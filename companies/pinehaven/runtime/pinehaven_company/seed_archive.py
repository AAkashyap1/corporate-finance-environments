"""Verified transport for Pinehaven's immutable ERP seed.

The uncompressed SQLite seed is slightly larger than GitHub's ordinary blob
limit. The release therefore commits a deterministic gzip archive and a
manifest containing both byte lengths and SHA-256 digests. Episode reset
streams the archive directly into the mutable state directory when the local
developer seed is absent.
"""

from __future__ import annotations

import contextlib
import gzip
import hashlib
import json
import lzma
import os
import stat
import tempfile
from pathlib import Path
from pathlib import PurePosixPath
from typing import Any
from typing import BinaryIO
from typing import Iterator


DATABASE_NAME = "erp_seed.db"
ARCHIVE_NAME = "erp_seed.db.gz"
MANIFEST_NAME = "ERP_SEED_MANIFEST.json"
TRANSPORT_DIRECTORY = "transport"
MAX_MANIFEST_BYTES = 1024 * 1024
MAX_ARCHIVE_BYTES = 64 * 1024 * 1024
MAX_DATABASE_BYTES = 256 * 1024 * 1024
MAX_PART_BYTES = 1024 * 1024
MAX_TRANSPORT_PARTS = 256
COPY_BLOCK_BYTES = 1024 * 1024


def _is_sha256(value: object) -> bool:
    return (
        isinstance(value, str)
        and len(value) == 64
        and all(character in "0123456789abcdef" for character in value)
    )


def _reject_duplicate_keys(
    pairs: list[tuple[str, Any]],
) -> dict[str, Any]:
    payload: dict[str, Any] = {}
    for key, value in pairs:
        if key in payload:
            raise ValueError(f"ERP seed manifest contains a duplicate key: {key!r}")
        payload[key] = value
    return payload


def _validated_relative_path(value: object, *, expected: str | None = None) -> str:
    if not isinstance(value, str) or not value:
        raise ValueError("ERP seed manifest contains an invalid path")
    path = PurePosixPath(value)
    if (
        value in {".", ".."}
        or path.is_absolute()
        or "." in path.parts
        or ".." in path.parts
        or path.as_posix() != value
        or "\\" in value
        or (path.parts and path.parts[0].endswith(":"))
    ):
        raise ValueError(f"ERP seed manifest contains an unsafe path: {value!r}")
    if expected is not None and value != expected:
        raise ValueError(
            f"ERP seed manifest path mismatch: expected {expected!r}, got {value!r}"
        )
    return value


@contextlib.contextmanager
def _open_seed_file(
    seed: Path,
    relative: str,
    *,
    maximum_bytes: int,
) -> Iterator[tuple[BinaryIO, os.stat_result]]:
    """Open a regular seed file without following any path-component symlink."""

    _validated_relative_path(relative)
    directory_flags = os.O_RDONLY | getattr(os, "O_DIRECTORY", 0)
    no_follow = getattr(os, "O_NOFOLLOW", 0)
    close_on_exec = getattr(os, "O_CLOEXEC", 0)
    current_fd = os.open(seed, directory_flags | no_follow | close_on_exec)
    try:
        components = PurePosixPath(relative).parts
        for component in components[:-1]:
            next_fd = os.open(
                component,
                directory_flags | no_follow | close_on_exec,
                dir_fd=current_fd,
            )
            os.close(current_fd)
            current_fd = next_fd
        file_fd = os.open(
            components[-1],
            os.O_RDONLY | no_follow | close_on_exec,
            dir_fd=current_fd,
        )
        try:
            metadata = os.fstat(file_fd)
            if not stat.S_ISREG(metadata.st_mode):
                raise ValueError(
                    f"ERP seed artifact is not a regular file: {relative}"
                )
            if metadata.st_size > maximum_bytes:
                raise ValueError(
                    f"ERP seed artifact exceeds its safety limit: {relative}"
                )
            with os.fdopen(file_fd, "rb", closefd=True) as handle:
                file_fd = -1
                yield handle, metadata
        finally:
            if file_fd >= 0:
                os.close(file_fd)
    finally:
        os.close(current_fd)


def _stream_sha256(
    handle: BinaryIO,
    *,
    expected_bytes: int,
) -> str:
    digest = hashlib.sha256()
    byte_count = 0
    while block := handle.read(min(COPY_BLOCK_BYTES, expected_bytes - byte_count + 1)):
        byte_count += len(block)
        if byte_count > expected_bytes:
            raise ValueError("ERP seed artifact exceeds its declared byte length")
        digest.update(block)
    if byte_count != expected_bytes:
        raise ValueError("ERP seed artifact has an unexpected byte length")
    return digest.hexdigest()


def _copy_exact(
    source: BinaryIO,
    target: BinaryIO,
    *,
    expected_bytes: int,
    expected_sha256: str,
) -> None:
    digest = hashlib.sha256()
    byte_count = 0
    while block := source.read(min(COPY_BLOCK_BYTES, expected_bytes - byte_count + 1)):
        byte_count += len(block)
        if byte_count > expected_bytes:
            raise ValueError("ERP seed materialization exceeds its declared byte length")
        digest.update(block)
        target.write(block)
    if byte_count != expected_bytes:
        raise ValueError("ERP seed materialization has an unexpected byte length")
    if digest.hexdigest() != expected_sha256:
        raise ValueError("ERP seed materialization has an unexpected SHA-256")


def _validate_manifest(payload: object) -> dict[str, Any]:
    if not isinstance(payload, dict):
        raise ValueError("ERP seed manifest must be a JSON object")
    required_top_level = {
        "schema_version",
        "company_id",
        "snapshot",
        "database",
    }
    if payload.get("schema_version") != 1:
        raise ValueError("Unsupported ERP seed manifest schema")
    if not isinstance(payload.get("company_id"), str) or not payload["company_id"]:
        raise ValueError("ERP seed manifest has an invalid company_id")
    if not isinstance(payload.get("snapshot"), str) or not payload["snapshot"]:
        raise ValueError("ERP seed manifest has an invalid snapshot")
    transport = payload.get("transport")
    archive = payload.get("archive")
    if isinstance(transport, dict) == isinstance(archive, dict):
        raise ValueError(
            "ERP seed manifest must declare exactly one transport representation"
        )
    expected_top_level = required_top_level | (
        {"transport"} if isinstance(transport, dict) else {"archive"}
    )
    if set(payload) != expected_top_level:
        raise ValueError("ERP seed manifest has unexpected or missing fields")

    database = payload.get("database")
    if not isinstance(database, dict) or set(database) != {
        "path",
        "bytes",
        "sha256",
        "sqlite_quick_check",
        "sqlite_page_count",
        "sqlite_page_size",
    }:
        raise ValueError("ERP seed manifest has an invalid database record")
    _validated_relative_path(
        database.get("path"),
        expected=f"company/seed/{DATABASE_NAME}",
    )
    database_bytes = database.get("bytes")
    page_count = database.get("sqlite_page_count")
    page_size = database.get("sqlite_page_size")
    if (
        not isinstance(database_bytes, int)
        or isinstance(database_bytes, bool)
        or database_bytes <= 0
        or database_bytes > MAX_DATABASE_BYTES
        or not _is_sha256(database.get("sha256"))
        or database.get("sqlite_quick_check") != "ok"
        or not isinstance(page_count, int)
        or isinstance(page_count, bool)
        or page_count <= 0
        or not isinstance(page_size, int)
        or isinstance(page_size, bool)
        or page_size < 512
        or page_size > 65536
        or page_size & (page_size - 1)
        or page_count * page_size != database_bytes
    ):
        raise ValueError("ERP seed manifest has invalid database metadata")

    if isinstance(transport, dict):
        if set(transport) != {
            "format",
            "part_bytes",
            "archive_bytes",
            "archive_sha256",
            "parts",
        }:
            raise ValueError("ERP seed manifest has an invalid transport record")
        part_bytes = transport.get("part_bytes")
        archive_bytes = transport.get("archive_bytes")
        parts = transport.get("parts")
        if (
            transport.get("format") != "xz"
            or not isinstance(part_bytes, int)
            or isinstance(part_bytes, bool)
            or part_bytes <= 0
            or part_bytes > MAX_PART_BYTES
            or not isinstance(archive_bytes, int)
            or isinstance(archive_bytes, bool)
            or archive_bytes <= 0
            or archive_bytes > MAX_ARCHIVE_BYTES
            or not _is_sha256(transport.get("archive_sha256"))
            or not isinstance(parts, list)
            or not parts
            or len(parts) > MAX_TRANSPORT_PARTS
        ):
            raise ValueError("ERP seed manifest has invalid transport metadata")
        declared_archive_bytes = 0
        for index, record in enumerate(parts):
            expected_path = f"{TRANSPORT_DIRECTORY}/erp-{index:04d}.part"
            if not isinstance(record, dict) or set(record) != {
                "path",
                "bytes",
                "sha256",
            }:
                raise ValueError("ERP seed manifest has an invalid transport part")
            _validated_relative_path(record.get("path"), expected=expected_path)
            size = record.get("bytes")
            if (
                not isinstance(size, int)
                or isinstance(size, bool)
                or size <= 0
                or size > part_bytes
                or (index < len(parts) - 1 and size != part_bytes)
                or not _is_sha256(record.get("sha256"))
            ):
                raise ValueError("ERP seed manifest has invalid transport part metadata")
            declared_archive_bytes += size
        if declared_archive_bytes != archive_bytes:
            raise ValueError("ERP seed transport declared byte lengths do not add up")
    else:
        if not isinstance(archive, dict) or set(archive) != {
            "path",
            "bytes",
            "sha256",
        }:
            raise ValueError("ERP seed manifest has an invalid archive record")
        _validated_relative_path(archive.get("path"), expected=ARCHIVE_NAME)
        archive_bytes = archive.get("bytes")
        if (
            not isinstance(archive_bytes, int)
            or isinstance(archive_bytes, bool)
            or archive_bytes <= 0
            or archive_bytes > MAX_ARCHIVE_BYTES
            or not _is_sha256(archive.get("sha256"))
        ):
            raise ValueError("ERP seed manifest has invalid archive metadata")
    return payload


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def load_seed_manifest(seed: Path) -> dict[str, Any]:
    try:
        with _open_seed_file(
            seed,
            MANIFEST_NAME,
            maximum_bytes=MAX_MANIFEST_BYTES,
        ) as (handle, _):
            payload = json.load(handle, object_pairs_hook=_reject_duplicate_keys)
    except FileNotFoundError as exc:
        raise FileNotFoundError(
            f"ERP seed manifest is missing: {seed / MANIFEST_NAME}"
        ) from exc
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("ERP seed manifest is not valid JSON") from exc
    return _validate_manifest(payload)


def verify_seed_artifact(seed: Path, *, require_raw: bool = False) -> dict[str, Any]:
    manifest = load_seed_manifest(seed)
    transport = manifest.get("transport")
    if isinstance(transport, dict):
        digest = hashlib.sha256()
        byte_count = 0
        for record in transport["parts"]:
            relative = record["path"]
            with _open_seed_file(
                seed,
                relative,
                maximum_bytes=record["bytes"],
            ) as (handle, metadata):
                if metadata.st_size != record["bytes"]:
                    raise ValueError(
                        f"ERP seed transport part size mismatch: {relative}"
                    )
                part_digest = hashlib.sha256()
                part_bytes = 0
                for block in iter(lambda: handle.read(COPY_BLOCK_BYTES), b""):
                    part_digest.update(block)
                    digest.update(block)
                    part_bytes += len(block)
                    byte_count += len(block)
                if part_bytes != record["bytes"]:
                    raise ValueError(
                        f"ERP seed transport part size mismatch: {relative}"
                    )
                if part_digest.hexdigest() != record["sha256"]:
                    raise ValueError(
                        f"ERP seed transport part SHA-256 mismatch: {relative}"
                    )
        if byte_count != transport["archive_bytes"]:
            raise ValueError("ERP seed transport total byte length mismatch")
        if digest.hexdigest() != transport["archive_sha256"]:
            raise ValueError("ERP seed transport aggregate SHA-256 mismatch")
    else:
        archive_record = manifest["archive"]
        with _open_seed_file(
            seed,
            ARCHIVE_NAME,
            maximum_bytes=archive_record["bytes"],
        ) as (handle, metadata):
            if metadata.st_size != archive_record["bytes"]:
                raise ValueError(
                    "ERP seed archive byte length does not match its manifest"
                )
            if (
                _stream_sha256(handle, expected_bytes=archive_record["bytes"])
                != archive_record["sha256"]
            ):
                raise ValueError(
                    "ERP seed archive SHA-256 does not match its manifest"
                )
    raw_record = manifest["database"]
    try:
        with _open_seed_file(
            seed,
            DATABASE_NAME,
            maximum_bytes=raw_record["bytes"],
        ) as (handle, metadata):
            if metadata.st_size != raw_record["bytes"]:
                raise ValueError(
                    "ERP seed database byte length does not match its manifest"
                )
            if (
                _stream_sha256(handle, expected_bytes=raw_record["bytes"])
                != raw_record["sha256"]
            ):
                raise ValueError(
                    "ERP seed database SHA-256 does not match its manifest"
                )
    except FileNotFoundError:
        if require_raw:
            raise FileNotFoundError(
                f"Uncompressed ERP seed is missing: {seed / DATABASE_NAME}"
            ) from None
    return manifest


@contextlib.contextmanager
def _atomic_destination(destination: Path) -> Iterator[BinaryIO]:
    destination.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{destination.name}.",
        suffix=".tmp",
        dir=destination.parent,
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "w+b", closefd=True) as handle:
            yield handle
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, destination)
    except BaseException:
        temporary.unlink(missing_ok=True)
        raise


@contextlib.contextmanager
def _assembled_transport(
    seed: Path,
    transport: dict[str, Any],
    destination_parent: Path,
) -> Iterator[BinaryIO]:
    descriptor, assembled_name = tempfile.mkstemp(
        prefix=".erp-seed.",
        suffix=".xz",
        dir=destination_parent,
    )
    assembled = Path(assembled_name)
    digest = hashlib.sha256()
    byte_count = 0
    try:
        with os.fdopen(descriptor, "w+b", closefd=True) as target:
            for record in transport["parts"]:
                with _open_seed_file(
                    seed,
                    record["path"],
                    maximum_bytes=record["bytes"],
                ) as (source, metadata):
                    if metadata.st_size != record["bytes"]:
                        raise ValueError(
                            "ERP seed transport part size changed during materialization"
                        )
                    part_digest = hashlib.sha256()
                    part_bytes = 0
                    while block := source.read(COPY_BLOCK_BYTES):
                        part_bytes += len(block)
                        byte_count += len(block)
                        if (
                            part_bytes > record["bytes"]
                            or byte_count > transport["archive_bytes"]
                        ):
                            raise ValueError(
                                "ERP seed transport exceeds its declared byte length"
                            )
                        part_digest.update(block)
                        digest.update(block)
                        target.write(block)
                    if (
                        part_bytes != record["bytes"]
                        or part_digest.hexdigest() != record["sha256"]
                    ):
                        raise ValueError(
                            "ERP seed transport part changed during materialization"
                        )
            if (
                byte_count != transport["archive_bytes"]
                or digest.hexdigest() != transport["archive_sha256"]
            ):
                raise ValueError(
                    "ERP seed transport changed during materialization"
                )
            target.flush()
            os.fsync(target.fileno())
            target.seek(0)
            yield target
    finally:
        assembled.unlink(missing_ok=True)


def copy_seed_database(seed: Path, destination: Path) -> Path:
    """Copy or decompress the verified immutable seed into ``destination``."""

    manifest = verify_seed_artifact(seed)
    expected = manifest["database"]
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        with _open_seed_file(
            seed,
            DATABASE_NAME,
            maximum_bytes=expected["bytes"],
        ) as (source, metadata):
            if metadata.st_size != expected["bytes"]:
                raise ValueError("ERP seed database changed during materialization")
            with _atomic_destination(destination) as target:
                _copy_exact(
                    source,
                    target,
                    expected_bytes=expected["bytes"],
                    expected_sha256=expected["sha256"],
                )
            return destination
    except FileNotFoundError:
        pass

    transport = manifest.get("transport")
    with _atomic_destination(destination) as target:
        if isinstance(transport, dict):
            with _assembled_transport(
                seed,
                transport,
                destination.parent,
            ) as assembled:
                with lzma.LZMAFile(assembled, "rb") as source:
                    _copy_exact(
                        source,
                        target,
                        expected_bytes=expected["bytes"],
                        expected_sha256=expected["sha256"],
                    )
        else:
            archive_record = manifest["archive"]
            with _open_seed_file(
                seed,
                ARCHIVE_NAME,
                maximum_bytes=archive_record["bytes"],
            ) as (archive, metadata):
                if metadata.st_size != archive_record["bytes"]:
                    raise ValueError(
                        "ERP seed archive changed during materialization"
                    )
                if (
                    _stream_sha256(
                        archive,
                        expected_bytes=archive_record["bytes"],
                    )
                    != archive_record["sha256"]
                ):
                    raise ValueError(
                        "ERP seed archive changed during materialization"
                    )
                archive.seek(0)
                with gzip.GzipFile(fileobj=archive, mode="rb") as source:
                    _copy_exact(
                        source,
                        target,
                        expected_bytes=expected["bytes"],
                        expected_sha256=expected["sha256"],
                    )
    return destination
