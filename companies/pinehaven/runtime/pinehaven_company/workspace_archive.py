"""Verified transport and extraction for the immutable visible workspace."""

from __future__ import annotations

import contextlib
import hashlib
import json
import os
import shutil
import stat
import tarfile
import tempfile
from pathlib import Path
from pathlib import PurePosixPath
from typing import Any
from typing import BinaryIO
from typing import Iterator

from pinehaven_company.seed_archive import sha256_file


MANIFEST_NAME = "WORKSPACE_SEED_MANIFEST.json"
SOURCE_MANIFEST_MEMBER = "SOURCE_MANIFEST.json"
SOURCE_MANIFEST_PATH = "company/seed/SOURCE_MANIFEST.json"
MAX_SOURCE_MANIFEST_BYTES = 16 * 1024 * 1024
MAX_WORKSPACE_MANIFEST_BYTES = 1024 * 1024
MAX_WORKSPACE_ARCHIVE_BYTES = 64 * 1024 * 1024
MAX_WORKSPACE_UNPACKED_BYTES = 256 * 1024 * 1024
MAX_WORKSPACE_MEMBER_BYTES = 32 * 1024 * 1024
MAX_WORKSPACE_FILES = 4_096
MAX_WORKSPACE_PART_BYTES = 1024 * 1024
MAX_WORKSPACE_TRANSPORT_PARTS = 256
MAX_WORKSPACE_PATH_BYTES = 4_096
MAX_WORKSPACE_PATH_DEPTH = 64
COPY_BLOCK_BYTES = 1024 * 1024
TRANSPORT_DIRECTORY = "transport"
SOURCE_ARTIFACT_COUNT_KEYS = frozenset(
    {
        "csv",
        "docx",
        "eml",
        "markdown",
        "pdf",
        "pptx",
        "xlsx_analysis_workbooks",
        "xlsx_close_binders",
        "xlsx_plant_scorecards",
    }
)


def _strict_json_object(
    pairs: list[tuple[str, Any]],
) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(
                f"Workspace manifest contains a duplicate JSON key: {key}"
            )
        result[key] = value
    return result


def _is_sha256(value: object) -> bool:
    return (
        isinstance(value, str)
        and len(value) == 64
        and all(character in "0123456789abcdef" for character in value)
    )


def _validated_relative_path(
    value: object,
    *,
    expected: str | None = None,
) -> str:
    if not isinstance(value, str) or not value:
        raise ValueError("Workspace seed manifest contains an invalid path")
    path = PurePosixPath(value)
    if (
        value in {".", ".."}
        or path.is_absolute()
        or "." in path.parts
        or ".." in path.parts
        or path.as_posix() != value
        or "\\" in value
        or len(path.parts) > MAX_WORKSPACE_PATH_DEPTH
        or len(value.encode("utf-8")) > MAX_WORKSPACE_PATH_BYTES
        or (path.parts and path.parts[0].endswith(":"))
    ):
        raise ValueError(
            f"Workspace seed manifest contains an unsafe path: {value!r}"
        )
    if expected is not None and value != expected:
        raise ValueError(
            "Workspace seed manifest path mismatch: "
            f"expected {expected!r}, got {value!r}"
        )
    return value


@contextlib.contextmanager
def _open_workspace_seed_file(
    seed: Path,
    relative: str,
    *,
    maximum_bytes: int,
) -> Iterator[tuple[BinaryIO, os.stat_result]]:
    """Open a bounded regular seed file without following component links."""

    _validated_relative_path(relative)
    directory_flags = (
        os.O_RDONLY
        | getattr(os, "O_DIRECTORY", 0)
        | getattr(os, "O_NOFOLLOW", 0)
        | getattr(os, "O_CLOEXEC", 0)
    )
    file_flags = (
        os.O_RDONLY
        | getattr(os, "O_NOFOLLOW", 0)
        | getattr(os, "O_CLOEXEC", 0)
    )
    current_fd = os.open(seed, directory_flags)
    try:
        components = PurePosixPath(relative).parts
        for component in components[:-1]:
            next_fd = os.open(
                component,
                directory_flags,
                dir_fd=current_fd,
            )
            os.close(current_fd)
            current_fd = next_fd
        file_fd = os.open(
            components[-1],
            file_flags,
            dir_fd=current_fd,
        )
        try:
            metadata = os.fstat(file_fd)
            if not stat.S_ISREG(metadata.st_mode):
                raise ValueError(
                    "Workspace seed artifact is not a regular file: "
                    f"{relative}"
                )
            if metadata.st_size > maximum_bytes:
                raise ValueError(
                    "Workspace seed artifact exceeds its safety limit: "
                    f"{relative}"
                )
            with os.fdopen(file_fd, "rb", closefd=True) as handle:
                file_fd = -1
                yield handle, metadata
        finally:
            if file_fd >= 0:
                os.close(file_fd)
    finally:
        os.close(current_fd)


def _validate_workspace_manifest(payload: object) -> dict[str, Any]:
    if not isinstance(payload, dict) or set(payload) != {
        "schema_version",
        "company_id",
        "snapshot",
        "files",
        "unpacked_bytes",
        "source_manifest_path",
        "source_manifest_sha256",
        "transport",
    }:
        raise ValueError(
            "Workspace seed manifest has unexpected or missing fields"
        )
    if payload.get("schema_version") != 1:
        raise ValueError("Unsupported workspace seed manifest schema")
    if (
        not isinstance(payload.get("company_id"), str)
        or not payload["company_id"]
        or not isinstance(payload.get("snapshot"), str)
        or not payload["snapshot"]
    ):
        raise ValueError("Workspace seed manifest identity is invalid")
    files = payload.get("files")
    unpacked_bytes = payload.get("unpacked_bytes")
    if (
        not isinstance(files, int)
        or isinstance(files, bool)
        or files <= 0
        or files > MAX_WORKSPACE_FILES
        or not isinstance(unpacked_bytes, int)
        or isinstance(unpacked_bytes, bool)
        or unpacked_bytes <= 0
        or unpacked_bytes > MAX_WORKSPACE_UNPACKED_BYTES
    ):
        raise ValueError("Workspace seed manifest file bounds are invalid")
    if payload.get("source_manifest_path") != SOURCE_MANIFEST_PATH:
        raise ValueError("Workspace source_manifest_path mismatch")
    if not _is_sha256(payload.get("source_manifest_sha256")):
        raise ValueError("Workspace source manifest SHA-256 is invalid")
    transport = payload.get("transport")
    if not isinstance(transport, dict) or set(transport) != {
        "format",
        "part_bytes",
        "archive_bytes",
        "archive_sha256",
        "parts",
    }:
        raise ValueError("Workspace seed transport record is invalid")
    part_bytes = transport.get("part_bytes")
    archive_bytes = transport.get("archive_bytes")
    parts = transport.get("parts")
    if (
        transport.get("format") != "xz"
        or not isinstance(part_bytes, int)
        or isinstance(part_bytes, bool)
        or part_bytes <= 0
        or part_bytes > MAX_WORKSPACE_PART_BYTES
        or not isinstance(archive_bytes, int)
        or isinstance(archive_bytes, bool)
        or archive_bytes <= 0
        or archive_bytes > MAX_WORKSPACE_ARCHIVE_BYTES
        or not _is_sha256(transport.get("archive_sha256"))
        or not isinstance(parts, list)
        or not parts
        or len(parts) > MAX_WORKSPACE_TRANSPORT_PARTS
    ):
        raise ValueError("Workspace seed transport metadata is invalid")
    declared_archive_bytes = 0
    for index, record in enumerate(parts):
        if not isinstance(record, dict) or set(record) != {
            "path",
            "bytes",
            "sha256",
        }:
            raise ValueError("Workspace seed transport part is invalid")
        _validated_relative_path(
            record.get("path"),
            expected=f"{TRANSPORT_DIRECTORY}/workspace-{index:04d}.part",
        )
        size = record.get("bytes")
        if (
            not isinstance(size, int)
            or isinstance(size, bool)
            or size <= 0
            or size > part_bytes
            or (index < len(parts) - 1 and size != part_bytes)
            or not _is_sha256(record.get("sha256"))
        ):
            raise ValueError(
                "Workspace seed transport part metadata is invalid"
            )
        declared_archive_bytes += size
    if declared_archive_bytes != archive_bytes:
        raise ValueError(
            "Workspace seed transport byte lengths do not add up"
        )
    return payload


def load_workspace_manifest(seed: Path) -> dict[str, Any]:
    try:
        with _open_workspace_seed_file(
            seed,
            MANIFEST_NAME,
            maximum_bytes=MAX_WORKSPACE_MANIFEST_BYTES,
        ) as (handle, _metadata):
            payload = json.load(
                handle,
                object_pairs_hook=_strict_json_object,
            )
    except FileNotFoundError as exc:
        raise FileNotFoundError(
            f"Workspace seed manifest is missing: {seed / MANIFEST_NAME}"
        ) from exc
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(
            "Workspace seed manifest is not valid JSON"
        ) from exc
    return _validate_workspace_manifest(payload)


@contextlib.contextmanager
def _assembled_workspace_transport(
    seed: Path,
    transport: dict[str, Any],
) -> Iterator[Path]:
    with tempfile.TemporaryDirectory(
        prefix="pinehaven-workspace-transport-"
    ) as directory:
        assembled = Path(directory) / "workspace-seed.tar.xz"
        descriptor = os.open(
            assembled,
            os.O_WRONLY
            | os.O_CREAT
            | os.O_EXCL
            | getattr(os, "O_NOFOLLOW", 0)
            | getattr(os, "O_CLOEXEC", 0),
            0o600,
        )
        digest = hashlib.sha256()
        byte_count = 0
        try:
            with os.fdopen(descriptor, "wb", closefd=True) as target:
                descriptor = -1
                for record in transport["parts"]:
                    relative = record["path"]
                    with _open_workspace_seed_file(
                        seed,
                        relative,
                        maximum_bytes=record["bytes"],
                    ) as (source, metadata):
                        if metadata.st_size != record["bytes"]:
                            raise ValueError(
                                "Workspace transport part size mismatch: "
                                f"{relative}"
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
                                    "Workspace transport exceeds its "
                                    "declared byte length"
                                )
                            part_digest.update(block)
                            digest.update(block)
                            target.write(block)
                        if (
                            part_bytes != record["bytes"]
                            or part_digest.hexdigest()
                            != record["sha256"]
                        ):
                            raise ValueError(
                                "Workspace transport part changed during "
                                f"assembly: {relative}"
                            )
                if (
                    byte_count != transport["archive_bytes"]
                    or digest.hexdigest() != transport["archive_sha256"]
                ):
                    raise ValueError(
                        "Workspace transport aggregate bytes or SHA-256 "
                        "mismatch"
                    )
                target.flush()
                os.fsync(target.fileno())
        finally:
            if descriptor >= 0:
                os.close(descriptor)
        yield assembled


def verify_workspace_artifact(seed: Path) -> dict[str, Any]:
    manifest = load_workspace_manifest(seed)
    transport = manifest["transport"]
    with _assembled_workspace_transport(seed, transport) as assembled:
        _verify_workspace_members(assembled, manifest)
    return manifest


def _safe_members(archive: tarfile.TarFile) -> list[tarfile.TarInfo]:
    members: list[tarfile.TarInfo] = []
    seen: set[str] = set()
    total_size = 0
    for member in archive:
        if len(members) >= MAX_WORKSPACE_FILES:
            raise ValueError(
                "Workspace archive member count exceeds its safety limit"
            )
        if member.name in seen:
            raise ValueError(
                f"Workspace archive duplicate member: {member.name}"
            )
        seen.add(member.name)
        path = PurePosixPath(member.name)
        if (
            member.name in {"", "."}
            or path.is_absolute()
            or ".." in path.parts
            or path.as_posix() != member.name
            or "\\" in member.name
            or len(path.parts) > MAX_WORKSPACE_PATH_DEPTH
            or len(member.name.encode("utf-8")) > MAX_WORKSPACE_PATH_BYTES
            or (path.parts and path.parts[0].endswith(":"))
        ):
            raise ValueError(f"Unsafe path in workspace archive: {member.name}")
        if not member.isfile():
            raise ValueError(
                f"Workspace archive contains a non-file member: {member.name}"
            )
        if member.size < 0 or member.size > MAX_WORKSPACE_MEMBER_BYTES:
            raise ValueError(
                "Workspace archive member exceeds its safety limit: "
                f"{member.name}"
            )
        total_size += member.size
        if total_size > MAX_WORKSPACE_UNPACKED_BYTES:
            raise ValueError(
                "Workspace archive expansion exceeds its safety limit"
            )
        members.append(member)
    return members


def _source_manifest_records(
    payload: bytes,
    outer_manifest: dict[str, Any],
) -> dict[str, dict[str, Any]]:
    try:
        source_manifest = json.loads(
            payload,
            object_pairs_hook=_strict_json_object,
        )
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(
            "Workspace archive SOURCE_MANIFEST.json is invalid"
        ) from exc
    if not isinstance(source_manifest, dict):
        raise ValueError(
            "Workspace archive SOURCE_MANIFEST.json must be an object"
        )
    if set(source_manifest) != {
        "artifact_counts",
        "files",
        "generated_at",
        "snapshot",
        "total_files",
        "company_id",
    }:
        raise ValueError(
            "Workspace archive SOURCE_MANIFEST.json schema mismatch"
        )
    if (
        not isinstance(source_manifest.get("generated_at"), str)
        or not source_manifest["generated_at"]
    ):
        raise ValueError(
            "Workspace archive SOURCE_MANIFEST.json generated_at is invalid"
        )
    for field in ("company_id", "snapshot"):
        if source_manifest.get(field) != outer_manifest.get(field):
            raise ValueError(
                f"Workspace source manifest {field} mismatch"
            )
    records = source_manifest.get("files")
    artifact_counts = source_manifest.get("artifact_counts")
    if not isinstance(records, list):
        raise ValueError(
            "Workspace archive SOURCE_MANIFEST.json lacks a files list"
        )
    if (
        not isinstance(artifact_counts, dict)
        or set(artifact_counts) != SOURCE_ARTIFACT_COUNT_KEYS
        or any(
            not isinstance(count, int)
            or isinstance(count, bool)
            or count < 0
            or count > MAX_WORKSPACE_FILES
            for count in artifact_counts.values()
        )
    ):
        raise ValueError(
            "Workspace archive SOURCE_MANIFEST.json artifact_counts "
            "schema mismatch"
        )
    declared: dict[str, dict[str, Any]] = {}
    declared_bytes = 0
    for record in records:
        if not isinstance(record, dict) or set(record) != {
            "path",
            "sha256",
            "size",
        }:
            raise ValueError(
                "Workspace archive SOURCE_MANIFEST.json has an invalid file record"
            )
        relative = record.get("path")
        size = record.get("size")
        digest = record.get("sha256")
        if (
            not isinstance(relative, str)
            or not relative
            or relative == SOURCE_MANIFEST_MEMBER
            or PurePosixPath(relative).is_absolute()
            or ".." in PurePosixPath(relative).parts
            or PurePosixPath(relative).as_posix() != relative
            or "\\" in relative
            or len(PurePosixPath(relative).parts)
            > MAX_WORKSPACE_PATH_DEPTH
            or len(relative.encode("utf-8")) > MAX_WORKSPACE_PATH_BYTES
            or (
                PurePosixPath(relative).parts
                and PurePosixPath(relative).parts[0].endswith(":")
            )
        ):
            raise ValueError(
                "Workspace archive SOURCE_MANIFEST.json has an unsafe file path"
            )
        if relative in declared:
            raise ValueError(
                f"Workspace source manifest duplicate member: {relative}"
            )
        if (
            not isinstance(size, int)
            or isinstance(size, bool)
            or size < 0
            or size > MAX_WORKSPACE_MEMBER_BYTES
        ):
            raise ValueError(
                f"Workspace source manifest has an invalid size: {relative}"
            )
        if not _is_sha256(digest):
            raise ValueError(
                f"Workspace source manifest has an invalid SHA-256: {relative}"
            )
        declared[relative] = record
        declared_bytes += size
        if declared_bytes > MAX_WORKSPACE_UNPACKED_BYTES:
            raise ValueError(
                "Workspace source manifest byte total exceeds its safety limit"
            )
    total_files = source_manifest.get("total_files")
    if (
        total_files != len(declared) + 1
        or total_files > MAX_WORKSPACE_FILES
        or sum(artifact_counts.values()) != len(declared)
    ):
        raise ValueError(
            "Workspace source manifest total_files mismatch"
        )
    return declared


def _verify_workspace_members(
    archive_path: Path,
    manifest: dict[str, Any],
) -> None:
    actual: dict[str, tuple[int, str]] = {}
    source_manifest_payload: bytes | None = None
    unpacked_bytes = 0
    with tarfile.open(archive_path, "r:xz") as archive:
        members = _safe_members(archive)
        if sum(member.size for member in members) != manifest.get(
            "unpacked_bytes"
        ):
            raise ValueError("Workspace archive unpacked byte count mismatch")
        for member in members:
            if (
                member.name == SOURCE_MANIFEST_MEMBER
                and member.size > MAX_SOURCE_MANIFEST_BYTES
            ):
                raise ValueError(
                    "Workspace archive SOURCE_MANIFEST.json is too large"
                )
            extracted = archive.extractfile(member)
            if extracted is None:
                raise ValueError(
                    f"Workspace archive member is unreadable: {member.name}"
                )
            digest = hashlib.sha256()
            payload = (
                bytearray()
                if member.name == SOURCE_MANIFEST_MEMBER
                else None
            )
            size = 0
            with extracted:
                for block in iter(
                    lambda: extracted.read(1024 * 1024),
                    b"",
                ):
                    digest.update(block)
                    size += len(block)
                    if size > member.size:
                        raise ValueError(
                            "Workspace archive member exceeds its declared "
                            f"size: {member.name}"
                        )
                    if payload is not None:
                        payload.extend(block)
            if size != member.size:
                raise ValueError(
                    f"Workspace archive member size mismatch: {member.name}"
                )
            actual[member.name] = (size, digest.hexdigest())
            unpacked_bytes += size
            if payload is not None:
                source_manifest_payload = bytes(payload)

    if source_manifest_payload is None:
        raise ValueError(
            "Workspace archive member set mismatch: "
            f"missing=['{SOURCE_MANIFEST_MEMBER}']"
        )
    source_manifest_hash = hashlib.sha256(
        source_manifest_payload
    ).hexdigest()
    if source_manifest_hash != manifest.get("source_manifest_sha256"):
        raise ValueError(
            "Workspace archive SOURCE_MANIFEST.json SHA-256 mismatch"
        )

    declared = _source_manifest_records(
        source_manifest_payload,
        manifest,
    )
    expected_members = set(declared) | {SOURCE_MANIFEST_MEMBER}
    actual_members = set(actual)
    missing = sorted(expected_members - actual_members)
    unexpected = sorted(actual_members - expected_members)
    if missing or unexpected:
        details = []
        if missing:
            details.append(f"missing={missing!r}")
        if unexpected:
            details.append(f"unexpected={unexpected!r}")
        raise ValueError(
            "Workspace archive member set mismatch: " + "; ".join(details)
        )

    for relative, record in declared.items():
        size, digest = actual[relative]
        if size != record["size"]:
            raise ValueError(
                f"Workspace archive member size mismatch: {relative}"
            )
        if digest != record["sha256"]:
            raise ValueError(
                f"Workspace archive member SHA-256 mismatch: {relative}"
            )

    if manifest.get("files") != len(actual):
        raise ValueError("Workspace archive file count mismatch")
    if manifest.get("unpacked_bytes") != unpacked_bytes:
        raise ValueError("Workspace archive unpacked byte count mismatch")


def copy_seed_workspace(seed: Path, destination: Path) -> Path:
    source_workspace = seed / "workspace"
    if source_workspace.is_dir():
        shutil.copytree(source_workspace, destination, dirs_exist_ok=True)
        return destination

    manifest = load_workspace_manifest(seed)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.mkdir(parents=True, exist_ok=True)
    if any(destination.iterdir()):
        raise ValueError(
            "Workspace seed destination must be empty before extraction"
        )
    with _assembled_workspace_transport(
        seed,
        manifest["transport"],
    ) as assembled:
        _verify_workspace_members(assembled, manifest)
        with tarfile.open(assembled, "r:xz") as archive:
            members = _safe_members(archive)
            archive.extractall(destination, members=members, filter="data")
    source_manifest = destination / "SOURCE_MANIFEST.json"
    if not source_manifest.is_file():
        raise ValueError("Materialized workspace lacks SOURCE_MANIFEST.json")
    if sha256_file(source_manifest) != manifest["source_manifest_sha256"]:
        raise ValueError("Materialized SOURCE_MANIFEST.json has an unexpected hash")
    return destination
