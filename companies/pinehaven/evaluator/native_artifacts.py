from __future__ import annotations

import hashlib
import io
import os
import posixpath
import re
import secrets
import shutil
import stat
import subprocess
import tempfile
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import BinaryIO, Iterator
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree
from zipfile import (
    ZIP_DEFLATED,
    ZIP_STORED,
    BadZipFile,
    ZipFile,
    ZipInfo,
)

from openpyxl.formula import Tokenizer


NATIVE_PACKAGE_SAFETY_VERSION = "native-ooxml-v3-2026-07-31"
NATIVE_PACKAGE_MAX_FILE_BYTES = 32 * 1024 * 1024
NATIVE_PACKAGE_MAX_MEMBERS = 2_048
NATIVE_PACKAGE_MAX_MEMBER_BYTES = 32 * 1024 * 1024
NATIVE_PACKAGE_MAX_EXPANDED_BYTES = 128 * 1024 * 1024
NATIVE_PACKAGE_RATIO_MIN_MEMBER_BYTES = 1 * 1024 * 1024
NATIVE_PACKAGE_MAX_COMPRESSION_RATIO = 200.0
NATIVE_PACKAGE_STREAM_CHUNK_BYTES = 1024 * 1024
NATIVE_PACKAGE_MAX_XML_BYTES = 16 * 1024 * 1024
NATIVE_PACKAGE_MAX_XML_NODES = 250_000
NATIVE_PACKAGE_MAX_XML_DEPTH = 128
NATIVE_PACKAGE_MAX_XML_ATTRIBUTES = 500_000
NATIVE_PACKAGE_MAX_ATTRIBUTES_PER_ELEMENT = 256
NATIVE_PACKAGE_MAX_CHART_EMBEDDINGS = 64
NATIVE_PACKAGE_MAX_EMBEDDED_EXPANDED_BYTES = 128 * 1024 * 1024
NATIVE_PACKAGE_ALLOWED_COMPRESSION = frozenset(
    {ZIP_STORED, ZIP_DEFLATED}
)
NATIVE_PACKAGE_SUFFIXES = frozenset({".docx", ".pptx", ".xlsx"})
LIBREOFFICE_SANDBOX_VERSION = "pinehaven-office-bwrap-v2-2026-07-26"
LIBREOFFICE_SANDBOX_REQUIRED_ENV = (
    "PINEHAVEN_REQUIRE_SHELL_ISOLATION"
)
# Calc reserves more than 3 GiB of virtual address space while opening even
# small XLSX files on the pinned 64-bit LibreOffice runtime.  Four GiB is the
# smallest tested bound that permits deterministic recalculation; RSS remains
# independently bounded by the hosted runner/container.
LIBREOFFICE_SANDBOX_MAX_ADDRESS_SPACE_BYTES = 4 * 1024 * 1024 * 1024
LIBREOFFICE_SANDBOX_MAX_FILE_BYTES = 256 * 1024 * 1024
LIBREOFFICE_SANDBOX_MAX_PROCESSES = 128
LIBREOFFICE_SANDBOX_MAX_OPEN_FILES = 256
LIBREOFFICE_SANDBOX_MAX_CPU_SECONDS = 40
LIBREOFFICE_SANDBOX_MAX_WALL_SECONDS = 45
LIBREOFFICE_SANDBOX_PROBE_MARKER = "pinehaven-office-sandbox-ok"
LIBREOFFICE_SANDBOX_MAX_DIAGNOSTIC_CHARS = 2_048
_LIBREOFFICE_SANDBOX_PROBE_VERIFIED = False
_RELATIONSHIPS_NAMESPACE = (
    "http://schemas.openxmlformats.org/package/2006/relationships"
)
_RELATIONSHIPS_TAG = f"{{{_RELATIONSHIPS_NAMESPACE}}}Relationships"
_RELATIONSHIP_TAG = f"{{{_RELATIONSHIPS_NAMESPACE}}}Relationship"
_CONTENT_TYPES_NAMESPACE = (
    "http://schemas.openxmlformats.org/package/2006/content-types"
)
_CONTENT_TYPES_TAG = f"{{{_CONTENT_TYPES_NAMESPACE}}}Types"
_CONTENT_TYPE_DEFAULT_TAG = f"{{{_CONTENT_TYPES_NAMESPACE}}}Default"
_CONTENT_TYPE_OVERRIDE_TAG = f"{{{_CONTENT_TYPES_NAMESPACE}}}Override"
_MAIN_PARTS = {
    ".docx": (
        "/word/document.xml",
        "application/vnd.openxmlformats-officedocument."
        "wordprocessingml.document.main+xml",
    ),
    ".pptx": (
        "/ppt/presentation.xml",
        "application/vnd.openxmlformats-officedocument."
        "presentationml.presentation.main+xml",
    ),
    ".xlsx": (
        "/xl/workbook.xml",
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet.main+xml",
    ),
}
_ACTIVE_MEMBER_SEGMENTS = (
    "/activex/",
    "/ctrlprops/",
    "/customui/",
    "/embeddings/",
    "/externallinks/",
)
_ACTIVE_MEMBER_BASENAMES = frozenset(
    {
        "vbadata.xml",
        "vbaproject.bin",
        "vbaprojectsignature.bin",
    }
)
_ACTIVE_FILE_SUFFIXES = frozenset(
    {
        ".bat",
        ".cmd",
        ".com",
        ".dll",
        ".exe",
        ".hta",
        ".jar",
        ".js",
        ".jse",
        ".msi",
        ".ps1",
        ".scr",
        ".vbe",
        ".vbs",
        ".wsf",
        ".wsh",
    }
)
_ACTIVE_CONTENT_TYPE_MARKERS = (
    "activex",
    "macroenabled",
    "oleobject",
    "vbaproject",
    "x-msdownload",
)
_BLOCKED_RELATIONSHIP_TYPE_SUFFIXES = (
    "/activexcontrol",
    "/activexcontrolbinary",
    "/attachedtemplate",
    "/control",
    "/customui",
    "/externallinkpath",
    "/oleobject",
    "/package",
    "/vbaproject",
)
_ACTIVE_XML_ELEMENTS = frozenset(
    {
        "activexcontrol",
        "altchunk",
        "control",
        "ddelink",
        "externalbook",
        "externallink",
        "olelink",
        "oleobject",
    }
)
_FIELD_INSTRUCTION_ELEMENTS = frozenset({"fldsimple", "instrtext"})
_ACTIVE_FIELD_INSTRUCTION = re.compile(
    r"(?i)\b(?:DDE|DDEAUTO|INCLUDEPICTURE|INCLUDETEXT|LINK)\b"
)
_ACTIVE_XML_ACTION = re.compile(
    r"(?i)(?:ppaction://(?:macro|program)|macro://)"
)
_FORBIDDEN_XML_DECLARATION = re.compile(
    br"(?i)<!\s*(?:doctype|entity)\b|<\?xml-stylesheet\b"
)
_EXTERNAL_FORMULA_LINK = re.compile(
    r"""(?ix)
    (?:
        \b(?:file|https?|ftp|smb):
        | \\\\
        | (?:^|['"(=,\s]) [A-Z]:[\\/]
        | \[[^\]\r\n]+\.(?:xlsx|xlsm|xlsb|xls|ods|csv)\]
    )
    """
)
_ACTIVE_FORMULA = re.compile(
    r"""(?ix)
    \b(?:CALL|DDE|EXEC|FILTERXML|REGISTER\.ID|RTD|RUN|WEBSERVICE)\s*\(
    """
)
_DDE_RANGE_OPERAND = re.compile(r"\|[^\r\n!]{0,512}!")
_DDE_FORMULA = re.compile(
    r"""(?ix)
    ^\s*=?\s*[A-Z_][A-Z0-9_.-]{0,63}\s*\|\s*
    (?:'[^'\r\n]{1,512}'|"[^"\r\n]{1,512}"|[^!\r\n]{1,512})
    \s*!\s*[^\r\n]{1,512}$
    """
)


def _has_active_formula(formula: str) -> bool:
    """Detect active functions and DDE range operands, not display literals."""

    if _ACTIVE_FORMULA.search(formula) or _DDE_FORMULA.search(formula):
        return True
    expression = formula if formula.startswith("=") else f"={formula}"
    try:
        tokens = Tokenizer(expression).items
    except Exception:
        # Malformed formulas are rejected elsewhere; do not turn an ordinary
        # display literal into an active-content false positive here.
        return False
    return any(
        token.type == "OPERAND"
        and token.subtype == "RANGE"
        and _DDE_RANGE_OPERAND.search(str(token.value))
        for token in tokens
    )


class NativePackageSafetyError(ValueError):
    """A native Office artifact violates the bounded package policy."""


@dataclass(frozen=True)
class NativePackageInspection:
    path: str
    file_bytes: int
    members: int
    expanded_bytes: int
    largest_member_bytes: int
    maximum_compression_ratio: float

    def to_dict(self) -> dict[str, str | int | float]:
        return asdict(self)


def native_package_policy_manifest() -> dict[str, object]:
    """Return the exact native-package limits pinned into release evidence."""

    return {
        "version": NATIVE_PACKAGE_SAFETY_VERSION,
        "allowed_suffixes": sorted(NATIVE_PACKAGE_SUFFIXES),
        "regular_non_symlink_path": True,
        "component_containment_required": True,
        "max_file_bytes": NATIVE_PACKAGE_MAX_FILE_BYTES,
        "max_members": NATIVE_PACKAGE_MAX_MEMBERS,
        "max_member_bytes": NATIVE_PACKAGE_MAX_MEMBER_BYTES,
        "max_expanded_bytes": NATIVE_PACKAGE_MAX_EXPANDED_BYTES,
        "compression_ratio_min_member_bytes": (
            NATIVE_PACKAGE_RATIO_MIN_MEMBER_BYTES
        ),
        "max_compression_ratio": NATIVE_PACKAGE_MAX_COMPRESSION_RATIO,
        "max_xml_bytes": NATIVE_PACKAGE_MAX_XML_BYTES,
        "max_xml_nodes": NATIVE_PACKAGE_MAX_XML_NODES,
        "max_xml_depth": NATIVE_PACKAGE_MAX_XML_DEPTH,
        "max_xml_attributes": NATIVE_PACKAGE_MAX_XML_ATTRIBUTES,
        "max_attributes_per_element": (
            NATIVE_PACKAGE_MAX_ATTRIBUTES_PER_ELEMENT
        ),
        "max_chart_workbook_embeddings": (
            NATIVE_PACKAGE_MAX_CHART_EMBEDDINGS
        ),
        "max_embedded_expanded_bytes": (
            NATIVE_PACKAGE_MAX_EMBEDDED_EXPANDED_BYTES
        ),
        "allowed_compression": ["stored", "deflated"],
        "encrypted_members_allowed": False,
        "symlink_members_allowed": False,
        "active_content_allowed": False,
        "macro_enabled_package_types_allowed": False,
        "external_relationships_allowed": False,
        "external_formula_links_allowed": False,
        "dtd_or_entity_declarations_allowed": False,
        "canonical_unique_member_names": True,
        "streaming_crc_verified": True,
        "libreoffice_sandbox": libreoffice_sandbox_policy_manifest(),
    }


def libreoffice_sandbox_policy_manifest() -> dict[str, object]:
    return {
        "version": LIBREOFFICE_SANDBOX_VERSION,
        "required_environment": LIBREOFFICE_SANDBOX_REQUIRED_ENV,
        "network_available": False,
        "host_environment_inherited": False,
        "host_filesystem_visible": False,
        "writable_mounts": ["/tmp", "/work"],
        "max_address_space_bytes": (
            LIBREOFFICE_SANDBOX_MAX_ADDRESS_SPACE_BYTES
        ),
        "max_file_bytes": LIBREOFFICE_SANDBOX_MAX_FILE_BYTES,
        "max_processes": LIBREOFFICE_SANDBOX_MAX_PROCESSES,
        "max_open_files": LIBREOFFICE_SANDBOX_MAX_OPEN_FILES,
        "max_cpu_seconds": LIBREOFFICE_SANDBOX_MAX_CPU_SECONDS,
        "max_wall_seconds": LIBREOFFICE_SANDBOX_MAX_WALL_SECONDS,
    }


def _lexical_absolute(path: str | Path) -> Path:
    return Path(os.path.abspath(os.fspath(path)))


def _isolation_required() -> bool:
    return os.environ.get(
        LIBREOFFICE_SANDBOX_REQUIRED_ENV,
        "",
    ).casefold() in {"1", "true", "yes", "on"}


def _sandbox_relative(
    path: str | Path,
    *,
    sandbox_root: Path,
) -> str:
    target = _lexical_absolute(path)
    try:
        relative = target.relative_to(sandbox_root)
    except ValueError as exc:
        raise NativePackageSafetyError(
            f"LibreOffice sandbox path escapes its private root: {target}"
        ) from exc
    if not relative.parts:
        return "/work"
    return f"/work/{relative.as_posix()}"


def _sandbox_binaries() -> tuple[str, str, str, str] | None:
    bwrap = shutil.which("bwrap")
    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    prlimit = shutil.which("prlimit")
    timeout_binary = shutil.which("timeout")
    if not all((bwrap, soffice, prlimit, timeout_binary)):
        return None
    resolved: list[str] = []
    for binary in (soffice, prlimit, timeout_binary):
        candidate = Path(str(binary)).resolve()
        if not candidate.is_relative_to("/usr"):
            return None
        resolved.append(candidate.as_posix())
    return str(bwrap), resolved[0], resolved[1], resolved[2]


def _sandbox_base_argv(
    *,
    bwrap: str,
    sandbox_root: Path,
) -> list[str]:
    argv = [
        bwrap,
        "--die-with-parent",
        "--new-session",
        "--unshare-user",
        "--unshare-pid",
        "--unshare-ipc",
        "--unshare-uts",
        "--unshare-cgroup-try",
        "--unshare-net",
        "--cap-drop",
        "ALL",
        "--ro-bind",
        "/usr",
        "/usr",
        "--symlink",
        "usr/lib",
        "/lib",
        "--symlink",
        "usr/lib64",
        "/lib64",
        "--symlink",
        "usr/bin",
        "/bin",
        "--symlink",
        "usr/sbin",
        "/sbin",
        "--dir",
        "/etc",
    ]
    for path in (
        "/etc/fonts",
        "/etc/group",
        "/etc/libreoffice",
        "/etc/localtime",
        "/etc/nsswitch.conf",
        "/etc/passwd",
    ):
        if Path(path).exists():
            argv.extend(["--ro-bind", path, path])
    argv.extend(
        [
            "--proc",
            "/proc",
            "--dev",
            "/dev",
            "--tmpfs",
            "/tmp",
            "--dir",
            "/run",
            "--bind",
            sandbox_root.as_posix(),
            "/work",
            "--chdir",
            "/work",
            "--clearenv",
            "--setenv",
            "HOME",
            "/work/home",
            "--setenv",
            "PATH",
            "/usr/bin:/bin",
            "--setenv",
            "TMPDIR",
            "/tmp",
            "--setenv",
            "XDG_CACHE_HOME",
            "/work/cache",
            "--setenv",
            "XDG_CONFIG_HOME",
            "/work/config",
            "--setenv",
            "LC_ALL",
            "C.UTF-8",
            "--setenv",
            "SAL_USE_VCLPLUGIN",
            "svp",
            "--",
        ]
    )
    return argv


def sandboxed_libreoffice_argv(
    *,
    sandbox_root: str | Path,
    input_path: str | Path,
    output_dir: str | Path,
    profile_dir: str | Path,
    convert_to: str,
    binaries: tuple[str, str, str, str] | None = None,
) -> list[str]:
    """Build the exact no-network, resource-bounded LibreOffice command."""

    root = _lexical_absolute(sandbox_root)
    selected = binaries or _sandbox_binaries()
    if selected is None:
        raise NativePackageSafetyError(
            "LibreOffice sandbox dependencies are unavailable"
        )
    bwrap, soffice, prlimit, timeout_binary = selected
    input_inside = _sandbox_relative(input_path, sandbox_root=root)
    output_inside = _sandbox_relative(output_dir, sandbox_root=root)
    profile_inside = _sandbox_relative(profile_dir, sandbox_root=root)
    return [
        *_sandbox_base_argv(bwrap=bwrap, sandbox_root=root),
        prlimit,
        f"--as={LIBREOFFICE_SANDBOX_MAX_ADDRESS_SPACE_BYTES}",
        f"--fsize={LIBREOFFICE_SANDBOX_MAX_FILE_BYTES}",
        f"--nproc={LIBREOFFICE_SANDBOX_MAX_PROCESSES}",
        f"--nofile={LIBREOFFICE_SANDBOX_MAX_OPEN_FILES}",
        f"--cpu={LIBREOFFICE_SANDBOX_MAX_CPU_SECONDS}",
        "--",
        timeout_binary,
        "--signal=KILL",
        "--kill-after=5s",
        f"{LIBREOFFICE_SANDBOX_MAX_WALL_SECONDS}s",
        soffice,
        "--headless",
        "--nologo",
        "--nodefault",
        "--nolockcheck",
        "--norestore",
        "--invisible",
        f"-env:UserInstallation={Path(profile_inside).as_uri()}",
        "--convert-to",
        convert_to,
        "--outdir",
        output_inside,
        input_inside,
    ]


def _safe_sandbox_environment() -> dict[str, str]:
    return {
        "HOME": "/nonexistent",
        "LC_ALL": "C.UTF-8",
        "PATH": "/usr/bin:/bin",
    }


def _bounded_process_diagnostic(value: str | None) -> str:
    text = value or ""
    limit = LIBREOFFICE_SANDBOX_MAX_DIAGNOSTIC_CHARS
    if len(text) <= limit:
        return text
    return f"{text[:limit]}...[truncated]"


def verify_libreoffice_sandbox() -> bool:
    """Probe both the isolated boundary and a real Office conversion once."""

    global _LIBREOFFICE_SANDBOX_PROBE_VERIFIED
    if _LIBREOFFICE_SANDBOX_PROBE_VERIFIED:
        return True
    binaries = _sandbox_binaries()
    if binaries is None:
        if _isolation_required():
            raise NativePackageSafetyError(
                "release grading requires bwrap, LibreOffice, prlimit, and "
                "timeout for the Office sandbox"
            )
        return False
    try:
        host_network_namespace = os.readlink("/proc/self/ns/net")
    except OSError as exc:
        if _isolation_required():
            raise NativePackageSafetyError(
                "LibreOffice sandbox cannot identify the host network "
                f"namespace: {type(exc).__name__}: {exc}"
            ) from exc
        return False
    if not re.fullmatch(r"net:\[\d+\]", host_network_namespace):
        if _isolation_required():
            raise NativePackageSafetyError(
                "LibreOffice sandbox received an invalid host network "
                f"namespace identity: {host_network_namespace!r}"
            )
        return False
    bwrap, _soffice, prlimit, timeout_binary = binaries
    with tempfile.TemporaryDirectory(
        prefix="pinehaven-office-probe-"
    ) as directory:
        root = Path(directory)
        for name in (
            "cache",
            "config",
            "home",
            "input",
            "output",
            "profile",
        ):
            (root / name).mkdir(mode=0o700)
        command = (
            "set -eu; "
            'test "$(pwd -P)" = /work; '
            "test ! -e /workspace; "
            'test "$(/usr/bin/readlink /proc/self/ns/net)" != '
            f"'{host_network_namespace}'; "
            "touch .pinehaven-office-probe; "
            "rm .pinehaven-office-probe; "
            f"printf {LIBREOFFICE_SANDBOX_PROBE_MARKER}"
        )
        argv = [
            *_sandbox_base_argv(bwrap=bwrap, sandbox_root=root),
            prlimit,
            f"--as={LIBREOFFICE_SANDBOX_MAX_ADDRESS_SPACE_BYTES}",
            f"--fsize={LIBREOFFICE_SANDBOX_MAX_FILE_BYTES}",
            f"--nproc={LIBREOFFICE_SANDBOX_MAX_PROCESSES}",
            f"--nofile={LIBREOFFICE_SANDBOX_MAX_OPEN_FILES}",
            f"--cpu={LIBREOFFICE_SANDBOX_MAX_CPU_SECONDS}",
            "--",
            timeout_binary,
            "--signal=KILL",
            "--kill-after=2s",
            "10s",
            "/usr/bin/bash",
            "-c",
            command,
        ]
        try:
            completed = subprocess.run(
                argv,
                cwd=root,
                env=_safe_sandbox_environment(),
                capture_output=True,
                text=True,
                timeout=15,
                check=False,
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            if _isolation_required():
                raise NativePackageSafetyError(
                    "LibreOffice sandbox boundary probe failed: "
                    f"{type(exc).__name__}: {exc}"
                ) from exc
            return False
        if (
            completed.returncode != 0
            or completed.stdout != LIBREOFFICE_SANDBOX_PROBE_MARKER
        ):
            if _isolation_required():
                raise NativePackageSafetyError(
                    "LibreOffice sandbox boundary probe failed closed: "
                    f"returncode={completed.returncode}; "
                    "stdout="
                    f"{_bounded_process_diagnostic(completed.stdout)!r}; "
                    "stderr="
                    f"{_bounded_process_diagnostic(completed.stderr)!r}"
                )
            return False

        probe_input = root / "input" / "probe.csv"
        probe_input.write_text(
            "label,value\npinehaven,1\n",
            encoding="utf-8",
        )
        office_argv = sandboxed_libreoffice_argv(
            sandbox_root=root,
            input_path=probe_input,
            output_dir=root / "output",
            profile_dir=root / "profile",
            convert_to="xlsx",
            binaries=binaries,
        )
        try:
            office_completed = subprocess.run(
                office_argv,
                cwd=root,
                env=_safe_sandbox_environment(),
                capture_output=True,
                text=True,
                timeout=LIBREOFFICE_SANDBOX_MAX_WALL_SECONDS + 10,
                check=False,
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            if _isolation_required():
                raise NativePackageSafetyError(
                    "LibreOffice sandbox conversion probe failed: "
                    f"{type(exc).__name__}: {exc}"
                ) from exc
            return False

        probe_output = root / "output" / "probe.xlsx"
        output_error = ""
        if office_completed.returncode == 0:
            try:
                output_stat = probe_output.lstat()
                if (
                    not stat.S_ISREG(output_stat.st_mode)
                    or output_stat.st_size <= 0
                    or output_stat.st_size > NATIVE_PACKAGE_MAX_FILE_BYTES
                ):
                    raise NativePackageSafetyError(
                        "conversion output is not a bounded regular file"
                    )
                inspect_native_package(
                    probe_output,
                    workspace_root=root,
                )
            except (OSError, NativePackageSafetyError) as exc:
                output_error = f"{type(exc).__name__}: {exc}"
        if office_completed.returncode != 0 or output_error:
            if _isolation_required():
                raise NativePackageSafetyError(
                    "LibreOffice sandbox conversion probe failed closed: "
                    f"returncode={office_completed.returncode}; "
                    f"output_error={output_error!r}; "
                    "stdout="
                    f"{_bounded_process_diagnostic(office_completed.stdout)!r}; "
                    "stderr="
                    f"{_bounded_process_diagnostic(office_completed.stderr)!r}"
                )
            return False
    _LIBREOFFICE_SANDBOX_PROBE_VERIFIED = True
    return True


def run_sandboxed_libreoffice(
    *,
    sandbox_root: str | Path,
    input_path: str | Path,
    output_dir: str | Path,
    profile_dir: str | Path,
    convert_to: str,
) -> subprocess.CompletedProcess[str] | None:
    """Run LibreOffice only inside the verified no-network bwrap profile."""

    if not verify_libreoffice_sandbox():
        return None
    root = _lexical_absolute(sandbox_root)
    argv = sandboxed_libreoffice_argv(
        sandbox_root=root,
        input_path=input_path,
        output_dir=output_dir,
        profile_dir=profile_dir,
        convert_to=convert_to,
    )
    try:
        return subprocess.run(
            argv,
            cwd=root,
            env=_safe_sandbox_environment(),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            text=True,
            timeout=LIBREOFFICE_SANDBOX_MAX_WALL_SECONDS + 10,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        if _isolation_required():
            raise NativePackageSafetyError(
                "sandboxed LibreOffice execution failed: "
                f"{type(exc).__name__}: {exc}"
            ) from exc
        return None


def _path_scope(
    path: str | Path,
    *,
    workspace_root: str | Path,
) -> tuple[Path, Path, Path]:
    target = _lexical_absolute(path)
    root = _lexical_absolute(workspace_root)
    try:
        relative = target.relative_to(root)
    except ValueError as exc:
        raise NativePackageSafetyError(
            f"native artifact is outside workspace root: {target}"
        ) from exc
    if not relative.parts:
        raise NativePackageSafetyError(
            "native artifact path must name a file below workspace root"
        )
    return target, root, relative


def _entry_stat(
    name: str,
    *,
    parent_fd: int,
    relative: Path,
) -> os.stat_result:
    try:
        return os.stat(name, dir_fd=parent_fd, follow_symlinks=False)
    except OSError as exc:
        raise NativePackageSafetyError(
            "native artifact path component is unavailable: "
            f"{relative.as_posix()}: {type(exc).__name__}: {exc}"
        ) from exc


def _external_reference(value: str) -> bool:
    decoded = value
    for _ in range(16):
        replacement = unquote(decoded)
        if replacement == decoded:
            break
        decoded = replacement
    else:
        return True
    return bool(_EXTERNAL_FORMULA_LINK.search(decoded))


def _validate_active_member_names(infos: list[ZipInfo]) -> None:
    for info in infos:
        if info.is_dir():
            continue
        folded = f"/{info.filename.casefold().lstrip('/')}"
        basename = posixpath.basename(folded)
        suffix = posixpath.splitext(basename)[1]
        chart_workbook = (
            folded.startswith("/ppt/embeddings/")
            and folded.endswith(".xlsx")
            and folded.count("/") == 3
        )
        if (
            basename in _ACTIVE_MEMBER_BASENAMES
            or suffix in _ACTIVE_FILE_SUFFIXES
            or any(
                segment in folded and not (
                    segment == "/embeddings/" and chart_workbook
                )
                for segment in _ACTIVE_MEMBER_SEGMENTS
            )
        ):
            raise NativePackageSafetyError(
                "native package contains active or embedded content: "
                f"{info.filename!r}"
            )


def _validate_all_xml_parts(
    archive: ZipFile,
    infos: list[ZipInfo],
) -> None:
    """Bound every XML tree before any Office library or LibreOffice sees it."""

    for info in infos:
        folded = info.filename.casefold()
        if (
            info.is_dir()
            or not (
                folded.endswith(".xml")
                or folded.endswith(".rels")
            )
        ):
            continue
        if info.file_size > NATIVE_PACKAGE_MAX_XML_BYTES:
            raise NativePackageSafetyError(
                f"{info.filename}: XML size exceeds "
                f"{NATIVE_PACKAGE_MAX_XML_BYTES}"
            )
        payload = archive.read(info)
        declaration = _FORBIDDEN_XML_DECLARATION.search(payload)
        if declaration:
            raise NativePackageSafetyError(
                f"{info.filename}: XML contains a forbidden DTD, entity, "
                "or stylesheet declaration"
            )
        depth = 0
        nodes = 0
        attributes = 0
        try:
            for event, element in ElementTree.iterparse(
                io.BytesIO(payload),
                events=("start", "end"),
            ):
                local_name = _local_name(element.tag).casefold()
                if event == "start":
                    depth += 1
                    nodes += 1
                    element_attributes = len(element.attrib)
                    attributes += element_attributes
                    if depth > NATIVE_PACKAGE_MAX_XML_DEPTH:
                        raise NativePackageSafetyError(
                            f"{info.filename}: XML depth exceeds "
                            f"{NATIVE_PACKAGE_MAX_XML_DEPTH}"
                        )
                    if nodes > NATIVE_PACKAGE_MAX_XML_NODES:
                        raise NativePackageSafetyError(
                            f"{info.filename}: XML node count exceeds "
                            f"{NATIVE_PACKAGE_MAX_XML_NODES}"
                        )
                    if (
                        element_attributes
                        > NATIVE_PACKAGE_MAX_ATTRIBUTES_PER_ELEMENT
                    ):
                        raise NativePackageSafetyError(
                            f"{info.filename}: XML element attribute count "
                            "exceeds "
                            f"{NATIVE_PACKAGE_MAX_ATTRIBUTES_PER_ELEMENT}"
                        )
                    if attributes > NATIVE_PACKAGE_MAX_XML_ATTRIBUTES:
                        raise NativePackageSafetyError(
                            f"{info.filename}: XML attribute count exceeds "
                            f"{NATIVE_PACKAGE_MAX_XML_ATTRIBUTES}"
                        )
                    if local_name in _ACTIVE_XML_ELEMENTS:
                        raise NativePackageSafetyError(
                            f"{info.filename}: XML contains active element "
                            f"{local_name!r}"
                        )
                    for attribute, value in element.attrib.items():
                        if not isinstance(value, str):
                            continue
                        if _ACTIVE_XML_ACTION.search(value):
                            raise NativePackageSafetyError(
                                f"{info.filename}: XML contains an active "
                                "macro or program action"
                            )
                        attribute_name = _local_name(attribute).casefold()
                        if (
                            attribute_name
                            in {
                                "action",
                                "data",
                                "href",
                                "link",
                                "path",
                                "source",
                                "src",
                                "target",
                                "url",
                            }
                            and not (
                                folded.endswith(".rels")
                                and attribute_name == "target"
                            )
                            and _external_reference(value)
                        ):
                            raise NativePackageSafetyError(
                                f"{info.filename}: XML contains an external "
                                "resource reference"
                            )
                else:
                    if local_name in _FIELD_INSTRUCTION_ELEMENTS:
                        instruction = " ".join(
                            [
                                element.text or "",
                                *(
                                    value
                                    for value in element.attrib.values()
                                    if isinstance(value, str)
                                ),
                            ]
                        )
                        if (
                            _ACTIVE_FIELD_INSTRUCTION.search(instruction)
                            or _external_reference(instruction)
                        ):
                            raise NativePackageSafetyError(
                                f"{info.filename}: XML contains an active or "
                                "external field instruction"
                            )
                    depth -= 1
                    element.clear()
        except NativePackageSafetyError:
            raise
        except ElementTree.ParseError as exc:
            raise NativePackageSafetyError(
                f"{info.filename}: XML is malformed: {exc}"
            ) from exc
        if depth != 0:
            raise NativePackageSafetyError(
                f"{info.filename}: XML element depth is unbalanced"
            )


def _validate_content_types(
    archive: ZipFile,
    infos: list[ZipInfo],
    *,
    suffix: str,
) -> None:
    names = {info.filename for info in infos if not info.is_dir()}
    content_types_name = "[Content_Types].xml"
    if content_types_name not in names:
        raise NativePackageSafetyError(
            "native package is missing [Content_Types].xml"
        )
    try:
        root = ElementTree.fromstring(archive.read(content_types_name))
    except ElementTree.ParseError as exc:
        raise NativePackageSafetyError(
            f"[Content_Types].xml is malformed: {exc}"
        ) from exc
    if root.tag != _CONTENT_TYPES_TAG:
        raise NativePackageSafetyError(
            "[Content_Types].xml has an invalid root"
        )
    main_part, main_content_type = _MAIN_PARTS[suffix]
    observed_main_types: list[str] = []
    default_content_types: dict[str, str] = {}
    for element in root:
        if element.tag == _CONTENT_TYPE_DEFAULT_TAG:
            extension = str(element.get("Extension") or "").casefold()
            content_type = str(element.get("ContentType") or "")
            if not extension or not content_type:
                raise NativePackageSafetyError(
                    "[Content_Types].xml has an incomplete default"
                )
            if f".{extension}" in _ACTIVE_FILE_SUFFIXES:
                raise NativePackageSafetyError(
                    "[Content_Types].xml declares an executable extension"
                )
            if extension in default_content_types:
                raise NativePackageSafetyError(
                    "[Content_Types].xml repeats a default extension"
                )
            default_content_types[extension] = content_type
        elif element.tag == _CONTENT_TYPE_OVERRIDE_TAG:
            part_name = str(element.get("PartName") or "")
            content_type = str(element.get("ContentType") or "")
            if not part_name.startswith("/") or not content_type:
                raise NativePackageSafetyError(
                    "[Content_Types].xml has an incomplete override"
                )
            if part_name == main_part:
                observed_main_types.append(content_type)
        else:
            raise NativePackageSafetyError(
                "[Content_Types].xml has an unexpected element"
            )
        folded_content_type = content_type.casefold()
        if any(
            marker in folded_content_type
            for marker in _ACTIVE_CONTENT_TYPE_MARKERS
        ):
            raise NativePackageSafetyError(
                "[Content_Types].xml declares active or macro-enabled content"
            )
    if len(observed_main_types) > 1:
        raise NativePackageSafetyError(
            "[Content_Types].xml repeats the native package main part"
        )
    effective_main_type = (
        observed_main_types[0]
        if observed_main_types
        else default_content_types.get("xml")
    )
    if effective_main_type != main_content_type:
        raise NativePackageSafetyError(
            "native package main content type does not exactly match its "
            f"{suffix} extension"
        )
    if main_part.lstrip("/") not in names:
        raise NativePackageSafetyError(
            f"native package is missing its main part {main_part!r}"
        )


@contextmanager
def _open_confined_regular_file(
    path: str | Path,
    *,
    workspace_root: str | Path,
) -> Iterator[tuple[BinaryIO, Path, os.stat_result]]:
    """Open one regular file through fd-relative, no-follow traversal.

    Each directory descriptor anchors the next lookup, so renaming or
    replacing a previously checked path component cannot redirect the final
    open. This is required because an episode may have spawned processes that
    continue mutating its workspace while host-side grading begins.
    """

    target, root, relative = _path_scope(
        path,
        workspace_root=workspace_root,
    )
    directory_flags = (
        os.O_RDONLY
        | getattr(os, "O_CLOEXEC", 0)
        | getattr(os, "O_DIRECTORY", 0)
        | getattr(os, "O_NOFOLLOW", 0)
    )
    file_flags = (
        os.O_RDONLY
        | getattr(os, "O_CLOEXEC", 0)
        | getattr(os, "O_NOFOLLOW", 0)
        | getattr(os, "O_NONBLOCK", 0)
    )
    if os.open not in os.supports_dir_fd:
        raise NativePackageSafetyError(
            "fd-relative native artifact traversal is unavailable"
        )

    root_stat = None
    root_fd = -1
    parent_fd = -1
    final_fd = -1
    handle: BinaryIO | None = None
    try:
        try:
            root_stat = root.lstat()
        except OSError as exc:
            raise NativePackageSafetyError(
                f"workspace root is unavailable: {type(exc).__name__}: {exc}"
            ) from exc
        if stat.S_ISLNK(root_stat.st_mode) or not stat.S_ISDIR(
            root_stat.st_mode
        ):
            raise NativePackageSafetyError(
                "workspace root must be a real directory, not a link or "
                "special file"
            )
        try:
            root_fd = os.open(root, directory_flags)
        except OSError as exc:
            raise NativePackageSafetyError(
                "workspace root could not be opened without following links: "
                f"{type(exc).__name__}: {exc}"
            ) from exc
        opened_root = os.fstat(root_fd)
        if (
            opened_root.st_dev != root_stat.st_dev
            or opened_root.st_ino != root_stat.st_ino
            or not stat.S_ISDIR(opened_root.st_mode)
        ):
            raise NativePackageSafetyError(
                "workspace root changed during path validation"
            )

        parent_fd = root_fd
        root_fd = -1
        for index, component in enumerate(relative.parts[:-1]):
            component_relative = Path(*relative.parts[: index + 1])
            component_stat = _entry_stat(
                component,
                parent_fd=parent_fd,
                relative=component_relative,
            )
            if stat.S_ISLNK(component_stat.st_mode):
                raise NativePackageSafetyError(
                    "native artifact path contains a symlink: "
                    f"{component_relative.as_posix()}"
                )
            if not stat.S_ISDIR(component_stat.st_mode):
                raise NativePackageSafetyError(
                    "native artifact parent is not a directory: "
                    f"{component_relative.as_posix()}"
                )
            try:
                next_fd = os.open(
                    component,
                    directory_flags,
                    dir_fd=parent_fd,
                )
            except OSError as exc:
                raise NativePackageSafetyError(
                    "native artifact parent changed or contains a symlink: "
                    f"{component_relative.as_posix()}: "
                    f"{type(exc).__name__}: {exc}"
                ) from exc
            opened_component = os.fstat(next_fd)
            if (
                opened_component.st_dev != component_stat.st_dev
                or opened_component.st_ino != component_stat.st_ino
                or not stat.S_ISDIR(opened_component.st_mode)
            ):
                os.close(next_fd)
                raise NativePackageSafetyError(
                    "native artifact parent changed during path validation: "
                    f"{component_relative.as_posix()}"
                )
            os.close(parent_fd)
            parent_fd = next_fd

        final_name = relative.parts[-1]
        final_stat = _entry_stat(
            final_name,
            parent_fd=parent_fd,
            relative=relative,
        )
        if stat.S_ISLNK(final_stat.st_mode):
            raise NativePackageSafetyError(
                "native artifact path contains a symlink: "
                f"{relative.as_posix()}"
            )
        if not stat.S_ISREG(final_stat.st_mode):
            raise NativePackageSafetyError(
                "native artifact must be a regular file"
            )
        try:
            final_fd = os.open(
                final_name,
                file_flags,
                dir_fd=parent_fd,
            )
        except OSError as exc:
            raise NativePackageSafetyError(
                "native artifact changed or contains a symlink: "
                f"{relative.as_posix()}: {type(exc).__name__}: {exc}"
            ) from exc
        opened_stat = os.fstat(final_fd)
        if (
            not stat.S_ISREG(opened_stat.st_mode)
            or opened_stat.st_dev != final_stat.st_dev
            or opened_stat.st_ino != final_stat.st_ino
        ):
            raise NativePackageSafetyError(
                "native artifact changed during path validation"
            )
        handle = os.fdopen(final_fd, "rb")
        final_fd = -1
        yield handle, target, opened_stat
    finally:
        if handle is not None:
            handle.close()
        if final_fd >= 0:
            os.close(final_fd)
        if parent_fd >= 0:
            os.close(parent_fd)
        if root_fd >= 0:
            os.close(root_fd)


def _validated_relationship_target(
    target: str,
    *,
    relationship_part: str,
) -> str:
    decoded = target
    for _ in range(4):
        replacement = unquote(decoded)
        if replacement == decoded:
            break
        decoded = replacement
    if (
        not decoded
        or "\x00" in decoded
        or "\\" in decoded
        or any(ord(character) < 32 for character in decoded)
    ):
        raise NativePackageSafetyError(
            f"{relationship_part}: relationship target is unsafe: "
            f"{target!r}"
        )
    parsed = urlsplit(decoded)
    if parsed.scheme or parsed.netloc:
        raise NativePackageSafetyError(
            f"{relationship_part}: relationship target is external: "
            f"{target!r}"
        )
    package_path = parsed.path
    if relationship_part == "_rels/.rels":
        source_directory = ""
    else:
        marker = "/_rels/"
        if marker not in relationship_part:
            raise NativePackageSafetyError(
                f"relationship part has a noncanonical location: "
                f"{relationship_part!r}"
            )
        prefix, relationship_name = relationship_part.rsplit(marker, 1)
        if not relationship_name.endswith(".rels"):
            raise NativePackageSafetyError(
                f"relationship part has a noncanonical name: "
                f"{relationship_part!r}"
            )
        source_part = posixpath.join(
            prefix,
            relationship_name.removesuffix(".rels"),
        )
        source_directory = posixpath.dirname(source_part)
    normalized = posixpath.normpath(
        package_path.lstrip("/")
        if package_path.startswith("/")
        else posixpath.join(source_directory, package_path)
    )
    if normalized in {"", ".", ".."} or normalized.startswith("../"):
        raise NativePackageSafetyError(
            f"{relationship_part}: relationship target escapes package: "
            f"{target!r}"
        )
    return normalized


def _validate_relationship_parts(
    archive: ZipFile,
    infos: list[ZipInfo],
) -> None:
    chart_workbook_embeddings: set[str] = set()
    for info in infos:
        name = info.filename
        if info.is_dir() or not name.casefold().endswith(".rels"):
            continue
        try:
            payload = archive.read(info)
            if b"<!doctype" in payload.lower():
                raise NativePackageSafetyError(
                    f"{name}: relationship XML contains a DTD"
                )
            root = ElementTree.fromstring(payload)
        except NativePackageSafetyError:
            raise
        except ElementTree.ParseError as exc:
            raise NativePackageSafetyError(
                f"{name}: relationship XML is malformed: {exc}"
            ) from exc
        if root.tag != _RELATIONSHIPS_TAG:
            raise NativePackageSafetyError(
                f"{name}: relationship XML has an invalid root"
            )
        for relationship in root:
            if relationship.tag != _RELATIONSHIP_TAG:
                raise NativePackageSafetyError(
                    f"{name}: relationship XML has an unexpected element"
                )
            mode = relationship.get("TargetMode", "Internal")
            target = relationship.get("Target")
            relationship_type = relationship.get("Type")
            if not relationship_type:
                raise NativePackageSafetyError(
                    f"{name}: relationship is missing its type"
                )
            if mode.casefold() != "internal":
                raise NativePackageSafetyError(
                    f"{name}: external relationship is forbidden: "
                    f"{target!r}"
                )
            if target is None:
                raise NativePackageSafetyError(
                    f"{name}: relationship is missing its target"
                )
            normalized_target = _validated_relationship_target(
                target,
                relationship_part=name,
            )
            folded_type = relationship_type.casefold().rstrip("/")
            blocked_suffix = next(
                (
                    suffix
                    for suffix in _BLOCKED_RELATIONSHIP_TYPE_SUFFIXES
                    if folded_type.endswith(suffix)
                ),
                None,
            )
            chart_workbook = (
                blocked_suffix == "/package"
                and bool(
                    re.fullmatch(
                        r"ppt/charts/_rels/chart\d+\.xml\.rels",
                        name,
                        flags=re.IGNORECASE,
                    )
                )
                and normalized_target.casefold().startswith(
                    "ppt/embeddings/"
                )
                and normalized_target.casefold().endswith(".xlsx")
                and normalized_target.count("/") == 2
            )
            if blocked_suffix and not chart_workbook:
                raise NativePackageSafetyError(
                    f"{name}: active or embedded relationship is forbidden: "
                    f"{relationship_type!r}"
                )
            if chart_workbook:
                chart_workbook_embeddings.add(normalized_target)

    embedding_members = {
        info.filename
        for info in infos
        if (
            not info.is_dir()
            and "/embeddings/" in f"/{info.filename.casefold()}"
        )
    }
    if embedding_members != chart_workbook_embeddings:
        raise NativePackageSafetyError(
            "native package contains an unbound or ambiguously bound "
            "embedded object"
        )
    if len(chart_workbook_embeddings) > NATIVE_PACKAGE_MAX_CHART_EMBEDDINGS:
        raise NativePackageSafetyError(
            "native package chart workbook embedding count exceeds "
            f"{NATIVE_PACKAGE_MAX_CHART_EMBEDDINGS}"
        )
    nested_expanded_bytes = 0
    for name in sorted(chart_workbook_embeddings):
        payload = archive.read(name)
        opened_stat = os.stat_result(
            (
                stat.S_IFREG | 0o600,
                0,
                0,
                1,
                0,
                0,
                len(payload),
                0,
                0,
                0,
            )
        )
        inspection = _inspect_open_native_package(
            io.BytesIO(payload),
            target=Path(posixpath.basename(name)),
            opened_stat=opened_stat,
        )
        nested_expanded_bytes += inspection.expanded_bytes
        if (
            nested_expanded_bytes
            > NATIVE_PACKAGE_MAX_EMBEDDED_EXPANDED_BYTES
        ):
            raise NativePackageSafetyError(
                "native package embedded chart workbook expansion exceeds "
                f"{NATIVE_PACKAGE_MAX_EMBEDDED_EXPANDED_BYTES}"
            )


def _validate_xlsx_formula_links(
    archive: ZipFile,
    infos: list[ZipInfo],
    *,
    suffix: str,
) -> None:
    if suffix != ".xlsx":
        return
    for info in infos:
        name = info.filename
        folded = name.casefold()
        if (
            info.is_dir()
            or not folded.startswith("xl/")
            or not folded.endswith(".xml")
        ):
            continue
        try:
            payload = archive.read(info)
            if b"<!doctype" in payload.lower():
                raise NativePackageSafetyError(
                    f"{name}: workbook XML contains a DTD"
                )
            root = ElementTree.fromstring(payload)
        except NativePackageSafetyError:
            raise
        except ElementTree.ParseError as exc:
            raise NativePackageSafetyError(
                f"{name}: workbook XML is malformed: {exc}"
            ) from exc
        for element in root.iter():
            if _local_name(element.tag) not in {"f", "definedName"}:
                continue
            formula = "".join(element.itertext())
            if (
                _EXTERNAL_FORMULA_LINK.search(formula)
                or _has_active_formula(formula)
            ):
                raise NativePackageSafetyError(
                    f"{name}: external or active workbook formula is forbidden"
                )


def _inspect_open_native_package(
    handle: BinaryIO,
    *,
    target: Path,
    opened_stat: os.stat_result,
) -> NativePackageInspection:
    if target.suffix.casefold() not in NATIVE_PACKAGE_SUFFIXES:
        raise NativePackageSafetyError(
            f"unsupported native artifact suffix: {target.suffix!r}"
        )
    if opened_stat.st_size > NATIVE_PACKAGE_MAX_FILE_BYTES:
        raise NativePackageSafetyError(
            f"native package file size {opened_stat.st_size} exceeds "
            f"{NATIVE_PACKAGE_MAX_FILE_BYTES}"
        )
    handle.seek(0)
    try:
        with ZipFile(handle) as archive:
            infos = archive.infolist()
            if len(infos) > NATIVE_PACKAGE_MAX_MEMBERS:
                raise NativePackageSafetyError(
                    f"native package member count {len(infos)} exceeds "
                    f"{NATIVE_PACKAGE_MAX_MEMBERS}"
                )

            names: set[str] = set()
            casefolded_names: set[str] = set()
            expanded_bytes = 0
            largest_member = 0
            maximum_ratio = 0.0
            for info in infos:
                name = info.filename
                if not _canonical_member_name(
                    name,
                    directory=info.is_dir(),
                ):
                    raise NativePackageSafetyError(
                        f"native package member name is noncanonical: "
                        f"{name!r}"
                    )
                folded = name.casefold()
                if name in names or folded in casefolded_names:
                    raise NativePackageSafetyError(
                        f"native package member name is duplicated: "
                        f"{name!r}"
                    )
                names.add(name)
                casefolded_names.add(folded)
                if _member_is_symlink(info):
                    raise NativePackageSafetyError(
                        f"native package contains a symlink member: "
                        f"{name!r}"
                    )
                if info.flag_bits & 0x1:
                    raise NativePackageSafetyError(
                        f"native package contains an encrypted member: "
                        f"{name!r}"
                    )
                if (
                    info.compress_type
                    not in NATIVE_PACKAGE_ALLOWED_COMPRESSION
                ):
                    raise NativePackageSafetyError(
                        f"native package member uses unsupported "
                        f"compression: {name!r}"
                    )
                if (
                    info.file_size < 0
                    or info.compress_size < 0
                    or info.file_size > NATIVE_PACKAGE_MAX_MEMBER_BYTES
                ):
                    raise NativePackageSafetyError(
                        f"native package member size is invalid: "
                        f"{name!r}"
                    )
                expanded_bytes += info.file_size
                if expanded_bytes > NATIVE_PACKAGE_MAX_EXPANDED_BYTES:
                    raise NativePackageSafetyError(
                        f"native package expanded size exceeds "
                        f"{NATIVE_PACKAGE_MAX_EXPANDED_BYTES}"
                    )
                largest_member = max(
                    largest_member,
                    info.file_size,
                )
                ratio = (
                    info.file_size / max(info.compress_size, 1)
                    if info.file_size
                    else 0.0
                )
                maximum_ratio = max(maximum_ratio, ratio)
                if (
                    info.file_size
                    >= NATIVE_PACKAGE_RATIO_MIN_MEMBER_BYTES
                    and ratio > NATIVE_PACKAGE_MAX_COMPRESSION_RATIO
                ):
                    raise NativePackageSafetyError(
                        f"native package member compression ratio "
                        f"{ratio:.2f} exceeds "
                        f"{NATIVE_PACKAGE_MAX_COMPRESSION_RATIO}: "
                        f"{name!r}"
                    )

            streamed_total = 0
            for info in infos:
                if info.is_dir():
                    continue
                streamed_member = 0
                with archive.open(info, "r") as member:
                    while True:
                        chunk = member.read(
                            NATIVE_PACKAGE_STREAM_CHUNK_BYTES
                        )
                        if not chunk:
                            break
                        streamed_member += len(chunk)
                        streamed_total += len(chunk)
                        if (
                            streamed_member
                            > NATIVE_PACKAGE_MAX_MEMBER_BYTES
                            or streamed_member > info.file_size
                            or streamed_total
                            > NATIVE_PACKAGE_MAX_EXPANDED_BYTES
                        ):
                            raise NativePackageSafetyError(
                                "native package expanded beyond declared "
                                "safety bounds"
                            )
                if streamed_member != info.file_size:
                    raise NativePackageSafetyError(
                        f"native package member length differs: "
                        f"{info.filename!r}"
                    )
            if streamed_total != expanded_bytes:
                raise NativePackageSafetyError(
                    "native package expanded length differs"
                )
            _validate_active_member_names(infos)
            _validate_all_xml_parts(archive, infos)
            _validate_content_types(
                archive,
                infos,
                suffix=target.suffix.casefold(),
            )
            _validate_relationship_parts(archive, infos)
            _validate_xlsx_formula_links(
                archive,
                infos,
                suffix=target.suffix.casefold(),
            )
    except NativePackageSafetyError:
        raise
    except (
        BadZipFile,
        EOFError,
        NotImplementedError,
        OSError,
        RuntimeError,
        ValueError,
    ) as exc:
        raise NativePackageSafetyError(
            f"native package validation failed: "
            f"{type(exc).__name__}: {exc}"
        ) from exc
    return NativePackageInspection(
        path=target.as_posix(),
        file_bytes=opened_stat.st_size,
        members=len(infos),
        expanded_bytes=expanded_bytes,
        largest_member_bytes=largest_member,
        maximum_compression_ratio=maximum_ratio,
    )


def _canonical_member_name(name: str, *, directory: bool) -> bool:
    if (
        not name
        or "\x00" in name
        or "\\" in name
        or name.startswith("/")
        or re.match(r"^[A-Za-z]:", name)
        or any(ord(character) < 32 for character in name)
    ):
        return False
    if directory != name.endswith("/"):
        return False
    body = name[:-1] if directory else name
    if not body:
        return False
    parts = body.split("/")
    return all(part not in {"", ".", ".."} for part in parts)


def _member_is_symlink(info: ZipInfo) -> bool:
    unix_mode = (info.external_attr >> 16) & 0xFFFF
    return stat.S_IFMT(unix_mode) == stat.S_IFLNK


def inspect_native_package(
    path: str | Path,
    *,
    workspace_root: str | Path | None = None,
) -> NativePackageInspection:
    """Fail closed unless an OOXML artifact satisfies all safety bounds."""

    candidate = _lexical_absolute(path)
    root = (
        _lexical_absolute(workspace_root)
        if workspace_root is not None
        else candidate.parent
    )
    with _open_confined_regular_file(
        candidate,
        workspace_root=root,
    ) as (handle, target, opened_stat):
        return _inspect_open_native_package(
            handle,
            target=target,
            opened_stat=opened_stat,
        )


@contextmanager
def validated_native_package_copy(
    path: str | Path,
    *,
    workspace_root: str | Path | None = None,
) -> Iterator[tuple[Path, NativePackageInspection]]:
    """Yield a private, immutable snapshot validated from the same safe fd.

    Downstream libraries and LibreOffice must consume this path instead of
    reopening the mutable episode workspace path after preflight.
    """

    candidate = _lexical_absolute(path)
    root = (
        _lexical_absolute(workspace_root)
        if workspace_root is not None
        else candidate.parent
    )
    with _open_confined_regular_file(
        candidate,
        workspace_root=root,
    ) as (source, target, opened_stat):
        inspection = _inspect_open_native_package(
            source,
            target=target,
            opened_stat=opened_stat,
        )
        source.seek(0)
        with tempfile.TemporaryDirectory(
            prefix="pinehaven-native-snapshot-"
        ) as directory:
            snapshot_root = Path(directory)
            snapshot = snapshot_root / target.name
            flags = (
                os.O_WRONLY
                | os.O_CREAT
                | os.O_EXCL
                | getattr(os, "O_CLOEXEC", 0)
                | getattr(os, "O_NOFOLLOW", 0)
            )
            descriptor = os.open(snapshot, flags, 0o600)
            copied = 0
            try:
                with os.fdopen(descriptor, "wb") as destination:
                    descriptor = -1
                    while True:
                        chunk = source.read(
                            NATIVE_PACKAGE_STREAM_CHUNK_BYTES
                        )
                        if not chunk:
                            break
                        copied += len(chunk)
                        if copied > NATIVE_PACKAGE_MAX_FILE_BYTES:
                            raise NativePackageSafetyError(
                                "native artifact grew while its private "
                                "snapshot was created"
                            )
                        destination.write(chunk)
                    destination.flush()
                    os.fsync(destination.fileno())
            finally:
                if descriptor >= 0:
                    os.close(descriptor)
            after_copy = os.fstat(source.fileno())
            if (
                copied != opened_stat.st_size
                or after_copy.st_size != opened_stat.st_size
                or after_copy.st_mtime_ns != opened_stat.st_mtime_ns
                or after_copy.st_ctime_ns != opened_stat.st_ctime_ns
            ):
                raise NativePackageSafetyError(
                    "native artifact changed while its private snapshot was "
                    "created"
                )
            snapshot_inspection = inspect_native_package(
                snapshot,
                workspace_root=snapshot_root,
            )
            if (
                snapshot_inspection.file_bytes != inspection.file_bytes
                or snapshot_inspection.members != inspection.members
                or snapshot_inspection.expanded_bytes
                != inspection.expanded_bytes
            ):
                raise NativePackageSafetyError(
                    "native artifact snapshot differs from validated source"
                )
            yield snapshot, inspection


def _sha256_handle(handle: BinaryIO) -> str:
    digest = hashlib.sha256()
    handle.seek(0)
    for chunk in iter(
        lambda: handle.read(NATIVE_PACKAGE_STREAM_CHUNK_BYTES),
        b"",
    ):
        digest.update(chunk)
    handle.seek(0)
    return digest.hexdigest()


def confined_regular_file_sha256(
    path: str | Path,
    *,
    workspace_root: str | Path,
    max_bytes: int,
) -> tuple[str, int]:
    """Hash a bounded regular file from the same fd that was confined."""

    if max_bytes < 0:
        raise ValueError("max_bytes must be nonnegative")
    with _open_confined_regular_file(
        path,
        workspace_root=workspace_root,
    ) as (handle, _target, opened_stat):
        if opened_stat.st_size > max_bytes:
            raise NativePackageSafetyError(
                f"regular file size {opened_stat.st_size} exceeds {max_bytes}"
            )
        digest = hashlib.sha256()
        consumed = 0
        while True:
            chunk = handle.read(NATIVE_PACKAGE_STREAM_CHUNK_BYTES)
            if not chunk:
                break
            consumed += len(chunk)
            if consumed > max_bytes or consumed > opened_stat.st_size:
                raise NativePackageSafetyError(
                    "regular file grew while it was hashed"
                )
            digest.update(chunk)
        after_hash = os.fstat(handle.fileno())
        if (
            consumed != opened_stat.st_size
            or after_hash.st_size != opened_stat.st_size
            or after_hash.st_mtime_ns != opened_stat.st_mtime_ns
            or after_hash.st_ctime_ns != opened_stat.st_ctime_ns
        ):
            raise NativePackageSafetyError(
                "regular file changed while it was hashed"
            )
        return digest.hexdigest(), consumed


def confined_regular_file_contains(
    path: str | Path,
    *,
    workspace_root: str | Path,
    needle: bytes,
    max_bytes: int,
) -> bool:
    """Search a bounded confined file without reopening its mutable path."""

    if not needle:
        raise ValueError("needle must not be empty")
    with _open_confined_regular_file(
        path,
        workspace_root=workspace_root,
    ) as (handle, _target, opened_stat):
        if opened_stat.st_size > max_bytes:
            raise NativePackageSafetyError(
                f"regular file size {opened_stat.st_size} exceeds {max_bytes}"
            )
        consumed = 0
        overlap = b""
        found = False
        while True:
            chunk = handle.read(NATIVE_PACKAGE_STREAM_CHUNK_BYTES)
            if not chunk:
                break
            consumed += len(chunk)
            if consumed > max_bytes or consumed > opened_stat.st_size:
                raise NativePackageSafetyError(
                    "regular file grew while it was searched"
                )
            if needle in overlap + chunk:
                found = True
            overlap_size = len(needle) - 1
            overlap = (
                (overlap + chunk)[-overlap_size:]
                if overlap_size
                else b""
            )
        after_search = os.fstat(handle.fileno())
        if (
            consumed != opened_stat.st_size
            or after_search.st_size != opened_stat.st_size
            or after_search.st_mtime_ns != opened_stat.st_mtime_ns
            or after_search.st_ctime_ns != opened_stat.st_ctime_ns
        ):
            raise NativePackageSafetyError(
                "regular file changed while it was searched"
            )
        return found


@contextmanager
def _open_confined_parent_directory(
    path: str | Path,
    *,
    workspace_root: str | Path,
) -> Iterator[tuple[int, Path, str]]:
    """Anchor the target parent directory for an fd-relative atomic replace."""

    target, root, relative = _path_scope(
        path,
        workspace_root=workspace_root,
    )
    directory_flags = (
        os.O_RDONLY
        | getattr(os, "O_CLOEXEC", 0)
        | getattr(os, "O_DIRECTORY", 0)
        | getattr(os, "O_NOFOLLOW", 0)
    )
    if os.open not in os.supports_dir_fd:
        raise NativePackageSafetyError(
            "fd-relative native artifact traversal is unavailable"
        )
    root_fd = -1
    parent_fd = -1
    try:
        root_stat = root.lstat()
        if stat.S_ISLNK(root_stat.st_mode) or not stat.S_ISDIR(
            root_stat.st_mode
        ):
            raise NativePackageSafetyError(
                "workspace root must be a real directory, not a link or "
                "special file"
            )
        root_fd = os.open(root, directory_flags)
        opened_root = os.fstat(root_fd)
        if (
            opened_root.st_dev != root_stat.st_dev
            or opened_root.st_ino != root_stat.st_ino
            or not stat.S_ISDIR(opened_root.st_mode)
        ):
            raise NativePackageSafetyError(
                "workspace root changed during path validation"
            )
        parent_fd = root_fd
        root_fd = -1
        for index, component in enumerate(relative.parts[:-1]):
            component_relative = Path(*relative.parts[: index + 1])
            component_stat = _entry_stat(
                component,
                parent_fd=parent_fd,
                relative=component_relative,
            )
            if stat.S_ISLNK(component_stat.st_mode):
                raise NativePackageSafetyError(
                    "native artifact path contains a symlink: "
                    f"{component_relative.as_posix()}"
                )
            if not stat.S_ISDIR(component_stat.st_mode):
                raise NativePackageSafetyError(
                    "native artifact parent is not a directory: "
                    f"{component_relative.as_posix()}"
                )
            next_fd = os.open(
                component,
                directory_flags,
                dir_fd=parent_fd,
            )
            opened_component = os.fstat(next_fd)
            if (
                opened_component.st_dev != component_stat.st_dev
                or opened_component.st_ino != component_stat.st_ino
                or not stat.S_ISDIR(opened_component.st_mode)
            ):
                os.close(next_fd)
                raise NativePackageSafetyError(
                    "native artifact parent changed during path validation: "
                    f"{component_relative.as_posix()}"
                )
            os.close(parent_fd)
            parent_fd = next_fd
        yield parent_fd, target, relative.parts[-1]
    except NativePackageSafetyError:
        raise
    except OSError as exc:
        raise NativePackageSafetyError(
            "native artifact parent could not be opened safely: "
            f"{type(exc).__name__}: {exc}"
        ) from exc
    finally:
        if parent_fd >= 0:
            os.close(parent_fd)
        if root_fd >= 0:
            os.close(root_fd)


def _replace_native_package_if_unchanged(
    target: Path,
    replacement: Path,
    *,
    workspace_root: Path,
    expected_sha256: str,
) -> None:
    """Atomically install validated bytes without following mutable paths."""

    replacement_inspection = inspect_native_package(
        replacement,
        workspace_root=replacement.parent,
    )
    with _open_confined_parent_directory(
        target,
        workspace_root=workspace_root,
    ) as (parent_fd, _target, final_name):
        current_stat = _entry_stat(
            final_name,
            parent_fd=parent_fd,
            relative=Path(final_name),
        )
        if stat.S_ISLNK(current_stat.st_mode):
            raise NativePackageSafetyError(
                "native artifact changed to a symlink before sanitization"
            )
        if not stat.S_ISREG(current_stat.st_mode):
            raise NativePackageSafetyError(
                "native artifact changed to a non-regular file before "
                "sanitization"
            )
        read_flags = (
            os.O_RDONLY
            | getattr(os, "O_CLOEXEC", 0)
            | getattr(os, "O_NOFOLLOW", 0)
            | getattr(os, "O_NONBLOCK", 0)
        )
        current_fd = os.open(
            final_name,
            read_flags,
            dir_fd=parent_fd,
        )
        with os.fdopen(current_fd, "rb") as current:
            opened_current = os.fstat(current.fileno())
            if (
                opened_current.st_dev != current_stat.st_dev
                or opened_current.st_ino != current_stat.st_ino
                or _sha256_handle(current) != expected_sha256
            ):
                raise NativePackageSafetyError(
                    "native artifact changed before sanitization could be "
                    "installed"
                )

        temporary_name = (
            f".{Path(final_name).stem}.pinehaven-"
            f"{secrets.token_hex(12)}.tmp{target.suffix.casefold()}"
        )
        write_flags = (
            os.O_RDWR
            | os.O_CREAT
            | os.O_EXCL
            | getattr(os, "O_CLOEXEC", 0)
            | getattr(os, "O_NOFOLLOW", 0)
        )
        temporary_fd = -1
        try:
            temporary_fd = os.open(
                temporary_name,
                write_flags,
                current_stat.st_mode & 0o777,
                dir_fd=parent_fd,
            )
            copied = 0
            with replacement.open("rb") as source:
                while True:
                    chunk = source.read(NATIVE_PACKAGE_STREAM_CHUNK_BYTES)
                    if not chunk:
                        break
                    copied += len(chunk)
                    if copied > NATIVE_PACKAGE_MAX_FILE_BYTES:
                        raise NativePackageSafetyError(
                            "sanitized native package exceeds file-size limit"
                        )
                    remaining = memoryview(chunk)
                    while remaining:
                        written = os.write(temporary_fd, remaining)
                        if written <= 0:
                            raise NativePackageSafetyError(
                                "sanitized native package write made no "
                                "progress"
                            )
                        remaining = remaining[written:]
            if copied != replacement_inspection.file_bytes:
                raise NativePackageSafetyError(
                    "sanitized native package copy length differs"
                )
            os.fsync(temporary_fd)
            os.lseek(temporary_fd, 0, os.SEEK_SET)
            with os.fdopen(os.dup(temporary_fd), "rb") as staged:
                _inspect_open_native_package(
                    staged,
                    target=target.with_name(temporary_name),
                    opened_stat=os.fstat(temporary_fd),
                )
            os.close(temporary_fd)
            temporary_fd = -1
            os.replace(
                temporary_name,
                final_name,
                src_dir_fd=parent_fd,
                dst_dir_fd=parent_fd,
            )
            os.fsync(parent_fd)
        finally:
            if temporary_fd >= 0:
                os.close(temporary_fd)
            try:
                os.unlink(temporary_name, dir_fd=parent_fd)
            except FileNotFoundError:
                pass


_AXIS_ID = re.compile(
    rb"(<(?:[A-Za-z_][\w.-]*:)?(?:axId|crossAx)\b[^>]*\bval=[\"'])(-?\d+)([\"'])"
)
_INT32_MIN = -(2**31)
_UINT32_MAX = 2**32 - 1
_AXIS_ID_REMAP_START = 100_000_001
_AXIS_DEFINITION_TAGS = {"catAx", "dateAx", "serAx", "valAx"}
_AXIS_PLOT_TAGS = {
    "area3DChart",
    "areaChart",
    "bar3DChart",
    "barChart",
    "bubbleChart",
    "line3DChart",
    "lineChart",
    "radarChart",
    "scatterChart",
    "stockChart",
    "surface3DChart",
    "surfaceChart",
}


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _axis_value(
    element: ElementTree.Element,
    *,
    part_name: str,
    issues: list[str],
    allow_signed_int32: bool = False,
) -> int | None:
    kind = _local_name(element.tag)
    raw = element.get("val")
    if raw is None:
        issues.append(f"{part_name}: {kind} is missing its val attribute")
        return None
    pattern = r"-?\d+" if allow_signed_int32 else r"\d+"
    if not re.fullmatch(pattern, raw):
        issues.append(
            f"{part_name}: {kind} value {raw!r} is not an unsigned integer"
        )
        return None
    value = int(raw)
    minimum = _INT32_MIN if allow_signed_int32 else 0
    if value < minimum or value > _UINT32_MAX:
        issues.append(
            f"{part_name}: {kind} value {value} is outside "
            f"{'signed-Int32/UInt32 repair range' if allow_signed_int32 else 'UInt32'}"
        )
        return None
    return value


def _chart_axis_issues(
    root: ElementTree.Element,
    *,
    part_name: str,
    allow_signed_int32: bool = False,
) -> tuple[list[str], dict[ElementTree.Element, int | None]]:
    issues: list[str] = []
    parents = {
        child: parent
        for parent in root.iter()
        for child in parent
    }
    axis_elements = [
        element
        for element in root.iter()
        if _local_name(element.tag) in {"axId", "crossAx"}
    ]
    values = {
        element: _axis_value(
            element,
            part_name=part_name,
            issues=issues,
            allow_signed_int32=allow_signed_int32,
        )
        for element in axis_elements
    }

    definitions: dict[int, str] = {}
    for axis in (
        element
        for element in root.iter()
        if _local_name(element.tag) in _AXIS_DEFINITION_TAGS
    ):
        axis_ids = [
            child
            for child in axis
            if _local_name(child.tag) == "axId"
        ]
        cross_axes = [
            child
            for child in axis
            if _local_name(child.tag) == "crossAx"
        ]
        axis_kind = _local_name(axis.tag)
        if len(axis_ids) != 1:
            issues.append(
                f"{part_name}: {axis_kind} defines {len(axis_ids)} "
                "axId elements; exactly one is required"
            )
        if len(cross_axes) != 1:
            issues.append(
                f"{part_name}: {axis_kind} defines {len(cross_axes)} "
                "crossAx elements; exactly one is required"
            )
        if len(axis_ids) == 1 and values.get(axis_ids[0]) is not None:
            value = values[axis_ids[0]]
            assert value is not None
            if value in definitions:
                issues.append(
                    f"{part_name}: duplicate defined axis id {value} "
                    f"in {definitions[value]} and {axis_kind}"
                )
            else:
                definitions[value] = axis_kind

    for element in axis_elements:
        value = values.get(element)
        if value is None:
            continue
        parent = parents.get(element)
        parent_kind = _local_name(parent.tag) if parent is not None else ""
        element_kind = _local_name(element.tag)
        is_definition = (
            element_kind == "axId"
            and parent_kind in _AXIS_DEFINITION_TAGS
        )
        if not is_definition and value not in definitions:
            issues.append(
                f"{part_name}: dangling {element_kind} reference "
                f"{value} from {parent_kind or 'unknown'}"
            )

    for plot in (
        element
        for element in root.iter()
        if _local_name(element.tag) in _AXIS_PLOT_TAGS
    ):
        references = [
            child
            for child in plot
            if _local_name(child.tag) == "axId"
        ]
        if len(references) < 2:
            issues.append(
                f"{part_name}: {_local_name(plot.tag)} has "
                f"{len(references)} axis references; at least two "
                "are required"
            )
    return issues, values


def pptx_axis_id_issues(
    path: str | Path,
    *,
    workspace_root: str | Path | None = None,
) -> list[str]:
    """Return malformed, duplicate, or dangling chart-axis references."""

    target = Path(path)
    issues: list[str] = []
    try:
        with validated_native_package_copy(
            target,
            workspace_root=workspace_root,
        ) as (snapshot, _inspection):
            with ZipFile(snapshot) as archive:
                for name in archive.namelist():
                    if not (
                        name.casefold().startswith("ppt/charts/")
                        and name.casefold().endswith(".xml")
                    ):
                        continue
                    try:
                        root = ElementTree.fromstring(archive.read(name))
                    except ElementTree.ParseError as exc:
                        issues.append(
                            f"{name}: chart XML is malformed: {exc}"
                        )
                        continue
                    chart_issues, _ = _chart_axis_issues(
                        root,
                        part_name=name,
                    )
                    issues.extend(chart_issues)
    except (OSError, BadZipFile, NativePackageSafetyError) as exc:
        issues.append(f"PPTX package could not be inspected: {exc}")
    return issues


def sanitize_pptx_axis_ids(
    path: str | Path,
    *,
    workspace_root: str | Path | None = None,
) -> bool:
    """Normalize signed IDs emitted by python-pptx to low positive values.

    Some python-pptx versions serialize their random signed 32-bit chart IDs
    directly into OOXML fields whose schema is unsigned. LibreOffice is
    lenient, while strict importers reject the whole deck. Assigning each
    structurally valid signed ID a deterministic, unused positive value keeps
    every definition/reference equivalence without emitting high UInt32 IDs
    that alter rendering in some consumers.
    """

    target = _lexical_absolute(path)
    inspection_root = _lexical_absolute(workspace_root or target.parent)
    with validated_native_package_copy(
        target,
        workspace_root=inspection_root,
    ) as (snapshot, _inspection):
        with snapshot.open("rb") as snapshot_handle:
            expected_sha256 = _sha256_handle(snapshot_handle)
        with ZipFile(snapshot) as source:
            payloads: list[tuple[ZipInfo, bytes]] = []
            replacements = 0
            for info in source.infolist():
                payload = source.read(info.filename)
                if not (
                    info.filename.casefold().startswith("ppt/charts/")
                    and info.filename.casefold().endswith(".xml")
                ):
                    payloads.append((info, payload))
                    continue
                try:
                    root = ElementTree.fromstring(payload)
                except ElementTree.ParseError:
                    # Do not partly repair a package containing arbitrary
                    # malformed chart XML. The strict validator must see the
                    # original defect.
                    return False
                graph_issues, values = _chart_axis_issues(
                    root,
                    part_name=info.filename,
                    allow_signed_int32=True,
                )
                if graph_issues:
                    # A signed integer is the only defect this compatibility
                    # repair recognizes.
                    return False

                used = {
                    value
                    for value in values.values()
                    if value is not None and value >= 0
                }
                signed_ids: list[int] = []
                for value in values.values():
                    if (
                        value is not None
                        and value < 0
                        and value not in signed_ids
                    ):
                        signed_ids.append(value)
                mapping: dict[int, int] = {}
                candidate = _AXIS_ID_REMAP_START
                for signed_id in signed_ids:
                    while candidate in used:
                        candidate += 1
                    mapping[signed_id] = candidate
                    used.add(candidate)
                    candidate += 1

                expected_replacements = sum(
                    value in mapping
                    for value in values.values()
                    if value is not None
                )
                chart_replacements = 0

                def replace(match: re.Match[bytes]) -> bytes:
                    nonlocal chart_replacements
                    value = int(match.group(2))
                    normalized = mapping.get(value)
                    if normalized is None:
                        return match.group(0)
                    chart_replacements += 1
                    return (
                        match.group(1)
                        + str(normalized).encode()
                        + match.group(3)
                    )

                rewritten = _AXIS_ID.sub(replace, payload)
                if chart_replacements != expected_replacements:
                    raise ValueError(
                        f"{info.filename}: could not rewrite every signed axis "
                        "definition/reference"
                    )
                if chart_replacements:
                    rewritten_root = ElementTree.fromstring(rewritten)
                    rewritten_issues, _ = _chart_axis_issues(
                        rewritten_root,
                        part_name=info.filename,
                    )
                    if rewritten_issues:
                        raise ValueError("; ".join(rewritten_issues))
                replacements += chart_replacements
                payloads.append((info, rewritten))
        if not replacements:
            return False

        rewritten_package = snapshot.with_name("sanitized.pptx")
        with ZipFile(
            rewritten_package,
            "w",
            ZIP_DEFLATED,
        ) as destination:
            for info, payload in payloads:
                destination.writestr(info, payload)
        inspect_native_package(
            rewritten_package,
            workspace_root=snapshot.parent,
        )
        rewritten_issues = pptx_axis_id_issues(
            rewritten_package,
            workspace_root=snapshot.parent,
        )
        if rewritten_issues:
            raise ValueError("; ".join(rewritten_issues))
        _replace_native_package_if_unchanged(
            target,
            rewritten_package,
            workspace_root=inspection_root,
            expected_sha256=expected_sha256,
        )

    inspect_native_package(
        target,
        workspace_root=inspection_root,
    )
    issues = pptx_axis_id_issues(
        target,
        workspace_root=inspection_root,
    )
    if issues:
        raise ValueError("; ".join(issues))
    return True
