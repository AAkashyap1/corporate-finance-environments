"""Fail-closed agent workspace isolation and rollout trace evidence."""

from __future__ import annotations

import hashlib
import json
import os
import re
import shlex
import stat
import subprocess
import sys
from pathlib import Path
from typing import Any, Mapping

from hud.environment import Mount, Workspace


WORKSPACE_ISOLATION_POLICY_VERSION = "pinehaven-workspace-v4-2026-07-31"
STARTUP_ATTESTATION_SCHEMA_VERSION = 1
BOUNDARY_BLOCK_MARKER = "PINEHAVEN_WORKSPACE_BOUNDARY_BLOCKED"
REQUIRE_ISOLATION_ENV = "PINEHAVEN_REQUIRE_SHELL_ISOLATION"
AGENT_VENV = "/opt/pinehaven-agent-venv"
AGENT_RUNTIME_SOURCE_ENV = "PINEHAVEN_AGENT_RUNTIME_ROOT"
AGENT_RUNTIME_MANIFEST_ENV = "PINEHAVEN_AGENT_RUNTIME_MANIFEST"
AGENT_RUNTIME_REQUIREMENTS_ENV = "PINEHAVEN_AGENT_RUNTIME_REQUIREMENTS"
DEFAULT_AGENT_RUNTIME_RELATIVE = Path(".build") / "agent-runtime"
DEFAULT_AGENT_RUNTIME_MANIFEST_RELATIVE = (
    Path(".build") / "agent-runtime-manifest.json"
)
DEFAULT_AGENT_RUNTIME_REQUIREMENTS_RELATIVE = (
    Path(".build") / "agent-runtime-requirements.txt"
)
_PROJECT_ROOT = Path(__file__).resolve().parents[1]
_VERIFIED_STARTUP_ATTESTATION: dict[str, Any] | None = None
_FORBIDDEN_AGENT_RUNTIME_ENTRIES = frozenset(
    {
        "env.py",
        "tasks.py",
        "task_catalog.py",
        "task_templates.py",
        "task_blueprints.py",
        "evaluator",
        "pinehaven_company",
    }
)
_FORBIDDEN_AGENT_RUNTIME_NAME_PARTS = (
    "pinehaven_manufacturing_finance_company",
    "__editable___pinehaven",
    "__editable__.pinehaven",
)
_FORBIDDEN_AGENT_IMPORT_MODULES = (
    "env",
    "tasks",
    "task_catalog",
    "task_templates",
    "task_blueprints",
    "evaluator",
    "pinehaven_company",
)
_AGENT_RUNTIME_TEXT_SUFFIXES = frozenset(
    {
        "",
        ".cfg",
        ".json",
        ".pth",
        ".py",
        ".txt",
    }
)

_FORBIDDEN_COMMAND_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("parent-directory traversal", re.compile(r"(^|[\s'\"=])\.\.(?:/|\\)")),
    ("raw ERP database access", re.compile(
        r"(?i)\b(?:PINEHAVEN_ERP_DB|pinehaven_erp\.db|sqlite3)\b"
    )),
    ("private evaluator or reference access", re.compile(
        r"(?i)(?:evaluator[\\/]+(?:reference|gold)|gold_builder|private[_ -]?gold)"
    )),
    ("environment implementation access", re.compile(
        r"(?i)\b(?:task_catalog|task_blueprints|task_templates)\.py\b"
    )),
    ("runtime state access", re.compile(r"(?i)(?:\.runtime[\\/]+state|COMPANY_STATE_ROOT)")),
    ("Git metadata access", re.compile(
        r"(?i)(?:^|[\s;&|])git(?:\s|$)|(?:^|[\\/])\.git(?:[\\/]|$)"
    )),
    ("host runner access", re.compile(
        r"(?i)(?:/home/runner/work|/environment(?:/|$)|/root(?:/|$))"
    )),
    ("root filesystem enumeration", re.compile(
        r"(?i)(?:^|[\s;&|])(?:find|ls|du|tree)\s+/(?:\s|$)"
    )),
)


def isolation_required() -> bool:
    return os.environ.get(REQUIRE_ISOLATION_ENV, "").casefold() in {
        "1",
        "true",
        "yes",
        "on",
    }


def agent_runtime_source() -> Path:
    configured = os.environ.get(AGENT_RUNTIME_SOURCE_ENV)
    if configured:
        return Path(os.path.abspath(configured))
    return _PROJECT_ROOT / DEFAULT_AGENT_RUNTIME_RELATIVE


def _bounded_file_bytes_and_sha256(
    path: Path,
    *,
    max_bytes: int,
) -> tuple[bytes, str]:
    try:
        path_stat = path.lstat()
    except OSError as exc:
        raise RuntimeError(
            f"Attested dependency file is unavailable: {path.name}: {exc}"
        ) from exc
    if (
        stat.S_ISLNK(path_stat.st_mode)
        or not stat.S_ISREG(path_stat.st_mode)
        or path_stat.st_size > max_bytes
    ):
        raise RuntimeError(
            f"Attested dependency file is unsafe: {path.name}"
        )
    flags = os.O_RDONLY | getattr(os, "O_CLOEXEC", 0)
    if hasattr(os, "O_NOFOLLOW"):
        flags |= os.O_NOFOLLOW
    try:
        descriptor = os.open(path, flags)
    except OSError as exc:
        raise RuntimeError(
            f"Attested dependency file cannot be opened safely: "
            f"{path.name}: {exc}"
        ) from exc
    try:
        opened = os.fstat(descriptor)
        if (
            not stat.S_ISREG(opened.st_mode)
            or opened.st_dev != path_stat.st_dev
            or opened.st_ino != path_stat.st_ino
            or opened.st_size != path_stat.st_size
        ):
            raise RuntimeError(
                f"Attested dependency file changed before open: {path.name}"
            )
        chunks: list[bytes] = []
        consumed = 0
        while True:
            chunk = os.read(descriptor, min(1024 * 1024, max_bytes + 1))
            if not chunk:
                break
            consumed += len(chunk)
            if consumed > max_bytes or consumed > path_stat.st_size:
                raise RuntimeError(
                    f"Attested dependency file grew: {path.name}"
                )
            chunks.append(chunk)
        after = os.fstat(descriptor)
    finally:
        os.close(descriptor)
    if (
        consumed != path_stat.st_size
        or after.st_size != path_stat.st_size
        or after.st_mtime_ns != path_stat.st_mtime_ns
        or after.st_ctime_ns != path_stat.st_ctime_ns
    ):
        raise RuntimeError(
            f"Attested dependency file changed: {path.name}"
        )
    payload = b"".join(chunks)
    return payload, hashlib.sha256(payload).hexdigest()


def _bounded_sha256(path: Path, *, max_bytes: int) -> str:
    _, digest = _bounded_file_bytes_and_sha256(path, max_bytes=max_bytes)
    return digest


def _manifest_relative_path(
    value: Any,
    *,
    field: str,
) -> Path:
    if not isinstance(value, str) or not value:
        raise RuntimeError(
            f"Dependency runtime manifest {field} must be a relative path"
        )
    relative = Path(value)
    if (
        relative.is_absolute()
        or relative == Path(".")
        or any(part in {"", ".", ".."} for part in relative.parts)
    ):
        raise RuntimeError(
            f"Dependency runtime manifest {field} is unsafe"
        )
    return relative


def _strict_manifest_object(
    pairs: list[tuple[str, Any]],
) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise RuntimeError(
                "Dependency runtime manifest contains a duplicate "
                f"JSON key: {key}"
            )
        result[key] = value
    return result


def dependency_runtime_attestation(
    *,
    project_root: Path,
) -> dict[str, Any]:
    runtime = agent_runtime_source()
    manifest_path = Path(
        os.environ.get(
            AGENT_RUNTIME_MANIFEST_ENV,
            project_root / DEFAULT_AGENT_RUNTIME_MANIFEST_RELATIVE,
        )
    )
    requirements_path = Path(
        os.environ.get(
            AGENT_RUNTIME_REQUIREMENTS_ENV,
            project_root / DEFAULT_AGENT_RUNTIME_REQUIREMENTS_RELATIVE,
        )
    )
    lock_path = project_root / "uv.lock"
    runtime_proof = validate_agent_runtime(
        runtime,
        project_root=project_root,
        base_runtime=sys.base_prefix,
    )
    manifest_bytes, manifest_sha256 = _bounded_file_bytes_and_sha256(
        manifest_path,
        max_bytes=4 * 1024 * 1024,
    )
    requirements_sha256 = _bounded_sha256(
        requirements_path,
        max_bytes=16 * 1024 * 1024,
    )
    lock_sha256 = _bounded_sha256(
        lock_path,
        max_bytes=32 * 1024 * 1024,
    )
    workspace_policy_sha256 = _bounded_sha256(
        Path(__file__).resolve(),
        max_bytes=4 * 1024 * 1024,
    )
    try:
        manifest = json.loads(
            manifest_bytes.decode("utf-8"),
            object_pairs_hook=_strict_manifest_object,
        )
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise RuntimeError(
            f"Dependency runtime manifest is malformed: {exc}"
        ) from exc
    if not isinstance(manifest, dict) or set(manifest) != {
        "editable_installs_allowed",
        "filesystem_proof",
        "import_proof",
        "lock_sha256",
        "project_installed",
        "requirements_relative",
        "requirements_sha256",
        "runtime_relative",
        "schema_version",
    }:
        raise RuntimeError(
            "Dependency runtime manifest schema differs from the pinned "
            "release contract"
        )
    runtime_relative = _manifest_relative_path(
        manifest["runtime_relative"],
        field="runtime_relative",
    )
    requirements_relative = _manifest_relative_path(
        manifest["requirements_relative"],
        field="requirements_relative",
    )
    expected_runtime = Path(os.path.abspath(project_root / runtime_relative))
    expected_requirements = Path(
        os.path.abspath(project_root / requirements_relative)
    )
    filesystem_proof = manifest.get("filesystem_proof")
    import_proof = manifest.get("import_proof")
    if (
        manifest.get("schema_version") != 1
        or expected_runtime != Path(os.path.abspath(runtime))
        or expected_requirements != Path(os.path.abspath(requirements_path))
        or manifest.get("lock_sha256") != lock_sha256
        or manifest.get("requirements_sha256") != requirements_sha256
        or manifest.get("project_installed") is not False
        or manifest.get("editable_installs_allowed") is not False
        or not isinstance(filesystem_proof, dict)
        or filesystem_proof.get("private_project_modules_present") is not False
        or filesystem_proof.get("escaping_links_present") is not False
        or not isinstance(import_proof, dict)
        or any(
            name not in import_proof.get("forbidden_modules", [])
            for name in _FORBIDDEN_AGENT_IMPORT_MODULES
        )
        or runtime_proof.get("private_project_modules_present") is not False
        or runtime_proof.get("escaping_links_present") is not False
    ):
        raise RuntimeError(
            "Dependency runtime manifest evidence failed validation"
        )
    return {
        "dependency_runtime_manifest_sha256": manifest_sha256,
        "lock_sha256": lock_sha256,
        "requirements_sha256": requirements_sha256,
        "workspace_policy_sha256": workspace_policy_sha256,
        "runtime_manifest_schema_version": 1,
        "private_project_modules_absent": True,
    }


def startup_isolation_attestation() -> dict[str, Any]:
    """Return the process-local proof set only by a successful live probe."""

    if _VERIFIED_STARTUP_ATTESTATION is None:
        return {
            "schema_version": STARTUP_ATTESTATION_SCHEMA_VERSION,
            "policy_version": WORKSPACE_ISOLATION_POLICY_VERSION,
            "required": isolation_required(),
            "verified": False,
            "private_project_modules_absent": False,
            "dependency_runtime_manifest_sha256": None,
            "lock_sha256": None,
            "requirements_sha256": None,
            "workspace_policy_sha256": None,
            "runtime_manifest_schema_version": None,
        }
    return json.loads(json.dumps(_VERIFIED_STARTUP_ATTESTATION))


def validate_agent_runtime(
    runtime_root: str | Path,
    *,
    project_root: str | Path = _PROJECT_ROOT,
    base_runtime: str | Path | None = None,
) -> dict[str, Any]:
    """Reject project code, escaping links, and editable hooks in the agent venv."""

    runtime = Path(os.path.abspath(os.fspath(runtime_root)))
    project = Path(os.path.abspath(os.fspath(project_root)))
    base = Path(
        os.path.abspath(
            os.fspath(base_runtime)
            if base_runtime is not None
            else sys.base_prefix
        )
    )
    if runtime.is_symlink():
        raise RuntimeError(f"Agent runtime root is a symlink: {runtime}")
    try:
        runtime_stat = runtime.stat(follow_symlinks=False)
    except OSError as exc:
        raise RuntimeError(
            f"Agent runtime is unavailable: {runtime}: {type(exc).__name__}: {exc}"
        ) from exc
    if not stat.S_ISDIR(runtime_stat.st_mode):
        raise RuntimeError(f"Agent runtime root is not a directory: {runtime}")

    site_packages = sorted(runtime.glob("lib/python*/site-packages"))
    if len(site_packages) != 1 or not site_packages[0].is_dir():
        raise RuntimeError(
            "Agent runtime must contain exactly one Python site-packages "
            f"directory; found={site_packages!r}"
        )
    site = site_packages[0]
    forbidden_paths = [
        site / relative for relative in _FORBIDDEN_AGENT_RUNTIME_ENTRIES
    ]
    forbidden_paths.extend(
        path
        for path in site.iterdir()
        if any(
            marker in path.name.casefold()
            for marker in _FORBIDDEN_AGENT_RUNTIME_NAME_PARTS
        )
    )
    present = sorted(
        path.relative_to(runtime).as_posix()
        for path in forbidden_paths
        if path.exists() or path.is_symlink()
    )
    if present:
        raise RuntimeError(
            "Agent runtime contains private Pinehaven modules: "
            f"{present!r}"
        )

    checked_files = 0
    checked_links = 0
    for path in runtime.rglob("*"):
        try:
            entry_stat = path.stat(follow_symlinks=False)
        except OSError as exc:
            raise RuntimeError(
                "Agent runtime entry cannot be inspected: "
                f"{path}: {type(exc).__name__}: {exc}"
            ) from exc
        if stat.S_ISLNK(entry_stat.st_mode):
            checked_links += 1
            try:
                resolved = path.resolve(strict=True)
            except OSError as exc:
                raise RuntimeError(
                    f"Agent runtime contains a broken link: {path}: {exc}"
                ) from exc
            if not (
                resolved.is_relative_to(runtime)
                or resolved.is_relative_to(base)
            ):
                raise RuntimeError(
                    "Agent runtime link escapes the approved runtimes: "
                    f"{path} -> {resolved}"
                )
            continue
        if stat.S_ISDIR(entry_stat.st_mode):
            continue
        if not stat.S_ISREG(entry_stat.st_mode):
            raise RuntimeError(
                f"Agent runtime contains a special entry: {path}"
            )
        checked_files += 1
        if path.suffix.casefold() not in _AGENT_RUNTIME_TEXT_SUFFIXES:
            continue
        if entry_stat.st_size > 4 * 1024 * 1024:
            continue
        try:
            text = path.read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError):
            continue
        lowered = text.casefold()
        project_text = re.escape(str(project).casefold().rstrip("/\\"))
        if re.search(
            rf"(?<![a-z0-9_.-]){project_text}"
            r"(?=$|[/\\\s'\",;:)\]}])",
            lowered,
        ):
            raise RuntimeError(
                "Agent runtime text references the private project root: "
                f"{path}"
            )
        if path.suffix.casefold() == ".pth" and (
            "pinehaven" in lowered or "__editable__" in lowered
        ):
            raise RuntimeError(
                "Agent runtime contains a Pinehaven/editable path hook: "
                f"{path}"
            )

    return {
        "runtime_root": str(runtime),
        "site_packages": str(site),
        "checked_files": checked_files,
        "checked_links": checked_links,
        "private_project_modules_present": False,
        "escaping_links_present": False,
    }


def agent_runtime_mounts() -> tuple[Mount, ...]:
    """Expose only the frozen Python runtime needed to author native artifacts."""

    mounts: list[Mount] = []
    base = Path(sys.base_prefix).resolve()
    if not base.is_relative_to("/usr"):
        mounts.append(Mount("ro", src=str(base), dst=str(base)))
    environment = agent_runtime_source()
    if environment.exists() or environment.is_symlink():
        # The release startup attestation performs the complete recursive
        # runtime validation before the HUD control channel is bound.  Doing
        # the same scan while declaring this read-only mount would duplicate
        # that work and can exceed the platform's startup window.
        mounts.append(Mount("ro", src=str(environment), dst=AGENT_VENV))
    elif isolation_required():
        raise RuntimeError(
            "Release isolation requires a dependency-only agent runtime; run "
            "scripts/prepare_agent_runtime.py before loading env.py."
        )
    return tuple(mounts)


def safe_agent_environment() -> dict[str, str]:
    return {
        # Native Office/font runtimes eagerly create profiles and caches.
        # Keep those runtime side effects in the bwrap-private tmpfs instead
        # of allowing them to appear as task workspace mutations.
        "HOME": "/tmp/pinehaven-agent-home",
        "PATH": (
            f"{AGENT_VENV}/bin:"
            "/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin"
        ),
        "PYTHONNOUSERSITE": "1",
        "TMPDIR": "/tmp",
        "VIRTUAL_ENV": AGENT_VENV,
        "XDG_CACHE_HOME": "/tmp/pinehaven-agent-cache",
        "XDG_CONFIG_HOME": "/tmp/pinehaven-agent-config",
    }


class IsolatedWorkspace(Workspace):
    """Workspace whose bwrap child receives an allowlisted environment only."""

    def shell_argv(
        self,
        command: str | None = None,
        *,
        cwd: str | None = None,
        env: Mapping[str, str] | None = None,
    ) -> list[str]:
        if (
            sys.platform != "win32"
            and self.bwrap_available
            and not self._drops_privileges()
        ):
            inner: list[str] = (
                ["bash", "-c", command] if command is not None else ["bash"]
            )
            child_env = {**self.env, **(env or {})}
            return self.bwrap_argv(
                inner,
                cwd=cwd,
                env=child_env,
                inherit_host_env=False,
            )
        argv = super().shell_argv(command, cwd=cwd, env=env)
        if not self.bwrap_available and self._drops_privileges():
            # Workspace's non-bwrap fallback uses a login shell, whose
            # /etc/profile replaces the dependency-only PATH supplied above.
            # Preserve the already sanitized environment for the dropped UID.
            try:
                bash_index = argv.index("bash")
            except ValueError:
                return argv
            if command is not None and argv[bash_index + 1 : bash_index + 2] == ["-lc"]:
                argv[bash_index + 1] = "-c"
            elif command is None and argv[bash_index + 1 : bash_index + 2] == ["-l"]:
                del argv[bash_index + 1]
        return argv


def verify_workspace_isolation(
    workspace: Workspace,
    *,
    project_root: Path,
) -> dict[str, Any]:
    """Exercise the real agent argv before any paid model request can run."""

    global _VERIFIED_STARTUP_ATTESTATION
    _VERIFIED_STARTUP_ATTESTATION = None
    if not isolation_required():
        attestation = {
            "schema_version": STARTUP_ATTESTATION_SCHEMA_VERSION,
            "policy_version": WORKSPACE_ISOLATION_POLICY_VERSION,
            "required": False,
            "verified": False,
            "private_project_modules_absent": False,
            "dependency_runtime_manifest_sha256": None,
            "lock_sha256": None,
            "requirements_sha256": None,
            "workspace_policy_sha256": None,
            "runtime_manifest_schema_version": None,
        }
        _VERIFIED_STARTUP_ATTESTATION = attestation
        return json.loads(json.dumps(attestation))
    if not workspace.bwrap_available and not workspace._drops_privileges():
        raise RuntimeError(
            "Release evaluation requires either bubblewrap or a non-root "
            "shell UID enforced with setpriv."
        )
    dependency_attestation = dependency_runtime_attestation(
        project_root=project_root,
    )

    hidden_root = shlex.quote(str(project_root.resolve()))
    if workspace.bwrap_available:
        boundary_checks = (
            f"test ! -e {hidden_root}",
            "test ! -e /state",
        )
    else:
        boundary_checks = (
            'test "$(id -u)" != 0',
            f"test ! -r {hidden_root}/env.py",
            f"test ! -r {hidden_root}/evaluator",
            "test ! -r /state",
        )
    command = " && ".join(
        (
            'test "$(pwd -P)" = /workspace',
            *boundary_checks,
            'test -z "${OPENAI_API_KEY+x}"',
            'test -z "${HUD_API_KEY+x}"',
            'test -z "${PINEHAVEN_ERP_DB+x}"',
            (
                "python -c "
                + shlex.quote(
                    "import importlib.util; "
                    "import docx, openpyxl, pptx; "
                    "forbidden=('env','tasks','task_catalog','task_templates',"
                    "'task_blueprints','evaluator','pinehaven_company'); "
                    "assert all(importlib.util.find_spec(name) is None "
                    "for name in forbidden); "
                    "print('pinehaven-python-tooling-ok')"
                )
            ),
            "(command -v libreoffice >/dev/null || command -v soffice >/dev/null)",
            "touch .pinehaven-isolation-probe",
            "rm .pinehaven-isolation-probe",
            "printf pinehaven-isolation-ok",
        )
    )
    try:
        completed = subprocess.run(
            workspace.shell_argv(command),
            cwd=workspace.root,
            capture_output=True,
            text=True,
            timeout=45,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise RuntimeError(
            f"Agent workspace isolation probe failed: {type(exc).__name__}: {exc}"
        ) from exc
    if completed.returncode != 0 or completed.stdout != (
        "pinehaven-python-tooling-ok\npinehaven-isolation-ok"
    ):
        raise RuntimeError(
            "Agent workspace isolation probe failed closed: "
            f"returncode={completed.returncode}; "
            f"stdout={completed.stdout!r}; stderr={completed.stderr!r}"
        )
    attestation = {
        "schema_version": STARTUP_ATTESTATION_SCHEMA_VERSION,
        "policy_version": WORKSPACE_ISOLATION_POLICY_VERSION,
        "required": True,
        "verified": True,
        **dependency_attestation,
    }
    _VERIFIED_STARTUP_ATTESTATION = attestation
    return json.loads(json.dumps(attestation))


def shell_boundary_violations(arguments: Any) -> list[str]:
    """Return explicit hidden-boundary violations in a shell tool request."""

    if not isinstance(arguments, dict):
        return []
    commands = arguments.get("commands")
    if isinstance(commands, str):
        command_rows = [commands]
    elif isinstance(commands, list):
        command_rows = [str(command) for command in commands]
    else:
        return []
    violations: set[str] = set()
    for command in command_rows:
        for label, pattern in _FORBIDDEN_COMMAND_PATTERNS:
            if pattern.search(command):
                violations.add(label)
    return sorted(violations)


def blocked_shell_result(violations: list[str]) -> str:
    return (
        f"{BOUNDARY_BLOCK_MARKER}: use only /workspace and the named ERP tools; "
        f"blocked={', '.join(violations)}"
    )


def _tool_result_text(step: dict[str, Any]) -> str:
    result = step.get("result")
    if result is None:
        return ""
    return json.dumps(result, default=str, sort_keys=True)


def assess_trace_boundary(
    trace: Mapping[str, Any],
    *,
    project_root: Path,
    startup_attestation: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Record blocked attempts and prove that no host path reached tool output."""

    blocked_calls: list[dict[str, Any]] = []
    host_path_exposures: list[int] = []
    root_text = str(project_root.resolve())
    runtime_state = str((project_root / ".runtime" / "state").resolve())
    steps = trace.get("steps")
    if not isinstance(steps, list):
        steps = []
    for step in steps:
        if not isinstance(step, dict):
            continue
        if step.get("source") == "agent":
            for call in step.get("tool_calls") or []:
                if not isinstance(call, dict) or call.get("name") != "shell":
                    continue
                violations = shell_boundary_violations(call.get("arguments"))
                if violations:
                    blocked_calls.append(
                        {
                            "step_id": step.get("step_id"),
                            "call_id": call.get("id"),
                            "violations": violations,
                        }
                    )
        if step.get("source") == "tool":
            call = step.get("call")
            if not isinstance(call, dict) or call.get("name") != "shell":
                continue
            result_text = _tool_result_text(step)
            if root_text in result_text or runtime_state in result_text:
                host_path_exposures.append(int(step.get("step_id") or 0))

    required = isolation_required()
    attestation = dict(
        startup_attestation
        if startup_attestation is not None
        else startup_isolation_attestation()
    )
    expected_attestation_fields = {
        "schema_version",
        "policy_version",
        "required",
        "verified",
        "private_project_modules_absent",
        "dependency_runtime_manifest_sha256",
        "lock_sha256",
        "requirements_sha256",
        "workspace_policy_sha256",
        "runtime_manifest_schema_version",
    }
    release_startup_attested = (
        set(attestation) == expected_attestation_fields
        and attestation.get("schema_version")
        == STARTUP_ATTESTATION_SCHEMA_VERSION
        and attestation.get("policy_version")
        == WORKSPACE_ISOLATION_POLICY_VERSION
        and attestation.get("required") is True
        and attestation.get("verified") is True
        and attestation.get("private_project_modules_absent") is True
        and attestation.get("runtime_manifest_schema_version") == 1
        and all(
            isinstance(attestation.get(field), str)
            and re.fullmatch(r"[0-9a-f]{64}", attestation[field])
            for field in (
                "dependency_runtime_manifest_sha256",
                "lock_sha256",
                "requirements_sha256",
                "workspace_policy_sha256",
            )
        )
    )
    return {
        "policy_version": WORKSPACE_ISOLATION_POLICY_VERSION,
        "release_isolation_required": required,
        "release_startup_attested": release_startup_attested,
        "startup_attestation": attestation,
        "blocked_shell_calls": blocked_calls,
        "host_path_exposure_steps": host_path_exposures,
        "compliant": (
            not host_path_exposures
            and (not required or release_startup_attested)
        ),
    }


__all__ = [
    "BOUNDARY_BLOCK_MARKER",
    "AGENT_RUNTIME_SOURCE_ENV",
    "AGENT_RUNTIME_MANIFEST_ENV",
    "AGENT_RUNTIME_REQUIREMENTS_ENV",
    "IsolatedWorkspace",
    "REQUIRE_ISOLATION_ENV",
    "STARTUP_ATTESTATION_SCHEMA_VERSION",
    "WORKSPACE_ISOLATION_POLICY_VERSION",
    "agent_runtime_source",
    "agent_runtime_mounts",
    "assess_trace_boundary",
    "blocked_shell_result",
    "dependency_runtime_attestation",
    "isolation_required",
    "safe_agent_environment",
    "shell_boundary_violations",
    "startup_isolation_attestation",
    "validate_agent_runtime",
    "verify_workspace_isolation",
]
