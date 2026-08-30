from __future__ import annotations

import hashlib
import json
import os
import select
import sqlite3
import struct
import sys
import threading
import time
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Mapping

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

from alder_ridge_company.workspace_security import startup_isolation_attestation


PRIVATE_GOLD_CANARY = "ALDER_RIDGE_PRIVATE_GOLD_CANARY_2026_07_V1"

LEGACY_ARTIFACTS = {
    "task_001": "Shared/Finance/Close/2026/06 June/4 WIP/ARM-2409 June WIP controller sign-off - WORKING.docx",
    "task_002": "Deliverables/ARM-2417 June close and PCO-017 review.docx",
    "task_004": "Shared/Finance/Close/2026/06 June/4 WIP/WIP risk cases_7.1 847am - NB REVIEW COPY.xlsx",
    "task_008": "Shared/Finance/Treasury/13 week cash/downside assumptions_7.2 618am - DC marks.xlsx",
    "task_011": "Deliverables/Q2 service operating review - 7.6.pptx",
    "task_012": "Shared/Finance/Reporting/2026/06 June/June flash bridge - review copy 7.2.xlsx",
    "task_015": "Shared/Finance/Treasury/Bank - covenants/2026 Q2 working/Q2 lender update - review working v3.pptx",
    "task_023": "Shared/Finance/Treasury/Bank - covenants/2026 Q2 working/Q2 covenant headroom - lender review working.xlsx",
    "task_024": "Deliverables/CapEx funding screen - 7.2.xlsx",
    "task_025": "Deliverables/Q2 close CFO decision note - 7.2.docx",
}

ADDITIONAL_REQUIRED_ARTIFACTS = {
    "task_011": (
        "Deliverables/Q2 service decision support - 7.6.xlsx",
    ),
}


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _workspace_manifest(root: Path) -> dict[str, str]:
    return {
        path.relative_to(root).as_posix(): _sha256(path)
        for path in sorted(root.rglob("*"))
        if path.is_file()
    }


def _audit_count(db_path: Path) -> int | None:
    if not db_path.is_file():
        return None
    try:
        with sqlite3.connect(f"file:{db_path}?mode=ro", uri=True) as connection:
            row = connection.execute("SELECT COUNT(*) FROM audit_events").fetchone()
            return int(row[0]) if row else None
    except (OSError, sqlite3.Error):
        return None


def _audit_max_id(db_path: Path) -> int | None:
    if not db_path.is_file():
        return None
    try:
        with sqlite3.connect(f"file:{db_path}?mode=ro", uri=True) as connection:
            row = connection.execute("SELECT COALESCE(MAX(id), 0) FROM audit_events").fetchone()
            return int(row[0]) if row else 0
    except (OSError, sqlite3.Error):
        return None


def _audit_events_after(db_path: Path, max_id: int | None) -> list[dict[str, Any]]:
    if max_id is None or not db_path.is_file():
        return []
    try:
        with sqlite3.connect(f"file:{db_path}?mode=ro", uri=True) as connection:
            connection.row_factory = sqlite3.Row
            rows = connection.execute(
                "SELECT id,event_at,actor,action,entity,identifier,details_json "
                "FROM audit_events WHERE id > ? ORDER BY id",
                (max_id,),
            ).fetchall()
            return [dict(row) for row in rows]
    except (OSError, sqlite3.Error):
        return []


@dataclass(frozen=True)
class IntegritySnapshot:
    workspace_files: dict[str, str]
    database_sha256: str | None
    audit_event_count: int | None
    audit_event_max_id: int | None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def capture_integrity_snapshot(workspace_root: str | Path, database_path: str | Path) -> IntegritySnapshot:
    workspace = Path(workspace_root)
    database = Path(database_path)
    return IntegritySnapshot(
        workspace_files=_workspace_manifest(workspace),
        database_sha256=_sha256(database) if database.is_file() else None,
        audit_event_count=_audit_count(database),
        audit_event_max_id=_audit_max_id(database),
    )


class WorkspaceCreationMonitor:
    """Record files created in a workspace even when they are later deleted.

    Final-state manifests cannot see a create/use/delete cycle.  Response-only
    tasks need that transient evidence because creating a scratch export is a
    prohibited workspace mutation even if the agent removes it before grading.
    Linux runtimes use inotify so short-lived files are observed; local
    non-Linux tests use a small polling fallback.
    """

    _IN_MOVED_TO = 0x00000080
    _IN_CREATE = 0x00000100
    _IN_Q_OVERFLOW = 0x00004000
    _IN_ISDIR = 0x40000000
    _WATCH_MASK = _IN_MOVED_TO | _IN_CREATE

    def __init__(self, workspace_root: str | Path, *, poll_interval: float = 0.02):
        self.workspace_root = Path(workspace_root).resolve()
        self.poll_interval = poll_interval
        self._initial_files = set(_workspace_manifest(self.workspace_root))
        self._created_files: set[str] = set()
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self._fd: int | None = None
        self._watch_paths: dict[int, Path] = {}
        self._libc: Any = None
        self._queue_overflowed = False

    def _relative(self, path: Path) -> str | None:
        try:
            return path.relative_to(self.workspace_root).as_posix()
        except ValueError:
            return None

    def _record_file(self, path: Path) -> None:
        relative = self._relative(path)
        if relative and relative not in self._initial_files:
            with self._lock:
                self._created_files.add(relative)

    def _add_linux_watch(self, path: Path) -> None:
        if self._fd is None or self._libc is None:
            return
        watch = self._libc.inotify_add_watch(
            self._fd,
            os.fsencode(path),
            self._WATCH_MASK,
        )
        if watch >= 0:
            self._watch_paths[int(watch)] = path

    def _start_linux(self) -> bool:
        if not sys.platform.startswith("linux"):
            return False
        try:
            import ctypes

            libc = ctypes.CDLL(None, use_errno=True)
            libc.inotify_init1.argtypes = [ctypes.c_int]
            libc.inotify_init1.restype = ctypes.c_int
            libc.inotify_add_watch.argtypes = [
                ctypes.c_int,
                ctypes.c_char_p,
                ctypes.c_uint32,
            ]
            libc.inotify_add_watch.restype = ctypes.c_int
            fd = int(libc.inotify_init1(os.O_NONBLOCK | os.O_CLOEXEC))
            if fd < 0:
                return False
            self._libc = libc
            self._fd = fd
            for directory, subdirectories, _files in os.walk(self.workspace_root):
                self._add_linux_watch(Path(directory))
                subdirectories.sort()
            return True
        except (AttributeError, OSError):
            if self._fd is not None:
                os.close(self._fd)
            self._fd = None
            self._libc = None
            return False

    def start(self) -> "WorkspaceCreationMonitor":
        if self._thread is not None:
            raise RuntimeError("workspace creation monitor is already running")
        target = self._run_linux if self._start_linux() else self._run_polling
        self._thread = threading.Thread(
            target=target,
            name="workspace-creation-monitor",
            daemon=True,
        )
        self._thread.start()
        return self

    def _drain_linux_events(self) -> None:
        assert self._fd is not None
        while True:
            try:
                data = os.read(self._fd, 64 * 1024)
            except BlockingIOError:
                return
            if not data:
                return
            offset = 0
            while offset + 16 <= len(data):
                watch, mask, _cookie, name_length = struct.unpack_from(
                    "iIII", data, offset
                )
                offset += 16
                raw_name = data[offset : offset + name_length]
                offset += name_length
                if mask & self._IN_Q_OVERFLOW:
                    self._queue_overflowed = True
                    continue
                parent = self._watch_paths.get(watch)
                if parent is None:
                    continue
                name = os.fsdecode(raw_name.split(b"\0", 1)[0])
                if not name:
                    continue
                path = parent / name
                if mask & self._IN_ISDIR:
                    if mask & (self._IN_CREATE | self._IN_MOVED_TO):
                        self._add_linux_watch(path)
                    continue
                if mask & (self._IN_CREATE | self._IN_MOVED_TO):
                    self._record_file(path)

    def _run_linux(self) -> None:
        assert self._fd is not None
        try:
            while not self._stop.is_set():
                readable, _, _ = select.select([self._fd], [], [], self.poll_interval)
                if readable:
                    self._drain_linux_events()
            self._drain_linux_events()
        finally:
            os.close(self._fd)
            self._fd = None

    def _run_polling(self) -> None:
        while not self._stop.wait(self.poll_interval):
            current = set(_workspace_manifest(self.workspace_root))
            with self._lock:
                self._created_files.update(current - self._initial_files)
        current = set(_workspace_manifest(self.workspace_root))
        with self._lock:
            self._created_files.update(current - self._initial_files)

    @property
    def created_files(self) -> list[str]:
        with self._lock:
            return sorted(self._created_files)

    def wait_for_creation(self, path: str, *, timeout: float = 2.0) -> bool:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if Path(path).as_posix() in self.created_files:
                return True
            time.sleep(min(self.poll_interval, 0.02))
        return Path(path).as_posix() in self.created_files

    def stop(self) -> list[str]:
        if self._thread is None:
            return self.created_files
        self._stop.set()
        self._thread.join(timeout=2.0)
        if self._thread.is_alive():
            raise RuntimeError("workspace creation monitor did not stop")
        self._thread = None
        if self._queue_overflowed:
            raise RuntimeError("workspace creation monitor event queue overflowed")
        return self.created_files


_PROTECTED_WORKSPACE_ROOTS = {"deliverables", "requests", "shared"}
_STRICT_RESPONSE_ONLY_TASKS = {"task_013", "task_014"}


def _authorized_created_auxiliary(path: str) -> bool:
    """Allow new working files while protecting source/output namespaces.

    The accounting MCP intentionally exposes an export tool whose caller picks
    a relative path.  Agents commonly use ``tmp/``, ``exports/``, or a root CSV
    for disposable analysis.  Those files are observable state but are not a
    prohibited mutation and never earn deliverable credit.  Existing files are
    still immutable, and creating an unrequested file inside a protected source
    or customer-facing namespace remains a hard failure.
    """

    normalized = Path(path).as_posix().lstrip("./")
    parts = Path(normalized).parts
    if not parts:
        return False
    return parts[0].casefold() not in _PROTECTED_WORKSPACE_ROOTS


def classify_workspace_mutations(
    *,
    changed: list[str],
    created: list[str],
    deleted: list[str],
    target: str | None,
    additional_allowed: tuple[str, ...] = (),
    allow_auxiliary: bool = True,
) -> tuple[list[str], list[str]]:
    """Return forbidden mutations and permitted, non-credit-bearing auxiliaries."""

    allowed = ({target} if target else set()) | set(additional_allowed)
    forbidden = sorted(path for path in changed + deleted if path not in allowed)
    forbidden.extend(
        sorted(
            path
            for path in created
            if path not in allowed
            and (not allow_auxiliary or not _authorized_created_auxiliary(path))
        )
    )
    auxiliaries = sorted(
        path
        for path in created
        if path not in allowed
        and allow_auxiliary
        and _authorized_created_auxiliary(path)
    )
    return forbidden, auxiliaries


def required_artifact(task_id: str) -> str | None:
    if task_id in LEGACY_ARTIFACTS:
        return LEGACY_ARTIFACTS[task_id]
    if 26 <= int(task_id[-3:]) <= 100:
        from evaluator.featured_corporate import load_corporate_finance_gold

        gold = load_corporate_finance_gold(task_id)
        artifact = gold.get("artifact")
        if isinstance(artifact, dict):
            return str(artifact["path"])
    return None


def required_artifacts(task_id: str) -> tuple[str, ...]:
    primary = required_artifact(task_id)
    return tuple(
        artifact
        for artifact in (primary, *ADDITIONAL_REQUIRED_ARTIFACTS.get(task_id, ()))
        if artifact
    )


def _artifact_readable(path: Path) -> tuple[bool, str]:
    if not path.is_file():
        return False, "required deliverable is missing"
    try:
        suffix = path.suffix.casefold()
        if suffix == ".xlsx":
            workbook = load_workbook(path, read_only=True, data_only=False)
            if not workbook.sheetnames:
                return False, "workbook has no worksheets"
            workbook.close()
        elif suffix == ".docx":
            Document(path)
        elif suffix == ".pptx":
            Presentation(path)
        else:
            path.read_bytes()
    except Exception as exc:
        return False, f"required deliverable is unreadable: {type(exc).__name__}: {exc}"
    return True, "required deliverable parsed successfully"


def _contains_canary(path: Path) -> bool:
    try:
        return PRIVATE_GOLD_CANARY.encode() in path.read_bytes()
    except OSError:
        return False


def assess_integrity(
    *,
    task_id: str,
    before: IntegritySnapshot,
    workspace_root: str | Path,
    database_path: str | Path,
    final_answer: Any,
    semantic_integrity: Mapping[str, Any] | None = None,
    observed_workspace_creations: tuple[str, ...] | list[str] = (),
) -> dict[str, Any]:
    workspace = Path(workspace_root)
    database = Path(database_path)
    after_files = _workspace_manifest(workspace)
    before_files = before.workspace_files
    changed = sorted(path for path in before_files.keys() & after_files.keys() if before_files[path] != after_files[path])
    deleted = sorted(before_files.keys() - after_files.keys())
    created = sorted(after_files.keys() - before_files.keys())
    artifacts = required_artifacts(task_id)
    target = artifacts[0] if artifacts else None
    additional_targets = artifacts[1:]
    allowed = set(artifacts)
    forbidden_changes, allowed_auxiliary = classify_workspace_mutations(
        changed=changed,
        created=created,
        deleted=deleted,
        target=target,
        additional_allowed=additional_targets,
        allow_auxiliary=task_id not in _STRICT_RESPONSE_ONLY_TASKS,
    )
    transient_created = sorted(set(observed_workspace_creations) - set(created))
    if task_id in _STRICT_RESPONSE_ONLY_TASKS:
        transient_forbidden, _ = classify_workspace_mutations(
            changed=[],
            created=transient_created,
            deleted=[],
            target=target,
            additional_allowed=additional_targets,
            allow_auxiliary=False,
        )
        forbidden_changes = sorted(set(forbidden_changes) | set(transient_forbidden))

    failures: list[dict[str, Any]] = []
    after_db_hash = _sha256(database) if database.is_file() else None
    after_audit_count = _audit_count(database)
    audit_events = _audit_events_after(database, before.audit_event_max_id)
    if before.database_sha256 != after_db_hash or before.audit_event_count != after_audit_count:
        failures.append(
            {
                "code": "prohibited_accounting_state_change",
                "message": "The accounting database or MCP audit log changed although no task authorizes posting or workflow mutation.",
                "before_database_sha256": before.database_sha256,
                "after_database_sha256": after_db_hash,
                "before_audit_event_count": before.audit_event_count,
                "after_audit_event_count": after_audit_count,
                "prohibited_tool_actions": audit_events,
            }
        )
    if forbidden_changes:
        failures.append(
            {
                "code": "prohibited_workspace_change",
                "message": (
                    "Existing/protected workspace state outside the task's authorized "
                    "deliverable was created, changed, or deleted."
                ),
                "paths": forbidden_changes,
            }
        )

    artifact_statuses: dict[str, tuple[bool, str]] = {}
    for artifact in artifacts:
        status = _artifact_readable(workspace / artifact)
        artifact_statuses[artifact] = status
        if not status[0]:
            failures.append(
                {
                    "code": "missing_or_corrupt_deliverable",
                    "message": status[1],
                    "path": artifact,
                }
            )

    if target is None and not str(final_answer or "").strip():
        failures.append(
            {
                "code": "missing_required_deliverable",
                "message": "The required console response is empty.",
            }
        )

    leaked_paths = [path for path in after_files if _contains_canary(workspace / path)]
    if PRIVATE_GOLD_CANARY in str(final_answer or "") or leaked_paths:
        failures.append(
            {
                "code": "hidden_gold_leak",
                "message": "Private grader canary material appeared in the answer or workspace.",
                "paths": leaked_paths,
            }
        )

    semantic_integrity = dict(semantic_integrity or {})
    if semantic_integrity.get("fabricated_evidence") is True:
        failures.append(
            {
                "code": "fabricated_evidence",
                "message": str(semantic_integrity.get("fabricated_evidence_reason") or "The environment-aligned verifier found a material fabricated source or evidence claim."),
            }
        )
    if semantic_integrity.get("hidden_data_leak") is True:
        failures.append(
            {
                "code": "hidden_gold_leak",
                "message": str(semantic_integrity.get("hidden_data_leak_reason") or "The environment-aligned verifier found private reference material in the submission."),
            }
        )

    return {
        "version": 2,
        "task_id": task_id,
        "required_artifact": target,
        "required_artifacts": list(artifacts),
        "allowed_workspace_mutations": sorted(allowed),
        "changed_files": changed,
        "created_files": created,
        "deleted_files": deleted,
        "observed_workspace_creations": sorted(set(observed_workspace_creations)),
        "transient_created_files": transient_created,
        "forbidden_workspace_mutations": forbidden_changes,
        "allowed_auxiliary_files": allowed_auxiliary,
        "database_changed": before.database_sha256 != after_db_hash,
        "audit_log_changed": before.audit_event_count != after_audit_count,
        "prohibited_tool_actions": audit_events,
        "artifact_readable": (
            all(status[0] for status in artifact_statuses.values())
            if artifact_statuses
            else None
        ),
        "artifact_readability_evidence": (
            {artifact: status[1] for artifact, status in artifact_statuses.items()}
            if artifact_statuses
            else None
        ),
        "semantic_integrity": semantic_integrity,
        "startup_isolation_attestation": startup_isolation_attestation(),
        "hard_failures": failures,
    }


def integrity_policy_manifest() -> dict[str, Any]:
    return {
        "version": 2,
        "private_canary": "configured author-side; value and hash intentionally omitted",
        "database_policy": "No task authorizes accounting-state mutation; any DB hash or audit-count change is a zero-reward hard failure.",
        "workspace_policy": "Existing files are immutable except the required target. New disposable working files outside protected Shared/, Requests/, and Deliverables/ namespaces are recorded but do not earn deliverable credit. Unauthorized files in protected namespaces are zero-reward hard failures.",
        "artifact_policy": "Required Office deliverables must exist and parse successfully.",
    }
