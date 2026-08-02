from __future__ import annotations

import os
import shutil
from pathlib import Path

from .paths import resolve_project_root, resolve_seed_root


PROJECT_ROOT = resolve_project_root(__file__)


def _path_from_env(name: str, fallback: Path) -> Path:
    raw = os.environ.get(name)
    return Path(raw).resolve() if raw else fallback.resolve()


def _handoff_workspace(root: Path) -> None:
    """Give only the agent workspace to the configured non-root shell UID."""

    raw_uid = os.environ.get("COMPANY_SHELL_UID")
    if not raw_uid:
        return
    uid = int(raw_uid)
    if uid <= 0:
        raise RuntimeError("COMPANY_SHELL_UID must be a positive integer")
    for directory, names, files in os.walk(root):
        os.chown(directory, uid, uid)
        for name in names:
            os.chown(Path(directory) / name, uid, uid, follow_symlinks=False)
        for name in files:
            os.chown(Path(directory) / name, uid, uid, follow_symlinks=False)


def reset_company() -> tuple[Path, Path]:
    """Create a clean agent workspace and mutable accounting-state copy."""

    seed_root = resolve_seed_root(__file__)
    runtime_root = _path_from_env("COMPANY_RUNTIME_ROOT", PROJECT_ROOT / ".runtime" / "workspace")
    state_root = _path_from_env("COMPANY_STATE_ROOT", PROJECT_ROOT / ".runtime" / "state")

    if not (seed_root / "accounting_seed.db").exists():
        raise FileNotFoundError("Alder Ridge accounting data is missing from the package.")

    if runtime_root.exists():
        shutil.rmtree(runtime_root)
    if state_root.exists():
        shutil.rmtree(state_root)
    runtime_root.mkdir(parents=True, exist_ok=True)
    state_root.mkdir(parents=True, exist_ok=True)
    state_root.chmod(0o700)

    shutil.copytree(seed_root / "workspace", runtime_root, dirs_exist_ok=True)
    db_path = state_root / "accounting.db"
    shutil.copy2(seed_root / "accounting_seed.db", db_path)

    for task_dir in ("Task 01 - June WIP Close", "Task 02 - 13 Week Liquidity", "Task 03 - June Operating Review", "Task 04 - Q2 Bank Compliance"):
        (runtime_root / "Deliverables" / task_dir).mkdir(parents=True, exist_ok=True)

    _handoff_workspace(runtime_root)

    return runtime_root, db_path
