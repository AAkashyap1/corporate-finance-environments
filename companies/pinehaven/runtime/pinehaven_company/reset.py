"""Deterministic episode reset for Pinehaven's workspace and mutable ERP state."""

from __future__ import annotations

import os
import shutil
from pathlib import Path

from pinehaven_company.paths import project_root, seed_root
from pinehaven_company.seed_archive import (
    ARCHIVE_NAME,
    DATABASE_NAME,
    copy_seed_database,
)
from pinehaven_company.workspace_archive import copy_seed_workspace


PROJECT_ROOT = project_root()
_BROAD_RESET_ROOTS = frozenset(
    {
        Path("/"),
        Path("/home"),
        Path("/root"),
        Path("/tmp"),
        Path("/usr"),
        Path("/var"),
        Path("/workspace"),
    }
)


def _path_from_env(name: str, fallback: Path) -> Path:
    value = os.environ.get(name)
    if value:
        configured = Path(value)
        if not configured.is_absolute():
            raise RuntimeError(f"{name} must be an absolute path")
        candidate = Path(os.path.abspath(configured))
    else:
        candidate = Path(os.path.abspath(fallback))
    current = Path(candidate.anchor)
    for component in candidate.parts[1:]:
        current = current / component
        if current.is_symlink():
            raise RuntimeError(f"{name} contains a symlink component: {current}")
    resolved = candidate.resolve(strict=False)
    if (
        resolved in _BROAD_RESET_ROOTS
        or len(resolved.parts) < 3
        or PROJECT_ROOT.is_relative_to(resolved)
        or seed_root().is_relative_to(resolved)
    ):
        raise RuntimeError(
            f"{name} is too broad or contains protected Pinehaven data: {resolved}"
        )
    if resolved.exists() and not resolved.is_dir():
        raise RuntimeError(f"{name} is not a directory: {resolved}")
    return resolved


def runtime_roots() -> tuple[Path, Path]:
    """Return distinct, non-overlapping roots that reset may safely replace."""

    runtime_root = _path_from_env(
        "COMPANY_RUNTIME_ROOT", PROJECT_ROOT / ".runtime" / "workspace"
    )
    state_root = _path_from_env(
        "COMPANY_STATE_ROOT", PROJECT_ROOT / ".runtime" / "state"
    )
    if (
        runtime_root == state_root
        or runtime_root.is_relative_to(state_root)
        or state_root.is_relative_to(runtime_root)
    ):
        raise RuntimeError(
            "COMPANY_RUNTIME_ROOT and COMPANY_STATE_ROOT must not overlap"
        )
    return runtime_root, state_root


def reset_company() -> tuple[Path, Path]:
    """Create a clean visible workspace and an isolated mutable ERP copy."""

    seed = seed_root()
    runtime_root, state_root = runtime_roots()
    seed_database = seed / DATABASE_NAME
    seed_archive = seed / ARCHIVE_NAME
    seed_workspace = seed / "workspace"
    seed_manifest = seed / "ERP_SEED_MANIFEST.json"
    if (
        not seed_database.exists()
        and not seed_archive.exists()
        and not seed_manifest.exists()
    ):
        raise FileNotFoundError("Pinehaven ERP data is missing from the package.")
    if (
        not seed_workspace.exists()
        and not (seed / "WORKSPACE_SEED_MANIFEST.json").exists()
    ):
        raise FileNotFoundError("Pinehaven source files are missing from the package.")

    for exact_root in (runtime_root, state_root):
        if exact_root.exists():
            shutil.rmtree(exact_root)
        exact_root.mkdir(parents=True, exist_ok=True)

    copy_seed_workspace(seed, runtime_root)
    database_path = state_root / "pinehaven_erp.db"
    copy_seed_database(seed, database_path)
    (runtime_root / "Deliverables").mkdir(parents=True, exist_ok=True)
    return runtime_root, database_path


__all__ = ["PROJECT_ROOT", "reset_company", "runtime_roots"]
