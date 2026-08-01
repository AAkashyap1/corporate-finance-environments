#!/usr/bin/env python3
"""Build and prove a dependency-only Python runtime for the model's shell.

The orchestration environment intentionally contains Pinehaven task,
evaluator-reference, and ERP implementation modules. Mounting it into
the model workspace would disclose hidden evaluation logic even when the
project checkout itself is hidden. This script exports only locked third-party
dependencies into a separate venv and rejects any private project module or
escaping link before release evaluation starts.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from pinehaven_company.workspace_security import validate_agent_runtime


DEFAULT_RUNTIME = ROOT / ".build" / "agent-runtime"
DEFAULT_REQUIREMENTS = ROOT / ".build" / "agent-runtime-requirements.txt"
DEFAULT_MANIFEST = ROOT / ".build" / "agent-runtime-manifest.json"
REQUIRED_AGENT_MODULES = (
    "docx",
    "numpy",
    "openpyxl",
    "pandas",
    "pdfplumber",
    "PIL",
    "pptx",
    "pypdf",
    "reportlab",
)
FORBIDDEN_AGENT_MODULES = (
    "env",
    "tasks",
    "task_catalog",
    "task_templates",
    "task_blueprints",
    "evaluator",
    "pinehaven_company",
)


def _contained_build_path(path: Path, *, description: str) -> Path:
    absolute = Path(os.path.abspath(path))
    build_root = ROOT / ".build"
    try:
        relative = absolute.relative_to(build_root)
    except ValueError as exc:
        raise SystemExit(
            f"{description} must be below {build_root}: {absolute}"
        ) from exc
    if not relative.parts:
        raise SystemExit(f"{description} must not replace {build_root}")
    current = build_root
    for component in relative.parts:
        current = current / component
        if current.is_symlink():
            raise SystemExit(
                f"{description} contains a symlink component: {current}"
            )
    return absolute


def _run(command: list[str]) -> None:
    subprocess.run(command, cwd=ROOT, check=True)


def _python_probe(runtime: Path) -> dict[str, Any]:
    interpreter = runtime / "bin" / "python"
    probe = (
        "import importlib.metadata, importlib.util, json;"
        f"required={REQUIRED_AGENT_MODULES!r};"
        f"forbidden={FORBIDDEN_AGENT_MODULES!r};"
        "missing=[name for name in required "
        "if importlib.util.find_spec(name) is None];"
        "exposed=[name for name in forbidden "
        "if importlib.util.find_spec(name) is not None];"
        "assert not missing, f'missing required modules: {missing!r}';"
        "assert not exposed, f'private modules exposed: {exposed!r}';"
        "rows=sorted((dist.metadata.get('Name') or '',dist.version) "
        "for dist in importlib.metadata.distributions());"
        "print(json.dumps(rows,separators=(',',':')))"
    )
    completed = subprocess.run(
        [str(interpreter), "-I", "-c", probe],
        cwd=ROOT,
        check=True,
        capture_output=True,
        text=True,
    )
    distributions = json.loads(completed.stdout)
    if not isinstance(distributions, list):
        raise RuntimeError("Agent runtime distribution probe is malformed")
    return {
        "required_modules": list(REQUIRED_AGENT_MODULES),
        "forbidden_modules": list(FORBIDDEN_AGENT_MODULES),
        "distributions": distributions,
    }


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(
            payload,
            indent=2,
            sort_keys=True,
            ensure_ascii=False,
            allow_nan=False,
        )
        + "\n",
        encoding="utf-8",
    )
    temporary.replace(path)


def build_agent_runtime(
    *,
    runtime: Path,
    requirements: Path,
    manifest: Path,
    uv_executable: str,
) -> dict[str, Any]:
    runtime = _contained_build_path(runtime, description="agent runtime")
    requirements = _contained_build_path(
        requirements,
        description="agent requirements",
    )
    manifest = _contained_build_path(manifest, description="agent manifest")
    runtime.parent.mkdir(parents=True, exist_ok=True)
    requirements.parent.mkdir(parents=True, exist_ok=True)
    manifest.parent.mkdir(parents=True, exist_ok=True)

    _run(
        [
            uv_executable,
            "export",
            "--frozen",
            "--no-dev",
            "--no-emit-project",
            "--no-hashes",
            "--output-file",
            str(requirements),
        ]
    )
    requirements_text = requirements.read_text(encoding="utf-8")
    active_requirements = [
        line.strip()
        for line in requirements_text.splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    ]
    if any(
        line.startswith(("-e ", "--editable "))
        or " @ file:" in line
        or line.casefold().startswith(
            "pinehaven-manufacturing-finance"
        )
        for line in active_requirements
    ):
        raise RuntimeError(
            "Dependency export unexpectedly contains the Pinehaven project"
        )

    if runtime.exists():
        if runtime.is_symlink():
            raise RuntimeError(f"Refusing to replace symlink runtime: {runtime}")
        shutil.rmtree(runtime)
    base_python = Path(sys.base_prefix) / "bin" / "python"
    _run(
        [
            uv_executable,
            "venv",
            "--relocatable",
            "--python",
            str(base_python),
            str(runtime),
        ]
    )
    _run(
        [
            uv_executable,
            "pip",
            "sync",
            "--python",
            str(runtime / "bin" / "python"),
            str(requirements),
        ]
    )

    filesystem_proof = validate_agent_runtime(
        runtime,
        project_root=ROOT,
        base_runtime=sys.base_prefix,
    )
    import_proof = _python_probe(runtime)
    payload = {
        "schema_version": 1,
        "runtime_relative": runtime.relative_to(ROOT).as_posix(),
        "requirements_relative": requirements.relative_to(ROOT).as_posix(),
        "lock_sha256": _sha256(ROOT / "uv.lock"),
        "requirements_sha256": _sha256(requirements),
        "filesystem_proof": filesystem_proof,
        "import_proof": import_proof,
        "project_installed": False,
        "editable_installs_allowed": False,
    }
    _write_json(manifest, payload)
    return payload


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--runtime", type=Path, default=DEFAULT_RUNTIME)
    parser.add_argument(
        "--requirements",
        type=Path,
        default=DEFAULT_REQUIREMENTS,
    )
    parser.add_argument("--manifest", type=Path, default=DEFAULT_MANIFEST)
    parser.add_argument(
        "--uv",
        default="uv",
        help="Pinned uv executable already installed by the release workflow.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    payload = build_agent_runtime(
        runtime=args.runtime,
        requirements=args.requirements,
        manifest=args.manifest,
        uv_executable=args.uv,
    )
    print(json.dumps(payload, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
