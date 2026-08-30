from __future__ import annotations

import hashlib
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
TASK_IDS = ("task_001", "task_004", "task_015", "task_035", "task_068")


def _task(task_id: str) -> dict:
    return json.loads((ROOT / "tasks" / f"{task_id}.json").read_text(encoding="utf-8"))


def test_featured_prompts_are_objective_led() -> None:
    prohibited = (
        "return only one json object",
        "exactly these keys",
        "show the complete calculation chain",
        "populate the blue",
        "rows 6:9",
        "$185,000",
        "$1,276,325.70",
        "25%/50%/100%",
        "grader",
        "benchmark",
    )
    for task_id in TASK_IDS:
        prompt = _task(task_id)["prompt"]
        assert len(prompt) < 1_000
        lowered = prompt.casefold()
        assert not any(token.casefold() in lowered for token in prohibited)


def test_featured_targets_and_private_references_exist() -> None:
    for task_id in TASK_IDS:
        task = _task(task_id)
        target = task["output"]["target_path"]
        assert target
        assert (ROOT / "source-files" / target).is_file()
        reference = json.loads(
            (ROOT / task["evaluation"]["reference_file"]).read_text(encoding="utf-8")
        )
        assert task_id in reference


def test_selected_source_registry_matches_workspace() -> None:
    dependencies = json.loads(
        (ROOT / "data" / "control" / "task_source_dependencies.json").read_text(encoding="utf-8")
    )["tasks"]
    registry = {
        entry["path"]: entry
        for entry in json.loads(
            (ROOT / "data" / "control" / "file_registry.json").read_text(encoding="utf-8")
        )["files"]
    }
    required = {
        path
        for task_id in TASK_IDS
        for path in dependencies[task_id]["minimum_source_artifacts"]
    }
    required.add(_task("task_001")["output"]["target_path"])
    for relative in required:
        path = ROOT / "source-files" / relative
        assert path.is_file()
        assert registry[relative]["bytes"] == path.stat().st_size
        assert registry[relative]["sha256"] == hashlib.sha256(path.read_bytes()).hexdigest()


def test_current_featured_grader_revisions_are_pinned() -> None:
    apex = (ROOT / "evaluator" / "featured_apex.py").read_text(encoding="utf-8")
    expected = {
        "task-001-objective-led-controller-signoff-v23",
        "task-004-evidence-discovery-wip-review-v8",
        "task-015-objective-led-covenant-stress-v11",
        "task-035-executable-backlog-capacity-decision-v11",
        "task-068-executive-accounting-mitigation-decision-v10",
    }
    assert all(revision in apex for revision in expected)
