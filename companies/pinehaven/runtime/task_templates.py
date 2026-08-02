from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Mapping
from zipfile import BadZipFile

from evaluator.task_grader import grade_task
from evaluator.integrity import assess_integrity, capture_integrity_snapshot
from evaluator.native_artifacts import sanitize_pptx_axis_ids
from evaluator.production_semantic import (
    semantic_review_not_requested,
    submission_integrity_evidence,
    verify_semantic_review,
)
from evaluator.rubric import apply_reward_policy
from task_catalog import TASKS


def _invalidate_grading_error(result: dict[str, Any]) -> dict[str, Any]:
    if result.get("grading_error"):
        result["reward"] = 0.0
        result["strict_pass"] = False
        result["infrastructure_valid"] = False
        result.setdefault("hard_failures", []).append(
            {
                "code": "semantic_verifier_unavailable",
                "message": result["grading_error"],
                "reward": 0.0,
            }
        )
    else:
        result["infrastructure_valid"] = True
    return result


def _hud_result(task_id: str, result: dict[str, Any]):
    from hud.graders import EvaluationResult

    criteria = [
        {
            key: row[key]
            for key in (
                "id",
                "description",
                "category",
                "weight",
                "semantic",
                "failure_cap",
                "value",
            )
            if key in row
        }
        for row in result["criteria"]
    ]
    strict = bool(result["strict_pass"])
    return EvaluationResult(
        reward=float(result["reward"]),
        done=True,
        content=(
            f"{'PASS' if strict else 'INCOMPLETE'} — "
            f"{result['criteria_met']}/{result['criteria_total']} atomic criteria met; "
            f"weighted reward={result['reward']:.3f}"
        ),
        info={
            "task_id": task_id,
            "strict_pass": strict,
            "criteria_met": result["criteria_met"],
            "criteria_total": result["criteria_total"],
            "weight_earned": result.get("weight_earned"),
            "weight_total": result.get("weight_total"),
            "criteria": criteria,
            "pass_definition": (
                "strict pass only when every binary criterion equals 1 and "
                "environment integrity is clean"
            ),
            "reward_definition": result.get("reward_definition"),
            "reward_schema_version": result.get("reward_schema_version"),
            "raw_weighted_reward": result.get("raw_weighted_reward"),
            "applied_reward_caps": result.get("applied_reward_caps"),
            "hard_failures": result.get("hard_failures"),
            "integrity": result.get("integrity"),
            "semantic_review_result": result.get("semantic_review_result"),
            "grading_error": result.get("grading_error"),
            "infrastructure_valid": result.get("infrastructure_valid", True),
        },
        subscores=[],
        isError=bool(result.get("grading_error")),
    )


def _persist_grade(
    task_id: str, result: dict[str, Any], state_root: Path
) -> None:
    path = state_root / "latest_grade_result.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".json.tmp")
    temporary.write_text(
        json.dumps({"task_id": task_id, "result": result}, default=str),
        encoding="utf-8",
    )
    temporary.replace(path)


def _sanitize_presentation_submission(
    path: Path,
    workspace_root: Path | None = None,
) -> None:
    if not path.exists() and not path.is_symlink():
        return
    try:
        sanitize_pptx_axis_ids(
            path,
            workspace_root=workspace_root or path.parent,
        )
    except (OSError, ValueError, BadZipFile):
        # Deterministic grading reports the strict OOXML defect; a sanitizer
        # failure must not abort the task template or hide that failed grade.
        pass


def register_task_templates(
    env: Any,
    runtime_root: Path,
    state_root: Path,
    project_root: Path,
    *,
    semantic_credentials: Mapping[str, str | None],
) -> dict[str, Any]:
    registered: dict[str, Any] = {}
    for spec in TASKS:
        task_id = spec.task_id
        prompt = spec.prompt
        target_path = spec.target_path

        @env.template(id=task_id, description=spec.title)
        async def task_template(
            _task_id: str = task_id,
            _prompt: str = prompt,
            _target_path: str | None = target_path,
        ):
            database_path = state_root / "pinehaven_erp.db"
            before = capture_integrity_snapshot(
                runtime_root,
                database_path,
                task_id=_task_id,
            )
            answer = yield _prompt
            if _target_path and _target_path.casefold().endswith(".pptx"):
                _sanitize_presentation_submission(
                    runtime_root / _target_path,
                    runtime_root,
                )
            try:
                result = grade_task(
                    _task_id, answer, runtime_root, database_path
                )
            except Exception as exc:
                result = {
                    "task_id": _task_id,
                    "criteria": [],
                    "criteria_met": 0,
                    "criteria_total": 0,
                    "reward": 0.0,
                    "strict_pass": False,
                    "grading_error": (
                        f"deterministic grader failed: {type(exc).__name__}: {exc}"
                    ),
                }
            semantic_integrity: dict[str, Any] = {}
            if not result.get("grading_error"):
                try:
                    semantic_evidence = submission_integrity_evidence(
                        task_id=_task_id,
                        final_answer=answer,
                        workspace_root=runtime_root,
                    )
                except Exception as exc:
                    preparation_error = (
                        "production semantic verifier evidence failed: "
                        f"{type(exc).__name__}: {exc}"
                    )
                    result, semantic_integrity = (
                        semantic_review_not_requested(
                            result=result,
                            hud_api_key=semantic_credentials.get(
                                "hud_api_key"
                            ),
                            openai_api_key=semantic_credentials.get(
                                "openai_api_key"
                            ),
                            reason=preparation_error,
                        )
                    )
                    result["grading_error"] = preparation_error
                else:
                    result, semantic_integrity = await verify_semantic_review(
                        task_id=_task_id,
                        prompt=_prompt,
                        result=result,
                        project_root=project_root,
                        hud_api_key=semantic_credentials.get("hud_api_key"),
                        openai_api_key=semantic_credentials.get(
                            "openai_api_key"
                        ),
                        submission_evidence=semantic_evidence,
                    )
                verification_error = result.get(
                    "semantic_review_result",
                    {},
                ).get("verification_error")
                if verification_error and not result.get("grading_error"):
                    result["grading_error"] = (
                        "production semantic verifier failed: "
                        f"{verification_error}"
                    )
            else:
                result, semantic_integrity = semantic_review_not_requested(
                    result=result,
                    hud_api_key=semantic_credentials.get("hud_api_key"),
                    openai_api_key=semantic_credentials.get("openai_api_key"),
                    reason=str(result["grading_error"]),
                )
            integrity = assess_integrity(
                task_id=_task_id,
                before=before,
                workspace_root=runtime_root,
                database_path=database_path,
                final_answer=answer,
                semantic_integrity=semantic_integrity,
            )
            result = apply_reward_policy(result, integrity=integrity)
            result = _invalidate_grading_error(result)
            _persist_grade(_task_id, result, state_root)
            yield _hud_result(_task_id, result)

        registered[task_id] = task_template
    return registered
