from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Mapping

from evaluator.task_grader import grade_task
from evaluator.integrity import assess_integrity, capture_integrity_snapshot
from evaluator.production_semantic import (
    submission_integrity_evidence,
    verify_semantic_review,
)
from evaluator.rubric import apply_reward_policy, attach_default_policy
from task_catalog import TASKS


def invalidate_on_grading_error(result: dict[str, Any]) -> dict[str, Any]:
    """Turn mandatory-verifier failure into an excluded, zero-reward result."""
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
    """Expose an IPC-safe rubric summary and governed terminal reward in HUD.

    HUD's local subprocess protocol is newline-delimited.  A fully evidenced
    190-row workbook rubric can exceed asyncio's 64 KiB default line limit, so
    the full result is persisted in the grader-only state directory and this
    transport object carries only the fields needed to verify every atomic
    outcome.  The frontier runner reads and cross-checks the sidecar before it
    writes the commercial trace.
    """
    from hud.graders import EvaluationResult

    criteria = [
        {
            key: criterion[key]
            for key in (
                "id",
                "description",
                "category",
                "weight",
                "semantic",
                "failure_cap",
                "value",
            )
            if key in criterion
        }
        for criterion in result["criteria"]
    ]
    strict_pass = bool(result["strict_pass"])
    return EvaluationResult(
        reward=float(result["reward"]),
        done=True,
        content=(
            f"{'PASS' if strict_pass else 'INCOMPLETE'} — "
            f"{result['criteria_met']}/{result['criteria_total']} atomic criteria met; "
            f"weighted reward={result['reward']:.3f}"
        ),
        info={
            "task_id": task_id,
            "strict_pass": strict_pass,
            "criteria_met": result["criteria_met"],
            "criteria_total": result["criteria_total"],
            "weight_earned": result.get("weight_earned"),
            "weight_total": result.get("weight_total"),
            "criteria": criteria,
            "pass_definition": "strict pass only when every binary criterion equals 1",
            "reward_definition": result.get("reward_definition"),
            "reward_schema_version": result.get("reward_schema_version"),
            "raw_weighted_reward": result.get("raw_weighted_reward"),
            "decision_accuracy_adjustment": result.get("decision_accuracy_adjustment"),
            "applied_reward_caps": result.get("applied_reward_caps"),
            "quality_gate_failures": result.get("quality_gate_failures"),
            "hard_failures": result.get("hard_failures"),
            "integrity": result.get("integrity"),
            "semantic_review_result": result.get("semantic_review_result"),
            "grading_error": result.get("grading_error"),
            "infrastructure_valid": result.get("infrastructure_valid", True),
        },
        # Atomic rows and weights are already present in ``info``.  Duplicating
        # 190 of them as SubScores would recreate the transport-size failure.
        subscores=[],
        isError=bool(result.get("grading_error")),
    )


def persist_grade_sidecar(
    *,
    task_id: str,
    result: dict[str, Any],
    state_root: Path,
) -> None:
    """Persist full grader evidence outside the agent-visible workspace."""

    path = state_root / "latest_grade_result.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".json.tmp")
    temporary.write_text(
        json.dumps({"task_id": task_id, "result": result}, default=str),
        encoding="utf-8",
    )
    temporary.replace(path)


def register_task_templates(
    env: Any,
    runtime_root: Path,
    state_root: Path,
    project_root: Path,
    *,
    semantic_credentials: Mapping[str, str | None],
) -> dict[str, Any]:
    """Register 100 independently runnable tasks for this company."""

    registered: dict[str, Any] = {}
    for spec in TASKS:
        task_id = spec.task_id
        prompt = spec.prompt

        @env.template(
            id=task_id,
            description=spec.title,
        )
        async def task_template(_task_id: str = task_id, _prompt: str = prompt):
            database_path = state_root / "accounting.db"
            before = capture_integrity_snapshot(runtime_root, database_path)
            answer = yield _prompt
            result = grade_task(_task_id, answer, runtime_root)
            semantic_integrity: dict[str, Any] = {}
            try:
                result, semantic_integrity = await verify_semantic_review(
                    task_id=_task_id,
                    prompt=_prompt,
                    result=result,
                    project_root=project_root,
                    hud_api_key=semantic_credentials.get("hud_api_key"),
                    openai_api_key=semantic_credentials.get("openai_api_key"),
                    submission_evidence=submission_integrity_evidence(
                        task_id=_task_id,
                        final_answer=answer,
                        workspace_root=runtime_root,
                    ),
                )
            except Exception as exc:
                # A verifier outage is infrastructure failure, never an agent
                # miss or permission to fall back to brittle lexical matching.
                result["grading_error"] = (
                    f"production semantic verifier failed: {type(exc).__name__}: {exc}"
                )
            integrity = assess_integrity(
                task_id=_task_id,
                before=before,
                workspace_root=runtime_root,
                database_path=database_path,
                final_answer=answer,
                semantic_integrity=semantic_integrity,
            )
            result = apply_reward_policy(
                _task_id,
                attach_default_policy(result),
                integrity=integrity,
            )
            # Verifier availability is part of the grader contract.  Never
            # emit a partial lexical/deterministic score as valid training
            # reward when mandatory semantic review did not complete.
            result = invalidate_on_grading_error(result)
            persist_grade_sidecar(
                task_id=_task_id,
                result=result,
                state_root=state_root,
            )
            yield _hud_result(_task_id, result)

        registered[task_id] = task_template
    return registered
