from __future__ import annotations

import copy
from typing import Any, Mapping

from evaluator.featured_rubric import apply_reward_policy, attach_default_policy


HYBRID_TASKS = frozenset(f"task_{number:03d}" for number in range(1, 101))
# Backward-compatible import for the pilot workflow.  The implementation is
# now benchmark-wide, but keeping the old name avoids breaking archived replay
# commands and already-published workflow files.
PILOT_TASKS = HYBRID_TASKS


def semantic_requirement(
    *,
    criterion_id: str,
    description: str,
    expected_facts: Mapping[str, Any] | None = None,
    artifact_type: str,
) -> str:
    """Build a narrow, self-contained binary requirement for a semantic judge.

    Objective facts are supplied to the judge only after the deterministic
    grader has verified that they actually occur in the artifact.  The judge's
    role is therefore limited to professional-language association, scenario
    binding, negation, and substantive narrative sufficiency.
    """

    facts = dict(expected_facts or {})
    requirement = (
        f"[{criterion_id}] Decide whether the {artifact_type} substantively satisfies this "
        f"rubric criterion: {description}. Be flexible about professional wording, labels, "
        "abbreviations, plurals, table layout, footnote markers, and whether units are stated "
        "once for a schedule. Be strict about which metric and scenario each value belongs to, "
        "the direction/sign of conclusions, negation, and whether the requested analysis is "
        "actually present. Do not pass keyword lists, copied rubric prose, or an unlabeled pile "
        "of numbers as substantive evidence."
    )
    if facts:
        requirement += (
            " The deterministically verified expected facts that must be professionally "
            f"associated in the artifact are: {facts!r}."
        )
    return requirement


def apply_semantic_judgments(
    result: Mapping[str, Any],
    judgments: Mapping[str, Any],
    *,
    judge_model: str,
) -> dict[str, Any]:
    """Combine semantic verdicts with hard deterministic gates.

    A semantic pass can never rescue a failed numeric, formula, structure, or
    unit gate.  Conversely, exact professional facts are not rejected merely
    because their labels differ from a hand-maintained alias list.
    """

    updated = copy.deepcopy(dict(result))
    review = updated.get("semantic_review")
    if not isinstance(review, dict):
        return updated
    review_criteria = review.get("criteria")
    if not isinstance(review_criteria, list):
        return updated

    criteria_by_id = {
        str(criterion.get("id")): criterion
        for criterion in updated.get("criteria", [])
        if isinstance(criterion, dict)
    }
    audit: list[dict[str, Any]] = []
    for item in review_criteria:
        if not isinstance(item, dict):
            continue
        criterion_id = str(item.get("criterion_id") or "")
        criterion = criteria_by_id.get(criterion_id)
        if criterion is None:
            continue
        hard_gate_met = bool(item.get("hard_gate_met"))
        raw_judgment = judgments.get(criterion_id)
        if isinstance(raw_judgment, dict):
            semantic_met = bool(raw_judgment.get("met"))
            reason = str(raw_judgment.get("reason") or "")
            response_id = raw_judgment.get("response_id")
        elif raw_judgment is None:
            semantic_met = False
            reason = "semantic judge returned no verdict"
            response_id = None
        else:
            semantic_met = bool(raw_judgment)
            reason = ""
            response_id = None
        final_met = hard_gate_met and semantic_met
        lexical_value = int(bool(criterion.get("value")))
        criterion["value"] = int(final_met)
        criterion["grading_method"] = "deterministic_gate_and_semantic_judge"
        criterion["evidence"] = (
            f"hard_gate={hard_gate_met}: {item.get('hard_gate_evidence', '')}; "
            f"semantic_judge={'MET' if semantic_met else 'UNMET'} ({judge_model}): {reason}; "
            f"legacy_lexical_match={lexical_value}"
        )
        audit.append(
            {
                "criterion_id": criterion_id,
                "hard_gate_met": hard_gate_met,
                "semantic_met": semantic_met,
                "final_met": final_met,
                "legacy_lexical_match": lexical_value,
                "reason": reason,
                "response_id": response_id,
            }
        )

    updated["semantic_review_result"] = {
        "mode": "environment_aligned_hybrid_semantic_v3",
        "judge_model": judge_model,
        "criteria": audit,
        "policy": "semantic MET is accepted only when its deterministic hard gate also passes",
    }
    task_id = str(review.get("task_id") or "unknown")
    return apply_reward_policy(task_id, attach_default_policy(updated))
