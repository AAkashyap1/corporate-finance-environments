from __future__ import annotations

import copy
from collections.abc import Mapping
from typing import Any


REWARD_SCHEMA_VERSION = 1
RUBRIC_SCHEMA_VERSION = 1
ALLOWED_WEIGHTS = {1, 3, 5, 10}


def apply_reward_policy(
    result: Mapping[str, Any],
    *,
    integrity: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Apply weighted binary reward, declared caps, and integrity hard fails."""

    updated = copy.deepcopy(dict(result))
    criteria = [
        row for row in updated.get("criteria", []) if isinstance(row, dict)
    ]
    total_weight = 0
    earned_weight = 0
    caps: list[tuple[float, str]] = []
    for row in criteria:
        weight = int(row.get("weight", 10))
        if weight not in ALLOWED_WEIGHTS:
            raise ValueError(f"Invalid criterion weight {weight}: {row.get('id')}")
        value = int(bool(row.get("value")))
        row["value"] = value
        row["weight"] = weight
        row.setdefault("category", "core_finance")
        row.setdefault("semantic", False)
        total_weight += weight
        earned_weight += weight * value
        if not value and row.get("failure_cap") is not None:
            caps.append((float(row["failure_cap"]), str(row.get("id"))))

    raw_reward = earned_weight / total_weight if total_weight else 0.0
    reward = raw_reward
    hard_failures: list[dict[str, Any]] = []
    if integrity is not None:
        updated["integrity"] = copy.deepcopy(dict(integrity))
        hard_failures = [
            dict(row)
            for row in integrity.get("hard_failures", [])
            if isinstance(row, Mapping)
        ]
        if hard_failures:
            reward = 0.0

    applied_caps: list[dict[str, Any]] = []
    for cap, criterion_id in sorted(caps):
        before = reward
        reward = min(reward, cap)
        if reward < before:
            applied_caps.append({"criterion_id": criterion_id, "cap": cap})

    met = sum(int(bool(row.get("value"))) for row in criteria)
    updated.update(
        {
            "reward": round(reward, 6),
            "raw_weighted_reward": round(raw_reward, 6),
            "strict_pass": bool(criteria)
            and met == len(criteria)
            and not hard_failures,
            "criteria_met": met,
            "criteria_total": len(criteria),
            "weight_earned": earned_weight,
            "weight_total": total_weight,
            "reward_schema_version": REWARD_SCHEMA_VERSION,
            "reward_definition": (
                "weighted binary criteria using 1/3/5/10 importance, followed "
                "by declared decision/auditability caps and zero-reward "
                "environment-integrity hard failures"
            ),
            "applied_reward_caps": applied_caps,
            "hard_failures": hard_failures,
        }
    )
    return updated
