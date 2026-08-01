from __future__ import annotations

import copy
import re
from typing import Any, Iterable, Mapping


RUBRIC_SCHEMA_VERSION = 3
REWARD_SCHEMA_VERSION = 3
ALLOWED_WEIGHTS = frozenset({1, 3, 5, 10})


def _slug(value: Any) -> str:
    text = re.sub(r"[^a-z0-9]+", "_", str(value or "").casefold()).strip("_")
    return text[:80] or "item"


def _criterion(
    parent: Mapping[str, Any],
    *,
    criterion_id: str,
    description: str,
    kind: str,
    category: str,
    weight: int,
    semantic: bool = False,
    failure_cap: float | None = None,
    **fields: Any,
) -> dict[str, Any]:
    if weight not in ALLOWED_WEIGHTS:
        raise ValueError(f"unsupported rubric weight: {weight}")
    row = {
        "id": criterion_id,
        "description": description,
        "kind": kind,
        "category": category,
        "weight": weight,
        "semantic": semantic,
    }
    if failure_cap is not None:
        row["failure_cap"] = float(failure_cap)
    row.update(fields)
    return row


def atomicize_criterion(spec: Mapping[str, Any]) -> list[dict[str, Any]]:
    """Expand one authored rubric row into independently scoreable requirements.

    The source builders stay concise, while the released rubric never hides
    several values, sheets, headings, or content requirements behind one bit.
    """

    parent = copy.deepcopy(dict(spec))
    criterion_id = str(parent["id"])
    description = str(parent["description"])
    kind = str(parent["kind"])

    if kind in {"numeric", "string", "boolean"}:
        expected = dict(parent["expected"])
        rows = []
        for key, value in expected.items():
            child = {
                k: v for k, v in parent.items()
                if k not in {
                    "id", "description", "expected", "kind", "category",
                    "weight", "semantic", "failure_cap",
                }
            }
            semantic = kind in {"string", "boolean"}
            rows.append(
                _criterion(
                    parent,
                    criterion_id=f"{criterion_id}__{_slug(key)}",
                    description=f"{description}: `{key}` is correct",
                    kind=kind,
                    category="decision" if semantic else "core_finance",
                    weight=10,
                    semantic=semantic,
                    failure_cap=0.49 if semantic else None,
                    expected={key: value},
                    **child,
                )
            )
        return rows

    if kind == "list":
        key = str(parent["key"])
        expected = list(parent["expected"])
        rows = []
        for index, value in enumerate(expected, start=1):
            rows.append(
                _criterion(
                    parent,
                    criterion_id=f"{criterion_id}__item_{index:02d}_{_slug(value)}",
                    description=f"{description}: required item {value!r} is correctly included and classified",
                    kind="list_item",
                    category="decision",
                    weight=10,
                    semantic=True,
                    failure_cap=0.49,
                    key=key,
                    expected_item=value,
                    unordered=bool(parent.get("unordered")),
                )
            )
        rows.append(
            _criterion(
                parent,
                criterion_id=f"{criterion_id}__complete_set",
                description=f"{description}: the submitted set has no missing, extra, or duplicated items",
                kind="list_exact",
                category="decision",
                weight=10,
                semantic=True,
                failure_cap=0.49,
                key=key,
                expected=expected,
                unordered=bool(parent.get("unordered")),
            )
        )
        return rows

    if kind in {"xlsx_label_values", "artifact_label_values"}:
        rows = []
        for label, value in dict(parent["label_values"]).items():
            decision = bool(
                re.search(
                    r"recommend|decision|compliance|compliant|breach|selected|status|required|feasible",
                    str(label),
                    flags=re.I,
                )
            )
            rows.append(
                _criterion(
                    parent,
                    criterion_id=f"{criterion_id}__{_slug(label)}",
                    description=f"{description}: `{label}` is correctly labeled and associated",
                    kind=kind,
                    category="decision" if decision else "core_finance",
                    weight=10,
                    semantic=True,
                    failure_cap=0.49 if decision else None,
                    label_values={label: value},
                )
            )
        return rows

    if kind == "artifact_tokens":
        parent_id = criterion_id.casefold()
        if "source" in parent_id:
            category, weight = "provenance", 3
        elif "control" in parent_id or "reconciliation" in parent_id:
            category, weight = "controls", 3
        elif "narrative" in parent_id or "decision" in parent_id:
            category, weight = "decision_quality", 5
        else:
            category, weight = "auditability", 3
        rows = []
        for token in parent["tokens"]:
            decision = category == "decision_quality" and bool(
                re.search(r"recommend|decision|risk|gate|action", str(token), flags=re.I)
            )
            rows.append(
                _criterion(
                    parent,
                    criterion_id=f"{criterion_id}__{_slug(token)}",
                    description=f"{description}: the artifact substantively addresses {token!r}",
                    kind=kind,
                    category=category,
                    weight=weight,
                    semantic=True,
                    failure_cap=0.49 if decision else None,
                    tokens=[token],
                )
            )
        return rows

    if kind == "xlsx_structure":
        rows = [
            _criterion(
                parent,
                criterion_id=f"{criterion_id}__sheet_{_slug(sheet)}",
                description=f"Required worksheet {sheet!r} is present",
                kind="xlsx_sheet_present",
                category="structure",
                weight=1,
                sheet=sheet,
            )
            for sheet in parent.get("sheets", [])
        ]
        rows.extend(
            _criterion(
                parent,
                criterion_id=f"{criterion_id}__preserve_{_slug(sheet)}",
                description=f"Protected source worksheet {sheet!r} remains unchanged",
                kind="xlsx_sheet_preserved",
                category="integrity",
                weight=10,
                failure_cap=0.0,
                sheet=sheet,
            )
            for sheet in parent.get("preserve_source_sheets", [])
        )
        return rows

    if kind == "xlsx_model_integrity":
        return [
            _criterion(
                parent,
                criterion_id=f"{criterion_id}__formula_graph",
                description=f"The model contains at least {int(parent['min_formulas'])} substantive formulas",
                kind="xlsx_formula_count",
                category="auditability",
                weight=5,
                failure_cap=0.69,
                min_formulas=int(parent["min_formulas"]),
            ),
            _criterion(
                parent,
                criterion_id=f"{criterion_id}__no_errors",
                description="The submitted workbook contains no spreadsheet error values",
                kind="xlsx_no_errors",
                category="integrity",
                weight=10,
                failure_cap=0.0,
            ),
        ]

    if kind == "xlsx_formula_lineage":
        rows = [
            _criterion(
                parent,
                criterion_id=f"{criterion_id}__headline_{_slug(label)}",
                description=f"Headline output {label!r} is formula-driven from workbook references",
                kind="xlsx_headline_formula",
                category="auditability",
                weight=5,
                failure_cap=0.69,
                headline_label=label,
            )
            for label in parent.get("headline_labels", [])
        ]
        rows.extend(
            _criterion(
                parent,
                criterion_id=f"{criterion_id}__formula_sheet_{_slug(sheet)}",
                description=f"Required model worksheet {sheet!r} contains formulas",
                kind="xlsx_formula_sheet",
                category="auditability",
                weight=3,
                sheet=sheet,
            )
            for sheet in parent.get("formula_sheets", [])
        )
        rows.append(
            _criterion(
                parent,
                criterion_id=f"{criterion_id}__cross_sheet",
                description=(
                    "The model contains at least "
                    f"{int(parent['min_cross_sheet_formulas'])} cross-sheet formulas"
                ),
                kind="xlsx_cross_sheet_formulas",
                category="auditability",
                weight=5,
                failure_cap=0.69,
                min_cross_sheet_formulas=int(parent["min_cross_sheet_formulas"]),
            )
        )
        return rows

    if kind == "docx_structure":
        rows = [
            _criterion(
                parent,
                criterion_id=f"{criterion_id}__heading_{_slug(heading)}",
                description=f"The memorandum contains a substantive {heading!r} section",
                kind="docx_heading",
                category="structure",
                weight=1,
                semantic=True,
                heading=heading,
            )
            for heading in parent.get("headings", [])
        ]
        rows.append(
            _criterion(
                parent,
                criterion_id=f"{criterion_id}__tables",
                description=f"The memorandum contains at least {int(parent['min_tables'])} real Word table(s)",
                kind="docx_tables",
                category="structure",
                weight=3,
                min_tables=int(parent["min_tables"]),
            )
        )
        return rows

    if kind == "pptx_structure":
        count_fields = (
            {"exact_slides": int(parent["exact_slides"])}
            if parent.get("exact_slides") is not None
            else {"min_slides": int(parent["min_slides"])}
        )
        return [
            _criterion(
                parent,
                criterion_id=f"{criterion_id}__slide_count",
                description="The presentation has the required slide count",
                kind="pptx_slide_count",
                category="structure",
                weight=1,
                **count_fields,
            ),
            _criterion(
                parent,
                criterion_id=f"{criterion_id}__title",
                description=f"The presentation contains the required title concept {parent['title_token']!r}",
                kind="pptx_title",
                category="structure",
                weight=1,
                semantic=True,
                title_token=parent["title_token"],
            ),
        ]

    # Already-atomic authored criteria retain explicit metadata when present.
    row = copy.deepcopy(parent)
    row.setdefault("category", "core_finance")
    row.setdefault("weight", 10)
    row.setdefault("semantic", kind in {"string", "boolean", "list"})
    return [row]


def atomicize_task_gold(gold: Mapping[str, Any]) -> dict[str, Any]:
    updated = copy.deepcopy(dict(gold))
    updated["rubric_schema_version"] = RUBRIC_SCHEMA_VERSION
    updated["criteria"] = [
        child
        for parent in gold.get("criteria", [])
        for child in atomicize_criterion(parent)
    ]
    ids = [str(row["id"]) for row in updated["criteria"]]
    if not ids or len(ids) != len(set(ids)):
        raise ValueError("atomic rubric ids must be non-empty and unique within a task")
    return updated


def criterion_policy(criterion: Mapping[str, Any]) -> dict[str, Any]:
    weight = int(criterion.get("weight", 10))
    if weight not in ALLOWED_WEIGHTS:
        raise ValueError(f"invalid criterion weight {weight}: {criterion.get('id')}")
    return {
        "category": str(criterion.get("category") or "core_finance"),
        "weight": weight,
        "failure_cap": (
            None if criterion.get("failure_cap") is None else float(criterion["failure_cap"])
        ),
        "semantic": bool(criterion.get("semantic")),
    }


def attach_default_policy(result: Mapping[str, Any]) -> dict[str, Any]:
    """Classify legacy hand-written criteria that predate schema-v3 gold."""

    updated = copy.deepcopy(dict(result))
    for row in updated.get("criteria", []):
        if not isinstance(row, dict) or "weight" in row:
            continue
        text = f"{row.get('id', '')} {row.get('description', '')}".casefold()
        semantic = bool(
            re.search(
                r"status|compliance|classification|timing|decision|recommend|title|required item|complete set|professional conclusion",
                text,
            )
        )
        if "preserv" in text:
            category, weight, cap = "integrity", 10, 0.0
        elif re.search(r"formula|lineage|audit", text):
            category, weight, cap = "auditability", 5, 0.69
        elif re.search(r"source|evidence", text):
            category, weight, cap = "provenance", 3, None
        elif re.search(r"chart|format|structure|slide|table|length|page|identity", text):
            category, weight, cap = "structure", 1, None
        elif semantic:
            category, weight, cap = "decision", 10, 0.49
        else:
            category, weight, cap = "core_finance", 10, None
        row.update(
            {
                "category": category,
                "weight": weight,
                "semantic": semantic,
            }
        )
        if cap is not None:
            row["failure_cap"] = cap
    return updated


def apply_reward_policy(
    task_id: str,
    result: Mapping[str, Any],
    *,
    integrity: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Compute strict pass and a weighted training reward with hard caps."""

    updated = copy.deepcopy(dict(result))
    criteria = [row for row in updated.get("criteria", []) if isinstance(row, dict)]
    total_weight = 0
    earned_weight = 0
    criterion_caps: list[tuple[float, str]] = []
    for row in criteria:
        policy = criterion_policy(row)
        row.update(policy)
        value = int(bool(row.get("value")))
        row["value"] = value
        total_weight += policy["weight"]
        earned_weight += policy["weight"] * value
        if not value and policy["failure_cap"] is not None:
            criterion_caps.append((float(policy["failure_cap"]), str(row.get("id"))))

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
    applied_caps = []
    for cap, criterion_id in sorted(criterion_caps):
        before = reward
        reward = min(reward, cap)
        if reward < before:
            applied_caps.append({"criterion_id": criterion_id, "cap": cap})

    met_count = sum(int(bool(row.get("value"))) for row in criteria)
    updated.update(
        {
            "reward": round(reward, 6),
            "raw_weighted_reward": round(raw_reward, 6),
            "strict_pass": bool(criteria) and met_count == len(criteria) and not hard_failures,
            "criteria_met": met_count,
            "criteria_total": len(criteria),
            "weight_earned": earned_weight,
            "weight_total": total_weight,
            "reward_schema_version": REWARD_SCHEMA_VERSION,
            "reward_definition": (
                "weighted binary criteria using 1/3/5/10 importance, followed by declared "
                "criterion caps and zero-reward environment-integrity hard failures"
            ),
            "applied_reward_caps": applied_caps,
            "hard_failures": hard_failures,
        }
    )
    return updated


def rubric_summary(criteria: Iterable[Mapping[str, Any]]) -> dict[str, Any]:
    rows = list(criteria)
    return {
        "criteria": len(rows),
        "weight_total": sum(int(row.get("weight", 10)) for row in rows),
        "semantic_criteria": sum(bool(row.get("semantic")) for row in rows),
        "failure_capped_criteria": sum(row.get("failure_cap") is not None for row in rows),
    }
