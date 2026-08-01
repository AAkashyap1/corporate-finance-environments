from __future__ import annotations

import csv
import hashlib
import json
import re
import zipfile
from functools import lru_cache
from pathlib import Path
from typing import Any, Mapping

from openpyxl import load_workbook
from pypdf import PdfReader

from evaluator.hybrid_semantic import apply_semantic_judgments


SEMANTIC_JUDGE_MODEL = "gpt-5.6-terra"
SEMANTIC_REASONING_EFFORT = "high"
SEMANTIC_PROMPT_VERSION = "alder-ridge-semantic-v17-2026-07-22"

SYSTEM_PROMPT = """You are the pinned production verifier for a corporate-finance RL environment.
Evaluate only the supplied semantic criteria. Deterministic code has already checked objective
numbers, dates, workbook structure, and prerequisite evidence gates; never override or weaken
those checks. In particular, when a criterion's deterministic hard gate passed, every supplied
expected fact is already present and numerically valid under the benchmark's exact tolerance and
unit rules. Spreadsheet evidence may expose the unformatted full-precision cached value behind a
cell that is displayed with ordinary professional rounding. Never redo the numeric comparison or
reject extra raw precision. Your remaining job is only to decide whether the verified fact is
professionally associated with the required metric, period, scenario, and conclusion.

Accept professionally equivalent wording, abbreviations, pluralization, reordered presentation,
footnote markers, and units stated once for a schedule. Be strict about metric-to-value and
scenario association, sign/direction, negation, decision logic, and whether an explanation is
substantive. Reject copied rubric language, keyword stuffing, and unlabeled piles of facts.
For every MET verdict involving supplied expected facts, verify and briefly identify the visible
metric/fact association in the explanation. The mere presence of the expected number or conclusion
elsewhere in the artifact is never sufficient.

Apply these criterion rules consistently:
- For headline_values criteria, MET requires the deterministically verified value to be explicitly associated with
  the named metric (or a professional equivalent) and the correct scenario. Do not calculate or
  infer an unstated ratio, percentage, total, maximum, or variance from other displayed values, and
  do not substitute a different scenario's value. Do not compare a raw full-precision cached
  spreadsheet value with the rounded expected fact; that comparison was already completed by the
  deterministic hard gate. An unambiguous selected row/flag in a labeled selection table establishes
  association for every labeled column at that row; the table does not need to repeat "selected"
  in each cell or restate the row as a separate headline. A selected ID shown above a table and the
  row bearing that same ID are an explicit association. A merely available value elsewhere may not.
  A short row label may be qualified by its adjacent definition, note, source, or table heading;
  assess the complete row rather than demanding that every qualifier be repeated in one cell. A total
  is not a maximum unless it is explicitly identified or formula-labeled as the maximum. Do not sum
  bridge components to invent an unstated headline total. Likewise, do not treat one debt, cash,
  revenue, or other component as a requested total merely because every other displayed component
  is zero; the artifact must explicitly identify the total or a professional equivalent. For a
  requested maximum or minimum, a time-series row alone is insufficient: require an explicit
  max/min summary, label, formula result, selected extreme, or equivalent presentation, and never
  scan the series yourself to choose the extreme. Formula-driven selected-project flags
  beside project IDs count as an explicit selected portfolio and need not be concatenated into one
  text cell. In an ordered period schedule, the first positive value on a clearly labeled draw row
  explicitly establishes the first draw period; scanning that row is reading the schedule, not
  inventing an unstated calculation.
- For sources__policy, MET requires an identifiable governing policy, approved definitions, or
  other explicit policy authority. A generic operating/finance schedule, an analytical practice,
  a steering decision, or a row caption containing the word "policy" is not by itself an identifiable
  policy authority. Likewise, a controlling source designation, IC decision framework, or the
  artifact's own distinction between supported and unsupported adjustments is not policy authority
  unless it identifies the governing policy or approved definitions source. A specifically named,
  dated/versioned workpaper may itself qualify when the artifact explicitly identifies that source
  as current and approved for policy; "controller-tied" alone does not make an unidentified source
  policy authority. In particular, a source-reconciliation entry that identifies a dated/versioned
  workpaper as CURRENT and expressly includes approved policy in its controlling use is present
  policy authority and MUST be MET; do not demand a redundant standalone policy PDF. A separate
  condition requiring future lender approval or future covenant definitions does not erase that
  existing internal policy authority because it governs a different approval layer. A condition
  requiring only future approval, with no present approved authority, remains insufficient. For
  sources__management_correspondence, require an identifiable email/thread,
  review note, follow-up, or a dated/participant-qualified correspondence reference. Generic phrases
  such as "management decision," "IC decisions," or a person's "warning" without a source-identifying
  context are insufficient.
- For model_content__model_status or controls__model_status, accept an explicit substantive
  aggregate conclusion labeled model status, overall status, control status, feasibility, model
  integrity, or a professional equivalent. A clearly labeled Checks/control section whose substantive
  status rows collectively establish the overall condition qualifies and MUST NOT be rejected merely
  because it lacks a separate cell saying "overall model status." Accurate EXCEPTION/FAIL rows and
  stated remediation are a substantive status just as PASS rows are; the criterion asks whether the
  artifact addresses model status, not whether the model is necessarily clean. INFO/INPUT rows do not
  invalidate an otherwise complete status section. Source-level version statuses alone are not an
  overall model status. Conversely, a displayed PASS/OK is not sufficient when visible formulas or
  labels show that the aggregate references the wrong rows, omits displayed controls, or a
  scenario-labeled output references another scenario. Mark that case UNMET. Do not invent a conflict
  that the artifact explicitly resolves through a labeled reconciliation or adjustment; evaluate the
  reconciled amount. A narrow checklist returning OK does not establish overall model status when the
  same artifact visibly contains a materially wrong required headline, conclusion, period/scenario
  association, or formula-driven decision output that the checklist omits. In that case the aggregate
  OK is not substantively reliable and is UNMET, even if its COUNTIF or equivalent correctly counts
  every row inside the narrow checklist range.
- For decision_logic__owner, require an accountable person, role, or function tied to the stated
  action or decision gate; the generic word "management" alone is not an owner. First locate an active
  accountability assignment in the artifact: for example an Owner/Responsible/Action by field, a named
  role that will perform the action, or an explicit statement that a person, role, function, board, or
  committee is responsible for approving, certifying, releasing, or signing off the specific gate. A
  role actively assigned that responsibility qualifies even when there is no literal "Owner" column,
  and an accountable role/function is sufficient without an individual's name. In contrast, a statement
  that someone must obtain another party's approval, that an action is subject to validation, that a
  decision remains pending certification, or that approval is required before proceeding describes a
  condition or approval dependency. It does not assign an owner unless the artifact separately and
  expressly states who is responsible for executing or owning that gate. The party whose approval,
  validation, or certification is needed is not thereby the owner; never combine several passive
  conditions to infer an unstated owner. A named board or committee qualifies only when the artifact
  actively assigns it responsibility for the specific decision gate. For
  decision_logic__timing, require action timing, an event-based release milestone, a gate deadline,
  or a decision cadence. Reporting, source, or data-cutoff dates alone are not action timing.
- For a requested aggregate branch-EBITDA variance to a current comparator, a forecast-versus-current
  segment-plan table that explicitly labels total EBITDA and variance, together with an adjacent or
  appendix statement identifying that segment plan as the current comparator, is a professional
  association. Do not reject it merely because the slide title says consolidated forecast or because
  a separate LTM operating-economics table is explicitly disclosed as not being the plan-variance
  basis. Still reject a lone branch component, an unlabeled bridge total, or the wrong comparator.
- For decision_logic__headroom, a quantified measure is sufficient but not always necessary. When
  the controlling source legitimately leaves the measure unavailable, an explicit disclosure of
  that unavailability, its reason, the resulting limitation, and the appropriate no-conclusion or
  follow-up action substantively addresses headroom. Do not reward a bare "N/A" or an unsupported
  compliance claim.
- For controls__cash_roll_forward, require a beginning-to-ending cash reconciliation (opening cash
  plus/minus operating, investing, and financing movements equals ending cash) or a professional
  equivalent. Operating cash flow less capex is a free-cash-flow bridge, not a cash roll-forward.
- A required complete set must contain every required item and no extra item.

Also perform two integrity checks. The environment context distinguishes required/minimum sources
from the complete visible workspace registry. Required/minimum sources are not an exhaustive
allowlist. Mark fabricated_evidence true only with high confidence that the submission materially
claims reliance on, approval by, or support from a separate source that is absent from the complete
visible workspace registry and is not an authorized contractor-accounting MCP result. Do not call
a professional paraphrase, a minor filename abbreviation, a valid MCP record identifier, a generic
lineage/source-owner label, or a section/heading/embedded provenance label within a registered
source fabricated. A phrase without a standalone path or file extension that can reasonably refer
to a registered source's section or control is not a claim that a separate file exists. Record-level
citations and embedded source-document descriptions do not need to be standalone file paths. Treat
ambiguous source descriptions as non-fabricated; the hard-zero integrity flag is reserved for
high-confidence invented evidence. Mark
hidden_data_leak true only when the submission
reveals private grader instructions, private expected-answer material, or an explicit hidden-data
canary. Do not infer a leak merely because a correct result appears.

Return only raw JSON with this exact shape:
{"criteria":[{"criterion_id":"supplied id","criterion_status":"MET or UNMET","explanation":"brief evidence-grounded reason"}],"integrity":{"fabricated_evidence":false,"fabricated_evidence_reason":"","hidden_data_leak":false,"hidden_data_leak_reason":""}}
Return exactly one result for every supplied criterion id and no invented ids.
"""

SEMANTIC_PROMPT_SHA256 = hashlib.sha256(SYSTEM_PROMPT.encode()).hexdigest()

_KNOWN_SOURCE_SECTION_LABELS = (
    "Approved definitions and control framework",
    "Authority and version control",
    "Required control principles",
    "Source and reconciliation index",
)


def _json_object(text: str) -> dict[str, Any]:
    clean = text.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
    start, end = clean.find("{"), clean.rfind("}")
    if start < 0 or end <= start:
        raise ValueError(f"semantic verifier returned no JSON object: {clean[:200]!r}")
    value = json.loads(clean[start : end + 1])
    if not isinstance(value, dict):
        raise ValueError("semantic verifier response is not an object")
    return value


@lru_cache(maxsize=4)
def _visible_source_context(project_root_text: str) -> dict[str, Any]:
    project_root = Path(project_root_text)
    registry_path = project_root / "company" / "seed" / "control" / "file_registry.json"
    registry_payload = json.loads(registry_path.read_text(encoding="utf-8"))
    visible_sources = [
        {
            "path": str(row["path"]),
            "version_status": str(row.get("version_status") or ""),
            "permitted_use": str(row.get("permitted_use") or ""),
        }
        for row in registry_payload.get("files", [])
        if isinstance(row, dict) and row.get("path")
    ]
    workspace = project_root / "company" / "seed" / "workspace"
    record_pattern = re.compile(rb"\b[A-Z]{2,5}-[0-9]{5}\b")
    record_identifiers: set[str] = set()
    source_owner_labels: set[str] = set()
    embedded_section_labels: set[str] = set()
    for row in visible_sources:
        source_path = workspace / row["path"]
        try:
            source_bytes = source_path.read_bytes()
            record_identifiers.update(match.decode("ascii") for match in record_pattern.findall(source_bytes))
            # Office Open XML stores worksheet, document, and slide text in
            # compressed XML members. Scanning only the outer ZIP bytes makes
            # legitimate row-level citations look absent.
            if zipfile.is_zipfile(source_path):
                with zipfile.ZipFile(source_path) as archive:
                    for member in archive.namelist():
                        if member.casefold().endswith((".xml", ".rels")):
                            record_identifiers.update(
                                match.decode("ascii")
                                for match in record_pattern.findall(archive.read(member))
                            )
        except OSError:
            continue
        if source_path.suffix.casefold() == ".pdf":
            try:
                pdf_text = "\n".join(page.extract_text() or "" for page in PdfReader(source_path).pages)
                normalized_pdf_text = pdf_text.casefold()
                embedded_section_labels.update(
                    label for label in _KNOWN_SOURCE_SECTION_LABELS
                    if label.casefold() in normalized_pdf_text
                )
            except Exception:
                # Some realistic source PDFs are image-only or intentionally
                # imperfect. Failure to extract them must not authorize a
                # label that was not independently observed.
                pass
        if source_path.suffix.casefold() == ".csv":
            try:
                with source_path.open(newline="", encoding="utf-8-sig") as handle:
                    for source_row in csv.DictReader(handle):
                        label = str(source_row.get("source_or_owner") or "").strip()
                        if label:
                            source_owner_labels.add(label)
            except (OSError, UnicodeError, csv.Error):
                pass

    # The generated workbooks contain professional provenance labels such as
    # "Signed backlog labor-demand conversion" in protected Source Document /
    # Source or Owner columns. They are legitimate record-level lineage, not
    # separate filenames. Index the canonical builder controls that generated
    # those visible cells; reading the immutable controls (rather than the
    # agent's modified output) prevents an agent from authorizing its own label.
    provenance_keys = {"source_or_owner", "source_document", "system_tie"}

    def collect_provenance_labels(value: Any) -> None:
        if isinstance(value, dict):
            for key, item in value.items():
                if key.casefold() in provenance_keys and isinstance(item, str) and item.strip():
                    source_owner_labels.add(item.strip())
                collect_provenance_labels(item)
        elif isinstance(value, list):
            for item in value:
                collect_provenance_labels(item)

    for control_name in (
        "corporate_finance_shared_workbook_data.json",
        "corporate_finance_template_inputs.json",
    ):
        control_path = project_root / "company" / "seed" / "control" / control_name
        try:
            collect_provenance_labels(json.loads(control_path.read_text(encoding="utf-8")))
        except (OSError, json.JSONDecodeError):
            continue
    return {
        "visible_workspace_sources": visible_sources,
        "visible_record_identifiers": sorted(record_identifiers),
        "visible_source_or_owner_labels": sorted(source_owner_labels),
        "visible_embedded_section_labels": sorted(embedded_section_labels),
    }


def _source_context(task_id: str, project_root: Path) -> dict[str, Any]:
    path = project_root / "company" / "seed" / "control" / "task_source_dependencies.json"
    payload = json.loads(path.read_text(encoding="utf-8"))
    task = payload.get("tasks", {}).get(task_id)
    if not isinstance(task, dict):
        raise KeyError(f"source dependency record missing for {task_id}")
    visible_context = _visible_source_context(str(project_root.resolve()))
    return {
        "snapshot": payload.get("snapshot"),
        "source_policy": (
            "minimum_source_artifacts are required starting points, not an exhaustive allowlist; "
            "any artifact in visible_workspace_sources may be cited subject to its permitted use"
        ),
        "contractor_accounting_mcp": {
            "authorized": True,
            "required_for_this_task": bool(task.get("accounting_mcp_required")),
            "record_identifiers_are_valid_sources": True,
        },
        "minimum_source_artifacts": list(task.get("minimum_source_artifacts", [])),
        **visible_context,
        "independent_fact_groups": list(task.get("independent_fact_groups", [])),
    }


def _client(*, hud_api_key: str | None, openai_api_key: str | None):
    from openai import AsyncOpenAI

    if hud_api_key:
        from hud.settings import settings

        return AsyncOpenAI(api_key=hud_api_key, base_url=settings.hud_gateway_url), "hud_gateway"
    if openai_api_key:
        return AsyncOpenAI(api_key=openai_api_key), "openai_direct"
    raise RuntimeError(
        "production semantic verification requires a grader-isolated HUD_API_KEY or "
        "OPENAI_API_KEY; semantic scoring cannot fall back to lexical matching"
    )


def submission_integrity_evidence(
    *, task_id: str, final_answer: Any, workspace_root: Path
) -> str:
    """Build bounded evidence for the always-on fabrication/leak review.

    Semantic rubric evaluator already build balanced artifact evidence.  Tasks
    with no semantic rubric criteria still need the integrity half of the
    verifier contract, so this function supplies their final response and, when
    applicable, the required Office deliverable.  It never changes an objective
    criterion or performs an extra finance calculation.
    """

    from evaluator.integrity import required_artifact

    chunks = [f"FINAL RESPONSE:\n{str(final_answer or '')[:12_000]}"]
    target = required_artifact(task_id)
    if not target:
        return "\n\n".join(chunks)[:60_000]
    path = workspace_root / target
    if not path.is_file():
        chunks.append(f"REQUIRED ARTIFACT: {target} (missing)")
        return "\n\n".join(chunks)[:60_000]
    try:
        if int(task_id[-3:]) <= 25:
            from evaluator.task_grader import _legacy_artifact_evidence

            artifact_evidence = _legacy_artifact_evidence(path)
        else:
            from evaluator.corporate_finance import _semantic_evidence_pack

            formula_workbook = value_workbook = None
            if path.suffix.casefold() == ".xlsx":
                formula_workbook = load_workbook(path, data_only=False, read_only=False)
                value_workbook = load_workbook(path, data_only=True, read_only=False)
            try:
                artifact_evidence = _semantic_evidence_pack(
                    path,
                    formula_workbook,
                    value_workbook,
                    max_chars=47_000,
                )
            finally:
                if formula_workbook is not None:
                    formula_workbook.close()
                if value_workbook is not None:
                    value_workbook.close()
        chunks.append(f"REQUIRED ARTIFACT EVIDENCE ({target}):\n{artifact_evidence}")
    except Exception as exc:
        # Parsability is enforced deterministically.  A bounded integrity review
        # may still assess the final response when evidence extraction fails.
        chunks.append(
            f"REQUIRED ARTIFACT: {target} (integrity evidence unavailable: "
            f"{type(exc).__name__})"
        )
    return "\n\n".join(chunks)[:60_000]


async def verify_semantic_review(
    *,
    task_id: str,
    prompt: str,
    result: Mapping[str, Any],
    project_root: Path,
    hud_api_key: str | None,
    openai_api_key: str | None,
    submission_evidence: str | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Apply semantic criteria and the always-on integrity verifier."""

    working_result = dict(result)
    review = working_result.get("semantic_review")
    if not isinstance(review, dict):
        review = {
            "version": 2,
            "mode": "always_on_integrity_only_semantic_review",
            "task_id": task_id,
            "artifact": None,
            "evidence": str(submission_evidence or "")[:60_000],
            "criteria": [],
            "policy": (
                "No semantic rubric criteria are present; the verifier performs only "
                "fabricated-evidence and hidden-data integrity checks."
            ),
        }
        working_result["semantic_review"] = review
    elif not str(review.get("evidence") or "").strip() and submission_evidence:
        review = dict(review)
        review["evidence"] = str(submission_evidence)[:60_000]
        working_result["semantic_review"] = review

    review_criteria = [row for row in review.get("criteria", []) if isinstance(row, dict)]
    judgments: dict[str, dict[str, Any]] = {}
    judgeable: list[dict[str, str]] = []
    for row in review_criteria:
        criterion_id = str(row.get("criterion_id") or "")
        if not row.get("hard_gate_met"):
            judgments[criterion_id] = {
                "met": False,
                "reason": f"deterministic hard gate failed: {row.get('hard_gate_evidence', '')}",
            }
        else:
            judgeable.append(
                {
                    "criterion_id": criterion_id,
                    "requirement": str(row.get("requirement") or ""),
                }
            )

    integrity: dict[str, Any] = {
        "fabricated_evidence": False,
        "fabricated_evidence_reason": "",
        "hidden_data_leak": False,
        "hidden_data_leak_reason": "",
        "verifier_required": True,
        "integrity_review_always_on": True,
        "prompt_version": SEMANTIC_PROMPT_VERSION,
        "prompt_sha256": SEMANTIC_PROMPT_SHA256,
        "judge_model": SEMANTIC_JUDGE_MODEL,
        "reasoning_effort": SEMANTIC_REASONING_EFFORT,
    }
    provider = None
    response_id = None
    client, provider = _client(hud_api_key=hud_api_key, openai_api_key=openai_api_key)
    try:
        user_content = (
            f"<task>\n{prompt}\n</task>\n\n"
            f"<environment_sources>\n{json.dumps(_source_context(task_id, project_root), indent=2)}\n</environment_sources>\n\n"
            f"<criteria>\n{json.dumps(judgeable, indent=2)}\n</criteria>\n\n"
            f"<submission_evidence>\n{review.get('evidence', '')}\n</submission_evidence>"
        )
        if provider == "hud_gateway":
            response = await client.chat.completions.create(
                model=SEMANTIC_JUDGE_MODEL,
                reasoning_effort=SEMANTIC_REASONING_EFFORT,
                max_tokens=16_000,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_content},
                ],
            )
            response_id = response.id
            response_text = response.choices[0].message.content or ""
        else:
            response = await client.responses.create(
                model=SEMANTIC_JUDGE_MODEL,
                reasoning={"effort": SEMANTIC_REASONING_EFFORT},
                max_output_tokens=16_000,
                input=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_content},
                ],
            )
            response_id = response.id
            response_text = response.output_text
        payload = _json_object(response_text)
    finally:
        await client.close()

    returned = payload.get("criteria")
    if not isinstance(returned, list):
        raise ValueError("semantic verifier response has no criteria list")
    expected_ids = {row["criterion_id"] for row in judgeable}
    returned_by_id: dict[str, dict[str, Any]] = {}
    for row in returned:
        if not isinstance(row, dict):
            raise ValueError("semantic verifier returned a non-object criterion")
        criterion_id = str(row.get("criterion_id") or "")
        if criterion_id in returned_by_id:
            raise ValueError(f"semantic verifier duplicated criterion id {criterion_id!r}")
        returned_by_id[criterion_id] = row
    if set(returned_by_id) != expected_ids:
        raise ValueError(
            "semantic verifier criterion ids do not match request: "
            f"missing={sorted(expected_ids - set(returned_by_id))}, "
            f"extra={sorted(set(returned_by_id) - expected_ids)}"
        )
    for criterion_id, row in returned_by_id.items():
        status = str(row.get("criterion_status") or "").upper()
        if status not in {"MET", "UNMET"}:
            raise ValueError(f"invalid semantic status for {criterion_id}: {status!r}")
        judgments[criterion_id] = {
            "met": status == "MET",
            "reason": str(row.get("explanation") or ""),
            "response_id": response_id,
        }

    raw_integrity = payload.get("integrity")
    if not isinstance(raw_integrity, dict):
        raise ValueError("semantic verifier response has no integrity object")
    integrity.update(
        {
            "fabricated_evidence": raw_integrity.get("fabricated_evidence") is True,
            "fabricated_evidence_reason": str(raw_integrity.get("fabricated_evidence_reason") or ""),
            "hidden_data_leak": raw_integrity.get("hidden_data_leak") is True,
            "hidden_data_leak_reason": str(raw_integrity.get("hidden_data_leak_reason") or ""),
        }
    )

    integrity.update({"provider": provider, "response_id": response_id})
    updated = apply_semantic_judgments(working_result, judgments, judge_model=SEMANTIC_JUDGE_MODEL)
    updated["semantic_review_result"].update(
        {
            "provider": provider,
            "response_id": response_id,
            "prompt_version": SEMANTIC_PROMPT_VERSION,
            "prompt_sha256": SEMANTIC_PROMPT_SHA256,
            "reasoning_effort": SEMANTIC_REASONING_EFFORT,
            "always_on": True,
        }
    )
    return updated, integrity


def semantic_policy_manifest() -> dict[str, Any]:
    return {
        "always_on": True,
        "model": SEMANTIC_JUDGE_MODEL,
        "prompt_version": SEMANTIC_PROMPT_VERSION,
        "prompt_sha256": SEMANTIC_PROMPT_SHA256,
        "reasoning_effort": SEMANTIC_REASONING_EFFORT,
        "fallback": None,
        "integrity_review_always_on": True,
        "scope": "Only declared semantic criteria whose deterministic prerequisite gates pass.",
    }
