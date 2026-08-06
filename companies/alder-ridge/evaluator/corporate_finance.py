from __future__ import annotations

import json
import math
import re
from pathlib import Path
from typing import Any, Iterable

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

from evaluator.task_grader import (
    Criterion,
    SEED_WORKSPACE,
    _answer_mapping,
    _close,
    _get,
    _normalize,
    _number,
    _recalculated_data_workbook,
    _result,
)
from evaluator.semantic import (
    contains_concept,
    date_matches,
    iter_numeric_candidates,
    ordered_semantic_list_matches,
    semantic_equal,
    semantic_value_matches,
    unordered_semantic_list_matches,
)
from evaluator.hybrid_semantic import semantic_requirement


GOLD_PATH = Path(__file__).resolve().parent / "reference" / "tasks_026_100.json"

_TASK_031_LABEL_ALIASES = {
    "peak_hiring_need": ["Peak cumulative hiring need"],
    "labor_cost_variance_to_plan": [
        "Variance to approved FY27 field-labor plan",
        "Labor cost variance to approved plan",
    ],
}

_TASK_037_LABEL_ALIASES = {
    "selected_portfolio": ["Selected ID"],
    "selected_capex": ["Cash capex"],
    "portfolio_npv": ["Selected NPV"],
    "downside_portfolio_npv": ["Downside portfolio NPV"],
}

_TASK_053_LABEL_ALIASES = {
    "terminal_growth_rate": ["Terminal growth"],
    "cost_of_equity": ["Cost of equity", "CAPM cost of equity"],
    "enterprise_value": [
        "Enterprise value - Gordon growth",
        "Primary enterprise value",
    ],
    "equity_value": [
        "Equity value - Gordon growth",
        "Primary equity value",
    ],
    "present_value_of_explicit_forecast": [
        "Present value of explicit forecast",
        "PV of explicit forecast",
        "PV of forecast cash flows",
        "PV of explicit-period FCF",
    ],
    "terminal_value": [
        "Terminal value - Gordon growth",
        "Gordon growth terminal value",
    ],
    "present_value_of_terminal_value": [
        "Present value of terminal value",
        "PV of terminal value",
    ],
    "terminal_value_share_of_enterprise_value": [
        "Terminal value share of enterprise value",
        "PV of terminal value / enterprise value",
        "PV of terminal value as % of EV",
    ],
    "net_debt": ["Net debt"],
    "downside_equity_value": [
        "Downside equity value",
        "High WACC / low growth equity value",
    ],
    "upside_equity_value": [
        "Upside equity value",
        "Low WACC / high growth equity value",
    ],
    "equity_value_sensitivity_range": [
        "Equity value sensitivity range",
        "Equity value range",
    ],
}

_TASK_075_LABEL_ALIASES = {
    "total_allocated": ["Total recommended uses / reserves", "Recommended program"],
    "estimated_value_creation": ["Value creation", "committee-estimated value"],
    "downside_minimum_cash": ["Downside cash before mitigation", "Downside"],
    "pro_forma_leverage": ["After payoff + $2.5m debt-funded acquisition"],
}

_TASK_076_LABEL_ALIASES = {
    "fy27_revenue": ["FY27 revenue", "Revenue"],
    "fy27_gross_profit": ["FY27 gross profit", "Gross profit"],
    "fy27_ebitda": ["Consolidated EBITDA", "EBITDA"],
    "fy27_free_cash_flow": ["Free cash flow"],
    "fy27_ending_cash": ["Ending unrestricted cash", "Ending cash"],
    "downside_free_cash_flow": ["Free cash flow"],
    "downside_revenue": ["Downside revenue", "Revenue"],
    "downside_ebitda": ["Downside EBITDA", "EBITDA"],
    "annual_pre_financing_ending_cash": ["Annual pre-financing ending cash estimate", "Annual pre-financing ending cash"],
    "quarterly_cash_reconciliation_difference": ["Labeled reconciliation difference", "Detailed less annual estimate"],
    "q1_ending_cash": ["Q1 ending cash", "Q1 cash", "Ending unrestricted cash"],
    "minimum_quarter_pre_financing_cash": ["Minimum pre-financing cash", "Lowest quarterly cash before financing", "Pre-revolver cash after scheduled principal"],
    "peak_revolver_draw": ["Peak revolver", "Maximum revolver draw", "Revolver draw"],
    "first_revolver_draw_quarter": ["First draw quarter", "Initial revolver draw quarter"],
    "year_end_funded_debt": ["FY27 ending funded debt", "Ending funded debt"],
    "year_end_gross_leverage": ["FY27 ending gross leverage", "Ending gross leverage", "Gross leverage on FY27 plan EBITDA"],
    "year_end_leverage_compliant": ["Year-end leverage compliance", "Leverage compliant", "Year-end leverage vs internal"],
}

_TASK_040_LABEL_ALIASES = {
    "fy31_free_cash_flow": ["Free cash flow before debt service"],
    "fy31_debt": ["Ending debt"],
    "fy31_cash": ["Ending cash"],
    "downside_peak_financing_plug": ["Maximum single-year financing plug"],
}

_TASK_055_LABEL_ALIASES = {
    "year_one_gaap_eps_accretion": ["GAAP EPS accretion / (dilution) %"],
    "year_one_adjusted_eps_accretion": ["Adjusted EPS accretion / (dilution) %"],
    "year_two_gaap_eps_accretion": ["GAAP EPS accretion / (dilution) %"],
    "year_two_adjusted_eps_accretion": ["Adjusted EPS accretion / (dilution) %"],
}

_TASK_061_LABEL_ALIASES = {
    "ending_temporary_difference_dta": ["temporary_difference_dta"],
    "ending_state_credit_dta": ["state_credit_dta"],
    "ending_valuation_allowance": ["valuation_allowance"],
    "ending_net_dta": ["net_dta"],
    "ending_net_deferred_tax_liability": ["net_deferred_tax_liability"],
}

_TASK_072_LABEL_ALIASES = {
    "q2_organic_growth": ["Organic growth"],
    "q2_adjusted_ebitda": ["Adjusted EBITDA"],
    "q2_free_cash_flow": ["Free cash flow"],
    "maximum_revolver": [
        "Maximum downside revolver",
        "FY27 downside maximum revolver",
        "Downside revolver",
    ],
    "latest_full_year_revenue_outlook": [
        "Latest approved outlook",
        "Revenue outlook",
    ],
}

_TASK_081_LABEL_ALIASES = {
    "selected_initiative_count": ["Selected count", "Selected initiatives"],
    "portfolio_spend": ["FY27 spend", "FY27 cash spend", "FY27 cash", "Selected spend"],
    "risk_adjusted_annual_ebitda_benefit": ["PW annual EBITDA", "Probability-weighted annual EBITDA"],
    "three_year_risk_adjusted_npv": ["Three-year RA NPV", "3-year RA NPV"],
    "unallocated_budget": ["Residual budget", "Residual FY27 budget", "Cash budget headroom", "FY27 reallocation cash budget"],
    "fy28_portfolio_spend": ["FY28 spend", "FY28 cash spend", "FY28 cash", "Selected FY28 spend"],
    "fy28_budget_headroom": [
        "FY28 headroom",
        "Residual FY28 budget",
        "FY28 residual budget",
        "FY28 cash budget headroom",
        "FY28 cash budget",
    ],
    "technician_hours_used": ["Technician hours", "Selected technician hours"],
    "technician_hours_headroom": ["Technician headroom", "Technician capacity headroom", "Remaining technician hours", "Technician-hour capacity"],
    "selected_resilience_initiatives": [
        "Resilience initiatives",
        "Selected resilience count",
        "Resilience count",
    ],
}

_TASK_082_LABEL_ALIASES = {
    "minimum_pre_financing_cash": ["Cash before financing"],
    "first_revolver_draw_week": ["First draw week"],
    "maximum_revolver_balance": [
        "Maximum ending revolver",
        "Peak ending revolver",
        "Peak debt",
    ],
    "ending_revolver_balance": ["Ending revolver", "Ending debt"],
    "total_revolver_interest": ["Total interest"],
    "total_unused_commitment_fees": ["Total unused fees"],
    "remaining_commitment_at_peak": [
        "Minimum remaining commitment",
        "Remaining contractual commitment at peak debt",
    ],
    "minimum_borrowing_base_headroom": [
        "Minimum borrowing-base headroom",
        "Minimum availability headroom",
        "Minimum collateral headroom",
        "Collateral headroom at peak debt",
        "Collateral headroom at peak-debt week",
    ],
    "first_operating_floor_breach_week": [
        "First floor breach week",
        "Operating cash floor breach",
        "Operating-floor breaches",
        "Floor-breach weeks",
    ],
    "maximum_unfunded_liquidity_shortfall": ["Maximum unfunded shortfall", "Unfunded liquidity"],
}

_TASK_087_LABEL_ALIASES = {
    "quality_of_earnings_adjustments": ["Accepted QoE adjustments", "Accepted standalone QoE adjustments"],
    "normalized_nwc_peg": ["NWC peg", "Median NWC peg"],
    "purchase_price_nwc_adjustment": ["Closing NWC adjustment", "Estimated closing NWC adjustment"],
    "debt_like_items": ["Gross debt and debt-like", "Total debt and debt-like", "Debt and debt-like items"],
    "usable_cash_credit": ["Usable target cash", "Usable cash", "Target cash credit"],
    "implied_enterprise_value": ["Total-consideration enterprise value", "Total consideration enterprise value", "Enterprise value", "Closing enterprise value"],
    "enterprise_value_to_adjusted_ebitda": ["Total-consideration EV / Adjusted EBITDA", "Total consideration EV / Adjusted EBITDA", "EV / Adjusted EBITDA"],
    "unauthorized_locked_box_leakage": ["Unauthorized leakage", "Unpermitted leakage"],
    "permitted_locked_box_leakage": ["Permitted leakage"],
    "adjusted_equity_purchase_price": ["Adjusted equity consideration", "Adjusted equity price", "Closing equity price"],
    "closing_escrow": ["Escrow", "Escrow withheld"],
    "earnout_fair_value": ["Earnout FV", "Contingent consideration fair value"],
    "total_purchase_consideration": ["Total consideration"],
    "cash_paid_to_seller_at_close": ["Direct seller cash", "Closing cash to seller", "Cash paid at close", "Cash paid directly to seller at close"],
    "total_closing_cash_uses": ["Closing cash uses", "Total cash uses", "Gross closing cash uses"],
    "net_buyer_funding_requirement": ["Buyer funding requirement", "Net funding requirement", "Net buyer cash funding"],
    "closing_funds_flow_check": ["Funds flow check", "Closing reconciliation", "Closing funds-flow check"],
}

_TASK_092_LABEL_ALIASES = {
    "term_loan_source": ["Acquisition term loan", "Term loan"],
    "acquisition_facility_source": ["Committed acquisition facility", "Acquisition facility"],
    "seller_note_source": ["Seller note actual funded", "Seller note funded", "Seller note"],
    "funding_gap": ["Funding gap", "Funds-flow check", "Funds flow check"],
    "financing_fees": ["Upfront fees", "Total fees", "Fee"],
    "annual_cash_interest": ["Annual cash interest", "Year 1 interest", "Acquisition interest"],
    "annual_principal_amortization": ["Scheduled Year 1 principal", "Year 1 scheduled principal", "Year 1 scheduled amortization", "Year 1 principal", "Acquisition principal", "Scheduled principal"],
    "base_leverage": ["Base gross leverage", "Gross leverage"],
    "downside_leverage": ["Downside gross leverage", "Gross leverage"],
    "base_debt_service_coverage": ["Base DSCR", "Debt service coverage", "DSCR"],
    "downside_debt_service_coverage": ["Downside DSCR", "Debt service coverage", "DSCR"],
    "downside_cash_headroom": ["Downside cash headroom", "Cash headroom shortfall", "Downside liquidity surplus shortfall", "Liquidity surplus shortfall", "Liquidity shortfall"],
    "recommendation": ["Recommendation", "Executive Recommendation"],
    "stressed_annual_cash_interest": ["Stressed interest", "Stress interest", "Stressed acquisition interest", "Interest stressed rates", "Year 1 interest stressed rates", "Rate stress annual interest", "Total stressed Year 1 interest", "Year-1 acquisition interest"],
    "stressed_downside_debt_service_coverage": ["Stressed downside DSCR", "Stressed DSCR Downside", "Rate stress DSCR", "Year-1 IC DSCR", "Debt service coverage"],
    "year_2_ending_acquisition_debt": ["Year 2 ending debt", "Ending acquisition debt", "Ending acquisition", "Year 2 scheduled acquisition debt", "Year 2 ending acquisition"],
    "year_2_ending_gross_funded_debt": ["Year 2 ending gross funded debt", "Year 2 gross funded debt", "Ending gross funded debt"],
    "year_2_cash_before_backstop": [
        "Year 2 pre-backstop cash",
        "Cash before revolver backstop",
        "Cash before revolver",
    ],
    "year_2_revolver_backstop_draw": ["Year 2 backstop draw", "Revolver backstop", "Backstop draw", "Drawn backstop"],
    "year_2_ending_cash_after_backstop": ["Year 2 ending cash", "Cash after backstop", "Ending cash after full backstop", "Minimum liquidity after backstop"],
    "year_2_remaining_backstop_commitment": ["Remaining backstop", "Remaining revolver commitment", "Backstop headroom", "Backstop draw remaining"],
    "year_2_stressed_gross_leverage": ["Year 2 gross leverage", "Year 2 stressed gross leverage", "Y2 stressed gross leverage"],
    "year_2_stressed_debt_service_coverage": ["Year 2 stressed DSCR", "Year 2 debt service coverage", "Year 2 DSCR", "Y2 stressed DSCR", "Stressed DSCR"],
    "year_2_cash_headroom": ["Year 2 cash headroom", "Year 2 cash headroom shortfall", "Year 2 liquidity headroom", "Y2 cash headroom", "Cash headroom", "Cash headroom shortfall", "Cash headroom to minimum", "Cash headroom shortfall to minimum"],
    "stressed_financing_case_compliant": ["Stress case compliance", "Stressed financing compliant", "Stressed Year 2 cash", "Overall guardrail result"],
}

_TASK_098_LABEL_ALIASES = {
    "ltm_revenue": ["LTM revenue"],
    "ltm_restated_adjusted_ebitda": ["LTM adjusted EBITDA", "LTM restated EBITDA"],
    "ltm_operating_cash_flow": ["LTM operating cash flow"],
    "q2_revenue_growth": ["Q2 revenue growth"],
    "q2_restated_adjusted_ebitda_margin": ["Q2 adjusted EBITDA margin", "Q2 restated EBITDA margin"],
    "q2_restated_ebitda_growth": ["Q2 adjusted EBITDA growth", "Q2 restated EBITDA growth"],
    "q2_free_cash_flow": ["Q2 free cash flow"],
    "net_debt": ["Net debt"],
    "net_leverage": ["Investor net leverage", "Net leverage"],
    "ending_backlog": ["Q2 ending backlog", "Ending backlog"],
    "q2_organic_revenue_growth": ["Organic Q2 revenue growth", "Q2 organic growth"],
    "ltm_operating_cash_conversion": ["LTM cash conversion", "Operating cash conversion"],
    "covenant_ebitda": ["Lender covenant EBITDA", "Covenant EBITDA"],
    "covenant_net_debt": ["Lender net debt", "Covenant net debt"],
    "covenant_net_leverage": ["Lender leverage", "Covenant net leverage", "Covenant-basis net leverage"],
    "investor_to_covenant_leverage_gap": ["Leverage definition gap", "Investor versus covenant leverage"],
    "quarterly_ebitda_restatement_bridge_check": ["Restatement bridge check", "Quarterly EBITDA bridge check"],
}

_TASK_100_LABEL_ALIASES = {
    "forecast_revenue": ["FY26 outlook", "FY26 forecast revenue", "revenue"],
    "revenue_variance_to_board_aop": ["vs Board AOP", "versus Board AOP", "Board AOP"],
    "revenue_variance_pct_to_board_aop": ["vs Board AOP", "versus Board AOP", "Board AOP"],
    "forecast_branch_ebitda": ["FY26 forecast branch EBITDA", "forecast branch EBITDA", "branch EBITDA"],
    "branch_ebitda_variance_to_current_comparator": ["branch EBITDA variance", "branch EBITDA gain", "branch economics", "vs current comparator"],
    "largest_branch_ebitda_miss": ["largest branch EBITDA miss", "largest negative branch variance"],
    "largest_branch_ebitda_miss_amount": ["largest branch EBITDA miss", "largest negative branch variance", "Controls"],
    "downside_leverage": ["Downside funded-debt leverage", "Downside leverage", "Downside"],
    "severe_leverage": ["Severe funded-debt leverage", "Severe leverage", "Severe"],
    "downside_ending_cash": ["Downside ending cash", "Downside cash", "Downside"],
    "severe_ending_cash": ["Severe ending cash", "Severe cash", "Severe"],
    "severe_revolver_headroom": ["Severe revolver headroom", "Severe headroom", "Severe"],
    "total_probability_weighted_priority_ebitda": ["total probability-weighted benefit", "total priority benefit", "PW benefit", "unconstrained PW benefit", "all priorities", "full priority list", "priority value"],
    "executable_probability_weighted_priority_ebitda": ["executable", "Approved or Gated", "Approved plus Gated", "Approved + Gated PW"],
    "severe_covenant_breach": ["Severe covenant breach", "breach flags"],
    "optimized_board_priority_portfolio": ["Optimized priority portfolio", "Selected board priorities", "Selected portfolio"],
    "optimized_priority_cash_spend": ["Selected priority spend", "Optimized cash spend", "Selected cash", "Budget"],
    "optimized_probability_weighted_base_benefit": ["Optimized Base benefit", "Selected probability-weighted benefit", "PW Base benefit", "Approved + Gated", "Executable FY27 portfolio"],
    "optimized_probability_weighted_severe_benefit": ["Optimized Severe benefit", "Selected Severe benefit", "Severe benefit"],
    "post_action_severe_incremental_revolver_draw": ["Post-action incremental revolver draw", "Severe incremental draw", "Incremental revolver funding", "Incremental draw"],
    "post_action_severe_remaining_revolver_headroom": ["Post-action revolver headroom", "Remaining Severe revolver headroom", "Remaining revolver capacity", "Remaining headroom"],
    "post_action_severe_unfunded_liquidity_shortfall": ["Post-action unfunded shortfall", "Severe unfunded liquidity", "Residual liquidity shortfall"],
    "post_action_severe_ending_cash": ["Post-action Severe cash", "Severe cash after priorities"],
    "post_action_severe_leverage": ["Post-action Severe leverage", "Severe leverage after priorities"],
    "post_action_severe_cash_breach": ["Post-action Severe cash breach", "Severe cash gate"],
    "post_action_severe_leverage_breach": ["Post-action Severe leverage breach", "Severe leverage gate"],
}

_ARTIFACT_TOKEN_ALIASES = {
    # A normal source-tie tab usually names the actual request/email and its
    # path instead of repeating the rubric phrase verbatim.
    "management correspondence": [
        "management correspondence",
        "management email",
        "request thread",
        "Requests/",
        "controller follow-up",
        "source-tie thread",
        "board review",
    ],
    # A decision memo can explain risk through explicit downside failures,
    # guardrails, or quantified shortfalls without using the literal word.
    "risk": ["risk", "downside", "guardrail", "failure", "shortfall"],
    "funding gap": ["funding gap", "unfunded gap", "funds flow check", "excess shortfall", "funds exactly", "sources less uses", "source use difference"],
    # Workbooks commonly identify the governed file/status directly (for
    # example v6, v7.6, or "current approved") instead of spelling out the
    # rubric noun "version".
    "version": ["version", "source version", "document version", "file version", "v4", "v6", "v7.4", "v7.6", "v8", "current approved", "current source", "controller-tied", "controlling case", "current - controller tie complete"],
    "working capital peg": ["working capital peg", "NWC peg", "median NWC peg", "monthly NWC"],
    # Finance workbooks often state the display unit directly (for example
    # USD in thousands) instead of repeating the noun unit.
    "unit": ["unit", "units", "USD in thousands", "USD thousands", "$000", "$mm", "USD millions"],
    # Purchase-price models normally use the operative schedule labels rather
    # than rubric prose. These aliases preserve the finance meaning.
    "cash credit": ["cash credit", "usable cash", "usable target cash", "target cash credit"],
    "locked box": ["locked box", "locked-box", "unauthorized leakage", "permitted leakage", "leakage"],
    "investor leverage": ["investor leverage", "investor net leverage"],
    "covenant leverage": ["covenant leverage", "covenant net leverage", "covenant-basis net leverage", "lender leverage"],
    # A formula-driven model can expose its state as a feasibility/result
    # control and a matrix of PASS/FAIL checks without repeating the literal
    # rubric phrase "model status".
    "model status": ["model status", "overall status", "control status", "feasibility", "model integrity"],
}


def load_corporate_finance_gold(task_id: str | None = None) -> dict[str, Any]:
    payload = json.loads(GOLD_PATH.read_text(encoding="utf-8"))
    return payload[task_id] if task_id else payload


def _flatten(value: Any) -> list[str]:
    if isinstance(value, dict):
        return [item for key, child in value.items() for item in (str(key), *_flatten(child))]
    if isinstance(value, (list, tuple, set)):
        return [item for child in value for item in _flatten(child)]
    return [str(value)] if value is not None else []


def _numeric_matches(actual: Any, expected: float, abs_tol: float) -> bool:
    return any(_close(candidate, expected, abs_tol=abs_tol, rel_tol=0.0) for candidate in iter_numeric_candidates(actual))


def _numeric_matches_spec(actual: Any, expected: float, spec: dict[str, Any]) -> bool:
    abs_tol = float(spec.get("abs_tol", 0.02))
    aggregate_field = spec.get("aggregate_field")
    if aggregate_field and isinstance(actual, (list, tuple)):
        aggregate_aliases = [aggregate_field, *spec.get("aggregate_field_aliases", [])]
        amounts = []
        for item in actual:
            if not isinstance(item, dict):
                continue
            raw = next(
                (
                    value for candidate in aggregate_aliases
                    for key, value in item.items()
                    if semantic_equal(key, candidate)
                ),
                None,
            )
            if isinstance(raw, (int, float)) and not isinstance(raw, bool):
                amounts.append(float(raw))
        concepts = list(spec.get("expected_concepts", []))
        concepts_match = not concepts or (
            unordered_semantic_list_matches(actual, concepts)
            if spec.get("unordered_aggregate_concepts")
            else ordered_semantic_list_matches(actual, concepts)
        )
        if len(amounts) == len(actual) and amounts and concepts_match and _close(sum(amounts), expected, abs_tol=abs_tol, rel_tol=0.0):
            return True
    if spec.get("sum_numeric_sequence") and isinstance(actual, (list, tuple)):
        amounts = [
            float(item) for item in actual
            if isinstance(item, (int, float)) and not isinstance(item, bool)
        ]
        if len(amounts) == len(actual) and amounts and _close(sum(amounts), expected, abs_tol=abs_tol, rel_tol=0.0):
            return True
    if spec.get("aggregate_structured_amounts") and isinstance(actual, (list, tuple)):
        amounts: list[float] = []
        for item in actual:
            if not isinstance(item, dict):
                return False
            raw = next((item[key] for key in ("amount", "value") if key in item), None)
            if not isinstance(raw, (int, float)) or isinstance(raw, bool):
                return False
            amounts.append(float(raw))
        concepts = list(spec.get("expected_concepts", []))
        return (
            len(amounts) == len(concepts)
            and ordered_semantic_list_matches(actual, concepts)
            and _close(sum(amounts), expected, abs_tol=abs_tol, rel_tol=0.0)
        )
    if _numeric_matches(actual, expected, abs_tol):
        return True
    denominator = spec.get("ratio_alternative_denominator")
    if denominator:
        ratio = expected / float(denominator)
        ratio_tol = float(spec.get("ratio_alternative_abs_tol", 0.0001))
        if any(_close(candidate, ratio, abs_tol=ratio_tol, rel_tol=0.0) for candidate in iter_numeric_candidates(actual)):
            return True
    decimals = spec.get("display_decimals")
    if decimals is not None:
        tolerance = 0.5 * (10 ** -int(decimals)) + 1e-12
        return any(_close(candidate, expected, abs_tol=max(abs_tol, tolerance), rel_tol=0.0) for candidate in iter_numeric_candidates(actual))
    return False


def _boolean_matches(actual: Any, expected: bool) -> bool:
    if isinstance(actual, bool):
        return actual is expected
    normalized = _normalize(actual)
    positive = {"true", "yes", "y", "pass", "compliant", "met"}
    negative = {"false", "no", "n", "fail", "noncompliant"}
    if normalized in positive | negative | {"1", "0", "not met"}:
        return normalized in positive | {"1"} if expected else normalized in negative | {"0", "not met"}
    words = set(normalized.split())
    positive_hit = bool(words & positive)
    negative_hit = bool(words & negative) or "not met" in normalized
    return positive_hit and not negative_hit if expected else negative_hit and not positive_hit


def _semantic_get(mapping: dict[str, Any], key: str) -> Any:
    """Prefer the requested key, then accept harmless label morphology."""
    actual = _get(mapping, key)
    if actual is not None:
        return actual
    for candidate, value in mapping.items():
        if semantic_equal(candidate, key):
            return value
    return None


def _semantic_find(mapping: dict[str, Any], key: str) -> tuple[bool, Any]:
    if key in mapping:
        return True, mapping[key]
    for candidate, value in mapping.items():
        if semantic_equal(candidate, key):
            return True, value
    return False, None


def _period_label_matches(actual: Any, expected: Any) -> bool:
    """Accept a bare ordinal only when the output key already supplies the period type."""
    def parse(value: Any) -> tuple[str | None, int] | None:
        if isinstance(value, (int, float)) and not isinstance(value, bool) and float(value).is_integer():
            return None, int(value)
        normalized = _normalize(value)
        if re.fullmatch(r"\d+", normalized):
            return None, int(normalized)
        match = re.fullmatch(r"(?:fiscal )?(year|yr|y|month|mo|week|wk|quarter|q)\s*-?\s*(\d+)", normalized)
        if not match:
            return None
        period = {
            "yr": "year", "y": "year", "mo": "month",
            "wk": "week", "q": "quarter",
        }.get(match.group(1), match.group(1))
        return period, int(match.group(2))

    left, right = parse(actual), parse(expected)
    return bool(
        left and right and left[1] == right[1]
        and (left[0] is None or right[0] is None or left[0] == right[0])
    )


def _console_criterion(mapping: dict[str, Any], spec: dict[str, Any]) -> Criterion:
    kind = spec["kind"]
    if kind == "numeric":
        failures = []
        for key, expected in spec["expected"].items():
            actual = _semantic_get(mapping, key)
            if not _numeric_matches_spec(actual, float(expected), spec):
                failures.append(f"{key}={actual!r}; expected {expected}")
        met = not failures
        evidence = "reported value matched" if met else "; ".join(failures)
    elif kind == "string":
        failures = []
        for key, expected in spec["expected"].items():
            present, actual = _semantic_find(mapping, key)
            expected_none = (
                expected is None
                or _normalize(expected) in {"none", "null"}
            )
            actual_none = actual is None or _normalize(actual) in {"none", "no draw", "no breach", "not applicable", "n a"}
            typed_period = key.endswith(("_year", "_month", "_week", "_quarter")) and _period_label_matches(actual, expected)
            numeric_alternative = False
            if present and "numeric_alternative" in spec:
                numeric_alternative = _numeric_matches_spec(
                    actual,
                    float(spec["numeric_alternative"]),
                    {"abs_tol": float(spec.get("numeric_alternative_abs_tol", 0.02))},
                )
            if not present or not ((expected_none and actual_none) or typed_period or numeric_alternative or semantic_value_matches(actual, str(expected))):
                failures.append(f"{key}={actual!r}; expected {expected!r}")
        met = not failures
        evidence = "classification matched" if met else "; ".join(failures)
    elif kind == "boolean":
        failures = []
        for key, expected in spec["expected"].items():
            actual = _semantic_get(mapping, key)
            if not _boolean_matches(actual, bool(expected)):
                failures.append(f"{key}={actual!r}; expected {expected!r}")
        met = not failures
        evidence = "boolean conclusion matched" if met else "; ".join(failures)
    elif kind == "list_item":
        actual = _semantic_get(mapping, spec["key"])
        expected = spec["expected_item"]
        if isinstance(actual, dict):
            flattened = [item for key, value in actual.items() for item in (key, value)]
        elif isinstance(actual, (list, tuple, set)):
            flattened = list(actual)
        elif actual is not None:
            flattened = [actual]
        else:
            flattened = []
        present_in_required_bucket = any(
            semantic_value_matches(item, str(expected)) for item in flattened
        )
        conflicting_matches = {}
        for conflicting_key in spec.get("exclusive_with_keys", []):
            conflicting_actual = _semantic_get(mapping, conflicting_key)
            if isinstance(conflicting_actual, dict):
                conflicting_items = [
                    item
                    for key, value in conflicting_actual.items()
                    for item in (key, value)
                ]
            elif isinstance(conflicting_actual, (list, tuple, set)):
                conflicting_items = list(conflicting_actual)
            elif conflicting_actual is not None:
                conflicting_items = [conflicting_actual]
            else:
                conflicting_items = []
            if any(
                semantic_value_matches(item, str(expected))
                for item in conflicting_items
            ):
                conflicting_matches[str(conflicting_key)] = conflicting_actual
        met = present_in_required_bucket and not conflicting_matches
        evidence = (
            f"actual={actual!r}; required_item={expected!r}; "
            f"conflicting_matches={conflicting_matches!r}"
        )
    elif kind in {"list", "list_exact"}:
        actual = _semantic_get(mapping, spec["key"])
        expected_values = list(spec["expected"])
        if spec.get("unordered"):
            met = unordered_semantic_list_matches(actual, expected_values)
            expectation = "unordered concepts"
        else:
            met = (isinstance(actual, list) and not actual) if not expected_values else ordered_semantic_list_matches(actual, expected_values)
            expectation = "ordered concepts"
        evidence = f"actual={actual!r}; expected {expectation}={expected_values!r}"
    else:
        raise ValueError(f"Unsupported console criterion: {kind}")
    return Criterion(spec["id"], spec["description"], met, evidence)


def _grade_console(task_id: str, answer: Any) -> dict[str, Any]:
    gold = load_corporate_finance_gold(task_id)
    mapping = _answer_mapping(answer)
    criteria = [_console_criterion(mapping, spec) for spec in gold["criteria"]]
    result = _result(criteria)
    reviews: list[dict[str, Any]] = []
    by_id = {criterion.id: criterion for criterion in criteria}
    for spec in gold["criteria"]:
        if not spec.get("semantic"):
            continue
        if spec["kind"] in {"string", "boolean"}:
            expected_facts = dict(spec["expected"])
        elif spec["kind"] == "list_item":
            if spec.get("exclusive_with_keys"):
                expected_facts = {
                    spec["key"]: {"must_include": spec["expected_item"]}
                }
                expected_facts.update(
                    {
                        conflicting_key: {"must_exclude": spec["expected_item"]}
                        for conflicting_key in spec["exclusive_with_keys"]
                    }
                )
            else:
                expected_facts = {spec["key"]: spec["expected_item"]}
        else:
            expected_facts = {spec["key"]: list(spec["expected"])}
        reviews.append(
            {
                "criterion_id": spec["id"],
                "requirement": semantic_requirement(
                    criterion_id=spec["id"],
                    description=spec["description"],
                    expected_facts=expected_facts,
                    artifact_type="submitted answer",
                ) + (
                    " For a list criterion, require the complete requested set: reject missing, "
                    "extra, duplicated, or wrongly classified items."
                    if spec["kind"] in {"list", "list_exact"} else ""
                ) + (
                    " For this membership criterion, require the item in the named bucket and "
                    "absent from every conflicting bucket."
                    if spec.get("exclusive_with_keys") else ""
                ),
                # Boolean conclusions are objective typed facts.  They must
                # match deterministically before the semantic verifier may
                # assess professional association.  A verifier must never be
                # able to turn an explicitly wrong true/false answer into a
                # pass merely because the field name is present.
                "hard_gate_met": (
                    bool(by_id[spec["id"]].met)
                    if (
                        spec["kind"] == "boolean"
                        or (
                            spec["kind"] == "list_item"
                            and spec.get("exclusive_with_keys")
                        )
                    )
                    else bool(mapping)
                ),
                "hard_gate_evidence": (
                    by_id[spec["id"]].evidence
                    if (
                        spec["kind"] == "boolean"
                        or (
                            spec["kind"] == "list_item"
                            and spec.get("exclusive_with_keys")
                        )
                    )
                    else (
                        "answer parsed into a non-empty field mapping"
                        if mapping else "answer did not parse into a non-empty field mapping"
                    )
                ),
                "legacy_lexical_match": bool(by_id[spec["id"]].met),
            }
        )
    if reviews:
        result["semantic_review"] = {
            "version": 2,
            "mode": "deterministic_hard_gates_plus_bounded_semantic_judge",
            "task_id": task_id,
            "artifact": None,
            "evidence": str(answer or "")[:60_000],
            "criteria": reviews,
            "policy": (
                "A criterion passes only when its parsing gate passes and the semantic judge "
                "finds the exact professional conclusion or complete list MET."
            ),
        }
    by_spec = {str(spec["id"]): spec for spec in gold["criteria"]}
    for row in result["criteria"]:
        spec = by_spec[str(row["id"])]
        for key in ("category", "weight", "failure_cap", "semantic", "kind"):
            if key in spec:
                row[key] = spec[key]
    return result


def _xlsx_entries(formula_workbook, value_workbook) -> list[tuple[str, Any]]:
    entries: list[tuple[str, Any]] = []
    for sheet_name in formula_workbook.sheetnames:
        formula_sheet = formula_workbook[sheet_name]
        value_sheet = value_workbook[sheet_name] if sheet_name in value_workbook.sheetnames else formula_sheet
        for row in formula_sheet.iter_rows():
            for cell in row:
                value = value_sheet[cell.coordinate].value
                entries.append((str(cell.value or ""), value))
    return entries


def _xlsx_label_value(workbook, values, label: str, expected: Any) -> bool:
    wanted = _normalize(label)
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            for cell in row:
                if not (semantic_equal(cell.value, label) or contains_concept(cell.value, label)):
                    continue
                candidates = []
                for offset in (1, 2, 3):
                    candidates.append(value_sheet.cell(cell.row, cell.column + offset).value)
                candidates.append(value_sheet.cell(cell.row + 1, cell.column).value)
                if _value_candidates_match(candidates, expected):
                    return True
    return False


def _xlsx_label_formula(workbook, label: str) -> tuple[bool, str]:
    wanted = _normalize(label)
    label_locations: list[str] = []
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if not (semantic_equal(cell.value, label) or contains_concept(cell.value, label)):
                    continue
                label_locations.append(f"{sheet_name}!{cell.coordinate}")
                candidates = [
                    sheet.cell(cell.row, cell.column + offset)
                    for offset in (1, 2, 3)
                    if cell.column + offset <= sheet.max_column
                ]
                if cell.row + 1 <= sheet.max_row:
                    candidates.append(sheet.cell(cell.row + 1, cell.column))
                for candidate in candidates:
                    formula = candidate.value
                    if not isinstance(formula, str) or not formula.startswith("="):
                        continue
                    # A formula that merely wraps the released answer (for
                    # example =12345) is not model lineage. Require at least
                    # one cell/range reference in the headline formula.
                    if re.search(r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)?!?\$?[A-Z]{1,3}\$?\d+", formula):
                        return True, f"{sheet_name}!{candidate.coordinate}={formula}"
    if label_locations:
        return False, f"label found at {label_locations!r}, but no adjacent source-linked formula"
    return False, "headline label not found"


def _is_source_linked_formula(value: Any) -> bool:
    return (
        isinstance(value, str)
        and value.startswith("=")
        and re.search(
            r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)?!?\$?[A-Z]{1,3}\$?\d+",
            value,
        ) is not None
    )


def _xlsx_row_column_formula(
    workbook,
    row_label: str,
    column_label: str,
) -> tuple[bool, str]:
    """Find a source-linked formula at a semantic row/column intersection."""

    label_locations: list[str] = []
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            row_labels = [
                cell
                for cell in row
                if semantic_equal(cell.value, row_label)
                or contains_concept(cell.value, row_label)
            ]
            if not row_labels:
                continue
            label_locations.extend(
                f"{sheet.title}!{cell.coordinate}" for cell in row_labels
            )
            for candidate in row:
                if not _is_source_linked_formula(candidate.value):
                    continue
                headers = [
                    sheet.cell(header_row, candidate.column).value
                    for header_row in range(max(1, candidate.row - 12), candidate.row)
                ]
                if any(
                    semantic_equal(header, column_label)
                    or contains_concept(header, column_label)
                    for header in headers
                ):
                    return (
                        True,
                        f"{sheet.title}!{candidate.coordinate}={candidate.value}; "
                        f"row={row_label!r}; column={column_label!r}",
                    )
    if label_locations:
        return (
            False,
            f"row label found at {label_locations!r}, but no source-linked "
            f"formula was under column {column_label!r}",
        )
    return False, f"row label {row_label!r} not found"


def _xlsx_section_row_formula(
    workbook,
    section_label: str,
    row_label: str,
) -> tuple[bool, str]:
    """Find a formula row within the bounded schedule section named by the task."""

    for sheet in workbook.worksheets:
        section_rows = [
            cell.row
            for row in sheet.iter_rows()
            for cell in row
            if semantic_equal(cell.value, section_label)
            or contains_concept(cell.value, section_label)
        ]
        for section_row in section_rows:
            for row_number in range(section_row + 1, min(sheet.max_row, section_row + 20) + 1):
                row = list(sheet[row_number])
                if not any(
                    semantic_equal(cell.value, row_label)
                    or contains_concept(cell.value, row_label)
                    for cell in row
                ):
                    continue
                for candidate in row:
                    if _is_source_linked_formula(candidate.value):
                        return True, f"{sheet.title}!{candidate.coordinate}={candidate.value}"
                return False, f"{sheet.title}!{row_number} has no source-linked formula"
    return False, f"section {section_label!r} / row {row_label!r} not found"


def _task_081_binding_constraint_formula(workbook) -> tuple[bool, str]:
    """Accept a source-linked binding-constraint output near its dashboard label."""

    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if not (
                    semantic_equal(cell.value, "binding constraint")
                    or contains_concept(cell.value, "binding constraint")
                ):
                    continue
                candidates = []
                for row_offset, column_offset in (
                    (0, 1),
                    (1, 0),
                    (1, 1),
                    (0, 2),
                    (2, 0),
                    (2, 1),
                ):
                    candidate_row = cell.row + row_offset
                    candidate_column = cell.column + column_offset
                    if (
                        candidate_row <= sheet.max_row
                        and candidate_column <= sheet.max_column
                    ):
                        candidates.append(
                            sheet.cell(candidate_row, candidate_column)
                        )
                candidate = next(
                    (
                        candidate
                        for candidate in candidates
                        if _is_source_linked_formula(candidate.value)
                    ),
                    None,
                )
                if candidate is not None:
                    return True, f"{sheet.title}!{candidate.coordinate}={candidate.value}"
                return (
                    False,
                    f"binding-constraint label at {sheet.title}!{cell.coordinate}; "
                    "no nearby output is a source-linked formula",
                )
    return False, "binding-constraint label not found"


def _task_081_constraints_input_sheet(sheet) -> tuple[bool, str]:
    """Recognize a sourced hardcoded optimizer-constraints input schedule.

    The Constraints tab is an input/control surface. Requiring formulas on it
    is a hidden design proxy; however, a blank or unsourced sheet should not
    earn auditability credit. Require several numeric limits plus explicit
    constraint and source-reference labeling.
    """

    constants = [
        cell.value
        for row in sheet.iter_rows()
        for cell in row
        if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)
    ]
    text = " ".join(
        str(cell.value)
        for row in sheet.iter_rows()
        for cell in row
        if cell.value is not None
    )
    has_constraint_label = contains_concept(text, "constraint")
    has_source_label = any(
        contains_concept(text, label)
        for label in ("source reference", "source ref", "authority")
    )
    met = len(constants) >= 3 and has_constraint_label and has_source_label
    return (
        met,
        f"numeric_limits={len(constants)}; "
        f"constraint_label={has_constraint_label}; source_label={has_source_label}",
    )


def _task_037_selection_tie_control(workbook, values) -> tuple[bool, str]:
    """Require a real independent selected-set calculation that ties to zero."""

    if "Checks" not in workbook.sheetnames or "Checks" not in values.sheetnames:
        return False, "Checks sheet is missing"
    sheet = workbook["Checks"]
    value_sheet = values["Checks"]
    reference_pattern = re.compile(
        r"(?:'Portfolio Selection'|Portfolio Selection)!"
        r"\$?([A-Z]{1,3})\$?(\d+)"
        r"(?::\$?([A-Z]{1,3})\$?(\d+))?",
        flags=re.I,
    )
    for row in sheet.iter_rows():
        for cell in row:
            if not (
                contains_concept(cell.value, "selection tie")
                or contains_concept(cell.value, "selected portfolio tie")
            ):
                continue
            candidates = [
                sheet.cell(cell.row, cell.column + offset)
                for offset in (1, 2, 3)
                if cell.column + offset <= sheet.max_column
            ]
            if cell.row + 1 <= sheet.max_row:
                candidates.append(sheet.cell(cell.row + 1, cell.column))
            for candidate in candidates:
                formula = candidate.value
                if not isinstance(formula, str) or not formula.startswith("="):
                    continue
                references = reference_pattern.findall(formula)
                normalized_references = {
                    (
                        start_column.casefold(),
                        int(start_row),
                        end_column.casefold() if end_column else "",
                        int(end_row) if end_row else 0,
                    )
                    for start_column, start_row, end_column, end_row
                    in references
                }
                single_references = {
                    reference
                    for reference in normalized_references
                    if not reference[2]
                }
                range_references = {
                    reference
                    for reference in normalized_references
                    if reference[2]
                }
                recalculated = value_sheet[candidate.coordinate].value
                if (
                    "sumproduct(" in formula.casefold().replace(" ", "")
                    and len(single_references) >= 1
                    and len(range_references) >= 2
                    and len(normalized_references) >= 3
                    and isinstance(recalculated, (int, float))
                    and not isinstance(recalculated, bool)
                    and _close(
                        float(recalculated),
                        0.0,
                        abs_tol=0.01,
                        rel_tol=0.0,
                    )
                ):
                    return (
                        True,
                        f"Checks!{candidate.coordinate}={formula}; "
                        f"value={recalculated}",
                    )
            return (
                False,
                f"selection-tie label found at Checks!{cell.coordinate}, "
                "but no independent selected-set zero formula",
            )
    return False, "no selection-tie control on Checks"


def _task_053_row_match(
    workbook,
    values,
    label: str,
    expected: Any,
    *,
    require_formula: bool,
) -> tuple[bool, str]:
    """Read professional DCF labels in clearly disclosed USD units."""

    unit_text = " ".join(
        str(cell.value)
        for sheet in workbook.worksheets
        for row in sheet.iter_rows()
        for cell in row
        if cell.value is not None
    ).casefold()
    millions_disclosed = any(
        marker in unit_text
        for marker in ("$mm", "usd millions", "usd in millions")
    )
    usd_disclosed = bool(re.search(r"\busd\b", unit_text))
    monetary = (
        isinstance(expected, (int, float))
        and not isinstance(expected, bool)
        and abs(float(expected)) > 10_000
    )
    if monetary and not (millions_disclosed or usd_disclosed):
        return False, "workbook does not explicitly disclose USD units"

    aliases = [
        label.replace("_", " "),
        *_TASK_053_LABEL_ALIASES.get(label, []),
    ]
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = (
            values[sheet_name]
            if sheet_name in values.sheetnames
            else sheet
        )
        for row in sheet.iter_rows():
            row_text = " | ".join(
                str(cell.value) for cell in row if cell.value is not None
            )
            if not any(
                contains_concept(row_text, alias)
                for alias in aliases
            ):
                continue
            for cell in row:
                formula = cell.value
                cached = value_sheet[cell.coordinate].value
                if require_formula and not (
                    isinstance(formula, str) and formula.startswith("=")
                ):
                    continue
                if (
                    not isinstance(cached, (int, float))
                    or isinstance(cached, bool)
                ):
                    continue
                if monetary and millions_disclosed:
                    targets = [float(expected) / 1_000_000.0]
                else:
                    targets = [float(expected)]
                if any(
                    _close(
                        float(cached),
                        target,
                        abs_tol=max(2e-8, abs(target) * 1e-6),
                        rel_tol=0.0,
                    )
                    for target in targets
                ):
                    return (
                        True,
                        f"{sheet_name}!{cell.coordinate}={formula!r}; "
                        f"cached={cached!r}",
                    )
    return False, "no exact professional DCF row matched"


def _display_tolerance(expected: float) -> float:
    if abs(expected) <= 10:
        return 0.0005
    return max(0.02, abs(expected) * 0.0005)


def _value_candidates_match(candidates: Iterable[Any], expected: Any) -> bool:
    if isinstance(expected, list):
        return ordered_semantic_list_matches(list(candidates), expected)
    if isinstance(expected, str):
        return any(date_matches(candidate, expected) or semantic_value_matches(candidate, expected) for candidate in candidates)
    if isinstance(expected, bool):
        return any(_boolean_matches(candidate, expected) for candidate in candidates)
    return any(
        _close(candidate, float(expected), abs_tol=_display_tolerance(float(expected)), rel_tol=0.0)
        for candidate in candidates
    )


def _task_076_row_match(workbook, values, label: str, expected: Any, *, require_formula: bool) -> tuple[bool, str]:
    """Read the integrated plan's disclosed $000 schedules without losing lineage.

    The model may expose a required headline as the exact matching quarter,
    consolidated column, minimum, or maximum within a labeled formula row.
    Accept only an exact cached result (with the explicitly declared $000 unit)
    and, for lineage checks, require the matching cell itself to be a formula.
    """
    aliases = [label.replace("_", " "), *_TASK_076_LABEL_ALIASES.get(label, [])]
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            row_text = " | ".join(str(cell.value) for cell in row if cell.value is not None)
            if not any(contains_concept(row_text, alias) for alias in aliases):
                continue
            for cell in row:
                cached = value_sheet[cell.coordinate].value
                formula = cell.value
                if require_formula and not (isinstance(formula, str) and formula.startswith("=")):
                    continue
                matched = False
                if isinstance(expected, bool):
                    matched = _boolean_matches(cached, expected)
                elif isinstance(expected, str):
                    matched = date_matches(cached, expected) or semantic_value_matches(cached, expected)
                elif isinstance(cached, (int, float)) and not isinstance(cached, bool):
                    targets = [float(expected)]
                    if abs(float(expected)) > 10:
                        targets.append(float(expected) / 1_000.0)
                    matched = any(
                        _close(cached, target, abs_tol=max(0.00002, abs(target) * 0.000001), rel_tol=0.0)
                        for target in targets
                    )
                if matched:
                    return True, f"{sheet_name}!{cell.coordinate}={formula!r}; cached={cached!r}"
    return False, "no exact formula-row value matched"


def _task_087_row_match(workbook, values, label: str, expected: Any, *, require_formula: bool) -> tuple[bool, str]:
    """Read the QoE model's explicitly disclosed USD-thousands schedules.

    A normal transaction model states USD in thousands once per sheet and
    then presents 4,860 rather than 4,860,000 on every row. Keep this
    alternative task-scoped, require an explicit workbook unit disclosure,
    and match only exact values on a row carrying an approved headline label.
    """
    unit_text = " ".join(
        str(cell.value)
        for sheet in workbook.worksheets
        for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 4))
        for cell in row
        if cell.value is not None
    ).casefold()
    if not any(marker in unit_text for marker in ("usd in thousands", "usd thousands", "$000")):
        return False, "workbook does not explicitly disclose USD-thousands units"

    aliases = [label.replace("_", " "), *_TASK_087_LABEL_ALIASES.get(label, [])]
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            row_text = " | ".join(str(cell.value) for cell in row if cell.value is not None)
            if not any(contains_concept(row_text, alias) for alias in aliases):
                continue
            for cell in row:
                formula = cell.value
                cached = value_sheet[cell.coordinate].value
                if require_formula and not (isinstance(formula, str) and formula.startswith("=")):
                    continue
                if not isinstance(cached, (int, float)) or isinstance(cached, bool):
                    continue
                targets = [float(expected)]
                if abs(float(expected)) > 10:
                    targets.append(float(expected) / 1_000.0)
                if any(
                    _close(cached, target, abs_tol=max(0.00002, abs(target) * 0.000001), rel_tol=0.0)
                    for target in targets
                ):
                    return True, f"{sheet_name}!{cell.coordinate}={formula!r}; cached={cached!r}; disclosed USD thousands"
    return False, "no exact disclosed-USD-thousands formula row matched"


def _task_081_row_match(workbook, values, label: str, expected: Any, *, require_formula: bool) -> tuple[bool, str]:
    """Read optimizer headlines disclosed in the workbook's stated $000 unit.

    Portfolio workbooks commonly place several labeled formula outputs across
    one summary row.  Match the exact expected cached value anywhere on that
    semantically labeled row, while retaining formula-lineage requirements.
    """
    aliases = [label.replace("_", " "), *_TASK_081_LABEL_ALIASES.get(label, [])]
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            if not any(
                any(semantic_equal(cell.value, alias) or contains_concept(cell.value, alias) for alias in aliases)
                for cell in row
            ):
                continue
            for cell in row:
                cached = value_sheet[cell.coordinate].value
                formula = cell.value
                if require_formula and not (isinstance(formula, str) and formula.startswith("=")):
                    continue
                if isinstance(expected, bool):
                    matched = _boolean_matches(cached, expected)
                elif isinstance(expected, str):
                    matched = date_matches(cached, expected) or semantic_value_matches(cached, expected)
                elif isinstance(expected, list):
                    row_values = [value_sheet[cell.coordinate].value for cell in row]
                    matched = (
                        ordered_semantic_list_matches(row_values, expected)
                        or unordered_semantic_list_matches(row_values, expected)
                        or all(any(semantic_value_matches(value, item) for value in row_values) for item in expected)
                    )
                elif isinstance(cached, (int, float)) and not isinstance(cached, bool):
                    targets = [float(expected)]
                    if abs(float(expected)) > 10:
                        targets.append(float(expected) / 1_000.0)
                    matched = any(
                        _close(cached, target, abs_tol=max(0.00002, abs(target) * 0.000001), rel_tol=0.0)
                        for target in targets
                    )
                else:
                    matched = False
                if matched:
                    return True, f"{sheet_name}!{cell.coordinate}={formula!r}; cached={cached!r}"
    return False, "no exact optimizer formula-row value matched"


def _artifact_text_entries(path: Path) -> list[str]:
    suffix = path.suffix.lower()
    entries: list[str] = []
    if suffix == ".xlsx":
        workbook = load_workbook(path, data_only=False, read_only=False)
        for sheet in workbook.worksheets:
            entries.append(sheet.title)
            entries.extend(str(cell.value) for row in sheet.iter_rows() for cell in row if cell.value is not None)
    elif suffix == ".pptx":
        presentation = Presentation(path)
        for slide in presentation.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text:
                    entries.append(shape.text)
                if getattr(shape, "has_table", False):
                    entries.extend(cell.text for row in shape.table.rows for cell in row.cells)
                    entries.extend(" | ".join(cell.text for cell in row.cells) for row in shape.table.rows)
                if getattr(shape, "has_chart", False):
                    # Chart values are visible deliverable content even though
                    # python-pptx does not expose them through ``shape.text``.
                    # Materialize category/series rows so a model is not
                    # penalized merely for choosing an editable chart over a
                    # redundant text box or table.
                    for plot in shape.chart.plots:
                        try:
                            categories = [category.label for category in plot.categories]
                        except (AttributeError, TypeError, ValueError):
                            categories = []
                        for series in plot.series:
                            entries.append(str(series.name))
                            try:
                                values = list(series.values)
                            except (AttributeError, TypeError, ValueError):
                                continue
                            for category, value in zip(categories, values):
                                if value is not None:
                                    entries.append(f"{category} | {series.name}: {value}")
    elif suffix == ".docx":
        document = Document(path)
        entries.extend(paragraph.text for paragraph in document.paragraphs if paragraph.text)
        entries.extend(cell.text for table in document.tables for row in table.rows for cell in row.cells if cell.text)
    return entries


def _artifact_text(path: Path) -> str:
    return "\n".join(_artifact_text_entries(path))


def _semantic_evidence_pack(path: Path, workbook=None, values=None, *, max_chars: int = 60_000) -> str:
    """Render the artifact as compact, grader-side evidence for a bounded judge."""

    chunks: list[str] = [f"ARTIFACT: {path.name}"]
    if path.suffix.lower() == ".docx":
        document = Document(path)
        for index, paragraph in enumerate(document.paragraphs, start=1):
            if paragraph.text.strip():
                chunks.append(f"PARAGRAPH {index}: {paragraph.text.strip()}")
        for table_index, table in enumerate(document.tables, start=1):
            chunks.append(f"TABLE {table_index}:")
            for row_index, row in enumerate(table.rows, start=1):
                chunks.append(
                    f"  ROW {row_index}: " + " | ".join(cell.text.strip() for cell in row.cells)
                )
    elif path.suffix.lower() == ".xlsx" and workbook is not None and values is not None:
        sheet_chunks: list[str] = []
        per_sheet_budget = max(1_000, (max_chars - 1_000) // max(1, len(workbook.sheetnames)))
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
            rendered_rows: list[tuple[int, str, str, str, tuple[Any, ...]]] = []
            for row_number, row in enumerate(sheet.iter_rows(), start=1):
                rendered: list[str] = []
                compact: list[str] = []
                searchable: list[str] = []
                semantic_values: list[Any] = []
                for cell in row:
                    if cell.value is None:
                        continue
                    number_format = str(cell.number_format or "General")
                    format_suffix = (
                        f"[FORMAT={number_format!r}]"
                        if number_format != "General"
                        and any(marker in number_format.casefold() for marker in ("$", "%", "x"))
                        else ""
                    )
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        cached = value_sheet[cell.coordinate].value
                        rendered.append(
                            f"{cell.coordinate}=FORMULA({cell.value})=>{cached!r}{format_suffix}"
                        )
                        compact.append(f"{cell.coordinate}=FORMULA=>{cached!r}{format_suffix}")
                        if cached is not None:
                            searchable.append(str(cached))
                            semantic_values.append(cached)
                    else:
                        rendered.append(f"{cell.coordinate}={cell.value!r}{format_suffix}")
                        compact.append(f"{cell.coordinate}={cell.value!r}{format_suffix}")
                        searchable.append(str(cell.value))
                        semantic_values.append(cell.value)
                if rendered:
                    rendered_rows.append(
                        (
                            row_number,
                            "  " + " | ".join(rendered),
                            " ".join(searchable).casefold(),
                            "  " + " | ".join(compact),
                            tuple(semantic_values),
                        )
                    )
            rendered_sheet = "\n".join(
                [f"SHEET: {sheet_name}", *(row[1] for row in rendered_rows)]
            )
            if len(rendered_sheet) > per_sheet_budget:
                # Large finance engines often put the selected-portfolio flags,
                # recommendation, controls, or source authority between a long
                # formula table and its tail. A plain head/tail truncation can
                # hide exactly the rows the bounded judge must assess. Preserve
                # labeled semantic landmarks and a short following window in a
                # dedicated middle excerpt; deterministic grading still reads
                # the complete workbook.
                landmark_terms = (
                    "selected project", "selected portfolio", "winning combo",
                    "recommendation", "decision", "owner", "timing", "deadline",
                    "model status", "overall status", "control status", "model integrity",
                    "source", "policy", "correspondence", "mcp", "check", "control",
                    "maximum", "variance", "headroom", "unfunded", "conclusion",
                )
                selected_rows: set[int] = set()
                follow_through_terms = ("selected project", "selected portfolio", "winning combo")
                for index, (_, _, searchable, _, _) in enumerate(rendered_rows):
                    if any(term in searchable for term in landmark_terms):
                        selected_rows.update(range(max(0, index - 1), min(len(rendered_rows), index + 2)))
                    if any(term in searchable for term in follow_through_terms):
                        selected_rows.update(range(index, min(len(rendered_rows), index + 11)))

                # A selected/winning ID is commonly shown in a compact control block while
                # the corresponding optimization row is hundreds of lines below it. Preserve
                # that exact row as evidence. This is a general table-association rule, not a
                # task-specific answer key.
                selector_terms = (
                    "selected portfolio id", "selected combination id", "selected row id",
                    "winning combo", "winning combination",
                )

                def selector_key(value: Any) -> str | None:
                    if isinstance(value, bool) or value is None:
                        return None
                    if isinstance(value, (int, float)):
                        numeric = float(value)
                        return str(int(numeric)) if numeric.is_integer() else format(numeric, ".15g")
                    text = str(value).strip().casefold()
                    if not text or any(term in text for term in selector_terms):
                        return None
                    return text if re.fullmatch(r"[a-z]{0,8}-?\d{1,8}", text) else None

                selector_values: set[str] = set()
                for _, _, searchable, _, row_values in rendered_rows:
                    if any(term in searchable for term in selector_terms):
                        selector_values.update(
                            key for value in row_values if (key := selector_key(value)) is not None
                        )
                if selector_values:
                    for index, (_, _, _, _, row_values) in enumerate(rendered_rows):
                        row_keys = {
                            key for value in row_values if (key := selector_key(value)) is not None
                        }
                        if row_keys & selector_values:
                            selected_rows.update(
                                range(max(0, index - 1), min(len(rendered_rows), index + 2))
                            )

                salient = "\n".join(rendered_rows[index][3] for index in sorted(selected_rows))
                salient_budget = int(per_sheet_budget * 0.42)
                if len(salient) > salient_budget:
                    salient = salient[:salient_budget] + "\n[SALIENT ROWS TRUNCATED]"
                head_budget = int(per_sheet_budget * 0.40)
                tail_budget = per_sheet_budget - head_budget - len(salient) - 120
                tail_budget = max(int(per_sheet_budget * 0.12), tail_budget)
                rendered_sheet = (
                    rendered_sheet[:head_budget]
                    + f"\n[SALIENT LABELED ROWS FROM {sheet_name}]\n"
                    + salient
                    + f"\n[TRUNCATED MIDDLE OF {sheet_name}; TAIL PRESERVED]\n"
                    + rendered_sheet[-tail_budget:]
                )
                if len(rendered_sheet) > per_sheet_budget:
                    rendered_sheet = rendered_sheet[:per_sheet_budget]
            sheet_chunks.append(rendered_sheet)
        chunks.extend(sheet_chunks)
    else:
        chunks.extend(_artifact_text_entries(path))
    rendered = "\n".join(chunks)
    if len(rendered) <= max_chars:
        return rendered
    return rendered[: max_chars - 120] + "\n[TRUNCATED AFTER BALANCED DETERMINISTIC EXTRACTION]"


def _workbook_value_present(values, expected: Any, *, formula_workbook=None) -> bool:
    candidates: list[Any]
    cached_candidates = (
        getattr(values, "_alder_all_value_candidates", None)
        if formula_workbook is None else None
    )
    if cached_candidates is not None:
        candidates = cached_candidates
    else:
        candidates = []
        for sheet_name in values.sheetnames:
            value_sheet = values[sheet_name]
            formula_sheet = formula_workbook[sheet_name] if formula_workbook is not None else None
            for row in value_sheet.iter_rows():
                for cell in row:
                    if formula_sheet is not None:
                        formula = formula_sheet[cell.coordinate].value
                        if not (isinstance(formula, str) and formula.startswith("=")):
                            continue
                    candidates.append(cell.value)
        if formula_workbook is None:
            setattr(values, "_alder_all_value_candidates", candidates)
    if isinstance(expected, bool):
        return any(_boolean_matches(candidate, expected) for candidate in candidates)
    if isinstance(expected, str):
        return any(date_matches(candidate, expected) or semantic_value_matches(candidate, expected) for candidate in candidates)
    if isinstance(expected, list):
        return ordered_semantic_list_matches(candidates, expected) or unordered_semantic_list_matches(candidates, expected)
    target = float(expected)
    targets = [target]
    if abs(target) > 10:
        targets.append(target / 1_000.0)
    if abs(target) >= 100_000:
        unit_text = getattr(values, "_alder_unit_text_cache", None)
        if unit_text is None:
            unit_text = "\n".join(
                str(candidate) for candidate in candidates if candidate is not None
            ).casefold()
            setattr(values, "_alder_unit_text_cache", unit_text)
        if any(
            marker in unit_text
            for marker in ("usd millions", "$ in millions", "$ millions", "$mm", "$ mm")
        ):
            targets.append(target / 1_000_000.0)
    return any(
        isinstance(candidate, (int, float))
        and not isinstance(candidate, bool)
        and any(
            _close(candidate, candidate_target, abs_tol=max(0.00002, abs(candidate_target) * 0.000001), rel_tol=0.0)
            for candidate_target in targets
        )
        for candidate in candidates
    )


def _workbook_numeric_targets(values, expected: float) -> list[float]:
    """Return only unit transformations explicitly disclosed by the workbook."""

    target = float(expected)
    targets = [target]
    if abs(target) > 10:
        targets.append(target / 1_000.0)
    unit_text = getattr(values, "_alder_unit_text_cache", None)
    if unit_text is None:
        unit_text = "\n".join(
            str(cell.value)
            for sheet in values.worksheets
            for row in sheet.iter_rows()
            for cell in row
            if cell.value is not None
        ).casefold()
        setattr(values, "_alder_unit_text_cache", unit_text)
    if abs(target) >= 100_000 and any(
        marker in unit_text
        for marker in (
            "usd millions", "$ in millions", "$ millions", "$mm", "$ mm",
            "all dollars in millions", "dollars in millions",
        )
    ):
        targets.append(target / 1_000_000.0)
    return list(dict.fromkeys(targets))


def _workbook_semantic_row_index(workbook, values):
    """Build the reusable row/label index once for all semantic hard gates."""

    cache = getattr(workbook, "_alder_semantic_row_index", None)
    if cache is not None and cache[0] == id(values):
        return cache[1], cache[2]
    rows: list[dict[str, Any]] = []
    label_index: dict[str, list[dict[str, Any]]] = {}
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            numbers: list[float] = []
            labels: list[str] = []
            formulas: list[tuple[str, str, Any]] = []
            for cell in row:
                cached = value_sheet[cell.coordinate].value
                if isinstance(cached, (int, float)) and not isinstance(cached, bool):
                    numbers.append(float(cached))
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formulas.append((cell.coordinate, cell.value, cached))
                elif cell.value is not None:
                    normalized = _normalize(cell.value)
                    if normalized:
                        labels.append(normalized)
            if not numbers and not labels and not formulas:
                continue
            indexed_row = {
                "sheet": sheet_name,
                "row": row[0].row,
                "numbers": tuple(numbers),
                "labels": tuple(labels),
                "formulas": tuple(formulas),
            }
            rows.append(indexed_row)
            for normalized in set(labels):
                label_index.setdefault(normalized, []).append(indexed_row)
    setattr(workbook, "_alder_semantic_row_index", (id(values), rows, label_index))
    return rows, label_index


def _explicit_label_value_conflict(workbook, values, label: str, expected: Any) -> str | None:
    """Reject a directly labeled numeric contradiction without enforcing aliases.

    Semantic review remains responsible for professional-equivalent labels.
    This gate is deliberately narrower: when the artifact itself uses the exact
    requested metric label, that row may not display a different number while
    the expected number happens to occur under another metric elsewhere.
    """

    if isinstance(expected, (bool, str, list)):
        return None
    wanted = _normalize(label.replace("_", " "))
    targets = _workbook_numeric_targets(values, float(expected))
    rows, label_index = _workbook_semantic_row_index(workbook, values)
    labeled_rows: list[tuple[str, int, tuple[float, ...]]] = []
    wanted_tokens = set(wanted.split())

    def row_matches_expected(numbers) -> bool:
        return any(
            _close(
                number,
                target,
                abs_tol=max(0.00002, abs(target) * 0.000001),
                rel_tol=0.0,
            )
            for number in numbers
            for target in targets
        )

    # A more-qualified professional label (for example
    # "probability-weighted run-rate synergy") legitimately resolves a
    # shorter ambiguous row elsewhere.  It must contain every requested
    # metric token and carry the expected value on its own row; mere occurrence
    # of the value elsewhere still does not pass.
    for row_label, indexed_rows in label_index.items():
        if not wanted_tokens <= set(row_label.split()):
            continue
        if any(row_matches_expected(row["numbers"]) for row in indexed_rows):
            return None
    for row in label_index.get(wanted, []):
        if row["numbers"]:
            labeled_rows.append((row["sheet"], row["row"], row["numbers"]))
    if not labeled_rows:
        return None
    for _sheet_name, _row_number, numbers in labeled_rows:
        if row_matches_expected(numbers):
            return None
    rendered = [
        f"{sheet_name}!{row_number}={numbers!r}"
        for sheet_name, row_number, numbers in labeled_rows
    ]
    return (
        f"exact metric label {label!r} displays conflicting numeric row(s) "
        f"{rendered!r}; expected one of {targets!r}"
    )


def _extreme_headline_association_present(
    workbook, values, label: str, expected: Any
) -> tuple[bool, str] | None:
    """Require an explicit max/min association when the requested fact is an extreme."""

    if isinstance(expected, (bool, str, list)):
        return None
    normalized_label = _normalize(label.replace("_", " "))
    if any(word in normalized_label.split() for word in ("maximum", "max", "peak", "highest")):
        direction = "maximum"
        terms = ("maximum", "max", "peak", "highest")
        formula_function = "MAX"
    elif any(word in normalized_label.split() for word in ("minimum", "min", "lowest")):
        direction = "minimum"
        terms = ("minimum", "min", "lowest")
        formula_function = "MIN"
    else:
        return None
    targets = _workbook_numeric_targets(values, float(expected))
    rows, _label_index = _workbook_semantic_row_index(workbook, values)

    def matches(value: Any) -> bool:
        return (
            isinstance(value, (int, float))
            and not isinstance(value, bool)
            and any(
                _close(
                    value,
                    target,
                    abs_tol=max(0.00002, abs(target) * 0.000001),
                    rel_tol=0.0,
                )
                for target in targets
            )
        )

    for row in rows:
        if not any(matches(number) for number in row["numbers"]):
            continue
        row_text = " ".join(row["labels"])
        if any(re.search(rf"\b{re.escape(term)}\b", row_text) for term in terms):
            return True, f"{direction} is explicitly labeled on {row['sheet']}!{row['row']}"
        for coordinate, formula, cached in row["formulas"]:
            if matches(cached) and re.search(
                rf"\b{formula_function}\s*\(", formula, flags=re.I
            ):
                return True, (
                    f"{direction} is formula-identified at "
                    f"{row['sheet']}!{coordinate}={formula}"
                )
    return (
        False,
        f"expected fact is present but is not explicitly identified as a {direction} "
        "by a professional label, selected-extreme row, or MAX/MIN formula",
    )


def _task_048_quarter_index(value: Any) -> int | None:
    """Convert a professional year-quarter label into a sortable index."""

    match = re.search(r"\b(20\d{2})\s*[- /]?\s*q([1-4])\b", str(value or ""), flags=re.I)
    if not match:
        return None
    return int(match.group(1)) * 4 + int(match.group(2)) - 1


def _task_048_decision_support(workbook, values, gold: dict[str, Any]) -> dict[str, Any]:
    """Measure useful covenant decision support without relaxing exact headlines.

    The approved-case and downside headline criteria remain binary strict-pass requirements. This
    companion signal distinguishes a workbook whose detailed schedule contains
    the right decision fact from one that omits the fact entirely.  It is
    intentionally capped at 0.5 for any non-passing headline and is consumed
    only by task_048's partial-reward cap.
    """

    answer = gold["answer"]
    expected_quarter = _task_048_quarter_index(answer["first_covenant_breach"])
    if expected_quarter is None:
        raise ValueError("task_048 gold has an invalid first-covenant-breach quarter")

    def row_text(sheet, row_number: int) -> str:
        return " ".join(
            _normalize(cell.value)
            for cell in sheet[row_number]
            if cell.value is not None and not (
                isinstance(cell.value, str) and cell.value.startswith("=")
            )
        )

    def row_has_concept(text: str, concepts: tuple[str, ...]) -> bool:
        return any(contains_concept(text, concept) for concept in concepts)

    def numeric_row_support(
        expected: float,
        concepts: tuple[str, ...],
        *,
        required_sheet_concept: str | None = None,
    ) -> tuple[float, str]:
        targets = _workbook_numeric_targets(values, expected)
        for sheet_name in workbook.sheetnames:
            if required_sheet_concept and not contains_concept(
                sheet_name,
                required_sheet_concept,
            ):
                continue
            sheet = workbook[sheet_name]
            value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
            for row_number in range(1, sheet.max_row + 1):
                text = row_text(sheet, row_number)
                if not row_has_concept(text, concepts):
                    continue
                for cell in value_sheet[row_number]:
                    candidate = cell.value
                    if not isinstance(candidate, (int, float)) or isinstance(candidate, bool):
                        continue
                    if any(
                        _close(
                            candidate,
                            target,
                            abs_tol=max(0.00002, abs(target) * 0.000001),
                            rel_tol=0.0,
                        )
                        for target in targets
                    ):
                        return (
                            0.5,
                            f"correct fact appears in a labeled supporting schedule at "
                            f"{sheet_name}!{cell.coordinate}",
                        )
        return 0.0, "correct fact is absent from a professionally labeled supporting schedule"

    def quarter_row_support(
        expected: Any,
        *,
        required_sheet_concept: str,
        concepts: tuple[str, ...] = (
            "first covenant breach",
            "first leverage breach",
            "first breach quarter",
            "initial covenant breach",
        ),
    ) -> tuple[float, str]:
        expected_index = _task_048_quarter_index(expected)
        if expected_index is None:
            expected_text = _normalize(expected)
            if expected_text not in {"none", "no breach"}:
                return 0.0, "expected quarter is invalid"
            for sheet_name in workbook.sheetnames:
                if not contains_concept(sheet_name, required_sheet_concept):
                    continue
                sheet = workbook[sheet_name]
                for row_number in range(1, sheet.max_row + 1):
                    text = row_text(sheet, row_number)
                    if row_has_concept(text, concepts) and (
                        contains_concept(text, "no breach")
                        or re.search(r"\bnone\b", text, flags=re.I)
                    ):
                        return 0.5, f"no-breach conclusion appears in a labeled schedule at {sheet_name}!{row_number}"
            return 0.0, f"no-breach conclusion is absent from the {required_sheet_concept} schedule"
        candidates: list[tuple[int, str]] = []
        for sheet_name in workbook.sheetnames:
            if not contains_concept(sheet_name, required_sheet_concept):
                continue
            sheet = workbook[sheet_name]
            value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
            for row_number in range(1, sheet.max_row + 1):
                text = row_text(sheet, row_number)
                if not row_has_concept(text, concepts):
                    continue
                for cell in value_sheet[row_number]:
                    quarter = _task_048_quarter_index(cell.value)
                    if quarter is not None:
                        candidates.append((quarter, f"{sheet_name}!{cell.coordinate}"))
                if row_number < value_sheet.max_row:
                    for cell in value_sheet[row_number + 1]:
                        quarter = _task_048_quarter_index(cell.value)
                        if quarter is not None:
                            candidates.append((quarter, f"{sheet_name}!{cell.coordinate}"))
        if not candidates:
            return 0.0, f"correct breach quarter is absent from the {required_sheet_concept} schedule"
        distance, location = min(
            (
                (abs(quarter - expected_index), location)
                for quarter, location in candidates
            ),
            key=lambda item: item[0],
        )
        if distance <= 1:
            return 0.5, f"breach headline is within one quarter of the expected result at {location}"
        if distance == 2:
            return 0.25, f"breach headline is two quarters from the expected result at {location}"
        return 0.0, f"breach headline is {distance} quarters from the expected result"

    direct_quarters: list[tuple[int, str]] = []
    scheduled_quarters: list[tuple[int, str]] = []
    direct_labels = (
        "first covenant breach",
        "first leverage breach",
        "first breach quarter",
        "initial covenant breach",
    )
    status_labels = (
        "covenant status",
        "leverage status",
        "covenant compliance",
        "breach status",
    )
    breach_markers = ("breach", "fail", "not compliant", "noncompliant")

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row_number in range(1, sheet.max_row + 1):
            text = row_text(sheet, row_number)
            if row_has_concept(text, direct_labels):
                for cell in value_sheet[row_number]:
                    quarter = _task_048_quarter_index(cell.value)
                    if quarter is not None:
                        direct_quarters.append((quarter, f"{sheet_name}!{cell.coordinate}"))
                if row_number < value_sheet.max_row:
                    for cell in value_sheet[row_number + 1]:
                        quarter = _task_048_quarter_index(cell.value)
                        if quarter is not None:
                            direct_quarters.append((quarter, f"{sheet_name}!{cell.coordinate}"))

            if not row_has_concept(text, status_labels):
                continue
            for cell in value_sheet[row_number]:
                marker = _normalize(cell.value)
                if not marker or not any(contains_concept(marker, term) for term in breach_markers):
                    continue
                for header_row in range(row_number - 1, max(0, row_number - 6), -1):
                    quarter = _task_048_quarter_index(value_sheet.cell(header_row, cell.column).value)
                    if quarter is not None:
                        scheduled_quarters.append(
                            (quarter, f"{sheet_name}!{cell.coordinate} under row {header_row}")
                        )
                        break

    if direct_quarters:
        distance, location = min(
            ((abs(quarter - expected_quarter), location) for quarter, location in direct_quarters),
            key=lambda item: item[0],
        )
        if distance == 0:
            breach_score = 0.5
            breach_evidence = f"correct breach quarter appears beside a breach headline at {location}"
        elif distance == 1:
            breach_score = 0.5
            breach_evidence = f"breach headline is one quarter from the executed-definition result at {location}"
        elif distance == 2:
            breach_score = 0.25
            breach_evidence = f"breach headline is two quarters from the executed-definition result at {location}"
        else:
            breach_score = 0.0
            breach_evidence = f"breach headline is {distance} quarters from the executed-definition result"
    elif scheduled_quarters and any(quarter == expected_quarter for quarter, _ in scheduled_quarters):
        location = next(location for quarter, location in scheduled_quarters if quarter == expected_quarter)
        breach_score = 0.4
        breach_evidence = f"correct first breach is identified in the supporting schedule at {location}"
    else:
        breach_score = 0.0
        breach_evidence = "correct first breach is not identified in a headline or status schedule"

    paydown_score, paydown_evidence = numeric_row_support(
        float(answer["required_debt_paydown"]),
        ("debt paydown", "required paydown", "covenant cure", "debt cure"),
    )
    leverage_score, leverage_evidence = numeric_row_support(
        float(answer["maximum_leverage"]),
        ("leverage", "funded debt leverage"),
    )
    coverage_score, coverage_evidence = numeric_row_support(
        float(answer["minimum_fixed_charge_coverage"]),
        ("fixed charge coverage", "fccr"),
    )
    downside_breach_score, downside_breach_evidence = quarter_row_support(
        answer["downside_first_covenant_breach"],
        required_sheet_concept="downside",
    )
    downside_paydown_score, downside_paydown_evidence = numeric_row_support(
        float(answer["downside_required_debt_paydown"]),
        ("debt paydown", "required paydown", "covenant cure", "debt cure"),
        required_sheet_concept="downside",
    )
    downside_leverage_score, downside_leverage_evidence = numeric_row_support(
        float(answer["downside_maximum_leverage"]),
        ("leverage", "funded debt leverage"),
        required_sheet_concept="downside",
    )
    downside_coverage_score, downside_coverage_evidence = numeric_row_support(
        float(answer["downside_minimum_fixed_charge_coverage"]),
        ("fixed charge coverage", "fccr"),
        required_sheet_concept="downside",
    )
    liquidity_breach_score, liquidity_breach_evidence = quarter_row_support(
        answer["first_liquidity_floor_breach"],
        required_sheet_concept="covenant",
        concepts=("first liquidity floor breach", "first cash floor breach", "initial liquidity shortfall"),
    )
    minimum_cash_score, minimum_cash_evidence = numeric_row_support(
        float(answer["minimum_pre_financing_cash"]),
        ("minimum pre financing cash", "minimum cash before financing", "lowest pre financing cash"),
        required_sheet_concept="covenant",
    )
    cure_cash_score, cure_cash_evidence = numeric_row_support(
        float(answer["cash_available_for_cure"]),
        ("cash available for cure", "cash funded cure capacity", "cash cure capacity"),
        required_sheet_concept="covenant",
    )
    cure_shortfall_score, cure_shortfall_evidence = numeric_row_support(
        float(answer["cash_cure_funding_shortfall"]),
        ("cash cure funding shortfall", "cure funding shortfall", "unfunded covenant cure"),
        required_sheet_concept="covenant",
    )
    tangible_net_worth_breach_score, tangible_net_worth_breach_evidence = quarter_row_support(
        answer["first_tangible_net_worth_breach"],
        required_sheet_concept="covenant",
        concepts=("first tangible net worth breach", "first tnw breach", "initial tangible net worth breach"),
    )
    minimum_tangible_net_worth_score, minimum_tangible_net_worth_evidence = numeric_row_support(
        float(answer["minimum_tangible_net_worth"]),
        ("minimum tangible net worth", "minimum tnw", "lowest tangible net worth"),
        required_sheet_concept="covenant",
    )
    downside_liquidity_breach_score, downside_liquidity_breach_evidence = quarter_row_support(
        answer["downside_first_liquidity_floor_breach"],
        required_sheet_concept="downside",
        concepts=("first liquidity floor breach", "first cash floor breach", "initial liquidity shortfall"),
    )
    downside_minimum_cash_score, downside_minimum_cash_evidence = numeric_row_support(
        float(answer["downside_minimum_pre_financing_cash"]),
        ("minimum pre financing cash", "minimum cash before financing", "lowest pre financing cash"),
        required_sheet_concept="downside",
    )
    downside_cure_cash_score, downside_cure_cash_evidence = numeric_row_support(
        float(answer["downside_cash_available_for_cure"]),
        ("cash available for cure", "cash funded cure capacity", "cash cure capacity"),
        required_sheet_concept="downside",
    )
    downside_cure_shortfall_score, downside_cure_shortfall_evidence = numeric_row_support(
        float(answer["downside_cash_cure_funding_shortfall"]),
        ("cash cure funding shortfall", "cure funding shortfall", "unfunded covenant cure"),
        required_sheet_concept="downside",
    )
    downside_tangible_net_worth_breach_score, downside_tangible_net_worth_breach_evidence = quarter_row_support(
        answer["downside_first_tangible_net_worth_breach"],
        required_sheet_concept="downside",
        concepts=("first tangible net worth breach", "first tnw breach", "initial tangible net worth breach"),
    )
    downside_minimum_tangible_net_worth_score, downside_minimum_tangible_net_worth_evidence = numeric_row_support(
        float(answer["downside_minimum_tangible_net_worth"]),
        ("minimum tangible net worth", "minimum tnw", "lowest tangible net worth"),
        required_sheet_concept="downside",
    )

    return {
        "method": "deterministic_professional_schedule_support_v2",
        "policy": (
            "Exact professionally associated headlines are required for strict pass; "
            "non-passing headlines can earn at most 0.5 support credit when the correct "
            "fact is present in a relevant schedule."
        ),
        "criteria": {
            "headline_values__first_covenant_breach": {
                "score": breach_score,
                "evidence": breach_evidence,
            },
            "headline_values__required_debt_paydown": {
                "score": paydown_score,
                "evidence": paydown_evidence,
            },
            "headline_values__maximum_leverage": {
                "score": leverage_score,
                "evidence": leverage_evidence,
            },
            "headline_values__minimum_fixed_charge_coverage": {
                "score": coverage_score,
                "evidence": coverage_evidence,
            },
            "headline_values__first_liquidity_floor_breach": {
                "score": liquidity_breach_score,
                "evidence": liquidity_breach_evidence,
            },
            "headline_values__minimum_pre_financing_cash": {
                "score": minimum_cash_score,
                "evidence": minimum_cash_evidence,
            },
            "headline_values__cash_available_for_cure": {
                "score": cure_cash_score,
                "evidence": cure_cash_evidence,
            },
            "headline_values__cash_cure_funding_shortfall": {
                "score": cure_shortfall_score,
                "evidence": cure_shortfall_evidence,
            },
            "headline_values__first_tangible_net_worth_breach": {
                "score": tangible_net_worth_breach_score,
                "evidence": tangible_net_worth_breach_evidence,
            },
            "headline_values__minimum_tangible_net_worth": {
                "score": minimum_tangible_net_worth_score,
                "evidence": minimum_tangible_net_worth_evidence,
            },
            "downside_values__downside_first_covenant_breach": {
                "score": downside_breach_score,
                "evidence": downside_breach_evidence,
            },
            "downside_values__downside_required_debt_paydown": {
                "score": downside_paydown_score,
                "evidence": downside_paydown_evidence,
            },
            "downside_values__downside_maximum_leverage": {
                "score": downside_leverage_score,
                "evidence": downside_leverage_evidence,
            },
            "downside_values__downside_minimum_fixed_charge_coverage": {
                "score": downside_coverage_score,
                "evidence": downside_coverage_evidence,
            },
            "downside_values__downside_first_liquidity_floor_breach": {
                "score": downside_liquidity_breach_score,
                "evidence": downside_liquidity_breach_evidence,
            },
            "downside_values__downside_minimum_pre_financing_cash": {
                "score": downside_minimum_cash_score,
                "evidence": downside_minimum_cash_evidence,
            },
            "downside_values__downside_cash_available_for_cure": {
                "score": downside_cure_cash_score,
                "evidence": downside_cure_cash_evidence,
            },
            "downside_values__downside_cash_cure_funding_shortfall": {
                "score": downside_cure_shortfall_score,
                "evidence": downside_cure_shortfall_evidence,
            },
            "downside_values__downside_first_tangible_net_worth_breach": {
                "score": downside_tangible_net_worth_breach_score,
                "evidence": downside_tangible_net_worth_breach_evidence,
            },
            "downside_values__downside_minimum_tangible_net_worth": {
                "score": downside_minimum_tangible_net_worth_score,
                "evidence": downside_minimum_tangible_net_worth_evidence,
            },
        },
    }


_FORMULA_REFERENCE_PATTERN = re.compile(
    r"(?:(?:'(?P<quoted>[^']+)'|(?P<plain>[A-Za-z_][A-Za-z0-9_. -]*))!)?"
    r"\$?[A-Z]{1,3}\$?(?P<row_start>\d+)"
    r"(?::\$?[A-Z]{1,3}\$?(?P<row_end>\d+))?"
)


def _scenario_marker(text: str) -> str | None:
    normalized = _normalize(text)
    markers: set[str] = set()
    if any(token in normalized for token in ("downside", "severe downside", "stress")):
        markers.add("downside")
    if "upside" in normalized:
        markers.add("upside")
    if re.search(r"\bbase(?: case)?\b", normalized):
        markers.add("base")
    return next(iter(markers)) if len(markers) == 1 else None


def _scenario_aggregate_formula_miswires(workbook) -> list[str]:
    """Find scenario summaries aggregating a visibly different scenario row.

    This intentionally does not reject ordinary downside formulas that use a
    base-case input.  It targets the much narrower, auditable defect where a
    scenario-labeled summary applies MAX/MIN/SUM/AVERAGE to a row explicitly
    labeled as another scenario, such as a downside maximum pointing at an
    upside free-cash-flow row.
    """

    cached = getattr(workbook, "_alder_scenario_miswire_cache", None)
    if cached is not None:
        return list(cached)
    comparison_terms = (
        "variance", "difference", "delta", " versus ", " vs ",
        "comparison", "bridge", "reconciliation", "sensitivity",
    )
    findings: list[str] = []
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows():
            target_text = " | ".join(
                str(cell.value)
                for cell in row
                if cell.value is not None
                and not (isinstance(cell.value, str) and cell.value.startswith("="))
            )
            target_scenario = _scenario_marker(target_text)
            normalized_target = f" {_normalize(target_text)} "
            if target_scenario is None or any(term in normalized_target for term in comparison_terms):
                continue
            for cell in row:
                formula = cell.value
                if not (
                    isinstance(formula, str)
                    and formula.startswith("=")
                    and re.search(r"\b(?:MAX|MIN|SUM|AVERAGE)\s*\(", formula, flags=re.I)
                ):
                    continue
                for match in _FORMULA_REFERENCE_PATTERN.finditer(formula):
                    referenced_sheet_name = (
                        match.group("quoted") or match.group("plain") or sheet_name
                    )
                    if referenced_sheet_name not in workbook.sheetnames:
                        continue
                    referenced_sheet = workbook[referenced_sheet_name]
                    row_start = int(match.group("row_start"))
                    row_end = int(match.group("row_end") or row_start)
                    if row_end - row_start > 250:
                        continue
                    for referenced_row_number in range(row_start, row_end + 1):
                        source_text = " | ".join(
                            str(source_cell.value)
                            for source_cell in referenced_sheet[referenced_row_number]
                            if source_cell.value is not None
                            and not (
                                isinstance(source_cell.value, str)
                                and source_cell.value.startswith("=")
                            )
                        )
                        source_scenario = _scenario_marker(source_text)
                        if source_scenario is None or source_scenario == target_scenario:
                            continue
                        findings.append(
                            f"{sheet_name}!{cell.coordinate} {target_scenario!r} summary "
                            f"uses {referenced_sheet_name}!{referenced_row_number} "
                            f"labeled {source_scenario!r}: {formula}"
                        )
    setattr(workbook, "_alder_scenario_miswire_cache", tuple(findings))
    return findings


def _task_092_expected_fact_present(path: Path, expected: Any) -> bool:
    if isinstance(expected, (str, bool, list)):
        # Direction, negation, classification, and list association are the
        # bounded judge's job.  Numeric facts remain a mandatory hard gate.
        return True
    text = _artifact_text(path)
    return any(
        _task_092_numeric_literal_matches(literal, float(expected))
        for literal in _TASK_092_NUMERIC_LITERAL_PATTERN.findall(text)
    )


def _numeric_literal_is_compatible_with_label(label: str | None, literal: str) -> bool:
    """Keep a displayed unit from satisfying an unrelated metric type.

    In particular, a rounded monetary literal such as ``$0.00m`` has a wide
    dollar display tolerance.  It must never satisfy a percentage criterion
    merely because the expected decimal happens to fall inside that dollar
    rounding band.
    """

    if not label:
        return True
    label_tokens = set(_normalize(label).split())
    if label_tokens & {"pct", "percent", "percentage"}:
        return "%" in literal
    return True


def _artifact_expected_fact_present(
    path: Path,
    expected: Any,
    *,
    label: str | None = None,
) -> bool:
    """Require objective displayed facts without requiring a brittle label.

    Numeric facts are deterministic hard gates.  Categorical conclusions and
    list membership stay with the bounded semantic judge because their visible
    expression necessarily depends on wording, negation, and table context.
    """

    if isinstance(expected, (str, bool, list)):
        return True
    text = _artifact_text(path)
    return any(
        _numeric_literal_is_compatible_with_label(label, literal)
        and _task_092_numeric_literal_matches(literal, float(expected))
        for literal in _TASK_092_NUMERIC_LITERAL_PATTERN.findall(text)
    )


def _workbook_error_values(values) -> list[str]:
    """Return recalculated spreadsheet errors with stable cell evidence."""

    cached = getattr(values, "_alder_error_values_cache", None)
    if cached is not None:
        return list(cached)
    errors = [
        f"{sheet.title}!{cell.coordinate}={cell.value}"
        for sheet in values.worksheets
        for row in sheet.iter_rows()
        for cell in row
        if cell.data_type == "e"
        or (
            isinstance(cell.value, str)
            and re.fullmatch(r"#(?:REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|NULL!)", cell.value)
        )
    ]
    setattr(values, "_alder_error_values_cache", tuple(errors))
    return errors


_A1_RANGE_PATTERN = re.compile(
    r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)?!?"
    r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)",
    flags=re.I,
)

_A1_REFERENCE_PATTERN = re.compile(
    r"(?:(?:'(?P<quoted>[^']+)'|"
    r"(?P<plain>[A-Za-z_][A-Za-z0-9_. -]*))!)?"
    r"\$?(?P<left_column>[A-Z]{1,3})\$?(?P<left_row>\d+)"
    r"(?::\$?(?P<right_column>[A-Z]{1,3})\$?(?P<right_row>\d+))?",
    flags=re.I,
)


def _column_number(column: str) -> int:
    number = 0
    for character in column.upper():
        number = number * 26 + ord(character) - ord("A") + 1
    return number


def _formula_period_count(formula: str) -> int:
    """Count periods represented by ranges or explicitly enumerated A1 refs.

    Finance controls are commonly written either as one range expression or
    as ``MAX(ABS(...), ABS(...), ...)`` with one equation per period.  Count
    distinct rows/columns on the same referenced sheet so the latter remains
    auditable without treating unrelated scattered cells as weekly coverage.
    """

    singleton_rows_by_sheet: dict[str, set[int]] = {}
    singleton_columns_by_sheet: dict[str, set[int]] = {}
    range_spans: list[int] = []
    for match in _A1_REFERENCE_PATTERN.finditer(formula):
        sheet = (match.group("quoted") or match.group("plain") or "<local>").casefold()
        left_row = int(match.group("left_row"))
        right_row = int(match.group("right_row") or left_row)
        left_column = _column_number(match.group("left_column"))
        right_column = _column_number(match.group("right_column") or match.group("left_column"))
        if match.group("right_row") is not None:
            range_spans.append(
                max(
                    abs(right_row - left_row) + 1,
                    abs(right_column - left_column) + 1,
                )
            )
        else:
            singleton_rows_by_sheet.setdefault(sheet, set()).add(left_row)
            singleton_columns_by_sheet.setdefault(sheet, set()).add(left_column)
    return max(
        [0]
        + range_spans
        + [len(rows) for rows in singleton_rows_by_sheet.values()]
        + [len(columns) for columns in singleton_columns_by_sheet.values()]
    )


def _formula_covers_period_count(formula: str, minimum_periods: int) -> bool:
    return _formula_period_count(formula) >= minimum_periods


def _task_082_control_coverage(workbook) -> tuple[bool, str]:
    """Require complete cash and debt feedback controls for the 13-week model.

    Cash has 13 weekly balances.  Debt feedback has 12 transitions between
    those 13 balances because the first week's beginning debt is an opening
    input, not a prior-week transition.
    """

    required_periods = {
        "cash rollforward": 13,
        "debt feedback": 12,
    }
    failures: list[str] = []
    evidence: list[str] = []
    for control, minimum_periods in required_periods.items():
        found = False
        covered = False
        for sheet in workbook.worksheets:
            for row in range(1, sheet.max_row + 1):
                row_values = [
                    sheet.cell(row, column).value
                    for column in range(1, sheet.max_column + 1)
                ]
                if not any(contains_concept(value, control) for value in row_values):
                    continue
                found = True
                formulas = [
                    str(value)
                    for value in row_values
                    if isinstance(value, str) and value.startswith("=")
                ]
                covered_periods = max(
                    [_formula_period_count(formula) for formula in formulas],
                    default=0,
                )
                if len(formulas) >= minimum_periods or any(
                    _formula_covers_period_count(formula, minimum_periods)
                    for formula in formulas
                ):
                    covered = True
                    evidence.append(
                        f"{control}={sheet.title}!{row} "
                        f"({max(len(formulas), covered_periods)} periods)"
                    )
                    break
            if covered:
                break
        if not found:
            failures.append(f"missing {control} control")
        elif not covered:
            failures.append(
                f"{control} control does not cover all "
                f"{minimum_periods} required periods"
            )
    if failures:
        return False, "; ".join(failures)
    return True, "full-period controls=" + ", ".join(evidence)


def _task_082_label_value_present(
    workbook,
    values,
    label: str,
    expected: Any,
) -> bool:
    """Apply the task's professional label equivalences without weakening values."""

    aliases = [label.replace("_", " "), *_TASK_082_LABEL_ALIASES.get(label, [])]
    if any(_xlsx_label_value(workbook, values, alias, expected) for alias in aliases):
        return True
    if label == "minimum_pre_financing_cash":
        return _task_082_minimum_pre_financing_cash(workbook, values, expected)
    if label == "first_operating_floor_breach_week" and expected in {None, "None"}:
        return _task_082_first_floor_breach(workbook, values, expected)
    if label == "maximum_unfunded_liquidity_shortfall":
        return _task_082_maximum_unfunded_shortfall(
            workbook, values, expected
        )
    return False


def _task_076_funded_debt_association(
    workbook,
    values,
    expected: Any,
) -> tuple[bool, str]:
    """Require the requested debt total, not inference from components."""

    labels = [
        "year end funded debt",
        *_TASK_076_LABEL_ALIASES["year_end_funded_debt"],
        "year end total debt",
        "ending total debt",
        "year end gross debt",
        "ending gross debt",
    ]
    if any(_xlsx_label_value(workbook, values, label, expected) for label in labels):
        return True, "explicit year-end funded/total/gross-debt label is associated with the value"
    return (
        False,
        "year-end funded debt is not explicitly labeled; a term-debt component plus a zero revolver cannot establish the requested total",
    )


def _hybrid_review_hard_gate(
    task_id: str,
    spec: dict[str, Any],
    gold: dict[str, Any],
    path: Path,
    workbook,
    values,
) -> tuple[bool, str]:
    kind = spec["kind"]
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        if spec["id"].endswith("__model_status"):
            errors = _workbook_error_values(values)
            if errors:
                return False, "recalculated spreadsheet errors=" + repr(errors[:8])
            miswires = _scenario_aggregate_formula_miswires(workbook)
            if miswires:
                return (
                    False,
                    "visible cross-scenario aggregate formula miswire(s): "
                    + "; ".join(miswires[:8]),
                )
            if task_id == "task_082":
                controls_met, controls_evidence = _task_082_control_coverage(workbook)
                if not controls_met:
                    return False, controls_evidence
        if kind == "xlsx_label_values":
            missing = [
                label
                for label, expected in spec["label_values"].items()
                if not (
                    _task_082_label_value_present(
                        workbook, values, label, expected
                    )
                    if task_id == "task_082"
                    else _workbook_value_present(values, expected)
                )
            ]
            if missing:
                return False, f"exact cached facts missing anywhere in workbook={missing!r}"
            if task_id == "task_076" and "year_end_funded_debt" in spec["label_values"]:
                associated, association_evidence = _task_076_funded_debt_association(
                    workbook,
                    values,
                    spec["label_values"]["year_end_funded_debt"],
                )
                if not associated:
                    return False, association_evidence
            task_082_equivalent_labels = {
                "minimum_pre_financing_cash",
                "minimum_borrowing_base_headroom",
                "first_operating_floor_breach_week",
                "maximum_unfunded_liquidity_shortfall",
            } if task_id == "task_082" else set()
            extreme_associations = [
                outcome
                for label, expected in spec["label_values"].items()
                if label not in task_082_equivalent_labels
                if (outcome := _extreme_headline_association_present(
                    workbook, values, label, expected
                )) is not None
            ]
            failed_extremes = [evidence for met, evidence in extreme_associations if not met]
            if failed_extremes:
                return False, "; ".join(failed_extremes)
            conflicts = [
                conflict
                for label, expected in spec["label_values"].items()
                if (conflict := _explicit_label_value_conflict(
                    workbook, values, label, expected
                )) is not None
            ]
            return (
                not conflicts,
                (
                    "exact cached facts are present and no exact-label contradiction exists"
                    if not conflicts
                    else "; ".join(conflicts)
                ),
            )
        if kind == "artifact_tokens":
            return True, "workbook parsed; criterion is intentionally semantic"
        if kind == "xlsx_formula_lineage":
            by_sheet = {
                sheet.title: sum(
                    1
                    for row in sheet.iter_rows()
                    for cell in row
                    if isinstance(cell.value, str) and cell.value.startswith("=")
                )
                for sheet in workbook.worksheets
            }
            formulas = [
                cell.value
                for sheet in workbook.worksheets
                for row in sheet.iter_rows()
                for cell in row
                if isinstance(cell.value, str) and cell.value.startswith("=")
            ]
            cross_sheet = sum(
                1
                for formula in formulas
                if re.search(r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)!\$?[A-Z]{1,3}\$?\d+", formula)
            )
            missing_sheets = [
                name for name in spec["formula_sheets"] if by_sheet.get(name, 0) == 0
            ]
            met = (
                not missing_sheets
                and cross_sheet >= int(spec["min_cross_sheet_formulas"])
            )
            return (
                met,
                f"cross_sheet={cross_sheet}/{spec['min_cross_sheet_formulas']}; "
                f"missing_formula_sheets={missing_sheets!r}. "
                "Headline facts are enforced by their separate exact-value gate; "
                "the semantic judge verifies professional association to the formula-driven schedules.",
            )

    if suffix == ".docx":
        if kind == "docx_heading":
            return True, "document parsed; substantive heading equivalence is semantic"
        if kind == "docx_structure":
            tables = len(Document(path).tables)
            return tables >= int(spec["min_tables"]), f"tables={tables}; required={spec['min_tables']}"
        if kind == "artifact_label_values":
            missing = [
                label
                for label, expected in spec["label_values"].items()
                if not _artifact_expected_fact_present(path, expected, label=label)
            ]
            return not missing, f"exact numeric facts missing anywhere in artifact={missing!r}"
        if kind == "artifact_tokens":
            return True, "artifact parsed; criterion is intentionally semantic"

    if suffix == ".pptx":
        if kind == "pptx_title":
            return True, "presentation parsed; substantive title equivalence is semantic"
        if kind == "pptx_structure":
            presentation = Presentation(path)
            expected_slides = spec.get("exact_slides")
            met = (
                len(presentation.slides) == int(expected_slides)
                if expected_slides is not None
                else len(presentation.slides) >= int(spec["min_slides"])
            )
            return met, f"slides={len(presentation.slides)}; required={expected_slides or spec['min_slides']}"
        if kind == "artifact_label_values":
            missing = [
                label
                for label, expected in spec["label_values"].items()
                if not _artifact_expected_fact_present(path, expected, label=label)
            ]
            return not missing, f"exact displayed numeric facts missing anywhere in deck={missing!r}"
        if kind == "artifact_tokens":
            return True, "deck parsed; criterion is intentionally semantic"
    return False, "criterion has no benchmark-wide hybrid hard gate"


def _hybrid_semantic_review(
    task_id: str,
    gold: dict[str, Any],
    path: Path,
    workbook,
    values,
    criteria: list[Criterion],
) -> dict[str, Any] | None:
    by_id = {criterion.id: criterion for criterion in criteria}
    reviews: list[dict[str, Any]] = []
    for spec in gold["criteria"]:
        if not spec.get("semantic"):
            continue
        hard_gate_met, hard_gate_evidence = _hybrid_review_hard_gate(
            task_id, spec, gold, path, workbook, values
        )
        expected_facts = spec.get("label_values")
        if expected_facts is None and spec.get("tokens"):
            expected_facts = {"required_concept": spec["tokens"][0]}
        if expected_facts is None and spec.get("heading"):
            expected_facts = {"required_section": spec["heading"]}
        if expected_facts is None and spec.get("title_token"):
            expected_facts = {"required_title": spec["title_token"]}
        if task_id == "task_092" and spec["id"].startswith("downside_case_values__") and expected_facts and "downside_leverage" in expected_facts:
            # The canonical downside leverage is the closing gross-leverage
            # test.  A professional memo may also show the lower post-Year-1-
            # amortization leverage in the same schedule.  Give the semantic
            # judge the temporal definition so it does not treat two correctly
            # labeled leverage views as a contradiction.
            expected_facts = dict(expected_facts)
            closing_leverage = expected_facts.pop("downside_leverage")
            expected_facts = {
                "closing_original_rate_downside_gross_leverage_before_year_1_amortization": closing_leverage,
                **expected_facts,
            }
        requirement = semantic_requirement(
            criterion_id=spec["id"],
            description=spec["description"],
            expected_facts=expected_facts,
            artifact_type={
                ".xlsx": "workbook",
                ".docx": "memorandum",
                ".pptx": "board presentation",
            }.get(path.suffix.lower(), "artifact"),
        )
        if task_id == "task_087" and (
            spec["id"].startswith("closing_funds_flow_values__")
            or spec["id"].startswith("model_content__")
        ):
            requirement += (
                " For the requested formula-driven zero funds-flow check, a professional "
                "Checks schedule showing formula-linked Gross closing cash uses and Net buyer "
                "funding requirement actual-versus-expected rows with zero differences and OK "
                "statuses is sufficient; do not require a literal label named 'funds-flow check'."
            )
        reviews.append(
            {
                "criterion_id": spec["id"],
                "requirement": requirement,
                "hard_gate_met": hard_gate_met,
                "hard_gate_evidence": hard_gate_evidence,
                "legacy_lexical_match": bool(by_id[spec["id"]].met),
            }
        )
    if not reviews:
        return None
    return {
        "version": 2,
        "mode": "deterministic_hard_gates_plus_bounded_semantic_judge",
        "task_id": task_id,
        "artifact": gold["artifact"]["path"],
        "evidence": _semantic_evidence_pack(path, workbook, values),
        "criteria": reviews,
        "policy": (
            "A criterion passes only when its deterministic hard gate passes and the semantic judge "
            "finds the professional-language association or narrative substance MET."
        ),
    }


def _artifact_label_value(entries: list[str], label: str, expected: Any) -> bool:
    for index, entry in enumerate(entries):
        normalized = _normalize(entry)
        if not contains_concept(entry, label):
            continue
        candidates: list[Any] = [entry]
        candidates.extend(entries[index + 1:index + 4])
        if isinstance(expected, list):
            if ordered_semantic_list_matches(candidates, expected):
                return True
            continue
        if isinstance(expected, str):
            if any(date_matches(candidate, expected) or semantic_value_matches(candidate, expected) for candidate in candidates):
                return True
            continue
        if isinstance(expected, bool):
            if any(_boolean_matches(candidate, expected) for candidate in candidates):
                return True
            continue
        numbers = []
        for candidate in candidates:
            # Do not let optional unit suffixes cross a newline and consume
            # the first letter of the next label (for example ``10.77\nb...``
            # being misread as 10.77 billion).
            numbers.extend(re.findall(r"\(?-?\$?\d[\d,]*(?:\.\d+)?[ \t]*(?:%|[kmbx×])?\)?", str(candidate), flags=re.I))
        def literal_tolerance(number: str) -> float:
            cleaned = number.strip().replace("$", "").replace(",", "").strip("() ")
            suffix = cleaned[-1:].casefold()
            multiplier = {"k": 1_000.0, "m": 1_000_000.0, "b": 1_000_000_000.0}.get(suffix, 1.0)
            if suffix in {"k", "m", "b", "x", "×"}:
                cleaned = cleaned[:-1].strip()
            percent = cleaned.endswith("%")
            cleaned = cleaned.rstrip("%").strip()
            decimals = len(cleaned.rsplit(".", 1)[1]) if "." in cleaned else 0
            tolerance = 0.5 * (10 ** -decimals) * multiplier
            return tolerance / 100 if percent else tolerance

        if any(
            _close(
                number,
                float(expected),
                abs_tol=max(_display_tolerance(float(expected)), literal_tolerance(number) + 1e-12),
                rel_tol=0.0,
            )
            for number in numbers
        ):
            return True
    return False



_TASK_092_SCENARIO_BINDINGS = {
    "base_leverage": ("base", ("gross leverage", "leverage")),
    "downside_leverage": ("downside", ("gross leverage", "leverage")),
    "base_debt_service_coverage": ("base", ("debt service coverage", "debt-service coverage", "dscr")),
    "downside_debt_service_coverage": ("downside", ("debt service coverage", "debt-service coverage", "dscr")),
    "downside_cash_headroom": ("downside", ("cash headroom", "liquidity headroom", "liquidity shortfall")),
    "stressed_downside_debt_service_coverage": ("stress", ("debt service coverage", "debt-service coverage", "dscr")),
    "stressed_annual_cash_interest": ("stress", ("acquisition interest", "cash interest", "interest")),
}
_TASK_092_NUMERIC_LITERAL_PATTERN = re.compile(
    r"""
    (?<![A-Za-z0-9])
    (?:-?\$?[ \t]*\(?|\([ \t]*\$?-?[ \t]*)
    \d[\d,]*(?:\.\d+)?[ \t]*\)?
    (?:[ \t]*(?:
        %
        |[x×]
        |k(?![A-Za-z])
        |thousand(?![A-Za-z])
        |m(?:m|illion)?(?![A-Za-z])
        |b(?:n|illion)?(?![A-Za-z])
    ))?
    [ \t]*\)?
    """,
    flags=re.I | re.X,
)


def _task_092_numeric_literal_matches(literal: str, expected: float) -> bool:
    candidates = list(iter_numeric_candidates(literal))
    accounting_negative = "(" in literal and ")" in literal
    explicit_negative = "-" in literal.partition(next((c for c in literal if c.isdigit()), ""))[0]
    percent = literal.rstrip().rstrip(")").rstrip().endswith("%")
    cleaned_for_precision = (
        literal.strip().replace("$", "").replace(",", "")
        .replace("(", "").replace(")", "").strip()
    )
    scale_match = re.search(
        r"(?:\s*)(thousand|k|million|mm|m|billion|bn|b)\s*$",
        cleaned_for_precision,
        flags=re.I,
    )
    scale_token = scale_match.group(1).casefold() if scale_match else ""
    display_multiplier = {
        "thousand": 1_000.0,
        "k": 1_000.0,
        "million": 1_000_000.0,
        "mm": 1_000_000.0,
        "m": 1_000_000.0,
        "billion": 1_000_000_000.0,
        "bn": 1_000_000_000.0,
        "b": 1_000_000_000.0,
    }.get(scale_token, 1.0)
    if scale_match:
        cleaned_for_precision = cleaned_for_precision[: scale_match.start()].strip()
    cleaned_for_precision = cleaned_for_precision.rstrip("%x×").strip()
    try:
        displayed_value = float(cleaned_for_precision) * display_multiplier
        if accounting_negative or explicit_negative:
            displayed_value = -abs(displayed_value)
        if percent:
            displayed_value /= 100.0
        candidates.append(displayed_value)
    except ValueError:
        pass
    decimals = len(cleaned_for_precision.rsplit(".", 1)[1]) if "." in cleaned_for_precision else 0
    literal_tolerance = 0.5 * (10 ** -decimals) * display_multiplier
    if percent:
        literal_tolerance /= 100.0
    if any(
        _close(
            candidate,
            expected,
            abs_tol=max(_display_tolerance(expected), literal_tolerance + 1e-12),
            rel_tol=0.0,
        )
        for candidate in candidates
    ):
        return True

    # IC tables use a single "$mm" header and compact dollar cells. Infer that
    # display unit only for monetary headlines; ratios and multiples stay raw.
    cleaned = literal.strip().replace("$", "").replace(",", "").strip("() ")
    if abs(expected) <= 10_000 or cleaned[-1:].casefold() in {"k", "m", "b", "%", "x", "×"}:
        return False
    try:
        displayed = float(cleaned)
    except ValueError:
        return False
    if accounting_negative or explicit_negative:
        displayed = -abs(displayed)
    scaled = displayed * 1_000_000.0
    return _close(
        scaled,
        expected,
        abs_tol=max(5_000.0, _display_tolerance(expected)),
        rel_tol=0.0,
    )


def _task_092_scenario_kind(text: str) -> str | None:
    normalized = _normalize(text)
    has_stress = "stress" in normalized
    has_downside = "downside" in normalized
    has_base = bool(re.search(r"\bbase(?: case)?\b", normalized))
    if has_stress:
        return "stress"
    if has_base and has_downside:
        return "mixed"
    if has_downside:
        return "downside"
    if has_base:
        return "base"
    return None


def _task_092_scenario_chunks(text: str) -> list[str]:
    return [
        chunk.strip()
        for chunk in re.split(
            r"(?i)(?=\b(?:rate[- ]stressed downside|downside stressed|stressed downside|"
            r"original[- ]rate downside|base(?: case)?|downside(?: case)?|stressed)\b)",
            text,
        )
        if chunk.strip()
    ]


def _task_092_scenario_value(document: Document, label: str, expected: float) -> bool | None:
    """Bind scenario-specific metrics to their own row or matrix column.

    Return None only when the memo has no explicit scenario presentation for
    this metric, allowing the normal standalone-label matcher to run.
    """
    binding = _TASK_092_SCENARIO_BINDINGS.get(label)
    if binding is None:
        return None
    target_scenario, metric_aliases = binding
    explicit_presentation_seen = False

    def metric_matches(text: str) -> bool:
        return any(contains_concept(text, alias) for alias in metric_aliases)

    def literals_match(text: str) -> bool:
        return any(
            _task_092_numeric_literal_matches(literal, float(expected))
            for literal in _TASK_092_NUMERIC_LITERAL_PATTERN.findall(text)
        )

    for table in document.tables:
        rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
        if not rows:
            continue

        header_scenarios = [_task_092_scenario_kind(cell) for cell in rows[0]]
        target_columns = [
            index for index, scenario in enumerate(header_scenarios)
            if scenario == target_scenario
        ]
        if target_columns:
            for row in rows[1:]:
                if not row or not metric_matches(row[0]):
                    continue
                explicit_presentation_seen = True
                if any(index < len(row) and literals_match(row[index]) for index in target_columns):
                    return True

        for row in rows:
            if not row or not metric_matches(row[0]):
                continue
            row_scenario = _task_092_scenario_kind(row[0])
            if row_scenario is not None:
                explicit_presentation_seen = True
                if row_scenario == target_scenario and literals_match(" | ".join(row[1:])):
                    return True
                continue

            scenario_chunks = [
                chunk
                for cell in row[1:]
                for chunk in _task_092_scenario_chunks(cell)
                if _task_092_scenario_kind(chunk) is not None
            ]
            if scenario_chunks:
                explicit_presentation_seen = True
                if any(
                    _task_092_scenario_kind(chunk) == target_scenario and literals_match(chunk)
                    for chunk in scenario_chunks
                ):
                    return True

    for paragraph in document.paragraphs:
        if not metric_matches(paragraph.text):
            continue
        scenario_chunks = [
            chunk
            for chunk in _task_092_scenario_chunks(paragraph.text)
            if _task_092_scenario_kind(chunk) is not None
        ]
        if not scenario_chunks:
            continue
        explicit_presentation_seen = True
        if any(
            _task_092_scenario_kind(chunk) == target_scenario and literals_match(chunk)
            for chunk in scenario_chunks
        ):
            return True

    return False if explicit_presentation_seen else None

def _task_092_artifact_value(path: Path, label: str, expected: Any) -> bool:
    """Read normal IC-memo tables, including tables presented in USD millions.

    The generic artifact matcher cannot infer that a cell containing ``$12.00``
    means $12 million when the unit appears once in the table header.  This
    task-specific reader keeps the accepted labels narrow and compares values
    only within the matching row or paragraph.
    """
    aliases = [label.replace("_", " "), *_TASK_092_LABEL_ALIASES.get(label, [])]
    document = Document(path)
    segments = [paragraph.text for paragraph in document.paragraphs if paragraph.text]
    table_rows: list[list[str]] = []
    contextual_table_rows: list[str] = []
    for table in document.tables:
        rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
        table_rows.extend(rows)
        if rows:
            # A professional memo normally states scenario/metric labels once
            # in the header and puts the values in subsequent rows.  Preserve
            # that local header context so correct matrix-style tables grade
            # equivalently to repeated standalone labels.
            header = rows[0]
            contextual_table_rows.extend(" | ".join([*header, *row]) for row in rows[1:])
    segments.extend(" | ".join(cells) for cells in table_rows)
    segments.extend(contextual_table_rows)
    if isinstance(expected, bool):
        if label == "stressed_financing_case_compliant" and expected is False:
            normalized = _normalize("\n".join(segments))
            if any(phrase in normalized for phrase in (
                "overall guardrail result fail",
                "downside fails all three gates",
                "reject current structure",
                "reject as presented",
                "current structure should not close",
                "structure cannot proceed",
                "withhold approval",
                "no approval on the current capital stack",
                "all three required tests fail",
                "all three year 2 compliance tests fail",
                "stressed conclusion fail",
                "no feasible draw within the commitment satisfies all gates",
            )):
                return True
        return any(
            any(contains_concept(segment, alias) for alias in aliases)
            and _boolean_matches(segment, expected)
            for segment in segments
        )
    if isinstance(expected, str):
        if label == "recommendation" and contains_concept(expected, "do not approve"):
            normalized = _normalize("\n".join(segments))
            if any(phrase in normalized for phrase in (
                "reject current structure",
                "reject as currently financed",
                "reject as presented",
                "reject the financing structure as presented",
                "do not authorize",
                "current structure should not close",
                "withhold approval",
                "no approval on the current capital stack",
            )):
                return True
        return any(
            any(contains_concept(segment, alias) for alias in aliases)
            and semantic_value_matches(segment, expected)
            for segment in segments
        )
    scenario_match = _task_092_scenario_value(document, label, float(expected))
    if scenario_match is not None:
        return scenario_match
    if float(expected) == 0.0:
        financial_zero_tokens = {"-", "–", "—", "$-", "$–", "$—"}
        for cells in table_rows:
            if not cells or not any(contains_concept(cells[0], alias) for alias in aliases):
                continue
            if any(cell.strip() in financial_zero_tokens for cell in cells[1:]):
                return True
    for segment in segments:
        if not any(contains_concept(segment, alias) for alias in aliases):
            continue
        if any(
            _task_092_numeric_literal_matches(literal, float(expected))
            for literal in _TASK_092_NUMERIC_LITERAL_PATTERN.findall(segment)
        ):
            return True
    return False


def _task_100_artifact_value(path: Path, label: str, expected: Any) -> bool:
    """Grade a board deck by meaning, including editable chart data and $m units.

    Board materials normally state units once at the slide or deck level and
    then show compact values such as ``10.77`` or ``3.17x``.  The generic
    artifact reader cannot safely infer those units across every task, so this
    reader is intentionally scoped to task 100 and its controlled labels.
    """
    aliases = [label.replace("_", " "), *_TASK_100_LABEL_ALIASES.get(label, [])]
    entries = _artifact_text_entries(path)
    text = _normalize("\n".join(entries))

    if isinstance(expected, list):
        if label == "optimized_board_priority_portfolio":
            # The board template lists the two selected initiatives as
            # separate shapes beneath the portfolio heading.
            start = next((i for i, entry in enumerate(entries) if contains_concept(entry, "optimized decision portfolio")), None)
            if start is None:
                return False
            stop = next((i for i in range(start + 1, len(entries)) if contains_concept(entries[i], "constraints")), min(len(entries), start + 24))
            selected_entries = entries[start:stop]
            return all(any(semantic_value_matches(entry, item) for entry in selected_entries) for item in expected)
        for index, entry in enumerate(entries):
            if not any(contains_concept(entry, alias) for alias in aliases):
                continue
            candidates = entries[index:index + max(4, len(expected) + 2)]
            if unordered_semantic_list_matches(candidates[1:1 + len(expected)], expected):
                return True
            if ordered_semantic_list_matches(candidates, expected):
                return True
        return False

    if isinstance(expected, bool):
        if label == "severe_covenant_breach":
            severe_is_flagged = bool(
                re.search(r"severe.{0,240}breach", text)
                or re.search(r"breach.{0,240}severe", text)
                or ("breach flags" in text and "severe" in text and "executed leverage" in text)
            )
            return severe_is_flagged if expected else not severe_is_flagged
        if label == "post_action_severe_cash_breach":
            flagged = bool(re.search(r"breach.{0,120}cash\s*<", text) or re.search(r"cash.{0,120}breach", text))
            return flagged if expected else not flagged
        if label == "post_action_severe_leverage_breach":
            flagged = bool(re.search(r"breach.{0,180}leverage\s*>", text) or re.search(r"leverage.{0,120}breach", text))
            return flagged if expected else not flagged
        return any(
            any(contains_concept(segment, alias) for alias in aliases)
            and _boolean_matches(segment, expected)
            for segment in entries
        )

    if isinstance(expected, str):
        if label == "largest_branch_ebitda_miss":
            # Accept a narrative callout or a clearly negative Controls row;
            # merely mentioning Controls elsewhere in the deck is not enough.
            named_callout = bool(
                re.search(r"controls.{0,180}(?:miss|negative|below)", text)
                or re.search(r"(?:miss|negative|below).{0,180}controls", text)
            )
            negative_row = any(
                contains_concept(segment, expected)
                and bool(re.search(r"(?:\(\s*\$?\d|[-−]\s*\$?\d)", segment))
                for segment in entries
            )
            return named_callout or negative_row
        return any(
            any(contains_concept(segment, alias) for alias in aliases)
            and semantic_value_matches(segment, expected)
            for segment in entries
        )

    pattern = re.compile(r"\(?[-−]?\$?\d[\d,]*(?:\.\d+)?[ \t]*(?:%|[kmbx×])?\)?", flags=re.I)
    for index, entry in enumerate(entries):
        if not any(contains_concept(entry, alias) for alias in aliases):
            continue
        # Slide templates often put a label and value in adjacent text boxes.
        # Keep the window narrow so values from unrelated cards cannot leak in.
        segment = " | ".join(entries[index:index + 4])
        for literal in pattern.findall(segment):
            normalized_literal = literal.replace("−", "-")
            candidates = list(iter_numeric_candidates(normalized_literal))
            precision_text = normalized_literal.strip().replace("$", "").replace(",", "").strip("() ")
            precision_suffix = precision_text[-1:].casefold()
            precision_multiplier = {"k": 1_000.0, "m": 1_000_000.0, "b": 1_000_000_000.0}.get(precision_suffix, 1.0)
            if precision_suffix in {"k", "m", "b", "x", "×"}:
                precision_text = precision_text[:-1].strip()
            precision_text = precision_text.rstrip("%").strip()
            precision_decimals = len(precision_text.rsplit(".", 1)[1]) if "." in precision_text else 0
            precision_tolerance = 0.5 * (10 ** -precision_decimals) * precision_multiplier
            if normalized_literal.strip().endswith("%"):
                precision_tolerance /= 100.0
            if any(_close(candidate, float(expected), abs_tol=max(_display_tolerance(float(expected)), precision_tolerance + 1e-12), rel_tol=0.0) for candidate in candidates):
                return True
            cleaned = normalized_literal.strip().replace("$", "").replace(",", "").strip("() ")
            suffix = cleaned[-1:].casefold()
            if abs(float(expected)) <= 10_000 or suffix in {"k", "m", "b", "%", "x", "×"}:
                continue
            try:
                displayed = float(cleaned)
            except ValueError:
                continue
            if normalized_literal.strip().startswith("("):
                displayed = -abs(displayed)
            scaled = displayed * 1_000_000.0
            decimals = len(cleaned.rsplit(".", 1)[1]) if "." in cleaned else 0
            display_tolerance = 0.5 * (10 ** -decimals) * 1_000_000.0 + 1e-9
            if _close(scaled, float(expected), abs_tol=max(display_tolerance, _display_tolerance(float(expected))), rel_tol=0.0):
                return True
    return False


def _task_098_xlsx_value(workbook, values, label: str, expected: Any) -> bool:
    """Accept formula-driven KPI headlines displayed in the stated $m unit."""
    aliases = [label.replace("_", " "), *_TASK_098_LABEL_ALIASES.get(label, [])]
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            for cell in row:
                if not any(semantic_equal(cell.value, alias) or contains_concept(cell.value, alias) for alias in aliases):
                    continue
                candidates = [
                    value_sheet.cell(cell.row, cell.column + offset).value
                    for offset in (1, 2, 3)
                    if cell.column + offset <= value_sheet.max_column
                ]
                candidates.append(value_sheet.cell(cell.row + 1, cell.column).value)
                if _value_candidates_match(candidates, expected):
                    return True
                if isinstance(expected, (int, float)) and not isinstance(expected, bool) and abs(float(expected)) > 10_000:
                    for candidate in candidates:
                        if isinstance(candidate, (int, float)) and not isinstance(candidate, bool):
                            if _close(float(candidate) * 1_000_000.0, float(expected), abs_tol=_display_tolerance(float(expected)), rel_tol=0.0):
                                return True
    return False


def _task_037_selected_projects(workbook, values, expected: list[str]) -> bool:
    selected: list[str] = []
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in range(1, sheet.max_row + 1):
            row_values = [value_sheet.cell(row, column).value for column in range(1, min(sheet.max_column, 12) + 1)]
            status = next((value for value in row_values if semantic_equal(value, "select") or semantic_equal(value, "selected")), None)
            if status is None:
                continue
            project = next((value for value in row_values if re.fullmatch(r"CP-\d{2}", str(value or ""), flags=re.I)), None)
            if project and str(project).upper() not in selected:
                selected.append(str(project).upper())
    return unordered_semantic_list_matches(selected, expected)


def _task_081_selected_initiatives(workbook, values, expected: list[str]) -> bool:
    """Accept a normal selection matrix instead of demanding a TEXTJOIN cell."""
    selected: list[str] = []
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        selection_columns: list[tuple[int, int]] = []
        for row in range(1, min(sheet.max_row, 30) + 1):
            for column in range(1, sheet.max_column + 1):
                value = sheet.cell(row, column).value
                if any(semantic_equal(value, label) for label in ("Selected flag", "Selected?", "Portfolio decision")):
                    selection_columns.append((row, column))
        for header_row, selection_column in selection_columns:
            for row in range(header_row + 1, sheet.max_row + 1):
                row_values = [value_sheet.cell(row, column).value for column in range(1, min(sheet.max_column, 16) + 1)]
                initiative = next((str(value) for value in row_values if any(semantic_equal(value, name) for name in expected)), None)
                decision = value_sheet.cell(row, selection_column).value
                chosen = (
                    semantic_equal(decision, "selected") or semantic_equal(decision, "select")
                    or decision is True
                    or (isinstance(decision, (int, float)) and not isinstance(decision, bool) and float(decision) == 1.0)
                )
                if initiative and chosen and not any(semantic_equal(initiative, item) for item in selected):
                    selected.append(initiative)
    return unordered_semantic_list_matches(selected, expected)


def _task_081_selection_formula(workbook) -> tuple[bool, str]:
    for sheet in workbook.worksheets:
        for row in range(1, min(sheet.max_row, 20) + 1):
            for cell in sheet[row]:
                if not (semantic_equal(cell.value, "selected") or semantic_equal(cell.value, "select") or contains_concept(cell.value, "selection")):
                    continue
                formulas = [sheet.cell(candidate_row, cell.column).value for candidate_row in range(row + 1, sheet.max_row + 1)]
                if any(isinstance(value, str) and value.startswith("=") and re.search(r"[A-Z]{1,3}\$?\d+", value) for value in formulas):
                    return True, f"formula-driven selection column at {sheet.title}!{cell.coordinate}"
    return False, "no formula-driven initiative-selection column"


def _task_082_minimum_pre_financing_cash(workbook, values, expected: Any) -> bool:
    """Accept the formula-driven weekly cash row when the minimum is not repeated in a summary box."""
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in range(1, sheet.max_row + 1):
            if not any(semantic_equal(sheet.cell(row, column).value, "Cash before financing") for column in range(1, sheet.max_column + 1)):
                continue
            candidates = [value_sheet.cell(row, column).value for column in range(1, sheet.max_column + 1)]
            numeric = [float(value) for value in candidates if isinstance(value, (int, float)) and not isinstance(value, bool)]
            if numeric and _numeric_matches(min(numeric), float(expected), 0.02):
                return True
    return False


def _task_082_first_floor_breach(workbook, values, expected: Any) -> bool:
    """Accept an explicit zero-breach control as equivalent to no first week.

    A well-designed cash model may summarize the result as a formula-driven
    breach count instead of a redundant text cell containing ``None``.
    """
    if _normalize(expected) != "none":
        return False
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            for cell in row:
                if not (
                    semantic_equal(cell.value, "Operating-floor breaches")
                    or semantic_equal(cell.value, "Floor breach count")
                    or semantic_equal(cell.value, "Floor breach weeks")
                    or semantic_equal(cell.value, "Floor-breach weeks")
                ):
                    continue
                candidates = [
                    value_sheet.cell(cell.row, cell.column + offset).value
                    for offset in (1, 2, 3)
                    if cell.column + offset <= value_sheet.max_column
                ]
                if any(isinstance(value, (int, float)) and not isinstance(value, bool) and abs(float(value)) <= 1e-9 for value in candidates):
                    return True
    return False


def _task_082_maximum_unfunded_shortfall(
    workbook,
    values,
    expected: Any,
) -> bool:
    """Prove a zero maximum from a complete formula-driven weekly series.

    When every one of the 13 weekly unfunded-liquidity outputs is zero, both
    the total and the maximum are necessarily zero. This accepts that
    professional representation without accepting a stray zero elsewhere.
    """

    if not isinstance(expected, (int, float)) or isinstance(expected, bool):
        return False
    if not _close(float(expected), 0.0, abs_tol=1e-9, rel_tol=0.0):
        return False
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            for header in row:
                if not (
                    semantic_equal(header.value, "Unfunded liquidity")
                    or contains_concept(header.value, "Unfunded liquidity")
                ):
                    continue
                formula_values: list[float] = []
                for row_number in range(header.row + 1, sheet.max_row + 1):
                    formula = sheet.cell(row_number, header.column).value
                    cached = value_sheet.cell(row_number, header.column).value
                    if not (isinstance(formula, str) and formula.startswith("=")):
                        continue
                    if isinstance(cached, (int, float)) and not isinstance(cached, bool):
                        formula_values.append(float(cached))
                if (
                    len(formula_values) >= 13
                    and all(abs(value) <= 0.02 for value in formula_values)
                ):
                    return True
    return False


def _task_066_matrix_value(workbook, values, label: str, expected: Any) -> bool:
    aliases = [label.replace("_", " ")]
    if label == "ltm_revenue": aliases += ["Revenue"]
    if label == "organic_growth": aliases += ["Organic growth"]
    if label == "adjusted_ebitda": aliases += ["Adjusted EBITDA"]
    if label == "adjusted_ebitda_margin": aliases += ["Adjusted EBITDA margin"]
    if label == "backlog": aliases += ["Signed backlog"]
    if label == "free_cash_flow": aliases += ["Free cash flow"]
    if label == "leverage": aliases += ["Gross leverage"]
    if label == "fixed_charge_coverage": aliases += ["Investor fixed-charge coverage", "Investor FCCR"]
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in range(1, sheet.max_row + 1):
            label_cell = next((cell for cell in sheet[row] if any(semantic_equal(cell.value, alias) for alias in aliases)), None)
            if not label_cell:
                continue
            row_values = [value_sheet.cell(row, column).value for column in range(1, sheet.max_column + 1)]
            if label in {"ltm_revenue", "adjusted_ebitda", "free_cash_flow"}:
                # The fact sheet is an eight-quarter matrix. These three gold
                # headlines are trailing-four-quarter totals, so derive them
                # from the latest four columns instead of demanding a
                # redundant standalone scalar label.
                numbers = [value_sheet.cell(row, column).value for column in range(7, 11)]
                if all(isinstance(value, (int, float)) for value in numbers) and _numeric_matches(sum(numbers), float(expected), _display_tolerance(float(expected))):
                    return True
            elif label == "adjusted_ebitda_margin":
                revenue_row = next(
                    (candidate for candidate in range(1, sheet.max_row + 1)
                     if _normalize(sheet.cell(candidate, 1).value) == _normalize("Revenue")),
                    None,
                )
                ebitda_row = next(
                    (candidate for candidate in range(1, sheet.max_row + 1)
                     if _normalize(sheet.cell(candidate, 1).value) == _normalize("Adjusted EBITDA")),
                    None,
                )
                if revenue_row and ebitda_row:
                    revenue = [value_sheet.cell(revenue_row, column).value for column in range(7, 11)]
                    ebitda = [value_sheet.cell(ebitda_row, column).value for column in range(7, 11)]
                    if all(isinstance(value, (int, float)) for value in revenue + ebitda):
                        margin = sum(ebitda) / sum(revenue)
                        if _numeric_matches(margin, float(expected), _display_tolerance(float(expected))):
                            return True
            elif _value_candidates_match(row_values, expected):
                return True
    return False


def _task_066_matrix_formula(workbook, label: str) -> tuple[bool, str]:
    aliases = [label.replace("_", " ")]
    aliases.extend({
        "ltm_revenue": ["Revenue"],
        "backlog": ["Signed backlog"],
        "leverage": ["Gross leverage"],
        "fixed_charge_coverage": ["Investor fixed-charge coverage", "Investor FCCR"],
    }.get(label, []))
    for sheet in workbook.worksheets:
        for row in range(1, sheet.max_row + 1):
            if not any(semantic_equal(sheet.cell(row, 1).value, alias) for alias in aliases):
                continue
            columns = range(7, 11) if label in {"ltm_revenue", "adjusted_ebitda", "adjusted_ebitda_margin", "free_cash_flow"} else (10,)
            formulas = [sheet.cell(row, column) for column in columns]
            if formulas and all(isinstance(cell.value, str) and cell.value.startswith("=") for cell in formulas):
                return True, f"matrix formulas at {sheet.title}!{formulas[0].coordinate}:{formulas[-1].coordinate}"
    return False, "no formula-driven matrix row"


def _preserved_sheets(output, seed, names: list[str]) -> tuple[bool, str]:
    for name in names:
        if name not in output.sheetnames or name not in seed.sheetnames:
            return False, f"missing protected sheet {name!r}"
        left, right = output[name], seed[name]
        max_row = max(left.max_row, right.max_row)
        max_col = max(left.max_column, right.max_column)
        for row in range(1, max_row + 1):
            for column in range(1, max_col + 1):
                left_value = left.cell(row, column).value
                right_value = right.cell(row, column).value
                if isinstance(left_value, (int, float)) and isinstance(right_value, (int, float)):
                    same = math.isclose(float(left_value), float(right_value), rel_tol=1e-12, abs_tol=1e-9)
                else:
                    same = left_value == right_value
                if not same:
                    return False, f"protected source changed at {name}!{left.cell(row, column).coordinate}"
    return True, "protected source sheets match the seeded template"


def _file_failure(gold: dict[str, Any], reason: str) -> dict[str, Any]:
    result = _result([Criterion(spec["id"], spec["description"], False, reason) for spec in gold["criteria"]])
    return _attach_gold_policy(result, gold)


def _attach_gold_policy(result: dict[str, Any], gold: dict[str, Any]) -> dict[str, Any]:
    by_spec = {str(spec["id"]): spec for spec in gold["criteria"]}
    for row in result.get("criteria", []):
        spec = by_spec[str(row["id"])]
        for key in ("category", "weight", "failure_cap", "semantic", "kind"):
            if key in spec:
                row[key] = spec[key]
    return result


def _headline_formula_match(task_id: str, workbook, values, label: str, gold: dict[str, Any]) -> tuple[bool, str]:
    candidates = [label.replace("_", " ")]
    if task_id == "task_031":
        candidates.extend(_TASK_031_LABEL_ALIASES.get(label, []))
    if task_id == "task_037":
        candidates.extend(_TASK_037_LABEL_ALIASES.get(label, []))
    if task_id == "task_053":
        candidates.extend(_TASK_053_LABEL_ALIASES.get(label, []))
    if task_id == "task_061":
        candidates.extend(_TASK_061_LABEL_ALIASES.get(label, []))
    if task_id == "task_076":
        candidates.extend(_TASK_076_LABEL_ALIASES.get(label, []))
    if task_id == "task_081":
        candidates.extend(_TASK_081_LABEL_ALIASES.get(label, []))
    if task_id == "task_082":
        candidates.extend(_TASK_082_LABEL_ALIASES.get(label, []))
    if task_id == "task_087":
        candidates.extend(_TASK_087_LABEL_ALIASES.get(label, []))
    if task_id == "task_098":
        candidates.extend(_TASK_098_LABEL_ALIASES.get(label, []))
    attempts = [_xlsx_label_formula(workbook, candidate) for candidate in candidates]
    matched, detail = next((attempt for attempt in attempts if attempt[0]), attempts[0])
    if task_id == "task_040" and not matched:
        row_column = {
            "fy31_free_cash_flow": ("Free cash flow before debt service", "FY31"),
            "fy31_debt": ("Ending debt", "FY31"),
            "fy31_cash": ("Ending cash", "FY31"),
        }
        if label in row_column:
            matched, detail = _xlsx_row_column_formula(
                workbook, *row_column[label]
            )
        elif label == "downside_peak_financing_plug":
            matched, detail = _xlsx_section_row_formula(
                workbook,
                "Downside scenario summary",
                "Maximum single-year financing plug",
            )
    if task_id == "task_055" and not matched:
        row_column = {
            "year_one_gaap_eps_accretion": (
                "GAAP EPS accretion / (dilution) %",
                "Year 1",
            ),
            "year_one_adjusted_eps_accretion": (
                "Adjusted EPS accretion / (dilution) %",
                "Year 1",
            ),
            "year_two_gaap_eps_accretion": (
                "GAAP EPS accretion / (dilution) %",
                "Year 2",
            ),
            "year_two_adjusted_eps_accretion": (
                "Adjusted EPS accretion / (dilution) %",
                "Year 2",
            ),
        }
        if label in row_column:
            matched, detail = _xlsx_row_column_formula(
                workbook, *row_column[label]
            )
    if task_id == "task_081" and label == "binding_constraint" and not matched:
        matched, detail = _task_081_binding_constraint_formula(workbook)
    if task_id == "task_076" and not matched:
        matched, detail = _task_076_row_match(workbook, values, label, gold["answer"][label], require_formula=True)
    if task_id == "task_081" and not matched:
        matched, detail = _task_081_row_match(workbook, values, label, gold["answer"][label], require_formula=True)
    if task_id == "task_087" and not matched:
        matched, detail = _task_087_row_match(workbook, values, label, gold["answer"][label], require_formula=True)
    if task_id == "task_053" and not matched:
        matched, detail = _task_053_row_match(
            workbook,
            values,
            label,
            gold["answer"][label],
            require_formula=True,
        )
    if task_id == "task_066" and not matched:
        matched, detail = _task_066_matrix_formula(workbook, label)
    if task_id == "task_081" and label == "selected_initiatives" and not matched:
        matched, detail = _task_081_selection_formula(workbook)
    return matched, detail


def _grade_artifact(task_id: str, workspace_root: Path) -> dict[str, Any]:
    gold = load_corporate_finance_gold(task_id)
    relative = Path(gold["artifact"]["path"])
    path = workspace_root / relative
    if not path.is_file():
        return _file_failure(gold, f"missing required artifact: {relative}")
    try:
        entries = _artifact_text_entries(path)
        normalized_text = _normalize("\n".join(entries))
        workbook = values = None
        if path.suffix.lower() == ".xlsx":
            workbook = load_workbook(path, data_only=False, read_only=False)
            values = _recalculated_data_workbook(path)
    except Exception as exc:
        return _file_failure(gold, f"unreadable artifact: {exc}")

    criteria: list[Criterion] = []
    for spec in gold["criteria"]:
        kind = spec["kind"]
        evidence = ""
        if kind == "xlsx_sheet_present":
            met = workbook is not None and spec["sheet"] in workbook.sheetnames
            evidence = f"required_sheet={spec['sheet']!r}; sheets={getattr(workbook, 'sheetnames', [])!r}"
        elif kind == "xlsx_sheet_preserved":
            seed_path = SEED_WORKSPACE / relative
            if workbook is None or not seed_path.is_file():
                met = False
                evidence = "submitted workbook or seeded edit template is missing"
            else:
                seed = load_workbook(seed_path, data_only=False, read_only=False)
                met, evidence = _preserved_sheets(workbook, seed, [spec["sheet"]])
        elif kind == "xlsx_formula_count":
            formulas = sum(
                1 for sheet in workbook.worksheets for row in sheet.iter_rows() for cell in row
                if isinstance(cell.value, str) and cell.value.startswith("=")
            )
            met = formulas >= int(spec["min_formulas"])
            evidence = f"formulas={formulas}; required={spec['min_formulas']}"
        elif kind == "xlsx_no_errors":
            errors = [
                f"{sheet.title}!{cell.coordinate}={cell.value}" for sheet in values.worksheets
                for row in sheet.iter_rows() for cell in row if cell.data_type == "e"
            ]
            met = not errors
            evidence = f"errors={errors[:10]!r}"
        elif kind == "xlsx_headline_formula":
            met, evidence = _headline_formula_match(
                task_id, workbook, values, str(spec["headline_label"]), gold
            )
        elif kind == "xlsx_formula_sheet":
            sheet = workbook[spec["sheet"]] if spec["sheet"] in workbook.sheetnames else None
            formulas = 0 if sheet is None else sum(
                1 for row in sheet.iter_rows() for cell in row
                if isinstance(cell.value, str) and cell.value.startswith("=")
            )
            if (
                task_id == "task_081"
                and spec["sheet"] == "Constraints"
                and sheet is not None
                and formulas == 0
            ):
                met, input_evidence = _task_081_constraints_input_sheet(sheet)
                evidence = (
                    f"sheet={spec['sheet']!r}; formulas=0; "
                    f"sourced input schedule={input_evidence}"
                )
            else:
                met = formulas > 0
                evidence = f"sheet={spec['sheet']!r}; formulas={formulas}"
        elif kind == "xlsx_cross_sheet_formulas":
            formulas = [
                cell.value for sheet in workbook.worksheets for row in sheet.iter_rows() for cell in row
                if isinstance(cell.value, str) and cell.value.startswith("=")
            ]
            cross_sheet = sum(
                1 for formula in formulas
                if re.search(r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)!\$?[A-Z]{1,3}\$?\d+", formula)
            )
            met = cross_sheet >= int(spec["min_cross_sheet_formulas"])
            evidence = f"cross_sheet={cross_sheet}; required={spec['min_cross_sheet_formulas']}"
        elif kind == "pptx_slide_count":
            presentation = Presentation(path)
            expected = spec.get("exact_slides")
            met = len(presentation.slides) == int(expected) if expected is not None else len(presentation.slides) >= int(spec["min_slides"])
            evidence = f"slides={len(presentation.slides)}; required={expected or spec.get('min_slides')}"
        elif kind == "pptx_title":
            met = contains_concept("\n".join(entries), spec["title_token"])
            evidence = f"title_concept={spec['title_token']!r}; present={met}"
        elif kind == "docx_heading":
            body = "\n".join(paragraph.text for paragraph in Document(path).paragraphs)
            met = contains_concept(body, spec["heading"])
            evidence = f"heading_concept={spec['heading']!r}; present={met}"
        elif kind == "docx_tables":
            tables = len(Document(path).tables)
            met = tables >= int(spec["min_tables"])
            evidence = f"tables={tables}; required={spec['min_tables']}"
        elif kind == "xlsx_structure":
            required = spec["sheets"]
            sheets_ok = workbook is not None and all(name in workbook.sheetnames for name in required)
            preserve_ok, preserve_evidence = (True, "not an edit task")
            if sheets_ok and spec.get("preserve_source_sheets"):
                seed_path = SEED_WORKSPACE / relative
                if not seed_path.is_file():
                    preserve_ok, preserve_evidence = False, "seeded edit template is missing"
                else:
                    seed = load_workbook(seed_path, data_only=False, read_only=False)
                    preserve_ok, preserve_evidence = _preserved_sheets(workbook, seed, spec["preserve_source_sheets"])
            met = sheets_ok and preserve_ok
            evidence = f"sheets={getattr(workbook, 'sheetnames', [])!r}; {preserve_evidence}"
        elif kind == "xlsx_model_integrity":
            formulas = sum(
                1 for sheet in workbook.worksheets for row in sheet.iter_rows() for cell in row
                if isinstance(cell.value, str) and cell.value.startswith("=")
            )
            errors = [
                f"{sheet.title}!{cell.coordinate}={cell.value}" for sheet in values.worksheets
                for row in sheet.iter_rows() for cell in row
                if cell.data_type == "e"
            ]
            met = formulas >= spec["min_formulas"] and not errors
            evidence = f"formulas={formulas}; errors={errors[:5]!r}"
        elif kind == "xlsx_formula_lineage":
            by_sheet = {
                sheet.title: sum(
                    1 for row in sheet.iter_rows() for cell in row
                    if isinstance(cell.value, str) and cell.value.startswith("=")
                )
                for sheet in workbook.worksheets
            }
            all_formulas = [
                cell.value for sheet in workbook.worksheets for row in sheet.iter_rows() for cell in row
                if isinstance(cell.value, str) and cell.value.startswith("=")
            ]
            cross_sheet = sum(
                1 for formula in all_formulas
                if re.search(r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)!\$?[A-Z]{1,3}\$?\d+", formula)
            )
            missing_formula_sheets = [name for name in spec["formula_sheets"] if by_sheet.get(name, 0) == 0]
            headline_failures = []
            headline_evidence = []
            for label in spec["headline_labels"]:
                candidates = [label.replace("_", " ")]
                if task_id == "task_031":
                    candidates.extend(_TASK_031_LABEL_ALIASES.get(label, []))
                if task_id == "task_037":
                    candidates.extend(_TASK_037_LABEL_ALIASES.get(label, []))
                if task_id == "task_053":
                    candidates.extend(_TASK_053_LABEL_ALIASES.get(label, []))
                if task_id == "task_076":
                    candidates.extend(_TASK_076_LABEL_ALIASES.get(label, []))
                if task_id == "task_081":
                    candidates.extend(_TASK_081_LABEL_ALIASES.get(label, []))
                if task_id == "task_082":
                    candidates.extend(_TASK_082_LABEL_ALIASES.get(label, []))
                if task_id == "task_087":
                    candidates.extend(_TASK_087_LABEL_ALIASES.get(label, []))
                if task_id == "task_098":
                    candidates.extend(_TASK_098_LABEL_ALIASES.get(label, []))
                attempts = [_xlsx_label_formula(workbook, candidate) for candidate in candidates]
                matched, detail = next((attempt for attempt in attempts if attempt[0]), attempts[0])
                if task_id == "task_076" and not matched:
                    expected = gold["answer"][label]
                    matched, detail = _task_076_row_match(workbook, values, label, expected, require_formula=True)
                if task_id == "task_081" and not matched:
                    expected = gold["answer"][label]
                    matched, detail = _task_081_row_match(workbook, values, label, expected, require_formula=True)
                if task_id == "task_087" and not matched:
                    expected = gold["answer"][label]
                    matched, detail = _task_087_row_match(workbook, values, label, expected, require_formula=True)
                if task_id == "task_053" and not matched:
                    expected = gold["answer"][label]
                    matched, detail = _task_053_row_match(
                        workbook,
                        values,
                        label,
                        expected,
                        require_formula=True,
                    )
                if task_id == "task_066" and not matched:
                    matched, detail = _task_066_matrix_formula(workbook, label)
                if task_id == "task_081" and label == "selected_initiatives" and not matched:
                    matched, detail = _task_081_selection_formula(workbook)
                if not matched:
                    headline_failures.append(label)
                else:
                    headline_evidence.append(detail)
            met = (
                not missing_formula_sheets
                and not headline_failures
                and cross_sheet >= spec["min_cross_sheet_formulas"]
            )
            evidence = (
                f"formula_sheets={by_sheet!r}; cross_sheet={cross_sheet}; "
                f"missing_formula_sheets={missing_formula_sheets!r}; "
                f"headline_failures={headline_failures!r}; headline_examples={headline_evidence[:3]!r}"
            )
        elif kind == "xlsx_label_values":
            failures = []
            for label, expected in spec["label_values"].items():
                labels = [label.replace("_", " ")]
                if task_id == "task_076":
                    labels.extend(_TASK_076_LABEL_ALIASES.get(label, []))
                if task_id == "task_081":
                    labels.extend(_TASK_081_LABEL_ALIASES.get(label, []))
                if task_id == "task_082":
                    labels.extend(_TASK_082_LABEL_ALIASES.get(label, []))
                if task_id == "task_087":
                    labels.extend(_TASK_087_LABEL_ALIASES.get(label, []))
                if task_id == "task_053":
                    labels.extend(_TASK_053_LABEL_ALIASES.get(label, []))
                if task_id == "task_098":
                    labels.extend(_TASK_098_LABEL_ALIASES.get(label, []))
                matched = any(_xlsx_label_value(workbook, values, candidate, expected) for candidate in labels)
                if task_id == "task_076" and not matched:
                    matched, _ = _task_076_row_match(workbook, values, label, expected, require_formula=False)
                if task_id == "task_081" and not matched:
                    matched, _ = _task_081_row_match(workbook, values, label, expected, require_formula=False)
                if task_id == "task_087" and not matched:
                    matched, _ = _task_087_row_match(workbook, values, label, expected, require_formula=False)
                if task_id == "task_053" and not matched:
                    matched, _ = _task_053_row_match(
                        workbook,
                        values,
                        label,
                        expected,
                        require_formula=False,
                    )
                if task_id == "task_037":
                    if label == "selected_portfolio" and isinstance(expected, list):
                        matched = matched or _task_037_selected_projects(workbook, values, expected)
                    else:
                        matched = matched or any(
                            _xlsx_label_value(workbook, values, alias, expected)
                            for alias in _TASK_037_LABEL_ALIASES.get(label, [])
                        )
                if task_id == "task_066":
                    matched = matched or _task_066_matrix_value(workbook, values, label, expected)
                if task_id == "task_081" and label == "selected_initiatives" and isinstance(expected, list):
                    matched = matched or _task_081_selected_initiatives(workbook, values, expected)
                if task_id == "task_082" and not matched:
                    matched = _task_082_label_value_present(
                        workbook, values, label, expected
                    )
                if task_id == "task_082" and label == "minimum_pre_financing_cash":
                    matched = matched or _task_082_minimum_pre_financing_cash(workbook, values, expected)
                if task_id == "task_082" and label == "first_operating_floor_breach_week":
                    matched = matched or _task_082_first_floor_breach(workbook, values, expected)
                if task_id == "task_098" and not matched:
                    matched = _task_098_xlsx_value(workbook, values, label, expected)
                if not matched:
                    failures.append(label)
            met = not failures
            evidence = "all labeled values matched" if met else f"missing or incorrect labels={failures!r}"
        elif kind == "task_037_selection_tie":
            met, evidence = _task_037_selection_tie_control(workbook, values)
        elif kind == "artifact_label_values":
            failures = []
            for label, expected in spec["label_values"].items():
                labels = [label.replace("_", " ")]
                if task_id == "task_072":
                    labels.extend(_TASK_072_LABEL_ALIASES.get(label, []))
                if task_id == "task_075":
                    labels.extend(_TASK_075_LABEL_ALIASES.get(label, []))
                if task_id == "task_092":
                    labels.extend(_TASK_092_LABEL_ALIASES.get(label, []))
                if task_id == "task_100":
                    labels.extend(_TASK_100_LABEL_ALIASES.get(label, []))
                matched = any(_artifact_label_value(entries, candidate, expected) for candidate in labels)
                if task_id == "task_092" and not matched:
                    matched = _task_092_artifact_value(path, label, expected)
                if task_id == "task_100" and not matched:
                    matched = _task_100_artifact_value(path, label, expected)
                if not matched:
                    failures.append(label)
            met = not failures
            evidence = "all labeled values matched" if met else f"missing or incorrect labels={failures!r}"
        elif kind == "artifact_tokens":
            missing = [
                token for token in spec["tokens"]
                if not any(
                    contains_concept(normalized_text, candidate)
                    for candidate in _ARTIFACT_TOKEN_ALIASES.get(token, [token])
                )
            ]
            met = not missing
            evidence = "all required content present" if met else f"missing={missing!r}"
        elif kind == "pptx_structure":
            presentation = Presentation(path)
            expected_slides = spec.get("exact_slides")
            count_ok = len(presentation.slides) == expected_slides if expected_slides is not None else len(presentation.slides) >= spec["min_slides"]
            # Preserve punctuation such as ``&`` until the semantic normalizer
            # sees it; pre-normalizing through the legacy artifact helper can
            # erase the conjunction before concept matching.
            title_present = contains_concept("\n".join(entries), spec["title_token"])
            met = count_ok and title_present
            evidence = f"slides={len(presentation.slides)}; title_present={title_present}"
        elif kind == "docx_structure":
            document = Document(path)
            body = _normalize("\n".join(paragraph.text for paragraph in document.paragraphs))
            missing = [heading for heading in spec["headings"] if not contains_concept(body, heading)]
            met = not missing and len(document.tables) >= spec["min_tables"]
            evidence = f"missing_headings={missing!r}; tables={len(document.tables)}"
        else:
            raise ValueError(f"Unsupported artifact criterion: {kind}")
        criteria.append(Criterion(spec["id"], spec["description"], met, evidence))
    result = _attach_gold_policy(_result(criteria), gold)
    if task_id == "task_048":
        result["decision_support"] = _task_048_decision_support(
            workbook,
            values,
            gold,
        )
    semantic_review = _hybrid_semantic_review(
        task_id, gold, path, workbook, values, criteria
    )
    if semantic_review is not None:
        result["semantic_review"] = semantic_review
    return result


def grade_corporate_finance_task(task_id: str, answer: Any, workspace_root: str | Path) -> dict[str, Any]:
    gold = load_corporate_finance_gold(task_id)
    if "artifact" in gold:
        return _grade_artifact(task_id, Path(workspace_root))
    return _grade_console(task_id, answer)
