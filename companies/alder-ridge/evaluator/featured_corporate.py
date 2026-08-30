from __future__ import annotations

import json
import math
import re
from pathlib import Path
from typing import Any, Iterable

from docx import Document
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pptx import Presentation

from evaluator.featured_apex import (
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
from evaluator.featured_semantic import (
    contains_concept,
    date_matches,
    iter_numeric_candidates,
    ordered_semantic_list_matches,
    semantic_equal,
    semantic_value_matches,
    unordered_semantic_list_matches,
)
from evaluator.featured_hybrid_semantic import semantic_requirement


GOLD_PATH = Path(__file__).resolve().parent / "reference" / "tasks_026_100.json"

_TASK_031_LABEL_ALIASES = {
    "peak_hiring_need": ["Peak cumulative hiring need"],
    "labor_cost_variance_to_plan": [
        "Variance to approved FY27 field-labor plan",
        "Labor cost variance to approved plan",
    ],
}

_TASK_035_LABEL_ALIASES = {
    "signed_backlog": ["Signed backlog", "Portfolio signed backlog"],
    "fy27_backlog_revenue_burn": [
        "Approved probability plan",
        "Probability-plan revenue",
        "FY27 backlog revenue",
    ],
    "fy27_required_hours": [
        "Approved probability plan",
        "Probability-plan required hours",
    ],
    "fy27_available_hours": ["FY27 available hours", "Base available hours"],
    "constrained_month_count": ["Constrained months", "Months constrained"],
    "first_constrained_month": [
        "First constrained month",
        "Probability-plan first constrained month",
    ],
    "maximum_capacity_shortfall_hours": [
        "Maximum capacity shortfall",
        "Probability-plan maximum shortfall",
    ],
    "probability_plan_unresolved_hours": [
        "Approved probability plan",
        "Probability-plan unresolved hours",
    ],
    "probability_plan_remediation_cost": [
        "Approved probability plan",
        "Probability-plan remediation cost",
    ],
    "gross_commitment_revenue_burn": [
        "Full signed commitment",
        "Gross-commitment revenue",
    ],
    "gross_commitment_required_hours": [
        "Full signed commitment",
        "Gross-commitment required hours",
    ],
    "gross_commitment_first_constrained_month": [
        "Full signed commitment",
        "Gross-commitment first constrained month",
    ],
    "gross_commitment_maximum_capacity_shortfall_hours": [
        "Gross-commitment maximum shortfall",
        "Full-commitment maximum shortfall",
    ],
    "gross_commitment_unresolved_hours": [
        "Full signed commitment",
        "Gross-commitment unresolved hours",
    ],
    "gross_commitment_remediation_cost": [
        "Full signed commitment",
        "Gross-commitment remediation cost",
    ],
    "gross_commitment_revenue_variance_to_probability_plan": [
        "Gross-commitment revenue variance",
        "Revenue variance to probability plan",
    ],
    "gross_commitment_required_hours_variance_to_probability_plan": [
        "Gross-commitment required-hours variance",
        "Required-hours variance to probability plan",
    ],
    "gross_commitment_unresolved_hours_variance_to_probability_plan": [
        "Gross-commitment unresolved-hours variance",
        "Unresolved-hours variance to probability plan",
    ],
    "gross_commitment_remediation_cost_variance_to_probability_plan": [
        "Gross-commitment remediation-cost variance",
        "Remediation-cost variance to probability plan",
    ],
    "gross_commitment_release_status": [
        "Full signed commitment",
        "Gross-commitment release conclusion",
    ],
    "controlling_capacity_case": ["Controlling case", "Capacity case controlling"],
    "execution_portfolio_completed_revenue": [
        "Portfolio completed revenue",
        "Execution completed revenue",
    ],
    "execution_portfolio_ending_deferred_revenue": [
        "Portfolio deferred revenue",
        "Ending deferred revenue",
    ],
    "execution_portfolio_revenue_check_delta": [
        "Portfolio revenue check",
        "Revenue bridge check",
    ],
    "execution_portfolio_ending_deferred_hours": [
        "Portfolio deferred hours",
        "Ending deferred hours",
    ],
    "execution_portfolio_completed_gross_profit": [
        "Portfolio completed gross profit",
        "Completed gross profit",
    ],
    "execution_portfolio_liquidated_damages": [
        "Portfolio liquidated damages",
        "Liquidated damages",
    ],
    "execution_portfolio_net_gross_profit_after_damages": [
        "Portfolio net gross profit",
        "Executable gross profit after damages",
    ],
    "execution_portfolio_highest_priority_deferred_project": [
        "Highest-priority deferred project",
        "Priority deferred item",
    ],
    "execution_portfolio_release_status": [
        "Portfolio release decision",
        "Execution release status",
    ],
    "release_bridge_probability_plan_gross_profit_after_remediation": [
        "Probability-plan GP after response",
        "Probability-plan GP after remediation",
    ],
    "release_bridge_gross_commitment_gross_profit_after_remediation": [
        "Gross-commitment GP after response",
        "Gross-commitment GP after remediation",
    ],
    "release_bridge_execution_net_gross_profit_after_damages": [
        "Executable GP after damages",
        "Execution net GP after damages",
    ],
    "release_bridge_deferred_revenue_at_risk": [
        "Deferred revenue at risk",
        "Deferred revenue",
    ],
    "release_bridge_management_decision": [
        "Management decision",
        "Release decision",
    ],
    "executive_recovery_probability_plan_gp_after_remediation": [
        "Probability-plan GP",
        "Probability-plan GP after response",
    ],
    "executive_recovery_gross_commitment_gp_after_remediation": [
        "Gross-commitment GP",
        "Gross-commitment GP after response",
    ],
    "executive_recovery_execution_net_gp_after_damages": [
        "Executable GP after damages",
        "Execution net GP after damages",
    ],
    "executive_recovery_gross_commitment_shortfall": [
        "Gross-commitment shortfall",
        "Earnings shortfall",
    ],
    "executive_recovery_deferred_revenue_at_risk": [
        "Deferred revenue",
        "Deferred revenue at risk",
    ],
    "executive_recovery_deferred_gross_profit_at_risk": [
        "Deferred gross profit",
        "Deferred gross profit at risk",
    ],
    "executive_recovery_liquidated_damages": ["Damages", "Liquidated damages"],
    "executive_recovery_required_recovery": [
        "Recovery required",
        "Required recovery",
    ],
    "executive_recovery_highest_priority_deferred_project": [
        "Priority item",
        "Highest-priority deferred project",
    ],
    "executive_recovery_decision": ["Recommendation", "Executive decision"],
}

_TASK_037_LABEL_ALIASES = {
    "selected_portfolio": ["Selected ID", "Selected portfolio", "Portfolio membership"],
    "selected_capex": ["Cash capex", "Cash capex used", "Selected cash capex"],
    "portfolio_npv": ["Selected NPV", "Portfolio NPV", "Max feasible portfolio NPV"],
    "downside_portfolio_npv": ["Downside portfolio NPV"],
    "selected_debt_eligible_basis": ["Eligible debt used", "Debt-eligible basis"],
    "selected_cash_funding": ["Cash funding used", "Selected cash funding"],
    "selected_technician_capacity": ["Technicians used", "Technician capacity used"],
    "cash_capex_headroom": ["Cash headroom", "Cash capex headroom"],
    "debt_capacity_headroom": ["Debt headroom", "Debt capacity headroom"],
    "technician_capacity_headroom": ["Technician headroom", "Technician capacity headroom"],
    "mandatory_safety_projects_selected": ["Mandatory safety selected", "Mandatory projects selected"],
    "cash_constraint_check": ["Cash constraint", "Cash capex check"],
    "debt_constraint_check": ["Debt constraint", "Debt capacity check"],
    "technician_constraint_check": ["Technician constraint", "Technician capacity check"],
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
    "acquisition_capacity": ["Acquisition - Pioneer HVAC", "Acquisition", "Identified acquisition", "Pioneer HVAC commitment"],
    "total_allocated": ["Total recommended uses / reserves", "Recommended program", "Total recommended deployment", "Total recommended uses"],
    "remaining_capacity": ["Remaining unused capacity", "Headroom remaining", "Remaining FY27 capacity headroom"],
    "estimated_value_creation": ["Value creation", "committee-estimated value"],
    "base_ending_cash_before_mitigation": ["Base"],
    "downside_minimum_cash": ["Downside cash before mitigation", "Downside"],
    "downside_cash_headroom_to_minimum": ["Downside (controlling)", "Downside"],
    "downside_cash_headroom_to_distribution_buffer": ["Downside-prudence", "Downside"],
    "severe_downside_ending_cash_before_mitigation": ["Severe downside"],
    "severe_cash_shortfall_to_minimum": ["Severe downside", "Below minimum cash"],
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

_TASK_076_QUARTER_METRIC_ALIASES = {
    "revenue": ["Revenue"],
    "ebitda": ["EBITDA"],
    "beginning_cash": ["Beginning cash", "Opening cash"],
    "ending_nwc": ["Ending NWC", "Ending net working capital"],
    "change_in_nwc": ["Change in NWC", "Net working capital change", "Increase in NWC"],
    "cash_taxes": ["Cash taxes", "Cash tax"],
    "capex": ["Capex", "Capital expenditures"],
    "term_interest": ["Term interest", "Term debt interest"],
    "revolver_interest": ["Revolver interest"],
    "total_interest": ["Total interest", "Cash interest"],
    "scheduled_term_amortization": ["Scheduled term amortization", "Scheduled principal", "Term debt amortization"],
    "pre_financing_cash": ["Pre-financing cash", "Cash before financing", "Pre-revolver cash"],
    "revolver_draw": ["Revolver draw", "Draw"],
    "revolver_repayment": ["Revolver repayment", "Repayment"],
    "ending_cash": ["Ending cash", "Ending unrestricted cash"],
    "beginning_term_debt": ["Beginning term debt", "Opening term debt"],
    "ending_term_debt": ["Ending term debt"],
    "beginning_revolver": ["Beginning revolver", "Opening revolver"],
    "ending_revolver": ["Ending revolver"],
    "ending_funded_debt": ["Ending funded debt", "Ending total debt", "Ending gross debt"],
}

_TASK_076_BRANCH_METRIC_ALIASES = {
    "fy26_baseline_revenue": ["FY26 baseline revenue", "FY26 revenue", "Baseline revenue"],
    "fy27_revenue": ["FY27 revenue", "Revenue"],
    "gross_profit": ["Gross profit"],
    "fixed_opex": ["Fixed opex", "Fixed operating expense"],
    "variable_opex": ["Variable opex", "Variable operating expense"],
    "ebitda_before_corporate": ["EBITDA before corporate", "Branch EBITDA", "EBITDA contribution"],
}

_TASK_040_LABEL_ALIASES = {
    "fy31_free_cash_flow": ["Free cash flow before debt service"],
    "fy31_debt": ["Ending debt"],
    "fy31_cash": ["Ending cash"],
    "downside_peak_financing_plug": ["Maximum single-year financing plug"],
}

_TASK_055_LABEL_ALIASES = {
    "buyer_standalone_eps": ["Standalone diluted EPS"],
    "seller_shares_issued": ["Seller shares issued (equity ÷ price)"],
    "pro_forma_diluted_shares": ["Pro forma diluted shares"],
    "target_ebit": ["Target EBIT"],
    "total_incremental_financing_cost": ["Total incremental financing cost"],
    "purchase_enterprise_value": ["Purchase enterprise value", "Enterprise value"],
    "total_transaction_sources": ["Total sources", "Sources"],
    "total_transaction_uses": ["Total uses", "Uses"],
    "sources_and_uses_check": ["Sources less uses", "Sources - uses", "Check: total sources minus total uses"],
    "incremental_debt_interest": ["Incremental new-debt interest", "Incremental interest (new debt)"],
    "foregone_cash_yield": ["Foregone cash yield (interest)", "Foregone cash yield"],
    "year_one_gaap_eps_accretion": ["GAAP EPS accretion / (dilution) %"],
    "year_one_adjusted_eps_accretion": ["Adjusted EPS accretion / (dilution) %"],
    "year_two_gaap_eps_accretion": ["GAAP EPS accretion / (dilution) %"],
    "year_two_adjusted_eps_accretion": ["Adjusted EPS accretion / (dilution) %"],
}

_TASK_055_YEAR_ROW_ALIASES = {
    "run_rate_synergy": ["Cost synergies", "Run-rate synergy"],
    "integration_expense": ["One-time integration expense"],
    "gaap_incremental_pre_tax_income": ["Target-side pre-tax income (GAAP)"],
    "gaap_incremental_after_tax_income": ["Target-side after-tax (GAAP)", "Target-side after-tax income (GAAP)"],
    "adjusted_incremental_pre_tax_income": ["Target-side pre-tax income (Adjusted)"],
    "adjusted_incremental_after_tax_income": ["Target-side after-tax (Adjusted)"],
    "gaap_pro_forma_net_income": ["Combined net income (GAAP)", "GAAP pro forma net income"],
    "adjusted_pro_forma_net_income": ["Combined net income (Adjusted)", "Adjusted pro forma net income"],
    "gaap_pro_forma_eps": ["Combined GAAP EPS", "GAAP pro forma EPS"],
    "adjusted_pro_forma_eps": ["Combined adjusted EPS", "Adjusted pro forma EPS"],
}

_TASK_055_SENSITIVITY_CASE_ALIASES = {
    "synergy_realization_50_percent": (
        "50% synergy realization",
        "50 percent synergy realization",
        "half synergy realization",
    ),
    "debt_rate_up_200_bps": (
        "+200 bps debt cost",
        "200 bps debt cost",
        "debt cost up 200 bps",
        "debt rate up 200 bps",
    ),
    "combined_downside": (
        "combined downside",
        "combined stress",
        "combined 50% synergy + 200 bps",
        "combined 50 percent synergy and 200 bps",
    ),
}

_TASK_055_SENSITIVITY_METRIC_ALIASES = {
    "realized_synergy": ("realized synergy", "cost synergy", "synergy"),
    "gaap_incremental_pre_tax_income": (
        "gaap incremental pre tax income",
        "gaap incremental pretax income",
        "gaap pretax income",
    ),
    "adjusted_incremental_pre_tax_income": (
        "adjusted incremental pre tax income",
        "adjusted incremental pretax income",
        "adjusted pretax income",
        "adj pretax income",
    ),
    "gaap_pro_forma_eps": ("gaap pro forma eps", "gaap eps"),
    "adjusted_pro_forma_eps": (
        "adjusted pro forma eps",
        "adjusted eps",
        "adj eps",
    ),
    "gaap_eps_accretion": (
        "gaap eps accretion",
        "gaap accretion",
        "gaap acc",
        "gaap delta",
        "gaap change",
    ),
    "adjusted_eps_accretion": (
        "adjusted eps accretion",
        "adjusted accretion",
        "adjusted acc",
        "adjusted delta",
        "adjusted change",
        "adj acc",
        "adj delta",
    ),
    "incremental_financing_cost": (
        "incremental financing cost",
        "financing cost",
    ),
    "committee_release_status": (
        "committee release status",
        "release status",
        "committee status",
        "release",
        "status",
    ),
}

_TASK_061_LABEL_ALIASES = {
    "ending_temporary_difference_dta": ["temporary_difference_dta"],
    "ending_state_credit_dta": ["state_credit_dta"],
    "ending_valuation_allowance": ["valuation_allowance"],
    "ending_net_dta": ["net_dta"],
    "ending_net_deferred_tax_liability": ["net_deferred_tax_liability"],
}

_TASK_072_LABEL_ALIASES = {
    "q2_revenue": ["Q2 revenue", "Board Performance Q2 revenue"],
    "q2_approved_plan_revenue": ["Q2 revenue plan", "Approved Q2 plan revenue"],
    "q2_revenue_variance_to_plan": [
        "Q2 revenue variance to plan", "Revenue variance vs plan",
        "Revenue variance to plan", "Q2 revenue variance",
    ],
    "q2_organic_growth": ["Organic growth"],
    "q2_adjusted_ebitda": ["Adjusted EBITDA"],
    "q2_approved_plan_adjusted_ebitda": [
        "Q2 adjusted EBITDA plan", "Approved Q2 plan adjusted EBITDA",
        "Adjusted EBITDA plan",
    ],
    "q2_adjusted_ebitda_variance_to_plan": [
        "Q2 adjusted EBITDA variance to plan", "Adjusted EBITDA variance vs plan",
        "EBITDA variance to plan", "Adjusted EBITDA variance",
    ],
    "construction_gross_profit_impact": ["Construction gross profit impact", "Construction gross profit downside"],
    "service_labor_productivity_impact": ["Service labor productivity impact", "Service labor downside"],
    "controls_mix_impact": ["Controls mix impact", "Controls mix downside"],
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
    "ltm_adjusted_ebitda": ["LTM adjusted EBITDA", "Approved LTM adjusted EBITDA", "Lender adjusted EBITDA"],
    "funded_debt": ["Posted funded debt", "Lender funded debt"],
    "lender_leverage": ["Lender leverage", "Covenant leverage", "Gross leverage"],
    "fixed_charge_coverage": ["Fixed-charge coverage", "Fixed charge coverage", "FCCR", "Lender FCCR"],
    "revenue_guidance_low": ["Revenue guidance low", "Published revenue guidance low"],
    "revenue_guidance_high": ["Revenue guidance high", "Published revenue guidance high"],
    "ebitda_guidance_low": ["EBITDA guidance low", "Published EBITDA guidance low"],
    "ebitda_guidance_high": ["EBITDA guidance high", "Published EBITDA guidance high"],
    "largest_downside_driver": ["Largest downside driver", "Principal downside driver"],
    "combined_stress_ebitda": ["Combined stress EBITDA", "Signed downside stress EBITDA"],
    "ebitda_shortfall_to_guidance_low": ["EBITDA shortfall to guidance low", "Shortfall to published EBITDA low end"],
    "additional_revenue_decline_to_update_trigger": ["Additional revenue decline to update trigger", "Headroom to formal update trigger"],
    "guidance_update_required": ["Guidance update required", "Revise published guidance"],
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

for _branch in ("construction", "service", "controls"):
    _display = _branch.title()
    _TASK_100_LABEL_ALIASES.setdefault(f"{_branch}_forecast_revenue", []).extend([f"{_display} revenue", f"{_display} Rev $M"])
    _TASK_100_LABEL_ALIASES.setdefault(f"{_branch}_forecast_ebitda", []).extend([f"{_display} EBITDA"])
    _TASK_100_LABEL_ALIASES.setdefault(f"{_branch}_ebitda_variance", []).extend([f"{_display} EBITDA variance", f"{_display} Δ vs comp"])
for _case in ("base", "downside", "severe"):
    _display = _case.title()
    for _suffix, _label in (
        ("ending_cash", "Ending cash (pre-fin)"),
        ("revolver_balance", "Revolver balance"),
        ("remaining_revolver_headroom", "Remaining headroom"),
        ("covenant_ebitda", "Covenant EBITDA"),
        ("funded_debt", "Funded debt"),
        ("leverage", "Gross leverage"),
    ):
        _TASK_100_LABEL_ALIASES.setdefault(f"{_case}_{_suffix}", []).append(f"{_display} {_label}")
    _TASK_100_LABEL_ALIASES.setdefault(f"{_case}_ending_cash_before_financing", []).extend([
        f"{_display} Ending cash (pre-fin)",
        f"{_display} cash before financing",
        f"{_display} cash b/f fin",
    ])
    _TASK_100_LABEL_ALIASES.setdefault(f"{_case}_revolver_balance", []).append(f"{_display} revolver")
    _TASK_100_LABEL_ALIASES.setdefault(f"{_case}_remaining_revolver_headroom", []).append(f"{_display} headroom")
    _TASK_100_LABEL_ALIASES.setdefault(f"{_case}_covenant_ebitda", []).append(f"{_display} cov EBITDA")
    _TASK_100_LABEL_ALIASES.setdefault(f"{_case}_leverage", []).append(f"{_display} leverage")
    _TASK_100_LABEL_ALIASES.setdefault(f"{_case}_minimum_cash_headroom", []).extend([f"{_display} cash guardrail", f"{_display} minimum cash headroom"])
    _TASK_100_LABEL_ALIASES.setdefault(f"{_case}_cash_breach", []).extend([
        f"{_display} cash guardrail",
        f"{_display} cash status",
        f"{_display} floor test",
    ])
    _TASK_100_LABEL_ALIASES.setdefault(f"{_case}_leverage_breach", []).extend([
        f"{_display} covenant guardrail",
        f"{_display} leverage status",
        f"{_display} 3.00x cov",
    ])

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
        "review notes",
    ],
    "contractor accounting mcp": [
        "contractor accounting mcp", "vista erp", "posted vista",
        "company accounting records",
    ],
    "executive summary": [
        "executive summary", "executive performance summary", "performance summary",
    ],
    "cash roll forward": [
        "cash roll forward", "cash roll-forward", "cash liquidity scenario",
        "ending cash before financing", "cash & liquidity",
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
    for alternative in spec.get("numeric_alternatives", []):
        if _numeric_matches(actual, float(alternative), abs_tol):
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


def _structured_value_matches(actual: Any, expected: Any) -> bool:
    """Match nested disclosed JSON values without stringifying containers."""
    if isinstance(expected, dict):
        if not isinstance(actual, dict):
            return False
        return all(
            (found := _semantic_find(actual, str(key)))[0]
            and _structured_value_matches(found[1], value)
            for key, value in expected.items()
        )
    if isinstance(expected, list):
        return (
            isinstance(actual, list)
            and len(actual) == len(expected)
            and all(
                _structured_value_matches(actual_value, expected_value)
                for actual_value, expected_value in zip(actual, expected, strict=True)
            )
        )
    if isinstance(expected, bool):
        return _boolean_matches(actual, expected)
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        return _numeric_matches(actual, float(expected), 0.02)
    if expected is None:
        return actual is None or _normalize(actual) in {"none", "null", "n a"}
    return semantic_value_matches(actual, str(expected))


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


def _structured_row(
    mapping: dict[str, Any], *, key: str, row_key: str, row_value: str
) -> dict[str, Any] | None:
    rows = _semantic_get(mapping, key)
    if not isinstance(rows, list):
        return None
    for item in rows:
        if not isinstance(item, dict):
            continue
        actual_row_value = _semantic_get(item, row_key)
        if semantic_value_matches(actual_row_value, row_value):
            return item
    return None


def _structured_field(row: dict[str, Any] | None, spec: dict[str, Any]) -> Any:
    """Read a disclosed structured field while honoring narrow rubric aliases."""
    if row is None:
        return None
    for field in (str(spec["field"]), *[str(value) for value in spec.get("field_aliases", [])]):
        present, value = _semantic_find(row, field)
        if present:
            return value
    return None


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


def _console_criterion(
    mapping: dict[str, Any], spec: dict[str, Any], *, task_id: str | None = None
) -> Criterion:
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
            text_match = (
                _sample_directional_semantic_value_matches(actual, str(expected))
                if task_id == "task_073"
                else semantic_value_matches(actual, str(expected))
            )
            if not present or not ((expected_none and actual_none) or typed_period or numeric_alternative or text_match):
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
            _structured_value_matches(item, expected)
            if isinstance(expected, (dict, list))
            else semantic_value_matches(item, str(expected))
            for item in flattened
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
        if any(isinstance(value, (dict, list)) for value in expected_values):
            if spec.get("unordered"):
                remaining = list(actual) if isinstance(actual, list) else []
                met = len(remaining) == len(expected_values)
                for expected_value in expected_values:
                    matched_index = next(
                        (
                            index
                            for index, actual_value in enumerate(remaining)
                            if _structured_value_matches(actual_value, expected_value)
                        ),
                        None,
                    )
                    if matched_index is None:
                        met = False
                        break
                    remaining.pop(matched_index)
                met = met and not remaining
                expectation = "unordered structured concepts"
            else:
                met = (
                    isinstance(actual, list)
                    and len(actual) == len(expected_values)
                    and all(
                        _structured_value_matches(actual_value, expected_value)
                        for actual_value, expected_value in zip(
                            actual, expected_values, strict=True
                        )
                    )
                )
                expectation = "ordered structured concepts"
        elif spec.get("unordered"):
            met = unordered_semantic_list_matches(actual, expected_values)
            expectation = "unordered concepts"
        else:
            met = (isinstance(actual, list) and not actual) if not expected_values else ordered_semantic_list_matches(actual, expected_values)
            expectation = "ordered concepts"
        evidence = f"actual={actual!r}; expected {expectation}={expected_values!r}"
    elif kind == "structured_row_order":
        actual = _semantic_get(mapping, spec["key"])
        actual_rows = []
        if isinstance(actual, list):
            actual_rows = [
                _semantic_get(item, spec["row_key"])
                for item in actual
                if isinstance(item, dict)
            ]
        expected_rows = list(spec["expected_rows"])
        met = (
            len(actual_rows) == len(expected_rows)
            and all(
                semantic_value_matches(actual_value, expected_value)
                for actual_value, expected_value in zip(actual_rows, expected_rows)
            )
        )
        evidence = f"actual_rows={actual_rows!r}; expected_rows={expected_rows!r}"
    elif kind == "structured_row_numeric":
        row = _structured_row(
            mapping,
            key=str(spec["key"]),
            row_key=str(spec["row_key"]),
            row_value=str(spec["row_value"]),
        )
        actual = _structured_field(row, spec)
        expected = float(spec["expected_value"])
        met = row is not None and _numeric_matches_spec(actual, expected, spec)
        evidence = (
            f"row={spec['row_value']!r}; field={spec['field']!r}; "
            f"actual={actual!r}; expected={expected!r}"
        )
    elif kind == "structured_row_boolean":
        row = _structured_row(
            mapping,
            key=str(spec["key"]),
            row_key=str(spec["row_key"]),
            row_value=str(spec["row_value"]),
        )
        actual = _structured_field(row, spec)
        expected = bool(spec["expected_value"])
        if isinstance(actual, bool):
            normalized = actual
        elif isinstance(actual, str):
            normalized = {"true": True, "false": False}.get(actual.strip().casefold())
        else:
            normalized = None
        met = row is not None and normalized is expected
        evidence = (
            f"row={spec['row_value']!r}; field={spec['field']!r}; "
            f"actual={actual!r}; expected={expected!r}"
        )
    elif kind == "structured_row_list":
        row = _structured_row(
            mapping,
            key=str(spec["key"]),
            row_key=str(spec["row_key"]),
            row_value=str(spec["row_value"]),
        )
        actual = _structured_field(row, spec)
        expected = list(spec["expected_value"])
        met = row is not None and unordered_semantic_list_matches(actual, expected)
        evidence = (
            f"row={spec['row_value']!r}; field={spec['field']!r}; "
            f"actual={actual!r}; expected unordered values={expected!r}"
        )
    elif kind == "structured_row_string":
        row = _structured_row(
            mapping,
            key=str(spec["key"]),
            row_key=str(spec["row_key"]),
            row_value=str(spec["row_value"]),
        )
        actual = _structured_field(row, spec)
        expected = str(spec["expected_value"])
        met = row is not None and semantic_value_matches(actual, expected)
        evidence = (
            f"row={spec['row_value']!r}; field={spec['field']!r}; "
            f"actual={actual!r}; expected={expected!r}"
        )
    elif kind == "structured_row_nullable_numeric":
        row = _structured_row(
            mapping,
            key=str(spec["key"]),
            row_key=str(spec["row_key"]),
            row_value=str(spec["row_value"]),
        )
        actual = None if row is None else _semantic_get(row, str(spec["field"]))
        met = row is not None and actual is None
        evidence = (
            f"row={spec['row_value']!r}; field={spec['field']!r}; "
            f"actual={actual!r}; expected=None"
        )
    elif kind == "structured_row_order_composite":
        actual = _semantic_get(mapping, spec["key"])
        row_keys = [str(value) for value in spec["row_keys"]]
        actual_rows = []
        if isinstance(actual, list):
            actual_rows = [
                [_semantic_get(item, row_key) for row_key in row_keys]
                for item in actual
                if isinstance(item, dict)
            ]
        expected_rows = [list(row) for row in spec["expected_rows"]]
        met = (
            len(actual_rows) == len(expected_rows)
            and all(
                len(actual_row) == len(expected_row)
                and all(
                    semantic_value_matches(actual_value, expected_value)
                    for actual_value, expected_value in zip(actual_row, expected_row)
                )
                for actual_row, expected_row in zip(actual_rows, expected_rows)
            )
        )
        evidence = f"actual_rows={actual_rows!r}; expected_rows={expected_rows!r}"
    elif kind == "structured_row_numeric_composite":
        actual_rows = _semantic_get(mapping, spec["key"])
        row = None
        if isinstance(actual_rows, list):
            for candidate in actual_rows:
                if not isinstance(candidate, dict):
                    continue
                if all(
                    semantic_value_matches(
                        _semantic_get(candidate, str(row_key)), row_value
                    )
                    for row_key, row_value in zip(spec["row_keys"], spec["row_values"])
                ):
                    row = candidate
                    break
        actual = None if row is None else _semantic_get(row, str(spec["field"]))
        expected = float(spec["expected_value"])
        met = row is not None and _numeric_matches_spec(actual, expected, spec)
        evidence = (
            f"row={spec['row_values']!r}; field={spec['field']!r}; "
            f"actual={actual!r}; expected={expected!r}"
        )
    else:
        raise ValueError(f"Unsupported console criterion: {kind}")
    return Criterion(spec["id"], spec["description"], met, evidence)


def _normalize_console_mapping(task_id: str, mapping: dict[str, Any]) -> dict[str, Any]:
    """Accept financially equivalent field names without weakening the math.

    Console tasks ask for professional schedules, not a private serialization
    dialect.  Normalize only exact, source-equivalent aliases or identities;
    every value still has to match the deterministic gold calculation.
    """

    normalized = dict(mapping)

    def copied_rows(key: str) -> list[dict[str, Any]]:
        raw = mapping.get(key)
        if not isinstance(raw, list):
            return []
        rows = [dict(row) for row in raw if isinstance(row, dict)]
        normalized[key] = rows
        return rows

    def canonical_label(value: Any, rules: dict[str, tuple[str, ...]]) -> str | None:
        text = _normalize(value)
        candidates = [
            (target, alias)
            for target, aliases in rules.items()
            for alias in (target, *aliases)
        ]
        # Prefer the most specific ordinary-language label.  This prevents a
        # long label such as "closing leverage (pro forma debt / EBITDA)" from
        # being captured by the shorter "pro forma funded debt" identity.
        candidates.sort(key=lambda pair: len(_normalize(pair[1])), reverse=True)
        for target, alias in candidates:
            if semantic_value_matches(value, alias) or _normalize(alias) in text:
                return target
        return None

    if task_id == "task_032":
        for row in copied_rows("monthly_backtest"):
            if _semantic_get(row, "included_in_mape") is None:
                excluded = next(
                    (
                        value
                        for alias in (
                            "acquisition_month",
                            "acquisition_month_excluded_from_mape",
                            "excluded_from_mape",
                        )
                        if (value := _semantic_get(row, alias)) is not None
                    ),
                    None,
                )
                if isinstance(excluded, (bool, int, float)):
                    row["included_in_mape"] = 0.0 if bool(excluded) else 1.0
        for rank, row in enumerate(copied_rows("business_unit_bias_analysis"), 1):
            signed = _semantic_get(row, "signed_bias")
            if _semantic_get(row, "absolute_bias") is None and isinstance(signed, (int, float)):
                row["absolute_bias"] = abs(float(signed))
            if _semantic_get(row, "absolute_bias_rank") is None:
                row["absolute_bias_rank"] = float(rank)

    if task_id == "task_033":
        for row in copied_rows("branch_analysis"):
            aliases = {
                "invoice_cost_allocation": ("invoice_pool_allocation",),
                "average_invested_capital": ("invested_capital",),
            }
            for target, candidates in aliases.items():
                if _semantic_get(row, target) is None:
                    for candidate in candidates:
                        value = _semantic_get(row, candidate)
                        if value is not None:
                            row[target] = value
                            break
            if _semantic_get(row, "total_corporate_allocation") is None:
                invoice = _semantic_get(row, "invoice_cost_allocation")
                people = _semantic_get(row, "people_systems_allocation")
                if isinstance(invoice, (int, float)) and isinstance(people, (int, float)):
                    row["total_corporate_allocation"] = float(invoice) + float(people)

    if task_id == "task_036":
        for row in copied_rows("branch_working_capital_analysis"):
            aliases = {
                "accounts_payable": ("ap",),
                "ordinary_receivable_days": ("dso_days", "dso"),
                "inventory_days": ("dio_days", "dio"),
                "payable_days": ("dpo_days", "dpo"),
            }
            for target, candidates in aliases.items():
                if _semantic_get(row, target) is None:
                    for candidate in candidates:
                        value = _semantic_get(row, candidate)
                        if value is not None:
                            row[target] = value
                            break
            if _semantic_get(row, "total_ar") is None:
                ordinary = _semantic_get(row, "ordinary_ar")
                retainage = _semantic_get(row, "retainage")
                if isinstance(ordinary, (int, float)) and isinstance(retainage, (int, float)):
                    row["total_ar"] = float(ordinary) + float(retainage)

    if task_id == "task_038":
        shortfall = _semantic_get(normalized, "irr_shortfall_to_hurdle")
        if isinstance(shortfall, (int, float)):
            normalized["irr_shortfall_to_hurdle"] = abs(float(shortfall))

    if task_id == "task_070":
        # Analysts often show a reconciliation as a tie-out object or boolean
        # rather than the scalar zero difference used by the gold record.  The
        # object still has to prove both sides are equal; a bare false or an
        # unequal pair receives no normalization.
        for key in (
            "segment_tam_reconciliation_check",
            "segment_sam_reconciliation_check",
            "segment_obtainable_reconciliation_check",
        ):
            value = _semantic_get(normalized, key)
            if value is True:
                normalized[key] = 0.0
            elif isinstance(value, dict):
                tied = _semantic_get(value, "ties")
                segment_sum = _semantic_get(value, "segment_sum")
                consolidated = _semantic_get(value, "consolidated")
                if (
                    _boolean_matches(tied, True)
                    and _number(segment_sum) is not None
                    and _number(consolidated) is not None
                    and _close(
                        segment_sum,
                        float(_number(consolidated)),
                        abs_tol=0.02,
                        rel_tol=0.0,
                    )
                ):
                    normalized[key] = 0.0

    if task_id == "task_071":
        rows = copied_rows("valuation_adjustment_analysis")
        if rows:
            # A clearly labeled peer-median baseline is useful presentation,
            # but is not one of the five signed adjustments being graded.
            rows[:] = [
                row
                for row in rows
                if _normalize(
                    _semantic_get(row, "adjustment")
                    or _semantic_get(row, "step")
                    or ""
                )
                not in {
                    "peer median start",
                    "peer median baseline",
                    "peer median",
                }
            ]
            for row in rows:
                if _semantic_get(row, "adjustment") is None:
                    step = _semantic_get(row, "step")
                    if step is not None:
                        row["adjustment"] = step
                if _semantic_get(row, "supported_multiple_after_adjustment") is None:
                    multiple = _semantic_get(row, "multiple_after_step")
                    if multiple is not None:
                        row["supported_multiple_after_adjustment"] = multiple

    if task_id == "task_078":
        for row in copied_rows("delivery_candidate_analysis"):
            aliases = {
                "required_labor_hours": ("labor_hours",),
                "mandatory_signed_award": ("mandatory",),
                "selection_status": ("selection_decision",),
            }
            for target, candidates in aliases.items():
                if _semantic_get(row, target) is not None:
                    continue
                for candidate in candidates:
                    value = _semantic_get(row, candidate)
                    if value is not None:
                        row[target] = value
                        break
            if _semantic_get(row, "gross_margin") is None:
                revenue = _number(_semantic_get(row, "risk_adjusted_revenue"))
                gross_profit = _number(
                    _semantic_get(row, "risk_adjusted_gross_profit")
                )
                if revenue not in {None, 0.0} and gross_profit is not None:
                    row["gross_margin"] = gross_profit / revenue
            if _semantic_get(row, "selection_status") is None:
                selected = _semantic_get(row, "selected")
                if isinstance(selected, bool):
                    row["selection_status"] = (
                        "Selected" if selected else "Excluded"
                    )
            else:
                # Explanatory decisions such as "Selected - mandatory signed
                # award" carry the same deterministic status prefix.
                decision = _normalize(_semantic_get(row, "selection_status"))
                if decision.startswith("selected"):
                    row["selection_status"] = "Selected"
                elif decision.startswith("excluded"):
                    row["selection_status"] = "Excluded"

    if task_id == "task_079":
        for row in copied_rows("customer_economic_analysis"):
            aliases = {
                "ltm_economic_profit": ("economic_profit",),
                "recommended_action": ("decision",),
                "expected_credit_loss": ("expected_credit_loss_raw",),
            }
            for target, candidates in aliases.items():
                if any(
                    _normalize(key) == _normalize(target)
                    for key in row
                ):
                    continue
                for candidate in candidates:
                    exact_key = next(
                        (
                            key
                            for key in row
                            if _normalize(key) == _normalize(candidate)
                        ),
                        None,
                    )
                    if exact_key is not None:
                        row[target] = row[exact_key]
                        break
            if _semantic_get(row, "below_economic_margin_floor") is None:
                margin = _number(_semantic_get(row, "economic_margin"))
                if margin is not None:
                    row["below_economic_margin_floor"] = margin < 0.045

    if task_id == "task_083":
        # The prompt describes the outage rows in ordinary portfolio language.
        # Normalize only exact submitted aliases; every amount and selected
        # action still has to tie to the deterministic refinancing gold.
        outage_aliases = {
            "selected_actions": ("replacement_selection",),
            "selected_gross_financial_effect": ("gross_financial_effect",),
            "selected_probability_weighted_effect": (
                "probability_weighted_effect",
            ),
            "selected_cash_cost": ("cash_cost",),
            "selected_expected_net_effect": ("expected_net_effect",),
            "expected_net_value_loss": (
                "value_loss_vs_base",
                "value_loss_vs_base_optimum",
            ),
        }
        for row in copied_rows("refinancing_protection_outage_analysis"):
            for target, aliases in outage_aliases.items():
                if _semantic_get(row, target) is not None:
                    continue
                for alias in aliases:
                    value = _semantic_get(row, alias)
                    if value is not None:
                        row[target] = value
                        break

    if task_id == "task_096":
        # The prompt asks for professional scenario/action, close-schedule,
        # and release-bridge tables without imposing a private nested-field
        # dialect. Normalize only values the analyst actually disclosed or
        # exact identities proved elsewhere in the same submitted answer.
        scenario_field_aliases = {
            "scenario_cash_cost": ("cash_cost",),
            "scenario_gross_financial_effect": ("gross_financial_effect",),
            "scenario_realization_probability": ("realization_probability",),
            "scenario_probability_weighted_effect": (
                "probability_weighted_effect",
            ),
            "scenario_capacity_units": ("capacity_units",),
        }
        scenario_rows = copied_rows(
            "structure_protection_scenario_action_analysis"
        )
        for row in scenario_rows:
            for target, aliases in scenario_field_aliases.items():
                if _semantic_get(row, target) is not None:
                    continue
                for alias in aliases:
                    value = _semantic_get(row, alias)
                    if value is not None:
                        row[target] = value
                        break

        summary_by_scenario: dict[str, dict[str, Any]] = {}
        for row in copied_rows("structure_protection_scenario_summary"):
            scenario = _normalize(_semantic_get(row, "scenario"))
            if not scenario:
                continue
            if _semantic_get(row, "value_loss_vs_management") is None:
                value = _semantic_get(row, "expected_net_value_loss")
                if value is not None:
                    row["value_loss_vs_management"] = value
            selected_cash = _number(_semantic_get(row, "selected_cash_cost"))
            cash_headroom = _number(_semantic_get(row, "cash_budget_headroom"))
            if (
                not any(
                    _normalize(key) == "cash budget"
                    for key in row
                )
                and selected_cash is not None
                and cash_headroom is not None
            ):
                row["cash_budget"] = selected_cash + cash_headroom
            selected_capacity = sum(
                float(capacity)
                for action_row in scenario_rows
                if _normalize(_semantic_get(action_row, "scenario")) == scenario
                and _boolean_matches(_semantic_get(action_row, "selected"), True)
                and (
                    capacity := _number(
                        _semantic_get(action_row, "scenario_capacity_units")
                    )
                ) is not None
            )
            capacity_headroom = _number(
                _semantic_get(row, "capacity_headroom")
            )
            if (
                _semantic_get(row, "capacity_limit") is None
                and capacity_headroom is not None
            ):
                row["capacity_limit"] = selected_capacity + capacity_headroom
            summary_by_scenario[scenario] = row

        bridge_aliases = {
            "base_net_asset_structure_benefit": ("base_net_asset_benefit",),
            "selected_expected_net_effect": ("expected_net_protection",),
            "selected_probability_weighted_effect": (
                "probability_weighted_effect",
            ),
            "protected_after_tax_transaction_value": (
                "protected_after_tax_value",
            ),
            "selected_cash_cost": ("cash_cost",),
            "scenario_threshold_ready": ("threshold_status",),
            "structure_release_ready": ("release_status",),
        }
        for row in copied_rows("transaction_structure_release_bridge"):
            for target, aliases in bridge_aliases.items():
                if _semantic_get(row, target) is not None:
                    continue
                for alias in aliases:
                    value = _semantic_get(row, alias)
                    if value is not None:
                        row[target] = value
                        break
            scenario = _normalize(_semantic_get(row, "scenario"))
            summary = summary_by_scenario.get(scenario)
            if summary is not None:
                for target in (
                    "selected_probability_weighted_effect",
                    "cash_budget_headroom",
                ):
                    if _semantic_get(row, target) is None:
                        value = _semantic_get(summary, target)
                        if value is not None:
                            row[target] = value
            for field in ("scenario_threshold_ready", "structure_release_ready"):
                value = _semantic_get(row, field)
                text = _normalize(value)
                if text in {"above threshold", "threshold met", "release ready"}:
                    row[field] = True
                elif text in {
                    "below threshold",
                    "threshold not met",
                    "not release ready",
                }:
                    row[field] = False

        close_schedule_aliases = {
            "step_up_basis": ("step_up",),
            "recovery_life": ("recovery_life_years", "life_years"),
            "year_1_mid_year_straight_line_deduction": (
                "year_1_midyear_straight_line_deduction",
                "year_1_midyear_sl_deduction",
            ),
            "year_1_total_deduction": ("total_deduction",),
            "year_1_tax_shield": ("tax_shield",),
            "remaining_basis_after_year_1": ("remaining_basis",),
            "pv_tax_shield": ("pv_shield",),
        }
        for row in copied_rows("tax_basis_close_schedule"):
            for target, aliases in close_schedule_aliases.items():
                if _semantic_get(row, target) is not None:
                    continue
                for alias in aliases:
                    value = _semantic_get(row, alias)
                    if value is not None:
                        row[target] = value
                        break

        sensitivity_case_aliases = {
            "Seller required gross-up": (
                "required gross-up documented",
                "required gross-up",
            ),
            "Buyer policy cap": (
                "buyer incremental-consideration cap",
                "buyer consideration-cap ceiling",
            ),
            "Economic gross-up ceiling": (
                "maximum economic gross-up",
                "maximum economic ceiling",
            ),
        }
        for row in copied_rows("seller_gross_up_sensitivity"):
            label = canonical_label(
                _semantic_get(row, "case"), sensitivity_case_aliases
            )
            if label is not None:
                row["case"] = label

        # The negotiation matrix is a finance schedule, not a private JSON
        # dialect.  Normalize ordinary labels that the two exact v21 pilots
        # used, while retaining every submitted amount and gate for the
        # deterministic comparison.  Derived headroom is accepted only when
        # it is an exact identity of values disclosed elsewhere in the same
        # answer.
        matrix_case_aliases = {
            "No seller gross-up": ("zero gross-up", "no gross up"),
            "Buyer policy cap": (
                "buyer cap-implied gross-up",
                "buyer incremental-consideration cap",
                "buyer consideration cap",
            ),
            "Seller required gross-up": (
                "seller-required gross-up",
                "required gross-up",
            ),
            "Economic gross-up ceiling": (
                "economic-ceiling gross-up",
                "economic ceiling",
            ),
        }
        matrix_field_aliases = {
            "tax_close_scenario": ("scenario",),
            "net_asset_structure_benefit_before_protection": (
                "pre_protection_value",
                "pre_protection_net_asset_benefit",
            ),
            "selected_expected_net_protection": (
                "expected_net_protection",
            ),
            "protected_after_tax_transaction_value": (
                "protected_value",
                "protected_after_tax_value",
            ),
            "selected_cash_cost": ("cash_cost",),
            "seller_consideration_within_cap": (
                "incremental_consideration_gate",
            ),
            "protected_positive_npv_gate": ("positive_npv_gate",),
            "scenario_threshold_ready": ("threshold_status",),
            "structure_release_ready": ("final_status", "release_status"),
        }
        maximum_buyer_consideration = _number(
            _semantic_get(normalized, "maximum_buyer_incremental_consideration")
        )
        economic_ceiling = _number(
            _semantic_get(
                normalized,
                "tax_basis_cured_maximum_economic_seller_gross_up",
            )
        )
        for row in copied_rows("tax_structure_negotiation_scenario_matrix"):
            label = canonical_label(
                _semantic_get(row, "gross_up_case"), matrix_case_aliases
            )
            if label is not None:
                row["gross_up_case"] = label
            for target, aliases in matrix_field_aliases.items():
                if _semantic_get(row, target) is not None:
                    continue
                for alias in aliases:
                    value = _semantic_get(row, alias)
                    if value is not None:
                        row[target] = value
                        break
            scenario = _normalize(_semantic_get(row, "tax_close_scenario"))
            summary = summary_by_scenario.get(scenario)
            if summary is not None and _semantic_get(
                row, "cash_budget_headroom"
            ) is None:
                value = _semantic_get(summary, "cash_budget_headroom")
                if value is not None:
                    row["cash_budget_headroom"] = value
            seller_gross_up = _number(_semantic_get(row, "seller_gross_up"))
            if (
                _semantic_get(row, "buyer_cap_headroom") is None
                and maximum_buyer_consideration is not None
                and seller_gross_up is not None
            ):
                row["buyer_cap_headroom"] = (
                    maximum_buyer_consideration - seller_gross_up
                )
            if (
                _semantic_get(row, "economic_headroom") is None
                and economic_ceiling is not None
                and seller_gross_up is not None
            ):
                row["economic_headroom"] = economic_ceiling - seller_gross_up
            for field in (
                "seller_consideration_within_cap",
                "protected_positive_npv_gate",
                "scenario_threshold_ready",
                "structure_release_ready",
            ):
                value = _semantic_get(row, field)
                text = _normalize(value)
                if text in {
                    "release",
                    "ready",
                    "true",
                    "yes",
                    "pass",
                }:
                    row[field] = True
                elif text in {
                    "do not release",
                    "not ready",
                    "false",
                    "no",
                    "fail",
                    "hold",
                }:
                    row[field] = False
            if _semantic_get(row, "case_scenario") is None:
                case = _semantic_get(row, "gross_up_case")
                scenario_value = _semantic_get(row, "tax_close_scenario")
                if case is not None and scenario_value is not None:
                    row["case_scenario"] = f"{case} | {scenario_value}"

        for key in (
            "class_level_year_1_tax_shield",
            "class_level_pv_tax_shield",
            "class_level_net_asset_structure_benefit",
        ):
            value = _semantic_get(normalized, key)
            amounts: list[float] = []
            if isinstance(value, dict):
                amounts = [
                    float(item)
                    for item in value.values()
                    if isinstance(item, (int, float)) and not isinstance(item, bool)
                ]
            elif isinstance(value, list):
                field = key.removeprefix("class_level_")
                amounts = [
                    float(item_value)
                    for item in value
                    if isinstance(item, dict)
                    and (
                        item_value := _number(_semantic_get(item, field))
                    ) is not None
                ]
            if len(amounts) == 4:
                normalized[key] = sum(amounts)

    if task_id == "task_099":
        # A controller may show the same annual capital-return roll-forward
        # using separate cadence/year labels and annual finance effects. Build
        # only identities proved by submitted values; do not infer missing
        # economics or override an explicit canonical value.
        rows = copied_rows("two_year_cadence_year_analysis")
        running_interest: dict[str, float] = {}
        running_foregone_yield: dict[str, float] = {}
        running_deployed: dict[str, float] = {}
        after_tax_interest_rate: dict[str, float] = {}
        after_tax_cash_yield: dict[str, float] = {}
        baseline_shares: dict[str, float] = {}
        for row in rows:
            cadence = _semantic_get(row, "cadence")
            fiscal_year = _semantic_get(row, "fiscal_year")
            if (
                _semantic_get(row, "cadence_year") is None
                and cadence is not None
                and fiscal_year is not None
            ):
                row["cadence_year"] = f"{cadence} | {fiscal_year}"

            cadence_key = _normalize(cadence)
            annual_interest = _number(
                _semantic_get(row, "after_tax_interest_savings")
            )
            submitted_foregone_yield = _number(
                _semantic_get(row, "after_tax_foregone_cash_yield")
            )
            actual_debt_paydown = _number(
                _semantic_get(row, "actual_debt_paydown")
            )
            total_outlay = _number(_semantic_get(row, "total_cash_outlay"))
            if annual_interest is not None:
                prior_interest = running_interest.get(cadence_key, 0.0)
                rate = after_tax_interest_rate.get(cadence_key)
                if (
                    rate is None
                    and actual_debt_paydown not in {None, 0.0}
                ):
                    rate = annual_interest / actual_debt_paydown
                    after_tax_interest_rate[cadence_key] = rate
                expected_annual_interest = (
                    actual_debt_paydown * rate
                    if actual_debt_paydown is not None and rate is not None
                    else None
                )
                submitted_is_cumulative = (
                    prior_interest != 0.0
                    and expected_annual_interest is not None
                    and math.isclose(
                        annual_interest,
                        prior_interest + expected_annual_interest,
                        rel_tol=1e-6,
                        abs_tol=0.05,
                    )
                )
                running_interest[cadence_key] = (
                    annual_interest
                    if submitted_is_cumulative
                    else prior_interest + annual_interest
                )
                if _semantic_get(row, "cumulative_interest_savings") is None:
                    row["cumulative_interest_savings"] = running_interest[cadence_key]
            if submitted_foregone_yield is not None:
                prior_foregone = running_foregone_yield.get(cadence_key, 0.0)
                yield_rate = after_tax_cash_yield.get(cadence_key)
                if yield_rate is None and total_outlay not in {None, 0.0}:
                    yield_rate = submitted_foregone_yield / total_outlay
                    after_tax_cash_yield[cadence_key] = yield_rate
                expected_annual_foregone = (
                    total_outlay * yield_rate
                    if total_outlay is not None and yield_rate is not None
                    else None
                )
                submitted_is_cumulative = (
                    prior_foregone != 0.0
                    and expected_annual_foregone is not None
                    and math.isclose(
                        submitted_foregone_yield,
                        prior_foregone + expected_annual_foregone,
                        rel_tol=1e-6,
                        abs_tol=0.05,
                    )
                )
                running_foregone_yield[cadence_key] = (
                    submitted_foregone_yield
                    if submitted_is_cumulative
                    else prior_foregone + submitted_foregone_yield
                )
            if total_outlay is not None:
                running_deployed[cadence_key] = (
                    running_deployed.get(cadence_key, 0.0) + total_outlay
                )
                if _semantic_get(row, "cumulative_deployed_cash") is None:
                    row["cumulative_deployed_cash"] = running_deployed[cadence_key]

            shares_repurchase = _number(_semantic_get(row, "share_repurchase"))
            shares_repurchased = _number(_semantic_get(row, "shares_repurchased"))
            ending_shares = _number(_semantic_get(row, "ending_diluted_shares"))
            volume_capacity = _number(_semantic_get(row, "volume_capacity_shares"))
            ending_cash = _number(_semantic_get(row, "ending_cash"))
            pro_forma_eps = _number(_semantic_get(row, "pro_forma_eps"))
            eps_accretion = _number(_semantic_get(row, "eps_accretion"))
            gross_leverage = _number(_semantic_get(row, "gross_leverage"))

            if (
                _semantic_get(row, "opening_diluted_shares") is None
                and ending_shares is not None
                and shares_repurchased is not None
            ):
                row["opening_diluted_shares"] = ending_shares + shares_repurchased
            opening_shares = _number(
                _semantic_get(row, "opening_diluted_shares")
            )
            if cadence_key not in baseline_shares and opening_shares is not None:
                baseline_shares[cadence_key] = opening_shares
            if (
                _semantic_get(row, "share_price") is None
                and shares_repurchase is not None
                and shares_repurchased not in {None, 0.0}
            ):
                row["share_price"] = shares_repurchase / shares_repurchased
            if (
                _semantic_get(row, "trading_volume") is None
                and volume_capacity is not None
            ):
                row["trading_volume"] = volume_capacity / 0.004
            if (
                _semantic_get(row, "allocation_income_effect") is None
                and cadence_key in running_interest
                and cadence_key in running_foregone_yield
            ):
                row["allocation_income_effect"] = (
                    running_interest[cadence_key]
                    - running_foregone_yield[cadence_key]
                )
            allocation_effect = _number(
                _semantic_get(row, "allocation_income_effect")
            )
            if (
                _semantic_get(row, "pro_forma_net_income") is None
                and pro_forma_eps is not None
                and ending_shares is not None
            ):
                row["pro_forma_net_income"] = round(
                    pro_forma_eps * ending_shares, 2
                )
            pro_forma_income = _number(
                _semantic_get(row, "pro_forma_net_income")
            )
            if (
                _semantic_get(row, "base_net_income") is None
                and pro_forma_income is not None
                and allocation_effect is not None
            ):
                row["base_net_income"] = round(
                    pro_forma_income - allocation_effect, 2
                )
            base_income = _number(_semantic_get(row, "base_net_income"))
            base_share_count = baseline_shares.get(cadence_key)
            if (
                _semantic_get(row, "base_eps") is None
                and base_income is not None
                and base_share_count not in {None, 0.0}
            ):
                row["base_eps"] = base_income / base_share_count
            if (
                _semantic_get(row, "minimum_cash_headroom") is None
                and ending_cash is not None
            ):
                row["minimum_cash_headroom"] = ending_cash - 2_250_000.0
            if (
                _semantic_get(row, "repurchase_volume_headroom") is None
                and volume_capacity is not None
                and shares_repurchased is not None
            ):
                row["repurchase_volume_headroom"] = (
                    volume_capacity - shares_repurchased
                )
            if (
                _semantic_get(row, "eps_accretion_headroom") is None
                and eps_accretion is not None
            ):
                row["eps_accretion_headroom"] = eps_accretion - 0.01
            if (
                _semantic_get(row, "gross_leverage_headroom") is None
                and gross_leverage is not None
            ):
                row["gross_leverage_headroom"] = 1.5 - gross_leverage
            if _semantic_get(row, "year_feasible") is None:
                annual_pass = _semantic_get(row, "annual_guardrails_pass")
                if isinstance(annual_pass, bool):
                    row["year_feasible"] = annual_pass

        for key in (
            "selected_cadence_fy27_eps_check",
            "selected_cadence_fy28_eps_check",
            "selected_cadence_fy28_cash_check",
            "selected_cadence_fy28_debt_check",
        ):
            if _semantic_get(normalized, key) is True:
                normalized[key] = 0.0

    if task_id == "task_050":
        # The quarterly release is a finance schedule, not a private JSON
        # dialect.  Accept ordinary controller-workpaper labels while keeping
        # every source amount and release decision deterministic.  These
        # aliases do not infer a missing value; they only copy a submitted
        # value into the disclosed canonical field before exact grading.
        quarterly_rows = copied_rows("quarterly_release_analysis")
        quarterly_aliases = {
            "downside_operating_cash_generation": ("downside_operating_cash",),
            "funded_capital_expenditure": ("funded_capex",),
            "required_debt_amortization": ("debt_amortization",),
            "requested_distribution": ("requested_this_quarter",),
            "pre_release_cash": ("cumulative_cash_before_release",),
            "cash_limited_release_capacity": ("cash_available_for_release",),
            "annual_fccr_capacity_remaining": ("remaining_annual_fccr_capacity",),
            "approved_distribution": ("approved_release",),
            "deferred_distribution": ("deferred_release",),
            "post_release_cash": ("ending_cash_after_release",),
        }
        canonical_quarterly_fields = {
            "quarter",
            "downside_operating_cash_generation",
            "funded_capital_expenditure",
            "required_debt_amortization",
            "requested_distribution",
            "pre_release_cash",
            "cash_limited_release_capacity",
            "annual_fccr_capacity_remaining",
            "approved_distribution",
            "deferred_distribution",
            "post_release_cash",
            "cash_floor_headroom",
            "release_status",
        }
        for index, row in enumerate(quarterly_rows, start=1):
            for target, candidates in quarterly_aliases.items():
                if _semantic_get(row, target) is not None:
                    continue
                for candidate in candidates:
                    value = _semantic_get(row, candidate)
                    if value is not None:
                        row[target] = value
                        break
            # The task asks for the ordered FY27 quarter schedule.  Analysts
            # commonly label the same four rows Q1..Q4, 2027-Q1..Q4, or by the
            # underlying calendar quarters.  When all four rows are supplied
            # in the requested order, normalize only that identity label.
            quarter = _normalize(_semantic_get(row, "quarter"))
            accepted_identity = {
                f"q{index}",
                f"fy27 q{index}",
                f"2027 q{index}",
                ("2026 q3", "2026 q4", "2027 q1", "2027 q2")[index - 1],
            }
            if len(quarterly_rows) == 4 and quarter in accepted_identity:
                row["quarter"] = f"FY27-Q{index}"
            if _semantic_get(row, "cash_floor_headroom") is None:
                post_cash = _semantic_get(row, "post_release_cash")
                floor = _semantic_get(row, "cash_floor_and_reserve")
                if isinstance(post_cash, (int, float)) and isinstance(floor, (int, float)):
                    row["cash_floor_headroom"] = float(post_cash) - float(floor)
            # Extra diagnostic columns are useful workpaper support, but the
            # deterministic item matcher should compare the requested finance
            # fields rather than reject a correct row for carrying them.
            for field in canonical_quarterly_fields:
                value = _semantic_get(row, field)
                if value is not None:
                    row[field] = value
            for field in tuple(row):
                if field not in canonical_quarterly_fields:
                    row.pop(field)

    if task_id == "task_056":
        # Analysts commonly show the raw probability-weighted value beside a
        # separate underwritten flag.  The committee case excludes any row
        # explicitly marked non-underwritten, regardless of that diagnostic.
        for row in copied_rows("synergy_initiative_analysis"):
            if _semantic_get(row, "underwritten") is False:
                row["probability_weighted_synergy"] = 0.0

        eps_rules = {
            "Buyer standalone net income": (),
            "Buyer diluted shares": (),
            "Buyer standalone EPS": (),
            "New shares": ("New shares issued",),
            "Pro forma diluted shares": (),
            "Required combined net income": (
                "Required pro forma adjusted net income at breakeven",
                "Required pro forma adjusted net income to hold standalone EPS",
                "Required pro forma net income for zero dilution",
                "Pro forma net income for zero dilution",
            ),
            "Target after-tax earnings before synergy": (
                "Less target after-tax earnings before synergy",
            ),
            "Incremental after-tax financing cost": (
                "Add back incremental after-tax financing cost",
            ),
            "Required after-tax synergy": (
                "After-tax synergy required at breakeven",
            ),
            "Tax rate": (),
            "Required pre-tax synergy": (
                "Pre-tax synergy required for zero EPS dilution",
                "Required pre-tax synergy EPS breakeven",
            ),
        }
        eps_rows = copied_rows("eps_breakeven_bridge")
        normalized_eps_by_label: dict[str, dict[str, Any]] = {}
        for row in eps_rows:
            if _normalize(_semantic_get(row, "metric")).startswith(
                "less buyer standalone net income"
            ):
                continue
            label = canonical_label(_semantic_get(row, "metric"), eps_rules)
            if label is None:
                continue
            row["metric"] = label
            if label in {"Buyer diluted shares", "New shares", "Pro forma diluted shares"}:
                if _normalize(_semantic_get(row, "unit")) in {"count", "share", "shares"}:
                    row["unit"] = "shares"
            if label in {
                "Target after-tax earnings before synergy",
                "Incremental after-tax financing cost",
            }:
                value = _semantic_get(row, "value")
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    row["value"] = abs(float(value))
            normalized_eps_by_label[label] = row
        normalized["eps_breakeven_bridge"] = [
            normalized_eps_by_label[label]
            for label in eps_rules
            if label in normalized_eps_by_label
        ]

        irr_rules = {
            "Purchase equity investment": (
                "Purchase equity investment t0 outflow",
                "Equity investment Year 0 outflow",
            ),
            "Annual base after-tax target cash flow": (
                "Base annual after-tax target cash flow Years 1-5",
            ),
            "Year 5 terminal proceeds": (),
            "IRR hurdle": (),
            "PV of base cash flows and terminal proceeds": (),
            "PV shortfall to equity investment": (
                "PV shortfall to be covered by incremental synergy",
                "PV shortfall filled by synergy equity investment less PV base and PV terminal",
                "Required PV from synergy",
            ),
            "Pre-tax synergy PV factor": (),
            "Required constant annual pre-tax synergy": (
                "Constant annual pre-tax synergy required to meet 12% equity IRR",
                "Required constant annual pre-tax synergy IRR hurdle",
                "Required constant annual pre-tax synergy to meet IRR hurdle",
                "IRR hurdle required constant pre-tax synergy",
            ),
        }
        irr_rows = copied_rows("irr_hurdle_bridge")
        by_label: dict[str, dict[str, Any]] = {}
        base_pv = terminal_pv = annuity = tax_rate = None
        for row in irr_rows:
            raw_label = _normalize(_semantic_get(row, "metric"))
            value = _semantic_get(row, "value")
            if "pv of base" in raw_label and "terminal" not in raw_label:
                base_pv = value
                continue
            if "pv of terminal" in raw_label or (
                "pv" in raw_label
                and "terminal proceeds" in raw_label
                and "base cash flows" not in raw_label
            ):
                terminal_pv = value
                continue
            if "annuity factor" in raw_label:
                annuity = value
                continue
            if "annuity discount factor" in raw_label:
                annuity = value
                continue
            if "discount factor" in raw_label:
                continue
            if "constant annual after tax synergy" in raw_label:
                continue
            if raw_label == "tax rate":
                tax_rate = value
                continue
            label = canonical_label(_semantic_get(row, "metric"), irr_rules)
            if label is not None:
                row["metric"] = label
                if label == "Purchase equity investment" and isinstance(value, (int, float)):
                    row["value"] = abs(float(value))
                by_label[label] = row
        if (
            "PV of base cash flows and terminal proceeds" not in by_label
            and isinstance(base_pv, (int, float))
            and isinstance(terminal_pv, (int, float))
        ):
            by_label["PV of base cash flows and terminal proceeds"] = {
                "metric": "PV of base cash flows and terminal proceeds",
                "value": float(base_pv) + float(terminal_pv),
                "unit": "USD",
            }
        if (
            "Pre-tax synergy PV factor" not in by_label
            and isinstance(annuity, (int, float))
            and isinstance(tax_rate, (int, float))
        ):
            by_label["Pre-tax synergy PV factor"] = {
                "metric": "Pre-tax synergy PV factor",
                "value": float(annuity) * (1.0 - float(tax_rate)),
                "unit": "multiple",
            }
        factor_row = by_label.get("Pre-tax synergy PV factor")
        if factor_row is not None and _normalize(_semantic_get(factor_row, "unit")) in {
            "factor", "multiple", "x"
        }:
            factor_row["unit"] = "multiple"
        normalized["irr_hurdle_bridge"] = [
            by_label[label] for label in irr_rules if label in by_label
        ]

        authority_rules = {
            "Branch consolidation gross synergy": (
                "Branch consolidation gross run-rate synergy",
                "Branch consolidation - Gross run-rate synergy",
            ),
            "Branch consolidation probability": (
                "Branch consolidation realization probability",
                "Branch consolidation - Realization probability",
            ),
            "Dispatch optimization gross synergy": (
                "Dispatch optimization gross run-rate synergy",
                "Dispatch optimization - Gross run-rate synergy",
            ),
            "Dispatch optimization probability": (
                "Dispatch optimization realization probability",
                "Dispatch optimization - Realization probability",
            ),
            "Cross-sell probability": (
                "Cross-sell realization probability",
                "Cross-sell - Realization probability",
            ),
            "Future pricing concept gross synergy": (
                "Future pricing concept gross run-rate synergy",
                "Future pricing concept - Gross run-rate synergy",
            ),
            "Future pricing concept probability": (
                "Future pricing concept realization probability",
                "Future pricing concept - Realization probability",
            ),
            "Incremental after-tax financing cost": (
                "Summit incremental after-tax financing cost",
            ),
            "Tax rate": ("Summit tax rate",),
            "Purchase equity investment": ("Summit purchase equity investment",),
        }
        normalized_authority_by_label: dict[str, dict[str, Any]] = {}
        excluded_authority = {
            "Cross-sell probability",
            "Future pricing concept gross synergy",
            "Future pricing concept probability",
        }
        for row in copied_rows("source_authority_reconciliation"):
            label = canonical_label(_semantic_get(row, "driver"), authority_rules)
            if label is None:
                continue
            row["driver"] = label
            treatment = _normalize(_semantic_get(row, "authority_treatment"))
            current_supported = any(token in treatment for token in (
                "current", "controller", "7 5", "7.5",
            ))
            prior_rejected = any(token in treatment for token in (
                "6 29", "6.29", "prior", "banker",
            )) and any(token in treatment for token in (
                "history", "trace", "superseded", "rejected", "audit",
            ))
            disclosed_bridge = all(
                _semantic_get(row, field) is not None
                for field in ("current_value", "prior_value", "current_less_prior")
            )
            excluded_supported = prior_rejected or label not in excluded_authority or any(
                token in treatment
                for token in ("excluded", "unapproved", "non actionable", "not underwritten")
            )
            if current_supported and excluded_supported and (
                prior_rejected or disclosed_bridge
            ):
                row["authority_treatment"] = (
                    "Current controller classification excludes this item; prior banker value is trace-only"
                    if label in excluded_authority
                    else "Use current controller-tied value; prior banker value is trace-only"
                )
            normalized_authority_by_label[label] = row
        normalized["source_authority_reconciliation"] = [
            normalized_authority_by_label[label]
            for label in authority_rules
            if label in normalized_authority_by_label
        ]

        approved_names = (
            "Branch consolidation", "Dispatch optimization", "Vendor terms",
        )

        def initiative_set(value: Any) -> tuple[str, ...]:
            if isinstance(value, (list, tuple, set)):
                text = " ".join(str(item) for item in value)
            elif isinstance(value, dict):
                text = " ".join(str(item) for item in value.values())
            else:
                text = str(value or "")
            normalized_text = _normalize(text)
            return tuple(
                name for name in approved_names
                if _normalize(name) in normalized_text
            )

        # The presentation label and JSON container are not finance facts.
        # Canonicalize only the three disclosed initiative identities, while
        # leaving every submitted amount and hurdle decision untouched.
        failure_by_name: dict[str, dict[str, Any]] = {}
        for row in copied_rows("initiative_failure_sensitivity"):
            case_text = _normalize(_semantic_get(row, "case"))
            if "all three" in case_text or "committee case" in case_text:
                continue
            matched = next(
                (name for name in approved_names if _normalize(name) in case_text),
                None,
            )
            if matched is not None:
                row["case"] = matched
                failure_by_name[matched] = row
        normalized["initiative_failure_sensitivity"] = [
            failure_by_name[name] for name in approved_names if name in failure_by_name
        ]

        realization_labels = (
            ("Approved base", ("approved base", "100")),
            ("75% realization", ("75",)),
            ("50% realization", ("50",)),
            ("No synergy", ("no synergy", "0")),
        )
        realization_by_label: dict[str, dict[str, Any]] = {}
        for row in copied_rows("realization_sensitivity"):
            case_text = _normalize(_semantic_get(row, "case"))
            for label, tokens in realization_labels:
                if any(token in case_text for token in tokens):
                    row["case"] = label
                    realization_by_label[label] = row
                    break
        normalized["realization_sensitivity"] = [
            realization_by_label[label]
            for label, _ in realization_labels
            if label in realization_by_label
        ]

        lattice_by_mask: dict[int, dict[str, Any]] = {}
        for row in copied_rows("approved_portfolio_lattice"):
            selected = initiative_set(_semantic_get(row, "selected_initiatives"))
            if not selected:
                selected = initiative_set(_semantic_get(row, "case"))
            mask = sum(1 << index for index, name in enumerate(approved_names) if name in selected)
            lost = tuple(name for name in approved_names if name not in selected)
            row["case"] = (
                "Retain " + " + ".join(selected)
                if selected else "Retain no approved initiatives"
            )
            row["selected_initiatives"] = "; ".join(selected) if selected else "None"
            row["lost_initiatives"] = "; ".join(lost) if lost else "None"
            row["retained_initiative_count"] = len(selected)
            lattice_by_mask[mask] = row
        normalized["approved_portfolio_lattice"] = [
            lattice_by_mask[mask] for mask in range(8) if mask in lattice_by_mask
        ]

        joint_by_mask: dict[int, dict[str, Any]] = {}
        for row in copied_rows("joint_realization_risk_analysis"):
            realized = initiative_set(_semantic_get(row, "realized_initiatives"))
            if not realized:
                realized = initiative_set(_semantic_get(row, "case"))
            mask = sum(1 << index for index, name in enumerate(approved_names) if name in realized)
            failed = tuple(name for name in approved_names if name not in realized)
            row["case"] = (
                "Realized " + " + ".join(realized)
                if realized else "Realized no approved initiatives"
            )
            row["realized_initiatives"] = "; ".join(realized) if realized else "None"
            row["failed_initiatives"] = "; ".join(failed) if failed else "None"
            row["realized_initiative_count"] = len(realized)
            joint_by_mask[mask] = row
        if joint_by_mask:
            normalized["joint_realization_risk_analysis"] = [
                joint_by_mask[mask] for mask in range(8) if mask in joint_by_mask
            ]

        price_by_mask: dict[int, dict[str, Any]] = {}
        for row in copied_rows("state_price_protection_analysis"):
            realized = initiative_set(_semantic_get(row, "realized_initiatives"))
            if not realized:
                realized = initiative_set(_semantic_get(row, "case"))
            mask = sum(
                1 << index for index, name in enumerate(approved_names)
                if name in realized
            )
            row["case"] = (
                "Realized " + " + ".join(realized)
                if realized else "Realized no approved initiatives"
            )
            row["realized_initiatives"] = "; ".join(realized) if realized else "None"
            current_ready = _semantic_get(row, "release_ready_at_current_price")
            protection_ready = _semantic_get(row, "price_protection_alone_release_ready")
            if current_ready is True:
                row["committee_price_action"] = "Retain current purchase equity; both hurdles clear"
            elif protection_ready is True:
                row["committee_price_action"] = "Require the state IRR price reduction before release"
            elif current_ready is False and protection_ready is False:
                row["committee_price_action"] = "Require both the state IRR price reduction and guaranteed pre-tax synergy cure before release"
            price_by_mask[mask] = row
        if price_by_mask:
            normalized["state_price_protection_analysis"] = [
                price_by_mask[mask] for mask in range(8) if mask in price_by_mask
            ]

        minimum_portfolio = _semantic_get(normalized, "minimum_viable_approved_portfolio")
        if initiative_set(minimum_portfolio) == approved_names:
            normalized["minimum_viable_approved_portfolio"] = "; ".join(approved_names)
        if "irr" in _normalize(_semantic_get(normalized, "binding_hurdle")):
            normalized["binding_hurdle"] = "Five-year equity IRR hurdle"
        hurdle_result = _semantic_get(normalized, "hurdles_met")
        if hurdle_result is True:
            normalized["hurdles_met"] = "EPS hurdle met; IRR hurdle met"
        elif isinstance(hurdle_result, dict):
            eps_met = _semantic_get(hurdle_result, "eps_hurdle_met")
            irr_met = _semantic_get(hurdle_result, "irr_hurdle_met")
            if eps_met is True and irr_met is True:
                normalized["hurdles_met"] = "EPS hurdle met; IRR hurdle met"

        failure_count = _semantic_get(
            normalized, "initiative_failure_cases_clearing_both"
        )
        if isinstance(failure_count, (list, tuple, set)) and not failure_count:
            normalized["initiative_failure_cases_clearing_both"] = 0

        release_text = _normalize(_semantic_get(normalized, "committee_release_status"))
        if (
            any(token in release_text for token in ("release ready", "release-ready", "approved for release"))
            and "eps" in release_text
            and "irr" in release_text
        ):
            normalized["committee_release_status"] = (
                "Release only with all three approved initiatives and at least the binding-hurdle realization rate"
            )
        portfolio_text = _normalize(_semantic_get(normalized, "portfolio_release_status"))
        if (
            any(token in portfolio_text for token in ("all three", "complete three", "full set"))
            and any(token in portfolio_text for token in ("zero", "no single", "any single"))
        ):
            normalized["portfolio_release_status"] = (
                "Release only the full three-initiative portfolio; no reduced approved subset clears both hurdles"
            )
        probabilistic_text = _normalize(
            _semantic_get(normalized, "probabilistic_release_status")
        )
        if (
            any(token in probabilistic_text for token in ("expected", "expectation"))
            and any(token in probabilistic_text for token in ("0 72", "72 00", "72 0", "72%", "72 %"))
            and any(token in probabilistic_text for token in ("seller protection", "unprotected downside", "shortfall"))
        ):
            normalized["probabilistic_release_status"] = (
                "Expected value clears both hurdles, but release requires all three approved initiatives because the probability of clearing both is below certainty"
            )
        price_text = _normalize(
            _semantic_get(normalized, "price_protection_release_status")
        )
        if (
            "price protection" in price_text
            and "irr" in price_text
            and "synergy guarantee" in price_text
            and "eps" in price_text
        ):
            normalized["price_protection_release_status"] = (
                "Do not release the current price without state-contingent IRR protection; states that miss EPS also require a guaranteed pre-tax synergy cure"
            )
        seller_text = _normalize(
            _semantic_get(normalized, "seller_protection_release_status")
        )
        if (
            any(token in seller_text for token in ("price credit", "price reduction"))
            and "synergy guarantee" in seller_text
            and any(token in seller_text for token in ("all eight", "every one of the eight"))
        ):
            normalized["seller_protection_release_status"] = (
                "Release only with the state-specific price credit and, where EPS still fails, the recurring pre-tax synergy guarantee documented before close"
            )

    if task_id == "task_057":
        def normalize_unavailable_sources(row: dict[str, Any]) -> None:
            for field in (
                "buyer_cash_source", "term_loan_source",
                "revolver_source", "seller_note_source",
            ):
                value = _semantic_get(row, field)
                if isinstance(value, str) and _normalize(value) in {
                    "unavailable", "not available", "not applicable", "n a", "na",
                }:
                    row[field] = 0.0

        use_rules = {
            "Purchase of equity": (),
            "Refinance target debt": (),
            "Transaction fees": (),
            "Financing fees": (),
            "Minimum target cash delivered": (),
        }
        normalized_uses = []
        for row in copied_rows("uses_analysis"):
            label = canonical_label(_semantic_get(row, "use_line"), use_rules)
            if label is None:
                continue
            row["use_line"] = label
            row["treatment"] = (
                "Closing cash delivered offsets uses"
                if label == "Minimum target cash delivered"
                else "Closing use"
            )
            normalized_uses.append(row)
        normalized["uses_analysis"] = normalized_uses

        cash_rules = {
            "Posted unrestricted cash": (),
            "Minimum buyer cash retained": (),
            "Cash above minimum balance": ("Cash above minimum retained balance",),
            "Board cash-use cap": (),
            "Approved buyer cash source": (
                "Buyer cash contribution",
                "Buyer cash deployed",
                "Buyer cash deployable",
            ),
        }
        normalized_cash = []
        for row in copied_rows("cash_capacity_bridge"):
            label = canonical_label(_semantic_get(row, "metric"), cash_rules)
            if label is None:
                continue
            row["metric"] = label
            normalized_cash.append(row)
        normalized["cash_capacity_bridge"] = normalized_cash

        source_rules = {
            "Buyer cash": ("Buyer cash contribution",),
            "Term loan": (),
            "Revolver": ("Revolving credit facility",),
            "Seller note": (),
        }
        normalized_sources = []
        for row in copied_rows("funding_source_analysis"):
            label = canonical_label(_semantic_get(row, "source"), source_rules)
            if label is None:
                continue
            row["source"] = label
            normalized_sources.append(row)
        priorities = [_semantic_get(row, "priority") for row in normalized_sources]
        if len(normalized_sources) == 4 and priorities == [0, 1, 2, 3]:
            for row in normalized_sources:
                row["priority"] = int(_semantic_get(row, "priority")) + 1
        normalized["funding_source_analysis"] = normalized_sources

        debt_rules = {
            "Opening buyer funded debt": ("Opening funded debt",),
            "New term loan funded": ("Term loan funded", "Term loan drawn"),
            "Revolver funded": ("Revolver drawn", "New revolver funded"),
            "Seller note funded": (),
            "Pro forma funded debt": (),
            "Approved covenant EBITDA": ("Approved pro forma covenant EBITDA",),
            "Closing leverage": ("Closing leverage pro forma funded debt covenant EBITDA",),
        }
        normalized_debt_by_label: dict[str, dict[str, Any]] = {}
        for row in copied_rows("debt_leverage_bridge"):
            raw_label = _normalize(_semantic_get(row, "metric"))
            if "opening buyer funded debt" in raw_label and any(
                token in raw_label for token in ("acct", "account")
            ) and "total" not in raw_label:
                continue
            label = canonical_label(_semantic_get(row, "metric"), debt_rules)
            if label is None:
                continue
            row["metric"] = label
            if label == "Closing leverage" and _normalize(_semantic_get(row, "unit")) in {
                "x", "x multiple", "multiple"
            }:
                row["unit"] = "multiple"
            normalized_debt_by_label[label] = row
        normalized["debt_leverage_bridge"] = [
            normalized_debt_by_label[label]
            for label in debt_rules
            if label in normalized_debt_by_label
        ]

        authority_rules = {
            "Purchase of equity": (),
            "Revolver availability": (),
            "Financing fees": (),
            "Pro forma covenant EBITDA": (),
            "Board cash-use cap": (),
            "Transaction fees": (),
            "Seller note committed": (),
            "Minimum buyer cash retained": (),
            "Refinance target debt": (),
            "Term loan commitment": (),
        }
        normalized_authority_by_label: dict[str, dict[str, Any]] = {}
        for row in copied_rows("source_authority_reconciliation"):
            label = canonical_label(_semantic_get(row, "driver"), authority_rules)
            if label is None:
                continue
            row["driver"] = label
            treatment = _normalize(_semantic_get(row, "authority_treatment"))
            current_supported = any(token in treatment for token in (
                "current", "controller", "7 5", "7.5", "approved",
                "committed", "board", "deal terms", "controls",
            ))
            prior_rejected = any(token in treatment for token in (
                "6 29", "6.29", "prior", "banker",
            )) and any(token in treatment for token in (
                "history", "trace", "superseded", "rejected", "audit",
            ))
            if current_supported and (
                prior_rejected
                or all(
                    _semantic_get(row, field) is not None
                    for field in ("current_value", "prior_value", "current_less_prior")
                )
            ):
                row["authority_treatment"] = "Use current controller-tied value; prior banker value is trace-only"
            normalized_authority_by_label[label] = row
        normalized["source_authority_reconciliation"] = [
            normalized_authority_by_label[label]
            for label in authority_rules
            if label in normalized_authority_by_label
        ]

        single_case_rules = {
            "Approved base": ("approved base", "approved_base"),
            "No buyer cash": ("no buyer cash", "no_buyer_cash"),
            "No revolver availability": ("no revolver", "no_revolver"),
            "No seller note": ("no seller note", "no_seller_note"),
            "No target cash delivered": ("no target cash", "no_target_cash"),
        }
        single_by_label: dict[str, dict[str, Any]] = {}
        for row in copied_rows("closing_stress_analysis"):
            normalize_unavailable_sources(row)
            case_text = _normalize(_semantic_get(row, "case"))
            for label, aliases in single_case_rules.items():
                if any(_normalize(alias) in case_text for alias in aliases):
                    row["case"] = label
                    single_by_label[label] = row
                    break
        normalized["closing_stress_analysis"] = [
            single_by_label[label]
            for label in single_case_rules
            if label in single_by_label
        ]

        outage_components = (
            ("buyer", "buyer cash"),
            ("revolver", "revolver"),
            ("seller", "seller note"),
            ("target", "target cash"),
        )
        outage_pairs = (
            (("buyer", "revolver"), "No buyer cash + no revolver"),
            (("buyer", "seller"), "No buyer cash + no seller note"),
            (("buyer", "target"), "No buyer cash + no target cash delivered"),
            (("revolver", "seller"), "No revolver + no seller note"),
            (("revolver", "target"), "No revolver + no target cash delivered"),
            (("seller", "target"), "No seller note + no target cash delivered"),
        )

        def outage_pair(value: Any) -> tuple[str, str] | None:
            text = _normalize(value)
            found = tuple(
                token for token, phrase in outage_components
                if phrase in text
                or (token == "seller" and "seller note" in text)
                or (token == "target" and "target cash" in text)
            )
            unique = tuple(token for token, _ in outage_components if token in found)
            return unique if len(unique) == 2 else None

        outage_label_by_pair = dict(outage_pairs)
        outage_by_pair: dict[tuple[str, str], dict[str, Any]] = {}
        for row in copied_rows("combined_source_outage_analysis"):
            normalize_unavailable_sources(row)
            pair = outage_pair(_semantic_get(row, "case"))
            if pair in outage_label_by_pair:
                row["case"] = outage_label_by_pair[pair]
                gap = _semantic_get(row, "funding_gap")
                if isinstance(gap, (int, float)):
                    row["committee_action"] = (
                        "Hold closing and cure the funding gap"
                        if float(gap) > 0.02
                        else "Funding remains complete; retain minimum cash and reapprove leverage"
                    )
                outage_by_pair[pair] = row
        normalized["combined_source_outage_analysis"] = [
            outage_by_pair[pair] for pair, _ in outage_pairs if pair in outage_by_pair
        ]

        cure_by_identity: dict[str, dict[str, Any]] = {}
        for row in copied_rows("combined_source_outage_cure_analysis"):
            normalize_unavailable_sources(row)
            pair = outage_pair(
                _semantic_get(row, "outage_case")
                or _semantic_get(row, "cure_case")
            )
            cure_text = _normalize(
                _semantic_get(row, "cure_type")
                or _semantic_get(row, "cure_case")
            )
            cure_type = (
                "Rescue equity" if "equity" in cure_text
                else "Purchase-price reduction"
                if "price" in cure_text or "purchase" in cure_text
                else None
            )
            if pair in outage_label_by_pair and cure_type is not None:
                outage_label = outage_label_by_pair[pair]
                identity = f"{outage_label} | {cure_type}"
                row["cure_case"] = identity
                row["outage_case"] = outage_label
                row["cure_type"] = cure_type
                residual_gap = _semantic_get(row, "residual_gap")
                release_ready = _semantic_get(row, "release_ready")
                if (
                    isinstance(residual_gap, (int, float))
                    and abs(float(residual_gap)) <= 0.02
                    and release_ready is True
                ):
                    row["committee_action"] = (
                        "Inject external rescue equity before closing; retain approved purchase price and minimum buyer cash"
                        if cure_type == "Rescue equity"
                        else "Reduce purchase price by the outage gap before closing; retain approved funding commitments and minimum buyer cash"
                    )
                cure_by_identity[identity] = row
        if cure_by_identity:
            normalized["combined_source_outage_cure_analysis"] = [
                cure_by_identity[f"{label} | {cure_type}"]
                for _, label in outage_pairs
                for cure_type in ("Rescue equity", "Purchase-price reduction")
                if f"{label} | {cure_type}" in cure_by_identity
            ]

        frontier_components = (
            ("buyer", "Buyer cash"),
            ("revolver", "Revolver"),
            ("seller", "Seller note"),
            ("target", "Target cash delivered"),
        )
        frontier_by_mask: dict[int, dict[str, Any]] = {}
        for row in copied_rows("financing_capacity_frontier"):
            raw_unavailable = _semantic_get(row, "unavailable_sources")
            if isinstance(raw_unavailable, (list, tuple, set)):
                unavailable_text = " ".join(str(value) for value in raw_unavailable)
            elif isinstance(raw_unavailable, dict):
                unavailable_text = " ".join(str(value) for value in raw_unavailable.values())
            else:
                unavailable_text = str(raw_unavailable or "")
            case_text = str(_semantic_get(row, "case") or "")
            combined_text = _normalize(unavailable_text + " " + case_text)
            base_case = any(token in combined_text for token in (
                "approved base", "base case", "none unavailable", "no outage",
            )) or _normalize(unavailable_text) in {"", "none", "n a", "na"}
            unavailable_tokens = [] if base_case else [
                token for token, phrase in frontier_components
                if _normalize(phrase) in combined_text
                or (token == "target" and "target cash" in combined_text)
            ]
            mask = sum(
                1 << index for index,(token,_) in enumerate(frontier_components)
                if token in unavailable_tokens
            )
            labels = [
                label for token,label in frontier_components
                if token in unavailable_tokens
            ]
            row["case"] = (
                "Approved base" if not labels else "Unavailable " + " + ".join(labels)
            )
            row["unavailable_sources"] = "; ".join(labels) if labels else "None"
            row["unavailable_source_count"] = len(labels)
            row["available_source_count"] = 4-len(labels)
            row["target_cash_available"] = "target" not in unavailable_tokens
            release_ready = _semantic_get(row, "release_ready_at_current_price")
            if release_ready is True:
                row["committee_action"] = "Release at current price within available committed sources"
            elif release_ready is False:
                row["committee_action"] = "Hold closing; require rescue equity or an equal signed purchase-price reduction"
            frontier_by_mask[mask] = row
        if frontier_by_mask:
            normalized["financing_capacity_frontier"] = [
                frontier_by_mask[mask] for mask in range(16) if mask in frontier_by_mask
            ]

        n_minus_one_by_mask: dict[int, dict[str, Any]] = {}
        for row in copied_rows("n_minus_one_signing_protection_analysis"):
            normalize_unavailable_sources(row)
            raw_unavailable = _semantic_get(row, "unavailable_sources")
            if isinstance(raw_unavailable, (list, tuple, set)):
                unavailable_text = " ".join(str(value) for value in raw_unavailable)
            elif isinstance(raw_unavailable, dict):
                unavailable_text = " ".join(str(value) for value in raw_unavailable.values())
            else:
                unavailable_text = str(raw_unavailable or "")
            case_text = str(_semantic_get(row, "case") or "")
            combined_text = _normalize(unavailable_text + " " + case_text)
            base_case = any(token in combined_text for token in (
                "approved base", "base case", "none unavailable", "no outage",
            )) or _normalize(unavailable_text) in {"", "none", "n a", "na"}
            unavailable_tokens = [] if base_case else [
                token for token, phrase in frontier_components
                if _normalize(phrase) in combined_text
                or (token == "target" and "target cash" in combined_text)
            ]
            mask = sum(
                1 << index for index,(token,_) in enumerate(frontier_components)
                if token in unavailable_tokens
            )
            labels = [
                label for token,label in frontier_components
                if token in unavailable_tokens
            ]
            row["case"] = (
                "Approved base" if not labels else "Unavailable " + " + ".join(labels)
            )
            row["unavailable_sources"] = "; ".join(labels) if labels else "None"
            release_ready = _semantic_get(row, "release_ready_without_external_equity")
            if release_ready is True:
                row["committee_action"] = "Release under the signed N-1 purchase price within available commitments"
            elif release_ready is False:
                row["committee_action"] = "Hold closing; fund the residual external equity before release"
            n_minus_one_by_mask[mask] = row
        if n_minus_one_by_mask:
            normalized["n_minus_one_signing_protection_analysis"] = [
                n_minus_one_by_mask[mask]
                for mask in range(16)
                if mask in n_minus_one_by_mask
            ]

        all_commitment_case_rules = {
            "No buyer cash": ("buyer cash",),
            "No term loan commitment": ("term loan",),
            "No revolver availability": ("revolver",),
            "No seller note": ("seller note",),
            "No target cash delivered": ("target cash",),
        }
        all_commitment_by_case: dict[str, dict[str, Any]] = {}
        for row in copied_rows("all_commitment_n_minus_one_analysis"):
            normalize_unavailable_sources(row)
            identity = _normalize(
                str(_semantic_get(row, "case") or "") + " "
                + str(_semantic_get(row, "unavailable_source") or "")
            )
            for label, aliases in all_commitment_case_rules.items():
                if any(alias in identity for alias in aliases):
                    row["case"] = label
                    row["unavailable_source"] = {
                        "No buyer cash": "Buyer cash",
                        "No term loan commitment": "Term loan",
                        "No revolver availability": "Revolver",
                        "No seller note": "Seller note",
                        "No target cash delivered": "Target cash delivered",
                    }[label]
                    release_ready = _semantic_get(row, "release_ready_without_backstop")
                    if release_ready is True:
                        row["required_protection"] = "None; committed sources fund closing and minimum buyer cash is retained"
                        row["committee_action"] = "Release this N-1 state within the surviving approved commitments"
                    elif label == "No term loan commitment":
                        row["required_protection"] = "Term-loan funding condition or equal alternate committed backstop"
                        row["committee_action"] = "Hold signing unless the term loan is a funding condition or a 43.4m alternate committed backstop is executed"
                    elif release_ready is False:
                        row["required_protection"] = "Signed purchase-price reduction or equal alternate committed backstop"
                        row["committee_action"] = "Hold signing until the funding gap is cured by signed price protection or an equal committed backstop"
                    all_commitment_by_case[label] = row
                    break
        if all_commitment_by_case:
            normalized["all_commitment_n_minus_one_analysis"] = [
                all_commitment_by_case[label]
                for label in all_commitment_case_rules
                if label in all_commitment_by_case
            ]

        severe = _semantic_get(normalized, "most_severe_combined_source_loss_case")
        severe_pair = outage_pair(severe)
        if severe_pair in outage_label_by_pair:
            normalized["most_severe_combined_source_loss_case"] = outage_label_by_pair[severe_pair]
        resilient = _semantic_get(normalized, "resilient_source_loss_case")
        resilient_text = _normalize(resilient)
        if "buyer" in resilient_text:
            normalized["resilient_source_loss_case"] = "No buyer cash"
        if "revolver" in _normalize(_semantic_get(normalized, "binding_downside")):
            normalized["binding_downside"] = "No revolver availability"

        release = _semantic_get(normalized, "closing_release_status")
        if isinstance(release, dict):
            decision = _normalize(_semantic_get(release, "decision"))
            gap_value = _semantic_get(release, "funding_gap")
            notes = " ".join(_normalize(value) for value in release.values())
            if (
                "release" in decision
                and isinstance(gap_value, (int, float))
                and abs(float(gap_value)) <= 0.02
                and any(token in notes for token in ("committed", "no uncommitted", "fully fund"))
                and any(token in notes for token in ("minimum", "cash use cap", "cash-use cap"))
            ):
                normalized["closing_release_status"] = "Fully funded within approved commitments; retain the minimum buyer cash balance and use no unapproved financing"
        elif isinstance(release, str):
            release_text = _normalize(release)
            gap_value = _semantic_get(normalized, "funding_gap")
            cash_value = _semantic_get(normalized, "post_close_buyer_cash")
            if (
                any(token in release_text for token in ("release", "approved", "fully fund"))
                and not any(token in release_text for token in ("not approved", "not release", "do not release", "hold"))
                and isinstance(gap_value, (int, float))
                and abs(float(gap_value)) <= 0.02
                and isinstance(cash_value, (int, float))
                and float(cash_value) >= 2_250_000
            ):
                normalized["closing_release_status"] = "Fully funded within approved commitments; retain the minimum buyer cash balance and use no unapproved financing"

        stress_release = _normalize(_semantic_get(normalized, "closing_stress_release_status"))
        if (
            any(token in stress_release for token in ("hold", "no release", "cannot release", "conditional"))
            and any(token in stress_release for token in ("stress", "revolver", "seller", "loss"))
        ):
            normalized["closing_stress_release_status"] = "Release approved base; hold if revolver or seller-note availability is lost"

        combined_release = _normalize(_semantic_get(normalized, "combined_source_loss_release_status"))
        if (
            any(token in combined_release for token in ("hold", "no release", "do not release", "requires cure"))
            and any(token in combined_release for token in ("two source", "outage", "cure"))
        ):
            normalized["combined_source_loss_release_status"] = "Hold any two-source outage with a funding gap; require rescue equity or an equal purchase-price reduction before release"

        cure_release = _normalize(_semantic_get(normalized, "outage_cure_release_status"))
        if (
            any(token in cure_release for token in ("every", "all", "two source"))
            and any(token in cure_release for token in ("cure", "curable"))
            and any(token in cure_release for token in ("release", "close gap", "closes its gap"))
        ):
            normalized["outage_cure_release_status"] = "Every two-source outage requires a funded rescue-equity injection or an equal signed purchase-price reduction before committee release"

        frontier_release = _normalize(_semantic_get(normalized, "financing_frontier_release_status"))
        if (
            "release" in frontier_release
            and any(token in frontier_release for token in ("zero gap", "cure", "availability", "full availability"))
        ):
            normalized["financing_frontier_release_status"] = "Release only zero-gap availability states; every other state requires signed rescue equity or an equal purchase-price reduction before closing"

    if task_id == "task_064":
        # Controller workpapers commonly use "approved/reversed" rather than
        # "approved/incorrect override" labels. Copy only values the agent
        # actually supplied; every amount and classification still has to pass
        # the exact deterministic finance checks.
        override_aliases = {
            "approved_policy_conclusion": ("approved_conclusion",),
            "incorrect_override": ("reversed_conclusion",),
            "override_adjusted_ebitda": (
                "reversed_adjusted_ebitda",
                "adjusted_ebitda_if_reversed",
            ),
            "override_ebitda_misstatement": ("adjusted_ebitda_misstatement",),
            "override_adjusted_ebitda_margin": (
                "reversed_adjusted_ebitda_margin",
                "adjusted_ebitda_margin_if_reversed",
            ),
        }
        for row in copied_rows("non_gaap_policy_override_analysis"):
            for target, aliases in override_aliases.items():
                if _semantic_get(row, target) is not None:
                    continue
                for alias in aliases:
                    value = _semantic_get(row, alias)
                    if value is not None:
                        row[target] = value
                        break

    return normalized


def _grade_console(task_id: str, answer: Any) -> dict[str, Any]:
    gold = load_corporate_finance_gold(task_id)
    mapping = _normalize_console_mapping(task_id, _answer_mapping(answer))
    criteria = [
        _console_criterion(mapping, spec, task_id=task_id)
        for spec in gold["criteria"]
    ]
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


def _xlsx_label_cell_index(workbook):
    """Index normalized workbook labels once for repeated atomic criteria.

    The production workbooks deliberately expose hundreds of controller-tied
    outputs under exact row labels.  Re-normalizing every cell for every atomic
    criterion is quadratic and can consume the residual rollout wall clock.
    Exact-label lookups cover the normal template path; callers retain their
    existing professional-alias fallback when an exact label is absent or does
    not carry the expected value.
    """

    cached = getattr(workbook, "_alder_label_cell_index", None)
    if cached is not None:
        return cached
    cells: list[tuple[str, int, int, Any]] = []
    exact: dict[str, list[tuple[str, int, int, Any]]] = {}
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                record = (sheet_name, cell.row, cell.column, cell.value)
                cells.append(record)
                normalized = _normalize(cell.value)
                if normalized:
                    exact.setdefault(normalized, []).append(record)
    cached = (cells, exact)
    setattr(workbook, "_alder_label_cell_index", cached)
    return cached


def _xlsx_label_value(
    workbook,
    values,
    label: str,
    expected: Any,
    *,
    directional_strings: bool = False,
) -> bool:
    wanted = _normalize(label)
    indexed_cells, exact_labels = _xlsx_label_cell_index(workbook)

    def record_matches(record: tuple[str, int, int, Any]) -> bool:
        sheet_name, row_number, column_number, _value = record
        value_sheet = (
            values[sheet_name]
            if sheet_name in values.sheetnames
            else workbook[sheet_name]
        )
        candidates = [
            value_sheet.cell(row_number, column_number + offset).value
            for offset in (1, 2, 3)
        ]
        candidates.append(value_sheet.cell(row_number + 1, column_number).value)
        return _value_candidates_match(
            candidates, expected, directional_strings=directional_strings
        )

    exact_records = exact_labels.get(wanted, [])
    if any(record_matches(record) for record in exact_records):
        return True
    exact_record_ids = {(record[0], record[1], record[2]) for record in exact_records}
    for record in indexed_cells:
        sheet_name, row_number, column_number, value = record
        if (sheet_name, row_number, column_number) in exact_record_ids:
            continue
        if not (semantic_equal(value, label) or contains_concept(value, label)):
            continue
        if record_matches(record):
            return True
    return False


def _xlsx_exact_label_value(
    workbook,
    values,
    label: str,
    expected: Any,
    *,
    directional_strings: bool = False,
) -> bool:
    """Require the value beside the exact normalized workbook label.

    Task 027 publishes a complete controller output registry, so an incorrect
    value on that named row must not be rescued by a similar label elsewhere
    in the model.  Other tasks retain the broader professional-alias fallback.
    """

    _indexed_cells, exact_labels = _xlsx_label_cell_index(workbook)
    for sheet_name, row_number, column_number, _value in exact_labels.get(
        _normalize(label), []
    ):
        value_sheet = (
            values[sheet_name]
            if sheet_name in values.sheetnames
            else workbook[sheet_name]
        )
        candidates = [
            value_sheet.cell(row_number, column_number + offset).value
            for offset in (1, 2, 3)
        ]
        candidates.append(value_sheet.cell(row_number + 1, column_number).value)
        if _value_candidates_match(
            candidates, expected, directional_strings=directional_strings
        ):
            return True
    return False


def _xlsx_label_formula(workbook, label: str) -> tuple[bool, str]:
    wanted = _normalize(label)
    indexed_cells, exact_labels = _xlsx_label_cell_index(workbook)
    label_locations: list[str] = []
    seen: set[tuple[str, int, int]] = set()

    def check_record(record: tuple[str, int, int, Any]) -> tuple[bool, str | None]:
        sheet_name, row_number, column_number, _value = record
        sheet = workbook[sheet_name]
        coordinate = sheet.cell(row_number, column_number).coordinate
        label_locations.append(f"{sheet_name}!{coordinate}")
        candidates = [
            sheet.cell(row_number, column_number + offset)
            for offset in (1, 2, 3)
            if column_number + offset <= sheet.max_column
        ]
        if row_number + 1 <= sheet.max_row:
            candidates.append(sheet.cell(row_number + 1, column_number))
        for candidate in candidates:
            formula = candidate.value
            if not isinstance(formula, str) or not formula.startswith("="):
                continue
            # A formula that merely wraps the released answer (for example
            # =12345) is not model lineage. Require a cell/range reference.
            if re.search(
                r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)?!?\$?[A-Z]{1,3}\$?\d+",
                formula,
            ):
                return True, f"{sheet_name}!{candidate.coordinate}={formula}"
        return False, None

    for record in exact_labels.get(wanted, []):
        seen.add((record[0], record[1], record[2]))
        matched, evidence = check_record(record)
        if matched:
            return True, str(evidence)
    for record in indexed_cells:
        identity = (record[0], record[1], record[2])
        if identity in seen:
            continue
        if not (semantic_equal(record[3], label) or contains_concept(record[3], label)):
            continue
        matched, evidence = check_record(record)
        if matched:
            return True, str(evidence)
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


def _task_053_relative_decision_score(actual: float, expected: float) -> float:
    """Give bounded partial credit for a disclosed but inexact DCF conclusion.

    A valuation conclusion that is close to the controller-tied result is more
    useful than one that is materially wrong.  The score therefore declines
    continuously with absolute relative error and reaches zero at a 15 percent
    miss.  Full decision credit remains reserved for the exact headline-value
    criterion; this companion score is capped at 0.5.
    """

    denominator = max(abs(float(expected)), 1e-12)
    relative_error = abs(float(actual) - float(expected)) / denominator
    return round(0.5 * max(0.0, 1.0 - relative_error / 0.15), 6)


def _task_053_decision_support(workbook, values, gold: dict[str, Any]) -> dict[str, Any]:
    """Measure the accuracy of labeled DCF conclusions that miss exact gold.

    This preserves truthful professional ordering among otherwise incomplete
    models.  It does not relax the exact headline criteria or strict-pass
    standard, and it does not award credit for an omitted or unlabeled value.
    """

    answer = gold["answer"]
    critical_labels = (
        "enterprise_value",
        "equity_value",
        "present_value_of_terminal_value",
        "terminal_value_share_of_enterprise_value",
        "downside_equity_value",
        "upside_equity_value",
        "equity_value_sensitivity_range",
    )
    unit_text = " ".join(
        str(cell.value)
        for sheet in workbook.worksheets
        for row in sheet.iter_rows()
        for cell in row
        if cell.value is not None
    ).casefold()
    millions_disclosed = any(
        marker in unit_text
        for marker in (
            "$mm",
            "$ mm",
            "usd millions",
            "$ in millions",
            "$ millions",
            "all dollars in millions",
            "dollars in millions",
        )
    )
    thousands_disclosed = any(
        marker in unit_text
        for marker in (
            "$000",
            "$ 000",
            "usd thousands",
            "$ in thousands",
            "$ thousands",
            "all dollars in thousands",
            "dollars in thousands",
        )
    )

    support: dict[str, dict[str, Any]] = {}
    for label in critical_labels:
        expected = float(answer[label])
        if abs(expected) >= 100_000 and millions_disclosed:
            target = expected / 1_000_000.0
        elif abs(expected) >= 100_000 and thousands_disclosed:
            target = expected / 1_000.0
        else:
            target = expected
        aliases = [label.replace("_", " "), *_TASK_053_LABEL_ALIASES.get(label, [])]
        candidates: list[tuple[float, str]] = []
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
            for row in sheet.iter_rows():
                row_text = " | ".join(
                    str(cell.value) for cell in row if cell.value is not None
                )
                if not any(contains_concept(row_text, alias) for alias in aliases):
                    continue
                for cell in row:
                    cached = value_sheet[cell.coordinate].value
                    if isinstance(cached, (int, float)) and not isinstance(cached, bool):
                        candidates.append((float(cached), f"{sheet_name}!{cell.coordinate}"))

        criterion_id = f"headline_values__{label}"
        if not candidates:
            support[criterion_id] = {
                "score": 0.0,
                "evidence": "no labeled numeric DCF conclusion was found",
            }
            continue

        actual, location = min(
            candidates,
            key=lambda item: abs(item[0] - target) / max(abs(target), 1e-12),
        )
        score = _task_053_relative_decision_score(actual, target)
        relative_error = abs(actual - target) / max(abs(target), 1e-12)
        support[criterion_id] = {
            "score": score,
            "evidence": (
                f"labeled value at {location}; relative error={relative_error:.6%}; "
                "partial decision credit declines continuously to zero at a 15% miss"
            ),
            "actual": actual,
            "expected": target,
            "relative_error": round(relative_error, 8),
        }
    return {"criteria": support}


def _display_tolerance(expected: float) -> float:
    if abs(expected) <= 10:
        return 0.0005
    return max(0.02, abs(expected) * 0.0005)


def _sample_directional_semantic_value_matches(actual: Any, expected: str) -> bool:
    """Accept professional sample-task wording without accepting its opposite.

    Several sample outputs contain decision states and classifications where
    keyword overlap alone is unsafe (``ready ... no``, ``hold is not
    required``, or ``leverage is not the binding constraint``). This guard is
    intentionally called only from the sample-task paths below so it does not
    change grading semantics for the other Alder Ridge tasks.
    """

    if not semantic_value_matches(actual, expected):
        return False
    normalized = _normalize(actual)
    expected_normalized = _normalize(expected)
    if not normalized or not expected_normalized:
        return False

    if "within commitment and cash floor" in expected_normalized:
        clean = normalized.replace("no breach", "").replace("no shortfall", "")
        if any(phrase in clean for phrase in (
            "breach", "shortfall", "over commitment", "below cash floor",
            "outside commitment", "insufficient liquidity",
        )):
            return False
    if "capacity cleared" in expected_normalized and any(
        phrase in normalized for phrase in (
            "not cleared", "sequencing required", "decision required",
            "recovery required", "rephase required", "shortfall remains",
            "constraint remains", "insufficient capacity",
        )
    ):
        return False
    if expected_normalized == "selected" and any(
        phrase in normalized for phrase in ("not selected", "unselected", "excluded")
    ):
        return False
    if expected_normalized == "not selected" and not any(
        phrase in normalized for phrase in ("not selected", "unselected", "excluded")
    ):
        return False
    if "feasible alternative" in expected_normalized:
        return not any(phrase in normalized for phrase in (
            "not feasible", "no feasible", "infeasible", "no alternative",
            "feasible alternative no", "feasible alternative is false",
        ))
    if "ready for controller posting" in expected_normalized and any(
        phrase in normalized for phrase in (
            "not ready", "do not post", "posting prohibited", "hold for",
        )
    ):
        return False
    guidance_text = re.sub(
        r"guidance update required\s*(?::|=|-)?\s*(?:false|no|0)\b",
        "",
        normalized,
    )
    if "maintain with heightened monitoring" in expected_normalized and any(
        phrase in guidance_text for phrase in (
            "do not maintain", "guidance update required", "update guidance now",
            "revise guidance now", "withdraw guidance now",
            "reduce guidance now", "cut guidance now",
        )
    ):
        return False
    if expected_normalized == "release" or expected_normalized.startswith("release only"):
        if any(phrase in normalized for phrase in (
            "do not release", "not released", "release not approved",
            "hold", "withhold", "not ready",
        )):
            return False
    if any(phrase in expected_normalized for phrase in (
        "hold", "sequencing required", "decision required",
        "mitigation required", "recovery plan",
    )) and any(phrase in normalized for phrase in (
        "hold not required", "no hold", "sequencing not required",
        "decision not required", "mitigation not required",
        "no mitigation", "unconditional release", "release approved",
        "approved to proceed", "clear to proceed",
    )):
        return False
    escaped = re.escape(expected_normalized).replace(r"\ ", r"\s+")
    if re.search(
        rf"\b{escaped}\b\s*(?:(?:is|are|was|were|remains?)\s+|[:\-]\s*)?(?:not|no)\b",
        normalized,
    ):
        return False
    return True


def _value_candidates_match(
    candidates: Iterable[Any],
    expected: Any,
    *,
    directional_strings: bool = False,
) -> bool:
    if isinstance(expected, list):
        return ordered_semantic_list_matches(list(candidates), expected)
    if isinstance(expected, str):
        matcher = (
            _sample_directional_semantic_value_matches
            if directional_strings
            else semantic_value_matches
        )
        return any(
            date_matches(candidate, expected) or matcher(candidate, expected)
            for candidate in candidates
        )
    if isinstance(expected, bool):
        return any(_boolean_matches(candidate, expected) for candidate in candidates)
    return any(
        _close(candidate, float(expected), abs_tol=_display_tolerance(float(expected)), rel_tol=0.0)
        for candidate in candidates
    )


def _task_035_row_match(
    workbook,
    values,
    label: str,
    expected: Any,
    *,
    require_formula: bool,
) -> tuple[bool, str]:
    """Accept ordinary planning labels while preserving value association.

    Task035 is intentionally laid out as normal case, portfolio, and executive
    tables rather than an answer-key registry.  A required output may therefore
    be identified by its case/portfolio row instead of a private snake_case
    label.  Match the exact expected value only on a row carrying an approved
    business alias, and require the matching cell itself to be source-linked
    for formula-lineage criteria.
    """

    aliases = [label.replace("_", " "), *_TASK_035_LABEL_ALIASES.get(label, [])]
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            if not any(
                semantic_equal(cell.value, alias) or contains_concept(cell.value, alias)
                for cell in row
                for alias in aliases
            ):
                continue
            for cell in row:
                cached = value_sheet[cell.coordinate].value
                if require_formula and not _is_source_linked_formula(cell.value):
                    continue
                if _value_candidates_match(
                    [cached],
                    expected,
                    directional_strings=True,
                ):
                    return (
                        True,
                        f"{sheet_name}!{cell.coordinate}={cell.value!r}; "
                        f"cached={cached!r}; associated aliases={aliases!r}",
                    )
    return False, f"no associated Task035 row value matched {label!r}"


def _task_076_quarter_schedule_match(
    workbook,
    values,
    quarter_number: int,
    metric: str,
    expected: Any,
    *,
    require_formula: bool,
) -> tuple[bool, str]:
    """Require the correct value at a disclosed quarter/metric intersection.

    A professional planning workbook may put quarters in columns or in rows.
    This accepts either orientation, including explicit USD-thousands display,
    but does not let a value from the wrong quarter satisfy the criterion.
    """

    quarter_aliases = (
        f"Q{quarter_number} FY27",
        f"FY27 Q{quarter_number}",
        f"Q{quarter_number}",
    )
    metric_aliases = _TASK_076_QUARTER_METRIC_ALIASES[metric]

    def concept_matches(value: Any, aliases: Iterable[str]) -> bool:
        return any(
            semantic_equal(value, alias) or contains_concept(value, alias)
            for alias in aliases
        )

    def value_matches(value: Any) -> bool:
        if isinstance(expected, bool):
            return _boolean_matches(value, expected)
        if isinstance(expected, str):
            return date_matches(value, expected) or semantic_value_matches(value, expected)
        if not isinstance(value, (int, float)) or isinstance(value, bool):
            return False
        targets = [float(expected)]
        if abs(float(expected)) > 10:
            targets.append(float(expected) / 1_000.0)
        return any(
            _close(
                float(value),
                target,
                abs_tol=max(0.00002, abs(target) * 0.0005),
                rel_tol=0.0,
            )
            for target in targets
        )

    def formula_ok(value: Any) -> bool:
        return not require_formula or _is_source_linked_formula(value)

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            row_has_metric = any(
                concept_matches(cell.value, metric_aliases) for cell in row
            )
            row_has_quarter = any(
                concept_matches(cell.value, quarter_aliases) for cell in row
            )
            if row_has_metric:
                for cell in row:
                    cached = value_sheet[cell.coordinate].value
                    if not value_matches(cached) or not formula_ok(cell.value):
                        continue
                    headers = [
                        sheet.cell(header_row, cell.column).value
                        for header_row in range(max(1, cell.row - 15), cell.row)
                    ]
                    if row_has_quarter or any(
                        concept_matches(header, quarter_aliases) for header in headers
                    ):
                        return (
                            True,
                            f"{sheet_name}!{cell.coordinate}={cell.value!r}; "
                            f"cached={cached!r}; quarter=Q{quarter_number}; metric={metric!r}",
                        )
            if row_has_quarter:
                for cell in row:
                    cached = value_sheet[cell.coordinate].value
                    if not value_matches(cached) or not formula_ok(cell.value):
                        continue
                    headers = [
                        sheet.cell(header_row, cell.column).value
                        for header_row in range(max(1, cell.row - 15), cell.row)
                    ]
                    if any(
                        concept_matches(header, metric_aliases) for header in headers
                    ):
                        return (
                            True,
                            f"{sheet_name}!{cell.coordinate}={cell.value!r}; "
                            f"cached={cached!r}; quarter=Q{quarter_number}; metric={metric!r}",
                        )
    return (
        False,
        f"no exact Q{quarter_number} FY27 / {metric} schedule value matched",
    )


def _task_076_branch_schedule_match(
    workbook,
    values,
    branch: str,
    metric: str,
    expected: Any,
    *,
    require_formula: bool,
) -> tuple[bool, str]:
    """Require a formula/value at the disclosed branch/metric intersection."""

    branch_aliases = (branch, f"{branch} branch")
    metric_aliases = _TASK_076_BRANCH_METRIC_ALIASES[metric]

    def concept_matches(value: Any, aliases: Iterable[str]) -> bool:
        return any(
            semantic_equal(value, alias) or contains_concept(value, alias)
            for alias in aliases
        )

    def value_matches(value: Any) -> bool:
        if not isinstance(value, (int, float)) or isinstance(value, bool):
            return False
        targets = [float(expected)]
        if abs(float(expected)) > 10:
            targets.append(float(expected) / 1_000.0)
        return any(
            _close(
                float(value),
                target,
                abs_tol=max(0.00002, abs(target) * 0.0005),
                rel_tol=0.0,
            )
            for target in targets
        )

    def formula_ok(value: Any) -> bool:
        return not require_formula or _is_source_linked_formula(value)

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in sheet.iter_rows():
            row_has_metric = any(
                concept_matches(cell.value, metric_aliases) for cell in row
            )
            row_has_branch = any(
                concept_matches(cell.value, branch_aliases) for cell in row
            )
            if row_has_metric:
                for cell in row:
                    cached = value_sheet[cell.coordinate].value
                    if not value_matches(cached) or not formula_ok(cell.value):
                        continue
                    headers = [
                        sheet.cell(header_row, cell.column).value
                        for header_row in range(max(1, cell.row - 15), cell.row)
                    ]
                    if row_has_branch or any(
                        concept_matches(header, branch_aliases) for header in headers
                    ):
                        return (
                            True,
                            f"{sheet_name}!{cell.coordinate}={cell.value!r}; "
                            f"cached={cached!r}; branch={branch!r}; metric={metric!r}",
                        )
            if row_has_branch:
                for cell in row:
                    cached = value_sheet[cell.coordinate].value
                    if not value_matches(cached) or not formula_ok(cell.value):
                        continue
                    headers = [
                        sheet.cell(header_row, cell.column).value
                        for header_row in range(max(1, cell.row - 15), cell.row)
                    ]
                    if any(
                        concept_matches(header, metric_aliases) for header in headers
                    ):
                        return (
                            True,
                            f"{sheet_name}!{cell.coordinate}={cell.value!r}; "
                            f"cached={cached!r}; branch={branch!r}; metric={metric!r}",
                        )
    return False, f"no exact {branch} / {metric} schedule value matched"


def _task_076_row_match(workbook, values, label: str, expected: Any, *, require_formula: bool) -> tuple[bool, str]:
    """Read the integrated plan's disclosed $000 schedules without losing lineage.

    The model may expose a required headline as the exact matching quarter,
    consolidated column, minimum, or maximum within a labeled formula row.
    Accept only an exact cached result (with the explicitly declared $000 unit)
    and, for lineage checks, require the matching cell itself to be a formula.
    """
    quarter_match = re.fullmatch(r"q([1-4])_(.+)", label)
    if quarter_match and quarter_match.group(2) in _TASK_076_QUARTER_METRIC_ALIASES:
        return _task_076_quarter_schedule_match(
            workbook,
            values,
            int(quarter_match.group(1)),
            quarter_match.group(2),
            expected,
            require_formula=require_formula,
        )
    branch_match = re.fullmatch(r"(construction|service|controls)_(.+)", label)
    if branch_match and branch_match.group(2) in _TASK_076_BRANCH_METRIC_ALIASES:
        return _task_076_branch_schedule_match(
            workbook,
            values,
            branch_match.group(1).title(),
            branch_match.group(2),
            expected,
            require_formula=require_formula,
        )
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


def _scoped_xlsx_label_evidence(
    workbook,
    values,
    label: str,
    *,
    aliases: Iterable[str] = (),
) -> tuple[str, bool, str]:
    """Return only the submitted workbook row needed for one semantic check."""

    indexed_cells, exact_labels = _xlsx_label_cell_index(workbook)
    accepted_labels = (label, *tuple(aliases))
    matches = [
        record
        for accepted in accepted_labels
        for record in exact_labels.get(_normalize(accepted), [])
    ]
    if not matches:
        matches = [
            record for record in indexed_cells
            if any(
                semantic_equal(record[3], accepted)
                or contains_concept(record[3], accepted)
                for accepted in accepted_labels
            )
        ]
    if not matches:
        return "required labeled output row is missing", False, "labeled output row is missing"

    rendered: list[str] = []
    has_submitted_value = False
    for sheet_name, row_number, column_number, _ in matches[:3]:
        formula_sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else formula_sheet
        parts: list[str] = []
        for column in range(1, formula_sheet.max_column + 1):
            formula_cell = formula_sheet.cell(row_number, column)
            cached = value_sheet.cell(row_number, column).value
            if formula_cell.value is None and cached is None:
                continue
            if isinstance(formula_cell.value, str) and formula_cell.value.startswith("="):
                parts.append(f"{formula_cell.coordinate}=FORMULA({formula_cell.value})=>{cached!r}")
            else:
                parts.append(f"{formula_cell.coordinate}={formula_cell.value!r}")
            if (
                column > column_number
                and cached not in (None, "")
                and _normalize(cached) not in {"-", "—", "tbd", "to be completed"}
            ):
                has_submitted_value = True
        if parts:
            rendered.append(f"SHEET {sheet_name} ROW {row_number}: " + " | ".join(parts))
    return (
        "\n".join(rendered)[:12_000],
        has_submitted_value,
        f"labeled row found and a submitted result cell is populated={has_submitted_value}",
    )


def _scoped_xlsx_token_evidence(
    workbook,
    values,
    token: str,
    *,
    aliases: Iterable[str] = (),
) -> tuple[str, bool, str]:
    """Return bounded workbook rows containing one provenance concept."""

    accepted_concepts = (token, *tuple(aliases))
    rows: list[str] = []
    completed = False
    for sheet_name in workbook.sheetnames:
        formula_sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else formula_sheet
        for row_number in range(1, formula_sheet.max_row + 1):
            row_values = [formula_sheet.cell(row_number, column).value for column in range(1, formula_sheet.max_column + 1)]
            if not any(
                contains_concept(value, concept)
                for value in row_values if value is not None
                for concept in accepted_concepts
            ):
                continue
            parts: list[str] = []
            for column in range(1, formula_sheet.max_column + 1):
                formula_cell = formula_sheet.cell(row_number, column)
                cached = value_sheet.cell(row_number, column).value
                if formula_cell.value is None and cached is None:
                    continue
                if isinstance(formula_cell.value, str) and formula_cell.value.startswith("="):
                    parts.append(f"{formula_cell.coordinate}=FORMULA({formula_cell.value})=>{cached!r}")
                else:
                    parts.append(f"{formula_cell.coordinate}={formula_cell.value!r}")
                if column > 1 and cached not in (None, "") and _normalize(cached) not in {"-", "—", "tbd", "to be completed"}:
                    completed = True
            if parts:
                rows.append(f"SHEET {sheet_name} ROW {row_number}: " + " | ".join(parts))
    return (
        "\n".join(rows[:8])[:12_000] or "required supporting row is missing",
        completed,
        f"supporting row found with populated submitted content={completed}",
    )


_TASK_035_SOURCE_REFERENCES = {
    "controller-tied": {
        "aliases": (
            "FY27 planning assumptions - v6 controller tie.xlsx",
            "FY27 planning assumptions v6 controller tie",
            "controller tie",
            "controller-tied",
        ),
        "answer_key": (
            "Identify the current controller-tied FY27 planning-assumptions workbook or "
            "the current controller-tied locked planning-input tabs, and distinguish them "
            "from prior/superseded drafts. The exact filename or an unambiguous shortened "
            "internal reference is sufficient because the task does not disclose a required filename."
        ),
    },
    "management correspondence": {
        "aliases": (
            "7.4.26_0711am - FY27 plan first pass + branch submissions.eml",
            "FY27 plan first pass + branch submissions",
            "7.5.26_0618am - FY27 plan review comments - no stretch case.eml",
            "FY27 plan review comments - no stretch case",
            "steering review comments",
            "controller follow-up",
            "management correspondence",
        ),
        "answer_key": (
            "Identify at least one of the two governing FY27 management threads: the "
            "July 4 first-pass/branch-submission thread or the July 5 review-comments/no-"
            "stretch-case thread. Exact filename, subject, or an unambiguous ordinary "
            "business reference is sufficient."
        ),
    },
    "policy": {
        "aliases": (
            "FY27 planning definitions + scenario guardrails - APPROVED.pdf",
            "FY27 planning definitions + scenario guardrails",
            "planning definitions",
            "scenario guardrails",
            "approved planning policy",
            "approved capacity-response policy",
            "capacity response policy",
            "executed contingent labor framework",
            "contingent labor framework",
            "policy",
        ),
        "answer_key": (
            "Identify an applicable approved authority used by the model: the FY27 "
            "planning-definitions/scenario-guardrails policy, the approved capacity-response "
            "policy, or the executed contingent-labor framework. The exact filename or an "
            "unambiguous shortened reference is sufficient."
        ),
    },
}


def _pptx_slide_evidence(path: Path, slide_numbers: Iterable[int]) -> tuple[str, bool, str]:
    """Render only the slides relevant to one presentation criterion."""

    presentation = Presentation(path)
    chunks: list[str] = []
    authored_values: list[str] = []
    placeholders = {
        "", "-", "—", "tbd", "to be completed", "complete with supported conclusion",
        "refresh coverage", "draft data not refreshed",
    }
    for slide_number in slide_numbers:
        if slide_number < 1 or slide_number > len(presentation.slides):
            continue
        slide = presentation.slides[slide_number - 1]
        entries: list[str] = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and str(shape.text or "").strip():
                entries.append(str(shape.text).strip())
            if getattr(shape, "has_table", False):
                entries.extend(" | ".join(cell.text.strip() for cell in row.cells) for row in shape.table.rows)
        authored_values.extend(
            entry for entry in entries
            if _normalize(entry) not in placeholders and not _normalize(entry).isdigit()
        )
        chunks.append(f"SLIDE {slide_number}:\n" + "\n".join(entries))
    substantive = bool(chunks) and len(authored_values) >= 2
    return (
        "\n\n".join(chunks)[:12_000] or "required slide is missing",
        substantive,
        f"required slide exists with non-placeholder content={substantive}",
    )


_TASK_035_DETERMINISTIC_TEXT_IDS = {
    "headline_values__first_constrained_month",
    "headline_values__gross_commitment_first_constrained_month",
    "headline_values__controlling_capacity_case",
    "headline_values__execution_portfolio_highest_priority_deferred_project",
    "headline_values__executive_recovery_highest_priority_deferred_project",
}


def _task_035_status_reference(label: str, expected: Any) -> dict[str, Any]:
    meaning_by_status = {
        "capacity cleared": "available capacity and permitted remedies fully cover required hours; no unresolved shortfall remains",
        "executive sequencing required": "a residual shortfall remains after permitted remedies and leadership must sequence or rephase work",
        "portfolio resequencing required": "the execution schedule leaves deferred work and requires portfolio-level resequencing",
        "executive decision required": "the project remains deferred or constrained and requires an explicit management decision",
        "executive portfolio decision required": "the portfolio retains deferred work and requires an executive portfolio decision before release",
        "hold for executive portfolio sequencing": "hold release until the deferred portfolio is sequenced or rephased",
        "recover or rephase deferred priority backlog before release": "recover capacity for, or rephase, the highest-priority deferred backlog before release",
        "hold for executive recovery plan": "hold release until management approves and demonstrates the required recovery plan",
    }
    normalized_expected = _normalize(expected)
    decision_scope = "full decision"
    grading_boundary = (
        "Judge the complete operational decision in this field. The separately scored "
        "amounts, project identifier, and formula outputs do not need to be repeated."
    )

    if re.fullmatch(r"(?:gross_commitment_)?20\d\d_\d\d_capacity_release_status", label):
        decision_scope = "monthly capacity status"
        if normalized_expected == "capacity cleared":
            required_meaning = (
                "the month has no unresolved capacity constraint. Accept Capacity cleared, "
                "Cleared, Within capacity, Capacity available, No constraint, Not constrained, "
                "or an equally clear professional equivalent"
            )
        else:
            required_meaning = (
                "the month retains a capacity shortfall and therefore requires executive "
                "sequencing, rephasing, or an equivalent constrained-capacity decision"
            )
        grading_boundary = (
            "Judge only whether this monthly capacity row communicates clear/available versus "
            "constrained/short. Detailed release and recovery actions are scored separately."
        )
    elif re.fullmatch(r"execution_20\d\d_\d\d_release_status", label):
        decision_scope = "monthly execution status"
        if normalized_expected == "capacity cleared":
            required_meaning = (
                "the month is released or cleared because no unresolved work remains. "
                "Accept Release/Released/Cleared or an equally clear release-status equivalent; "
                "a timing-only label such as On schedule is insufficient"
            )
        else:
            required_meaning = (
                "the month is held, deferred, or not fully released because work remains. "
                "Accept Deferred backlog, Deferred work carried, Hold, or an equivalent. "
                "Do not require this monthly status cell to repeat the separately scored "
                "portfolio-level resequencing action"
            )
        grading_boundary = (
            "Judge only whether this monthly row communicates released versus deferred/held. "
            "Portfolio escalation and recovery actions are scored in separate overall-decision fields."
        )
    elif re.fullmatch(r"release_bridge_20\d\d_\d\d_management_release_status", label):
        decision_scope = "monthly earnings-release status"
        if normalized_expected == "capacity cleared":
            required_meaning = (
                "the month is released or cleared. Accept Release, Release supported, "
                "Release cleared, or an equivalent"
            )
        else:
            required_meaning = (
                "the month is held or conditional because deferred work or revenue risk remains. "
                "Accept Hold, Hold - deferred revenue/backlog risk, Conditional - remediation + "
                "deferral, or an equivalent. Do not require the monthly cell to restate the "
                "separately scored portfolio action"
            )
        grading_boundary = (
            "Judge only the monthly release/hold conclusion. The portfolio sequencing action is "
            "scored in release_bridge_management_decision."
        )
    elif re.fullmatch(r"execution_p_\d+_release_status", label):
        decision_scope = "project execution status"
        if normalized_expected == "capacity cleared":
            required_meaning = (
                "the project is complete, released, or otherwise cleared with no deferred work. "
                "Accept Complete, Complete - release, Release/Released, or an equivalent"
            )
        else:
            required_meaning = (
                "the project is deferred or held and is not released. Accept Deferred, Deferred "
                "backlog, Hold - deferred backlog, or an equivalent. Do not require each project "
                "row to repeat the separately scored executive portfolio action"
            )
        grading_boundary = (
            "Judge only whether this project is released/complete versus deferred/held. Executive "
            "portfolio action is scored separately."
        )
    elif label == "execution_portfolio_release_status":
        decision_scope = "portfolio execution status"
        required_meaning = (
            "the portfolio is not a full unconditional release because deferred work remains. "
            "Accept a hold, conditional/partial release with deferred backlog, or an equivalent. "
            "The detailed sequencing action is scored separately"
        )
        grading_boundary = (
            "Judge the overall released-versus-deferred portfolio status, not whether the same "
            "cell repeats the management action or recovery condition."
        )
    elif label == "executive_recovery_required_action":
        decision_scope = "executive recovery action"
        required_meaning = (
            "management must recover capacity for, sequence, or rephase the remaining deferred "
            "priority/unresolved backlog. Accept an action that stages the approved remedies and "
            "then escalates residual unresolved hours for added capacity, customer schedule relief, "
            "or executive sequencing. The action field need not repeat the separately scored release timing"
        )
        grading_boundary = (
            "Judge the recovery action itself. The highest-priority project, recovery amount, release "
            "condition, and hold/release decision are independently scored."
        )
    elif label == "executive_recovery_release_condition":
        decision_scope = "full-portfolio release condition"
        required_meaning = (
            "full release requires both (1) eliminating or resolving the remaining capacity shortfall, "
            "unresolved hours, deferred priority backlog, or the associated deferred revenue/revenue-at-risk "
            "through recovery/sequencing/rephasing and "
            "(2) curing the required recovery amount or gross-commitment gross-profit shortfall. Accept "
            "equivalent causal wording. Merely recovering gross profit while still allowing backlog to "
            "remain deferred is insufficient"
        )
        grading_boundary = (
            "Judge the two-part condition for full portfolio release. Exact amounts and the separate "
            "executive decision are independently scored."
        )
    elif label == "release_bridge_management_decision":
        decision_scope = "portfolio management release decision"
        required_meaning = (
            "hold the gross-commitment/full-portfolio release until the deferred portfolio is sequenced "
            "or rephased. It may separately support release of the probability-weighted plan. Reject a "
            "conditional or partial full-portfolio release that expressly leaves backlog deferred instead "
            "of holding that release for executive sequencing"
        )
        grading_boundary = (
            "Judge the portfolio-level management release decision. Monthly and project statuses, recovery "
            "amounts, action mechanics, and the final recovery condition are independently scored."
        )
    elif label == "executive_recovery_decision":
        decision_scope = "current full-portfolio recovery decision"
        required_meaning = (
            "the gross-commitment/full-portfolio release is currently held or not authorized until "
            "executive recovery, sequencing, or rephasing resolves the deferred/constrained work. "
            "Accept HOLD plus executive sequencing before release, or conditional wording that "
            "unambiguously identifies the unresolved constraint, requires executive sequencing/recovery, "
            "and says that constraint persists until the work is resequenced or recovered. Because the "
            "submitted field is already labeled as the executive recovery decision, wording such as "
            "'Conditional — executive sequencing required; shortfall remains until deferred work is "
            "resequenced' is a valid hold-pending-action equivalent and need not repeat the word release. "
            "Reject an unconditional release and reject a present 'conditional release' that merely lists "
            "future remediation or sequencing without an explicit until/before dependency showing that "
            "the unresolved condition must be cured first"
        )
        grading_boundary = (
            "Judge only the current full-portfolio hold/release decision and its dependency on executive "
            "recovery or sequencing. The recovery action, release-condition mechanics, project, and exact "
            "amounts are independently scored and need not be repeated here."
        )
    else:
        required_meaning = meaning_by_status.get(normalized_expected, str(expected))
    return {
        "output_label": label,
        "decision_scope": decision_scope,
        "canonical_status": expected,
        "required_meaning": required_meaning,
        "grading_boundary": grading_boundary,
        "equivalence_rule": (
            "Accept any ordinary professional wording with the same operational decision and direction. "
            "Do not require the canonical phrase. Reject a status that reverses cleared versus held, "
            "or omits an action that this field itself is specifically responsible for communicating."
        ),
    }


def _task_035_model_or_control_check(workbook, values, criterion_id: str) -> tuple[bool, str]:
    model_requirements = {
        "model_content__burn_curve": ("Revenue Burn", ("probability", "revenue")),
        "model_content__required_hours": ("Revenue Burn", ("required", "hours")),
        "model_content__available_hours": ("Labor Capacity", ("available", "hours")),
        "model_content__overtime": ("Gap Analysis", ("overtime",)),
        "model_content__subcontract": ("Gap Analysis", ("subcontract",)),
        "model_content__capacity_gap": ("Gap Analysis", ("shortfall",)),
    }
    if criterion_id in model_requirements:
        sheet_name, required_terms = model_requirements[criterion_id]
        if sheet_name not in workbook.sheetnames:
            return False, f"required schedule {sheet_name!r} is missing"
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        matches: list[str] = []
        for row in sheet.iter_rows():
            label_text = _normalize(" ".join(str(cell.value or "") for cell in row))
            if not all(term in label_text for term in required_terms):
                continue
            for cell in row:
                cached = value_sheet[cell.coordinate].value
                if isinstance(cell.value, str) and cell.value.startswith("=") and cached is not None:
                    matches.append(f"{sheet_name}!{cell.coordinate}={cell.value}=>{cached!r}")
        return bool(matches), (
            f"formula-driven completed schedule examples={matches[:5]!r}"
            if matches else "the labeled schedule exists only as an empty starter row or lacks formula-driven results"
        )

    checks = {
        "controls__source": "Source population tie",
        "controls__version": "Current-versus-prior version",
        "controls__period": "Period completeness",
        "controls__scenario": "Scenario validity",
        "controls__unit": "Unit conversion",
        "controls__check": "Roll-forward / bridge",
        "controls__model_status": "MODEL STATUS",
        "model_content__model_status": "MODEL STATUS",
    }
    label = checks.get(criterion_id)
    if label is None or "Checks" not in workbook.sheetnames:
        return False, "no task-specific substantive control is defined"
    sheet = workbook["Checks"]
    value_sheet = values["Checks"] if "Checks" in values.sheetnames else sheet
    for row in sheet.iter_rows():
        if not any(
            semantic_equal(cell.value, label) or contains_concept(cell.value, label)
            for cell in row if cell.value is not None
        ):
            continue
        populated: list[str] = []
        formula_results: list[str] = []
        for cell in row[1:]:
            cached = value_sheet[cell.coordinate].value
            if cached not in (None, ""):
                populated.append(f"{cell.coordinate}={cached!r}")
            if isinstance(cell.value, str) and cell.value.startswith("=") and cached is not None:
                formula_results.append(f"{cell.coordinate}={cell.value}=>{cached!r}")
        status = value_sheet.cell(row[0].row, 6).value
        status_ok = any(contains_concept(status, token) for token in ("ok", "pass", "cleared", "complete"))
        met = len(populated) >= 3 and bool(formula_results) and status_ok
        return met, f"populated={populated!r}; formula_results={formula_results!r}; status={status!r}"
    return False, f"required control row {label!r} is missing"


_TASK_068_SLIDES_BY_CRITERION = {
    "preservation__title": (1,),
    "headline_values__guidance_release_status": (2, 8),
    "headline_values__largest_downside_driver": (8,),
    "headline_values__guidance_update_required": (2, 8),
    "narrative__executive_summary": (2,),
    "narrative__revenue": (3,),
    "narrative__ebitda": (4,),
    "narrative__cash": (6,),
    "narrative__backlog": (7,),
    "narrative__outlook": (7,),
    "narrative__risk": (8,),
    "narrative__owner": (8,),
    "narrative__action": (8,),
}


def _task_068_slide_numbers(criterion_id: str) -> tuple[int, ...]:
    if criterion_id in _TASK_068_SLIDES_BY_CRITERION:
        return _TASK_068_SLIDES_BY_CRITERION[criterion_id]
    if criterion_id.startswith(("sources__", "controls__")):
        return (9,)
    return tuple(range(1, 10))


def _task_068_slides_changed(path: Path, artifact_relative: str, slide_numbers: Iterable[int]) -> bool:
    seed_path = SEED_WORKSPACE / artifact_relative
    if not seed_path.is_file():
        return True
    current = Presentation(path)
    seed = Presentation(seed_path)
    for slide_number in slide_numbers:
        if slide_number > len(current.slides) or slide_number > len(seed.slides):
            return False
        current_text = _normalize("\n".join(
            str(shape.text) for shape in current.slides[slide_number - 1].shapes
            if hasattr(shape, "text") and str(shape.text or "").strip()
        ))
        seed_text = _normalize("\n".join(
            str(shape.text) for shape in seed.slides[slide_number - 1].shapes
            if hasattr(shape, "text") and str(shape.text or "").strip()
        ))
        if current_text != seed_text:
            return True
    return False


def _task_068_reference_context(
    criterion_id: str,
    expected_facts: Any,
    answer: dict[str, Any],
) -> dict[str, Any]:
    answer_keys = {
        "headline_values__guidance_release_status": "The current release decision is to maintain guidance with heightened monitoring. Accept ordinary concise equivalents such as `maintain guidance; monitor`, `maintain with monitoring`, `hold guidance and monitor`, or `guidance unchanged with close monitoring`. Do not require the literal word `heightened`. Reject a release, cut, withdrawal, or update conclusion that contradicts this state.",
        "narrative__executive_summary": "Summarize the direction and scale of Q2 revenue and EBITDA performance, cash/liquidity, latest outlook, principal risk, and management response. A concise monitoring or release posture is a sufficient management response here; owner names, detailed actions, and dates are independently scored on slide 8 and need not be repeated on the summary slide. Do not require a scripted sentence. If the summary displays a Q2 free-cash-flow, combined-stress EBITDA, post-mitigation EBITDA, or other verified headline amount, a materially wrong displayed amount makes this criterion unmet; correct directional prose or the correct final maintain-guidance wording cannot rescue that contradiction.",
        "narrative__revenue": "Explain that Construction, Service, and Controls together bridge approved-plan Q2 revenue to Q2 actual, with the three branch variances reconciling to the consolidated favorable variance. Accept normal business-unit and favorable-variance wording.",
        "narrative__ebitda": "Explain that Construction, Service, and Controls together bridge approved-plan adjusted EBITDA to actual, with the three adverse branch-level actual-versus-plan variances reconciling to the consolidated shortfall. Accept normal EBITDA and unfavorable-variance wording. Do not treat a consolidated operational-driver bridge or the larger diagnostic impacts for construction gross profit, service labor, and controls mix as the required business-unit EBITDA reconciliation.",
        "narrative__cash": "Explain the OCF-minus-capex free-cash-flow chain, the shortfall to plan, unrestricted cash/revolver posture, and the documented collections-timing drag. Reject an earnings explanation for the collections-only cash item.",
        "narrative__backlog": "Connect signed backlog to remaining FY26 revenue outlook and its resulting coverage, without treating probability-weighted pipeline as signed backlog.",
        "narrative__outlook": "State the latest FY26 outlook and guidance posture and identify the quantified update trigger/buffer or remaining revenue-decline headroom that would require reconsideration.",
        "narrative__risk": "Identify the signed principal downside and correctly connect combined-stress EBITDA, executable mitigation, and post-mitigation guidance headroom.",
        "narrative__owner": "Assign each material risk or mitigation action shown on the risk-and-action slide to an identifiable management owner; job titles or unambiguous role abbreviations are acceptable.",
        "narrative__action": "State concrete mitigation actions and their next evidence gate or decision date, and make the release conclusion consistent with the verified mitigation bridge. If this slide displays an action pool, conversion, executable mitigation, post-mitigation EBITDA, or headroom that materially contradicts the verified values, this criterion is unmet even when the action labels, owners, dates, or final maintain-guidance words are otherwise present.",
        "controls__numerical_tie_out": "Show a substantive control for each of these five subjects on the final slide: Q2 revenue to posted YTD revenue, adjusted EBITDA, cash/free cash flow, signed backlog, and the latest FY26 outlook. Each subject needs a displayed value or an unambiguous subject-specific tie row whose referenced value is also clear on the slide. A generic or incomplete list of TIES/OK labels is insufficient, and a stress-EBITDA control does not replace the latest-outlook control.",
        "controls__revenue_bridge": "Show that the three business-unit revenue amounts and variances reconcile approved-plan Q2 revenue to Q2 actual with no unexplained difference.",
        "controls__ebitda_bridge": "Show that the three business-unit adjusted-EBITDA amounts and variances reconcile approved-plan EBITDA to Q2 actual with no unexplained difference.",
        "controls__cash_roll_forward": "Show operating cash flow less capital expenditure equals free cash flow and relate the result to plan and liquidity. Equivalent cash-conversion or liquidity-tie wording is acceptable.",
        "controls__latest_forecast": "Identify the June reforecast as the latest approved FY26 outlook authority and tie the latest forecast to the deck or executive summary. The $50.8 million outlook is independently checked elsewhere and need not be repeated on this control slide when a clear `latest forecast ... TIES/OK` row and the latest-approved June-reforecast source are both present. An AOP or April working forecast is not sufficient, and naming the June reforecast only in a source list without a latest-forecast tie/control is insufficient.",
        "sources__contractor_accounting_mcp": "Identify Vista ERP or company accounting records as the authority for posted YTD revenue. The acronym MCP is not required.",
        "sources__controller_tied": "Identify `Q2 management reporting data book - v7 controller tie.xlsx` or an unambiguous v7 controller-tied reporting-book reference; reject the v5 CFO scratch book as controlling.",
        "sources__management_correspondence": "Identify the July 4 Q2 board/lender/IR refresh thread, the July 5 board-review/source-tie follow-up, or the Q2 board/lender narrative review notes as management authority for the outlook, risks, or actions.",
        "sources__policy": "Identify the signed KPI definitions/lender-presentation policy or signed FY26 non-GAAP policy for the definitions or reporting treatment it governs.",
    }
    branch_revenue_keys = tuple(
        f"{branch}_q2_{suffix}"
        for branch in ("construction", "service", "controls")
        for suffix in (
            "revenue", "approved_plan_revenue", "revenue_variance_to_plan"
        )
    )
    branch_ebitda_keys = tuple(
        f"{branch}_q2_{suffix}"
        for branch in ("construction", "service", "controls")
        for suffix in (
            "adjusted_ebitda", "approved_plan_adjusted_ebitda",
            "adjusted_ebitda_variance_to_plan",
        )
    )
    context_keys_by_criterion = {
        "narrative__executive_summary": (
            "q2_revenue", "q2_revenue_variance_to_plan", "q2_adjusted_ebitda",
            "q2_adjusted_ebitda_variance_to_plan", "q2_free_cash_flow",
            "q2_unrestricted_cash", "latest_full_year_revenue_outlook",
            "combined_stress_ebitda", "post_mitigation_combined_stress_ebitda",
            "guidance_release_status",
        ),
        "narrative__revenue": (
            "posted_ytd_revenue", "q2_posted_ytd_revenue_share",
            "q2_posted_ytd_revenue_reconciliation_check", "q2_revenue",
            "q2_approved_plan_revenue", "q2_revenue_variance_to_plan",
            *branch_revenue_keys,
        ),
        "narrative__ebitda": (
            "q2_adjusted_ebitda", "q2_approved_plan_adjusted_ebitda",
            "q2_adjusted_ebitda_variance_to_plan", *branch_ebitda_keys,
        ),
        "narrative__cash": (
            "q2_operating_cash_flow", "q2_capital_expenditure",
            "q2_free_cash_flow", "q2_free_cash_flow_plan",
            "q2_free_cash_flow_variance_to_plan", "q2_unrestricted_cash",
            "collections_timing_cash_impact", "maximum_revolver",
        ),
        "narrative__backlog": (
            "signed_backlog", "remaining_fy26_revenue_outlook",
            "signed_backlog_coverage_of_remaining_outlook",
            "latest_full_year_revenue_outlook",
        ),
        "narrative__outlook": (
            "latest_full_year_revenue_outlook", "revenue_guidance_low",
            "revenue_guidance_high", "ebitda_guidance_low",
            "ebitda_guidance_high", "minimum_guidance_update_headroom",
            "additional_revenue_decline_to_update_trigger",
            "guidance_update_required", "guidance_release_status",
        ),
        "narrative__risk": (
            "largest_downside_driver", "combined_stress_ebitda",
            "ebitda_shortfall_to_guidance_low", "identified_ebitda_action_pool",
            "approved_mitigation_conversion", "executable_ebitda_mitigation",
            "post_mitigation_combined_stress_ebitda",
            "post_mitigation_headroom_to_guidance_low",
        ),
        "narrative__owner": (),
        "narrative__action": (
            "identified_ebitda_action_pool", "approved_mitigation_conversion",
            "executable_ebitda_mitigation", "combined_stress_ebitda",
            "post_mitigation_combined_stress_ebitda",
            "post_mitigation_headroom_to_guidance_low",
            "minimum_guidance_update_headroom", "guidance_release_status",
        ),
        "controls__numerical_tie_out": (
            "posted_ytd_revenue", "q2_posted_ytd_revenue_share",
            "q2_posted_ytd_revenue_reconciliation_check", "q2_revenue",
            "q2_adjusted_ebitda", "q2_free_cash_flow", "signed_backlog",
            "latest_full_year_revenue_outlook",
        ),
        "controls__revenue_bridge": (
            "q2_revenue", "q2_approved_plan_revenue",
            "q2_revenue_variance_to_plan", *branch_revenue_keys,
        ),
        "controls__ebitda_bridge": (
            "q2_adjusted_ebitda", "q2_approved_plan_adjusted_ebitda",
            "q2_adjusted_ebitda_variance_to_plan", *branch_ebitda_keys,
        ),
        "controls__cash_roll_forward": (
            "q2_operating_cash_flow", "q2_capital_expenditure",
            "q2_free_cash_flow", "q2_free_cash_flow_plan",
            "q2_free_cash_flow_variance_to_plan", "q2_unrestricted_cash",
        ),
        "controls__latest_forecast": (
            "latest_full_year_revenue_outlook", "revenue_guidance_low",
            "revenue_guidance_high", "ebitda_guidance_low",
            "ebitda_guidance_high",
        ),
    }
    context_keys = context_keys_by_criterion.get(criterion_id, ())
    return {
        "criterion_answer_key": answer_keys.get(criterion_id, expected_facts),
        "expected_facts": expected_facts,
        "equivalence_rule": "Accept ordinary professional wording and equivalent labels; do not require the authored phrase.",
        "verified_finance_context": {
            key: answer[key]
            for key in context_keys
            if key in answer and answer[key] is not None
        },
    }


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
            if task_id == "task_035":
                label, _expected = next(iter(spec["label_values"].items()))
                _evidence, populated, gate_evidence = _scoped_xlsx_label_evidence(
                    workbook,
                    values,
                    label.replace("_", " "),
                    aliases=_TASK_035_LABEL_ALIASES.get(label, ()),
                )
                return populated, gate_evidence
            # Deterministic hard gates protect checkable finance facts.  Text,
            # booleans, dates, and lists are intentionally routed to the
            # independent semantic reviewer so ordinary professional wording
            # is not forced to reproduce the authored answer literally.
            checkable = {
                label: expected
                for label, expected in spec["label_values"].items()
                if isinstance(expected, (int, float)) and not isinstance(expected, bool)
            }
            missing = [
                label
                for label, expected in checkable.items()
                if not (
                    _task_082_label_value_present(
                        workbook, values, label, expected
                    )
                    if task_id == "task_082"
                    else _workbook_value_present(values, expected)
                )
            ]
            if missing:
                return False, f"required numeric facts missing anywhere in workbook={missing!r}"
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
                    "required numeric facts are present and no exact-label contradiction exists"
                    if not conflicts
                    else "; ".join(conflicts)
                ),
            )
        if kind == "artifact_tokens":
            if task_id == "task_035":
                token = spec.get("tokens", [""])[0]
                source_reference = _TASK_035_SOURCE_REFERENCES.get(token, {})
                _evidence, completed, gate_evidence = _scoped_xlsx_token_evidence(
                    workbook,
                    values,
                    token,
                    aliases=source_reference.get("aliases", ()),
                )
                return completed, gate_evidence
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
        if task_id == "task_068":
            slide_numbers = _task_068_slide_numbers(spec["id"])
            _evidence, substantive, evidence = _pptx_slide_evidence(path, slide_numbers)
            changed = _task_068_slides_changed(path, gold["artifact"]["path"], slide_numbers)
            if spec["id"] == "preservation__title":
                return substantive, evidence
            return substantive and changed, f"{evidence}; relevant slide content changed from starter={changed}"
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
        if task_id == "task_035":
            if spec["kind"] == "xlsx_label_values":
                label, expected = next(iter(spec["label_values"].items()))
                submitted_evidence, _populated, _gate = _scoped_xlsx_label_evidence(
                    workbook,
                    values,
                    label.replace("_", " "),
                    aliases=_TASK_035_LABEL_ALIASES.get(label, ()),
                )
                reference_context = _task_035_status_reference(label, expected)
                evidence_scope = f"exact workbook output row labeled {label!r}"
                requirement = (
                    f"Evaluate only whether the submitted value for {label!r} expresses "
                    "the correct operational decision under the supplied answer key. Accept "
                    "an unambiguous professional equivalent. Apply the answer key's explicit "
                    "grading boundary: do not demand that a monthly or project status repeat an "
                    "executive action scored in another criterion, but reject an opposite or "
                    "ambiguous released-versus-held decision."
                )
            else:
                token = spec.get("tokens", [""])[0]
                source_reference = _TASK_035_SOURCE_REFERENCES.get(token, {})
                submitted_evidence, _populated, _gate = _scoped_xlsx_token_evidence(
                    workbook,
                    values,
                    token,
                    aliases=source_reference.get("aliases", ()),
                )
                reference_context = {
                    "required_source_or_control": token,
                    "criterion_answer_key": source_reference.get(
                        "answer_key", spec["description"]
                    ),
                    "accepted_authority_references": list(
                        source_reference.get("aliases", ())
                    ),
                    "equivalence_rule": "Accept a clearly identified equivalent source/control; do not require the authored phrase.",
                }
                evidence_scope = f"workbook rows containing the {token!r} source/control concept"
                requirement = (
                    "Evaluate only whether this submitted source row identifies an authoritative "
                    f"source satisfying the supplied answer key for {token!r}. Accept a clear "
                    "filename, subject, abbreviation, or ordinary professional equivalent; reject "
                    "an unidentified generic claim or the wrong/superseded authority."
                )
            task_context = {
                "assignment": "Complete the FY27 signed-backlog burn, capacity, remediation, execution, earnings-release, and executive-recovery model.",
                "sign_convention": "capacity gap = required hours minus available hours; positive means shortage, negative means spare capacity; unresolved shortfall is floored at zero",
                "status_conventions": {
                    "cleared": "no unresolved shortfall remains after permitted remedies",
                    "sequencing_or_resequencing_required": "a residual shortfall or deferred portfolio remains",
                    "executive_decision_required": "the project or portfolio cannot be released without management action",
                },
                "grading_boundary": "All checkable amounts, months, project IDs, formulas, and calculation chains are graded deterministically. This judge grades only the supplied status/source row.",
            }
        elif task_id == "task_068":
            slide_numbers = _task_068_slide_numbers(spec["id"])
            submitted_evidence, _substantive, _gate = _pptx_slide_evidence(path, slide_numbers)
            reference_context = _task_068_reference_context(
                spec["id"], expected_facts, gold["answer"]
            )
            evidence_scope = "slide(s) " + ", ".join(str(number) for number in slide_numbers)
            requirement = (
                f"Evaluate only this criterion in the supplied slide content: {spec['description']} "
                "Use the criterion answer key and verified finance context below. Accept normal "
                "professional wording; reject a missing, contradictory, or unsupported conclusion."
            )
            task_context = {
                "assignment": "Complete the existing nine-slide June executive performance review for leadership.",
                "central_work": "Tie the accounting/reporting figures and build the mitigation-to-guidance release decision.",
                "grading_boundary": "Every checkable amount and calculation is graded separately and deterministically. This judge grades only the supplied slide content for the named narrative, source, control, or decision criterion.",
            }
        else:
            submitted_evidence = ""
            reference_context = expected_facts or {}
            evidence_scope = "legacy whole artifact"
            task_context = {}
        reviews.append(
            {
                "criterion_id": spec["id"],
                "requirement": requirement,
                "hard_gate_met": hard_gate_met,
                "hard_gate_evidence": hard_gate_evidence,
                "legacy_lexical_match": bool(by_id[spec["id"]].met),
                **(
                    {
                        "submitted_evidence": submitted_evidence,
                        "reference_context": reference_context,
                        "task_context": task_context,
                        "evidence_scope": evidence_scope,
                        "always_judge": True,
                    }
                    if task_id in {"task_035", "task_068"}
                    else {}
                ),
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
        "execution_mode": (
            "scoped_per_criterion"
            if task_id in {"task_035", "task_068"}
            else "legacy_batched"
        ),
        "criteria": reviews,
        "policy": (
            "A criterion passes only when its deterministic hard gate passes and the semantic judge "
            "finds the professional-language association or narrative substance MET."
        ),
    }


def _artifact_label_value(
    entries: list[str],
    label: str,
    expected: Any,
    *,
    directional_strings: bool = False,
) -> bool:
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
            matcher = (
                _sample_directional_semantic_value_matches
                if directional_strings
                else semantic_value_matches
            )
            if any(
                date_matches(candidate, expected) or matcher(candidate, expected)
                for candidate in candidates
            ):
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
    (?:[-−]?\$?[ \t]*\(?|\([ \t]*\$?[-−]?[ \t]*)
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
    normalized_literal = literal.replace("−", "-")
    candidates = list(iter_numeric_candidates(normalized_literal))
    accounting_negative = "(" in literal and ")" in literal
    explicit_negative = "-" in normalized_literal.partition(
        next((c for c in normalized_literal if c.isdigit()), "")
    )[0]
    percent = literal.rstrip().rstrip(")").rstrip().endswith("%")
    cleaned_for_precision = (
        normalized_literal.strip().replace("$", "").replace(",", "")
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
    cleaned = normalized_literal.strip().replace("$", "").replace(",", "").strip("() ")
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


def _presentation_contexts(path: Path) -> list[str]:
    """Return slide, table-row, and table-intersection contexts."""
    presentation = Presentation(path)
    contexts: list[str] = []
    for slide in presentation.slides:
        slide_parts = [
            shape.text.strip() for shape in slide.shapes
            if hasattr(shape, "text") and shape.text.strip()
        ]
        slide_text = " | ".join(slide_parts)
        contexts.append(slide_text)
        for shape in slide.shapes:
            if not getattr(shape, "has_table", False):
                continue
            rows = [[cell.text.strip() for cell in row.cells] for row in shape.table.rows]
            if not rows:
                continue
            headers = rows[0]
            for row in rows[1:]:
                contexts.append(" | ".join(row))
                for column, cell in enumerate(row):
                    header = headers[column] if column < len(headers) else ""
                    contexts.append(" | ".join((row[0] if row else "", header, cell)))
    return contexts


def _context_has_alias(context: str, alias: str) -> bool:
    if contains_concept(context, alias):
        return True
    context_tokens = set(_normalize(context).split())
    alias_tokens = [token for token in _normalize(alias).split() if len(token) > 1]
    return bool(alias_tokens) and all(token in context_tokens for token in alias_tokens)


def _task_068_value_slides(label: str) -> tuple[int, ...]:
    if label.startswith(("construction_q2_", "service_q2_", "controls_q2_")):
        return (5,)
    if label == "june_revenue":
        return (3,)
    if label == "june_ebitda":
        return (4,)
    if label in {
        "posted_ytd_revenue", "q2_posted_ytd_revenue_share",
        "q2_posted_ytd_revenue_reconciliation_check", "q2_revenue",
        "q2_approved_plan_revenue", "q2_revenue_variance_to_plan",
        "q2_organic_growth",
    }:
        return (2, 3, 5, 9)
    if label in {
        "q2_adjusted_ebitda",
        "q2_approved_plan_adjusted_ebitda",
        "q2_adjusted_ebitda_variance_to_plan",
    }:
        return (2, 4, 5, 9)
    if label in {
        "q2_operating_cash_flow", "q2_capital_expenditure", "q2_free_cash_flow",
        "q2_free_cash_flow_plan", "q2_free_cash_flow_variance_to_plan",
        "collections_timing_cash_impact", "q2_unrestricted_cash", "maximum_revolver",
    }:
        return (2, 6, 9)
    if label in {
        "signed_backlog", "remaining_fy26_revenue_outlook",
        "signed_backlog_coverage_of_remaining_outlook",
        "latest_full_year_revenue_outlook",
    }:
        return (2, 7, 9)
    if label in {
        "ltm_adjusted_ebitda", "funded_debt", "lender_leverage",
        "fixed_charge_coverage", "revenue_guidance_low", "revenue_guidance_high",
        "ebitda_guidance_low", "ebitda_guidance_high",
    }:
        return (2, 7, 8, 9)
    return (2, 4, 8, 9)


def _task_068_local_contexts(path: Path, slide_numbers: Iterable[int]) -> list[str]:
    """Return local shape/row contexts, never an indiscriminate whole-deck dump."""

    presentation = Presentation(path)
    contexts: list[str] = []
    for slide_number in slide_numbers:
        if slide_number < 1 or slide_number > len(presentation.slides):
            continue
        slide = presentation.slides[slide_number - 1]
        text_shapes = [
            shape for shape in slide.shapes
            if hasattr(shape, "text") and str(shape.text or "").strip()
        ]
        ordered = sorted(text_shapes, key=lambda shape: (int(shape.top), int(shape.left)))
        texts = [str(shape.text).strip() for shape in ordered]
        contexts.extend(texts)
        # Consecutive shape windows cover normal label/value and short formula
        # blocks without letting an unrelated figure elsewhere on the slide
        # satisfy the criterion.
        for start in range(len(texts)):
            for width in range(2, 7):
                window = texts[start : start + width]
                if len(window) == width:
                    contexts.append(" | ".join(window))

        # PowerPoint scorecards are often separate text boxes rather than real
        # tables. Join same-row boxes and, for each value box, its nearest row
        # label and column heading so branch/metric association remains local.
        row_tolerance = 180_000
        column_tolerance = 300_000
        for shape in text_shapes:
            same_row = sorted(
                (
                    candidate for candidate in text_shapes
                    if abs(int(candidate.top) - int(shape.top)) <= row_tolerance
                ),
                key=lambda candidate: int(candidate.left),
            )
            if len(same_row) >= 2:
                contexts.append(" | ".join(str(candidate.text).strip() for candidate in same_row))
            left_candidates = [
                candidate for candidate in same_row
                if int(candidate.left) < int(shape.left)
            ]
            above_candidates = [
                candidate for candidate in text_shapes
                if int(candidate.top) < int(shape.top)
                and abs(int(candidate.left) - int(shape.left)) <= column_tolerance
            ]
            if left_candidates and above_candidates:
                row_label = min(left_candidates, key=lambda candidate: int(candidate.left))
                header = max(above_candidates, key=lambda candidate: int(candidate.top))
                contexts.append(
                    " | ".join(
                        (str(row_label.text).strip(), str(header.text).strip(), str(shape.text).strip())
                    )
                )

        for shape in slide.shapes:
            if not getattr(shape, "has_table", False):
                continue
            rows = [[cell.text.strip() for cell in row.cells] for row in shape.table.rows]
            if not rows:
                continue
            headers = rows[0]
            contexts.extend(" | ".join(row) for row in rows)
            for row in rows[1:]:
                for column, cell in enumerate(row):
                    header = headers[column] if column < len(headers) else ""
                    contexts.append(" | ".join((row[0] if row else "", header, cell)))
    return list(dict.fromkeys(context for context in contexts if context.strip()))


_TASK_068_LABEL_ALIASES = {
    "june_revenue": ("June revenue", "June actual revenue", "June actual"),
    "june_ebitda": ("June EBITDA", "June adjusted EBITDA", "June actual"),
    "posted_ytd_revenue": ("posted YTD revenue", "Vista YTD revenue"),
    "q2_posted_ytd_revenue_share": ("Q2 share", "Q2 revenue share", "53% Q2 share"),
    "q2_posted_ytd_revenue_reconciliation_check": (
        "Q2 revenue reconciliation check", "posted YTD revenue tie", "no plug",
    ),
    "q2_revenue": ("Q2 revenue", "Q2 actual revenue"),
    "q2_approved_plan_revenue": (
        "Q2 revenue plan", "approved plan revenue", "revenue plan", "plan",
    ),
    "q2_revenue_variance_to_plan": ("Q2 revenue vs plan", "revenue variance to plan"),
    "q2_organic_growth": (
        "organic growth", "organic YoY", "organic year over year", "IR organic",
    ),
    "q2_adjusted_ebitda": ("Q2 adjusted EBITDA", "Q2 EBITDA actual"),
    "q2_approved_plan_adjusted_ebitda": (
        "Q2 adjusted EBITDA plan", "approved plan EBITDA", "EBITDA plan", "plan",
    ),
    "q2_adjusted_ebitda_variance_to_plan": (
        "Q2 EBITDA vs plan", "adjusted EBITDA variance to plan", "EBITDA variance",
    ),
    "q2_operating_cash_flow": ("Q2 operating cash flow", "operating cash flow", "OCF"),
    "q2_capital_expenditure": ("Q2 capital expenditure", "capital expenditure", "capex"),
    "q2_free_cash_flow": ("Q2 free cash flow", "free cash flow", "FCF"),
    "q2_free_cash_flow_plan": ("Q2 free cash flow plan", "FCF plan", "plan FCF"),
    "q2_free_cash_flow_variance_to_plan": ("FCF vs plan", "free cash flow variance"),
    "collections_timing_cash_impact": ("collections timing", "collections cash impact"),
    "q2_unrestricted_cash": ("Q2 unrestricted cash", "unrestricted cash"),
    "maximum_revolver": ("maximum revolver", "max revolver", "downside revolver"),
    "signed_backlog": ("signed backlog",),
    "remaining_fy26_revenue_outlook": (
        "remaining FY26 revenue", "remaining revenue outlook",
        "H2 remaining revenue", "H2 remaining revenue need",
    ),
    "signed_backlog_coverage_of_remaining_outlook": (
        "backlog coverage of remaining outlook", "FY26 remaining coverage", "backlog coverage",
    ),
    "latest_full_year_revenue_outlook": ("FY26 revenue outlook", "June reforecast"),
    "ltm_adjusted_ebitda": ("LTM adjusted EBITDA", "lender EBITDA"),
    "funded_debt": ("funded debt",),
    "lender_leverage": ("lender leverage", "gross leverage", "leverage"),
    "fixed_charge_coverage": ("fixed charge coverage", "FCCR"),
    "revenue_guidance_low": ("revenue guidance low", "revenue guidance"),
    "revenue_guidance_high": ("revenue guidance high", "revenue guidance"),
    "ebitda_guidance_low": ("EBITDA guidance low", "guidance low"),
    "ebitda_guidance_high": ("EBITDA guidance high", "EBITDA guidance"),
    "construction_gross_profit_impact": ("construction gross profit", "construction GP"),
    "service_labor_productivity_impact": ("service labor productivity", "service labor"),
    "controls_mix_impact": ("controls mix",),
    "identified_ebitda_action_pool": ("identified EBITDA action pool", "action pool"),
    "approved_mitigation_conversion": ("mitigation conversion", "conversion factor"),
    "executable_ebitda_mitigation": (
        "executable mitigation", "converted mitigation",
        "action pool mitigation conversion",
    ),
    "combined_stress_ebitda": ("combined stress EBITDA", "pre-mitigation stress EBITDA"),
    "ebitda_shortfall_to_guidance_low": ("shortfall to guidance low", "guidance shortfall"),
    "post_mitigation_combined_stress_ebitda": ("post-mitigation stress EBITDA",),
    "post_mitigation_headroom_to_guidance_low": (
        "post-mitigation headroom", "headroom to guidance low",
    ),
    "minimum_guidance_update_headroom": ("minimum update buffer", "minimum headroom"),
    "guidance_update_trigger": ("guidance update trigger", "update trigger"),
    "additional_revenue_decline_to_update_trigger": (
        "revenue decline headroom", "additional revenue decline", "revenue headroom",
    ),
    "largest_downside_driver": ("largest downside driver", "copper escalation"),
    "guidance_release_status": ("guidance release status", "guidance"),
    "guidance_update_required": ("guidance update", "formal update"),
}


def _task_068_numeric_literal_matches(
    literal: str,
    expected: float,
    *,
    context: str,
) -> bool:
    """Match a board-deck number without inventing an undisclosed unit.

    The generic investment-committee matcher accepts an unqualified decimal as
    millions because those tables use a single ``$mm`` heading.  Task068 also
    contains ratios such as the 0.62 mitigation-conversion factor.  Treating
    that ratio as $0.62 million can falsely satisfy an unrelated EBITDA value.
    Here an unqualified compact number receives million scaling only when its
    *local* context explicitly declares a dollar-in-millions display unit.
    Explicit ``k``/``M`` suffixes and full-dollar figures retain the normal
    professional board-rounding tolerance.
    """

    normalized_literal = literal.replace("−", "-").strip()
    cleaned = (
        normalized_literal.replace("$", "").replace(",", "")
        .replace("(", "").replace(")", "").strip()
    )
    explicit_unit = bool(re.search(
        r"(?:%|[x×]|thousand|k|million|mm|m|billion|bn|b)\s*$",
        cleaned,
        flags=re.I,
    ))
    numeric_text = re.sub(
        r"(?:%|[x×]|thousand|k|million|mm|m|billion|bn|b)\s*$",
        "",
        cleaned,
        flags=re.I,
    ).strip()
    try:
        displayed = float(numeric_text)
    except ValueError:
        return False

    if explicit_unit or abs(displayed) >= 10_000 or abs(expected) <= 10_000:
        return _task_092_numeric_literal_matches(literal, expected)

    dollars_in_millions = bool(re.search(
        r"(?:\$\s*(?:in\s+)?(?:millions?|mm|m)\b"
        r"|\busd\s*(?:in\s+)?(?:millions?|mm|m)\b"
        r"|\bdollars?\s+in\s+millions?\b)",
        context,
        flags=re.I,
    ))
    if not dollars_in_millions:
        return False

    accounting_negative = "(" in literal and ")" in literal
    explicit_negative = normalized_literal.lstrip().startswith("-")
    if accounting_negative or explicit_negative:
        displayed = -abs(displayed)
    scaled = displayed * 1_000_000.0
    decimals = len(numeric_text.rsplit(".", 1)[1]) if "." in numeric_text else 0
    rounding_tolerance = 0.5 * (10 ** -decimals) * 1_000_000.0
    return _close(
        scaled,
        expected,
        abs_tol=max(_display_tolerance(expected), rounding_tolerance + 1e-12),
        rel_tol=0.0,
    )


def _task_068_branch_artifact_value(path: Path, label: str, expected: float) -> bool:
    """Read a business-unit scorecard without confusing actual, plan, and variance.

    Board decks commonly implement scorecards as aligned text boxes rather
    than native PowerPoint tables.  A branch name alone is not enough to bind
    every number in that row to every branch criterion.  This parser first
    accepts an explicit branch/metric/value statement, then uses the column
    heading and row alignment for grid-style scorecards.
    """

    match = re.fullmatch(
        r"(construction|service|controls)_q2_(.+)", label
    )
    if not match:
        return False
    branch, metric = match.groups()
    branch_aliases = {
        "construction": ("Construction",),
        "service": ("Service",),
        "controls": ("Controls", "Building Controls"),
    }[branch]
    metric_aliases = {
        "revenue": ("Q2 revenue", f"{branch} revenue"),
        "approved_plan_revenue": (
            "approved plan revenue", "plan revenue", f"{branch} plan revenue",
        ),
        "revenue_variance_to_plan": (
            "revenue vs plan", "revenue variance to plan",
            f"{branch} revenue vs plan",
        ),
        "adjusted_ebitda": (
            "Q2 adjusted EBITDA", f"{branch} adjusted EBITDA", f"{branch} EBITDA",
        ),
        "approved_plan_adjusted_ebitda": (
            "approved plan adjusted EBITDA", "plan adjusted EBITDA", "plan EBITDA",
            f"{branch} plan EBITDA",
        ),
        "adjusted_ebitda_variance_to_plan": (
            "adjusted EBITDA vs plan", "EBITDA variance to plan",
            f"{branch} EBITDA vs plan",
        ),
    }[metric]
    contexts = _task_068_local_contexts(path, (5,))
    for context in contexts:
        if not any(_context_has_alias(context, alias) for alias in branch_aliases):
            continue
        if not any(_context_has_alias(context, alias) for alias in metric_aliases):
            continue
        if any(
            _task_068_numeric_literal_matches(literal, expected, context=context)
            for literal in _TASK_092_NUMERIC_LITERAL_PATTERN.findall(context)
        ):
            return True

    presentation = Presentation(path)
    if len(presentation.slides) < 5:
        return False
    shapes = [
        shape for shape in presentation.slides[4].shapes
        if hasattr(shape, "text") and str(shape.text or "").strip()
    ]
    branch_shapes = [
        shape for shape in shapes
        if any(_context_has_alias(str(shape.text), alias) for alias in branch_aliases)
    ]

    def is_header(text: str) -> bool:
        normalized = _normalize(text)
        has_revenue = "revenue" in normalized and "ebitda" not in normalized
        has_ebitda = "ebitda" in normalized
        has_plan = "plan" in normalized
        has_variance = "variance" in normalized or "vs plan" in normalized
        if metric == "revenue":
            return has_revenue and not has_plan and not has_variance
        if metric == "approved_plan_revenue":
            return has_revenue and has_plan and not has_variance
        if metric == "revenue_variance_to_plan":
            return (has_revenue and has_variance) or normalized == "vs plan"
        if metric == "adjusted_ebitda":
            return has_ebitda and not has_plan and not has_variance and "%" not in text
        if metric == "approved_plan_adjusted_ebitda":
            return has_ebitda and has_plan and not has_variance
        return has_ebitda and has_variance

    header_shapes = [shape for shape in shapes if is_header(str(shape.text))]
    row_tolerance = 180_000
    for branch_shape in branch_shapes:
        row_shapes = [
            shape for shape in shapes
            if abs(int(shape.top) - int(branch_shape.top)) <= row_tolerance
        ]
        for header in header_shapes:
            header_center = int(header.left) + int(header.width) // 2
            inherited_unit_headers: list[str] = []
            if metric == "revenue_variance_to_plan":
                same_header_row = [
                    shape for shape in shapes
                    if abs(int(shape.top) - int(header.top)) <= row_tolerance
                    and int(shape.left) < int(header.left)
                    and "revenue" in _normalize(str(shape.text))
                    and "ebitda" not in _normalize(str(shape.text))
                ]
                if same_header_row:
                    inherited_unit_headers.append(str(max(
                        same_header_row, key=lambda shape: int(shape.left)
                    ).text).strip())
            candidates = sorted(
                (
                    shape for shape in row_shapes
                    if shape is not branch_shape
                    and abs(
                        int(shape.left) + int(shape.width) // 2 - header_center
                    ) <= 300_000
                ),
                key=lambda shape: abs(
                    int(shape.left) + int(shape.width) // 2 - header_center
                ),
            )
            for value_shape in candidates[:1]:
                context = " | ".join((
                    str(branch_shape.text).strip(),
                    *inherited_unit_headers,
                    str(header.text).strip(),
                    str(value_shape.text).strip(),
                ))
                if any(
                    _task_068_numeric_literal_matches(
                        literal, expected, context=context
                    )
                    for literal in _TASK_092_NUMERIC_LITERAL_PATTERN.findall(context)
                ):
                    return True
    return False


def _task_068_bridge_artifact_value(path: Path, label: str, expected: float) -> bool:
    """Bind a Plan/Actual waterfall value to its local Q2 bridge and unit."""

    mapping = {
        "q2_approved_plan_revenue": (3, "revenue", "plan"),
        "q2_revenue": (3, "revenue", "actual"),
        "q2_approved_plan_adjusted_ebitda": (4, "ebitda", "plan"),
        "q2_adjusted_ebitda": (4, "ebitda", "actual"),
    }
    if label not in mapping:
        return False
    slide_number, metric, state = mapping[label]
    presentation = Presentation(path)
    if len(presentation.slides) < slide_number:
        return False
    shapes = [
        shape for shape in presentation.slides[slide_number - 1].shapes
        if hasattr(shape, "text") and str(shape.text or "").strip()
    ]
    headings = [
        shape for shape in shapes
        if metric in _normalize(str(shape.text))
        and "bridge" in _normalize(str(shape.text))
    ]
    labels = [
        shape for shape in shapes
        if _normalize(str(shape.text)) == state
    ]
    for state_shape in labels:
        state_center = int(state_shape.left) + int(state_shape.width) // 2
        aligned = sorted(
            (
                shape for shape in shapes
                if shape is not state_shape
                and abs(
                    int(shape.left) + int(shape.width) // 2 - state_center
                ) <= 300_000
                and int(shape.top) < int(state_shape.top)
            ),
            key=lambda shape: int(state_shape.top) - int(shape.top),
        )
        for value_shape in aligned[:2]:
            for heading in headings:
                context = " | ".join((
                    str(heading.text).strip(),
                    str(state_shape.text).strip(),
                    str(value_shape.text).strip(),
                ))
                if any(
                    _task_068_numeric_literal_matches(
                        literal, expected, context=context
                    )
                    for literal in _TASK_092_NUMERIC_LITERAL_PATTERN.findall(context)
                ):
                    return True
    return False


def _task_068_artifact_value(path: Path, label: str, expected: Any) -> bool:
    if re.fullmatch(
        r"(?:construction|service|controls)_q2_.+", label
    ) and not isinstance(expected, (bool, str)):
        return _task_068_branch_artifact_value(path, label, float(expected))
    if not isinstance(expected, (bool, str)) and _task_068_bridge_artifact_value(
        path, label, float(expected)
    ):
        return True

    aliases = [label.replace("_", " "), *_TASK_068_LABEL_ALIASES.get(label, ())]
    contexts = _task_068_local_contexts(path, _task_068_value_slides(label))
    if isinstance(expected, bool):
        if label == "guidance_update_required" and expected is False:
            text = _normalize("\n".join(contexts))
            maintained = any(phrase in text for phrase in (
                "no guidance update", "maintain guidance", "hold guidance",
                "guidance unchanged", "do not revise guidance",
            ))
            contrary = any(phrase in text for phrase in (
                "guidance update required", "update guidance now",
                "revise guidance now", "withdraw guidance now",
                "reduce guidance now", "cut guidance now",
            ))
            return maintained and not contrary
        return any(
            any(_context_has_alias(context, alias) for alias in aliases)
            and _boolean_matches(context, expected)
            for context in contexts
        )
    if isinstance(expected, str):
        return any(
            any(_context_has_alias(context, alias) for alias in aliases)
            and _sample_directional_semantic_value_matches(context, expected)
            for context in contexts
        )
    if label == "q2_posted_ytd_revenue_reconciliation_check" and float(expected) == 0.0:
        for context in contexts:
            if not any(_context_has_alias(context, alias) for alias in aliases):
                continue
            normalized = _normalize(context)
            if any(phrase in normalized for phrase in ("no plug", "zero difference", "reconciles", "ties")):
                return True

    if label in {
        "revenue_guidance_low", "revenue_guidance_high",
        "ebitda_guidance_low", "ebitda_guidance_high",
    }:
        range_pattern = re.compile(
            r"\$?\s*(\d[\d,]*(?:\.\d+)?)\s*[-–—]\s*\$?\s*"
            r"(\d[\d,]*(?:\.\d+)?)\s*(k|m|mm|million|b|bn|billion)\b",
            flags=re.I,
        )
        scale = {
            "k": 1_000.0, "m": 1_000_000.0, "mm": 1_000_000.0,
            "million": 1_000_000.0, "b": 1_000_000_000.0,
            "bn": 1_000_000_000.0, "billion": 1_000_000_000.0,
        }
        for context in contexts:
            if not any(_context_has_alias(context, alias) for alias in aliases):
                continue
            for match in range_pattern.finditer(context):
                multiplier = scale[match.group(3).casefold()]
                values = (
                    float(match.group(1).replace(",", "")) * multiplier,
                    float(match.group(2).replace(",", "")) * multiplier,
                )
                if any(
                    _close(value, float(expected), abs_tol=_display_tolerance(float(expected)), rel_tol=0.0)
                    for value in values
                ):
                    return True

    if label == "minimum_guidance_update_headroom":
        for context in contexts:
            normalized = _normalize(context)
            if "guidance" not in normalized or "trigger" not in normalized:
                continue
            literals = _TASK_092_NUMERIC_LITERAL_PATTERN.findall(context)
            values: list[float] = []
            for literal in literals:
                cleaned = literal.replace("−", "-").strip()
                negative = cleaned.startswith("-") or (
                    cleaned.startswith("(") and cleaned.endswith(")")
                )
                cleaned = cleaned.replace("$", "").replace(",", "").strip("() ")
                suffix = cleaned[-1:].casefold()
                multiplier = {"k": 1_000.0, "m": 1_000_000.0, "b": 1_000_000_000.0}.get(suffix, 1.0)
                if suffix in {"k", "m", "b"}:
                    cleaned = cleaned[:-1].strip()
                cleaned = cleaned.rstrip("%x×").strip()
                try:
                    value = float(cleaned) * multiplier
                except ValueError:
                    continue
                if negative:
                    value = -abs(value)
                if abs(value) >= 100_000:
                    values.append(value)
            if any(
                _close(high - low, float(expected), abs_tol=_display_tolerance(float(expected)), rel_tol=0.0)
                for high in values for low in values if high > low
            ):
                return True

    adverse_magnitude_labels = {
        "construction_gross_profit_impact", "service_labor_productivity_impact",
        "controls_mix_impact", "collections_timing_cash_impact",
    }
    for context in contexts:
        if not any(_context_has_alias(context, alias) for alias in aliases):
            continue
        literals = _TASK_092_NUMERIC_LITERAL_PATTERN.findall(context)
        if label == "remaining_fy26_revenue_outlook":
            normalized = _normalize(context)
            explanatory_need = (
                "remaining revenue need" in normalized
                or "remaining fy26 revenue need" in normalized
            )
            if explanatory_need and any(
                "(" in literal and ")" in literal
                and _task_068_numeric_literal_matches(
                    literal, -abs(float(expected)), context=context
                )
                for literal in literals
            ):
                return True
        if label == "q2_capital_expenditure" and any(
            _task_068_numeric_literal_matches(
                literal, abs(float(expected)), context=context
            )
            or _task_068_numeric_literal_matches(
                literal, -abs(float(expected)), context=context
            )
            for literal in literals
        ):
            return True
        if label in adverse_magnitude_labels:
            normalized = _normalize(context)
            adverse = any(term in normalized for term in (
                "unfavorable", "unfav", "adverse", "drag", "miss", "risk",
                "erosion", "pressure", "timing",
            ))
            recovery = any(term in normalized for term in (
                "recovery", "action", "mitigation", "opportunity", "action pool",
            ))
            if adverse and any(
                _task_068_numeric_literal_matches(
                    literal, -abs(float(expected)), context=context
                )
                or (
                    _task_068_numeric_literal_matches(
                        literal, abs(float(expected)), context=context
                    )
                    and not any(sign in literal for sign in ("+", "-", "−", "("))
                )
                for literal in literals
            ):
                return True
            if recovery and any(
                _task_068_numeric_literal_matches(
                    literal, abs(float(expected)), context=context
                )
                for literal in literals
            ):
                return True
            continue
        if any(
            _task_068_numeric_literal_matches(
                literal, float(expected), context=context
            )
            for literal in literals
        ):
            return True
    return False


def _task_072_artifact_value(path: Path, label: str, expected: Any) -> bool:
    """Read normal management-memo prose and comparison tables.

    The memo may put units once in a table heading and use ordinary labels
    such as ``Plan`` or ``Variance`` within a clearly named row.  Keep the
    association local to that paragraph/table context, while accepting the
    common dollar-in-millions presentation used in board materials.
    """

    aliases = [label.replace("_", " "), *_TASK_072_LABEL_ALIASES.get(label, [])]
    document = Document(path)
    contexts = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
    units_in_millions = any(
        token in _normalize("\n".join(contexts))
        for token in (
            "dollars in millions", "usd millions", "usd mm",
            "in millions", "millions", "mm",
        )
    )
    for table in document.tables:
        rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
        if not rows:
            continue
        header = rows[0]
        header_text = " | ".join(header)
        if any(token in _normalize(header_text) for token in ("in millions", "usd mm", "dollars mm")):
            units_in_millions = True
        for row in rows[1:]:
            row_text = " | ".join(row)
            contexts.append(row_text)
            contexts.append(" | ".join([*header, *row]))
            for column, cell in enumerate(row):
                contexts.append(" | ".join((row[0] if row else "", header[column] if column < len(header) else "", cell)))

    if isinstance(expected, bool):
        if label == "guidance_update_required" and expected is False:
            text = _normalize("\n".join(contexts))
            maintained = any(phrase in text for phrase in (
                "no guidance update", "maintain guidance", "hold guidance",
                "guidance unchanged", "do not revise guidance",
            ))
            contrary = any(phrase in text for phrase in (
                "guidance update required", "update guidance now",
                "revise guidance now", "withdraw guidance now",
                "reduce guidance now", "cut guidance now",
            ))
            return maintained and not contrary
        return any(
            any(_context_has_alias(context, alias) for alias in aliases)
            and _boolean_matches(context, expected)
            for context in contexts
        )
    if isinstance(expected, str):
        return any(
            any(_context_has_alias(context, alias) for alias in aliases)
            and _sample_directional_semantic_value_matches(context, expected)
            for context in contexts
        )

    for context in contexts:
        if not any(_context_has_alias(context, alias) for alias in aliases):
            continue
        literals = _TASK_092_NUMERIC_LITERAL_PATTERN.findall(context)
        if any(_task_092_numeric_literal_matches(literal, float(expected)) for literal in literals):
            return True
        if units_in_millions and abs(float(expected)) >= 100_000:
            for literal in literals:
                if re.search(r"%|[x×]|\b(?:k|m|mm|million|b|bn|billion)\b", literal, flags=re.I):
                    continue
                candidates = list(iter_numeric_candidates(literal))
                if any(
                    _close(value * 1_000_000.0, float(expected), abs_tol=5_000.0, rel_tol=0.0)
                    for value in candidates
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
    entries.extend(_presentation_contexts(path))
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
        local_boolean_text = _normalize("\n".join(
            entry
            for entry in entries
            if any(_context_has_alias(entry, alias) for alias in aliases)
        ))
        positive_breach_text = re.sub(
            r"\b(?:no|not a|without)\s+(?:cash\s+|leverage\s+|covenant\s+)?breach\b"
            r"|\bbreach\s*[:=\-]?\s*(?:false|no|none|not present)\b",
            "",
            local_boolean_text,
        )
        if label == "severe_covenant_breach":
            severe_is_flagged = bool(
                re.search(r"severe.{0,240}breach", positive_breach_text)
                or re.search(r"breach.{0,240}severe", positive_breach_text)
                or (
                    "breach flags" in positive_breach_text
                    and "severe" in positive_breach_text
                    and "executed leverage" in positive_breach_text
                )
            )
            return severe_is_flagged if expected else not severe_is_flagged
        if label == "post_action_severe_cash_breach":
            flagged = bool(
                re.search(r"breach.{0,120}cash\s*<", positive_breach_text)
                or re.search(r"cash.{0,120}breach", positive_breach_text)
            )
            return flagged if expected else not flagged
        if label == "post_action_severe_leverage_breach":
            flagged = bool(
                re.search(r"breach.{0,180}leverage\s*>", positive_breach_text)
                or re.search(r"leverage.{0,120}breach", positive_breach_text)
            )
            return flagged if expected else not flagged
        # A board slide often carries several yes/no gates in one text box.
        # Judging the entire box makes one nearby TRUE contaminate a correctly
        # labeled FALSE (or vice versa).  Bind the boolean to its own rendered
        # line first; then allow only a short suffix after the label.
        for entry in entries:
            for line in str(entry or "").splitlines():
                if not any(_context_has_alias(line, alias) for alias in aliases):
                    continue
                if _boolean_matches(line, expected):
                    return True
                if label.endswith("_breach"):
                    status = re.split(r"\||:", line)[-1].strip()
                    normalized_status = _normalize(status)
                    if normalized_status in {"breach", "failed", "fail"}:
                        return expected is True
                    if normalized_status in {"ok", "pass", "compliant", "within", "no breach"}:
                        return expected is False
            normalized_entry = _normalize(entry)
            for alias in aliases:
                normalized_alias = _normalize(alias)
                start = normalized_entry.find(normalized_alias)
                if start < 0:
                    continue
                local = normalized_entry[start:start + len(normalized_alias) + 40]
                if _boolean_matches(local, expected):
                    return True
        return False

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
            and _sample_directional_semantic_value_matches(segment, expected)
            for segment in entries
        )

    pattern = re.compile(r"\(?[-−]?\$?\d[\d,]*(?:\.\d+)?[ \t]*(?:%|[kmbx×])?\)?", flags=re.I)
    for index, entry in enumerate(entries):
        if not any(_context_has_alias(entry, alias) for alias in aliases):
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
    index = _task_037_workbook_index(workbook, values)
    for project, records in index["cached_projects"].items():
        for record in records:
            row_values = record["cached"][:12]
            chosen = any(
                _sample_directional_semantic_value_matches(value, "selected")
                or value is True
                or (isinstance(value, (int, float)) and not isinstance(value, bool) and float(value) == 1.0)
                for value in row_values
            )
            if chosen and project not in selected:
                selected.append(project)
    return unordered_semantic_list_matches(selected, expected)


def _task_037_workbook_index(workbook, values) -> dict[str, Any]:
    """Index Task037 rows once instead of rescanning 8,000+ cells per fact.

    Task037 deliberately has hundreds of atomic finance checks.  The prior
    implementation performed a full workbook scan for nearly every one, which
    was exact but quadratic on a realistic portfolio model and could exhaust
    HUD's rollout clock after the agent had already finished.  This index holds
    the same formula/cached values and does not change any matching rule.
    """

    cached = getattr(workbook, "_alder_task_037_index", None)
    if cached is not None:
        return cached
    rows: list[dict[str, Any]] = []
    by_sheet_row: dict[tuple[str, int], dict[str, Any]] = {}
    formula_projects: dict[str, list[dict[str, Any]]] = {}
    cached_projects: dict[str, list[dict[str, Any]]] = {}
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        max_row, max_column = sheet.max_row, sheet.max_column
        for row_index in range(1, max_row + 1):
            formula_values = tuple(
                sheet.cell(row_index, column).value
                for column in range(1, max_column + 1)
            )
            cached_values = tuple(
                value_sheet.cell(row_index, column).value
                for column in range(1, max_column + 1)
            )
            record = {
                "sheet": sheet_name,
                "row": row_index,
                "max_column": max_column,
                "formula": formula_values,
                "cached": cached_values,
            }
            rows.append(record)
            by_sheet_row[(sheet_name, row_index)] = record
            for value in formula_values:
                for project in re.findall(r"\bCP-\d{2}\b", str(value or ""), flags=re.I):
                    bucket = formula_projects.setdefault(project.upper(), [])
                    if not bucket or bucket[-1] is not record:
                        bucket.append(record)
            for value in cached_values:
                if re.fullmatch(r"CP-\d{2}", str(value or ""), flags=re.I):
                    cached_projects.setdefault(str(value).upper(), []).append(record)
                    break
    cached = {
        "rows": rows,
        "by_sheet_row": by_sheet_row,
        "formula_projects": formula_projects,
        "cached_projects": cached_projects,
    }
    setattr(workbook, "_alder_task_037_index", cached)
    return cached


def _professional_row_value_match(
    workbook,
    values,
    aliases: list[str],
    expected: Any,
    require_formula: bool = False,
    year: int | None = None,
    *,
    directional_strings: bool = False,
) -> tuple[bool, str]:
    """Match a normal finance schedule by row label and, when needed, year column."""
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        year_columns: set[int] = set()
        if year is not None:
            for row in sheet.iter_rows():
                for cell in row:
                    if semantic_equal(cell.value, f"Year {year}"):
                        year_columns.add(cell.column)
        for row_index in range(1, sheet.max_row + 1):
            row_labels = [sheet.cell(row_index, column).value for column in range(1, sheet.max_column + 1)]
            if not any(any(contains_concept(value, alias) for alias in aliases) for value in row_labels):
                continue
            columns = sorted(year_columns) if year_columns else list(range(1, sheet.max_column + 1))
            for column in columns:
                formula_value = sheet.cell(row_index, column).value
                cached_value = value_sheet.cell(row_index, column).value
                if require_formula and not (isinstance(formula_value, str) and formula_value.startswith("=")):
                    continue
                if _value_candidates_match(
                    [cached_value],
                    expected,
                    directional_strings=directional_strings,
                ):
                    return True, f"professional schedule match at {sheet_name}!{sheet.cell(row_index, column).coordinate}"
    return False, "no professional schedule row match"


def _task_055_release_status_matches(actual: Any, expected: Any) -> bool:
    """Accept the disclosed decision state, not a single committee phrase."""

    normalized = _normalize(actual)
    if not normalized:
        return False
    # The committee rule is directional: any dilutive sensitivity requires a
    # hold/condition/mitigation response.  Keyword overlap must not turn the
    # opposite conclusion (for example, ``mitigation not required``) into a
    # match.  Keep this scoped to Task055 rather than changing semantic
    # matching for the other 99 tasks.
    opposite_release = (
        "release",
        "released",
        "approve",
        "approved",
        "proceed",
        "clear to close",
        "no mitigation",
        "mitigation not required",
        "mitigation is not required",
        "review not required",
        "review is not required",
        "unconditional",
    )
    if any(phrase in normalized for phrase in opposite_release):
        return False
    blocked_or_conditional = (
        "hold",
        "conditional",
        "mitigation required",
        "requires mitigation",
        "review required",
        "requires review",
        "not ready",
        "do not release",
    )
    return (
        semantic_value_matches(actual, expected)
        or any(phrase in normalized for phrase in blocked_or_conditional)
    )


def _task_055_case_cell_matches(value: Any, case: str) -> bool:
    normalized = _normalize(value)
    if not normalized:
        return False
    aliases = _TASK_055_SENSITIVITY_CASE_ALIASES[case]
    if any(
        semantic_equal(value, alias) or contains_concept(value, alias)
        for alias in aliases
    ):
        return True
    # Parenthetical abbreviations such as ``Combined (50% syn + 200 bps)``
    # are common in an IC sensitivity table but do not survive generic phrase
    # matching cleanly.
    if case == "combined_downside":
        return (
            "combined" in normalized
            and "50" in normalized
            and "200" in normalized
            and ("syn" in normalized or "synergy" in normalized)
        )
    return False


def _task_055_metric_header_matches(value: Any, metric: str, year: int | None) -> bool:
    raw = str(value or "")
    normalized = _normalize(raw)
    if not normalized:
        return False
    if year is not None:
        year_markers = {
            str(year), f"y{year}", f"yr{year}", f"year{year}", f"year {year}",
        }
        if not any(marker in normalized.split() for marker in year_markers) and not any(
            marker in normalized for marker in (f"year {year}", f"year{year}")
        ):
            return False

    aliases = _TASK_055_SENSITIVITY_METRIC_ALIASES[metric]
    if any(
        semantic_equal(raw, alias) or contains_concept(raw, alias)
        for alias in aliases
    ):
        return True

    tokens = set(normalized.split())
    is_adjusted = "adjusted" in tokens or "adj" in tokens
    has_delta = any(
        token in tokens
        for token in (
            "accretion", "accretive", "dilution", "dilutive", "acc", "delta", "change",
        )
    ) or "Δ" in raw
    has_eps = "eps" in tokens
    if metric == "gaap_pro_forma_eps":
        return "gaap" in tokens and has_eps and not has_delta
    if metric == "adjusted_pro_forma_eps":
        return is_adjusted and has_eps and not has_delta
    if metric == "gaap_eps_accretion":
        return "gaap" in tokens and has_delta
    if metric == "adjusted_eps_accretion":
        return is_adjusted and has_delta
    return False


def _task_055_sensitivity_matrix_match(
    workbook,
    values,
    label: str,
    expected: Any,
    require_formula: bool,
) -> tuple[bool, str] | None:
    pattern = re.fullmatch(
        r"(synergy_realization_50_percent|debt_rate_up_200_bps|combined_downside)_"
        r"(?:(year_(one|two)_(realized_synergy|gaap_incremental_pre_tax_income|"
        r"adjusted_incremental_pre_tax_income|gaap_pro_forma_eps|"
        r"adjusted_pro_forma_eps|gaap_eps_accretion|adjusted_eps_accretion))|"
        r"(incremental_financing_cost|committee_release_status))",
        label,
    )
    if not pattern:
        return None
    case = pattern.group(1)
    year = 1 if pattern.group(3) == "one" else 2 if pattern.group(3) == "two" else None
    metric = pattern.group(4) or pattern.group(5)

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        value_sheet = values[sheet_name] if sheet_name in values.sheetnames else sheet
        for row in range(1, sheet.max_row + 1):
            if not any(
                _task_055_case_cell_matches(sheet.cell(row, column).value, case)
                for column in range(1, sheet.max_column + 1)
            ):
                continue
            for column in range(1, sheet.max_column + 1):
                headers = [
                    sheet.cell(header_row, column).value
                    for header_row in range(max(1, row - 6), row)
                ]
                if not any(
                    _task_055_metric_header_matches(header, metric, year)
                    for header in headers
                ):
                    continue
                formula_value = sheet.cell(row, column).value
                cached_value = value_sheet.cell(row, column).value
                if require_formula and not (
                    isinstance(formula_value, str) and formula_value.startswith("=")
                ):
                    continue
                matched = (
                    _task_055_release_status_matches(cached_value, expected)
                    if metric == "committee_release_status"
                    else _value_candidates_match([cached_value], expected)
                )
                if matched:
                    coordinate = sheet.cell(row, column).coordinate
                    return True, f"committee sensitivity matrix match at {sheet_name}!{coordinate}"
    return False, "no case-and-metric sensitivity matrix match"


def _task_055_row_match(workbook, values, label: str, expected: Any, require_formula: bool = False) -> tuple[bool, str]:
    sensitivity = _task_055_sensitivity_matrix_match(
        workbook, values, label, expected, require_formula
    )
    if sensitivity is not None:
        if sensitivity[0]:
            return sensitivity
        # Retain a vertical complete-key schedule, but do not let a value from
        # a different case or year satisfy a matrix criterion.
        matched, evidence = _professional_row_value_match(
            workbook,
            values,
            [label.replace("_", " ")],
            expected,
            require_formula,
            directional_strings=True,
        )
        if matched:
            return matched, evidence
        return sensitivity
    match = re.fullmatch(r"year_(one|two)_(.+)", label)
    if match:
        year = 1 if match.group(1) == "one" else 2
        suffix = match.group(2)
        aliases = _TASK_055_YEAR_ROW_ALIASES.get(suffix)
        if aliases:
            matched, evidence = _professional_row_value_match(
                workbook,
                values,
                aliases,
                expected,
                require_formula,
                year,
                directional_strings=True,
            )
            if matched:
                return matched, evidence
            if isinstance(expected, (int, float)) and expected and suffix in {
                "integration_expense", "incremental_debt_interest", "foregone_cash_yield",
            }:
                return _professional_row_value_match(workbook, values, aliases, -expected, require_formula, year)
            return matched, evidence
    matched, evidence = _professional_row_value_match(
        workbook,
        values,
        [label.replace("_", " "), *_TASK_055_LABEL_ALIASES.get(label, [])],
        expected,
        require_formula,
        directional_strings=True,
    )
    if matched:
        return matched, evidence
    # Expenses are commonly displayed as deductions in professional accretion
    # models even though the gold packet stores their positive magnitude.
    if isinstance(expected, (int, float)) and expected and label in {
        "incremental_debt_interest", "foregone_cash_yield",
        "total_incremental_financing_cost",
    }:
        return _professional_row_value_match(
            workbook,
            values,
            [label.replace("_", " "), *_TASK_055_LABEL_ALIASES.get(label, [])],
            -expected,
            require_formula,
        )
    return matched, evidence


def _task_061_row_match(workbook, values, label: str, expected: Any, require_formula: bool = False) -> tuple[bool, str]:
    aliases = [label.replace("_", " "), *_TASK_061_LABEL_ALIASES.get(label, [])]
    matched, evidence = _professional_row_value_match(
        workbook,
        values,
        aliases,
        expected,
        require_formula,
        directional_strings=True,
    )
    if matched or label != "ending_valuation_allowance" or not isinstance(expected, (int, float)):
        return matched, evidence
    # A valuation allowance is professionally presented either as a positive
    # contra-DTA magnitude or as a negative deduction. The independently
    # graded net-DTA roll-forward still enforces the correct economic sign.
    return _professional_row_value_match(
        workbook, values, aliases, -expected, require_formula
    )


def _task_037_row_match(workbook, values, label: str, expected: Any, require_formula: bool = False) -> tuple[bool, str]:
    index = _task_037_workbook_index(workbook, values)
    project_match = re.fullmatch(r"cp_(\d{2})_(.+)", label)
    if project_match:
        project = f"CP-{project_match.group(1)}"
        suffix = project_match.group(2)
        aliases = [suffix.replace("_", " "), *_TASK_037_LABEL_ALIASES.get(label, [])]
        year_flow = re.fullmatch(r"year_(\d)_(base|downside)_after_tax_cash_flow", suffix)
        if year_flow:
            year, case = year_flow.groups()
            aliases.extend([
                f"Yr{year} {'down' if case == 'downside' else 'CF'}",
                f"Year {year} {case} cash flow",
            ])
        aliases.extend({
            "technician_capacity": ["Technician req.", "Technician capacity required"],
            "mandatory_safety_flag": ["Mandatory safety"],
            "npv": ["NPV base", "NPV (base)"],
            "irr": ["IRR base", "IRR (base)"],
            "payback_years": ["Disc payback", "Payback"],
            "selected": ["Selected", "Select", "Portfolio decision"],
        }.get(suffix, []))
        for record in index["formula_projects"].get(project, []):
            sheet_name = record["sheet"]
            row_index = record["row"]
            for column in range(1, record["max_column"] + 1):
                headers = [
                    index["by_sheet_row"][(sheet_name, header_row)]["formula"][column - 1]
                    for header_row in range(max(1, row_index - 4), row_index)
                ]
                if not any(any(contains_concept(header, alias) for alias in aliases) for header in headers):
                    continue
                formula_value = record["formula"][column - 1]
                cached_value = record["cached"][column - 1]
                if require_formula and not (isinstance(formula_value, str) and formula_value.startswith("=")):
                    continue
                if _value_candidates_match(
                    [cached_value], expected, directional_strings=True
                ):
                    coordinate = f"{get_column_letter(column)}{row_index}"
                    return True, f"project schedule match at {sheet_name}!{coordinate}"
    aliases = [label.replace("_", " "), *_TASK_037_LABEL_ALIASES.get(label, [])]
    for record in index["rows"]:
        if not any(
            contains_concept(value, alias)
            for value in record["formula"]
            for alias in aliases
        ):
            continue
        for column, (formula_value, cached_value) in enumerate(
            zip(record["formula"], record["cached"]), start=1
        ):
            if require_formula and not (
                isinstance(formula_value, str) and formula_value.startswith("=")
            ):
                continue
            if _value_candidates_match(
                [cached_value], expected, directional_strings=True
            ):
                coordinate = f"{get_column_letter(column)}{record['row']}"
                return True, (
                    f"professional schedule match at {record['sheet']}!{coordinate}"
                )
    return False, "no professional schedule row match"


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
    if task_id == "task_035":
        candidates.extend(_TASK_035_LABEL_ALIASES.get(label, []))
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
    if task_id == "task_037":
        matched, detail = _task_037_row_match(
            workbook, values, label, gold["answer"][label], require_formula=True
        )
        if matched:
            return matched, detail
    attempts = [_xlsx_label_formula(workbook, candidate) for candidate in candidates]
    matched, detail = next((attempt for attempt in attempts if attempt[0]), attempts[0])
    if task_id == "task_035" and not matched:
        matched, detail = _task_035_row_match(
            workbook,
            values,
            label,
            gold["answer"][label],
            require_formula=True,
        )
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
        if not matched:
            matched, detail = _task_055_row_match(
                workbook, values, label, gold["answer"][label], require_formula=True
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
        elif kind == "xlsx_sheet_lineage":
            target_name = str(spec["target_sheet"])
            source_name = str(spec["source_sheet"])
            target_sheet = workbook[target_name] if target_name in workbook.sheetnames else None
            quoted_reference = f"'{source_name.replace(chr(39), chr(39) * 2)}'!".casefold()
            unquoted_reference = f"{source_name}!".casefold()
            matches = [] if target_sheet is None else [
                f"{target_name}!{cell.coordinate}={cell.value}"
                for row in target_sheet.iter_rows()
                for cell in row
                if isinstance(cell.value, str)
                and cell.value.startswith("=")
                and (
                    quoted_reference in cell.value.casefold()
                    or unquoted_reference in cell.value.casefold()
                )
            ]
            met = bool(matches)
            evidence = (
                f"target_sheet={target_name!r}; source_sheet={source_name!r}; "
                f"linked_formula_examples={matches[:3]!r}"
            )
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
                if task_id == "task_035":
                    candidates.extend(_TASK_035_LABEL_ALIASES.get(label, []))
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
                if task_id == "task_035" and not matched:
                    expected = gold["answer"][label]
                    matched, detail = _task_035_row_match(
                        workbook,
                        values,
                        label,
                        expected,
                        require_formula=True,
                    )
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
                if task_id == "task_055" and not matched:
                    expected = gold["answer"][label]
                    matched, detail = _task_055_row_match(
                        workbook, values, label, expected, require_formula=True
                    )
                if task_id == "task_037" and not matched:
                    expected = gold["answer"][label]
                    matched, detail = _task_037_row_match(
                        workbook, values, label, expected, require_formula=True
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
                if task_id == "task_035":
                    labels.extend(_TASK_035_LABEL_ALIASES.get(label, []))
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
                if task_id == "task_037":
                    if label == "selected_portfolio" and isinstance(expected, list):
                        matched = _task_037_selected_projects(workbook, values, expected)
                    else:
                        matched, _ = _task_037_row_match(
                            workbook, values, label, expected, require_formula=False
                        )
                    if not matched:
                        matched = any(
                            _xlsx_label_value(
                                workbook,
                                values,
                                candidate,
                                expected,
                                directional_strings=True,
                            )
                            for candidate in [
                                *labels,
                                *_TASK_037_LABEL_ALIASES.get(label, []),
                            ]
                        )
                elif task_id == "task_027":
                    # Task027 publishes one normalized controller output row
                    # for every graded headline. Formatting may change, but an
                    # incorrect value on that row must not be rescued by the
                    # same number appearing in a different scenario, bridge,
                    # or source sheet.
                    matched = _xlsx_exact_label_value(
                        workbook,
                        values,
                        label.replace("_", " "),
                        expected,
                        directional_strings=True,
                    )
                else:
                    matched = any(
                        _xlsx_label_value(
                            workbook,
                            values,
                            candidate,
                            expected,
                            directional_strings=task_id
                            in {"task_035", "task_055", "task_061"},
                        )
                        for candidate in labels
                    )
                if task_id == "task_076" and not matched:
                    matched, _ = _task_076_row_match(workbook, values, label, expected, require_formula=False)
                if task_id == "task_035" and not matched:
                    matched, _ = _task_035_row_match(
                        workbook,
                        values,
                        label,
                        expected,
                        require_formula=False,
                    )
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
                if task_id == "task_055" and not matched:
                    matched, _ = _task_055_row_match(workbook, values, label, expected, require_formula=False)
                if task_id == "task_061" and not matched:
                    matched, _ = _task_061_row_match(workbook, values, label, expected, require_formula=False)
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
                if task_id in {"task_068", "task_072"}:
                    labels.extend(_TASK_072_LABEL_ALIASES.get(label, []))
                if task_id == "task_075":
                    labels.extend(_TASK_075_LABEL_ALIASES.get(label, []))
                if task_id == "task_092":
                    labels.extend(_TASK_092_LABEL_ALIASES.get(label, []))
                if task_id == "task_100":
                    labels.extend(_TASK_100_LABEL_ALIASES.get(label, []))
                matched = any(
                    _artifact_label_value(
                        entries,
                        candidate,
                        expected,
                        directional_strings=task_id
                        in {"task_068", "task_072", "task_100"},
                    )
                    for candidate in labels
                )
                if task_id == "task_092" and not matched:
                    matched = _task_092_artifact_value(path, label, expected)
                if task_id == "task_068" and not matched:
                    matched = _task_068_artifact_value(path, label, expected)
                if task_id == "task_072" and not matched:
                    matched = _task_072_artifact_value(path, label, expected)
                if task_id == "task_100" and not matched:
                    matched = _task_100_artifact_value(path, label, expected)
                if not matched:
                    failures.append(label)
            met = not failures
            evidence = "all labeled values matched" if met else f"missing or incorrect labels={failures!r}"
        elif kind == "artifact_tokens":
            if task_id == "task_035" and spec["id"].startswith(("model_content__", "controls__")):
                met, evidence = _task_035_model_or_control_check(
                    workbook, values, spec["id"]
                )
            else:
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
            joined_entries = "\n".join(entries)
            title_present = contains_concept(joined_entries, spec["title_token"])
            if task_id == "task_100" and not title_present:
                normalized_title = _normalize(joined_entries)
                title_present = all(token in normalized_title for token in ("fy26", "outlook", "fy27", "priorities"))
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
    if task_id == "task_053":
        result["decision_support"] = _task_053_decision_support(
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
