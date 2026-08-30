from __future__ import annotations

import copy
import re
from typing import Any, Iterable, Mapping


RUBRIC_SCHEMA_VERSION = 3
REWARD_SCHEMA_VERSION = 8
ALLOWED_WEIGHTS = frozenset({1, 3, 5, 10})

# A finance deliverable is not professionally usable when its material decision
# outputs are wrong, even if its supporting model and audit trail are strong.
# Preserve credit for useful underlying work without allowing it to mask wrong
# business conclusions. The task-specific policy below adjusts the ordinary
# weighted score continuously using supporting accuracy and decision quality.
DECISION_ACCURACY_CRITERIA = {
    "task_010": (
        "top_three_projects__rank_1",
        "top_three_projects__rank_2",
        "top_three_projects__rank_3",
        "recommended_follow_up__arm_2514",
        "recommended_follow_up__arm_2510",
        "recommended_follow_up__arm_2506",
        "portfolio_intensity__highest_project",
        "portfolio_intensity__separate_follow_up",
        "decision__largest_absolute_rate_basis_variance_project",
        "decision__deck_action",
        "decision__cost_treatment",
    ),
    "task_024": (
        # The workbook is a funding-decision screen.  Correct supporting rows
        # and formulas do not, by themselves, make it decision-useful when the
        # committed-versus-indicative financing conclusions are wrong.
        "request__cx-26-017__equipment_line_proceeds",
        "request__cx-26-017__cash_required",
        "request__cx-26-017__classification",
        "summary_value__approved_spend",
        "summary_value__equipment_line_proceeds",
        "summary_value__cash_required",
        "summary_value__facility_remaining",
        "summary_value__aop_remaining_after_approved_requests",
        "summary_value__downside_maximum_revolver",
        "summary_value__revolver_capacity_after_downside",
        "summary_value__pro_forma_leverage",
        "funding_views__approved_spend__committed",
        "funding_views__approved_spend__indicative",
        "funding_views__equipment_line_proceeds__committed",
        "funding_views__equipment_line_proceeds__indicative",
        "funding_views__cash_required__committed",
        "funding_views__cash_required__indicative",
        "funding_views__pro_forma_funded_debt__committed",
        "funding_views__pro_forma_funded_debt__indicative",
        "funding_views__pro_forma_cash__committed",
        "funding_views__pro_forma_cash__indicative",
        "funding_views__pro_forma_net_debt__committed",
        "funding_views__pro_forma_net_debt__indicative",
        "funding_views__pro_forma_leverage__committed",
        "funding_views__pro_forma_leverage__indicative",
    ),
    "task_048": (
        "headline_values__first_covenant_breach",
        "headline_values__required_debt_paydown",
        "headline_values__maximum_leverage",
        "headline_values__minimum_fixed_charge_coverage",
        "headline_values__first_liquidity_floor_breach",
        "headline_values__minimum_pre_financing_cash",
        "headline_values__cash_available_for_cure",
        "headline_values__cash_cure_funding_shortfall",
        "headline_values__first_tangible_net_worth_breach",
        "headline_values__minimum_tangible_net_worth",
        "downside_values__downside_first_covenant_breach",
        "downside_values__downside_required_debt_paydown",
        "downside_values__downside_maximum_leverage",
        "downside_values__downside_minimum_fixed_charge_coverage",
        "downside_values__downside_first_liquidity_floor_breach",
        "downside_values__downside_minimum_pre_financing_cash",
        "downside_values__downside_cash_available_for_cure",
        "downside_values__downside_cash_cure_funding_shortfall",
        "downside_values__downside_first_tangible_net_worth_breach",
        "downside_values__downside_minimum_tangible_net_worth",
    ),
    "task_053": (
        "headline_values__enterprise_value",
        "headline_values__equity_value",
        "headline_values__present_value_of_terminal_value",
        "headline_values__terminal_value_share_of_enterprise_value",
        "headline_values__downside_equity_value",
        "headline_values__upside_equity_value",
        "headline_values__equity_value_sensitivity_range",
    ),
}

# Task 011 is a bounded operating-review package. Normalize the ten
# professionally distinct sections so the executive deck cannot crowd out the
# independent out-of-time production-readiness decision or its formula-driven
# audit trail.
SECTION_NORMALIZED_REWARD_POLICIES = {
    # Task 001 is a controller sign-off whose central evidence is the complete
    # source-to-ledger population reconciliation.  A memo that gets the headline
    # WIP conclusion right but misidentifies much of the 23-entry / 13-journal
    # population is not review-ready.  Keep every independently reviewable
    # section, but weight the full-population control in proportion to the work
    # and financial assurance it provides.
    "task_001": {
        "sections": (
            {
                "name": "artifact_identity_and_structure",
                "prefixes": ("artifact__", "identity__", "structure__", "metadata__"),
                "weight": 0.01,
            },
            {
                "name": "policy_corrected_june_wip",
                "prefixes": ("june__",),
                "weight": 0.02,
            },
            {
                "name": "may_to_june_bridge",
                "prefixes": ("bridge__",),
                "weight": 0.02,
            },
            {
                "name": "pm_to_policy_close_impact",
                "prefixes": ("impact__",),
                "weight": 0.02,
            },
            {
                "name": "accounting_source_values",
                "criterion_ids": tuple(
                    f"reconciliation__{scope}__{field}"
                    for scope in (
                        "labor", "material", "subcontract", "other",
                        "total_job_cost", "billings", "current_contract",
                    )
                    for field in ("base", "activity", "close")
                ),
                "weight": 0.01,
            },
            {
                "name": "accounting_reconciliation_controls",
                "criterion_ids": (
                    *(
                        f"reconciliation__{scope}__{field}"
                        for scope in (
                            "labor", "material", "subcontract", "other",
                            "total_job_cost", "billings", "current_contract",
                        )
                        for field in ("variance", "rollforward")
                    ),
                    "reconciliation__job_cost_crossfoot__base",
                    "reconciliation__job_cost_crossfoot__activity",
                    "reconciliation__job_cost_crossfoot__close",
                    "reconciliation__schedule_result_completeness",
                ),
                "weight": 0.01,
            },
            {
                "name": "june_transaction_support",
                "prefixes": ("transaction_support__", "transaction_control__", "billing_support__"),
                "weight": 0.04,
            },
            {
                "name": "source_to_ledger_posting_lineage",
                "prefixes": (
                    "posting_lineage__", "posting_account_control__",
                    "billing_lineage__",
                ),
                "weight": 0.04,
            },
            {
                "name": "full_population_june_posting_control",
                "prefixes": ("full_population__",),
                "weight": 0.52,
            },
            {
                "name": "etc_composition_and_recovery_correction",
                "prefixes": ("etc_composition__",),
                "weight": 0.04,
            },
            {
                "name": "commercial_and_accounting_judgment",
                "prefixes": ("judgment__",),
                "weight": 0.15,
            },
            {
                "name": "source_traceability",
                "prefixes": ("evidence__",),
                "weight": 0.02,
            },
            {
                "name": "ownership_and_review_usability",
                "prefixes": ("action__", "usability__"),
                "weight": 0.02,
            },
            {
                "name": "pco_source_authority_and_clearance",
                "prefixes": ("authority__",),
                "weight": 0.05,
            },
            {
                "name": "june_wip_posting_bridge",
                "prefixes": ("journal__",),
                "weight": 0.03,
            },
        ),
    },
    # Task 004 is a controller review, not a contest to fill the most repeated
    # workbook cells. Preserve every atomic check while weighting the six
    # independently reviewable finance workstreams by their close relevance.
    # The accounting control is material because it proves the four WIP rows
    # reconcile to agent-visible system records rather than copied workpaper
    # values; cosmetic preservation remains deliberately small.
    "task_004": {
        "sections": (
            {
                "name": "artifact_preservation",
                "prefixes": ("preservation__",),
                "weight": 0.03,
            },
            {
                "name": "source_inputs_and_provenance",
                "prefixes": ("source__",),
                "weight": 0.08,
            },
            {
                "name": "wip_finance_and_controls",
                "prefixes": (
                    "formula_lineage__", "result__", "total_formula_lineage",
                    "total_result__",
                ),
                "weight": 0.22,
            },
            {
                "name": "commercial_review_and_release_decision",
                "prefixes": ("review__", "decision_bridge__", "decision_control__"),
                "weight": 0.28,
            },
            {
                "name": "unbooked_commercial_sensitivity",
                "prefixes": ("commercial_sensitivity__",),
                "weight": 0.14,
            },
            {
                "name": "accounting_to_workpaper_control",
                "prefixes": ("accounting_control__",),
                "weight": 0.25,
            },
        ),
    },
    # The lender slide is a compact decision artifact.  Normalize the central
    # covenant re-performance, its calculation support, and the disclosed
    # reversal ladder separately so presentation mechanics cannot dominate the
    # finance work and a missing downside analysis cannot be hidden by copying
    # posted/pro-forma headline values.
    "task_015": {
        "sections": (
            {
                "name": "artifact_and_presentation",
                "criterion_ids": (
                    "structure__title",
                    "structure__table",
                ),
                "prefixes": ("preservation__", "style__"),
                "weight": 0.08,
            },
            {
                "name": "covenant_reperformance",
                "prefixes": ("metric__",),
                "weight": 0.25,
            },
            {
                "name": "compliance_and_binding_decision",
                "prefixes": ("status__", "headroom__"),
                "weight": 0.10,
            },
            {
                "name": "calculation_support",
                "prefixes": ("calculation_support__",),
                "weight": 0.27,
            },
            {
                "name": "downside_capacity",
                "criterion_ids": ("structure__wip_reversal_sensitivity",),
                "prefixes": ("sensitivity__",),
                "weight": 0.25,
            },
            {
                "name": "source_authority",
                "prefixes": ("source__",),
                "weight": 0.05,
            },
        ),
    },
    # The executive performance deck is principally an accounting tie and a
    # mitigation decision. Those central workstreams must outweigh mechanics
    # and narrative polish while every underlying finance criterion remains.
    "task_068": {
        "sections": (
            {
                "name": "artifact_sources_and_controls",
                "prefixes": ("preservation__", "sources__", "controls__"),
                "weight": 0.08,
            },
            {
                "name": "accounting_and_reporting_tie",
                "prefixes": tuple(
                    f"headline_values__{label}"
                    for label in (
                        "june_revenue", "june_ebitda", "posted_ytd_revenue",
                        "q2_posted_ytd_revenue_share",
                        "q2_posted_ytd_revenue_reconciliation_check", "q2_revenue",
                        "q2_approved_plan_revenue", "q2_revenue_variance_to_plan",
                        "q2_organic_growth", "q2_adjusted_ebitda",
                        "q2_approved_plan_adjusted_ebitda",
                        "q2_adjusted_ebitda_variance_to_plan",
                        "construction_q2_revenue",
                        "construction_q2_approved_plan_revenue",
                        "construction_q2_revenue_variance_to_plan",
                        "construction_q2_adjusted_ebitda",
                        "construction_q2_approved_plan_adjusted_ebitda",
                        "construction_q2_adjusted_ebitda_variance_to_plan",
                        "service_q2_revenue", "service_q2_approved_plan_revenue",
                        "service_q2_revenue_variance_to_plan",
                        "service_q2_adjusted_ebitda",
                        "service_q2_approved_plan_adjusted_ebitda",
                        "service_q2_adjusted_ebitda_variance_to_plan",
                        "controls_q2_revenue", "controls_q2_approved_plan_revenue",
                        "controls_q2_revenue_variance_to_plan",
                        "controls_q2_adjusted_ebitda",
                        "controls_q2_approved_plan_adjusted_ebitda",
                        "controls_q2_adjusted_ebitda_variance_to_plan",
                        "q2_operating_cash_flow", "q2_capital_expenditure",
                        "q2_free_cash_flow", "q2_free_cash_flow_plan",
                        "q2_free_cash_flow_variance_to_plan",
                        "collections_timing_cash_impact", "q2_unrestricted_cash",
                        "maximum_revolver", "latest_full_year_revenue_outlook",
                        "signed_backlog", "remaining_fy26_revenue_outlook",
                        "signed_backlog_coverage_of_remaining_outlook",
                        "ltm_adjusted_ebitda", "funded_debt", "lender_leverage",
                        "fixed_charge_coverage", "revenue_guidance_low",
                        "revenue_guidance_high", "ebitda_guidance_low",
                        "ebitda_guidance_high",
                    )
                ),
                "weight": 0.40,
            },
            {
                "name": "mitigation_and_release_decision",
                "prefixes": tuple(
                    f"headline_values__{label}"
                    for label in (
                        "identified_ebitda_action_pool", "approved_mitigation_conversion",
                        "executable_ebitda_mitigation",
                        "post_mitigation_combined_stress_ebitda",
                        "post_mitigation_headroom_to_guidance_low",
                        "minimum_guidance_update_headroom", "guidance_update_trigger",
                        "guidance_release_status",
                        "construction_gross_profit_impact",
                        "service_labor_productivity_impact", "controls_mix_impact",
                        "largest_downside_driver", "combined_stress_ebitda",
                        "ebitda_shortfall_to_guidance_low",
                        "additional_revenue_decline_to_update_trigger",
                        "guidance_update_required",
                    )
                ),
                "weight": 0.40,
            },
            {
                "name": "executive_narrative",
                "prefixes": ("narrative__",),
                "weight": 0.12,
            },
        ),
    },
    # Task 037 is a capital-committee decision package, not a contest to repeat
    # the largest number of mechanically similar cash-flow cells.  Preserve
    # every atomic finance and formula check, but grade the five independently
    # reviewable workstreams the way a controller or investment committee
    # would: source/control hygiene, model auditability, project underwriting,
    # the base portfolio decision, and the resilience/release decision.
    "task_037": {
        "sections": (
            {
                "name": "artifact_sources_and_controls",
                "prefixes": (
                    "structure__", "model_integrity__", "model_content__",
                    "sources__", "controls__",
                ),
                "weight": 0.15,
            },
            {
                "name": "formula_auditability",
                "prefixes": ("formula_lineage__",),
                "weight": 0.10,
            },
            {
                "name": "project_underwriting",
                "prefixes": tuple(
                    f"headline_values__cp_{project:02d}_{suffix}"
                    for project in range(1, 9)
                    for suffix in (
                        "initial_capex", "technician_capacity",
                        "mandatory_safety_flag", "npv", "downside_npv",
                        "irr", "payback_years", "selected", "year_",
                    )
                ),
                "weight": 0.25,
            },
            {
                "name": "base_portfolio_decision",
                "prefixes": tuple(
                    f"headline_values__{label}"
                    for label in (
                        "selected_portfolio", "selected_capex",
                        "portfolio_npv", "downside_portfolio_npv",
                        "selected_debt_eligible_basis",
                        "selected_cash_funding",
                        "selected_technician_capacity",
                        "cash_capex_headroom", "debt_capacity_headroom",
                        "technician_capacity_headroom",
                        "mandatory_safety_projects_selected",
                        "cash_constraint_check", "debt_constraint_check",
                        "technician_constraint_check",
                    )
                ) + ("selection_tie",),
                "weight": 0.25,
            },
            {
                "name": "resilience_and_committee_release",
                "prefixes": tuple(
                    f"headline_values__{prefix}"
                    for prefix in (
                        "cp_03_unavailable_", "cp_04_unavailable_",
                        "cash_capacity_down_20_percent_",
                        "debt_capacity_down_25_percent_",
                        "technician_capacity_down_25_percent_",
                        "downside_npv_objective_",
                    )
                ),
                "weight": 0.25,
            },
        ),
    },
    # Task 035's monthly and project-level rows provide necessary audit support,
    # but the deliverable is ultimately an executive release decision.  Keep
    # each professional section proportional while preventing hundreds of
    # mechanically repeated cells from overwhelming an incorrect capacity/P&L
    # release conclusion.
    "task_035": {
        "sections": (
            {
                "name": "artifact_and_sources",
                "prefixes": (
                    "structure__", "model_integrity__", "model_content__",
                    "sources__", "controls__",
                ),
                "weight": 0.08,
            },
            {
                "name": "formula_auditability",
                "prefixes": ("formula_lineage__",),
                "weight": 0.10,
            },
            {
                "name": "probability_plan",
                "prefixes": tuple(
                    f"headline_values__{prefix}"
                    for prefix in (
                        "signed_backlog", "fy27_", "constrained_",
                        "first_constrained_", "maximum_capacity_",
                        "probability_plan_", "controlling_capacity_case",
                    )
                ),
                "weight": 0.14,
            },
            {
                "name": "gross_commitment_stress",
                "prefixes": ("headline_values__gross_commitment_",),
                "weight": 0.18,
            },
            {
                "name": "execution_portfolio",
                "prefixes": ("headline_values__execution_",),
                "weight": 0.20,
            },
            {
                "name": "earnings_release_decision",
                "prefixes": ("headline_values__release_bridge_",),
                "weight": 0.15,
            },
            {
                "name": "executive_recovery_decision",
                "prefixes": ("headline_values__executive_recovery_",),
                "weight": 0.15,
            },
        ),
    },
    # Task 027 is a controller release decision supported by detailed scenario,
    # liquidity, attribution, action, and weekly schedules.  Hundreds of
    # formula-lineage and repeated weekly cells are necessary audit support but
    # must not overwhelm an incorrect covenant conclusion.  Every criterion
    # remains uncapped and proportional within its independently reviewable
    # finance section.
    "task_027": {
        "sections": (
            {
                "name": "artifact_and_sources",
                "prefixes": (
                    "structure__", "model_integrity__", "model_content__",
                    "sources__", "controls__",
                ),
                "weight": 0.05,
            },
            {
                "name": "formula_auditability",
                "prefixes": ("formula_lineage__",),
                "weight": 0.08,
            },
            {
                "name": "scenario_financials",
                "prefixes": (
                    "base_scenario_values__", "downside_scenario_values__",
                    "upside_scenario_values__", "scenario_comparison_values__",
                ),
                "weight": 0.14,
            },
            {
                "name": "quarterly_liquidity",
                "prefixes": (
                    "base_liquidity_values__", "downside_liquidity_values__",
                    "upside_liquidity_values__",
                ),
                "weight": 0.12,
            },
            {
                "name": "scenario_attribution",
                "prefixes": (
                    "downside_scenario_attribution_values__",
                    "upside_scenario_attribution_values__",
                ),
                "weight": 0.10,
            },
            {
                "name": "contingency_selection",
                "prefixes": tuple(
                    f"downside_contingency_values__{prefix}"
                    for prefix in (
                        "contingency_capex_deferral_",
                        "contingency_collections_acceleration_",
                        "contingency_construction_pricing_",
                        "contingency_controls_mix_",
                        "contingency_fixed_opex_freeze_",
                        "contingency_procurement_savings_",
                        "contingency_service_productivity_",
                        "contingency_targeted_overhead_reduction_",
                        "contingency_selected_", "contingency_annual_",
                        "contingency_implementation_", "contingency_adjusted_",
                        "contingency_ebitda_margin_headroom",
                        "contingency_free_cash_flow_headroom",
                    )
                ),
                "weight": 0.11,
            },
            {
                "name": "weekly_cash_support",
                "prefixes": (
                    "downside_contingency_values__contingency_week_",
                    "downside_contingency_values__contingency_13_week_",
                ),
                "weight": 0.10,
            },
            {
                "name": "lender_release_decision",
                "prefixes": ("covenant_release_values__",),
                "weight": 0.30,
            },
        ),
    },
    "task_011": {
        "sections": (
            {"name": "finance", "prefixes": ("finance__",), "weight": 0.10},
            {"name": "management", "prefixes": ("management__",), "weight": 0.05},
            {
                "name": "capacity_decision",
                "prefixes": ("capacity__", "portfolio__"),
                "weight": 0.06,
            },
            {
                "name": "execution_timing",
                "prefixes": ("timing__",),
                "weight": 0.06,
            },
            {
                "name": "development_signal",
                "prefixes": ("signal__",),
                "weight": 0.06,
            },
            {
                "name": "holdout_deployment",
                "prefixes": ("holdout__",),
                "weight": 0.14,
            },
            {
                "name": "out_of_time_validation",
                "prefixes": ("oot__",),
                "weight": 0.30,
            },
            {
                "name": "formula_workpaper",
                "prefixes": ("workpaper__",),
                "weight": 0.16,
            },
            {"name": "artifact_native", "prefixes": ("artifact__",), "weight": 0.04},
            {
                "name": "readability_and_sources",
                "prefixes": ("readability__", "source__"),
                "weight": 0.03,
            },
        ),
    },
    # Task 016 is principally a controller decision over the proposed WIP
    # close.  The cash/GL bridge and other workpapers support that judgment,
    # but hundreds of copied population cells must not overwhelm incorrect
    # risk-project margins, policy treatment, or release disposition.  Keep
    # every criterion proportional and uncapped while weighting the sections
    # by close materiality.
    "task_016": {
        "sections": (
            {
                "name": "format",
                "prefixes": ("format__",),
                "weight": 0.03,
            },
            {
                "name": "ledger_and_workpaper_support",
                "prefixes": (
                    "cash_account_review__", "cash_control__",
                    "cash_population__", "ar_control__", "ar_population__",
                    "ap_control__", "ap_population__", "debt_control__",
                    "debt_population__", "posted_close__", "close_view__",
                    "ar_population_review__", "ar_workpaper_control__",
                    "ap_population_review__", "ap_workpaper_control__",
                    "debt_note_review__", "debt_workpaper_control__",
                    "wip_population__", "wip_population_review__",
                ),
                "weight": 0.03,
            },
            {
                "name": "cash_authority_and_gl",
                "prefixes": (
                    "cash_cross_ledger__", "cash_source_bridge__",
                    "june_gl_rollforward__",
                ),
                "weight": 0.03,
            },
            {
                "name": "wip_risk_decision",
                "prefixes": (
                    "wip_review__", "wip_risk_release_control__",
                ),
                "weight": 0.28,
            },
            {
                "name": "posting_and_statement_release",
                "prefixes": (
                    "close_release_gate__", "close_release_source__",
                    "close_release_blocker__", "close_release_boundary__",
                    "cutoff_release_gate__", "posting_sequence_control__",
                    "financial_statement_release_bridge__",
                    "income_statement_release_bridge__",
                    "close_scenario_decision__", "close_release_queue__",
                ),
                "weight": 0.08,
            },
            {
                "name": "controller_journal_package",
                "prefixes": (
                    "close_journal_package_decision__",
                ),
                "weight": 0.55,
            },
        ),
    },
    # Task 017 is a working-capital actionability review, not merely a source-
    # population transcription exercise.  Its long AR/AP support schedules
    # contain many atomic cells, so ordinary row-weighted scoring otherwise
    # lets repeated support facts overwhelm a wrong executable cash decision.
    # Keep every criterion proportional and uncapped, while assigning the
    # independently reviewable sections their management relevance.
    "task_017": {
        "sections": (
            {
                "name": "format",
                "prefixes": ("format__",),
                "weight": 0.01,
            },
            {
                "name": "measurement_and_cutoff",
                "prefixes": (
                    "ltm_denominators__",
                    "ar_population__",
                    "ap_population__",
                    "days__",
                    "cash_release__",
                    "ar_actionability__",
                    "ap_actionability__",
                    "normalized_view__",
                    "target_bridge__",
                    "decision__",
                ),
                "weight": 0.15,
            },
            {
                "name": "theoretical_whole_item_plan",
                "prefixes": (
                    "whole_item_ar__",
                    "whole_item_ar_schedule__",
                    "whole_item_ap__",
                    "whole_item_combined__",
                    "ar_support__",
                    "ap_support__",
                ),
                "weight": 0.15,
            },
            {
                "name": "feasibility_evidence",
                "prefixes": (
                    "executable_ar_review__",
                    "whole_item_ap_schedule__",
                ),
                "weight": 0.23,
            },
            {
                "name": "executable_cash_decision",
                "prefixes": ("executable_plan__",),
                "weight": 0.46,
            },
        ),
    },
}


# These gates target omitted professional workstreams, not phrases or visual
# choices.  Reward remains proportional inside every section; only a
# materially absent/incorrect central module is prevented from looking like a
# passing submission because easier mechanics elsewhere were completed.
CRITICAL_SECTION_GATES = {
    "task_001": (
        {"section": "policy_corrected_june_wip", "minimum_accuracy": 0.50, "cap": 0.49},
        {"section": "commercial_and_accounting_judgment", "minimum_accuracy": 0.50, "cap": 0.49},
        {"section": "full_population_june_posting_control", "minimum_accuracy": 0.30, "cap": 0.49},
    ),
    "task_004": (
        {"section": "wip_finance_and_controls", "minimum_accuracy": 0.50, "cap": 0.49},
        {"section": "commercial_review_and_release_decision", "minimum_accuracy": 0.45, "cap": 0.49},
        {"section": "accounting_to_workpaper_control", "minimum_accuracy": 0.45, "cap": 0.49},
    ),
    "task_015": (
        {"section": "covenant_reperformance", "minimum_accuracy": 0.55, "cap": 0.49},
        {"section": "calculation_support", "minimum_accuracy": 0.45, "cap": 0.49},
        {"section": "downside_capacity", "minimum_accuracy": 0.35, "cap": 0.49},
    ),
    "task_035": (
        {"section": "gross_commitment_stress", "minimum_accuracy": 0.45, "cap": 0.49},
        {"section": "execution_portfolio", "minimum_accuracy": 0.40, "cap": 0.49},
        {"section": "executive_recovery_decision", "minimum_accuracy": 0.45, "cap": 0.49},
    ),
    "task_068": (
        {"section": "accounting_and_reporting_tie", "minimum_accuracy": 0.55, "cap": 0.49},
        {"section": "mitigation_and_release_decision", "minimum_accuracy": 0.55, "cap": 0.49},
    ),
}


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
            numeric_value = isinstance(value, (int, float)) and not isinstance(value, bool)
            deterministic_numeric = bool(parent.get("deterministic_numeric")) and numeric_value
            proportional = bool(parent.get("proportional"))
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
                    semantic=not deterministic_numeric,
                    failure_cap=None if proportional or deterministic_numeric else (0.49 if decision else None),
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
        rows = []
        if parent.get("require_formula_count", True):
            rows.append(
                _criterion(
                    parent,
                    criterion_id=f"{criterion_id}__formula_graph",
                    description=f"The model contains at least {int(parent['min_formulas'])} substantive formulas",
                    kind="xlsx_formula_count",
                    category="auditability",
                    weight=5,
                    failure_cap=0.69,
                    min_formulas=int(parent["min_formulas"]),
                )
            )
        rows.append(
            _criterion(
                parent,
                criterion_id=f"{criterion_id}__no_errors",
                description="The submitted workbook contains no spreadsheet error values",
                kind="xlsx_no_errors",
                category="integrity",
                weight=10,
                failure_cap=0.0,
            )
        )
        return rows

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
        lineage_pairs = list(parent.get("lineage_pairs", []))
        if lineage_pairs:
            rows.extend(
                _criterion(
                    parent,
                    criterion_id=(
                        f"{criterion_id}__link_{_slug(pair['target_sheet'])}"
                        f"_from_{_slug(pair['source_sheet'])}"
                    ),
                    description=(
                        f"Material calculations on worksheet {pair['target_sheet']!r} "
                        f"are linked to supporting schedule {pair['source_sheet']!r}"
                    ),
                    kind="xlsx_sheet_lineage",
                    category="auditability",
                    weight=3,
                    target_sheet=str(pair["target_sheet"]),
                    source_sheet=str(pair["source_sheet"]),
                )
                for pair in lineage_pairs
            )
        else:
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
    """Compute a monotonic training reward and independent quality gates.

    Partial credit must preserve the professional ordering of two valid but
    incomplete submissions.  Non-zero legacy caps therefore remain visible as
    critical quality-gate metadata and strict-pass blockers, but they do not
    flatten distinct weighted scores.  Only declared zero-reward criteria and
    environment-integrity failures invalidate the scalar reward.
    """

    updated = copy.deepcopy(dict(result))
    criteria = [row for row in updated.get("criteria", []) if isinstance(row, dict)]
    total_weight = 0
    earned_weight = 0
    failed_declared_gates: list[tuple[float, str]] = []
    for row in criteria:
        policy = criterion_policy(row)
        row.update(policy)
        value = int(bool(row.get("value")))
        row["value"] = value
        total_weight += policy["weight"]
        earned_weight += policy["weight"] * value
        if not value and policy["failure_cap"] is not None:
            failed_declared_gates.append(
                (float(policy["failure_cap"]), str(row.get("id")))
            )

    raw_reward = earned_weight / total_weight if total_weight else 0.0
    reward = raw_reward
    decision_accuracy_adjustment = None
    section_normalized_reward = None
    critical_section_gate_failures: list[dict[str, Any]] = []
    core_model_failure = None
    critical_ids = DECISION_ACCURACY_CRITERIA.get(task_id)
    if critical_ids:
        by_id = {str(row.get("id")): row for row in criteria}
        missing_ids = [criterion_id for criterion_id in critical_ids if criterion_id not in by_id]
        if missing_ids:
            raise ValueError(
                f"missing decision-accuracy criteria for {task_id}: {missing_ids}"
            )
        critical_rows = [by_id[criterion_id] for criterion_id in critical_ids]
        critical_met = sum(int(bool(row.get("value"))) for row in critical_rows)
        critical_weight_total = sum(int(row.get("weight", 10)) for row in critical_rows)
        critical_weight_earned = sum(
            int(row.get("weight", 10)) * int(bool(row.get("value")))
            for row in critical_rows
        )
        supporting_weight_total = total_weight - critical_weight_total
        supporting_weight_earned = earned_weight - critical_weight_earned
        if supporting_weight_total <= 0:
            raise ValueError(f"missing supporting criteria for {task_id}")
        supporting_accuracy = supporting_weight_earned / supporting_weight_total
        support_payload = updated.get("decision_support")
        support_by_id = (
            support_payload.get("criteria", {})
            if isinstance(support_payload, Mapping) else {}
        )
        decision_quality_scores: dict[str, float] = {}
        for row in critical_rows:
            criterion_id = str(row.get("id"))
            if bool(row.get("value")):
                score = 1.0
            else:
                support = support_by_id.get(criterion_id, {})
                score = float(support.get("score", 0.0)) if isinstance(support, Mapping) else 0.0
                if not 0.0 <= score <= 0.5:
                    raise ValueError(
                        f"invalid non-passing decision-support score for {criterion_id}: {score}"
                    )
            decision_quality_scores[criterion_id] = score
        decision_accuracy = critical_met / len(critical_ids)
        decision_support_quality = sum(decision_quality_scores.values()) / len(critical_ids)
        maximum_reward_factor = 0.25 + 0.75 * decision_support_quality
        decision_adjusted_reward = supporting_accuracy * maximum_reward_factor
        reward = min(raw_reward, decision_adjusted_reward)
        decision_accuracy_adjustment = {
            "method": "cap_by_supporting_accuracy_and_decision_quality",
            "criteria_ids": list(critical_ids),
            "criteria_met": critical_met,
            "criteria_total": len(critical_ids),
            "decision_accuracy": round(decision_accuracy, 6),
            "decision_support_quality": round(decision_support_quality, 6),
            "decision_quality_scores": {
                criterion_id: round(score, 6)
                for criterion_id, score in decision_quality_scores.items()
            },
            "supporting_weight_earned": supporting_weight_earned,
            "supporting_weight_total": supporting_weight_total,
            "supporting_accuracy": round(supporting_accuracy, 6),
            "maximum_reward_factor": round(maximum_reward_factor, 6),
            "decision_adjusted_reward": round(decision_adjusted_reward, 6),
        }
    section_policy = SECTION_NORMALIZED_REWARD_POLICIES.get(task_id)
    if section_policy and criteria:
        sections = tuple(section_policy.get("sections", ()))
        if not sections:
            raise ValueError(f"missing section-normalized policy for {task_id}")
        section_results: dict[str, dict[str, Any]] = {}
        matched_ids: set[str] = set()
        normalized_reward = 0.0
        declared_weight = 0.0
        for section in sections:
            name = str(section["name"])
            prefixes = tuple(
                str(prefix) for prefix in section.get("prefixes", ())
            )
            criterion_ids = {
                str(criterion_id)
                for criterion_id in section.get("criterion_ids", ())
            }
            rows = [
                row
                for row in criteria
                if str(row.get("id", "")) in criterion_ids
                or (
                    prefixes
                    and str(row.get("id", "")).startswith(prefixes)
                )
            ]
            if not rows:
                raise ValueError(
                    f"missing section-normalized criteria for {task_id}: {name}"
                )
            row_ids = {str(row["id"]) for row in rows}
            if matched_ids.intersection(row_ids):
                raise ValueError(
                    f"overlapping section-normalized criteria for {task_id}: {name}"
                )
            matched_ids.update(row_ids)
            total = sum(int(row["weight"]) for row in rows)
            earned = sum(
                int(row["weight"]) * int(bool(row.get("value")))
                for row in rows
            )
            accuracy = earned / total
            section_weight = float(section["weight"])
            declared_weight += section_weight
            normalized_reward += section_weight * accuracy
            section_results[name] = {
                "prefixes": list(prefixes),
                "criterion_ids": sorted(criterion_ids),
                "section_weight": section_weight,
                "weight_earned": earned,
                "weight_total": total,
                "accuracy": round(accuracy, 6),
            }
        all_ids = {str(row["id"]) for row in criteria}
        if matched_ids != all_ids:
            raise ValueError(
                f"unassigned section-normalized criteria for {task_id}: "
                f"{sorted(all_ids - matched_ids)}"
            )
        if abs(declared_weight - 1.0) > 1e-9:
            raise ValueError(f"invalid section weights for {task_id}")
        reward = normalized_reward
        section_normalized_reward = {
            "method": "linear_multi_section_normalization",
            "sections": section_results,
            "section_normalized_reward": round(reward, 6),
        }
        for gate in CRITICAL_SECTION_GATES.get(task_id, ()):
            section_name = str(gate["section"])
            if section_name not in section_results:
                raise ValueError(
                    f"missing critical section gate for {task_id}: {section_name}"
                )
            accuracy = float(section_results[section_name]["accuracy"])
            minimum_accuracy = float(gate["minimum_accuracy"])
            if accuracy < minimum_accuracy:
                cap = float(gate["cap"])
                reward = min(reward, cap)
                critical_section_gate_failures.append({
                    "criterion_id": f"critical_section__{section_name}",
                    "section": section_name,
                    "section_accuracy": round(accuracy, 6),
                    "minimum_accuracy": minimum_accuracy,
                    "cap": cap,
                    "reward_effect": "central_workstream_pass_cap",
                })
    required_editable_artifact_by_task = {
        "task_001": (
            "Shared/Finance/Close/2026/06 June/4 WIP/"
            "ARM-2409 June WIP controller sign-off - WORKING.docx"
        ),
        "task_027": (
            "Shared/Finance/FP&A/FY27 plan/"
            "FY27 EBITDA scenarios - WORKING.xlsx"
        ),
        "task_031": (
            "Shared/Finance/FP&A/FY27 plan/"
            "FY27 field workforce plan - WORKING.xlsx"
        ),
        "task_035": (
            "Shared/Finance/FP&A/Backlog/"
            "FY27 backlog burn and capacity - WORKING.xlsx"
        ),
        "task_041": (
            "Shared/Finance/Treasury/Cash positioning/"
            "July daily cash position - WORKING.xlsx"
        ),
        "task_060": (
            "Shared/Finance/Corporate Development/"
            "Orion integration value capture - WORKING.xlsx"
        ),
        "task_067": (
            "Shared/Finance/Reporting/Board/Q2 2026/"
            "Board performance dashboard - WORKING.xlsx"
        ),
        "task_068": (
            "Shared/Finance/Reporting/2026/06 June/"
            "June executive performance review - WORKING.pptx"
        ),
    }
    if task_id in required_editable_artifact_by_task and integrity is not None:
        required_artifact = required_editable_artifact_by_task[task_id]
        declared_artifact = str(integrity.get("required_artifact") or "")
        changed_files = {
            str(path)
            for path in integrity.get("changed_files", [])
            if path is not None
        }
        created_files = {
            str(path)
            for path in integrity.get("created_files", [])
            if path is not None
        }
        if (
            declared_artifact == required_artifact
            and required_artifact not in changed_files | created_files
        ):
            core_model_failure = {
                "code": (
                    "missing_required_artifact_modification"
                    if task_id == "task_001"
                    else "missing_required_workbook_modification"
                ),
                "required_artifact": required_artifact,
                "changed_files_count": len(changed_files),
                "created_files_count": len(created_files),
                "reward_effect": "zero_reward_hard_failure",
            }
    if task_id == "task_001":
        task_001_structure_only_ids = {
            "full_population__detail__row_count",
            "full_population__journal__row_count",
        }
        substantive_prefixes = (
            "june__", "bridge__", "impact__", "reconciliation__",
            "transaction_support__", "transaction_control__", "billing_support__",
            "posting_lineage__", "posting_account_control__", "billing_lineage__",
            "full_population__",
            "etc_composition__", "judgment__", "evidence__", "action__",
        )
        substantive_met_ids = sorted(
            str(row.get("id"))
            for row in criteria
            if str(row.get("id", "")).startswith(substantive_prefixes)
            and str(row.get("id", "")) not in task_001_structure_only_ids
            and bool(row.get("value"))
        )
        if not substantive_met_ids:
            core_model_failure = {
                "code": "no_substantive_finance_work",
                "substantive_criteria_met": 0,
                "reward_effect": "zero_reward_hard_failure",
            }
    # Tasks079/080 are response-only finance analyses, so there is no workbook
    # mutation to prove that substantive work occurred.  A blank/default JSON
    # can otherwise receive a small reward when reference constants happen to
    # be false or zero.  Exclude only those known default-collision criteria
    # when deciding whether any substantive finance work was completed; one
    # genuinely correct calculated value still retains proportional credit.
    if task_id in {"task_079", "task_080"}:
        default_collision_ids = {
            "task_079": {
                "renewal_downside_release_ready__renewal_downside_release_ready",
            },
            "task_080": {
                "revolver_capacity_breach__revolver_capacity_breach",
                (
                    "severe_unfunded_liquidity_shortfall__"
                    "severe_unfunded_liquidity_shortfall"
                ),
            },
        }[task_id]
        met_ids = {
            str(row.get("id")) for row in criteria if bool(row.get("value"))
        }
        substantive_met_ids = sorted(met_ids - default_collision_ids)
        if not substantive_met_ids:
            core_model_failure = {
                "code": "no_substantive_finance_work",
                "criteria_met": len(met_ids),
                "substantive_criteria_met": 0,
                "excluded_default_match_criteria": sorted(
                    met_ids.intersection(default_collision_ids)
                ),
                "reward_effect": "zero_reward_hard_failure",
            }
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
    quality_gate_failures = []
    for gate in critical_section_gate_failures:
        applied_caps.append({
            "criterion_id": gate["criterion_id"],
            "cap": gate["cap"],
        })
        quality_gate_failures.append(dict(gate))
    if core_model_failure is not None:
        reward = 0.0
        applied_caps.append(
            {
                "criterion_id": core_model_failure["code"],
                "cap": 0.0,
            }
        )
        quality_gate_failures.append(dict(core_model_failure))
    for cap, criterion_id in sorted(failed_declared_gates):
        gate = {"criterion_id": criterion_id, "declared_threshold": cap}
        if cap == 0.0:
            reward = 0.0
            applied_caps.append({"criterion_id": criterion_id, "cap": cap})
            gate["reward_effect"] = "zero_reward_hard_failure"
        else:
            gate["reward_effect"] = "strict_pass_blocker_only"
        quality_gate_failures.append(gate)

    met_count = sum(int(bool(row.get("value"))) for row in criteria)
    updated.update(
        {
            "reward": round(reward, 6),
            "raw_weighted_reward": round(raw_reward, 6),
            "decision_accuracy_adjustment": decision_accuracy_adjustment,
            "section_normalized_reward": section_normalized_reward,
            "core_model_failure": core_model_failure,
            "strict_pass": (
                bool(criteria)
                and met_count == len(criteria)
                and not hard_failures
                and core_model_failure is None
                and not any(cap == 0.0 for cap, _criterion_id in failed_declared_gates)
            ),
            "criteria_met": met_count,
            "criteria_total": len(criteria),
            "weight_earned": earned_weight,
            "weight_total": total_weight,
            "reward_schema_version": REWARD_SCHEMA_VERSION,
            "reward_definition": (
                "weighted binary criteria using 1/3/5/10 importance; task-declared "
                "linear section normalization keeps repeated support schedules "
                "proportional to separately weighted management judgment; "
                "tasks with a declared required editable artifact receive zero when that "
                "artifact is not modified or created; "
                "Tasks079 and 080 blank/default responses with no substantive finance work "
                "receive zero; "
                "decision-quality caps preserve supporting-work credit, distinguish correct "
                "schedule support from omitted facts, and prevent "
                "wrong critical business outputs from being masked; non-zero declared quality thresholds block strict "
                "pass, while task-declared central-workstream gates cap only materially "
                "incomplete professional modules; "
                "non-zero declared quality thresholds otherwise do not clip partial credit; "
                "declared zero-reward criteria and environment-integrity failures score zero"
            ),
            "applied_reward_caps": applied_caps,
            "quality_gate_failures": quality_gate_failures,
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
