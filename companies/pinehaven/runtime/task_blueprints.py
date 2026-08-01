from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class TaskBlueprint:
    task_id: str
    title: str
    workflow: str
    output_mode: str
    bundle: str
    parameters: dict[str, Any]
    source_hint: str
    slug: str
    target_path: str | None = None


BUNDLE_KEYS: dict[str, tuple[str, ...]] = {
    "close_control": (
        "system_cash",
        "system_ar",
        "system_ap",
        "system_inventory",
        "system_wip",
        "system_production_variance",
        "gl_5110_production_variance",
        "production_variance_control_delta",
        "system_debt",
        "journal_imbalance",
        "maximum_control_delta",
    ),
    "pnl": (
        "revenue",
        "cost_of_goods_sold",
        "gross_profit",
        "gross_margin",
        "operating_expense",
        "operating_income",
        "period_status",
    ),
    "production_variance": (
        "order_count",
        "standard_output_cost",
        "material_usage_variance",
        "labor_rate_variance",
        "labor_efficiency_variance",
        "variable_overhead_variance",
        "fixed_overhead_volume_variance",
        "scrap_variance",
        "rounding_adjustment",
        "total_variance",
        "variance_rate",
    ),
    "wip": (
        "order_count",
        "started_orders",
        "reported_finished_orders",
        "cost_to_date",
        "reported_output",
        "wip_value",
        "largest_order",
        "largest_order_wip",
        "top_five_concentration",
    ),
    "plant_output": (
        "order_count",
        "completed_units",
        "scrap_units",
        "yield_rate",
        "standard_output",
        "total_variance",
        "variance_rate",
    ),
    "inventory_value": (
        "item_site_rows",
        "on_hand_quantity",
        "available_quantity",
        "reserved_quantity",
        "quality_hold_quantity",
        "inventory_value",
        "quality_hold_value",
        "average_unit_cost",
    ),
    "inventory_aging": (
        "item_site_rows",
        "inventory_value",
        "value_over_threshold",
        "phase_out_value",
        "quality_hold_value",
        "over_threshold_percent",
        "top_item",
        "top_item_value",
        "top_item_site",
        "top_item_site_value",
    ),
    "inventory_reserve": (
        "inventory_value",
        "aged_over_365_value",
        "phase_out_value",
        "aged_phase_out_overlap",
        "incremental_phase_out_overlay",
        "quality_hold_value",
        "incremental_quality_hold_overlay",
        "gross_reserve_exposure",
        "supported_recovery_value",
        "recommended_reserve",
        "reserve_percent",
        "top_item_site",
        "top_item_site_exposure",
    ),
    "inventory_turns": (
        "ltm_cost_of_goods_sold",
        "ending_inventory",
        "inventory_turns",
        "days_inventory",
    ),
    "inventory_transfer": (
        "transfer_out",
        "transfer_in",
        "net_difference",
        "transfer_out_references",
        "transfer_in_references",
        "unmatched_references",
    ),
    "forecast": (
        "forecast_units",
        "forecast_revenue",
        "forecast_months",
        "item_count",
        "average_price",
        "top_month",
        "top_month_revenue",
    ),
    "backlog": (
        "backlog_units",
        "backlog_value",
        "open_orders",
        "forecast_units",
        "forecast_revenue",
        "coverage",
        "overdue_value",
        "top_customer",
        "top_customer_value",
    ),
    "planning": (
        "planned_orders",
        "planned_quantity",
        "expedite_count",
        "defer_count",
        "increase_count",
        "decrease_count",
        "earliest_required_date",
    ),
    "capacity": (
        "work_centers",
        "net_available_hours",
        "planned_downtime_hours",
        "required_hours",
        "utilization",
        "overloaded_centers",
        "top_constraint",
        "top_constraint_utilization",
    ),
    "capacity_scenario": (
        "work_centers",
        "base_net_available_hours",
        "base_required_hours",
        "base_utilization",
        "base_overloaded_centers",
        "base_top_constraint",
        "base_top_constraint_utilization",
        "required_hours_change",
        "net_available_hours_change",
        "scenario_net_available_hours",
        "scenario_required_hours",
        "scenario_utilization",
        "scenario_overloaded_centers",
        "scenario_top_constraint",
        "scenario_top_constraint_utilization",
        "overloaded_center_reduction",
    ),
    "profitability": (
        "revenue",
        "standard_cost",
        "gross_profit",
        "gross_margin",
        "units",
        "group_count",
        "top_group",
        "top_group_revenue",
        "lowest_margin_group",
        "lowest_margin",
    ),
    "standard_cost": (
        "item_site_rows",
        "material_cost",
        "labor_cost",
        "variable_overhead",
        "fixed_overhead",
        "outside_processing",
        "total_standard_cost",
        "material_percent",
        "conversion_percent",
    ),
    "margin_sensitivity": (
        "baseline_revenue",
        "baseline_standard_cost",
        "baseline_material_cost",
        "baseline_gross_profit",
        "baseline_margin",
        "price_change",
        "cost_change",
        "scenario_revenue",
        "scenario_standard_cost",
        "scenario_gross_profit",
        "scenario_margin",
        "gross_profit_change",
        "required_price_change_for_gp",
        "required_price_change_for_margin",
    ),
    "ppv": (
        "invoice_count",
        "purchase_quantity",
        "actual_cost",
        "standard_cost",
        "purchase_price_variance",
        "ppv_rate",
        "top_vendor",
        "top_vendor_ppv",
    ),
    "ppv_causal": (
        "invoice_count",
        "purchase_quantity",
        "actual_cost",
        "standard_cost",
        "purchase_price_variance",
        "po_price_variance",
        "invoice_price_variance",
        "bridge_difference",
        "top_vendor",
        "top_vendor_ppv",
    ),
    "supplier": (
        "vendor_count",
        "receipt_count",
        "on_time_percent",
        "accepted_percent",
        "receipt_value",
        "worst_vendor",
        "worst_vendor_score",
    ),
    "commitments": (
        "open_po_lines",
        "open_quantity",
        "open_commitment",
        "overdue_commitment",
        "next_30_day_commitment",
        "top_vendor",
        "top_vendor_commitment",
    ),
    "invoice_holds": (
        "held_invoices",
        "held_amount",
        "quality_hold_invoices",
        "quality_hold_amount",
        "past_due_held_amount",
        "top_vendor",
        "top_vendor_held_amount",
    ),
    "payroll": (
        "employees",
        "regular_hours",
        "overtime_hours",
        "overtime_rate",
        "gross_pay",
        "employer_cost",
        "average_cost_per_employee",
        "top_department",
    ),
    "scrap": (
        "production_orders",
        "completed_units",
        "scrap_units",
        "yield_rate",
        "standard_scrap_cost",
        "top_family",
        "top_family_scrap_cost",
    ),
    "quality": (
        "open_quality_orders",
        "failed_quantity",
        "financial_exposure",
        "nonconformances",
        "blocked_inventory_value",
        "top_item",
        "top_item_exposure",
    ),
    "copq": (
        "standard_scrap_cost",
        "open_quality_order_exposure",
        "blocked_inventory_value",
        "known_copq_exposure",
        "scrap_percent_of_known_copq",
        "top_family",
        "top_family_scrap_cost",
        "top_open_quality_item",
        "top_open_quality_item_exposure",
    ),
    "maintenance": (
        "work_orders",
        "downtime_hours",
        "maintenance_cost",
        "unplanned_work_orders",
        "unplanned_downtime",
        "preventive_work_orders",
        "preventive_downtime",
        "unplanned_downtime_percent",
        "top_asset",
        "top_asset_downtime",
    ),
    "maintenance_capacity": (
        "work_orders",
        "maintenance_cost",
        "unplanned_work_orders",
        "unplanned_downtime",
        "preventive_work_orders",
        "preventive_downtime",
        "ytd_net_available_hours",
        "unplanned_capacity_loss_percent",
        "top_asset",
        "top_asset_downtime",
    ),
    "fixed_assets": (
        "asset_count",
        "gross_cost",
        "accumulated_depreciation",
        "net_book_value",
        "ytd_additions",
        "fully_depreciated_assets",
        "average_age_years",
    ),
    "npv": (
        "initial_investment",
        "annual_savings",
        "annual_maintenance",
        "net_annual_benefit",
        "discount_rate",
        "life_years",
        "npv",
        "irr",
        "payback_years",
        "profitability_index",
    ),
    "capital_portfolio": (
        "project_count",
        "capital_budget",
        "total_requested_investment",
        "total_project_npv",
        "selected_project_count",
        "selected_investment",
        "selected_npv",
        "remaining_budget",
        "priority_project",
        "priority_project_npv",
        "priority_project_profitability_index",
    ),
    "make_buy": (
        "annual_units",
        "make_variable_cost",
        "buy_unit_cost",
        "avoidable_fixed_cost",
        "one_time_cost",
        "make_total_cost",
        "buy_total_cost",
        "annual_savings",
        "payback_years",
        "payback_status",
        "three_year_undiscounted_savings",
        "three_year_npv",
    ),
    "cash_debt": (
        "cash",
        "total_debt",
        "undrawn_revolver",
        "total_liquidity",
        "current_debt",
        "weighted_interest_rate",
        "annual_cash_interest",
    ),
    "debt_profile": (
        "active_instrument_count",
        "revolver_count",
        "equipment_loan_count",
        "term_loan_count",
        "revolver_outstanding",
        "equipment_loan_outstanding",
        "term_loan_outstanding",
        "total_debt",
        "current_debt",
        "long_term_debt",
        "revolver_annual_interest",
        "equipment_loan_annual_interest",
        "term_loan_annual_interest",
        "annual_cash_interest",
        "weighted_interest_rate",
    ),
    "borrowing_base": (
        "gross_open_ar",
        "eligible_ar",
        "gross_inventory",
        "eligible_inventory",
        "ar_advance_rate",
        "inventory_advance_rate",
        "borrowing_base",
        "revolver_commitment",
        "revolver_outstanding",
        "availability",
    ),
    "cash_forecast": (
        "opening_cash",
        "customer_collections",
        "other_inflows",
        "supplier_payments",
        "payroll_and_benefits",
        "interest",
        "capital_expenditures",
        "other_outflows",
        "minimum_cash",
        "minimum_cash_week",
        "ending_cash",
        "maximum_revolver",
    ),
    "covenant": (
        "funded_debt",
        "ltm_adjusted_ebitda",
        "leverage",
        "maximum_leverage",
        "leverage_headroom",
        "cash",
        "undrawn_revolver",
        "liquidity",
        "tangible_net_worth",
        "minimum_tangible_net_worth",
        "compliant",
    ),
    "tax": (
        "pretax_income",
        "interest_expense",
        "book_depreciation",
        "tax_depreciation",
        "depreciation_temporary_difference",
        "permanent_items",
        "estimated_taxable_income",
        "federal_rate",
        "federal_tax",
        "state_rate",
        "state_tax",
        "current_tax",
        "deferred_tax_expense",
        "total_tax_provision",
        "effective_tax_rate",
    ),
    "interplant": (
        "transfer_out",
        "transfer_in",
        "net_difference",
        "markup_rate",
        "transfer_price_adjustment",
    ),
    "qoe": (
        "reported_ebitda",
        "production_variance_normalization",
        "quality_normalization",
        "maintenance_normalization",
        "normalized_ebitda",
        "quality_of_earnings_ratio",
    ),
    "dcf": (
        "base_cash_flow",
        "growth_rate",
        "discount_rate",
        "terminal_growth",
        "forecast_years",
        "enterprise_value",
        "net_debt",
        "equity_value",
    ),
    "roic": (
        "revenue",
        "operating_income",
        "tax_rate",
        "nopat",
        "invested_capital",
        "roic",
    ),
    "executive": (
        "period_start",
        "period_end",
        "status",
        "ytd_revenue",
        "ytd_gross_profit",
        "ytd_gross_margin",
        "ytd_operating_income",
        "inventory",
        "wip",
        "backlog",
        "cash",
        "total_debt",
        "total_liquidity",
        "ytd_ppv",
        "ytd_production_variance",
    ),
    "forecast_scenario": (
        "baseline_period_start",
        "baseline_period_end",
        "scenario_period",
        "status",
        "baseline_revenue",
        "baseline_gross_margin",
        "volume_change",
        "price_change",
        "material_cost_change",
        "scenario_revenue",
        "scenario_gross_profit",
        "scenario_gross_margin",
        "scenario_operating_income",
        "incremental_working_capital",
    ),
    "kpi_reconciliation": (
        "period_start",
        "period_end",
        "status",
        "revenue",
        "gross_profit",
        "operating_income",
        "inventory",
        "wip",
        "backlog",
        "cash",
        "debt",
        "reconciliation_failures",
    ),
    "balance_cash_bridge": (
        "as_of_date",
        "status",
        "cash",
        "accounts_receivable",
        "inventory",
        "work_in_process",
        "fixed_assets",
        "total_assets",
        "accounts_payable",
        "accrued_payroll",
        "other_liabilities",
        "debt",
        "total_liabilities",
        "equity",
        "operating_working_capital",
        "net_debt",
    ),
    "erp_draft_journal": (
        "reference",
        "status",
        "debits",
        "credits",
        "audit_action",
    ),
    "erp_posted_journal": (
        "reference",
        "status",
        "debits",
        "credits",
        "audit_actions",
    ),
    "erp_purchase_order": (
        "po_number",
        "status",
        "line_count",
        "order_value",
        "audit_action",
    ),
    "erp_production_order": (
        "order_number",
        "status",
        "item_id",
        "site_code",
        "order_quantity",
        "standard_unit_cost",
        "audit_actions",
    ),
    "erp_quality_hold": (
        "quality_order_id",
        "status",
        "placed_quantity",
        "estimated_financial_exposure",
        "audit_actions",
    ),
}


FAMILY_SOURCE_HINTS = {
    1: (
        "Use the ERP general ledger, production-cost, WIP, and reconciliation "
        "reports together with the Monthly Close Policy, production-order close "
        "policy, June close instructions, and the applicable close binder."
    ),
    2: (
        "Use ERP on-hand, valuation, aging, movement, and financial reports. "
        "Reconcile to the Inventory Aging and E&O workbook, inventory reserve "
        "policy, aftermarket transfer process, and June inventory-control email."
    ),
    3: (
        "Use the approved `2026-06-SOP` demand and planned-order version, current "
        "sales backlog, production orders, routings, work-center calendar, MRP "
        "action workbook, S&OP charter, and consensus demand plan."
    ),
    4: (
        "Use ERP invoice, shipment, customer, item/site standard-cost, BOM, and "
        "routing data. Reconcile to the customer/product profitability and "
        "standard-cost rollup workbooks and the pricing/margin review email."
    ),
    5: (
        "Use ERP purchase orders, receipts, vendor invoices, standard material "
        "cost, supplier performance, quality status, and open commitments. "
        "Reconcile to the PPV and supplier-scorecard workbooks and match policy."
    ),
    6: (
        "Use ERP payroll, production confirmations, scrap/yield, quality orders, "
        "nonconformances, maintenance work orders, and work-center data together "
        "with the plant scorecards and audit summaries."
    ),
    7: (
        "Use the ERP fixed-asset register, plant output, maintenance, payroll, "
        "capacity, and cash/debt data. Treat scenario assumptions as prospective "
        "and cite the Capital Committee Charter and capitalization policy."
    ),
    8: (
        "Use ERP cash, debt register, AR/AP aging, inventory, P&L, balance sheet, "
        "and fixed assets together with the treasury framework and executed debt "
        "summaries. Keep posted balances distinct from forecast assumptions."
    ),
    9: (
        "Use posted ERP financials, fixed assets, payroll, transfers, quality, "
        "maintenance, and debt. Clearly separate management scenarios and "
        "valuation assumptions from posted company results."
    ),
    10: (
        "Use the full ERP and the June pre-close source package. Reconcile "
        "every headline KPI to its authoritative module, label assumptions and "
        "cutoffs, and keep May-closed versus June-pre-close status explicit."
    ),
}


WORKFLOW_SOURCE_SETS: dict[str, tuple[str, ...]] = {
    "manufacturing-close": (
        "ERP production variance, WIP, P&L, and subledger reconciliation",
        "Finance/Monthly Close/2026-06 Close Binder.xlsx",
        "Corporate Memos/June 2026 Close Instructions.docx",
    ),
    "inventory-working-capital": (
        "ERP inventory on-hand, valuation, aging, and movement reports",
        "Supply Chain/Analyses/Inventory Aging and E&O Candidates.xlsx",
        "Accounting Policies/Inventory Reserve Policy.docx",
    ),
    "sop-reforecast": (
        "ERP forecast, backlog, planned-order, and capacity reports",
        "Supply Chain/Analyses/Consensus Demand Plan.xlsx",
        "Corporate Memos/S&OP Governance Charter.docx",
    ),
    "commercial-finance": (
        "ERP customer/product profitability and standard-cost reports",
        "Finance/Analyses/Product Family Profitability.xlsx",
        "Finance/Analyses/Standard Cost Rollup Review.xlsx",
    ),
    "procurement-finance": (
        "ERP PO, receipt, vendor-invoice, PPV, and supplier reports",
        "Supply Chain/Analyses/Purchase Price Variance.xlsx",
        "Accounting Policies/Three Way Match Policy.docx",
    ),
    "operations-finance": (
        "ERP payroll, production, quality, and maintenance reports",
        "Supply Chain/Analyses/Quality Exposure.xlsx",
        "Supply Chain/Analyses/Maintenance Downtime.xlsx",
    ),
    "capital-strategy": (
        "ERP fixed assets, capacity, maintenance, and cash/debt reports",
        "Finance/Analyses/Capital Expenditure Pipeline.xlsx",
        "Corporate Memos/Capital Committee Charter.docx",
    ),
    "treasury": (
        "ERP cash, debt, AR/AP, inventory, and P&L reports",
        "Finance/Analyses/Cash and Debt.xlsx",
        "Finance/Analyses/Debt Covenant Model.xlsx",
    ),
    "production-execution": (
        "ERP item/site master and site standard-cost records",
        "ERP active BOM and routing versions",
        "ERP production-order state and audit-event history",
    ),
    "tax-corpdev": (
        "ERP posted financials, assets, quality, maintenance, and debt",
        "Finance/Analyses/Tax Provision Inputs.xlsx",
        "Tax/2026 Q2 Provision Calendar.pdf",
    ),
    "executive-synthesis": (
        "ERP authoritative June 30 and YTD reports across all modules",
        "Presentations/2026-06 Preclose Operating Review.pptx",
        "Communications/2026-06-29 - June 30 executive package.eml",
    ),
}

BUNDLE_SOURCE_SETS: dict[str, tuple[str, ...]] = {
    "profitability": (
        "ERP customer/product profitability reports",
        "Finance/Analyses/Customer Profitability.xlsx",
        "Finance/Analyses/Product Family Profitability.xlsx",
        "Communications/2026-06-15 - Product family margin review.eml",
    ),
    "margin_sensitivity": (
        "ERP product-family profitability and standard-cost reports",
        "Finance/Analyses/Product Family Profitability.xlsx",
        "Finance/Analyses/Standard Cost Rollup Review.xlsx",
        "Communications/2026-06-15 - Product family margin review.eml",
    ),
    "ppv": (
        "ERP purchase-price-variance report",
        "Supply Chain/Analyses/Purchase Price Variance.xlsx",
        "Communications/2026-06-09 - June PPV actions.eml",
        "Accounting Policies/Purchase Price Variance Policy.docx",
    ),
    "ppv_causal": (
        "ERP vendor-invoice, PO-price, and standard-material-cost reports",
        "Supply Chain/Analyses/Purchase Price Variance.xlsx",
        "Communications/2026-06-09 - June PPV actions.eml",
        "Accounting Policies/Purchase Price Variance Policy.docx",
    ),
    "supplier": (
        "ERP supplier-performance and receipt reports",
        "Supply Chain/Analyses/Supplier Scorecard.xlsx",
        "Corporate Memos/Supplier Risk Review.docx",
        "Communications/2026-02-03 - Supplier delivery recovery.eml",
    ),
    "commitments": (
        "ERP open purchase-order and commitment reports",
        "Supply Chain/Analyses/Open Purchase Commitments.xlsx",
        "Accounting Policies/Three Way Match Policy.docx",
    ),
    "invoice_holds": (
        "ERP open vendor-invoice and match-status reports",
        "Supply Chain/Analyses/Vendor Invoice Holds.xlsx",
        "Communications/2026-04-22 - Quality-hold invoices.eml",
        "Accounting Policies/Three Way Match Policy.docx",
    ),
    "payroll": (
        "ERP posted payroll and employee reports",
        "Finance/Analyses/Payroll and Overtime.xlsx",
        "Operations/Plant Scorecards",
    ),
    "scrap": (
        "ERP production completion and variance reports",
        "Finance/Analyses/Production Variance Bridge.xlsx",
        "Operations/Plant Scorecards",
    ),
    "quality": (
        "ERP quality-order, nonconformance, and inventory-hold reports",
        "Supply Chain/Analyses/Quality Exposure.xlsx",
        "Presentations/Supplier and Quality Risk Review.pptx",
    ),
    "copq": (
        "ERP production-variance, quality-order, and inventory-hold reports",
        "Finance/Analyses/Production Variance Bridge.xlsx",
        "Supply Chain/Analyses/Quality Exposure.xlsx",
        "Presentations/Supplier and Quality Risk Review.pptx",
    ),
    "maintenance": (
        "ERP maintenance work-order and asset reports",
        "Supply Chain/Analyses/Maintenance Downtime.xlsx",
        "Communications/2026-04-05 - Preventive maintenance backlog.eml",
    ),
    "maintenance_capacity": (
        "ERP maintenance work-order and work-center calendar reports",
        "Supply Chain/Analyses/Maintenance Downtime.xlsx",
        "Supply Chain/Analyses/Work Center Capacity.xlsx",
        "Communications/2026-04-05 - Preventive maintenance backlog.eml",
    ),
    "plant_output": (
        "ERP production completion and variance reports by site",
        "Operations/Plant Scorecards",
        "Finance/Analyses/Production Variance Bridge.xlsx",
    ),
    "fixed_assets": (
        "ERP active fixed-asset register",
        "Finance/Analyses/Fixed Asset Register.xlsx",
        "Accounting Policies/Capitalization Policy.docx",
    ),
    "npv": (
        "ERP capacity, maintenance, and fixed-asset reports",
        "Finance/Analyses/Capital Expenditure Pipeline.xlsx",
        "Corporate Memos/Capital Committee Charter.docx",
        "Accounting Policies/Capitalization Policy.docx",
        "Communications/2026-06-13 - Capacity capital requests.eml",
    ),
    "capital_portfolio": (
        "ERP capacity, maintenance, and fixed-asset reports",
        "Finance/Analyses/Capital Expenditure Pipeline.xlsx",
        "Corporate Memos/Capital Committee Charter.docx",
        "Communications/2026-06-13 - Capacity capital requests.eml",
    ),
    "make_buy": (
        "ERP production-cost and capacity reports",
        "Finance/Analyses/Capital Expenditure Pipeline.xlsx",
        "Corporate Memos/Capital Committee Charter.docx",
        "Accounting Policies/Capitalization Policy.docx",
        "Communications/2026-06-13 - Capacity capital requests.eml",
    ),
    "cash_debt": (
        "ERP posted cash and active debt-register reports",
        "Finance/Analyses/Cash and Debt.xlsx",
        "Corporate Memos/Treasury and Liquidity Framework.docx",
        "Treasury/2022 Sponsor Term Loan Summary.pdf",
        "Treasury/2025 ABL Facility Summary.pdf",
        "Treasury/Equipment Facility Summary.pdf",
    ),
    "debt_profile": (
        "ERP active debt-register report",
        "Finance/Analyses/Cash and Debt.xlsx",
        "Corporate Memos/Treasury and Liquidity Framework.docx",
        "Treasury/2022 Sponsor Term Loan Summary.pdf",
        "Treasury/2025 ABL Facility Summary.pdf",
        "Treasury/Equipment Facility Summary.pdf",
        "Communications/2026-06-23 - Debt classification check.eml",
    ),
    "borrowing_base": (
        "ERP open-AR, inventory, reserve, and revolver reports",
        "Finance/Analyses/Cash and Debt.xlsx",
        "Corporate Memos/Treasury and Liquidity Framework.docx",
        "Treasury/2022 Sponsor Term Loan Summary.pdf",
        "Treasury/2025 ABL Facility Summary.pdf",
        "Treasury/Equipment Facility Summary.pdf",
        "Communications/2026-02-12 - ABL availability package.eml",
    ),
    "cash_forecast": (
        "ERP posted cash, collections, disbursement, and debt reports",
        "Finance/Analyses/Cash and Debt.xlsx",
        "Corporate Memos/Treasury and Liquidity Framework.docx",
        "Treasury/2022 Sponsor Term Loan Summary.pdf",
        "Treasury/2025 ABL Facility Summary.pdf",
        "Treasury/Equipment Facility Summary.pdf",
        "Communications/2026-05-27 - Liquidity downside case.eml",
    ),
    "covenant": (
        "ERP LTM P&L, balance-sheet, and active debt reports",
        "Finance/Analyses/Debt Covenant Model.xlsx",
        "Finance/Analyses/Cash and Debt.xlsx",
        "Corporate Memos/Treasury and Liquidity Framework.docx",
        "Treasury/2022 Sponsor Term Loan Summary.pdf",
        "Treasury/2025 ABL Facility Summary.pdf",
        "Treasury/Equipment Facility Summary.pdf",
    ),
    "tax": (
        "ERP posted book-income, debt-interest, fixed-asset, and payroll reports",
        "Finance/Analyses/Tax Provision Inputs.xlsx",
        "Tax/2026 Q2 Provision Calendar.pdf",
        "Communications/2026-06-11 - Q2 provision data.eml",
    ),
    "interplant": (
        "ERP inventory-transfer movement ledger",
        "Corporate Memos/Aftermarket Transfer Process.docx",
        "Communications/2026-06-07 - Aftermarket transfer reconciliation.eml",
    ),
    "qoe": (
        "ERP LTM P&L, production-variance, open-quality, and corrective-maintenance reports",
        "Finance/Analyses/Debt Covenant Model.xlsx",
        "Finance/Analyses/Production Variance Bridge.xlsx",
        "Supply Chain/Analyses/Quality Exposure.xlsx",
        "Supply Chain/Analyses/Maintenance Downtime.xlsx",
    ),
    "dcf": (
        "ERP LTM P&L, quality, maintenance, fixed-asset, cash, and debt reports",
        "Finance/Analyses/Debt Covenant Model.xlsx",
        "Finance/Analyses/Fixed Asset Register.xlsx",
        "Finance/Analyses/Cash and Debt.xlsx",
        "Management valuation assumptions: cash tax, growth, discount rate, terminal growth, and forecast years",
    ),
    "roic": (
        "ERP LTM P&L and June 30 balance-sheet reports",
        "Finance/Analyses/Debt Covenant Model.xlsx",
        "Finance/Analyses/Fixed Asset Register.xlsx",
        "Finance/Analyses/Working Capital Dashboard.xlsx",
    ),
    "erp_purchase_order": (
        "ERP active vendor, item/site, buyer, and recent PO master data",
        "Supply Chain/Analyses/Open Purchase Commitments.xlsx",
        "Accounting Policies/Three Way Match Policy.docx",
    ),
}


TASK_SOURCE_SETS: dict[str, tuple[str, ...]] = {
    "task_093": (
        "Presentations/2026-06 Preclose Operating Review.pptx",
        "ERP independent YTD P&L, inventory, WIP, backlog, cash, and active debt reports",
        "Communications/2026-06-29 - June 30 executive package.eml",
    ),
    "task_096": (
        "ERP authoritative June 30 and YTD reports across all modules",
        "Finance/Monthly Close/2026-06 Close Binder.xlsx",
        "Corporate Memos/June 2026 Close Instructions.docx",
        "Presentations/2026-06 Preclose Operating Review.pptx",
        "Communications/2026-06-29 - June 30 executive package.eml",
    ),
}


def authoritative_sources(
    blueprint: TaskBlueprint,
) -> tuple[str, ...]:
    if blueprint.task_id in TASK_SOURCE_SETS:
        return TASK_SOURCE_SETS[blueprint.task_id]
    return BUNDLE_SOURCE_SETS.get(
        blueprint.bundle,
        WORKFLOW_SOURCE_SETS[blueprint.workflow],
    )


WORKFLOW_SOURCE_HINT_OVERRIDES = {
    "production-execution": (
        "Use the authorized ERP item/site setup, site standard cost, active BOM "
        "and routing versions, production-order state, and audit history. "
        "Confirm that the order copies the approved structure and that only "
        "the required status transitions occur."
    ),
}


def _slug(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")


def _blueprint(
    number: int,
    title: str,
    workflow: str,
    output_mode: str,
    bundle: str,
    parameters: dict[str, Any] | None = None,
) -> TaskBlueprint:
    artifact_title = re.sub(
        r"^(?:build|write|create)\s+the\s+",
        "",
        title,
        flags=re.IGNORECASE,
    )
    slug = _slug(artifact_title)
    suffix = {
        "spreadsheet": ".xlsx",
        "document": ".docx",
        "presentation": ".pptx",
    }.get(output_mode)
    target = f"Deliverables/{slug}{suffix}" if suffix else None
    family = (number - 1) // 10 + 1
    return TaskBlueprint(
        task_id=f"task_{number:03d}",
        title=title,
        workflow=workflow,
        output_mode=output_mode,
        bundle=bundle,
        parameters=parameters or {},
        source_hint=WORKFLOW_SOURCE_HINT_OVERRIDES.get(
            workflow, FAMILY_SOURCE_HINTS[family]
        ),
        slug=slug,
        target_path=target,
    )


TASK_BLUEPRINTS: tuple[TaskBlueprint, ...] = (
    # 001-010: manufacturing close (6 console, 2 spreadsheet, 1 document, 1 ERP)
    _blueprint(1, "June cross-ledger manufacturing close control", "manufacturing-close", "console", "close_control"),
    _blueprint(2, "June pre-close profit and margin bridge", "manufacturing-close", "console", "pnl", {"date_from": "2026-06-01", "date_to": "2026-06-30"}),
    _blueprint(3, "YTD settled production variance bridge", "manufacturing-close", "console", "production_variance", {"date_from": "2026-01-01", "date_to": "2026-06-30"}),
    _blueprint(4, "June open production WIP concentration", "manufacturing-close", "console", "wip"),
    _blueprint(5, "Grand Rapids June output and variance", "manufacturing-close", "console", "plant_output", {"site_code": "100", "date_from": "2026-06-01", "date_to": "2026-06-30"}),
    _blueprint(6, "Dayton June output and variance", "manufacturing-close", "console", "plant_output", {"site_code": "200", "date_from": "2026-06-01", "date_to": "2026-06-30"}),
    _blueprint(7, "Build the June manufacturing close control workbook", "manufacturing-close", "spreadsheet", "close_control"),
    _blueprint(8, "Build the YTD production variance bridge workbook", "manufacturing-close", "spreadsheet", "production_variance", {"date_from": "2026-01-01", "date_to": "2026-06-30"}),
    _blueprint(9, "Write the June close risk memorandum", "manufacturing-close", "document", "executive"),
    _blueprint(10, "Create the approved June scrap reclass draft", "manufacturing-close", "erp", "erp_draft_journal", {"posting_date": "2026-06-30", "reference": "PH-JE-010", "memo": "Reclass June abnormal scrap for management review", "lines": [{"account_code": "5110", "site_code": "100", "debit": 125000.0}, {"account_code": "5100", "site_code": "100", "credit": 125000.0}]}),

    # 011-020: inventory and working capital (6 console, 3 spreadsheet, 1 document)
    _blueprint(11, "Inventory valuation across all sites", "inventory-working-capital", "console", "inventory_value"),
    _blueprint(12, "Fort Worth aftermarket availability and holds", "inventory-working-capital", "console", "inventory_value", {"site_code": "400", "product_family": "AFT"}),
    _blueprint(13, "Inventory aging over 180 days", "inventory-working-capital", "console", "inventory_aging", {"days": 180}),
    _blueprint(14, "Phase-out inventory exposure", "inventory-working-capital", "console", "inventory_aging", {"days": 90, "lifecycle_status": "Phase Out"}),
    _blueprint(15, "Company inventory turns and days on hand", "inventory-working-capital", "console", "inventory_turns"),
    _blueprint(16, "Aftermarket plant-to-DC transfer reconciliation", "inventory-working-capital", "console", "inventory_transfer", {"product_family": "AFT"}),
    _blueprint(17, "Build the inventory aging and E&O workbook", "inventory-working-capital", "spreadsheet", "inventory_aging", {"days": 180}),
    _blueprint(18, "Build the inventory reserve waterfall", "inventory-working-capital", "spreadsheet", "inventory_reserve", {"days": 365, "supported_recovery_value": 0.0}),
    _blueprint(19, "Build the inventory turns dashboard", "inventory-working-capital", "spreadsheet", "inventory_turns"),
    _blueprint(20, "Write the inventory disposition recommendation", "inventory-working-capital", "document", "inventory_aging", {"days": 180}),

    # 021-030: S&OP and reforecast (5 console, 3 spreadsheet, 1 presentation, 1 ERP)
    _blueprint(21, "Second-half consensus demand plan", "sop-reforecast", "console", "forecast"),
    _blueprint(22, "DRV backlog coverage of consensus demand", "sop-reforecast", "console", "backlog", {"product_family": "DRV"}),
    _blueprint(23, "Company MRP action-message profile", "sop-reforecast", "console", "planning"),
    _blueprint(24, "Dayton Q3 work-center capacity", "sop-reforecast", "console", "capacity", {"site_code": "200", "date_from": "2026-07-01", "date_to": "2026-09-30"}),
    _blueprint(25, "Greenville Q3 work-center capacity", "sop-reforecast", "console", "capacity", {"site_code": "300", "date_from": "2026-07-01", "date_to": "2026-09-30"}),
    _blueprint(26, "Build the demand-to-backlog bridge", "sop-reforecast", "spreadsheet", "backlog"),
    _blueprint(27, "Build the Q3 constraint scenario model", "sop-reforecast", "spreadsheet", "capacity_scenario", {"site_code": "200", "date_from": "2026-07-01", "date_to": "2026-09-30", "required_hours_change": -0.10, "net_available_hours_change": 0.10}),
    _blueprint(28, "Build the MRP executive action workbook", "sop-reforecast", "spreadsheet", "planning"),
    _blueprint(29, "Create the executive S&OP decision deck", "sop-reforecast", "presentation", "forecast", {"capacity_recovery_plan_days": 10}),
    _blueprint(30, "Create the DRV production order from the approved plan", "sop-reforecast", "erp", "erp_production_order", {"item_id": "DRV-0001", "site_code": "100", "order_quantity": 5.0, "scheduled_start": "2026-07-01", "scheduled_finish": "2026-07-11", "source_reference": "PLAN-000790"}),

    # 031-040: product and customer economics (5 console, 3 spreadsheet, 1 document, 1 presentation)
    _blueprint(31, "YTD product-family profitability", "commercial-finance", "console", "profitability", {"group_by": "product_family"}),
    _blueprint(32, "YTD customer profitability and concentration", "commercial-finance", "console", "profitability", {"group_by": "customer"}),
    _blueprint(33, "YTD regional profitability", "commercial-finance", "console", "profitability", {"group_by": "region"}),
    _blueprint(34, "Grand Rapids DRV standard-cost composition", "commercial-finance", "console", "standard_cost", {"site_code": "100", "product_family": "DRV"}),
    _blueprint(35, "DRV price and material-cost sensitivity", "commercial-finance", "console", "margin_sensitivity", {"product_family": "DRV", "price_change": 0.02, "cost_change": 0.04}),
    _blueprint(36, "Build the customer profitability workbook", "commercial-finance", "spreadsheet", "profitability", {"group_by": "customer"}),
    _blueprint(37, "Build the product-family margin bridge", "commercial-finance", "spreadsheet", "profitability", {"group_by": "product_family"}),
    _blueprint(38, "Build the CTL pricing sensitivity model", "commercial-finance", "spreadsheet", "margin_sensitivity", {"product_family": "CTL", "price_change": 0.03, "cost_change": 0.05}),
    _blueprint(39, "Write the DRV pricing recommendation", "commercial-finance", "document", "margin_sensitivity", {"product_family": "DRV", "price_change": 0.02, "cost_change": 0.04, "decision_owner": "Commercial Finance Director", "decision_date": "2026-07-10"}),
    _blueprint(40, "Create the commercial margin review deck", "commercial-finance", "presentation", "profitability", {"group_by": "product_family"}),

    # 041-050: procurement and supplier finance (6 console, 2 spreadsheet, 1 document, 1 ERP)
    _blueprint(41, "YTD purchase-price variance summary and top vendor", "procurement-finance", "console", "ppv"),
    _blueprint(42, "Packaging supplier delivery and quality", "procurement-finance", "console", "supplier", {"category": "Packaging"}),
    _blueprint(43, "Companywide open purchase commitments", "procurement-finance", "console", "commitments"),
    _blueprint(44, "Quality-hold vendor invoice exposure", "procurement-finance", "console", "invoice_holds"),
    _blueprint(45, "Dayton open purchase commitments", "procurement-finance", "console", "commitments", {"site_code": "200"}),
    _blueprint(46, "Metals supplier performance", "procurement-finance", "console", "supplier", {"category": "Metals"}),
    _blueprint(47, "Build the supplier recovery scorecard", "procurement-finance", "spreadsheet", "supplier", {"recovery_score_threshold": 0.8, "critical_score_threshold": 0.65, "action_due_date": "2026-07-17"}),
    _blueprint(48, "Build the YTD PPV causal bridge", "procurement-finance", "spreadsheet", "ppv_causal"),
    _blueprint(49, "Write the supplier-risk council memorandum", "procurement-finance", "document", "supplier", {"decision_owner": "VP Supply Chain", "decision_date": "2026-07-10"}),
    _blueprint(50, "Create the approved raw-material purchase order", "procurement-finance", "erp", "erp_purchase_order", {"vendor_id": "VEND-0014", "site_code": "100", "order_date": "2026-06-30", "expected_date": "2026-07-20", "buyer_employee_id": "EMP-0005", "lines": [{"item_id": "RM-0001", "quantity": 240.0, "unit_price": 12.76}]}),

    # 051-060: labor, capacity, quality, maintenance (5 console, 3 spreadsheet, 1 presentation, 1 ERP)
    _blueprint(51, "Companywide YTD payroll and overtime", "operations-finance", "console", "payroll"),
    _blueprint(52, "Dayton payroll and overtime", "operations-finance", "console", "payroll", {"site_code": "200"}),
    _blueprint(53, "YTD scrap and yield summary and top product family", "operations-finance", "console", "scrap"),
    _blueprint(54, "Open quality and nonconformance exposure", "operations-finance", "console", "quality"),
    _blueprint(55, "YTD maintenance downtime and cost", "operations-finance", "console", "maintenance"),
    _blueprint(56, "Build the labor and overtime workbook", "operations-finance", "spreadsheet", "payroll"),
    _blueprint(57, "Build the cost-of-poor-quality workbook", "operations-finance", "spreadsheet", "copq"),
    _blueprint(58, "Build the maintenance-capacity workbook", "operations-finance", "spreadsheet", "maintenance_capacity"),
    _blueprint(59, "Create the multi-plant operating review deck", "operations-finance", "presentation", "plant_output", {"date_from": "2026-01-01", "date_to": "2026-06-30", "decision_owner": "VP Operations", "action_due_date": "2026-07-17"}),
    _blueprint(60, "Place the approved RM-0001 inventory quality hold", "operations-finance", "erp", "erp_quality_hold", {"item_id": "RM-0001", "site_code": "100", "quantity": 12.0, "reason": "Supplier lot inspection pending"}),

    # 061-070: capex and operating strategy (5 console, 3 spreadsheet, 1 document, 1 ERP)
    _blueprint(61, "Fixed-asset portfolio and YTD additions", "capital-strategy", "console", "fixed_assets"),
    _blueprint(62, "Dayton fixed-asset replacement profile", "capital-strategy", "console", "fixed_assets", {"site_code": "200"}),
    _blueprint(63, "Dayton machining automation NPV", "capital-strategy", "console", "npv", {"project_id": "CAP-DAY-001", "project_name": "Dayton machining automation", "site_name": "Dayton", "site_code": "200", "initial_investment": 4200000.0, "annual_savings": 1180000.0, "annual_maintenance": 135000.0, "discount_rate": 0.105, "life_years": 7}),
    _blueprint(64, "Cabinet assembly make-versus-buy", "capital-strategy", "console", "make_buy", {"annual_units": 4200.0, "make_variable_cost": 2380.0, "buy_unit_cost": 2615.0, "avoidable_fixed_cost": 480000.0, "one_time_cost": 725000.0, "discount_rate": 0.105}),
    _blueprint(65, "Greenville test-cell automation NPV", "capital-strategy", "console", "npv", {"project_id": "CAP-GRV-001", "project_name": "Greenville test-cell automation", "site_name": "Greenville", "site_code": "300", "initial_investment": 2850000.0, "annual_savings": 830000.0, "annual_maintenance": 95000.0, "discount_rate": 0.105, "life_years": 6}),
    _blueprint(66, "Build the capital-project portfolio workbook", "capital-strategy", "spreadsheet", "capital_portfolio", {"capital_budget": 5000000.0, "projects": [{"project_id": "CAP-DAY-001", "project_name": "Dayton machining automation", "site_name": "Dayton", "site_code": "200", "initial_investment": 4200000.0, "annual_savings": 1180000.0, "annual_maintenance": 135000.0, "discount_rate": 0.105, "life_years": 7}, {"project_id": "CAP-GRV-001", "project_name": "Greenville test-cell automation", "site_name": "Greenville", "site_code": "300", "initial_investment": 2850000.0, "annual_savings": 830000.0, "annual_maintenance": 95000.0, "discount_rate": 0.105, "life_years": 6}]}),
    _blueprint(67, "Build the Dayton automation business case", "capital-strategy", "spreadsheet", "npv", {"project_id": "CAP-DAY-001", "project_name": "Dayton machining automation", "site_name": "Dayton", "site_code": "200", "initial_investment": 4200000.0, "annual_savings": 1180000.0, "annual_maintenance": 135000.0, "discount_rate": 0.105, "life_years": 7, "decision_owner": "VP Operations", "implementation_gate_date": "2026-07-31"}),
    _blueprint(68, "Build the cabinet make-buy model", "capital-strategy", "spreadsheet", "make_buy", {"annual_units": 4200.0, "make_variable_cost": 2380.0, "buy_unit_cost": 2615.0, "avoidable_fixed_cost": 480000.0, "one_time_cost": 725000.0, "discount_rate": 0.105}),
    _blueprint(69, "Write the Dayton automation capital committee decision memorandum", "capital-strategy", "document", "npv", {"site_name": "Dayton", "site_code": "200", "initial_investment": 4200000.0, "annual_savings": 1180000.0, "annual_maintenance": 135000.0, "discount_rate": 0.105, "life_years": 7}),
    _blueprint(70, "Create and post the approved depreciation reclass", "capital-strategy", "erp", "erp_posted_journal", {"posting_date": "2026-06-30", "reference": "PH-JE-070", "memo": "Approved June depreciation site reclass", "lines": [{"account_code": "6400", "site_code": "200", "debit": 85000.0}, {"account_code": "6400", "site_code": "100", "credit": 85000.0}]}),

    # 071-079: treasury and balance sheet; 080: production execution
    _blueprint(71, "June cash, debt, and total liquidity", "treasury", "console", "cash_debt"),
    _blueprint(72, "Debt classification and cash-interest run rate", "treasury", "console", "debt_profile"),
    _blueprint(73, "ABL borrowing-base availability", "treasury", "console", "borrowing_base", {"ar_advance_rate": 0.85, "inventory_advance_rate": 0.55}),
    _blueprint(74, "Base 13-week cash and revolver forecast", "treasury", "console", "cash_forecast"),
    _blueprint(75, "June leverage and tangible-net-worth covenant", "treasury", "console", "covenant", {"maximum_leverage": 4.5, "minimum_tangible_net_worth": 85000000.0}),
    _blueprint(76, "Build the 13-week cash forecast workbook", "treasury", "spreadsheet", "cash_forecast"),
    _blueprint(77, "Build the covenant and liquidity workbook", "treasury", "spreadsheet", "covenant", {"maximum_leverage": 4.5, "minimum_tangible_net_worth": 85000000.0}),
    _blueprint(78, "Write the liquidity and covenant memorandum", "treasury", "document", "covenant", {"maximum_leverage": 4.5, "minimum_tangible_net_worth": 85000000.0}),
    _blueprint(79, "Create the Q2 lender update deck", "treasury", "presentation", "covenant", {"maximum_leverage": 4.5, "minimum_tangible_net_worth": 85000000.0}),
    _blueprint(80, "Create and release the approved AFT production order", "production-execution", "erp", "erp_production_order", {"item_id": "AFT-0001", "site_code": "100", "order_quantity": 40.0, "scheduled_start": "2026-07-06", "scheduled_finish": "2026-07-15", "source_reference": "AFT-REPLENISHMENT", "final_status": "Released"}),

    # 081-090: tax, corp dev, portfolio strategy (5 console, 2 spreadsheet, 2 document, 1 presentation)
    _blueprint(81, "Estimated 2026 first-half tax provision", "tax-corpdev", "console", "tax", {"permanent_items": 425000.0, "tax_depreciation": 15750000.0, "federal_rate": 0.21, "state_rate": 0.045}),
    _blueprint(82, "Aftermarket transfer-pricing adjustment", "tax-corpdev", "console", "interplant", {"markup_rate": 0.06}),
    _blueprint(83, "Manufacturing quality-of-earnings normalization", "tax-corpdev", "console", "qoe"),
    _blueprint(84, "Pinehaven standalone discounted cash flow", "tax-corpdev", "console", "dcf", {"cash_tax_rate": 0.255, "growth_rate": 0.04, "discount_rate": 0.11, "terminal_growth": 0.025, "forecast_years": 5}),
    _blueprint(85, "Pinehaven operating return on invested capital", "tax-corpdev", "console", "roic", {"tax_rate": 0.255}),
    _blueprint(86, "Build the standalone DCF workbook", "tax-corpdev", "spreadsheet", "dcf", {"cash_tax_rate": 0.255, "growth_rate": 0.04, "discount_rate": 0.11, "terminal_growth": 0.025, "forecast_years": 5}),
    _blueprint(87, "Build the quality-of-earnings bridge", "tax-corpdev", "spreadsheet", "qoe"),
    _blueprint(88, "Write the first-half tax provision memorandum", "tax-corpdev", "document", "tax", {"permanent_items": 425000.0, "tax_depreciation": 15750000.0, "federal_rate": 0.21, "state_rate": 0.045}),
    _blueprint(89, "Write the standalone valuation committee memorandum", "tax-corpdev", "document", "dcf", {"cash_tax_rate": 0.255, "growth_rate": 0.04, "discount_rate": 0.11, "terminal_growth": 0.025, "forecast_years": 5}),
    _blueprint(90, "Create the valuation committee deck", "tax-corpdev", "presentation", "dcf", {"cash_tax_rate": 0.255, "growth_rate": 0.04, "discount_rate": 0.11, "terminal_growth": 0.025, "forecast_years": 5}),

    # 091-100: executive synthesis (4 console, 1 spreadsheet, 2 document, 1 presentation, 2 ERP)
    _blueprint(91, "Integrated June executive scorecard", "executive-synthesis", "console", "executive"),
    _blueprint(92, "Second-half downside operating forecast", "executive-synthesis", "console", "forecast_scenario", {"baseline_start": "2025-07-01", "baseline_end": "2025-12-31", "volume_change": -0.08, "price_change": -0.01, "material_cost_change": 0.05}),
    _blueprint(93, "Board and lender KPI reconciliation", "executive-synthesis", "console", "kpi_reconciliation"),
    _blueprint(94, "Balance-sheet and net-debt bridge", "executive-synthesis", "console", "balance_cash_bridge"),
    _blueprint(95, "Build the FY27 integrated operating plan", "executive-synthesis", "spreadsheet", "forecast_scenario", {"baseline_start": "2025-01-01", "baseline_end": "2025-12-31", "volume_change": 0.06, "price_change": 0.025, "material_cost_change": 0.03}),
    _blueprint(96, "Write the CFO June close decision note", "executive-synthesis", "document", "executive"),
    _blueprint(97, "Write the FY27 board priorities memorandum", "executive-synthesis", "document", "forecast_scenario", {"baseline_start": "2025-01-01", "baseline_end": "2025-12-31", "volume_change": 0.06, "price_change": 0.025, "material_cost_change": 0.03}),
    _blueprint(98, "Create the FY26 outlook and FY27 priorities board deck", "executive-synthesis", "presentation", "executive"),
    _blueprint(99, "Create, release, and start the approved CTL order", "executive-synthesis", "erp", "erp_production_order", {"item_id": "CTL-0001", "site_code": "300", "order_quantity": 18.0, "scheduled_start": "2026-07-06", "scheduled_finish": "2026-07-24", "source_reference": "BOARD-PRIORITY", "final_status": "Started"}),
    _blueprint(100, "Place and release the final-inspection inventory hold", "executive-synthesis", "erp", "erp_quality_hold", {"item_id": "RM-0002", "site_code": "100", "quantity": 8.0, "reason": "Final inspection sample", "release": True, "disposition": "Released after final inspection"}),
)


def validate_blueprints() -> None:
    if len(TASK_BLUEPRINTS) != 100:
        raise AssertionError(f"Expected 100 tasks; found {len(TASK_BLUEPRINTS)}")
    ids = [task.task_id for task in TASK_BLUEPRINTS]
    if ids != [f"task_{index:03d}" for index in range(1, 101)]:
        raise AssertionError("Task IDs are not complete and sequential")
    if len({task.slug for task in TASK_BLUEPRINTS}) != 100:
        raise AssertionError("Task slugs are not unique")
    counts = {
        mode: sum(task.output_mode == mode for task in TASK_BLUEPRINTS)
        for mode in ("console", "spreadsheet", "document", "presentation", "erp")
    }
    expected = {
        "console": 52,
        "spreadsheet": 24,
        "document": 10,
        "presentation": 6,
        "erp": 8,
    }
    if counts != expected:
        raise AssertionError(f"Output-mode mix mismatch: {counts}")
    unknown = sorted(
        {task.bundle for task in TASK_BLUEPRINTS} - set(BUNDLE_KEYS)
    )
    if unknown:
        raise AssertionError(f"Unknown task bundles: {unknown}")


validate_blueprints()
