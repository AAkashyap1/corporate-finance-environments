from __future__ import annotations

import json
from dataclasses import dataclass

from task_blueprints import (
    BUNDLE_KEYS,
    TASK_BLUEPRINTS,
    TaskBlueprint,
    authoritative_sources,
)
from task_prompts import TASK_PROMPTS


@dataclass(frozen=True)
class TaskSpec:
    task_id: str
    slug: str
    title: str
    workflow: str
    output_mode: str
    prompt: str
    difficulty: str
    target_path: str | None


PERIOD_GUIDANCE = {
    "close_control": (
        "Use the June 30, 2026 pre-close snapshot and reconcile every control "
        "account to its authoritative ERP subledger. Production variance and "
        "the GL account 5110 control are June-only activity from 2026-06-01 "
        "through 2026-06-30, not YTD activity. If you quantify negative "
        "WIP residuals, use the WIP report's exact negative-residual scope "
        "count and signed value, which aggregate atomic open-order values "
        "before rounding once; do not sum rounded detail rows or hardcode an "
        "exception amount."
    ),
    "pnl": "Use posted June activity from 2026-06-01 through 2026-06-30. June remains pre-close; do not call it final.",
    "production_variance": (
        "Use settled production orders in the specified period. Preserve "
        "signed favorable/unfavorable variance convention. Use the production "
        "variance report's exact scope totals, which aggregate atomic order "
        "values before rounding once; do not sum the displayed rounded "
        "site/family rows."
    ),
    "wip": (
        "Use all production orders not Ended as of 2026-06-30. WIP is cost "
        "to date less standard value of reported output. Use the WIP report's "
        "exact scope totals, which aggregate atomic order values before "
        "rounding once; do not sum displayed rounded order rows."
    ),
    "plant_output": "Use production orders whose actual or scheduled finish falls in the specified period and the requested plant.",
    "inventory_value": "Use the June 30 perpetual on-hand dimensions. Available quantity is on hand less reserved and quality-held quantity.",
    "inventory_aging": (
        "Age positive on-hand value from the last movement date through "
        "2026-06-30 and apply the supplied lifecycle/site filter to the entire "
        "in-scope population. `item_site_rows`, `inventory_value`, and the "
        "`over_threshold_percent` denominator are that full filtered "
        "positive-on-hand population; they are not the aged subset. "
        "`value_over_threshold`, the top item, and the top item/site use the "
        "strict `age_days > days` test. `top_item` aggregates across sites; "
        "`top_item_site` ranks one item/site exposure in accordance with the "
        "reserve policy. `phase_out_value` is the full positive-on-hand Phase "
        "Out population within the same other filters, independent of age, "
        "without double-counting it into the aged measure."
    ),
    "inventory_reserve": (
        "Build an item/site-specific candidate reserve-exposure waterfall at "
        "2026-06-30. Start with full positive on-hand inventory value. The "
        "aged layer is full item/site value with strict `age_days > 365`; the "
        "Phase Out layer contributes only item/site value not already in the "
        "aged layer; and the quality-hold layer contributes only held value "
        "from rows not already aged or Phase Out. Show the aged/Phase Out "
        "overlap explicitly and do not double count it. Deduct only the "
        "supplied supported-recovery value. `recommended_reserve` is gross "
        "reserve exposure less that supported recovery, never below zero. "
        "Rank the top exposure at item/site level, not by item across sites. "
        "Label the result as a June pre-close candidate requiring Controller "
        "approval, not a booked reserve."
    ),
    "inventory_turns": (
        "Use LTM posted manufacturing cost of sales and June 30 discrete "
        "inventory; report turns and 365-day inventory days. Manufacturing "
        "cost of sales includes material cost of sales plus posted PPV, "
        "production variance, direct-labor clearing, and variable/fixed "
        "overhead absorption accounts."
    ),
    "inventory_transfer": (
        "Use 2026 transfer movements for the requested family. Aggregate each "
        "reference's raw `quantity * unit_cost` before rounding final money "
        "values. Preserve the ledger signs: `transfer_out` is negative and "
        "`transfer_in` is positive. Calculate `net_difference` as "
        "`transfer_out + transfer_in` so balanced references reconcile to zero."
    ),
    "forecast": (
        "Use approved forecast version 2026-06-SOP for July through December "
        "2026. Use the demand report's exact scope totals, which aggregate raw "
        "forecast rows before rounding; `item_count` is the distinct forecast "
        "item count, not the number of displayed month/site/family groups."
    ),
    "backlog": (
        "Use open, unshipped sales-order quantities at 2026-06-30 and compare "
        "them with July-December 2026 consensus demand for the same scope. "
        "Use the backlog report's exact scope totals, which aggregate raw "
        "sales-order lines before rounding once; do not sum the displayed "
        "rounded detail rows. "
        "Calculate `coverage = backlog_units / forecast_units`; "
        "`forecast_revenue` is the separately reported value of that same "
        "consensus demand and is not the coverage denominator."
    ),
    "planning": "Use plan version 2026-06-SOP and its explicit action messages; do not substitute production-order status.",
    "capacity": (
        "Use the work-center calendar net of planned downtime and planned "
        "operation load for the requested Q3 window. Use the capacity "
        "report's exact scope totals, which aggregate raw operation load "
        "before rounding once; do not sum the displayed rounded "
        "work-center rows. Return `top_constraint` as the work-center code, "
        "not its display name."
    ),
    "capacity_scenario": (
        "Use the work-center calendar net of planned downtime and planned "
        "operation load for the requested Q3 site/window. Preserve the base "
        "case before applying the supplied scenario. Apply "
        "`required_hours_change` uniformly to each work center's required "
        "hours and `net_available_hours_change` uniformly to each work "
        "center's net available hours; both are decimal ratios. Recalculate "
        "site totals, utilization, overloaded-center count, and top constraint "
        "from the adjusted work-center rows. Do not present a site-level ratio "
        "alone as proof that individual constraints are relieved."
    ),
    "profitability": "Use posted customer-invoice lines from 2026-01-01 through 2026-06-30 and current item/site standard cost.",
    "standard_cost": "Use all requested item/site records and separately total material, labor, variable overhead, fixed overhead, and outside processing.",
    "margin_sensitivity": (
        "Use the requested product-family row from the YTD profitability "
        "report. Its `units` and `material_cost` are exact invoice-line scope "
        "inputs. Calculate `scenario_standard_cost = baseline_standard_cost + "
        "material_cost * cost_change`, then apply only the stated price and "
        "material-cost shocks. Calculate the break-even price change once to "
        "preserve baseline gross-profit dollars and separately once to "
        "preserve baseline gross-margin percentage."
    ),
    "ppv": (
        "Use YTD posted vendor-invoice quantities and prices against item/site "
        "standard material cost. Use the PPV report's exact scope totals, "
        "which aggregate atomic invoice-line values before rounding once; do "
        "not sum the displayed rounded vendor rows."
    ),
    "ppv_causal": (
        "Use YTD posted vendor-invoice quantities. Bridge standard material "
        "cost to PO price as `po_price_variance = quantity * "
        "(PO unit price - standard unit cost)`, then PO price to invoice price "
        "as `invoice_price_variance = quantity * (invoice unit price - PO "
        "unit price)`. The two causes must reconcile exactly to total PPV; "
        "report the residual in `bridge_difference` and require zero."
    ),
    "supplier": (
        "Use distinct receipts dated from 2026-01-01 through 2026-06-30. "
        "`vendor_count` is the number of distinct vendors with at least one "
        "qualifying receipt, and `receipt_count` is the number of distinct "
        "qualifying receipt IDs. `on_time_percent` is the pooled on-time "
        "receipt count divided by pooled `receipt_count`; "
        "`accepted_percent` is the pooled Accepted receipt count divided by "
        "pooled `receipt_count`; do not average vendor percentages. "
        "`receipt_value` is the sum of receipt-line quantity times "
        "receipt-line unit price. A receipt is on time only when it was "
        "received on or before the PO expected date. For each vendor, calculate "
        "`supplier_score = 0.5 * (on_time_receipts / vendor_receipts) + "
        "0.5 * (accepted_receipts / vendor_receipts)` from unrounded count "
        "ratios. Select the lowest score, break an exact tie by ascending "
        "vendor ID, and return that decimal ratio rounded to 6 decimals in "
        "`worst_vendor_score`. Preserve supplier category exactly as recorded."
    ),
    "commitments": (
        "Include only POs ordered on or before 2026-06-30 with status `Open` "
        "and lines where order quantity exceeds received quantity. Open "
        "quantity is order quantity minus received quantity; open commitment "
        "is that quantity times unit price, aggregated before final rounding. "
        "When no `site_code` is supplied, return one companywide scalar per "
        "key rather than objects by site. `overdue_commitment` uses expected "
        "dates before 2026-06-30. `next_30_day_commitment` includes both "
        "2026-06-30 and 2026-07-30. Rank `top_vendor` by total open "
        "commitment and break an exact tie by ascending vendor ID."
    ),
    "invoice_holds": "Use open vendor invoices whose match status is not Matched; isolate quality-related holds and past-due exposure.",
    "payroll": (
        "Use posted payroll checks from 2026-01-01 through 2026-06-30, "
        "preserving employee distinct counts and employer taxes/benefits. "
        "When no site is supplied, return one companywide scalar per key "
        "across all plants. `overtime_rate` is overtime hours divided by "
        "regular plus overtime hours; `top_department` is ranked by employer "
        "cost."
    ),
    "scrap": "Use YTD production completions and atomic production-variance scrap cost.",
    "quality": (
        "Use open quality orders, open nonconformances, and June 30 "
        "blocked/quality-held inventory. Rank `top_item` and calculate "
        "`top_item_exposure` from open quality-order estimated financial "
        "exposure only; report nonconformance count and blocked inventory "
        "value separately without adding them to `top_item_exposure`."
    ),
    "copq": (
        "Define known YTD cost of poor quality as settled-date atomic "
        "production-variance scrap cost plus open quality-order estimated "
        "financial exposure plus June 30 blocked-inventory value. Keep the "
        "three components separate and do not add open nonconformance counts "
        "as dollars. Disclose that warranty, returns, and rework are outside "
        "the available source scope."
    ),
    "maintenance": (
        "Use all YTD maintenance work orders. Treat Corrective work as "
        "unplanned, Preventive work as planned preventive, and Calibration "
        "and Inspection as other scheduled/control work. Total work orders, "
        "downtime, and cost include all four work types; calculate "
        "`unplanned_downtime_percent` as Corrective downtime divided by total "
        "YTD maintenance downtime. Use the asset-level maintenance report to "
        "rank `top_asset` by total downtime hours; break an exact tie by "
        "ascending asset ID."
    ),
    "maintenance_capacity": (
        "Use YTD maintenance work orders and the YTD work-center calendar. "
        "Separate Corrective work as unplanned and Preventive work as planned. "
        "Calculate `ytd_net_available_hours` as available calendar hours less "
        "planned downtime, then calculate unplanned maintenance downtime as a "
        "percentage of those net available hours. Use the asset-level "
        "maintenance report to rank the top asset by total downtime, breaking "
        "ties by ascending asset ID."
    ),
    "fixed_assets": "Use the active fixed-asset register at 2026-06-30; YTD additions are assets placed in service during 2026.",
    "npv": (
        "Assume benefits occur at each year-end, no terminal value, no tax "
        "effect, and no residual value. Calculate NPV, IRR, payback, and "
        "profitability index from year-by-year cash-flow logic without "
        "rounding intermediate present values."
    ),
    "capital_portfolio": (
        "Evaluate every supplied project independently using year-end benefits "
        "with no terminal value, tax effect, or residual value. Rank projects "
        "by positive NPV and profitability index, then select the highest-NPV "
        "combination that does not exceed `capital_budget`. Show requested, "
        "selected, and remaining capital and do not present one project as a "
        "portfolio."
    ),
    "make_buy": (
        "Compare annual relevant make cost to buy cost, include the stated "
        "one-time transition cost, and discount three year-end savings. "
        "Calculate `annual_savings = make_total_cost - buy_total_cost`. Every "
        "repeated Annual savings result in summaries, bridges, and controls "
        "must formula-link or reconcile to that authoritative value; never "
        "sum make components and the already-totaled make cost together. "
        "When annual savings are nonpositive, report `payback_years` as "
        "`No payback` and `payback_status` as an explicit no-recovery status; "
        "never report 0.0 years or imply immediate recovery. Also report the "
        "formula-backed three-year undiscounted savings including the year-0 "
        "transition outflow."
    ),
    "cash_debt": (
        "Use June 30 posted cash and the active debt register. Focus the "
        "decision on available liquidity: undrawn revolver equals active "
        "commitment less outstanding principal, and total liquidity equals "
        "posted cash plus that undrawn committed revolver."
    ),
    "debt_profile": (
        "Use every debt instrument whose status is Active at 2026-06-30 and "
        "classify it by its exact recorded instrument type: Revolver, "
        "Equipment Loan, or Term Loan. Count instruments independently from "
        "their balances. `long_term_debt = total_debt - current_debt`, where "
        "current debt is the sum of each instrument's recorded current "
        "portion. For each instrument type, annual cash interest equals "
        "outstanding principal multiplied by its stated annual interest rate; "
        "total those independently derived amounts and calculate the weighted "
        "interest rate as total annual cash interest divided by total debt."
    ),
    "borrowing_base": (
        "Treat positive open AR and positive on-hand inventory as the gross "
        "and eligible collateral pools. Report those pools before advance "
        "rates in `eligible_ar` and `eligible_inventory`; apply the stated "
        "rates only when calculating the gross `borrowing_base`. Use the "
        "commitment and outstanding principal from the Active Revolver. "
        "Calculate `availability = max(0, min(borrowing_base, "
        "revolver_commitment) - revolver_outstanding)`; availability may "
        "never exceed undrawn committed capacity."
    ),
    "cash_forecast": (
        "Use the approved Base 13-week treasury assumptions in the source "
        "corpus, beginning with posted June 30 cash. Report supplier payments, "
        "payroll and benefits, interest, and capital expenditures as positive "
        "cash-use totals that are subtracted in the cash bridge. "
        "`minimum_cash_week` is the ISO week-ending date. "
        "`maximum_revolver` is the maximum cumulative incremental forecast "
        "draw outstanding across the 13 weeks needed to maintain the "
        "minimum-cash buffer, excluding the opening revolver balance. Add "
        "draws from successive shortfall weeks rather than taking only the "
        "largest single-week draw."
    ),
    "covenant": (
        "Use funded debt, June 30 liquidity, tangible net worth, and LTM "
        "activity from 2025-07-01 through 2026-06-30. Calculate "
        "`ltm_adjusted_ebitda = LTM operating_income + posted account 6400 "
        "Depreciation & Amortization`. Operating income includes revenue, "
        "cost of goods sold, and operating-expense accounts 6000-6999, and "
        "excludes account 7000 Interest Expense, account 7100 Other Income / "
        "Expense, and account 8000 tax expense. Account 7100 is non-operating "
        "and is excluded by definition, not treated as a discretionary "
        "management add-back. When reconciling from pretax income, calculate "
        "`ltm_adjusted_ebitda = pretax_income + account_7000_interest_expense "
        "+ signed_account_7100_other_expense + account_6400_depreciation_and_"
        "amortization`; use account 7100's signed debit-minus-credit balance "
        "exactly once. Make no other EBITDA add-backs. State both covenant "
        "tests."
    ),
    "tax": "Use YTD posted pretax income plus the stated permanent items; apply federal tax first and the stated state rate to the same taxable-income base.",
    "interplant": (
        "Reconcile 2026 AFT transfer movements by aggregating raw "
        "`quantity * unit_cost` before rounding final money values. Preserve "
        "the ledger signs: `transfer_out` is negative, `transfer_in` is "
        "positive, and `net_difference = transfer_out + transfer_in`. "
        "Calculate `transfer_price_adjustment = transfer_out * markup_rate`, "
        "so the adjustment retains the signed transfer-out direction."
    ),
    "qoe": (
        "Use 2025-07-01 through 2026-06-30 for LTM measures. "
        "`reported_ebitda` equals LTM operating income plus posted account "
        "6400 depreciation and amortization; operating income excludes "
        "interest, other non-operating expense, and tax. "
        "`production_variance_normalization` is the signed sum of "
        "`total_variance` for production variances settled in that LTM "
        "period. `quality_normalization` is the positive estimated financial "
        "exposure of quality orders Open at 2026-06-30. "
        "`maintenance_normalization` is the positive labor, material, and "
        "outside-service cost of Corrective maintenance work orders requested "
        "from 2026-01-01 through 2026-06-30. Add all three normalizations to "
        "`reported_ebitda`; `normalized_ebitda = reported_ebitda + "
        "production_variance_normalization + quality_normalization + "
        "maintenance_normalization`. Calculate "
        "`quality_of_earnings_ratio = reported_ebitda / normalized_ebitda`."
    ),
    "dcf": (
        "First derive normalized EBITDA using the quality-of-earnings formula: "
        "LTM is 2025-07-01 through 2026-06-30; reported EBITDA is LTM "
        "operating income plus posted account 6400 depreciation and "
        "amortization; add the signed LTM settled production `total_variance`, "
        "the positive estimated exposure of quality orders Open at "
        "2026-06-30, and the positive cost of Corrective maintenance work "
        "orders requested from 2026-01-01 through 2026-06-30. Thus "
        "`normalized_ebitda = reported_ebitda + "
        "production_variance_normalization + quality_normalization + "
        "maintenance_normalization`. Maintenance capital expenditures are "
        "the gross cost of active fixed assets "
        "placed in service from 2026-01-01 through 2026-06-30. Calculate "
        "`base_cash_flow = normalized_ebitda * (1 - cash_tax_rate) - "
        "maintenance_capital_expenditures`. For each forecast year `t` from "
        "1 through `forecast_years`, calculate `cash_flow_t = base_cash_flow "
        "* (1 + growth_rate) ^ t`. Calculate terminal value at the end of the "
        "last forecast year as `cash_flow_N * (1 + terminal_growth) / "
        "(discount_rate - terminal_growth)`. `enterprise_value` is the sum "
        "of each forecast cash flow discounted by `(1 + discount_rate) ^ t` "
        "plus terminal value discounted by `(1 + discount_rate) ^ N`. "
        "`net_debt` is active debt principal less posted cash across "
        "accounts 1000, 1010, and 1020, and `equity_value = enterprise_value "
        "- net_debt`. "
        "Do not round intermediate forecast or present-value calculations; "
        "round only requested final money values."
    ),
    "roic": (
        "Use LTM operating income and `nopat = operating_income * "
        "(1 - tax_rate)`. At 2026-06-30, calculate `total_assets` as posted "
        "debits minus credits across every balance-sheet Asset account, so "
        "credit-balance contra assets reduce the total; include Work in "
        "Process Inventory account 1210. Cash is accounts 1000, 1010, and "
        "1020. Calculate `invested_capital = total_assets - cash - "
        "accounts_payable_account_2000 - "
        "accrued_payroll_account_2010`; accrued payroll is an operating "
        "liability and must be subtracted. Calculate `roic = nopat / "
        "invested_capital`."
    ),
    "executive": (
        "Use authoritative June 30/YTD modules, not stale presentation "
        "figures. For production variance, invoke the authoritative report "
        "with `settled_only=true`. For backlog, PPV, and production variance, "
        "use each report's exact scope totals, which aggregate atomic values "
        "before rounding once; do not sum displayed rounded detail rows. Return "
        "`period_start` as `2026-01-01`, `period_end` as `2026-06-30`, and "
        "`status` as `June pre-close`; do not imply that June is closed."
    ),
    "forecast_scenario": (
        "Use the exact `baseline_start` through `baseline_end` date range. "
        "Return those dates unchanged as `baseline_period_start` and "
        "`baseline_period_end`. Return `scenario_period` as `FY26 H2 "
        "downside` for the July-December baseline task and `FY27 operating "
        "plan` for the full-year baseline tasks; return `status` as "
        "`Management scenario`. "
        "Baseline revenue is posted customer-invoice-line revenue in that "
        "range. Baseline standard cost is invoice quantity times the current "
        "item/site total standard cost, split between material and all other "
        "standard-cost components. Calculate `scenario_revenue = "
        "baseline_revenue * (1 + volume_change) * (1 + price_change)`. "
        "Calculate scenario material cost as baseline material cost times "
        "`(1 + volume_change) * (1 + material_cost_change)` and scenario "
        "nonmaterial cost as baseline nonmaterial standard cost times "
        "`(1 + volume_change)`. Scenario gross profit is scenario revenue "
        "less those two costs. Hold posted operating expense from the same "
        "baseline date range fixed when calculating scenario operating "
        "income. Use June 30 operating working capital (AR plus inventory plus "
        "WIP less AP and accrued payroll) divided by baseline revenue as the "
        "working-capital rate; apply that rate only to the scenario revenue "
        "change. Do not round intermediate calculations."
    ),
    "kpi_reconciliation": (
        "Compare the eight package KPIs (revenue, gross profit, operating "
        "income, inventory, WIP, backlog, cash, and debt) with independent "
        "recomputations from their authoritative ERP modules. Compare money "
        "after rounding both sides to cents; a KPI passes only when the "
        "absolute difference is no more than $0.005. Count exactly the eight "
        "PASS/FAIL results to derive `reconciliation_failures`; do not "
        "hard-code the count or compare KPIs with different periods/scopes. "
        "Return `period_start` as `2026-01-01`, `period_end` as "
        "`2026-06-30`, and `status` as `June pre-close` so the named "
        "package and recomputation use the same explicit period contract."
    ),
    "balance_cash_bridge": (
        "Use the June 30 balance sheet and active debt register on a consistent "
        "sign basis. Return `as_of_date` as `2026-06-30` and `status` as "
        "`June pre-close`. Calculate `operating_working_capital = "
        "accounts_receivable + inventory + work_in_process - "
        "accounts_payable - accrued_payroll`; calculate `net_debt = debt - "
        "cash`. Reconcile total assets, liabilities, and equity separately so "
        "other liabilities are not silently included in operating working "
        "capital."
    ),
}


def _format_parameters(blueprint: TaskBlueprint) -> str:
    if not blueprint.parameters:
        return ""
    return "Use these assumptions:\n```json\n" + json.dumps(
        blueprint.parameters, indent=2, sort_keys=True
    ) + "\n```"


def _format_source_contract(blueprint: TaskBlueprint) -> str:
    sources = authoritative_sources(blueprint)
    return (
        "Use these as the controlling sources:\n"
        + "\n".join(f"- `{source}`" for source in sources)
    )


def _request_line(blueprint: TaskBlueprint, *, context: str = "") -> str:
    number = int(blueprint.task_id.rsplit("_", 1)[1])
    title = blueprint.title.rstrip(".")
    lower_title = title[:1].lower() + title[1:]
    starts_with_action = title.split(maxsplit=1)[0].lower() in {
        "build", "create", "enter", "place", "post", "prepare", "reverse", "write"
    }
    if starts_with_action:
        options = (
            f"Please {lower_title}{context}.",
            f"Can you {lower_title}{context}?",
            f"I need you to {lower_title}{context}.",
            f"Please {lower_title}{context}.",
        )
    else:
        options = (
            f"Please complete {lower_title}{context}.",
            f"Can you prepare {lower_title}{context}?",
            f"I need {lower_title}{context}.",
            f"Please run {lower_title}{context}.",
        )
    return options[(number - 1) % len(options)]


def _console_prompt(blueprint: TaskBlueprint) -> str:
    keys = BUNDLE_KEYS[blueprint.bundle]
    task_specific = {
        "task_021": (
            "For `top_month_revenue`, use the selected month's raw atomic "
            "forecast-row sum and round the final sum once to cents; do not "
            "sum already-rounded detail groups."
        ),
        "task_032": (
            "Return the customer display name, not the customer ID, in "
            "`top_group` and `lowest_margin_group`."
        ),
        "task_063": (
            "Report `irr` and `profitability_index` to four decimal places."
        ),
        "task_064": (
            "For non-positive annual savings, return `No payback` and a "
            "plain-language status explaining that the investment does not "
            "recover its cost."
        ),
        "task_065": (
            "Report `irr` and `profitability_index` to four decimal places."
        ),
    }.get(blueprint.task_id, "")
    response = f"""Return one valid JSON object and no surrounding prose. Use exactly these top-level keys, in this order:
{json.dumps(list(keys))}

Formatting requirements:
- Money is a JSON number in USD rounded to 2 decimals; quantities use up to 4 decimals.
- Rates are decimal ratios, not percentage points (for example, 0.125 means 12.5%).
- Counts are integers. Names, IDs, periods, and statuses are strings.
- Every top-level value is a scalar JSON number, integer, string, or boolean; do not return nested objects or arrays.
- Show signed values exactly as calculated. Do not omit a key, add a key, use `null`, or invent evidence.
- Reconcile the result across the ERP and the listed source files. Do not add narrative or citation fields to the response.
"""
    sections = (
        _request_line(blueprint),
        blueprint.source_hint,
        _format_source_contract(blueprint),
        f"Use this calculation basis: {PERIOD_GUIDANCE[blueprint.bundle]}",
        _format_parameters(blueprint),
        task_specific,
        response,
    )
    return "\n\n".join(section.strip() for section in sections if section.strip()) + "\n"


def _artifact_prompt(blueprint: TaskBlueprint) -> str:
    assert blueprint.target_path
    metric_keys = list(BUNDLE_KEYS[blueprint.bundle])
    project_scope = ""
    if (
        blueprint.output_mode == "document"
        and blueprint.parameters.get("site_name")
        and blueprint.parameters.get("site_code")
    ):
        project_scope = (
            "\nProject scope: The memorandum and its approval recommendation "
            f"must explicitly identify {blueprint.parameters['site_name']} "
            f"(site {blueprint.parameters['site_code']}) as the project site. "
            "Evidence from another plant may be used only as a clearly labeled "
            "comparison, never as the recommended project scope.\n"
        )
    if blueprint.task_id == "task_028":
        project_scope += (
            "\nAction-register scope: Include a complete native row-level "
            "register for every nonblank `2026-06-SOP` MRP action message. "
            "Use exactly these native register columns: `planned_order_id`, "
            "`item_id`, `site_code`, `quantity`, `action_message`, `priority`, "
            "`release_date`, `required_date`, and "
            "`source_demand_reference`. Make the register filterable and sort "
            "it by priority, required date, and planned-order ID so planners "
            "can execute it; summary counts alone are insufficient.\n"
        )
    if blueprint.task_id == "task_029":
        project_scope += (
            "\nDecision-deck scope: Bridge the approved consensus demand to "
            "company backlog, strict past-due backlog, the Dayton and "
            "Greenville Q3 work-center constraints, and the explicit MRP "
            "action-message population. The Decision and Actions slides must "
            "treat consensus demand as a planning baseline, not unconstrained "
            "supply, and assign owners and the supplied capacity-recovery "
            "timing. Forecast-only approval is insufficient.\n"
        )
    task_specific_scope = {
        "task_019": (
            "\nWorking-capital decision: Use the reconciled turns and days on "
            "hand to recommend a source-supported working-capital action. A "
            "quantified aged- or slow-moving-inventory reduction/disposition "
            "plan is appropriate; do not merely restate the metrics.\n"
        ),
        "task_026": (
            "\nBacklog decision: Recommend a source-supported backlog-to-cash, "
            "overdue-risk, capacity-mitigation, or re-promise action. Identify "
            "the top customer exposure when concentration is material. The "
            "recommendation must be supported by the quantified bridge.\n"
        ),
        "task_027": (
            "\nScenario input precision: Use the approved capacity report's "
            "two-decimal `scope_required_hours` and "
            "`scope_net_available_hours` as the authoritative base inputs, "
            "then apply the supplied changes. Do not substitute hidden "
            "database precision that the report does not expose.\n"
        ),
        "task_036": (
            "\nRequired detail: Include one filterable row for every in-scope "
            "customer, with customer ID/name, end market, region, units, "
            "revenue, standard cost, gross profit, gross margin, and revenue "
            "rank. The complete 240-customer schedule must reconcile to the "
            "headline totals; a top/bottom summary alone is insufficient.\n"
        ),
        "task_037": (
            "\nRequired bridge: Include all five product-family rows and a "
            "formula-driven revenue-to-standard-cost-to-gross-profit bridge. "
            "Show family margin and ranking and reconcile the family schedule "
            "to company totals.\n"
        ),
        "task_038": (
            "\nRequired sensitivity: Keep baseline material and nonmaterial "
            "cost visible, formula-calculate price-only, material-cost-only, "
            "combined, gross-profit break-even, and margin break-even cases, "
            "and chart comparable margin or gross-profit outputs using "
            "consistent units.\n"
        ),
        "task_039": (
            "\nPricing decision: Recommend a specific minimum DRV price action. "
            "State separately the formula-driven increase required to preserve "
            "baseline gross-profit dollars and the increase required to "
            "preserve baseline gross-margin percentage. Assign the supplied "
            "owner and decision date and explain which threshold management "
            "should approve.\n"
        ),
        "task_040": (
            "\nCommercial-review scope: Compare all five product families in a "
            "native table and chart using consistent units. Identify the "
            "revenue-protection priority, lowest-margin family, quantified "
            "economics, accountable owner, and dated corrective action.\n"
        ),
        "task_047": (
            "\nRecovery-register scope: Include one filterable row for every "
            "in-scope supplier with vendor ID/name, category, country, receipt "
            "count, on-time rate, acceptance rate, receipt value, composite "
            "score, risk tier, recovery action, owner, and due date. Apply the "
            "supplied thresholds with visible formulas.\n"
        ),
        "task_048": (
            "\nCausal-bridge scope: Bridge standard cost to PO price and PO "
            "price to invoice price. Show both formula-driven cause amounts, "
            "their total, and a zero residual to authoritative PPV; a "
            "vendor-only summary is insufficient.\n"
        ),
        "task_049": (
            "\nSupplier-risk decision: Evaluate the priority supplier using "
            "delivery, quality, lead time, PPV, open commitment, primary-item "
            "dependency, and country exposure. Select a concrete mitigation "
            "such as recovery, dual-source qualification, or safety stock, "
            "with the supplied owner and decision date. Use natural business "
            "labels for those fields; do not add a risk tier that is not in "
            "the supplied sources.\n"
        ),
        "task_056": (
            "\nRequired detail: Include one filterable row for every in-scope "
            "site/department combination with employee count, regular hours, "
            "overtime hours, overtime rate, gross pay, employer cost, and "
            "employer-cost rank. Reconcile the detailed schedule to company "
            "totals.\n"
        ),
        "task_057": (
            "\nCOPQ scope: Separately show production scrap, open "
            "quality-order exposure, and blocked inventory, then reconcile "
            "them to known COPQ. Explicitly disclose warranty, returns, and "
            "rework as unavailable exclusions rather than implying they are "
            "zero. Recommend a source-supported COPQ action; prioritizing the "
            "largest scrap family and resolving the top open-quality item is "
            "an appropriate course of action.\n"
        ),
        "task_058": (
            "\nMaintenance-capacity scope: Separate corrective and preventive "
            "work while retaining Calibration and Inspection rows so the "
            "native schedule reconciles all YTD work orders and maintenance "
            "cost. Calculate unplanned downtime as a percentage of YTD net "
            "available work-center hours, identify the top downtime asset, "
            "and state the capacity-recovery action.\n"
        ),
        "task_059": (
            "\nMulti-plant scope: Include separate native rows for sites 100, "
            "200, and 300 with orders, completed units, scrap units, yield, "
            "standard output, and production variance. The chart must compare "
            "like-for-like plant metrics and Actions must assign the supplied "
            "owner and due date.\n"
        ),
        "task_066": (
            "\nPortfolio scope: Include every supplied project, its site, "
            "investment, annual net benefit, NPV, IRR, payback, profitability "
            "index, rank, and select/defer decision. Optimize selection within "
            "the supplied capital budget and show remaining budget.\n"
        ),
        "task_067": (
            "\nBusiness-case scope: Identify Dayton site 200, include status "
            "quo and investment alternatives, a year-0 through year-7 "
            "cash-flow and present-value schedule, NPV, IRR, payback, "
            "profitability index, sensitivities, implementation risks, the "
            "supplied owner/gate date, and a post-audit measure.\n"
        ),
        "task_068": (
            "\nMake-buy model scope: Build a visible annual relevant-cost "
            "bridge from units and unit costs, separate avoidable fixed cost "
            "and the one-time transition outflow, and include a year 0 through "
            "year 3 discounted-cash-flow schedule. Formula-calculate annual "
            "savings, payback, three-year undiscounted savings, and NPV. "
            "Show a formula-backed payback status; if annual savings are "
            "non-positive, report `No payback` rather than 0.0 or immediate "
            "recovery. "
            "Include a make-versus-buy chart with like-for-like dollar units.\n"
        ),
        "task_076": (
            "\nCash-forecast scope: Include all 13 weekly periods with opening "
            "cash, customer collections, operating disbursements, interest, "
            "other inflows/outflows, pre-financing cash, revolver draw or "
            "repayment, and closing cash. Every week must roll from the prior "
            "week, preserve the minimum-cash buffer, and reconcile the maximum "
            "incremental revolver need and ending cash to the summary.\n"
        ),
        "task_077": (
            "\nCovenant-model scope: Show active debt, LTM EBITDA, leverage "
            "threshold/headroom, tangible net worth, its minimum/headroom, "
            "cash, undrawn revolver, and total liquidity. Formula-drive both "
            "covenant tests and a combined PASS/FAIL conclusion from visibly "
            "sourced inputs; include no pass-through-only controls.\n"
        ),
        "task_078": (
            "\nMemorandum decision scope: Explain liquidity, active debt, both "
            "covenant tests and their headroom, downside triggers, and the "
            "13-week cash outlook, plus specific reporting/mitigation actions. "
            "Retain explicit Fact and Assumption labels. Assign Treasurer, CFO, "
            "and Controller ownership under explicit Owner and Timing labels "
            "with dated deliverables, and distinguish executed debt terms from "
            "management forecasts.\n"
        ),
        "task_079": (
            "\nLender-update scope: Present the active debt profile and "
            "liquidity, leverage and tangible-net-worth tests with headroom, "
            "cash outlook, risks/triggers, and dated actions owned by the CFO, "
            "Treasurer, and Controller. Charts must keep dollars and ratios on "
            "separate axes or separate visuals.\n"
        ),
        "task_086": (
            "\nDCF model scope: Include a visible reported-to-normalized "
            "EBITDA bridge, maintenance-capex derivation, year 1 through year "
            "5 cash-flow and present-value schedule, terminal-value formula, "
            "enterprise-to-equity bridge, and a terminal-growth sensitivity "
            "whose base case remains exactly 0.025. No central output may be "
            "hard-coded or self-reconciled.\n"
        ),
        "task_087": (
            "\nQoE bridge scope: Show reported EBITDA and each signed "
            "production-variance, quality, and corrective-maintenance "
            "normalization in a formula-driven bridge to normalized EBITDA. "
            "Show the normalization ratio and a zero-difference tie-out to the "
            "summary; chart bridge amounts in consistent dollar units.\n"
        ),
        "task_088": (
            "\nTax-memo scope: Reconcile book income to taxable income, show "
            "permanent and depreciation differences, federal and state rates, "
            "current tax, and the combined effective rate. Use benefit wording "
            "when the calculated tax amount is negative, and specify review, "
            "support-retention, and close-control ownership with dates.\n"
        ),
        "task_089": (
            "\nValuation-memo scope: Explain the reported-to-normalized EBITDA "
            "bridge, maintenance capital, explicit forecast, terminal value, "
            "enterprise-to-equity bridge, core DCF assumptions, sensitivity, "
            "valuation risks, and an approval recommendation with an owner and "
            "dated next gate.\n"
        ),
        "task_090": (
            "\nValuation-deck scope: Tell the full valuation story from "
            "reported EBITDA through normalizations, forecast cash flows and "
            "terminal value to enterprise value, net debt, and equity value. "
            "Show assumptions and sensitivity, decision risks, and dated "
            "committee actions; use separate visuals for dollars and rates.\n"
        ),
        "task_095": (
            "\nIntegrated-plan scope: Build a January-through-December FY27 "
            "schedule using the corresponding 2025 monthly actual pattern as "
            "the baseline. For every month, visibly formula-calculate revenue, "
            "material cost, nonmaterial cost, gross profit, gross margin, "
            "operating expense, operating income, and incremental working "
            "capital from the supplied volume, price, and material-cost "
            "assumptions. Annual results must be SUMs of the 12 monthly rows, "
            "and the base price assumption must remain exactly 0.025 after "
            "native recalculation.\n"
        ),
        "task_096": (
            "\nCFO-note scope: State which June close controls are complete or "
            "open, identify the largest margin, backlog-to-cash, working-"
            "capital, and liquidity exposures, and make explicit approve/hold "
            "decisions for the KPI package, final period lock, and discretionary "
            "capital. Include every open close control in the decision register, "
            "with an accountable owner and dated next action for each priority; "
            "June must remain labeled pre-close.\n"
        ),
        "task_097": (
            "\nBoard-priorities scope: Connect the FY27 monthly operating plan "
            "to revenue, gross profit, operating income, and incremental "
            "working capital. Identify quantified value-creation priorities, "
            "funding/capacity risks, milestones, accountable executives, and "
            "dated board gates rather than presenting a metric recap alone.\n"
        ),
        "task_098": (
            "\nBoard-deck scope: Separate FY26 YTD current state from the FY26 "
            "second-half downside outlook and FY27 upside plan. Show revenue, "
            "gross profit/margin, operating income, working-capital and "
            "liquidity implications for both forward scenarios. Define each "
            "scenario's pro forma liquidity as June 30 total liquidity minus "
            "that scenario's incremental working capital: positive working "
            "capital is a liquidity use and negative working capital is a "
            "liquidity source. Then convert "
            "them into quantified FY27 priorities with owners, dates, triggers, "
            "and explicit board decisions. A June YTD metric snapshot alone is "
            "insufficient; charts must not mix dollars and ratios on one axis.\n"
        ),
    }
    project_scope += task_specific_scope.get(blueprint.task_id, "")
    common = "\n\n".join(
        section.strip()
        for section in (
            _request_line(blueprint),
            blueprint.source_hint,
            _format_source_contract(blueprint),
            f"Use this calculation basis: {PERIOD_GUIDANCE[blueprint.bundle]}",
            _format_parameters(blueprint),
            project_scope,
            (
                f"Create the deliverable at `{blueprint.target_path}` and leave the source files "
                "unchanged. The finished file must show the calculation chain for these metrics:\n"
                f"{json.dumps(metric_keys)}"
            ),
            (
                "Keep the shared company workspace clean: create only the requested deliverable "
                "and do not leave behind draft scripts or temporary exports."
            ),
        )
        if section.strip()
    )
    if blueprint.output_mode == "spreadsheet":
        body = """
For the workbook:
- Include sheets named `Read Me`, `Inputs`, `Analysis`, and `Control`.
- Preserve source lineage in Inputs with source object/file, date or version, and extraction cutoff.
- Put scenario assumptions in visibly distinct input cells. Drive Analysis and Control by formulas; do not hard-code calculated results.
- Include a decision-ready summary, the requested detail/bridge, and at least one useful native chart that clarifies a KPI, comparison, trend, composition, or control result.
- Include zero-difference controls tying every central metric to an independent authoritative total; a pass-through reference is not a control.
- Use professional number formats, frozen headers, filters where useful, readable widths, and no formula errors. Wrap or widen populated text cells so labels, dates, and source notes are not clipped by adjacent values.
- Round monetary controls to cents before evaluating zero/tolerance or PASS/FAIL results.
- Normalize dates to a common date type before comparing them; do not compare mixed text and date representations.
"""
    elif blueprint.output_mode == "document":
        body = """
For the document:
- Maximum 3 pages, with a clear title, date, audience, and June pre-close labeling. Use Word's built-in Title style for the title.
- Use Word's built-in heading styles for `Executive conclusion`, `Evidence`, `Economics`, `Risks and controls`, and `Sources`.
- Include a native calculation table with metric, source/input, calculation logic, and result columns, plus a native decision table with the recommended decision/action, named owner, and timing.
- Cite visible source filenames and ERP report names. Distinguish facts from assumptions and do not claim June is closed.
- Number every page in the footer with native page numbering and render every page to confirm the number is visible.
"""
    else:
        body = """
For the presentation:
- Exactly 5 slides titled `Decision`, `Evidence`, `Economics`, `Risks`, and `Actions`.
- Include exact central metrics, one native editable chart, one native editable table, a clear decision, named owners, timing, and source footnotes.
- Use a coherent executive visual hierarchy; no screenshots of tables, overlapping objects, clipped text, placeholders, or unsupported claims.
"""
    completion = f"""When the file is complete, return one JSON object only:
{{"deliverable":"{blueprint.target_path}","status":"complete","central_decision":"<concise decision>","sources_checked":["<source 1>","<source 2>"]}}
"""
    return "\n\n".join(part.strip() for part in (common, body, completion)) + "\n"


def _erp_prompt(blueprint: TaskBlueprint) -> str:
    action = {
        "erp_draft_journal": (
            "Create exactly one balanced draft journal with the supplied header and lines. "
            "Do not post it."
        ),
        "erp_posted_journal": (
            "Create exactly one balanced draft journal with the supplied header and lines, "
            "then post that same journal."
        ),
        "erp_purchase_order": (
            "Create exactly one purchase order with the supplied header and line values."
        ),
        "erp_production_order": (
            "Create exactly one production order with the supplied values. Apply only the "
            "status transitions needed to reach `final_status`; when omitted, leave it Scheduled."
        ),
        "erp_quality_hold": (
            "Use ERP on-hand to choose a valid available warehouse/location for the supplied "
            "item and site, place exactly one hold for the supplied quantity and reason, and "
            "release only when `release` is true. The quality order must retain the full "
            "placed quantity in `quantity_inspected`; this neutral inventory hold has "
            "`quantity_failed = 0`. A release changes the historical order to Closed "
            "with the supplied release disposition, without erasing its inspected "
            "quantity, and must exactly restore the prior blocked inventory balance. "
            "`placed_quantity` is the historical quantity placed and inspected; it is "
            "not the current blocked quantity after a release."
        ),
    }[blueprint.bundle]
    source_context = ""
    if (
        blueprint.workflow == "production-execution"
        or blueprint.bundle == "erp_purchase_order"
    ):
        source_context = (
            f"{blueprint.source_hint}\n\n"
            f"{_format_source_contract(blueprint)}\n\n"
        )
    controls = f"""Before you finish:
- Use the Pinehaven ERP workflow tools; do not alter files or use direct database access.
- Make no additional ERP write, retry duplicate, substitute item/site/account, or amount change.
- Confirm the returned business object and audit events after the write.

Return one valid JSON object and no surrounding prose, with exactly these keys:
{json.dumps(list(BUNDLE_KEYS[blueprint.bundle]))}
"""
    sections = (
        _request_line(blueprint, context=" in Pinehaven ERP"),
        source_context,
        action,
        _format_parameters(blueprint),
        controls,
    )
    return "\n\n".join(section.strip() for section in sections if section.strip()) + "\n"


def _difficulty(blueprint: TaskBlueprint) -> str:
    if blueprint.output_mode in {"presentation", "erp"} or blueprint.task_id in {
        "task_029", "task_059", "task_079", "task_084", "task_090",
        "task_091", "task_092", "task_093", "task_094", "task_095",
        "task_096", "task_097", "task_098", "task_099", "task_100",
    }:
        return "expert"
    return "advanced"


def _spec(blueprint: TaskBlueprint) -> TaskSpec:
    if blueprint.task_id in TASK_PROMPTS:
        prompt = TASK_PROMPTS[blueprint.task_id]
    elif blueprint.output_mode == "console":
        prompt = _console_prompt(blueprint)
    elif blueprint.output_mode in {"spreadsheet", "document", "presentation"}:
        prompt = _artifact_prompt(blueprint)
    else:
        prompt = _erp_prompt(blueprint)
    return TaskSpec(
        task_id=blueprint.task_id,
        slug=blueprint.slug,
        title=blueprint.title,
        workflow=blueprint.workflow,
        output_mode=blueprint.output_mode,
        prompt=prompt,
        difficulty=_difficulty(blueprint),
        target_path=blueprint.target_path,
    )


TASKS: tuple[TaskSpec, ...] = tuple(_spec(blueprint) for blueprint in TASK_BLUEPRINTS)
TASK_BY_ID = {task.task_id: task for task in TASKS}

if len(TASKS) != 100 or len(TASK_BY_ID) != 100:
    raise AssertionError("The task catalog must contain exactly 100 unique tasks")
