"""Natural-language task instructions keyed by task ID."""

from __future__ import annotations


TASK_PROMPTS: dict[str, str] = {
    "task_001": """Please pull together Pinehaven's June cross-ledger manufacturing close control.

Use the June 30, 2026 pre-close general ledger and the authoritative ERP subledgers for cash, AR, AP, inventory, WIP, production variance, debt, and journal balance. Follow the Monthly Close Policy and production-order close policy, and cross-check your work against `Finance/Monthly Close/2026-06 Close Binder.xlsx` and `Corporate Memos/June 2026 Close Instructions.docx`.

For production variance and GL account 5110, use activity from June 1 through June 30 rather than YTD activity. If there are negative WIP residuals, take the exact count and signed total from the WIP report. That total is calculated from the underlying open orders before rounding, so do not add the displayed rounded rows.

Reply here with one JSON object containing exactly these keys in this order:
["system_cash", "system_ar", "system_ap", "system_inventory", "system_wip", "system_production_variance", "gl_5110_production_variance", "production_variance_control_delta", "system_debt", "journal_imbalance", "maximum_control_delta"]

Show USD amounts to two decimals and preserve their signs. Include every key, use scalar values only, and do not add commentary or `null` values.""",

    "task_002": """Can you send me Pinehaven's June pre-close profit and margin bridge?

Pull posted activity from June 1 through June 30, 2026 from the ERP P&L and related production, WIP, and subledger reports. Use `Finance/Monthly Close/2026-06 Close Binder.xlsx`, `Corporate Memos/June 2026 Close Instructions.docx`, the Monthly Close Policy, and the production-order close policy as your cross-checks. June is still open, so the period status should remain pre-close rather than final.

Reply with one JSON object using exactly these keys in this order:
["revenue", "cost_of_goods_sold", "gross_profit", "gross_margin", "operating_expense", "operating_income", "period_status"]

Round USD amounts to two decimals, report gross margin as a decimal ratio, and preserve calculated signs. Include every key and no others; return only the JSON object, with no `null` values.""",

    "task_003": """Please prepare Pinehaven's settled production-variance bridge for the first half of 2026.

Use the ERP production-variance report for settled orders from January 1 through June 30, 2026. Check the result against the June close binder, June close instructions, Monthly Close Policy, and production-order close policy. Keep the favorable and unfavorable variance signs as reported. Use the report's exact scope totals, which are calculated from the underlying orders before rounding, rather than adding the displayed site or product-family rows.

Send back one JSON object with exactly these keys in this order:
["order_count", "standard_output_cost", "material_usage_variance", "labor_rate_variance", "labor_efficiency_variance", "variable_overhead_variance", "fixed_overhead_volume_variance", "scrap_variance", "rounding_adjustment", "total_variance", "variance_rate"]

Use an integer for the order count, two decimals for USD amounts, and a decimal ratio for the variance rate. Preserve all signs, include every key, and return only the JSON object without `null` values.""",

    "task_004": """I need Pinehaven's open-production WIP concentration as of June 30, 2026.

Include every production order that was not `Ended` at the cutoff. Calculate WIP as cost to date less the standard value of reported output, using the authoritative ERP WIP report and the June close binder. Use the report's exact scope totals; they are calculated from the underlying orders before rounding, so do not add the displayed rounded order rows. Follow the Monthly Close Policy, production-order close policy, and June close instructions.

Reply with one JSON object containing exactly these keys in this order:
["order_count", "started_orders", "reported_finished_orders", "cost_to_date", "reported_output", "wip_value", "largest_order", "largest_order_wip", "top_five_concentration"]

Use integers for counts, two decimals for USD amounts, and a decimal ratio for concentration. Preserve signs, include every key, and return only the JSON object without `null` values.""",

    "task_005": """Please send me the June output and production-variance analysis for Pinehaven's Grand Rapids plant.

For site 100, include production orders whose actual or scheduled finish falls between June 1 and June 30, 2026. Use the ERP production-completion and variance reports by site, the Grand Rapids plant scorecard, and `Finance/Analyses/Production Variance Bridge.xlsx`.

Reply with one JSON object using exactly these keys in this order:
["order_count", "completed_units", "scrap_units", "yield_rate", "standard_output", "total_variance", "variance_rate"]

Use an integer for the order count, up to four decimals for quantities, two decimals for USD amounts, and decimal ratios for the rates. Preserve calculated signs, include every key and no others, and return only the JSON object without `null` values.""",

    "task_006": """Can you run the June output and production-variance analysis for Pinehaven's Dayton plant?

Use site 200 production orders whose actual or scheduled finish falls between June 1 and June 30, 2026. Reconcile the ERP production-completion and variance reports by site to the Dayton plant scorecard and `Finance/Analyses/Production Variance Bridge.xlsx`.

Send me one JSON object containing exactly these keys in this order:
["order_count", "completed_units", "scrap_units", "yield_rate", "standard_output", "total_variance", "variance_rate"]

Use an integer for the order count, up to four decimals for quantities, two decimals for USD amounts, and decimal ratios for the rates. Preserve calculated signs, include every key and no others, and return only the JSON object without `null` values.""",

    "task_007": """Please build Pinehaven's June manufacturing close control workbook and save it as `Deliverables/june-manufacturing-close-control-workbook.xlsx`.

Base it on the June 30, 2026 pre-close general ledger and `ERP production variance, WIP, P&L, and subledger reconciliation`. Follow the Monthly Close Policy and production-order close policy, and use `Finance/Monthly Close/2026-06 Close Binder.xlsx` and `Corporate Memos/June 2026 Close Instructions.docx` as supporting sources. Production variance and GL account 5110 should reflect June 1 through June 30 activity, not YTD. Use the WIP report's exact count and signed total for negative residuals rather than adding rounded detail rows.

The workbook should cover:
["system_cash", "system_ar", "system_ap", "system_inventory", "system_wip", "system_production_variance", "gl_5110_production_variance", "production_variance_control_delta", "system_debt", "journal_imbalance", "maximum_control_delta"]

Include `Read Me`, `Inputs`, `Analysis`, and `Control` sheets. In `Inputs`, identify each source, its date or version, and the extraction cutoff. Keep assumptions visibly separate and use formulas for the analysis and controls. Include a decision-ready summary, supporting detail or bridge, and a useful native chart. Tie each central metric to an independent authoritative total with a zero-difference control; a pass-through reference is not enough.

Use clear professional formatting, frozen headers, filters where helpful, readable widths, and no formula errors. Make sure text is not clipped, normalize dates before comparing them, and round monetary controls to cents before applying tolerances or PASS/FAIL tests. Leave the source files unchanged. Keep the shared company workspace clean: create only the requested deliverable and do not leave behind draft scripts or temporary exports.

When it is ready, reply with only:
{"deliverable":"Deliverables/june-manufacturing-close-control-workbook.xlsx","status":"complete","central_decision":"<concise decision>","sources_checked":["<source 1>","<source 2>"]}""",

    "task_008": """Build a YTD production-variance bridge for Pinehaven and save the workbook as `Deliverables/ytd-production-variance-bridge-workbook.xlsx`.

Cover settled production orders from January 1 through June 30, 2026. Use `ERP production variance, WIP, P&L, and subledger reconciliation`, the Monthly Close Policy, production-order close policy, `Finance/Monthly Close/2026-06 Close Binder.xlsx`, and `Corporate Memos/June 2026 Close Instructions.docx`. Keep the reported favorable and unfavorable signs. Use the production-variance report's exact scope totals instead of adding rounded site or product-family rows.

Show the calculation chain for:
["order_count", "standard_output_cost", "material_usage_variance", "labor_rate_variance", "labor_efficiency_variance", "variable_overhead_variance", "fixed_overhead_volume_variance", "scrap_variance", "rounding_adjustment", "total_variance", "variance_rate"]

Create `Read Me`, `Inputs`, `Analysis`, and `Control` sheets. Document each source, date or version, and extraction cutoff in `Inputs`; distinguish assumptions from sourced inputs; and drive the analysis and controls with formulas. Include a concise decision summary, the variance bridge, and a useful native chart. Add independent zero-difference controls for the central metrics rather than simple pass-through references.

Format the workbook professionally with frozen headers, filters where useful, readable widths, visible labels and notes, and no formula errors. Normalize dates before comparing them and round monetary controls to cents before evaluating tolerances or PASS/FAIL results. Do not alter the source files. Keep the shared company workspace clean: create only the requested deliverable and do not leave behind draft scripts or temporary exports.

Once complete, reply with only:
{"deliverable":"Deliverables/ytd-production-variance-bridge-workbook.xlsx","status":"complete","central_decision":"<concise decision>","sources_checked":["<source 1>","<source 2>"]}""",

    "task_009": """Please prepare Pinehaven's June close risk memorandum and save it as `Deliverables/june-close-risk-memorandum.docx`.

Use `ERP production variance, WIP, P&L, and subledger reconciliation`, `Finance/Monthly Close/2026-06 Close Binder.xlsx`, and `Corporate Memos/June 2026 Close Instructions.docx`; do not use stale presentation figures. Run production variance with `settled_only=true`. For backlog, PPV, and production variance, use each report's exact scope total rather than adding rounded detail rows. The reporting period is January 1 through June 30, 2026, and the status is `June pre-close`; do not describe June as closed.

The memo should support these metrics:
["period_start", "period_end", "status", "ytd_revenue", "ytd_gross_profit", "ytd_gross_margin", "ytd_operating_income", "inventory", "wip", "backlog", "cash", "total_debt", "total_liquidity", "ytd_ppv", "ytd_production_variance"]

Keep it to three pages or fewer. Use Word's built-in Title style for the title and built-in heading styles for `Executive conclusion`, `Evidence`, `Economics`, `Risks and controls`, and `Sources`. Include the date, audience, and pre-close label. Add a native calculation table with metric, source/input, calculation logic, and result columns, plus a native decision table covering the recommended action, accountable owner, and timing. Name the source files and ERP reports, distinguish facts from assumptions, and use native page numbers in every footer. Render the finished memo to confirm the page numbers are visible. Leave the source files unchanged. Keep the shared company workspace clean: create only the requested deliverable and do not leave behind draft scripts or temporary exports.

When it is done, reply with only:
{"deliverable":"Deliverables/june-close-risk-memorandum.docx","status":"complete","central_decision":"<concise decision>","sources_checked":["<source 1>","<source 2>"]}""",

    "task_010": """Please enter the approved June scrap reclass in Pinehaven ERP as a draft journal. Do not post it.

Use this header and these lines exactly:
```json
{
  "lines": [
    {
      "account_code": "5110",
      "debit": 125000.0,
      "site_code": "100"
    },
    {
      "account_code": "5100",
      "credit": 125000.0,
      "site_code": "100"
    }
  ],
  "memo": "Reclass June abnormal scrap for management review",
  "posting_date": "2026-06-30",
  "reference": "PH-JE-010"
}
```

Create this journal once. Do not retry it, create a duplicate, change an account, site, amount, or file, or use direct database access. Afterward, verify the returned journal and its audit event.

Send me only one JSON object with exactly these keys in this order:
["reference", "status", "debits", "credits", "audit_action"]""",
}
