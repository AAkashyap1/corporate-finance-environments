"""FastMCP server exposing Pinehaven's application-style ERP boundary."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Literal

from fastmcp import FastMCP
from pydantic import BaseModel, ConfigDict

from pinehaven_company.erp import PinehavenERP


DATABASE_ENV = "PINEHAVEN_ERP_DB"
_DATABASE_PATH: Path | None = None
mcp = FastMCP(
    "Pinehaven Manufacturing ERP",
    instructions=(
        "Use the ERP's business-object and report tools. Do not infer that raw "
        "database access exists. Dates are ISO YYYY-MM-DD and months are YYYY-MM."
    ),
)
server = mcp

AuditEntity = Literal[
    "journal", "purchase_order", "production_order", "inventory"
]
ItemType = Literal[
    "Aftermarket",
    "Finished Good",
    "Manufactured Subassembly",
    "Purchased Component",
    "Raw Material",
]
ItemLifecycle = Literal["Active", "Phase Out"]
EngineeringChangeStatus = Literal[
    "Approved", "Draft", "Implemented", "Under Review"
]
ProductionOrderStatus = Literal[
    "Scheduled",
    "Released",
    "Started",
    "Reported Finished",
    "Ended",
    "Cancelled",
]
InventoryTransactionType = Literal[
    "Customer Shipment",
    "Opening Balance",
    "Production Issue",
    "Production Receipt",
    "Purchase Receipt",
    "Transfer In",
    "Transfer Out",
]
QualityStatus = Literal["Closed", "Open"]
MaintenanceType = Literal[
    "Calibration", "Corrective", "Inspection", "Preventive"
]
PurchaseOrderStatus = Literal["Cancelled", "Closed", "Open"]
SalesOrderStatus = Literal["Cancelled", "Invoiced", "Open"]
ForecastVersion = Literal[
    "2026-05-RF", "2026-06-DOWNSIDE", "2026-06-SOP", "FY26-AOP"
]
JournalSource = Literal[
    "Cash Receipt",
    "Debt",
    "Equity",
    "Factory Cost Summary",
    "Fixed Assets",
    "Historical ERP Summary",
    "Inventory Control",
    "Inventory Opening",
    "Manufacturing",
    "Manufacturing Control",
    "Opening",
    "Order to Cash",
    "Payroll",
    "Procure to Pay",
    "Treasury",
    "Vendor Payment",
]
FixedAssetClass = Literal[
    "Building Improvement", "IT & Furniture", "Machinery"
]


class JournalLineInput(BaseModel):
    """Strict, application-level journal line accepted by the ERP workflow."""

    model_config = ConfigDict(extra="forbid", allow_inf_nan=False)

    account_code: str
    debit: float | None = None
    credit: float | None = None
    site_code: str | None = None
    department_code: str | None = None
    item_id: str | None = None
    customer_id: str | None = None
    vendor_id: str | None = None
    line_memo: str | None = None


class PurchaseOrderLineInput(BaseModel):
    """Strict purchase-order line accepted by the controlled PO workflow."""

    model_config = ConfigDict(extra="forbid", allow_inf_nan=False)

    item_id: str
    quantity: float
    unit_price: float


def configure_database(path: str | Path) -> None:
    """Bind the episode database without exposing its path to the agent shell."""

    global _DATABASE_PATH
    _DATABASE_PATH = Path(path).resolve()


def _erp() -> PinehavenERP:
    database = os.environ.get(DATABASE_ENV)
    path = Path(database).resolve() if database else _DATABASE_PATH
    if path is None:
        raise RuntimeError(f"{DATABASE_ENV} is not configured")
    return PinehavenERP(path)


@mcp.tool
def get_company_profile() -> dict[str, Any]:
    """Return company, site, period, scale, and ERP object-count context."""
    return _erp().get_company_profile()


@mcp.tool
def list_erp_modules() -> list[dict[str, Any]]:
    """List installed ERP modules and the business objects in each."""
    return _erp().list_erp_modules()


@mcp.tool
def list_erp_reports() -> list[str]:
    """List the standard report families available through this ERP."""
    return [
        "Production order cost analysis",
        "Production variance",
        "WIP",
        "Scrap and yield",
        "Work-center capacity",
        "Inventory on hand, valuation, movement, and aging",
        "Quality and nonconformance",
        "Maintenance cost and downtime",
        "Open purchase commitments and supplier performance",
        "Purchase price variance",
        "Backlog, shipments, and customer/product profitability",
        "Demand forecast, planned orders, and planning exceptions",
        "Trial balance, P&L, balance sheet, and GL detail",
        "AR/AP aging and subledger reconciliation",
        "Cash, debt, fixed assets, and payroll",
    ]


@mcp.tool
def list_fiscal_periods() -> list[dict[str, Any]]:
    """Return fiscal periods and their Open, Closed, or Future status."""
    return _erp().list_fiscal_periods()


@mcp.tool
def get_audit_events(
    entity: AuditEntity | None = None,
    identifier: str | None = None,
    limit: int = 100,
) -> list[dict[str, Any]]:
    """Retrieve immutable audit events for controlled ERP writes."""
    return _erp().get_audit_events(entity, identifier, limit)


@mcp.tool
def search_items(
    query: str | None = None,
    item_type: ItemType | None = None,
    product_family: str | None = None,
    site_code: str | None = None,
    lifecycle_status: ItemLifecycle | None = None,
    limit: int = 100,
) -> list[dict[str, Any]]:
    """Search item master records by text, type, family, site, or lifecycle."""
    return _erp().search_items(
        query, item_type, product_family, site_code, lifecycle_status, limit
    )


@mcp.tool
def get_item(item_id: str) -> dict[str, Any]:
    """Return one item master and all site-specific planning/cost records."""
    return _erp().get_item(item_id)


@mcp.tool
def get_bom(
    item_id: str,
    site_code: str,
    as_of: str = "2026-06-30",
    explode_levels: int = 1,
) -> dict[str, Any]:
    """Return the effective BOM, optionally exploded through four levels."""
    return _erp().get_bom(item_id, site_code, as_of, explode_levels)


@mcp.tool
def get_routing(
    item_id: str, site_code: str, as_of: str = "2026-06-30"
) -> dict[str, Any]:
    """Return the effective routing and work-center rates for an item/site."""
    return _erp().get_routing(item_id, site_code, as_of)


@mcp.tool
def get_standard_cost(item_id: str, site_code: str) -> dict[str, Any]:
    """Return material, labor, overhead, outside-process, and total standard cost."""
    return _erp().get_standard_cost(item_id, site_code)


@mcp.tool
def list_engineering_changes(
    status: EngineeringChangeStatus | None = None,
    effective_from: str | None = None,
    limit: int = 100,
) -> list[dict[str, Any]]:
    """List engineering changes and their effectivity/financial impact."""
    return _erp().list_engineering_changes(status, effective_from, limit)


@mcp.tool
def search_production_orders(
    status: ProductionOrderStatus | None = None,
    site_code: str | None = None,
    item_id: str | None = None,
    finish_from: str | None = None,
    finish_to: str | None = None,
    limit: int = 100,
) -> list[dict[str, Any]]:
    """Search production orders by lifecycle, plant, item, and finish window."""
    return _erp().search_production_orders(
        status, site_code, item_id, finish_from, finish_to, limit
    )


@mcp.tool
def get_production_order(order_number: str) -> dict[str, Any]:
    """Return a production order with copied material, operation, and variance detail."""
    return _erp().get_production_order(order_number)


@mcp.tool
def get_production_cost_analysis(order_number: str) -> dict[str, Any]:
    """Compare actual order cost to standard reported-output cost."""
    return _erp().get_production_cost_analysis(order_number)


@mcp.tool
def get_production_variance_report(
    date_from: str,
    date_to: str,
    site_code: str | None = None,
    product_family: str | None = None,
    settled_only: bool = True,
) -> dict[str, Any]:
    """Summarize atomic production variance components by plant and family."""
    return _erp().get_production_variance_report(
        date_from, date_to, site_code, product_family, settled_only
    )


@mcp.tool
def get_wip_report(
    as_of: str = "2026-06-30", site_code: str | None = None
) -> dict[str, Any]:
    """Return open-order cost-to-date, reported output, and residual WIP."""
    return _erp().get_wip_report(as_of, site_code)


@mcp.tool
def get_scrap_yield_report(
    date_from: str, date_to: str, site_code: str | None = None
) -> list[dict[str, Any]]:
    """Report started, completed, scrapped units, yield, and scrap cost."""
    return _erp().get_scrap_yield_report(date_from, date_to, site_code)


@mcp.tool
def get_work_center_capacity(
    date_from: str, date_to: str, site_code: str | None = None
) -> list[dict[str, Any]]:
    """Compare work-center capacity and load, including exact scope totals."""
    return _erp().get_work_center_capacity(date_from, date_to, site_code)


@mcp.tool
def get_inventory_on_hand(
    item_id: str | None = None,
    site_code: str | None = None,
    warehouse_code: str | None = None,
    include_zero: bool = False,
    limit: int = 500,
) -> list[dict[str, Any]]:
    """Return on-hand, reserved, held, available, cost, and value by dimension."""
    return _erp().get_inventory_on_hand(
        item_id, site_code, warehouse_code, include_zero, limit
    )


@mcp.tool
def get_inventory_valuation(
    site_code: str | None = None,
    group_by: Literal[
        "item_type", "product_family", "site", "warehouse", "item"
    ] = "item_type",
) -> dict[str, Any]:
    """Value inventory by item_type, product_family, site, warehouse, or item."""
    return _erp().get_inventory_valuation(site_code, group_by)


@mcp.tool
def get_inventory_movements(
    date_from: str,
    date_to: str,
    item_id: str | None = None,
    site_code: str | None = None,
    transaction_type: InventoryTransactionType | None = None,
    limit: int = 500,
    product_family: str | None = None,
) -> list[dict[str, Any]]:
    """Return perpetual movements with raw and display-rounded extensions."""
    return _erp().get_inventory_movements(
        date_from,
        date_to,
        item_id,
        site_code,
        transaction_type,
        limit,
        product_family,
    )


@mcp.tool
def get_inventory_aging(
    as_of: str = "2026-06-30",
    site_code: str | None = None,
    lifecycle_status: ItemLifecycle | None = None,
    threshold_days: int = 180,
) -> dict[str, Any]:
    """Age in-scope inventory and return exact threshold and top-item controls."""
    return _erp().get_inventory_aging(
        as_of, site_code, lifecycle_status, threshold_days
    )


@mcp.tool
def get_quality_orders(
    status: QualityStatus | None = None,
    site_code: str | None = None,
    opened_from: str | None = None,
    limit: int = 200,
) -> list[dict[str, Any]]:
    """Search quality orders, failure rates, dispositions, and financial exposure."""
    return _erp().get_quality_orders(status, site_code, opened_from, limit)


@mcp.tool
def get_nonconformances(
    status: QualityStatus | None = None,
    site_code: str | None = None,
    limit: int = 200,
) -> list[dict[str, Any]]:
    """Return nonconformance root causes, owners, dispositions, and cost."""
    return _erp().get_nonconformances(status, site_code, limit)


@mcp.tool
def get_maintenance_report(
    date_from: str,
    date_to: str,
    site_code: str | None = None,
    maintenance_type: MaintenanceType | None = None,
) -> list[dict[str, Any]]:
    """Summarize maintenance work orders, downtime, and cost."""
    return _erp().get_maintenance_report(
        date_from, date_to, site_code, maintenance_type
    )


@mcp.tool
def get_maintenance_asset_report(
    date_from: str,
    date_to: str,
    site_code: str | None = None,
    limit: int = 200,
) -> list[dict[str, Any]]:
    """Rank asset-level maintenance downtime, work orders, and cost."""
    return _erp().get_maintenance_asset_report(
        date_from, date_to, site_code, limit
    )


@mcp.tool
def search_vendors(
    query: str | None = None,
    commodity: str | None = None,
    active_only: bool = True,
    limit: int = 100,
) -> list[dict[str, Any]]:
    """Search supplier master records and open AP."""
    return _erp().search_vendors(query, commodity, active_only, limit)


@mcp.tool
def search_purchase_orders(
    status: PurchaseOrderStatus | None = None,
    vendor_id: str | None = None,
    site_code: str | None = None,
    expected_from: str | None = None,
    expected_to: str | None = None,
    limit: int = 100,
) -> list[dict[str, Any]]:
    """Search POs and open commitments by supplier, plant, status, or due date."""
    return _erp().search_purchase_orders(
        status, vendor_id, site_code, expected_from, expected_to, limit
    )


@mcp.tool
def get_purchase_order(po_number: str) -> dict[str, Any]:
    """Return PO header, lines, receipts, quality status, and matched invoices."""
    return _erp().get_purchase_order(po_number)


@mcp.tool
def get_open_purchase_commitments(
    as_of: str = "2026-06-30", site_code: str | None = None
) -> dict[str, Any]:
    """Return grouped commitments plus exact scope, due-window, and vendor totals."""
    return _erp().get_open_purchase_commitments(as_of, site_code)


@mcp.tool
def get_purchase_price_variance(
    date_from: str,
    date_to: str,
    site_code: str | None = None,
    vendor_id: str | None = None,
    group_by: Literal["vendor", "item", "site", "commodity"] = "vendor",
) -> dict[str, Any]:
    """Compare purchase cost to standard by vendor, item, site, or commodity."""
    return _erp().get_purchase_price_variance(
        date_from, date_to, site_code, vendor_id, group_by
    )


@mcp.tool
def get_supplier_performance(
    date_from: str,
    date_to: str,
    vendor_id: str | None = None,
    limit: int = 200,
) -> list[dict[str, Any]]:
    """Report supplier on-time delivery, acceptance, and actual lead time."""
    return _erp().get_supplier_performance(date_from, date_to, vendor_id, limit)


@mcp.tool
def search_customers(
    query: str | None = None,
    end_market: str | None = None,
    region: str | None = None,
    limit: int = 100,
) -> list[dict[str, Any]]:
    """Search customer master records and open AR."""
    return _erp().search_customers(query, end_market, region, limit)


@mcp.tool
def search_sales_orders(
    status: SalesOrderStatus | None = None,
    customer_id: str | None = None,
    product_family: str | None = None,
    promised_from: str | None = None,
    promised_to: str | None = None,
    limit: int = 100,
) -> list[dict[str, Any]]:
    """Search sales orders and backlog value by customer/family/date/status."""
    return _erp().search_sales_orders(
        status, customer_id, product_family, promised_from, promised_to, limit
    )


@mcp.tool
def get_sales_order(order_number: str) -> dict[str, Any]:
    """Return sales order, customer, line, shipment, and invoice detail."""
    return _erp().get_sales_order(order_number)


@mcp.tool
def get_backlog(
    as_of: str = "2026-06-30",
    site_code: str | None = None,
    product_family: str | None = None,
) -> dict[str, Any]:
    """Return unshipped backlog by promise date, plant, and family."""
    return _erp().get_backlog(as_of, site_code, product_family)


@mcp.tool
def get_shipments(
    date_from: str,
    date_to: str,
    customer_id: str | None = None,
    site_code: str | None = None,
    limit: int = 500,
) -> list[dict[str, Any]]:
    """Retrieve shipment, carrier, customer, item, quantity, and standard cost."""
    return _erp().get_shipments(date_from, date_to, customer_id, site_code, limit)


@mcp.tool
def get_customer_product_profitability(
    date_from: str,
    date_to: str,
    group_by: Literal[
        "customer", "product_family", "item", "end_market", "region"
    ] = "customer",
    limit: int = 200,
) -> dict[str, Any]:
    """Report invoiced units, material cost, and profitability by dimension."""
    return _erp().get_customer_product_profitability(
        date_from, date_to, group_by, limit
    )


@mcp.tool
def get_demand_forecast(
    forecast_version: ForecastVersion,
    month_from: str | None = None,
    month_to: str | None = None,
    product_family: str | None = None,
    site_code: str | None = None,
) -> list[dict[str, Any]]:
    """Return grouped demand plus exact once-rounded scope totals and item count."""
    return _erp().get_demand_forecast(
        forecast_version, month_from, month_to, product_family, site_code
    )


@mcp.tool
def get_planned_orders(
    plan_version: str,
    order_type: str | None = None,
    site_code: str | None = None,
    action_message: str | None = None,
    required_to: str | None = None,
    limit: int = 500,
) -> list[dict[str, Any]]:
    """Return planned purchase/production supply and pegging/action messages."""
    return _erp().get_planned_orders(
        plan_version, order_type, site_code, action_message, required_to, limit
    )


@mcp.tool
def get_planning_exceptions(
    plan_version: str, site_code: str | None = None
) -> dict[str, Any]:
    """Summarize non-null master-planning action messages."""
    return _erp().get_planning_exceptions(plan_version, site_code)


@mcp.tool
def get_trial_balance(as_of: str = "2026-06-30") -> dict[str, Any]:
    """Return posted debits, credits, and balance by account through a date."""
    return _erp().get_trial_balance(as_of)


@mcp.tool
def get_profit_and_loss(
    date_from: str, date_to: str, site_code: str | None = None
) -> dict[str, Any]:
    """Return exact P&L summary totals plus rounded account rows."""
    return _erp().get_profit_and_loss(date_from, date_to, site_code)


@mcp.tool
def get_operating_scenario_baseline(
    date_from: str,
    date_to: str,
    as_of: str = "2026-06-30",
) -> dict[str, Any]:
    """Return invoice-cost, opex, and working-capital scenario components."""
    return _erp().get_operating_scenario_baseline(
        date_from, date_to, as_of
    )


@mcp.tool
def get_balance_sheet(as_of: str = "2026-06-30") -> dict[str, Any]:
    """Return balance-sheet accounts plus current and unclosed earnings."""
    return _erp().get_balance_sheet(as_of)


@mcp.tool
def get_gl_detail(
    date_from: str,
    date_to: str,
    account_code: str | None = None,
    source: JournalSource | None = None,
    reference: str | None = None,
    site_code: str | None = None,
    status: Literal["Posted"] = "Posted",
    limit: int = 500,
) -> list[dict[str, Any]]:
    """Retrieve dimensioned journal detail without exposing arbitrary SQL."""
    return _erp().get_gl_detail(
        date_from,
        date_to,
        account_code,
        source,
        reference,
        site_code,
        status,
        limit,
    )


@mcp.tool
def get_subledger_reconciliation(as_of: str = "2026-06-30") -> dict[str, Any]:
    """Reconcile AR, AP, discrete inventory, and WIP to the GL."""
    return _erp().get_subledger_reconciliation(as_of)


@mcp.tool
def get_ar_aging(as_of: str = "2026-06-30") -> dict[str, Any]:
    """Return open customer invoices in standard aging buckets."""
    return _erp().get_ar_aging(as_of)


@mcp.tool
def get_ap_aging(as_of: str = "2026-06-30") -> dict[str, Any]:
    """Return open vendor invoices in standard aging buckets."""
    return _erp().get_ap_aging(as_of)


@mcp.tool
def get_cash_and_debt(as_of: str = "2026-06-30") -> dict[str, Any]:
    """Return cash, debt instruments, net debt, and available liquidity."""
    return _erp().get_cash_and_debt(as_of)


@mcp.tool
def get_fixed_asset_register(
    site_code: str | None = None,
    asset_class: FixedAssetClass | None = None,
    status: Literal["Active"] | None = "Active",
    limit: int = 500,
) -> list[dict[str, Any]]:
    """Return fixed-asset cost, accumulated depreciation, and net book value."""
    return _erp().get_fixed_asset_register(site_code, asset_class, status, limit)


@mcp.tool
def get_payroll_summary(
    date_from: str,
    date_to: str,
    site_code: str | None = None,
    department_code: str | None = None,
) -> list[dict[str, Any]]:
    """Summarize headcount, hours, pay, taxes, benefits, and employer cost."""
    return _erp().get_payroll_summary(
        date_from, date_to, site_code, department_code
    )


@mcp.tool
def create_draft_journal(
    posting_date: str,
    reference: str,
    memo: str,
    lines: list[JournalLineInput],
) -> dict[str, Any]:
    """Create an audited balanced draft journal in an open fiscal period."""
    return _erp().create_draft_journal(
        posting_date,
        reference,
        memo,
        [line.model_dump(exclude_none=True) for line in lines],
    )


@mcp.tool
def post_draft_journal(reference: str) -> dict[str, Any]:
    """Post an existing balanced draft journal if its period remains open."""
    return _erp().post_draft_journal(reference)


@mcp.tool
def create_purchase_order(
    vendor_id: str,
    site_code: str,
    order_date: str,
    expected_date: str,
    buyer_employee_id: str,
    lines: list[PurchaseOrderLineInput],
) -> dict[str, Any]:
    """Create an audited open purchase order after master-data validation."""
    return _erp().create_purchase_order(
        vendor_id,
        site_code,
        order_date,
        expected_date,
        buyer_employee_id,
        [line.model_dump() for line in lines],
    )


@mcp.tool
def create_production_order(
    item_id: str,
    site_code: str,
    order_quantity: float,
    scheduled_start: str,
    scheduled_finish: str,
    source_reference: str | None = None,
) -> dict[str, Any]:
    """Create a scheduled production order with copied effective BOM/routing."""
    return _erp().create_production_order(
        item_id,
        site_code,
        order_quantity,
        scheduled_start,
        scheduled_finish,
        source_reference,
    )


@mcp.tool
def update_production_order_status(
    order_number: str,
    new_status: Literal["Released", "Started", "Reported Finished", "Cancelled"],
) -> dict[str, Any]:
    """Apply a validated and audited production lifecycle transition."""
    return _erp().update_production_order_status(order_number, new_status)


@mcp.tool
def place_inventory_quality_hold(
    item_id: str,
    site_code: str,
    warehouse_code: str,
    location_code: str,
    quantity: float,
    reason: str,
    lot_number: str = "",
) -> dict[str, Any]:
    """Move available quantity to quality hold and open a quality order."""
    return _erp().place_inventory_quality_hold(
        item_id,
        site_code,
        warehouse_code,
        location_code,
        quantity,
        reason,
        lot_number,
    )


@mcp.tool
def release_inventory_quality_hold(
    quality_order_id: str, disposition: str
) -> dict[str, Any]:
    """Close an agent-created inventory hold and restore held availability."""
    return _erp().release_inventory_quality_hold(quality_order_id, disposition)


def main() -> None:
    host = os.environ.get("PINEHAVEN_MCP_HOST", "127.0.0.1")
    port = int(os.environ.get("PINEHAVEN_MCP_PORT", "8765"))
    mcp.run(transport="http", host=host, port=port)


if __name__ == "__main__":
    main()
