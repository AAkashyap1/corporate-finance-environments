from __future__ import annotations

import os
from pathlib import Path

from fastmcp import FastMCP
from pydantic import BaseModel, ConfigDict

from .accounting import AccountingRepository, json_result


server = FastMCP(name="Alder Ridge Contractor Accounting")


class _StrictInput(BaseModel):
    """Closed JSON objects that every hosted model provider can validate."""

    model_config = ConfigDict(extra="forbid")


class ExportParameters(_StrictInput):
    start_date: str | None = None
    end_date: str | None = None
    as_of_date: str | None = None
    group_by: str | None = None
    account_code: str | None = None
    project_id: str | None = None
    customer_id: str | None = None
    technician_id: str | None = None
    employee_id: str | None = None
    service_work_order_id: str | None = None
    department_code: str | None = None
    status: str | None = None
    cleared: bool | None = None
    limit: int | None = None
    offset: int | None = None


class JournalLine(_StrictInput):
    account_code: str
    department_code: str | None = None
    project_id: str | None = None
    cost_code: str | None = None
    debit: float = 0.0
    credit: float = 0.0
    line_memo: str | None = None


class InvoiceLine(_StrictInput):
    account_code: str
    amount: float
    line_number: int | None = None
    department_code: str | None = None
    project_id: str | None = None
    service_work_order_id: str | None = None
    description: str | None = None
    quantity: float | None = None
    unit_price: float | None = None


class InvoiceInput(_StrictInput):
    invoice_number: str
    customer_id: str
    invoice_date: str
    due_date: str
    amount: float
    project_id: str | None = None
    retainage_amount: float = 0.0
    description: str | None = None
    revenue_account_code: str | None = None
    department_code: str | None = None
    line_items: list[InvoiceLine] | None = None


class BillLine(_StrictInput):
    account_code: str
    amount: float
    line_number: int | None = None
    department_code: str | None = None
    project_id: str | None = None
    cost_code: str | None = None
    description: str | None = None
    quantity: float | None = None
    units: float | None = None
    unit_cost: float | None = None


class BillInput(_StrictInput):
    bill_number: str
    vendor_id: str
    invoice_date: str
    due_date: str
    amount: float
    project_id: str | None = None
    department_code: str | None = None
    cost_code: str | None = None
    expense_account_code: str | None = None
    payment_hold: bool = False
    hold_reason: str | None = None
    description: str | None = None
    line_items: list[BillLine] | None = None


class PurchaseOrderInput(_StrictInput):
    po_number: str
    vendor_id: str
    order_date: str
    expected_date: str
    buyer: str
    project_id: str | None = None


class PurchaseOrderLine(_StrictInput):
    cost_code: str
    description: str
    committed_amount: float
    line_number: int | None = None
    invoiced_amount: float = 0.0


class TransactionChanges(_StrictInput):
    due_date: str | None = None
    description: str | None = None
    retainage_amount: float | None = None
    payment_hold: bool | None = None
    hold_reason: str | None = None
    status: str | None = None
    expected_date: str | None = None
    buyer: str | None = None
    posting_date: str | None = None
    reference: str | None = None
    memo: str | None = None


def _repo() -> AccountingRepository:
    state_root = Path(os.environ.get("COMPANY_STATE_ROOT", ".runtime/state")).resolve()
    return AccountingRepository(state_root / "accounting.db")


def _workspace() -> Path:
    return Path(os.environ.get("COMPANY_RUNTIME_ROOT", ".runtime/workspace")).resolve()


@server.tool
def get_company_profile() -> str:
    """Company, reporting cutoff, legal entity, departments, and source-of-truth metadata."""
    return json_result(_repo().company_profile())


@server.tool
def list_available_reports() -> str:
    """List report names and their output grain before requesting data."""
    return json_result(_repo().list_reports())


@server.tool
def list_accounts(account_type: str | None = None) -> str:
    """Chart of accounts with code, name, type, normal balance and cash flag; optionally filter by account type."""
    return json_result(_repo().list_accounts(account_type))


@server.tool
def list_master_data(entity: str, search: str | None = None, limit: int = 250) -> str:
    """List/search customers, vendors, employees, projects, departments, cost codes, agreements, fixed assets, or prepaids."""
    return json_result(_repo().list_master_data(entity, search, limit))


@server.tool
def query_transactions(entity: str, start_date: str | None = None, end_date: str | None = None,
                       status: str | None = None, counterparty_id: str | None = None,
                       project_id: str | None = None, limit: int = 250, offset: int = 0) -> str:
    """Search invoices, bills, POs, journals, work orders, timecards, payroll runs, bank activity, or change orders."""
    return json_result(_repo().query_transactions(entity, start_date, end_date, status, counterparty_id, project_id, limit, offset))


@server.tool
def get_transaction(entity: str, identifier: str) -> str:
    """Retrieve one accounting or operating transaction; invoices, bills, POs, journals, and payroll runs include lines."""
    return json_result(_repo().get_transaction(entity, identifier))


@server.tool
def get_audit_log(entity: str | None = None, identifier: str | None = None, limit: int = 250) -> str:
    """Read immutable MCP write-event history, optionally filtered by entity and document identifier."""
    return json_result(_repo().audit_log(entity, identifier, limit))


@server.tool
def get_trial_balance(start_date: str, end_date: str) -> str:
    """Posted trial balance activity for an inclusive date range (YYYY-MM-DD)."""
    return json_result(_repo().trial_balance(start_date, end_date))


@server.tool
def get_profit_and_loss(start_date: str, end_date: str, group_by: str = "month") -> str:
    """Posted P&L grouped by month, department, or account."""
    return json_result(_repo().profit_and_loss(start_date, end_date, group_by))


@server.tool
def get_balance_sheet(as_of_date: str) -> str:
    """Posted balance sheet balances as of a date."""
    return json_result(_repo().balance_sheet(as_of_date))


@server.tool
def get_general_ledger(
    start_date: str,
    end_date: str,
    account_code: str | None = None,
    project_id: str | None = None,
    limit: int = 100,
    offset: int = 0,
) -> str:
    """Posted journal-line detail filtered by date and optional account/project; page with limit (max 100) and offset."""
    return json_result(_repo().general_ledger(
        start_date, end_date, account_code, project_id, min(max(limit, 1), 100), max(offset, 0)
    ))


@server.tool
def get_project_summary(project_id: str | None = None, as_of_date: str = "2026-06-30") -> str:
    """Current contract, cost budget, cost-to-date, and billings for one or all open jobs."""
    return json_result(_repo().project_summary(project_id, as_of_date))


@server.tool
def get_job_cost_detail(
    project_id: str,
    as_of_date: str = "2026-06-30",
    limit: int = 100,
    offset: int = 0,
) -> str:
    """Job-cost transaction detail through an as-of date; page with limit (max 100) and offset."""
    return json_result(_repo().job_cost_detail(
        project_id, as_of_date, min(max(limit, 1), 100), max(offset, 0)
    ))


@server.tool
def get_ar_aging(as_of_date: str = "2026-06-30") -> str:
    """Open customer invoices and retainage with aging buckets."""
    return json_result(_repo().ar_aging(as_of_date))


@server.tool
def get_ap_aging(as_of_date: str = "2026-06-30") -> str:
    """Open vendor bills, payment holds, and due dates."""
    return json_result(_repo().ap_aging(as_of_date))


@server.tool
def get_open_purchase_orders(as_of_date: str = "2026-06-30") -> str:
    """Remaining commitments on open and partially received purchase orders."""
    return json_result(_repo().open_purchase_orders(as_of_date))


@server.tool
def get_cash_balances(as_of_date: str = "2026-06-30") -> str:
    """Book cash by bank account as of a date."""
    return json_result(_repo().cash_balances(as_of_date))


@server.tool
def get_bank_activity(start_date: str, end_date: str, account_code: str | None = None,
                      cleared: bool | None = None, limit: int = 100, offset: int = 0) -> str:
    """Bank-feed detail with book links; page with limit (max 100) and offset."""
    return json_result(_repo().bank_activity(
        start_date, end_date, account_code, cleared, min(max(limit, 1), 100), max(offset, 0)
    ))


@server.tool
def get_service_work_orders(start_date: str, end_date: str, customer_id: str | None = None,
                            technician_id: str | None = None, status: str | None = None,
                            limit: int = 100, offset: int = 0) -> str:
    """Service dispatch detail; page with limit (max 100) and offset."""
    return json_result(_repo().service_work_orders(
        start_date, end_date, customer_id, technician_id, status,
        min(max(limit, 1), 100), max(offset, 0)
    ))


@server.tool
def get_timecard_detail(start_date: str, end_date: str, employee_id: str | None = None,
                        project_id: str | None = None, service_work_order_id: str | None = None,
                        status: str | None = None, limit: int = 100, offset: int = 0) -> str:
    """Time-entry detail; page with limit (max 100) and offset."""
    return json_result(_repo().timecard_detail(
        start_date, end_date, employee_id, project_id, service_work_order_id, status,
        min(max(limit, 1), 100), max(offset, 0)
    ))


@server.tool
def get_payroll_register(start_date: str, end_date: str, employee_id: str | None = None,
                         department_code: str | None = None, limit: int = 100,
                         offset: int = 0) -> str:
    """Payroll-register detail; page with limit (max 100) and offset."""
    return json_result(_repo().payroll_register(
        start_date, end_date, employee_id, department_code,
        min(max(limit, 1), 100), max(offset, 0)
    ))


@server.tool
def get_fixed_asset_register(as_of_date: str = "2026-06-30") -> str:
    """In-service fixed assets, cost, accumulated depreciation, net book value, location, and vendor."""
    return json_result(_repo().fixed_asset_register(as_of_date))


@server.tool
def get_prepaid_schedule(as_of_date: str = "2026-06-30") -> str:
    """Prepaid contracts through an as-of date with service term, original amount, amortization, and remaining balance."""
    return json_result(_repo().prepaid_schedule(as_of_date))


@server.tool
def get_project_change_orders(project_id: str | None = None, status: str | None = None) -> str:
    """Project change-order register with submission/approval status and revenue/cost-budget amounts."""
    return json_result(_repo().project_change_orders(project_id, status))


@server.tool
def get_service_agreements(as_of_date: str = "2026-06-30", status: str | None = None) -> str:
    """Customer service agreements active at an as-of date, including site, billing frequency, and annual value."""
    return json_result(_repo().service_agreements(as_of_date, status))


@server.tool
def export_accounting_report(report: str, relative_path: str, parameters: ExportParameters) -> str:
    """Write a report CSV inside the agent workspace and return path and row count."""
    kwargs = parameters.model_dump(exclude_none=True)
    return json_result(_repo().export_csv(report, _workspace(), relative_path, **kwargs))


@server.tool
def create_journal(posting_date: str, reference: str, memo: str, lines: list[JournalLine], post: bool = False) -> str:
    """Create a draft journal with dimensional lines; optionally post immediately if balanced. Prefer draft for review tasks."""
    payload = [line.model_dump(exclude_none=True) for line in lines]
    return json_result(_repo().create_journal(posting_date, reference, memo, payload, post))


@server.tool
def post_or_void_journal(journal_id: int, action: str) -> str:
    """Post a balanced draft journal or void a journal. Action is `post` or `void`."""
    return json_result(_repo().post_or_void_journal(journal_id, action))


@server.tool
def create_invoice(invoice: InvoiceInput) -> str:
    """Create and post an open customer invoice. Supports invoice_number, customer_id, optional project_id, dates, amount, retainage_amount, description, and optional revenue line_items."""
    return json_result(_repo().create_invoice(invoice.model_dump(exclude_none=True)))


@server.tool
def create_bill(bill: BillInput) -> str:
    """Create and post an open vendor bill. Supports bill_number, vendor_id, optional project/cost-code dimensions, dates, amount, hold fields, description, and optional expense line_items."""
    return json_result(_repo().create_bill(bill.model_dump(exclude_none=True)))


@server.tool
def create_purchase_order(purchase_order: PurchaseOrderInput, lines: list[PurchaseOrderLine]) -> str:
    """Create an open purchase order with vendor/project header data and cost-code-level commitment lines."""
    header = purchase_order.model_dump(exclude_none=True)
    detail = [line.model_dump(exclude_none=True) for line in lines]
    return json_result(_repo().create_purchase_order(header, detail))


@server.tool
def update_transaction(entity: str, identifier: str, changes: TransactionChanges) -> str:
    """Update supported workflow fields on invoices, bills, POs or draft journals; immutable posted accounting fields are protected."""
    payload = changes.model_dump(exclude_none=True)
    return json_result(_repo().update_transaction(entity, identifier, payload))


@server.tool
def apply_payment(entity: str, document_number: str, amount: float, payment_date: str, cash_account: str = "1000") -> str:
    """Apply a customer receipt or vendor payment and create the balanced posted cash journal. Entity is invoice or bill."""
    return json_result(_repo().apply_payment(entity, document_number, amount, payment_date, cash_account))


@server.tool
def update_document_status(entity: str, document_number: str, status: str, note: str | None = None) -> str:
    """Update workflow status for an invoice, bill or purchase order; bill notes may document hold/release rationale."""
    return json_result(_repo().update_document_status(entity, document_number, status, note))


@server.tool
def list_workspace_attachments(relative_directory: str = "Shared", pattern: str = "*") -> str:
    """List files available to the accounting workflow without exposing the database or host filesystem."""
    root = _workspace(); target = (root / relative_directory).resolve()
    if target != root and root not in target.parents: raise ValueError("directory must stay inside workspace")
    return json_result([{"path": str(p.relative_to(root)), "bytes": p.stat().st_size} for p in target.rglob(pattern) if p.is_file()][:2000])
