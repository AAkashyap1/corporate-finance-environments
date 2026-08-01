from __future__ import annotations

import csv
import io
import json
import sqlite3
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Iterable


class AccountingRepository:
    """Application-style facade over the simulated accounting system.

    The database is an implementation detail outside the agent workspace. Agents
    use MCP operations for reports, master data, transactions and controlled
    posting workflows; they never receive SQL or a database path.
    """

    def __init__(self, db_path: str | Path):
        self.db_path = Path(db_path)

    def _connect(self, write: bool = False) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path if write else f"file:{self.db_path}?mode=ro", uri=not write)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")
        return conn

    @staticmethod
    def _rows(rows: Iterable[sqlite3.Row]) -> list[dict[str, Any]]:
        return [dict(row) for row in rows]

    @staticmethod
    def _audit(conn: sqlite3.Connection, action: str, entity: str, identifier: str, details: dict[str, Any]) -> None:
        conn.execute(
            "INSERT INTO audit_events(event_at,actor,action,entity,identifier,details_json) VALUES(?,?,?,?,?,?)",
            (datetime.now(timezone.utc).isoformat(), "agent via MCP", action, entity, str(identifier), json.dumps(details, sort_keys=True)),
        )

    def audit_log(self, entity: str | None = None, identifier: str | None = None, limit: int = 250) -> list[dict[str, Any]]:
        clauses, params = [], []
        if entity:
            clauses.append("entity=?"); params.append(entity)
        if identifier:
            clauses.append("identifier=?"); params.append(str(identifier))
        where = " WHERE " + " AND ".join(clauses) if clauses else ""
        params.append(min(max(limit, 1), 1000))
        with self._connect() as conn:
            rows = self._rows(conn.execute(f"SELECT * FROM audit_events{where} ORDER BY id DESC LIMIT ?", params))
        for row in rows:
            row["details"] = json.loads(row.pop("details_json"))
        return rows

    def company_profile(self) -> dict[str, Any]:
        with self._connect() as conn:
            company = dict(conn.execute("SELECT * FROM company LIMIT 1").fetchone())
            company["departments"] = self._rows(
                conn.execute("SELECT code, name, branch FROM departments ORDER BY code")
            )
            company["reporting_cutoff"] = "2026-06-30"
            company["closed_through"] = "2026-05-31"
            company["system_history_begins"] = company["founded_date"]
            company["system_history_note"] = (
                "Detailed accounting and operating history is available from incorporation on "
                f"{company['founded_date']} through the reporting cutoff."
            )
            return company

    def list_reports(self) -> list[dict[str, str]]:
        return [
            {"name": "profit_and_loss", "grain": "month, department, account"},
            {"name": "balance_sheet", "grain": "account as of date"},
            {"name": "trial_balance", "grain": "account for date range"},
            {"name": "general_ledger", "grain": "posted journal line"},
            {"name": "project_summary", "grain": "job through as-of date"},
            {"name": "job_cost_detail", "grain": "job cost transaction"},
            {"name": "ar_aging", "grain": "customer invoice"},
            {"name": "ap_aging", "grain": "vendor bill"},
            {"name": "open_purchase_orders", "grain": "PO line"},
            {"name": "cash_balances", "grain": "bank account"},
            {"name": "bank_activity", "grain": "bank transaction"},
            {"name": "service_work_orders", "grain": "service dispatch work order"},
            {"name": "timecard_detail", "grain": "employee/day/job or work order"},
            {"name": "payroll_register", "grain": "payroll run/employee"},
            {"name": "fixed_asset_register", "grain": "fixed asset"},
            {"name": "prepaid_schedule", "grain": "prepaid item"},
            {"name": "project_change_orders", "grain": "project/change number"},
            {"name": "service_agreements", "grain": "customer/site agreement"},
        ]

    def list_accounts(self, account_type: str | None = None) -> list[dict[str, Any]]:
        clause, params = (" WHERE account_type=?", [account_type]) if account_type else ("", [])
        with self._connect() as conn:
            return self._rows(conn.execute(f"SELECT code,name,account_type,normal_balance,is_cash FROM accounts{clause} ORDER BY code", params))

    def list_master_data(self, entity: str, search: str | None = None, limit: int = 250) -> list[dict[str, Any]]:
        specs = {
            "customers": ("customers", "id", "name"), "vendors": ("vendors", "id", "name"),
            "employees": ("employees", "id", "full_name"), "projects": ("projects", "id", "name"),
            "departments": ("departments", "code", "name"), "cost_codes": ("cost_codes", "code", "name"),
            "service_agreements": ("service_agreements", "id", "site_name"),
            "fixed_assets": ("fixed_assets", "id", "description"),
            "prepaid_items": ("prepaid_items", "id", "description"),
        }
        if entity not in specs: raise ValueError(f"entity must be one of {sorted(specs)}")
        table, key, label = specs[entity]
        params: list[Any] = []
        where = ""
        if search:
            where = f" WHERE lower({key} || ' ' || {label}) LIKE ?"
            params.append(f"%{search.lower()}%")
        params.append(min(max(limit, 1), 1000))
        with self._connect() as conn:
            return self._rows(conn.execute(f"SELECT * FROM {table}{where} ORDER BY {label} LIMIT ?", params))

    def query_transactions(self, entity: str, start_date: str | None = None, end_date: str | None = None,
                           status: str | None = None, counterparty_id: str | None = None,
                           project_id: str | None = None, limit: int = 250, offset: int = 0) -> list[dict[str, Any]]:
        specs = {
            "invoices": ("customer_invoices", "invoice_date", "customer_id", "status", "project_id"),
            "bills": ("vendor_bills", "invoice_date", "vendor_id", "status", "project_id"),
            "purchase_orders": ("purchase_orders", "order_date", "vendor_id", "status", "project_id"),
            "journals": ("journal_headers", "posting_date", None, "status", None),
            "work_orders": ("service_work_orders", "date(completed_at)", "customer_id", "status", None),
            "timecards": ("timecard_entries", "work_date", "employee_id", "status", "project_id"),
            "payroll_runs": ("payroll_runs", "check_date", None, "status", None),
            "bank_transactions": ("bank_transactions", "posted_date", None, None, None),
            "change_orders": ("project_change_orders", "submitted_date", None, "status", "project_id"),
        }
        if entity not in specs: raise ValueError(f"entity must be one of {sorted(specs)}")
        table, date_col, party_col, status_col, project_col = specs[entity]
        clauses, params = [], []
        if start_date: clauses.append(f"{date_col}>=?"); params.append(start_date)
        if end_date: clauses.append(f"{date_col}<=?"); params.append(end_date)
        if status and status_col: clauses.append(f"{status_col}=?"); params.append(status)
        if counterparty_id and party_col: clauses.append(f"{party_col}=?"); params.append(counterparty_id)
        if project_id and project_col: clauses.append(f"{project_col}=?"); params.append(project_id)
        where = " WHERE " + " AND ".join(clauses) if clauses else ""
        params.extend([min(max(limit, 1), 1000), max(offset, 0)])
        with self._connect() as conn:
            return self._rows(conn.execute(f"SELECT * FROM {table}{where} ORDER BY {date_col} DESC,id DESC LIMIT ? OFFSET ?", params))

    def get_transaction(self, entity: str, identifier: str) -> dict[str, Any] | None:
        specs = {
            "invoice": ("customer_invoices", "invoice_number"), "bill": ("vendor_bills", "bill_number"),
            "purchase_order": ("purchase_orders", "po_number"), "journal": ("journal_headers", "id"),
            "work_order": ("service_work_orders", "id"), "timecard": ("timecard_entries", "id"),
            "payroll_run": ("payroll_runs", "id"), "bank_transaction": ("bank_transactions", "id"),
            "fixed_asset": ("fixed_assets", "id"), "prepaid_item": ("prepaid_items", "id"),
            "change_order": ("project_change_orders", "id"), "service_agreement": ("service_agreements", "id"),
        }
        if entity not in specs: raise ValueError(f"entity must be one of {sorted(specs)}")
        table, key = specs[entity]
        with self._connect() as conn:
            row = conn.execute(f"SELECT * FROM {table} WHERE {key}=?", (identifier,)).fetchone()
            if not row: return None
            result = dict(row)
            if entity == "invoice": result["lines"] = self._rows(conn.execute("SELECT * FROM customer_invoice_lines WHERE invoice_id=? ORDER BY line_number", (row["id"],)))
            if entity == "bill": result["lines"] = self._rows(conn.execute("SELECT * FROM vendor_bill_lines WHERE bill_id=? ORDER BY line_number", (row["id"],)))
            if entity == "purchase_order": result["lines"] = self._rows(conn.execute("SELECT * FROM purchase_order_lines WHERE po_id=? ORDER BY line_number", (row["id"],)))
            if entity == "journal": result["lines"] = self._rows(conn.execute("SELECT * FROM journal_lines WHERE journal_id=? ORDER BY id", (row["id"],)))
            if entity == "payroll_run": result["lines"] = self._rows(conn.execute("SELECT * FROM payroll_run_lines WHERE payroll_run_id=? ORDER BY id", (row["id"],)))
            return result

    def create_journal(self, posting_date: str, reference: str, memo: str, lines: list[dict[str, Any]], post: bool = False) -> dict[str, Any]:
        if len(lines) < 2: raise ValueError("A journal requires at least two lines")
        debit = round(sum(float(x.get("debit", 0)) for x in lines), 2); credit = round(sum(float(x.get("credit", 0)) for x in lines), 2)
        if post and abs(debit-credit) > .01: raise ValueError("Posted journal must balance")
        with self._connect(write=True) as conn:
            cur = conn.execute("INSERT INTO journal_headers(posting_date,source,reference,memo,status,created_at) VALUES(?,?,?,?,?,datetime('now'))", (posting_date,"MCP",reference,memo,"Posted" if post else "Draft"))
            jid = cur.lastrowid
            for line in lines:
                conn.execute("INSERT INTO journal_lines(journal_id,account_code,department_code,project_id,cost_code,debit,credit,line_memo) VALUES(?,?,?,?,?,?,?,?)",
                             (jid,line["account_code"],line.get("department_code"),line.get("project_id"),line.get("cost_code"),float(line.get("debit",0)),float(line.get("credit",0)),line.get("line_memo",memo)))
            self._audit(conn, "create", "journal", str(jid), {"reference": reference, "status": "Posted" if post else "Draft", "debit": debit, "credit": credit})
            conn.commit()
        return self.get_transaction("journal", str(jid)) or {}

    def post_or_void_journal(self, journal_id: int, action: str) -> dict[str, Any]:
        if action not in {"post", "void"}: raise ValueError("action must be post or void")
        with self._connect(write=True) as conn:
            row = conn.execute("SELECT status FROM journal_headers WHERE id=?", (journal_id,)).fetchone()
            if not row: raise ValueError("journal not found")
            if action == "post":
                totals = conn.execute("SELECT ROUND(SUM(debit),2),ROUND(SUM(credit),2) FROM journal_lines WHERE journal_id=?", (journal_id,)).fetchone()
                if abs((totals[0] or 0)-(totals[1] or 0)) > .01: raise ValueError("journal does not balance")
                status = "Posted"
            else: status = "Voided"
            conn.execute("UPDATE journal_headers SET status=? WHERE id=?", (status,journal_id))
            self._audit(conn, action, "journal", str(journal_id), {"new_status": status})
            conn.commit()
        return self.get_transaction("journal", str(journal_id)) or {}

    def create_invoice(self, invoice: dict[str, Any]) -> dict[str, Any]:
        amount = round(float(invoice.get("amount", 0)), 2)
        if amount <= 0: raise ValueError("invoice amount must be positive")
        project_id = invoice.get("project_id")
        with self._connect(write=True) as conn:
            project = conn.execute("SELECT department_code FROM projects WHERE id=?", (project_id,)).fetchone() if project_id else None
            default_department = invoice.get("department_code") or (project["department_code"] if project else "30")
            line_items = invoice.get("line_items") or [{
                "account_code": invoice.get("revenue_account_code") or ("4000" if project_id else "4100"),
                "amount": amount,
                "department_code": default_department,
                "project_id": project_id,
                "description": invoice.get("description", "Customer invoice"),
            }]
            line_total = round(sum(float(line.get("amount", 0)) for line in line_items), 2)
            if abs(line_total - amount) > .01: raise ValueError(f"invoice lines must total invoice amount: {line_total} vs {amount}")
            for line in line_items:
                account = conn.execute("SELECT account_type FROM accounts WHERE code=?", (line["account_code"],)).fetchone()
                if not account or account["account_type"] != "Revenue": raise ValueError(f"invoice line account must be a Revenue account: {line['account_code']}")
            vals = [invoice.get(f) for f in ("invoice_number","customer_id","project_id","invoice_date","due_date")]
            vals.extend([amount, float(invoice.get("retainage_amount", 0)), invoice.get("description", "Customer invoice")])
            invoice_cur = conn.execute("INSERT INTO customer_invoices(invoice_number,customer_id,project_id,invoice_date,due_date,amount,paid_amount,retainage_amount,status,description) VALUES(?,?,?,?,?,?,0,?,'Open',?)", vals)
            cur = conn.execute(
                "INSERT INTO journal_headers(posting_date,source,reference,memo,status,created_at) VALUES(?,?,?,?, 'Posted',datetime('now'))",
                (invoice["invoice_date"], "MCP Invoice", invoice["invoice_number"], invoice.get("description", "Customer invoice")),
            )
            journal_id = cur.lastrowid
            conn.execute(
                "INSERT INTO journal_lines(journal_id,account_code,department_code,project_id,debit,credit,line_memo) VALUES(?,?,?,?,?,?,?)",
                (journal_id, "1100", default_department, project_id, amount, 0, f"Invoice {invoice['invoice_number']}"),
            )
            for line_number, line in enumerate(line_items, 1):
                conn.execute(
                    """INSERT INTO customer_invoice_lines(invoice_id,line_number,account_code,department_code,project_id,
                       service_work_order_id,description,quantity,unit_price,amount) VALUES(?,?,?,?,?,?,?,?,?,?)""",
                    (invoice_cur.lastrowid, line.get("line_number", line_number), line["account_code"],
                     line.get("department_code", default_department), line.get("project_id", project_id),
                     line.get("service_work_order_id"), line.get("description", invoice.get("description", "Customer invoice")),
                     float(line.get("quantity", 1)), float(line.get("unit_price", line["amount"])), float(line["amount"])),
                )
                conn.execute(
                    "INSERT INTO journal_lines(journal_id,account_code,department_code,project_id,debit,credit,line_memo) VALUES(?,?,?,?,?,?,?)",
                    (journal_id, line["account_code"], line.get("department_code", default_department), line.get("project_id", project_id), 0, float(line["amount"]), line.get("description", invoice.get("description", "Customer invoice"))),
                )
            self._audit(conn, "create", "invoice", invoice["invoice_number"], {"amount": amount, "journal_id": journal_id, "line_count": len(line_items)})
            conn.commit()
        result = self.get_transaction("invoice", invoice["invoice_number"]) or {}
        result["posting_journal_id"] = journal_id
        return result

    def create_bill(self, bill: dict[str, Any]) -> dict[str, Any]:
        amount = round(float(bill.get("amount", 0)), 2)
        if amount <= 0: raise ValueError("bill amount must be positive")
        project_id = bill.get("project_id")
        with self._connect(write=True) as conn:
            project = conn.execute("SELECT department_code FROM projects WHERE id=?", (project_id,)).fetchone() if project_id else None
            default_department = bill.get("department_code") or (project["department_code"] if project else "90")
            default_cost_code = bill.get("cost_code")
            default_account = bill.get("expense_account_code")
            if not default_account and default_cost_code:
                mapped = conn.execute("SELECT gl_account_code FROM cost_codes WHERE code=?", (default_cost_code,)).fetchone()
                if not mapped: raise ValueError(f"unknown cost code: {default_cost_code}")
                default_account = mapped["gl_account_code"]
            line_items = bill.get("line_items") or [{
                "account_code": default_account or "6900",
                "amount": amount,
                "department_code": default_department,
                "project_id": project_id,
                "cost_code": default_cost_code,
                "description": bill.get("description", "Vendor bill"),
            }]
            line_total = round(sum(float(line.get("amount", 0)) for line in line_items), 2)
            if abs(line_total - amount) > .01: raise ValueError(f"bill lines must total bill amount: {line_total} vs {amount}")
            for line in line_items:
                account = conn.execute("SELECT account_type FROM accounts WHERE code=?", (line["account_code"],)).fetchone()
                if not account or account["account_type"] not in {"Cost of Revenue", "Operating Expense", "Asset", "Other Expense"}:
                    raise ValueError(f"unsupported bill line account: {line['account_code']}")
                if line.get("cost_code"):
                    mapped = conn.execute("SELECT gl_account_code FROM cost_codes WHERE code=?", (line["cost_code"],)).fetchone()
                    if not mapped or mapped["gl_account_code"] != line["account_code"]:
                        raise ValueError(f"cost code {line['cost_code']} must post to {mapped['gl_account_code'] if mapped else 'a valid mapped account'}")
            vals = [bill.get(f) for f in ("bill_number","vendor_id","project_id","invoice_date","due_date")]
            vals.extend([amount, int(bool(bill.get("payment_hold", 0))), bill.get("hold_reason"), bill.get("description", "Vendor bill")])
            bill_cur = conn.execute("INSERT INTO vendor_bills(bill_number,vendor_id,project_id,invoice_date,due_date,amount,paid_amount,payment_hold,hold_reason,status,description) VALUES(?,?,?,?,?,?,0,?,?,'Open',?)", vals)
            cur = conn.execute(
                "INSERT INTO journal_headers(posting_date,source,reference,memo,status,created_at) VALUES(?,?,?,?, 'Posted',datetime('now'))",
                (bill["invoice_date"], "MCP Bill", bill["bill_number"], bill.get("description", "Vendor bill")),
            )
            journal_id = cur.lastrowid
            for line_number, line in enumerate(line_items, 1):
                line_project = line.get("project_id", project_id)
                line_cost_code = line.get("cost_code")
                conn.execute(
                    """INSERT INTO vendor_bill_lines(bill_id,line_number,account_code,department_code,project_id,cost_code,
                       description,quantity,unit_cost,amount) VALUES(?,?,?,?,?,?,?,?,?,?)""",
                    (bill_cur.lastrowid, line.get("line_number", line_number), line["account_code"],
                     line.get("department_code", default_department), line_project, line_cost_code,
                     line.get("description", bill.get("description", "Vendor bill")), float(line.get("quantity", line.get("units", 1))),
                     float(line.get("unit_cost", line["amount"])), float(line["amount"])),
                )
                conn.execute(
                    "INSERT INTO journal_lines(journal_id,account_code,department_code,project_id,cost_code,debit,credit,line_memo) VALUES(?,?,?,?,?,?,?,?)",
                    (journal_id, line["account_code"], line.get("department_code", default_department), line_project, line_cost_code, float(line["amount"]), 0, line.get("description", bill.get("description", "Vendor bill"))),
                )
                if line_project and line_cost_code:
                    conn.execute(
                        "INSERT INTO job_cost_entries(posting_date,document_number,source,project_id,cost_code,vendor_id,employee_id,units,amount,description) VALUES(?,?,?,?,?,?,NULL,?,?,?)",
                        (bill["invoice_date"], bill["bill_number"], "MCP Bill", line_project, line_cost_code, bill["vendor_id"], float(line.get("units", 1)), float(line["amount"]), line.get("description", bill.get("description", "Vendor bill"))),
                    )
            conn.execute(
                "INSERT INTO journal_lines(journal_id,account_code,department_code,project_id,debit,credit,line_memo) VALUES(?,?,?,?,?,?,?)",
                (journal_id, "2000", default_department, project_id, 0, amount, f"Bill {bill['bill_number']}"),
            )
            self._audit(conn, "create", "bill", bill["bill_number"], {"amount": amount, "journal_id": journal_id, "line_count": len(line_items)})
            conn.commit()
        result = self.get_transaction("bill", bill["bill_number"]) or {}
        result["posting_journal_id"] = journal_id
        return result

    def create_purchase_order(self, purchase_order: dict[str, Any], lines: list[dict[str, Any]]) -> dict[str, Any]:
        if not lines: raise ValueError("purchase order requires at least one line")
        with self._connect(write=True) as conn:
            cur = conn.execute("INSERT INTO purchase_orders(po_number,vendor_id,project_id,order_date,expected_date,status,buyer) VALUES(?,?,?,?,?,'Open',?)",
                               tuple(purchase_order.get(f) for f in ("po_number","vendor_id","project_id","order_date","expected_date","buyer")))
            po_id=cur.lastrowid
            for i,line in enumerate(lines,1):
                conn.execute("INSERT INTO purchase_order_lines(po_id,line_number,cost_code,description,committed_amount,invoiced_amount) VALUES(?,?,?,?,?,?)",
                             (po_id,line.get("line_number",i),line["cost_code"],line["description"],float(line["committed_amount"]),float(line.get("invoiced_amount",0))))
            self._audit(conn, "create", "purchase_order", purchase_order["po_number"], {"line_count": len(lines), "committed_amount": sum(float(line["committed_amount"]) for line in lines)})
            conn.commit()
        return self.get_transaction("purchase_order",purchase_order["po_number"]) or {}

    def update_transaction(self, entity: str, identifier: str, changes: dict[str, Any]) -> dict[str, Any]:
        specs = {
            "invoice": ("customer_invoices","invoice_number",{"due_date","description","retainage_amount","status"}),
            "bill": ("vendor_bills","bill_number",{"due_date","description","payment_hold","hold_reason","status"}),
            "purchase_order": ("purchase_orders","po_number",{"expected_date","buyer","status"}),
            "journal": ("journal_headers","id",{"posting_date","reference","memo"}),
        }
        if entity not in specs: raise ValueError(f"entity must be one of {sorted(specs)}")
        table,key,allowed=specs[entity]
        bad=set(changes)-allowed
        if bad: raise ValueError(f"unsupported fields: {sorted(bad)}; allowed: {sorted(allowed)}")
        if not changes: return self.get_transaction(entity,identifier) or {}
        with self._connect(write=True) as conn:
            if entity=="journal":
                status=conn.execute("SELECT status FROM journal_headers WHERE id=?",(identifier,)).fetchone()
                if not status or status[0] != "Draft": raise ValueError("only draft journals can be edited")
            sets=",".join(f"{f}=?" for f in changes); params=list(changes.values())+[identifier]
            cur=conn.execute(f"UPDATE {table} SET {sets} WHERE {key}=?",params)
            if not cur.rowcount: raise ValueError("document not found")
            self._audit(conn, "update", entity, identifier, {"changes": changes})
            conn.commit()
        return self.get_transaction(entity,identifier) or {}

    def apply_payment(self, entity: str, document_number: str, amount: float, payment_date: str, cash_account: str = "1000") -> dict[str, Any]:
        if amount <= 0: raise ValueError("amount must be positive")
        if entity == "invoice": table,key,ar_acct = "customer_invoices","invoice_number","1100"
        elif entity == "bill": table,key,ar_acct = "vendor_bills","bill_number","2000"
        else: raise ValueError("entity must be invoice or bill")
        with self._connect(write=True) as conn:
            row = conn.execute(f"SELECT id,amount,paid_amount,status FROM {table} WHERE {key}=?", (document_number,)).fetchone()
            if not row: raise ValueError("document not found")
            if row["paid_amount"] + amount > row["amount"] + .01: raise ValueError("payment exceeds open amount")
            new_paid = row["paid_amount"] + amount; status = "Paid" if abs(new_paid-row["amount"]) <= .01 else "Partially Paid"
            conn.execute(f"UPDATE {table} SET paid_amount=?,status=? WHERE id=?", (new_paid,status,row["id"]))
            cur = conn.execute("INSERT INTO journal_headers(posting_date,source,reference,memo,status,created_at) VALUES(?,?,?,?, 'Posted',datetime('now'))", (payment_date,"MCP Payment",f"PAY-{document_number}",f"Payment applied to {document_number}"))
            jid = cur.lastrowid
            if entity == "invoice": lines=[(cash_account,amount,0),(ar_acct,0,amount)]
            else: lines=[(ar_acct,amount,0),(cash_account,0,amount)]
            for acct,dr,cr in lines: conn.execute("INSERT INTO journal_lines(journal_id,account_code,debit,credit,line_memo) VALUES(?,?,?,?,?)", (jid,acct,dr,cr,f"Payment {document_number}"))
            self._audit(conn, "apply_payment", entity, document_number, {"amount": amount, "payment_date": payment_date, "cash_account": cash_account, "journal_id": jid})
            conn.commit()
        return self.get_transaction(entity, document_number) or {}

    def update_document_status(self, entity: str, document_number: str, status: str, note: str | None = None) -> dict[str, Any]:
        specs = {"invoice":("customer_invoices","invoice_number"),"bill":("vendor_bills","bill_number"),"purchase_order":("purchase_orders","po_number")}
        if entity not in specs: raise ValueError(f"entity must be one of {sorted(specs)}")
        table,key=specs[entity]
        with self._connect(write=True) as conn:
            cur=conn.execute(f"UPDATE {table} SET status=? WHERE {key}=?",(status,document_number))
            if not cur.rowcount: raise ValueError("document not found")
            if entity=="bill" and note is not None: conn.execute("UPDATE vendor_bills SET hold_reason=? WHERE bill_number=?",(note,document_number))
            self._audit(conn, "status_update", entity, document_number, {"status": status, "note": note})
            conn.commit()
        return self.get_transaction(entity,document_number) or {}

    def trial_balance(self, start_date: str, end_date: str) -> list[dict[str, Any]]:
        with self._connect() as conn:
            return self._rows(
                conn.execute(
                    """
                    SELECT a.code AS account_code, a.name AS account_name,
                           a.account_type,
                           ROUND(SUM(jl.debit), 2) AS debit,
                           ROUND(SUM(jl.credit), 2) AS credit,
                           ROUND(SUM(jl.debit - jl.credit), 2) AS net_debit
                    FROM journal_lines jl
                    JOIN journal_headers jh ON jh.id = jl.journal_id
                    JOIN accounts a ON a.code = jl.account_code
                    WHERE jh.status = 'Posted' AND jh.posting_date BETWEEN ? AND ?
                    GROUP BY a.code, a.name, a.account_type
                    HAVING ABS(SUM(jl.debit - jl.credit)) > 0.004
                    ORDER BY a.code
                    """,
                    (start_date, end_date),
                )
            )

    def profit_and_loss(
        self, start_date: str, end_date: str, group_by: str = "month"
    ) -> list[dict[str, Any]]:
        allowed = {
            "month": "substr(jh.posting_date,1,7)",
            "department": "COALESCE(d.name,'Unassigned')",
            "account": "a.code || ' - ' || a.name",
        }
        if group_by not in allowed:
            raise ValueError(f"group_by must be one of {sorted(allowed)}")
        expr = allowed[group_by]
        with self._connect() as conn:
            return self._rows(
                conn.execute(
                    f"""
                    SELECT {expr} AS grouping, a.account_type,
                           ROUND(SUM(CASE WHEN a.account_type = 'Revenue'
                               THEN jl.credit - jl.debit ELSE jl.debit - jl.credit END), 2) AS amount
                    FROM journal_lines jl
                    JOIN journal_headers jh ON jh.id = jl.journal_id
                    JOIN accounts a ON a.code = jl.account_code
                    LEFT JOIN departments d ON d.code = jl.department_code
                    WHERE jh.status='Posted' AND jh.posting_date BETWEEN ? AND ?
                      AND a.account_type IN ('Revenue','Cost of Revenue','Operating Expense','Other Income','Other Expense')
                    GROUP BY grouping, a.account_type
                    ORDER BY grouping, a.account_type
                    """,
                    (start_date, end_date),
                )
            )

    def balance_sheet(self, as_of_date: str) -> list[dict[str, Any]]:
        with self._connect() as conn:
            return self._rows(
                conn.execute(
                    """
                    SELECT a.code AS account_code, a.name AS account_name, a.account_type,
                           ROUND(SUM(CASE WHEN a.normal_balance='Debit'
                               THEN jl.debit-jl.credit ELSE jl.credit-jl.debit END),2) AS balance
                    FROM journal_lines jl
                    JOIN journal_headers jh ON jh.id=jl.journal_id
                    JOIN accounts a ON a.code=jl.account_code
                    WHERE jh.status='Posted' AND jh.posting_date <= ?
                      AND a.account_type IN ('Asset','Liability','Equity')
                    GROUP BY a.code,a.name,a.account_type
                    HAVING ABS(balance) > 0.004
                    ORDER BY a.code
                    """,
                    (as_of_date,),
                )
            )

    def general_ledger(
        self,
        start_date: str,
        end_date: str,
        account_code: str | None = None,
        project_id: str | None = None,
        limit: int = 1000,
        offset: int = 0,
    ) -> list[dict[str, Any]]:
        clauses = ["jh.status='Posted'", "jh.posting_date BETWEEN ? AND ?"]
        params: list[Any] = [start_date, end_date]
        if account_code:
            clauses.append("jl.account_code = ?")
            params.append(account_code)
        if project_id:
            clauses.append("jl.project_id = ?")
            params.append(project_id)
        params.extend([min(max(limit, 1), 5000), max(offset, 0)])
        with self._connect() as conn:
            return self._rows(
                conn.execute(
                    f"""
                    SELECT jh.posting_date, jh.source, jh.reference, jh.memo,
                           jl.account_code, a.name AS account_name,
                           jl.department_code, jl.project_id, jl.cost_code,
                           ROUND(jl.debit,2) AS debit, ROUND(jl.credit,2) AS credit,
                           jl.line_memo
                    FROM journal_lines jl
                    JOIN journal_headers jh ON jh.id=jl.journal_id
                    JOIN accounts a ON a.code=jl.account_code
                    WHERE {' AND '.join(clauses)}
                    ORDER BY jh.posting_date, jh.id, jl.id
                    LIMIT ? OFFSET ?
                    """,
                    params,
                )
            )

    def project_summary(self, project_id: str | None, as_of_date: str) -> list[dict[str, Any]]:
        params: list[Any] = [as_of_date, as_of_date]
        clause = ""
        if project_id:
            clause = "AND p.id = ?"
            params.append(project_id)
        with self._connect() as conn:
            return self._rows(
                conn.execute(
                    f"""
                    SELECT p.id AS project_id, p.name AS project_name, c.name AS customer,
                           p.branch, p.project_manager, p.status, p.start_date, p.expected_end_date,
                           ROUND(p.original_contract + p.approved_change_orders,2) AS current_contract,
                           ROUND(p.original_estimated_cost + p.approved_cost_budget_changes,2) AS current_cost_budget,
                           ROUND(COALESCE((SELECT SUM(j.amount) FROM job_cost_entries j
                               WHERE j.project_id=p.id AND j.posting_date<=?),0),2) AS cost_to_date,
                           ROUND(COALESCE((SELECT SUM(i.amount) FROM customer_invoices i
                               WHERE i.project_id=p.id AND i.invoice_date<=? AND i.status<>'Voided'),0),2) AS billings_to_date
                    FROM projects p JOIN customers c ON c.id=p.customer_id
                    WHERE p.status IN ('Active','Substantially Complete','Warranty') {clause}
                    ORDER BY p.id
                    """,
                    params,
                )
            )

    def job_cost_detail(
        self,
        project_id: str,
        as_of_date: str,
        limit: int = 5000,
        offset: int = 0,
    ) -> list[dict[str, Any]]:
        with self._connect() as conn:
            return self._rows(
                conn.execute(
                    """
                    SELECT j.posting_date, j.document_number, j.source,
                           j.cost_code, cc.name AS cost_code_name, j.vendor_id,
                           j.employee_id, ROUND(j.units,2) AS units,
                           ROUND(j.amount,2) AS amount, j.description
                    FROM job_cost_entries j JOIN cost_codes cc ON cc.code=j.cost_code
                    WHERE j.project_id=? AND j.posting_date<=?
                    ORDER BY j.posting_date,j.id LIMIT ? OFFSET ?
                    """,
                    (project_id, as_of_date, min(max(limit, 1), 10000), max(offset, 0)),
                )
            )

    def ar_aging(self, as_of_date: str) -> list[dict[str, Any]]:
        with self._connect() as conn:
            return self._rows(
                conn.execute(
                    """
                    SELECT i.invoice_number, i.project_id, c.name AS customer,
                           i.invoice_date, i.due_date, ROUND(i.amount,2) AS original_amount,
                           ROUND(i.amount-i.paid_amount,2) AS open_amount,
                           CAST(julianday(?) - julianday(i.due_date) AS INTEGER) AS days_past_due,
                           CASE WHEN date(?) <= date(i.due_date) THEN 'Current'
                                WHEN julianday(?) - julianday(i.due_date) <= 30 THEN '1-30'
                                WHEN julianday(?) - julianday(i.due_date) <= 60 THEN '31-60'
                                WHEN julianday(?) - julianday(i.due_date) <= 90 THEN '61-90'
                                ELSE '90+' END AS aging_bucket,
                           i.retainage_amount, i.status
                    FROM customer_invoices i JOIN customers c ON c.id=i.customer_id
                    WHERE i.invoice_date<=? AND i.amount-i.paid_amount>0.004 AND i.status<>'Voided'
                    ORDER BY i.due_date,i.invoice_number
                    """,
                    (as_of_date, as_of_date, as_of_date, as_of_date, as_of_date, as_of_date),
                )
            )

    def ap_aging(self, as_of_date: str) -> list[dict[str, Any]]:
        with self._connect() as conn:
            return self._rows(
                conn.execute(
                    """
                    SELECT b.bill_number, v.name AS vendor, b.project_id, b.invoice_date,
                           b.due_date, ROUND(b.amount,2) AS original_amount,
                           ROUND(b.amount-b.paid_amount,2) AS open_amount,
                           CAST(julianday(?) - julianday(b.due_date) AS INTEGER) AS days_past_due,
                           b.payment_hold, b.hold_reason, b.status
                    FROM vendor_bills b JOIN vendors v ON v.id=b.vendor_id
                    WHERE b.invoice_date<=? AND b.amount-b.paid_amount>0.004 AND b.status<>'Voided'
                    ORDER BY b.payment_hold DESC,b.due_date,b.bill_number
                    """,
                    (as_of_date, as_of_date),
                )
            )

    def open_purchase_orders(self, as_of_date: str) -> list[dict[str, Any]]:
        with self._connect() as conn:
            return self._rows(
                conn.execute(
                    """
                    SELECT po.po_number, v.name AS vendor, po.project_id, po.order_date,
                           po.expected_date, po.status, pol.line_number, pol.description,
                           ROUND(pol.committed_amount,2) AS committed_amount,
                           ROUND(pol.invoiced_amount,2) AS invoiced_amount,
                           ROUND(pol.committed_amount-pol.invoiced_amount,2) AS remaining_commitment
                    FROM purchase_orders po
                    JOIN purchase_order_lines pol ON pol.po_id=po.id
                    JOIN vendors v ON v.id=po.vendor_id
                    WHERE po.order_date<=? AND po.status IN ('Open','Partially Received')
                      AND pol.committed_amount-pol.invoiced_amount>0.004
                    ORDER BY po.expected_date,po.po_number,pol.line_number
                    """,
                    (as_of_date,),
                )
            )

    def cash_balances(self, as_of_date: str) -> list[dict[str, Any]]:
        with self._connect() as conn:
            return self._rows(
                conn.execute(
                    """
                    SELECT a.code AS account_code, a.name AS bank_account,
                           ROUND(SUM(jl.debit-jl.credit),2) AS ledger_balance
                    FROM journal_lines jl JOIN journal_headers jh ON jh.id=jl.journal_id
                    JOIN accounts a ON a.code=jl.account_code
                    WHERE a.is_cash=1 AND jh.status='Posted' AND jh.posting_date<=?
                    GROUP BY a.code,a.name ORDER BY a.code
                    """,
                    (as_of_date,),
                )
            )

    def bank_activity(self, start_date: str, end_date: str, account_code: str | None = None,
                      cleared: bool | None = None, limit: int = 5000,
                      offset: int = 0) -> list[dict[str, Any]]:
        clauses = ["b.posted_date BETWEEN ? AND ?"]
        params: list[Any] = [start_date, end_date]
        if account_code:
            clauses.append("b.bank_account_code=?"); params.append(account_code)
        if cleared is not None:
            clauses.append("b.cleared=?"); params.append(int(cleared))
        params.extend([min(max(limit, 1), 10000), max(offset, 0)])
        with self._connect() as conn:
            return self._rows(conn.execute(
                f"""SELECT b.id,b.bank_account_code,a.name AS bank_account,b.transaction_date,b.posted_date,
                           b.transaction_type,b.reference,b.payee_payer,ROUND(b.amount,2) AS amount,
                           b.journal_id,b.cleared,b.description
                    FROM bank_transactions b JOIN accounts a ON a.code=b.bank_account_code
                    WHERE {' AND '.join(clauses)} ORDER BY b.posted_date,b.id LIMIT ? OFFSET ?""", params))

    def service_work_orders(self, start_date: str, end_date: str, customer_id: str | None = None,
                            technician_id: str | None = None, status: str | None = None,
                            limit: int = 5000, offset: int = 0) -> list[dict[str, Any]]:
        clauses = ["date(w.opened_at)<=?", "date(COALESCE(w.completed_at,w.opened_at))>=?"]
        params: list[Any] = [end_date, start_date]
        if customer_id:
            clauses.append("w.customer_id=?"); params.append(customer_id)
        if technician_id:
            clauses.append("w.technician_id=?"); params.append(technician_id)
        if status:
            clauses.append("w.status=?"); params.append(status)
        params.extend([min(max(limit, 1), 10000), max(offset, 0)])
        with self._connect() as conn:
            return self._rows(conn.execute(
                f"""SELECT w.id,w.agreement_id,w.customer_id,c.name AS customer,w.opened_at,w.completed_at,
                           w.work_type,w.priority,w.technician_id,e.full_name AS technician,w.branch,w.status,
                           w.callback,ROUND(w.billable_hours,2) AS billable_hours,
                           ROUND(w.labor_amount,2) AS labor_amount,ROUND(w.material_amount,2) AS material_amount,
                           w.invoice_number,w.description
                    FROM service_work_orders w JOIN customers c ON c.id=w.customer_id
                    LEFT JOIN employees e ON e.id=w.technician_id
                    WHERE {' AND '.join(clauses)} ORDER BY w.opened_at,w.id LIMIT ? OFFSET ?""", params))

    def timecard_detail(self, start_date: str, end_date: str, employee_id: str | None = None,
                        project_id: str | None = None, service_work_order_id: str | None = None,
                        status: str | None = None, limit: int = 5000,
                        offset: int = 0) -> list[dict[str, Any]]:
        clauses = ["t.work_date BETWEEN ? AND ?"]
        params: list[Any] = [start_date, end_date]
        for column, value in (("t.employee_id", employee_id), ("t.project_id", project_id),
                              ("t.service_work_order_id", service_work_order_id), ("t.status", status)):
            if value:
                clauses.append(f"{column}=?"); params.append(value)
        params.extend([min(max(limit, 1), 10000), max(offset, 0)])
        with self._connect() as conn:
            return self._rows(conn.execute(
                f"""SELECT t.id,t.work_date,t.employee_id,e.full_name AS employee,t.project_id,
                           t.service_work_order_id,t.cost_code,ROUND(t.regular_hours,2) AS regular_hours,
                           ROUND(t.overtime_hours,2) AS overtime_hours,ROUND(t.loaded_labor_cost,2) AS loaded_labor_cost,
                           t.payroll_run_id,t.status,t.description
                    FROM timecard_entries t JOIN employees e ON e.id=t.employee_id
                    WHERE {' AND '.join(clauses)} ORDER BY t.work_date,t.employee_id,t.id LIMIT ? OFFSET ?""", params))

    def payroll_register(self, start_date: str, end_date: str, employee_id: str | None = None,
                         department_code: str | None = None, limit: int = 10000,
                         offset: int = 0) -> list[dict[str, Any]]:
        clauses = ["r.check_date BETWEEN ? AND ?"]
        params: list[Any] = [start_date, end_date]
        if employee_id:
            clauses.append("l.employee_id=?"); params.append(employee_id)
        if department_code:
            clauses.append("l.department_code=?"); params.append(department_code)
        params.extend([min(max(limit, 1), 20000), max(offset, 0)])
        with self._connect() as conn:
            return self._rows(conn.execute(
                f"""SELECT r.id AS payroll_run_id,r.period_start,r.period_end,r.check_date,r.status,
                           l.employee_id,e.full_name AS employee,l.department_code,
                           ROUND(l.regular_hours,2) AS regular_hours,ROUND(l.overtime_hours,2) AS overtime_hours,
                           ROUND(l.gross_pay,2) AS gross_pay,ROUND(l.employer_taxes,2) AS employer_taxes,
                           ROUND(l.benefits,2) AS benefits
                    FROM payroll_runs r JOIN payroll_run_lines l ON l.payroll_run_id=r.id
                    JOIN employees e ON e.id=l.employee_id
                    WHERE {' AND '.join(clauses)} ORDER BY r.check_date,l.employee_id LIMIT ? OFFSET ?""", params))

    def fixed_asset_register(self, as_of_date: str = "2026-06-30") -> list[dict[str, Any]]:
        with self._connect() as conn:
            return self._rows(conn.execute(
                """SELECT f.*,ROUND(f.gross_cost-f.accumulated_depreciation,2) AS net_book_value,
                          v.name AS vendor,a.name AS asset_account
                   FROM fixed_assets f LEFT JOIN vendors v ON v.id=f.vendor_id
                   JOIN accounts a ON a.code=f.account_code
                   WHERE f.acquired_date<=? AND (f.disposal_date IS NULL OR f.disposal_date>?)
                   ORDER BY f.account_code,f.acquired_date,f.id""", (as_of_date, as_of_date)))

    def prepaid_schedule(self, as_of_date: str = "2026-06-30") -> list[dict[str, Any]]:
        with self._connect() as conn:
            return self._rows(conn.execute(
                """SELECT p.*,ROUND(p.original_amount-p.cumulative_amortization,2) AS unamortized_balance,
                          v.name AS vendor,a.name AS prepaid_account
                   FROM prepaid_items p LEFT JOIN vendors v ON v.id=p.vendor_id
                   JOIN accounts a ON a.code=p.account_code
                   WHERE p.service_start<=? ORDER BY p.service_start,p.id""", (as_of_date,)))

    def project_change_orders(self, project_id: str | None = None, status: str | None = None) -> list[dict[str, Any]]:
        clauses, params = [], []
        if project_id:
            clauses.append("co.project_id=?"); params.append(project_id)
        if status:
            clauses.append("co.status=?"); params.append(status)
        where = " WHERE " + " AND ".join(clauses) if clauses else ""
        with self._connect() as conn:
            return self._rows(conn.execute(
                f"""SELECT co.*,p.name AS project_name FROM project_change_orders co
                    JOIN projects p ON p.id=co.project_id{where}
                    ORDER BY co.project_id,co.change_number""", params))

    def service_agreements(self, as_of_date: str = "2026-06-30", status: str | None = None) -> list[dict[str, Any]]:
        clauses = ["s.start_date<=?", "(s.end_date IS NULL OR s.end_date>=?)"]
        params: list[Any] = [as_of_date, as_of_date]
        if status:
            clauses.append("s.status=?"); params.append(status)
        with self._connect() as conn:
            return self._rows(conn.execute(
                f"""SELECT s.*,c.name AS customer FROM service_agreements s
                    JOIN customers c ON c.id=s.customer_id
                    WHERE {' AND '.join(clauses)} ORDER BY c.name,s.site_name""", params))

    def export_csv(self, report: str, workspace_root: Path, relative_path: str, **kwargs: Any) -> dict[str, Any]:
        methods = {
            "trial_balance": self.trial_balance,
            "profit_and_loss": self.profit_and_loss,
            "balance_sheet": self.balance_sheet,
            "general_ledger": self.general_ledger,
            "project_summary": self.project_summary,
            "job_cost_detail": self.job_cost_detail,
            "ar_aging": self.ar_aging,
            "ap_aging": self.ap_aging,
            "open_purchase_orders": self.open_purchase_orders,
            "cash_balances": self.cash_balances,
            "bank_activity": self.bank_activity,
            "service_work_orders": self.service_work_orders,
            "timecard_detail": self.timecard_detail,
            "payroll_register": self.payroll_register,
            "fixed_asset_register": self.fixed_asset_register,
            "prepaid_schedule": self.prepaid_schedule,
            "project_change_orders": self.project_change_orders,
            "service_agreements": self.service_agreements,
        }
        if report not in methods:
            raise ValueError(f"Unknown report {report!r}")
        target = (workspace_root / relative_path).resolve()
        if workspace_root.resolve() not in target.parents:
            raise ValueError("relative_path must stay inside the workspace")
        rows = methods[report](**kwargs)
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open("w", newline="", encoding="utf-8") as fh:
            writer = csv.DictWriter(fh, fieldnames=list(rows[0]) if rows else ["no_rows"])
            writer.writeheader()
            writer.writerows(rows)
        return {"path": str(target.relative_to(workspace_root)), "row_count": len(rows)}


def json_result(value: Any) -> str:
    return json.dumps(value, indent=2, default=str)
