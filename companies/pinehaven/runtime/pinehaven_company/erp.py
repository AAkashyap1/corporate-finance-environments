"""Application-level repository for Pinehaven's integrated manufacturing ERP.

The agent-facing MCP delegates to this class.  SQL is deliberately kept behind
the repository so callers work with ERP objects, reports, validations, and
workflow transitions rather than a database console.
"""

from __future__ import annotations

import json
import math
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Iterator


SNAPSHOT_DATE = "2026-06-30"
OPEN_PERIOD = "2026-06"


class ERPError(ValueError):
    """A business-rule or validation error suitable for an MCP response."""


def _finite_number(
    value: float | int | None,
    *,
    label: str,
) -> float:
    try:
        numeric = float(value or 0)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ERPError(f"{label} must be a finite number") from exc
    if not math.isfinite(numeric):
        raise ERPError(f"{label} must be a finite number")
    return numeric


def _money(value: float | int | None) -> float:
    return round(_finite_number(value, label="Money") + 1e-10, 2)


def _qty(value: float | int | None) -> float:
    return round(_finite_number(value, label="Quantity") + 1e-10, 4)


def _ratio(
    numerator: float | int | None, denominator: float | int | None
) -> float:
    base = _finite_number(denominator, label="Ratio denominator")
    value = _finite_number(numerator, label="Ratio numerator")
    return round(value / base, 6) if base else 0.0


def _iso_date(value: str, *, label: str) -> str:
    if not isinstance(value, str):
        raise ERPError(f"{label} must use ISO YYYY-MM-DD")
    try:
        parsed = date.fromisoformat(value)
    except ValueError as exc:
        raise ERPError(f"{label} must use ISO YYYY-MM-DD") from exc
    if parsed.isoformat() != value:
        raise ERPError(f"{label} must use ISO YYYY-MM-DD")
    return value


def _as_dict(row: sqlite3.Row | None) -> dict[str, Any] | None:
    return dict(row) if row is not None else None


class PinehavenERP:
    """Transactional ERP facade over one resettable runtime database."""

    def __init__(self, database_path: str | Path, actor: str = "rl-agent") -> None:
        self.database_path = Path(database_path)
        self.actor = actor
        if not self.database_path.exists():
            raise FileNotFoundError(self.database_path)

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.database_path, timeout=30)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys=ON")
        connection.execute("PRAGMA busy_timeout=30000")
        return connection

    @contextmanager
    def _write(self) -> Iterator[sqlite3.Connection]:
        connection = self._connect()
        try:
            connection.execute("BEGIN IMMEDIATE")
            yield connection
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    @staticmethod
    def _limit(value: int, maximum: int = 500) -> int:
        return max(1, min(int(value), maximum))

    @staticmethod
    def _audit(
        connection: sqlite3.Connection,
        actor: str,
        action: str,
        entity: str,
        identifier: str,
        details: dict[str, Any],
    ) -> None:
        connection.execute(
            """
            INSERT INTO audit_events(
                event_at,actor,action,entity,identifier,details_json
            ) VALUES(?,?,?,?,?,?)
            """,
            (
                datetime.now(timezone.utc).isoformat(),
                actor,
                action,
                entity,
                identifier,
                json.dumps(details, sort_keys=True),
            ),
        )

    @staticmethod
    def _require_row(
        connection: sqlite3.Connection,
        query: str,
        params: tuple[Any, ...],
        label: str,
    ) -> sqlite3.Row:
        row = connection.execute(query, params).fetchone()
        if row is None:
            raise ERPError(f"{label} was not found")
        return row

    # ------------------------------------------------------------------
    # Discovery and master data
    # ------------------------------------------------------------------

    def get_company_profile(self) -> dict[str, Any]:
        with self._connect() as connection:
            company = _as_dict(connection.execute("SELECT * FROM company").fetchone())
            sites = [
                dict(row)
                for row in connection.execute(
                    "SELECT * FROM sites ORDER BY code"
                )
            ]
            headcount = int(
                connection.execute(
                    "SELECT COUNT(*) FROM employees WHERE active=1"
                ).fetchone()[0]
            )
            counts = {
                table: int(connection.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0])
                for table in (
                    "items",
                    "customers",
                    "vendors",
                    "production_orders",
                    "inventory_transactions",
                    "sales_orders",
                    "purchase_orders",
                    "journal_headers",
                )
            }
            return {
                "snapshot_date": SNAPSHOT_DATE,
                "closed_through": "2026-05-31",
                "open_period": OPEN_PERIOD,
                "company": company,
                "sites": sites,
                "active_headcount": headcount,
                "object_counts": counts,
            }

    def list_erp_modules(self) -> list[dict[str, Any]]:
        return [
            {
                "module": "Product Information Management",
                "objects": ["items", "site items", "BOMs", "routings", "engineering changes"],
            },
            {
                "module": "Production Control",
                "objects": ["production orders", "material issues", "operations", "variances"],
            },
            {
                "module": "Inventory Management",
                "objects": ["on hand", "movements", "lots", "locations", "transfers"],
            },
            {
                "module": "Procurement and Sourcing",
                "objects": ["vendors", "purchase orders", "receipts", "vendor invoices"],
            },
            {
                "module": "Sales and Distribution",
                "objects": ["customers", "sales orders", "shipments", "customer invoices"],
            },
            {
                "module": "Master Planning",
                "objects": ["forecasts", "planned orders", "action messages"],
            },
            {
                "module": "Quality Management",
                "objects": ["quality orders", "tests", "nonconformances", "holds"],
            },
            {
                "module": "Asset and Maintenance Management",
                "objects": ["maintenance assets", "work orders", "fixed assets"],
            },
            {
                "module": "Financial Management",
                "objects": ["GL", "AR", "AP", "cash", "debt", "payroll"],
            },
        ]

    def list_fiscal_periods(self) -> list[dict[str, Any]]:
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    "SELECT * FROM fiscal_periods ORDER BY period"
                )
            ]

    def get_audit_events(
        self,
        entity: str | None = None,
        identifier: str | None = None,
        limit: int = 100,
    ) -> list[dict[str, Any]]:
        clauses: list[str] = []
        params: list[Any] = []
        if entity:
            clauses.append("entity=?")
            params.append(entity)
        if identifier:
            clauses.append("identifier=?")
            params.append(identifier)
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        params.append(self._limit(limit))
        with self._connect() as connection:
            rows = connection.execute(
                f"""
                SELECT * FROM audit_events {where}
                ORDER BY id DESC LIMIT ?
                """,
                tuple(params),
            )
            output = []
            for row in rows:
                payload = dict(row)
                payload["details"] = json.loads(payload.pop("details_json"))
                output.append(payload)
            return output

    def search_items(
        self,
        query: str | None = None,
        item_type: str | None = None,
        product_family: str | None = None,
        site_code: str | None = None,
        lifecycle_status: str | None = None,
        limit: int = 100,
    ) -> list[dict[str, Any]]:
        clauses = ["1=1"]
        params: list[Any] = []
        if query:
            clauses.append("(i.id LIKE ? OR i.description LIKE ?)")
            pattern = f"%{query}%"
            params.extend([pattern, pattern])
        if item_type:
            clauses.append("i.item_type=?")
            params.append(item_type)
        if product_family:
            clauses.append("i.product_family=?")
            params.append(product_family)
        if site_code:
            clauses.append(
                "EXISTS(SELECT 1 FROM item_sites s WHERE s.item_id=i.id AND s.site_code=?)"
            )
            params.append(site_code)
        if lifecycle_status:
            clauses.append("i.lifecycle_status=?")
            params.append(lifecycle_status)
        params.append(self._limit(limit))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT i.*,
                           (SELECT GROUP_CONCAT(site_code,',')
                            FROM item_sites s WHERE s.item_id=i.id) AS sites
                    FROM items i
                    WHERE {' AND '.join(clauses)}
                    ORDER BY i.id LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def get_item(self, item_id: str) -> dict[str, Any]:
        with self._connect() as connection:
            item = self._require_row(
                connection, "SELECT * FROM items WHERE id=?", (item_id,), "Item"
            )
            sites = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT s.*,
                           ROUND(s.standard_material_cost+s.standard_labor_cost+
                                 s.standard_variable_overhead+s.standard_fixed_overhead+
                                 s.standard_outside_processing,2) AS total_standard_cost
                    FROM item_sites s WHERE item_id=? ORDER BY site_code
                    """,
                    (item_id,),
                )
            ]
            return {"item": dict(item), "site_records": sites}

    def get_bom(
        self,
        item_id: str,
        site_code: str,
        as_of: str = SNAPSHOT_DATE,
        explode_levels: int = 1,
    ) -> dict[str, Any]:
        levels = max(1, min(int(explode_levels), 4))
        with self._connect() as connection:
            header = self._require_row(
                connection,
                """
                SELECT * FROM bom_headers
                WHERE item_id=? AND site_code=? AND status='Active'
                  AND effective_from<=?
                  AND (effective_to IS NULL OR effective_to>=?)
                ORDER BY effective_from DESC LIMIT 1
                """,
                (item_id, site_code, as_of, as_of),
                "Effective BOM",
            )

            def explode(bom_id: int, level: int) -> list[dict[str, Any]]:
                rows = [
                    dict(row)
                    for row in connection.execute(
                        """
                        SELECT l.*,i.description,i.item_type,i.make_buy
                        FROM bom_lines l JOIN items i ON i.id=l.component_item_id
                        WHERE l.bom_id=? ORDER BY l.line_number
                        """,
                        (bom_id,),
                    )
                ]
                if level < levels:
                    for row in rows:
                        nested = connection.execute(
                            """
                            SELECT id FROM bom_headers
                            WHERE item_id=? AND site_code=? AND status='Active'
                              AND effective_from<=?
                              AND (effective_to IS NULL OR effective_to>=?)
                            ORDER BY effective_from DESC LIMIT 1
                            """,
                            (row["component_item_id"], site_code, as_of, as_of),
                        ).fetchone()
                        if nested:
                            row["children"] = explode(int(nested[0]), level + 1)
                return rows

            return {"header": dict(header), "lines": explode(int(header["id"]), 1)}

    def get_routing(
        self, item_id: str, site_code: str, as_of: str = SNAPSHOT_DATE
    ) -> dict[str, Any]:
        with self._connect() as connection:
            header = self._require_row(
                connection,
                """
                SELECT * FROM routing_headers
                WHERE item_id=? AND site_code=? AND status='Active'
                  AND effective_from<=?
                  AND (effective_to IS NULL OR effective_to>=?)
                ORDER BY effective_from DESC LIMIT 1
                """,
                (item_id, site_code, as_of, as_of),
                "Effective routing",
            )
            operations = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT o.*,w.name AS work_center_name,w.cost_center,
                           w.labor_rate,w.machine_rate,
                           w.variable_overhead_rate,w.fixed_overhead_rate
                    FROM routing_operations o
                    JOIN work_centers w ON w.code=o.work_center_code
                    WHERE o.routing_id=? ORDER BY o.operation_number
                    """,
                    (header["id"],),
                )
            ]
            return {"header": dict(header), "operations": operations}

    def get_standard_cost(self, item_id: str, site_code: str) -> dict[str, Any]:
        with self._connect() as connection:
            row = self._require_row(
                connection,
                """
                SELECT s.*,i.description,i.item_type,i.product_family,
                       ROUND(s.standard_material_cost+s.standard_labor_cost+
                             s.standard_variable_overhead+s.standard_fixed_overhead+
                             s.standard_outside_processing,2) AS total_standard_cost
                FROM item_sites s JOIN items i ON i.id=s.item_id
                WHERE s.item_id=? AND s.site_code=?
                """,
                (item_id, site_code),
                "Item/site standard cost",
            )
            return dict(row)

    def list_engineering_changes(
        self,
        status: str | None = None,
        effective_from: str | None = None,
        limit: int = 100,
    ) -> list[dict[str, Any]]:
        clauses = ["1=1"]
        params: list[Any] = []
        if status:
            clauses.append("status=?")
            params.append(status)
        if effective_from:
            clauses.append("effective_date>=?")
            params.append(effective_from)
        params.append(self._limit(limit))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT e.*,i.description FROM engineering_changes e
                    JOIN items i ON i.id=e.item_id
                    WHERE {' AND '.join(clauses)}
                    ORDER BY effective_date,e.id LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    # ------------------------------------------------------------------
    # Manufacturing, capacity, inventory, quality, and maintenance
    # ------------------------------------------------------------------

    def search_production_orders(
        self,
        status: str | None = None,
        site_code: str | None = None,
        item_id: str | None = None,
        finish_from: str | None = None,
        finish_to: str | None = None,
        limit: int = 100,
    ) -> list[dict[str, Any]]:
        clauses = ["1=1"]
        params: list[Any] = []
        for column, value in (
            ("p.status", status),
            ("p.site_code", site_code),
            ("p.item_id", item_id),
        ):
            if value:
                clauses.append(f"{column}=?")
                params.append(value)
        if finish_from:
            clauses.append("COALESCE(p.actual_finish,p.scheduled_finish)>=?")
            params.append(finish_from)
        if finish_to:
            clauses.append("COALESCE(p.actual_finish,p.scheduled_finish)<=?")
            params.append(finish_to)
        params.append(self._limit(limit))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT p.id,p.order_number,p.item_id,i.description,
                           i.product_family,p.site_code,p.source_type,
                           p.source_reference,p.scheduled_start,p.scheduled_finish,
                           p.actual_start,p.actual_finish,p.order_quantity,
                           p.completed_quantity,p.scrapped_quantity,p.status,
                           p.standard_unit_cost,
                           ROUND(p.completed_quantity*p.standard_unit_cost,2)
                               AS standard_output_value
                    FROM production_orders p JOIN items i ON i.id=p.item_id
                    WHERE {' AND '.join(clauses)}
                    ORDER BY COALESCE(p.actual_finish,p.scheduled_finish) DESC,p.id
                    LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def get_production_order(self, order_number: str) -> dict[str, Any]:
        with self._connect() as connection:
            order = self._require_row(
                connection,
                """
                SELECT p.*,i.description,i.product_family
                FROM production_orders p JOIN items i ON i.id=p.item_id
                WHERE p.order_number=?
                """,
                (order_number,),
                "Production order",
            )
            materials = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT m.*,i.description,i.item_type,
                           ROUND(m.planned_quantity*m.standard_unit_cost,2)
                               AS planned_standard_value,
                           ROUND(m.issued_quantity*m.actual_unit_cost,2)
                               AS actual_value
                    FROM production_order_materials m
                    JOIN items i ON i.id=m.component_item_id
                    WHERE m.production_order_id=?
                    ORDER BY m.line_number
                    """,
                    (order["id"],),
                )
            ]
            operations = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT o.*,w.name AS work_center_name,w.site_code,
                           ROUND((o.actual_setup_hours+o.actual_run_hours)*
                                 o.labor_rate_actual,2) AS actual_labor_cost
                    FROM production_order_operations o
                    JOIN work_centers w ON w.code=o.work_center_code
                    WHERE o.production_order_id=?
                    ORDER BY o.operation_number
                    """,
                    (order["id"],),
                )
            ]
            variance = _as_dict(
                connection.execute(
                    """
                    SELECT * FROM production_variances
                    WHERE production_order_id=?
                    """,
                    (order["id"],),
                ).fetchone()
            )
            return {
                "order": dict(order),
                "materials": materials,
                "operations": operations,
                "variance": variance,
            }

    def get_production_cost_analysis(self, order_number: str) -> dict[str, Any]:
        detail = self.get_production_order(order_number)
        order = detail["order"]
        standard_output = _money(
            order["completed_quantity"] * order["standard_unit_cost"]
        )
        actual = {
            "material": _money(order["actual_material_cost"]),
            "labor": _money(order["actual_labor_cost"]),
            "variable_overhead": _money(order["actual_variable_overhead"]),
            "fixed_overhead": _money(order["actual_fixed_overhead"]),
            "outside_processing": _money(order["actual_outside_processing"]),
        }
        actual["total"] = _money(sum(actual.values()))
        return {
            "order_number": order_number,
            "status": order["status"],
            "item_id": order["item_id"],
            "site_code": order["site_code"],
            "completed_quantity": order["completed_quantity"],
            "standard_output_cost": standard_output,
            "actual_cost": actual,
            "actual_less_standard": _money(actual["total"] - standard_output),
            "variance_components": detail["variance"],
        }

    def get_production_variance_report(
        self,
        date_from: str,
        date_to: str,
        site_code: str | None = None,
        product_family: str | None = None,
        settled_only: bool = True,
    ) -> dict[str, Any]:
        clauses = [
            "COALESCE(v.settled_date,p.actual_finish,p.scheduled_finish) BETWEEN ? AND ?"
        ]
        params: list[Any] = [date_from, date_to]
        if settled_only:
            clauses.append("v.settled_date IS NOT NULL")
        if site_code:
            clauses.append("p.site_code=?")
            params.append(site_code)
        if product_family:
            clauses.append("i.product_family=?")
            params.append(product_family)
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT p.site_code,COALESCE(i.product_family,'Subassembly') AS product_family,
                           COUNT(*) AS order_count,
                           ROUND(SUM(p.completed_quantity*p.standard_unit_cost),2)
                               AS standard_output_cost,
                           ROUND(SUM(v.material_price_variance),2)
                               AS material_price_variance,
                           ROUND(SUM(v.material_usage_variance),2)
                               AS material_usage_variance,
                           ROUND(SUM(v.labor_rate_variance),2)
                               AS labor_rate_variance,
                           ROUND(SUM(v.labor_efficiency_variance),2)
                               AS labor_efficiency_variance,
                           ROUND(SUM(v.variable_overhead_variance),2)
                               AS variable_overhead_variance,
                           ROUND(SUM(v.fixed_overhead_volume_variance),2)
                               AS fixed_overhead_volume_variance,
                           ROUND(SUM(v.scrap_variance),2) AS scrap_variance,
                           ROUND(SUM(v.substitution_variance),2)
                               AS substitution_variance,
                           ROUND(SUM(v.total_variance),2) AS total_variance
                    FROM production_variances v
                    JOIN production_orders p ON p.id=v.production_order_id
                    JOIN items i ON i.id=p.item_id
                    WHERE {' AND '.join(clauses)}
                    GROUP BY p.site_code,COALESCE(i.product_family,'Subassembly')
                    ORDER BY ABS(SUM(v.total_variance)) DESC
                    """,
                    tuple(params),
                )
            ]
            totals = dict(
                connection.execute(
                    f"""
                    SELECT COUNT(*) AS order_count,
                           SUM(p.completed_quantity*p.standard_unit_cost)
                               AS standard_output_cost,
                           SUM(v.material_price_variance)
                               AS material_price_variance,
                           SUM(v.material_usage_variance)
                               AS material_usage_variance,
                           SUM(v.labor_rate_variance)
                               AS labor_rate_variance,
                           SUM(v.labor_efficiency_variance)
                               AS labor_efficiency_variance,
                           SUM(v.variable_overhead_variance)
                               AS variable_overhead_variance,
                           SUM(v.fixed_overhead_volume_variance)
                               AS fixed_overhead_volume_variance,
                           SUM(v.scrap_variance) AS scrap_variance,
                           SUM(v.substitution_variance)
                               AS substitution_variance,
                           SUM(v.total_variance) AS total_variance
                    FROM production_variances v
                    JOIN production_orders p ON p.id=v.production_order_id
                    JOIN items i ON i.id=p.item_id
                    WHERE {' AND '.join(clauses)}
                    """,
                    tuple(params),
                ).fetchone()
            )
            standard_output = _money(totals["standard_output_cost"])
            total_variance = _money(totals["total_variance"])
            return {
                "date_from": date_from,
                "date_to": date_to,
                "settled_only": settled_only,
                "rows": rows,
                "order_count": int(totals["order_count"] or 0),
                "standard_output_cost": standard_output,
                "material_price_variance": _money(
                    totals["material_price_variance"]
                ),
                "material_usage_variance": _money(
                    totals["material_usage_variance"]
                ),
                "labor_rate_variance": _money(totals["labor_rate_variance"]),
                "labor_efficiency_variance": _money(
                    totals["labor_efficiency_variance"]
                ),
                "variable_overhead_variance": _money(
                    totals["variable_overhead_variance"]
                ),
                "fixed_overhead_volume_variance": _money(
                    totals["fixed_overhead_volume_variance"]
                ),
                "scrap_variance": _money(totals["scrap_variance"]),
                "substitution_variance": _money(
                    totals["substitution_variance"]
                ),
                "total_variance": total_variance,
                "variance_rate": _ratio(total_variance, standard_output),
            }

    def get_wip_report(
        self, as_of: str = SNAPSHOT_DATE, site_code: str | None = None
    ) -> dict[str, Any]:
        clauses = ["p.status<>'Ended'", "p.created_date<=?"]
        params: list[Any] = [as_of]
        if site_code:
            clauses.append("p.site_code=?")
            params.append(site_code)
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT p.order_number,p.item_id,i.description,p.site_code,p.status,
                           p.scheduled_finish,p.order_quantity,p.completed_quantity,
                           ROUND(p.actual_material_cost+p.actual_labor_cost+
                                 p.actual_variable_overhead+p.actual_fixed_overhead+
                                 p.actual_outside_processing,2) AS cost_to_date,
                           ROUND(p.completed_quantity*p.standard_unit_cost,2)
                               AS reported_output,
                           ROUND(p.actual_material_cost+p.actual_labor_cost+
                                 p.actual_variable_overhead+p.actual_fixed_overhead+
                                 p.actual_outside_processing-
                                 p.completed_quantity*p.standard_unit_cost,2)
                               AS wip_value
                    FROM production_orders p JOIN items i ON i.id=p.item_id
                    WHERE {' AND '.join(clauses)}
                    ORDER BY wip_value DESC,p.order_number
                    """,
                    tuple(params),
                )
            ]
            totals = dict(
                connection.execute(
                    f"""
                    WITH scoped AS (
                        SELECT p.order_number,p.status,
                               p.actual_material_cost+p.actual_labor_cost+
                               p.actual_variable_overhead+
                               p.actual_fixed_overhead+
                               p.actual_outside_processing AS cost_to_date,
                               p.completed_quantity*p.standard_unit_cost
                                   AS reported_output,
                               p.actual_material_cost+p.actual_labor_cost+
                               p.actual_variable_overhead+
                               p.actual_fixed_overhead+
                               p.actual_outside_processing-
                               p.completed_quantity*p.standard_unit_cost
                                   AS wip_value
                        FROM production_orders p
                        WHERE {' AND '.join(clauses)}
                    )
                    SELECT COUNT(*) AS order_count,
                           SUM(CASE WHEN status='Started' THEN 1 ELSE 0 END)
                               AS started_orders,
                           SUM(CASE WHEN status='Reported Finished'
                               THEN 1 ELSE 0 END)
                               AS reported_finished_orders,
                           SUM(cost_to_date) AS cost_to_date,
                           SUM(reported_output) AS reported_output,
                           SUM(wip_value) AS total_wip,
                           SUM(CASE WHEN wip_value<0 THEN 1 ELSE 0 END)
                               AS negative_residual_count,
                           SUM(CASE WHEN wip_value<0 THEN wip_value ELSE 0 END)
                               AS negative_residual_value
                    FROM scoped
                    """,
                    tuple(params),
                ).fetchone()
            )
            top_rows = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT p.order_number,
                           p.actual_material_cost+p.actual_labor_cost+
                           p.actual_variable_overhead+
                           p.actual_fixed_overhead+
                           p.actual_outside_processing-
                           p.completed_quantity*p.standard_unit_cost
                               AS wip_value
                    FROM production_orders p
                    WHERE {' AND '.join(clauses)}
                    ORDER BY wip_value DESC,p.order_number
                    LIMIT 5
                    """,
                    tuple(params),
                )
            ]
            total_wip = _money(totals["total_wip"])
            return {
                "as_of": as_of,
                "rows": rows,
                "order_count": int(totals["order_count"] or 0),
                "started_orders": int(totals["started_orders"] or 0),
                "reported_finished_orders": int(
                    totals["reported_finished_orders"] or 0
                ),
                "cost_to_date": _money(totals["cost_to_date"]),
                "reported_output": _money(totals["reported_output"]),
                "wip_value": total_wip,
                "total_wip": total_wip,
                "largest_order": (
                    top_rows[0]["order_number"] if top_rows else None
                ),
                "largest_order_wip": (
                    _money(top_rows[0]["wip_value"]) if top_rows else 0.0
                ),
                "top_five_concentration": _ratio(
                    sum(float(row["wip_value"]) for row in top_rows),
                    totals["total_wip"],
                ),
                "negative_residual_count": int(
                    totals["negative_residual_count"] or 0
                ),
                "negative_residual_value": _money(
                    totals["negative_residual_value"]
                ),
            }

    def get_scrap_yield_report(
        self,
        date_from: str,
        date_to: str,
        site_code: str | None = None,
    ) -> list[dict[str, Any]]:
        clauses = ["COALESCE(p.actual_finish,p.scheduled_finish) BETWEEN ? AND ?"]
        params: list[Any] = [date_from, date_to]
        if site_code:
            clauses.append("p.site_code=?")
            params.append(site_code)
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT p.site_code,COALESCE(i.product_family,'Subassembly') product_family,
                           COUNT(*) order_count,
                           ROUND(SUM(p.order_quantity),4) started_quantity,
                           ROUND(SUM(p.completed_quantity),4) completed_quantity,
                           ROUND(SUM(p.scrapped_quantity),4) scrapped_quantity,
                           ROUND(CASE WHEN SUM(p.completed_quantity+p.scrapped_quantity)=0
                                THEN 0 ELSE
                                100.0*SUM(p.completed_quantity)/
                                SUM(p.completed_quantity+p.scrapped_quantity) END,2)
                                AS first_pass_yield_percent,
                           ROUND(SUM(v.scrap_variance),2) scrap_cost
                    FROM production_orders p
                    JOIN items i ON i.id=p.item_id
                    JOIN production_variances v ON v.production_order_id=p.id
                    WHERE {' AND '.join(clauses)}
                    GROUP BY p.site_code,COALESCE(i.product_family,'Subassembly')
                    ORDER BY scrap_cost DESC
                    """,
                    tuple(params),
                )
            ]

    def get_work_center_capacity(
        self,
        date_from: str,
        date_to: str,
        site_code: str | None = None,
    ) -> list[dict[str, Any]]:
        site_filter = ""
        if site_code:
            site_filter = "WHERE w.site_code=?"
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    f"""
                    WITH capacity AS (
                        SELECT work_center_code,
                               SUM(available_hours-planned_downtime_hours)
                                   AS net_available_hours,
                               SUM(planned_downtime_hours) AS planned_downtime_hours
                        FROM work_center_calendar
                        WHERE work_date BETWEEN ? AND ?
                        GROUP BY work_center_code
                    ),
                    released_load AS (
                        SELECT o.work_center_code,
                               SUM(o.planned_setup_hours+o.planned_run_hours)
                                   AS planned_load_hours,
                               SUM(o.actual_setup_hours+o.actual_run_hours)
                                   AS actual_hours
                        FROM production_order_operations o
                        JOIN production_orders p ON p.id=o.production_order_id
                        WHERE COALESCE(o.completed_date,p.scheduled_finish)
                              BETWEEN ? AND ?
                        GROUP BY o.work_center_code
                    ),
                    planned_load AS (
                        SELECT ro.work_center_code,
                               SUM(ro.setup_hours+
                                   ro.run_hours_per_unit*p.quantity)
                                   AS planned_load_hours,
                               0.0 AS actual_hours
                        FROM planned_orders p
                        JOIN routing_headers rh
                          ON rh.item_id=p.item_id
                         AND rh.site_code=p.site_code
                         AND rh.status='Active'
                         AND rh.effective_from<=p.required_date
                         AND (rh.effective_to IS NULL
                              OR rh.effective_to>=p.required_date)
                        JOIN routing_operations ro ON ro.routing_id=rh.id
                        WHERE p.plan_version='2026-06-SOP'
                          AND p.order_type='Production'
                          AND p.status='Planned'
                          AND p.required_date BETWEEN ? AND ?
                        GROUP BY ro.work_center_code
                    ),
                    load AS (
                        SELECT work_center_code,
                               SUM(planned_load_hours) AS planned_load_hours,
                               SUM(actual_hours) AS actual_hours
                        FROM (
                            SELECT * FROM released_load
                            UNION ALL
                            SELECT * FROM planned_load
                        )
                        GROUP BY work_center_code
                    )
                    SELECT w.code,w.name,w.site_code,w.cost_center,
                           COALESCE(c.net_available_hours,0)
                               AS raw_net_available_hours,
                           COALESCE(c.planned_downtime_hours,0)
                               AS raw_planned_downtime_hours,
                           COALESCE(l.planned_load_hours,0)
                               AS raw_planned_load_hours,
                           COALESCE(l.actual_hours,0) AS raw_actual_hours,
                           ROUND(CASE WHEN COALESCE(c.net_available_hours,0)=0 THEN 0
                                ELSE 100.0*COALESCE(l.planned_load_hours,0)/
                                c.net_available_hours END,2) AS planned_utilization_percent
                    FROM work_centers w
                    LEFT JOIN capacity c ON c.work_center_code=w.code
                    LEFT JOIN load l ON l.work_center_code=w.code
                    {site_filter}
                    ORDER BY planned_utilization_percent DESC,w.code
                    """,
                    tuple([date_from, date_to, date_from, date_to,
                           date_from, date_to, *([site_code] if site_code else [])]),
                )
            ]
            raw_rows: list[dict[str, Any]] = []
            for row in rows:
                raw = {
                    "net_available_hours": float(
                        row.pop("raw_net_available_hours")
                    ),
                    "planned_downtime_hours": float(
                        row.pop("raw_planned_downtime_hours")
                    ),
                    "planned_load_hours": float(
                        row.pop("raw_planned_load_hours")
                    ),
                    "actual_hours": float(row.pop("raw_actual_hours")),
                }
                raw_rows.append(raw)
                for key, value in raw.items():
                    row[key] = _money(value)
            if not rows:
                return rows
            raw_utilizations = [
                (
                    raw["planned_load_hours"]
                    / raw["net_available_hours"]
                    if raw["net_available_hours"]
                    else 0.0
                )
                for raw in raw_rows
            ]
            top_index = max(
                range(len(rows)),
                key=lambda index: (
                    raw_utilizations[index],
                    rows[index]["code"],
                ),
            )
            required_hours = sum(
                raw["planned_load_hours"] for raw in raw_rows
            )
            net_available_hours = sum(
                raw["net_available_hours"] for raw in raw_rows
            )
            scope = {
                "scope_work_centers": len(rows),
                "scope_net_available_hours": _money(net_available_hours),
                "scope_planned_downtime_hours": _money(
                    sum(
                        raw["planned_downtime_hours"]
                        for raw in raw_rows
                    )
                ),
                "scope_required_hours": _money(required_hours),
                "scope_utilization": _ratio(
                    required_hours,
                    net_available_hours,
                ),
                "scope_overloaded_centers": sum(
                    utilization > 1
                    for utilization in raw_utilizations
                ),
                "scope_top_constraint": rows[top_index]["code"],
                "scope_top_constraint_utilization": round(
                    raw_utilizations[top_index],
                    6,
                ),
            }
            for row in rows:
                row.update(scope)
            return rows

    def get_inventory_on_hand(
        self,
        item_id: str | None = None,
        site_code: str | None = None,
        warehouse_code: str | None = None,
        include_zero: bool = False,
        limit: int = 500,
    ) -> list[dict[str, Any]]:
        clauses = ["1=1"]
        params: list[Any] = []
        if not include_zero:
            clauses.append("b.on_hand_quantity<>0")
        for column, value in (
            ("b.item_id", item_id),
            ("b.site_code", site_code),
            ("b.warehouse_code", warehouse_code),
        ):
            if value:
                clauses.append(f"{column}=?")
                params.append(value)
        params.append(self._limit(limit, 2000))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT b.*,i.description,i.item_type,i.product_family,
                           ROUND(b.on_hand_quantity-b.reserved_quantity-
                                 b.quality_hold_quantity,4) AS available_quantity,
                           ROUND(b.on_hand_quantity*b.standard_unit_cost,2)
                               AS inventory_value
                    FROM inventory_balances b JOIN items i ON i.id=b.item_id
                    WHERE {' AND '.join(clauses)}
                    ORDER BY inventory_value DESC,b.item_id,b.site_code
                    LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def get_inventory_valuation(
        self,
        site_code: str | None = None,
        group_by: str = "item_type",
    ) -> dict[str, Any]:
        dimensions = {
            "item_type": "i.item_type",
            "product_family": "COALESCE(i.product_family,i.item_type)",
            "site": "b.site_code",
            "warehouse": "b.warehouse_code",
            "item": "b.item_id",
        }
        if group_by not in dimensions:
            raise ERPError(f"Unsupported valuation grouping: {group_by}")
        clauses = ["1=1"]
        params: list[Any] = []
        if site_code:
            clauses.append("b.site_code=?")
            params.append(site_code)
        dimension = dimensions[group_by]
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT {dimension} AS dimension,
                           ROUND(SUM(b.on_hand_quantity),4) AS on_hand_quantity,
                           ROUND(SUM(b.reserved_quantity),4) AS reserved_quantity,
                           ROUND(SUM(b.quality_hold_quantity),4) AS hold_quantity,
                           ROUND(SUM(b.on_hand_quantity*b.standard_unit_cost),2)
                               AS inventory_value
                    FROM inventory_balances b JOIN items i ON i.id=b.item_id
                    WHERE {' AND '.join(clauses)}
                    GROUP BY {dimension}
                    ORDER BY inventory_value DESC
                    """,
                    tuple(params),
                )
            ]
            return {
                "group_by": group_by,
                "site_code": site_code,
                "rows": rows,
                "total_value": _money(sum(row["inventory_value"] for row in rows)),
            }

    def get_inventory_movements(
        self,
        date_from: str,
        date_to: str,
        item_id: str | None = None,
        site_code: str | None = None,
        transaction_type: str | None = None,
        limit: int = 500,
        product_family: str | None = None,
    ) -> list[dict[str, Any]]:
        clauses = ["t.transaction_date BETWEEN ? AND ?"]
        params: list[Any] = [date_from, date_to]
        for column, value in (
            ("t.item_id", item_id),
            ("t.site_code", site_code),
            ("t.transaction_type", transaction_type),
        ):
            if value:
                clauses.append(f"{column}=?")
                params.append(value)
        if product_family:
            clauses.append("i.product_family=?")
            params.append(product_family)
        params.append(self._limit(limit, 2000))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT t.*,i.description,i.product_family,
                           t.quantity*t.unit_cost AS extended_value,
                           ROUND(t.quantity*t.unit_cost,2)
                               AS display_extended_value
                    FROM inventory_transactions t JOIN items i ON i.id=t.item_id
                    WHERE {' AND '.join(clauses)}
                    ORDER BY t.transaction_date DESC,t.id DESC LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def get_inventory_aging(
        self,
        as_of: str = SNAPSHOT_DATE,
        site_code: str | None = None,
        lifecycle_status: str | None = None,
        threshold_days: int = 180,
    ) -> dict[str, Any]:
        if threshold_days < 0 or threshold_days > 3650:
            raise ValueError("threshold_days must be between 0 and 3650")
        clauses = ["b.on_hand_quantity>0"]
        params: list[Any] = [as_of]
        if site_code:
            clauses.append("b.site_code=?")
            params.append(site_code)
        if lifecycle_status:
            clauses.append("i.lifecycle_status=?")
            params.append(lifecycle_status)
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT b.item_id,b.site_code,b.on_hand_quantity,
                           b.quality_hold_quantity,b.standard_unit_cost,
                           i.lifecycle_status,
                           CAST(julianday(?)-
                                julianday(b.last_movement_date) AS INTEGER)
                               AS age_days
                    FROM inventory_balances b
                    JOIN items i ON i.id=b.item_id
                    WHERE {' AND '.join(clauses)}
                    """,
                    tuple(params),
                )
            ]
            bucket_order = (
                "0-90 days",
                "91-180 days",
                "181-365 days",
                "Over 365 days",
            )

            def bucket(age_days: int) -> str:
                if age_days <= 90:
                    return bucket_order[0]
                if age_days <= 180:
                    return bucket_order[1]
                if age_days <= 365:
                    return bucket_order[2]
                return bucket_order[3]

            grouped: dict[str, dict[str, float | int]] = {
                name: {"item_site_count": 0, "inventory_value": 0.0}
                for name in bucket_order
            }
            aged_by_item: dict[str, float] = {}
            aged_by_item_site: dict[tuple[str, str], float] = {}
            total_value = 0.0
            phase_out_value = 0.0
            quality_hold_value = 0.0
            value_over_threshold = 0.0
            for row in rows:
                inventory_value = (
                    float(row["on_hand_quantity"])
                    * float(row["standard_unit_cost"])
                )
                name = bucket(int(row["age_days"]))
                grouped[name]["item_site_count"] = (
                    int(grouped[name]["item_site_count"]) + 1
                )
                grouped[name]["inventory_value"] = (
                    float(grouped[name]["inventory_value"]) + inventory_value
                )
                total_value += inventory_value
                quality_hold_value += (
                    float(row["quality_hold_quantity"])
                    * float(row["standard_unit_cost"])
                )
                if row["lifecycle_status"] == "Phase Out":
                    phase_out_value += inventory_value
                if int(row["age_days"]) > threshold_days:
                    value_over_threshold += inventory_value
                    aged_by_item[row["item_id"]] = (
                        aged_by_item.get(row["item_id"], 0.0)
                        + inventory_value
                    )
                    item_site = (
                        str(row["item_id"]),
                        str(row["site_code"]),
                    )
                    aged_by_item_site[item_site] = (
                        aged_by_item_site.get(item_site, 0.0)
                        + inventory_value
                    )
            top_item, top_item_value = (
                max(
                    aged_by_item.items(),
                    key=lambda pair: (pair[1], pair[0]),
                )
                if aged_by_item
                else ("None", 0.0)
            )
            (top_site_item, top_site), top_item_site_value = (
                max(
                    aged_by_item_site.items(),
                    key=lambda pair: (
                        pair[1],
                        pair[0][0],
                        pair[0][1],
                    ),
                )
                if aged_by_item_site
                else (("None", "None"), 0.0)
            )
            summary_rows = [
                {
                    "aging_bucket": name,
                    "item_site_count": int(grouped[name]["item_site_count"]),
                    "inventory_value": _money(
                        grouped[name]["inventory_value"]
                    ),
                }
                for name in bucket_order
                if grouped[name]["item_site_count"]
            ]
            return {
                "as_of": as_of,
                "site_code": site_code,
                "lifecycle_status": lifecycle_status,
                "threshold_days": threshold_days,
                "rows": summary_rows,
                "item_site_rows": len(rows),
                "inventory_value": _money(total_value),
                "total_value": _money(total_value),
                "value_over_threshold": _money(value_over_threshold),
                "phase_out_value": _money(phase_out_value),
                "quality_hold_value": _money(quality_hold_value),
                "over_threshold_percent": _ratio(
                    value_over_threshold, total_value
                ),
                "top_item": top_item,
                "top_item_value": _money(top_item_value),
                "top_item_site": (
                    f"{top_site_item} / site {top_site}"
                ),
                "top_item_site_value": _money(top_item_site_value),
            }

    def get_quality_orders(
        self,
        status: str | None = None,
        site_code: str | None = None,
        opened_from: str | None = None,
        limit: int = 200,
    ) -> list[dict[str, Any]]:
        clauses = ["1=1"]
        params: list[Any] = []
        for column, value in (("q.status", status), ("q.site_code", site_code)):
            if value:
                clauses.append(f"{column}=?")
                params.append(value)
        if opened_from:
            clauses.append("q.opened_date>=?")
            params.append(opened_from)
        params.append(self._limit(limit, 1000))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT q.*,i.description,
                           ROUND(CASE WHEN q.quantity_inspected=0 THEN 0
                                ELSE 100.0*q.quantity_failed/q.quantity_inspected END,2)
                                AS failure_rate_percent
                    FROM quality_orders q JOIN items i ON i.id=q.item_id
                    WHERE {' AND '.join(clauses)}
                    ORDER BY q.opened_date DESC,q.id LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def get_nonconformances(
        self,
        status: str | None = None,
        site_code: str | None = None,
        limit: int = 200,
    ) -> list[dict[str, Any]]:
        clauses = ["1=1"]
        params: list[Any] = []
        for column, value in (("n.status", status), ("n.site_code", site_code)):
            if value:
                clauses.append(f"{column}=?")
                params.append(value)
        params.append(self._limit(limit, 1000))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT n.*,i.description,e.full_name AS owner_name
                    FROM nonconformances n
                    LEFT JOIN items i ON i.id=n.item_id
                    JOIN employees e ON e.id=n.owner_employee_id
                    WHERE {' AND '.join(clauses)}
                    ORDER BY n.estimated_cost DESC,n.id LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def get_maintenance_report(
        self,
        date_from: str,
        date_to: str,
        site_code: str | None = None,
        maintenance_type: str | None = None,
    ) -> list[dict[str, Any]]:
        clauses = ["w.request_date BETWEEN ? AND ?"]
        params: list[Any] = [date_from, date_to]
        if site_code:
            clauses.append("a.site_code=?")
            params.append(site_code)
        if maintenance_type:
            clauses.append("w.work_type=?")
            params.append(maintenance_type)
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT a.site_code,w.work_type AS maintenance_type,w.status,
                           COUNT(*) AS work_order_count,
                           ROUND(SUM(w.downtime_hours),2) AS downtime_hours,
                           ROUND(SUM(w.labor_cost),2) AS labor_cost,
                           ROUND(SUM(w.material_cost),2) AS parts_cost,
                           ROUND(SUM(w.outside_service_cost),2)
                               AS outside_service_cost,
                           ROUND(SUM(w.labor_cost+w.material_cost+
                                     w.outside_service_cost),2) AS total_cost
                    FROM maintenance_work_orders w
                    JOIN maintenance_assets a ON a.id=w.asset_id
                    WHERE {' AND '.join(clauses)}
                    GROUP BY a.site_code,w.work_type,w.status
                    ORDER BY total_cost DESC
                    """,
                    tuple(params),
                )
            ]

    def get_maintenance_asset_report(
        self,
        date_from: str,
        date_to: str,
        site_code: str | None = None,
        limit: int = 200,
    ) -> list[dict[str, Any]]:
        """Return the asset-level detail needed to audit downtime rankings."""

        clauses = ["w.request_date BETWEEN ? AND ?"]
        params: list[Any] = [date_from, date_to]
        if site_code:
            clauses.append("a.site_code=?")
            params.append(site_code)
        params.append(self._limit(limit, 1000))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT a.id AS asset_id,a.description AS asset_description,
                           a.site_code,a.work_center_code,a.criticality,
                           COUNT(*) AS work_order_count,
                           ROUND(SUM(w.downtime_hours),2) AS downtime_hours,
                           COUNT(CASE WHEN w.work_type='Corrective' THEN 1 END)
                               AS unplanned_work_orders,
                           ROUND(SUM(CASE WHEN w.work_type='Corrective'
                                          THEN w.downtime_hours ELSE 0 END),2)
                               AS unplanned_downtime,
                           ROUND(SUM(w.labor_cost+w.material_cost+
                                     w.outside_service_cost),2) AS total_cost
                    FROM maintenance_work_orders w
                    JOIN maintenance_assets a ON a.id=w.asset_id
                    WHERE {' AND '.join(clauses)}
                    GROUP BY a.id,a.description,a.site_code,a.work_center_code,
                             a.criticality
                    ORDER BY downtime_hours DESC,a.id
                    LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    # ------------------------------------------------------------------
    # Procurement, order-to-cash, and planning
    # ------------------------------------------------------------------

    def search_vendors(
        self,
        query: str | None = None,
        commodity: str | None = None,
        active_only: bool = True,
        limit: int = 100,
    ) -> list[dict[str, Any]]:
        clauses = ["1=1"]
        params: list[Any] = []
        if query:
            clauses.append("(v.id LIKE ? OR v.name LIKE ?)")
            pattern = f"%{query}%"
            params.extend([pattern, pattern])
        if commodity:
            clauses.append("v.category=?")
            params.append(commodity)
        if active_only:
            clauses.append("v.active=1")
        params.append(self._limit(limit))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT v.*,
                           (SELECT COUNT(*) FROM purchase_orders p
                            WHERE p.vendor_id=v.id) AS purchase_order_count,
                           (SELECT ROUND(SUM(open_amount),2) FROM vendor_invoices vi
                            WHERE vi.vendor_id=v.id) AS open_ap
                    FROM vendors v WHERE {' AND '.join(clauses)}
                    ORDER BY v.name LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def search_purchase_orders(
        self,
        status: str | None = None,
        vendor_id: str | None = None,
        site_code: str | None = None,
        expected_from: str | None = None,
        expected_to: str | None = None,
        limit: int = 100,
    ) -> list[dict[str, Any]]:
        clauses = ["1=1"]
        params: list[Any] = []
        for column, value in (
            ("p.status", status),
            ("p.vendor_id", vendor_id),
            ("p.site_code", site_code),
        ):
            if value:
                clauses.append(f"{column}=?")
                params.append(value)
        if expected_from:
            clauses.append("p.expected_date>=?")
            params.append(expected_from)
        if expected_to:
            clauses.append("p.expected_date<=?")
            params.append(expected_to)
        params.append(self._limit(limit, 1000))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT p.*,v.name AS vendor_name,e.full_name AS buyer_name,
                           ROUND(SUM((l.order_quantity-l.received_quantity)*
                                     l.unit_price),2) AS open_commitment
                    FROM purchase_orders p
                    JOIN vendors v ON v.id=p.vendor_id
                    JOIN employees e ON e.id=p.buyer_employee_id
                    JOIN purchase_order_lines l ON l.po_id=p.id
                    WHERE {' AND '.join(clauses)}
                    GROUP BY p.id ORDER BY p.expected_date,p.po_number LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def get_purchase_order(self, po_number: str) -> dict[str, Any]:
        with self._connect() as connection:
            header = self._require_row(
                connection,
                """
                SELECT p.*,v.name AS vendor_name,e.full_name AS buyer_name
                FROM purchase_orders p JOIN vendors v ON v.id=p.vendor_id
                JOIN employees e ON e.id=p.buyer_employee_id
                WHERE p.po_number=?
                """,
                (po_number,),
                "Purchase order",
            )
            lines = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT l.*,i.description,
                           ROUND(l.order_quantity*l.unit_price,2) AS ordered_value,
                           ROUND((l.order_quantity-l.received_quantity)*
                                 l.unit_price,2) AS open_value,
                           ROUND((l.unit_price-l.standard_unit_cost)*
                                 l.received_quantity,2) AS purchase_price_variance
                    FROM purchase_order_lines l JOIN items i ON i.id=l.item_id
                    WHERE l.po_id=? ORDER BY l.line_number
                    """,
                    (header["id"],),
                )
            ]
            receipts = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT r.receipt_number,r.receipt_date,r.status,
                           rl.item_id,rl.quantity,rl.unit_price,
                           rl.standard_unit_cost,rl.lot_number,rl.quality_status
                    FROM receipts r JOIN receipt_lines rl ON rl.receipt_id=r.id
                    WHERE r.po_id=? ORDER BY r.receipt_date,r.id
                    """,
                    (header["id"],),
                )
            ]
            invoices = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT * FROM vendor_invoices WHERE po_id=?
                    ORDER BY invoice_date,id
                    """,
                    (header["id"],),
                )
            ]
            return {
                "header": dict(header),
                "lines": lines,
                "receipts": receipts,
                "vendor_invoices": invoices,
            }

    def get_open_purchase_commitments(
        self, as_of: str = SNAPSHOT_DATE, site_code: str | None = None
    ) -> dict[str, Any]:
        clauses = [
            "p.order_date<=?",
            "l.order_quantity>l.received_quantity",
            "p.status='Open'",
        ]
        params: list[Any] = [as_of]
        if site_code:
            clauses.append("p.site_code=?")
            params.append(site_code)
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT p.site_code,v.category AS commodity,
                           COUNT(DISTINCT p.id) AS purchase_order_count,
                           ROUND(SUM(l.order_quantity-l.received_quantity),4)
                               AS open_quantity,
                           ROUND(SUM((l.order_quantity-l.received_quantity)*
                                     l.unit_price),2) AS open_commitment
                    FROM purchase_orders p
                    JOIN purchase_order_lines l ON l.po_id=p.id
                    JOIN vendors v ON v.id=p.vendor_id
                    WHERE {' AND '.join(clauses)}
                    GROUP BY p.site_code,v.category
                    ORDER BY open_commitment DESC
                    """,
                    tuple(params),
                )
            ]
            totals = connection.execute(
                f"""
                SELECT COUNT(*) AS open_po_lines,
                       SUM(l.order_quantity-l.received_quantity)
                           AS open_quantity,
                       SUM((l.order_quantity-l.received_quantity)*
                           l.unit_price) AS open_commitment,
                       SUM(CASE WHEN p.expected_date<?
                                THEN (l.order_quantity-l.received_quantity)*
                                     l.unit_price ELSE 0 END)
                           AS overdue_commitment,
                       SUM(CASE WHEN p.expected_date BETWEEN ? AND date(?,'+30 days')
                                THEN (l.order_quantity-l.received_quantity)*
                                     l.unit_price ELSE 0 END)
                           AS next_30_day_commitment
                FROM purchase_orders p
                JOIN purchase_order_lines l ON l.po_id=p.id
                JOIN vendors v ON v.id=p.vendor_id
                WHERE {' AND '.join(clauses)}
                """,
                (as_of, as_of, as_of, *params),
            ).fetchone()
            top_vendor = connection.execute(
                f"""
                SELECT p.vendor_id,v.name,
                       SUM((l.order_quantity-l.received_quantity)*
                           l.unit_price) AS open_commitment
                FROM purchase_orders p
                JOIN purchase_order_lines l ON l.po_id=p.id
                JOIN vendors v ON v.id=p.vendor_id
                WHERE {' AND '.join(clauses)}
                GROUP BY p.vendor_id,v.name
                ORDER BY open_commitment DESC,p.vendor_id ASC
                LIMIT 1
                """,
                tuple(params),
            ).fetchone()
            return {
                "as_of": as_of,
                "rows": rows,
                "open_po_lines": int(totals["open_po_lines"] or 0),
                "open_quantity": _qty(totals["open_quantity"]),
                "open_commitment": _money(totals["open_commitment"]),
                "total_open_commitment": _money(totals["open_commitment"]),
                "overdue_commitment": _money(
                    totals["overdue_commitment"]
                ),
                "next_30_day_commitment": _money(
                    totals["next_30_day_commitment"]
                ),
                "top_vendor": (
                    str(top_vendor["name"]) if top_vendor else ""
                ),
                "top_vendor_commitment": _money(
                    top_vendor["open_commitment"] if top_vendor else 0
                ),
            }

    def get_purchase_price_variance(
        self,
        date_from: str,
        date_to: str,
        site_code: str | None = None,
        vendor_id: str | None = None,
        group_by: str = "vendor",
    ) -> dict[str, Any]:
        dimensions = {
            "vendor": ("vi.vendor_id", "v.name"),
            "item": ("l.item_id", "i.description"),
            "site": ("l.site_code", "s.name"),
            "commodity": ("v.category", "v.category"),
        }
        if group_by not in dimensions:
            raise ERPError(f"Unsupported PPV grouping: {group_by}")
        key, label = dimensions[group_by]
        clauses = ["vi.invoice_date BETWEEN ? AND ?"]
        params: list[Any] = [date_from, date_to]
        if site_code:
            clauses.append("l.site_code=?")
            params.append(site_code)
        if vendor_id:
            clauses.append("vi.vendor_id=?")
            params.append(vendor_id)
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT {key} AS dimension,{label} AS description,
                           COUNT(DISTINCT vi.id) AS invoice_count,
                           ROUND(SUM(l.quantity),4) AS invoiced_quantity,
                           ROUND(SUM(l.quantity*l.unit_price),2) AS actual_cost,
                           ROUND(SUM(l.quantity*pol.standard_unit_cost),2)
                               AS standard_cost,
                           ROUND(SUM(l.quantity*(
                                     l.unit_price-pol.standard_unit_cost)),2)
                               AS purchase_price_variance
                    FROM vendor_invoice_lines l
                    JOIN vendor_invoices vi ON vi.id=l.vendor_invoice_id
                    JOIN purchase_order_lines pol ON pol.id=l.po_line_id
                    JOIN vendors v ON v.id=vi.vendor_id
                    JOIN items i ON i.id=l.item_id
                    JOIN sites s ON s.code=l.site_code
                    WHERE {' AND '.join(clauses)}
                    GROUP BY {key},{label}
                    ORDER BY ABS(SUM(l.quantity*(
                                     l.unit_price-pol.standard_unit_cost))) DESC
                    """,
                    tuple(params),
                )
            ]
            totals = dict(
                connection.execute(
                    f"""
                    SELECT COUNT(DISTINCT vi.id) AS invoice_count,
                           ROUND(SUM(l.quantity),4) AS purchase_quantity,
                           ROUND(SUM(l.quantity*l.unit_price),2) AS actual_cost,
                           ROUND(SUM(l.quantity*pol.standard_unit_cost),2)
                               AS standard_cost
                    FROM vendor_invoice_lines l
                    JOIN vendor_invoices vi ON vi.id=l.vendor_invoice_id
                    JOIN purchase_order_lines pol ON pol.id=l.po_line_id
                    JOIN vendors v ON v.id=vi.vendor_id
                    JOIN items i ON i.id=l.item_id
                    JOIN sites s ON s.code=l.site_code
                    WHERE {' AND '.join(clauses)}
                    """,
                    tuple(params),
                ).fetchone()
            )
            total_ppv = _money(
                float(totals["actual_cost"] or 0)
                - float(totals["standard_cost"] or 0)
            )
            return {
                "date_from": date_from,
                "date_to": date_to,
                "group_by": group_by,
                "rows": rows,
                "invoice_count": int(totals["invoice_count"] or 0),
                "purchase_quantity": _qty(totals["purchase_quantity"]),
                "actual_cost": _money(totals["actual_cost"]),
                "standard_cost": _money(totals["standard_cost"]),
                "purchase_price_variance": total_ppv,
                "ppv_rate": _ratio(total_ppv, totals["standard_cost"]),
                "total_ppv": total_ppv,
            }

    def get_supplier_performance(
        self,
        date_from: str,
        date_to: str,
        vendor_id: str | None = None,
        limit: int = 200,
    ) -> list[dict[str, Any]]:
        clauses = ["r.receipt_date BETWEEN ? AND ?"]
        params: list[Any] = [date_from, date_to]
        if vendor_id:
            clauses.append("p.vendor_id=?")
            params.append(vendor_id)
        params.append(self._limit(limit, 1000))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    WITH performance AS (
                        SELECT v.id AS vendor_id,v.name,
                               v.category AS commodity,
                               COUNT(DISTINCT p.id) AS purchase_order_count,
                               COUNT(DISTINCT r.id) AS receipt_count,
                               COUNT(DISTINCT CASE
                                   WHEN r.receipt_date<=p.expected_date
                                   THEN r.id END) AS on_time_receipt_count,
                               COUNT(DISTINCT CASE
                                   WHEN rl.quality_status='Accepted'
                                   THEN r.id END) AS accepted_receipt_count,
                               ROUND(SUM(rl.quantity*rl.unit_price),4)
                                   AS receipt_value,
                               ROUND(AVG(julianday(r.receipt_date)-
                                         julianday(p.order_date)),2)
                                   AS average_actual_lead_days
                        FROM vendors v
                        JOIN purchase_orders p ON p.vendor_id=v.id
                        JOIN receipts r ON r.po_id=p.id
                        JOIN receipt_lines rl ON rl.receipt_id=r.id
                        WHERE {' AND '.join(clauses)}
                        GROUP BY v.id,v.name,v.category
                    )
                    SELECT *,
                           ROUND(100.0*on_time_receipt_count/
                                 NULLIF(receipt_count,0),2)
                               AS on_time_percent,
                           ROUND(100.0*accepted_receipt_count/
                                 NULLIF(receipt_count,0),2)
                               AS accepted_percent,
                           ROUND(1.0*on_time_receipt_count/
                                 NULLIF(receipt_count,0),6)
                               AS on_time_ratio,
                           ROUND(1.0*accepted_receipt_count/
                                 NULLIF(receipt_count,0),6)
                               AS accepted_ratio,
                           ROUND(
                               0.5*on_time_receipt_count/
                                   NULLIF(receipt_count,0)
                               + 0.5*accepted_receipt_count/
                                   NULLIF(receipt_count,0),
                               6
                           ) AS supplier_score
                    FROM performance
                    ORDER BY supplier_score ASC,vendor_id ASC
                    LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def search_customers(
        self,
        query: str | None = None,
        end_market: str | None = None,
        region: str | None = None,
        limit: int = 100,
    ) -> list[dict[str, Any]]:
        clauses = ["c.active=1"]
        params: list[Any] = []
        if query:
            clauses.append("(c.id LIKE ? OR c.name LIKE ?)")
            pattern = f"%{query}%"
            params.extend([pattern, pattern])
        if end_market:
            clauses.append("c.end_market=?")
            params.append(end_market)
        if region:
            clauses.append("c.region=?")
            params.append(region)
        params.append(self._limit(limit))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT c.*,
                           ROUND(COALESCE(SUM(ci.open_amount),0),2) AS open_ar
                    FROM customers c LEFT JOIN customer_invoices ci
                      ON ci.customer_id=c.id
                    WHERE {' AND '.join(clauses)}
                    GROUP BY c.id ORDER BY c.name LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def search_sales_orders(
        self,
        status: str | None = None,
        customer_id: str | None = None,
        product_family: str | None = None,
        promised_from: str | None = None,
        promised_to: str | None = None,
        limit: int = 100,
    ) -> list[dict[str, Any]]:
        clauses = ["1=1"]
        params: list[Any] = []
        if status:
            clauses.append("o.status=?")
            params.append(status)
        if customer_id:
            clauses.append("o.customer_id=?")
            params.append(customer_id)
        if product_family:
            clauses.append("i.product_family=?")
            params.append(product_family)
        if promised_from:
            clauses.append("o.promised_ship_date>=?")
            params.append(promised_from)
        if promised_to:
            clauses.append("o.promised_ship_date<=?")
            params.append(promised_to)
        params.append(self._limit(limit, 1000))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT o.id,o.order_number,o.customer_id,c.name AS customer_name,
                           o.order_date,o.requested_ship_date,o.promised_ship_date,
                           o.status,o.sales_region,
                           ROUND(SUM(l.order_quantity*l.unit_price*
                                     (1-l.discount_percent)),2) AS order_value,
                           ROUND(SUM((l.order_quantity-l.shipped_quantity)*
                                     l.unit_price*(1-l.discount_percent)),2)
                               AS backlog_value
                    FROM sales_orders o JOIN customers c ON c.id=o.customer_id
                    JOIN sales_order_lines l ON l.sales_order_id=o.id
                    JOIN items i ON i.id=l.item_id
                    WHERE {' AND '.join(clauses)}
                    GROUP BY o.id ORDER BY o.promised_ship_date,o.order_number LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def get_sales_order(self, order_number: str) -> dict[str, Any]:
        with self._connect() as connection:
            header = self._require_row(
                connection,
                """
                SELECT o.*,c.name AS customer_name,c.end_market,c.region,
                       c.payment_terms_days,c.credit_limit
                FROM sales_orders o JOIN customers c ON c.id=o.customer_id
                WHERE o.order_number=?
                """,
                (order_number,),
                "Sales order",
            )
            lines = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT l.*,i.description,i.product_family,
                           ROUND(l.order_quantity*l.unit_price*
                                 (1-l.discount_percent),2) AS net_order_value,
                           ROUND(l.shipped_quantity*l.standard_unit_cost,2)
                               AS shipped_standard_cost
                    FROM sales_order_lines l JOIN items i ON i.id=l.item_id
                    WHERE l.sales_order_id=? ORDER BY l.line_number
                    """,
                    (header["id"],),
                )
            ]
            shipments = [
                dict(row)
                for row in connection.execute(
                    "SELECT * FROM shipments WHERE sales_order_id=? ORDER BY ship_date",
                    (header["id"],),
                )
            ]
            invoices = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT * FROM customer_invoices
                    WHERE sales_order_id=? ORDER BY invoice_date
                    """,
                    (header["id"],),
                )
            ]
            return {
                "header": dict(header),
                "lines": lines,
                "shipments": shipments,
                "customer_invoices": invoices,
            }

    def get_backlog(
        self,
        as_of: str = SNAPSHOT_DATE,
        site_code: str | None = None,
        product_family: str | None = None,
    ) -> dict[str, Any]:
        clauses = [
            "o.order_date<=?",
            "l.order_quantity>l.shipped_quantity",
            "o.status='Open'",
        ]
        params: list[Any] = [as_of]
        if site_code:
            clauses.append("l.site_code=?")
            params.append(site_code)
        if product_family:
            clauses.append("i.product_family=?")
            params.append(product_family)
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT o.promised_ship_date,l.site_code,i.product_family,
                           COUNT(DISTINCT o.id) AS order_count,
                           ROUND(SUM(l.order_quantity-l.shipped_quantity),4)
                               AS backlog_quantity,
                           ROUND(SUM((l.order_quantity-l.shipped_quantity)*
                                     l.unit_price*(1-l.discount_percent)),2)
                               AS backlog_value
                    FROM sales_orders o JOIN sales_order_lines l
                      ON l.sales_order_id=o.id
                    JOIN items i ON i.id=l.item_id
                    WHERE {' AND '.join(clauses)}
                    GROUP BY o.promised_ship_date,l.site_code,i.product_family
                    ORDER BY o.promised_ship_date,l.site_code,i.product_family
                    """,
                    tuple(params),
                )
            ]
            total_backlog = connection.execute(
                f"""
                SELECT SUM(l.order_quantity-l.shipped_quantity),
                       SUM((l.order_quantity-l.shipped_quantity)*
                           l.unit_price*(1-l.discount_percent)),
                       COUNT(DISTINCT o.id),
                       SUM(CASE WHEN o.promised_ship_date<?
                                THEN (l.order_quantity-l.shipped_quantity)*
                                     l.unit_price*(1-l.discount_percent)
                                ELSE 0 END)
                FROM sales_orders o JOIN sales_order_lines l
                  ON l.sales_order_id=o.id
                JOIN items i ON i.id=l.item_id
                WHERE {' AND '.join(clauses)}
                """,
                (as_of, *params),
            ).fetchone()
            top_customer = connection.execute(
                f"""
                SELECT o.customer_id,c.name,
                       SUM((l.order_quantity-l.shipped_quantity)*
                           l.unit_price*(1-l.discount_percent)) AS backlog_value
                FROM sales_orders o JOIN sales_order_lines l
                  ON l.sales_order_id=o.id
                JOIN items i ON i.id=l.item_id
                JOIN customers c ON c.id=o.customer_id
                WHERE {' AND '.join(clauses)}
                GROUP BY o.customer_id,c.name
                ORDER BY backlog_value DESC,o.customer_id
                LIMIT 1
                """,
                tuple(params),
            ).fetchone()
            return {
                "as_of": as_of,
                "rows": rows,
                "total_backlog_quantity": _qty(total_backlog[0]),
                "total_backlog": _money(total_backlog[1]),
                "open_order_count": int(total_backlog[2] or 0),
                "overdue_backlog": _money(total_backlog[3]),
                "top_customer": (
                    str(top_customer["name"]) if top_customer else ""
                ),
                "top_customer_backlog": _money(
                    top_customer["backlog_value"] if top_customer else 0
                ),
            }

    def get_shipments(
        self,
        date_from: str,
        date_to: str,
        customer_id: str | None = None,
        site_code: str | None = None,
        limit: int = 500,
    ) -> list[dict[str, Any]]:
        clauses = ["s.ship_date BETWEEN ? AND ?"]
        params: list[Any] = [date_from, date_to]
        if customer_id:
            clauses.append("o.customer_id=?")
            params.append(customer_id)
        if site_code:
            clauses.append("l.site_code=?")
            params.append(site_code)
        params.append(self._limit(limit, 2000))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT s.shipment_number,s.ship_date,s.carrier,s.freight_amount,
                           o.order_number,o.customer_id,c.name AS customer_name,
                           l.item_id,l.site_code,l.quantity,l.unit_standard_cost,
                           ROUND(l.quantity*l.unit_standard_cost,2)
                               AS standard_cost
                    FROM shipments s JOIN sales_orders o ON o.id=s.sales_order_id
                    JOIN customers c ON c.id=o.customer_id
                    JOIN shipment_lines l ON l.shipment_id=s.id
                    WHERE {' AND '.join(clauses)}
                    ORDER BY s.ship_date DESC,s.id DESC LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def get_customer_product_profitability(
        self,
        date_from: str,
        date_to: str,
        group_by: str = "customer",
        limit: int = 200,
    ) -> dict[str, Any]:
        dimensions = {
            "customer": ("ci.customer_id", "c.name"),
            "product_family": ("l.product_family", "l.product_family"),
            "item": ("l.item_id", "i.description"),
            "end_market": ("c.end_market", "c.end_market"),
            "region": ("c.region", "c.region"),
        }
        if group_by not in dimensions:
            raise ERPError(f"Unsupported profitability grouping: {group_by}")
        key, label = dimensions[group_by]
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT {key} AS dimension,{label} AS description,
                           ROUND(SUM(l.quantity),4) AS units,
                           ROUND(SUM(l.amount),2) AS revenue,
                           ROUND(SUM(
                               l.quantity*s.standard_material_cost
                           ),2) AS material_cost,
                           ROUND(SUM(l.quantity*(
                               s.standard_material_cost+s.standard_labor_cost+
                               s.standard_variable_overhead+s.standard_fixed_overhead+
                               s.standard_outside_processing)),2) AS standard_cost,
                           ROUND(SUM(l.amount-l.quantity*(
                               s.standard_material_cost+s.standard_labor_cost+
                               s.standard_variable_overhead+s.standard_fixed_overhead+
                               s.standard_outside_processing)),2) AS gross_profit,
                           ROUND(100.0*SUM(l.amount-l.quantity*(
                               s.standard_material_cost+s.standard_labor_cost+
                               s.standard_variable_overhead+s.standard_fixed_overhead+
                               s.standard_outside_processing))/
                               NULLIF(SUM(l.amount),0),2) AS gross_margin_percent
                    FROM customer_invoices ci
                    JOIN customer_invoice_lines l ON l.invoice_id=ci.id
                    JOIN customers c ON c.id=ci.customer_id
                    JOIN items i ON i.id=l.item_id
                    JOIN item_sites s ON s.item_id=l.item_id
                                     AND s.site_code=l.site_code
                    WHERE ci.invoice_date BETWEEN ? AND ?
                    GROUP BY {key},{label}
                    ORDER BY gross_profit DESC LIMIT ?
                    """,
                    (date_from, date_to, self._limit(limit, 1000)),
                )
            ]
            return {
                "date_from": date_from,
                "date_to": date_to,
                "group_by": group_by,
                "rows": rows,
                "totals": {
                    "units": _qty(sum(row["units"] for row in rows)),
                    "revenue": _money(sum(row["revenue"] for row in rows)),
                    "material_cost": _money(
                        sum(row["material_cost"] for row in rows)
                    ),
                    "standard_cost": _money(sum(row["standard_cost"] for row in rows)),
                    "gross_profit": _money(sum(row["gross_profit"] for row in rows)),
                },
            }

    def get_demand_forecast(
        self,
        forecast_version: str,
        month_from: str | None = None,
        month_to: str | None = None,
        product_family: str | None = None,
        site_code: str | None = None,
    ) -> list[dict[str, Any]]:
        clauses = ["f.forecast_version=?"]
        params: list[Any] = [forecast_version]
        if month_from:
            clauses.append("f.forecast_month>=?")
            params.append(month_from)
        if month_to:
            clauses.append("f.forecast_month<=?")
            params.append(month_to)
        if product_family:
            clauses.append("i.product_family=?")
            params.append(product_family)
        if site_code:
            clauses.append("f.site_code=?")
            params.append(site_code)
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT f.forecast_version,f.forecast_month,f.site_code,
                           i.product_family,
                           ROUND(SUM(f.quantity),4) AS forecast_quantity,
                           ROUND(SUM(f.quantity*f.unit_price),2) AS forecast_revenue
                    FROM demand_forecast f JOIN items i ON i.id=f.item_id
                    WHERE {' AND '.join(clauses)}
                    GROUP BY f.forecast_version,f.forecast_month,f.site_code,
                             i.product_family
                    ORDER BY f.forecast_month,f.site_code,i.product_family
                    """,
                    tuple(params),
                )
            ]
            totals = connection.execute(
                f"""
                SELECT COUNT(DISTINCT f.item_id),
                       COUNT(DISTINCT f.forecast_month),
                       SUM(f.quantity),
                       SUM(f.quantity*f.unit_price)
                FROM demand_forecast f
                JOIN items i ON i.id=f.item_id
                WHERE {' AND '.join(clauses)}
                """,
                tuple(params),
            ).fetchone()
            scope = {
                "scope_item_count": int(totals[0] or 0),
                "scope_forecast_months": int(totals[1] or 0),
                "scope_forecast_quantity": _qty(totals[2]),
                "scope_forecast_revenue": _money(totals[3]),
            }
            for row in rows:
                row.update(scope)
            return rows

    def get_planned_orders(
        self,
        plan_version: str,
        order_type: str | None = None,
        site_code: str | None = None,
        action_message: str | None = None,
        required_to: str | None = None,
        limit: int = 500,
    ) -> list[dict[str, Any]]:
        clauses = ["p.plan_version=?"]
        params: list[Any] = [plan_version]
        for column, value in (
            ("p.order_type", order_type),
            ("p.site_code", site_code),
            ("p.action_message", action_message),
        ):
            if value:
                clauses.append(f"{column}=?")
                params.append(value)
        if required_to:
            clauses.append("p.required_date<=?")
            params.append(required_to)
        params.append(self._limit(limit, 2000))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT p.*,i.description,i.product_family,
                           ROUND(p.quantity*(
                               s.standard_material_cost+s.standard_labor_cost+
                               s.standard_variable_overhead+s.standard_fixed_overhead+
                               s.standard_outside_processing),2) AS planned_value
                    FROM planned_orders p JOIN items i ON i.id=p.item_id
                    JOIN item_sites s ON s.item_id=p.item_id
                                     AND s.site_code=p.site_code
                    WHERE {' AND '.join(clauses)}
                    ORDER BY p.required_date,p.id LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def get_planning_exceptions(
        self,
        plan_version: str,
        site_code: str | None = None,
    ) -> dict[str, Any]:
        clauses = ["p.plan_version=?", "p.action_message<>'None'"]
        params: list[Any] = [plan_version]
        if site_code:
            clauses.append("p.site_code=?")
            params.append(site_code)
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT p.action_message,p.site_code,p.order_type,
                           COUNT(*) AS exception_count,
                           ROUND(SUM(p.quantity),4) AS quantity,
                           MIN(p.required_date) AS earliest_required_date
                    FROM planned_orders p
                    WHERE {' AND '.join(clauses)}
                    GROUP BY p.action_message,p.site_code,p.order_type
                    ORDER BY exception_count DESC
                    """,
                    tuple(params),
                )
            ]
            return {
                "plan_version": plan_version,
                "rows": rows,
                "exception_count": sum(row["exception_count"] for row in rows),
            }

    # ------------------------------------------------------------------
    # Finance and accounting reports
    # ------------------------------------------------------------------

    def get_trial_balance(self, as_of: str = SNAPSHOT_DATE) -> dict[str, Any]:
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT a.code AS account_code,a.name,a.account_type,
                           a.normal_balance,a.financial_statement,
                           ROUND(COALESCE(SUM(CASE WHEN h.id IS NOT NULL
                                                   THEN l.debit ELSE 0 END),0),2)
                               AS total_debits,
                           ROUND(COALESCE(SUM(CASE WHEN h.id IS NOT NULL
                                                   THEN l.credit ELSE 0 END),0),2)
                               AS total_credits,
                           ROUND(COALESCE(SUM(CASE WHEN h.id IS NOT NULL
                                                   THEN l.debit-l.credit ELSE 0 END),0),2)
                               AS debit_less_credit
                    FROM accounts a
                    LEFT JOIN journal_lines l ON l.account_code=a.code
                    LEFT JOIN journal_headers h ON h.id=l.journal_id
                                                AND h.status='Posted'
                                                AND h.posting_date<=?
                    GROUP BY a.code ORDER BY a.code
                    """,
                    (as_of,),
                )
            ]
            return {
                "as_of": as_of,
                "rows": rows,
                "total_debits": _money(sum(row["total_debits"] for row in rows)),
                "total_credits": _money(sum(row["total_credits"] for row in rows)),
            }

    def get_profit_and_loss(
        self,
        date_from: str,
        date_to: str,
        site_code: str | None = None,
    ) -> dict[str, Any]:
        clauses = [
            "h.status='Posted'",
            "h.posting_date BETWEEN ? AND ?",
            "a.financial_statement='Income Statement'",
        ]
        params: list[Any] = [date_from, date_to]
        if site_code:
            clauses.append("(l.site_code=? OR l.site_code IS NULL)")
            params.append(site_code)
        with self._connect() as connection:
            raw_rows = [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT a.code AS account_code,a.name,a.account_type,
                           SUM(CASE WHEN a.account_type='Revenue'
                                    THEN l.credit-l.debit
                                    ELSE l.debit-l.credit END) AS amount
                    FROM journal_lines l JOIN journal_headers h ON h.id=l.journal_id
                    JOIN accounts a ON a.code=l.account_code
                    WHERE {' AND '.join(clauses)}
                    GROUP BY a.code,a.name,a.account_type
                    ORDER BY a.code
                    """,
                    tuple(params),
                )
            ]
            rows = [
                {**row, "amount": _money(row["amount"])}
                for row in raw_rows
            ]
            raw_revenue = sum(
                row["amount"]
                for row in raw_rows
                if row["account_type"] == "Revenue"
            )
            raw_expenses = sum(
                row["amount"]
                for row in raw_rows
                if row["account_type"] != "Revenue"
            )
            raw_cogs = sum(
                row["amount"]
                for row in raw_rows
                if 5000 <= int(row["account_code"]) < 6000
            )
            raw_operating_expense = sum(
                row["amount"]
                for row in raw_rows
                if 6000 <= int(row["account_code"]) < 7000
            )
            raw_gross_profit = raw_revenue - raw_cogs
            return {
                "date_from": date_from,
                "date_to": date_to,
                "site_code": site_code,
                "rows": rows,
                "revenue": _money(raw_revenue),
                "cost_of_goods_sold": _money(raw_cogs),
                "gross_profit": _money(raw_gross_profit),
                "gross_margin": _ratio(raw_gross_profit, raw_revenue),
                "operating_expense": _money(raw_operating_expense),
                "operating_income": _money(
                    raw_gross_profit - raw_operating_expense
                ),
                "expenses": _money(raw_expenses),
                "net_income": _money(raw_revenue - raw_expenses),
                "aggregation_note": (
                    "Summary values aggregate unrounded journal amounts and "
                    "round once; displayed account rows are rounded separately."
                ),
            }

    def get_operating_scenario_baseline(
        self,
        date_from: str,
        date_to: str,
        as_of: str = SNAPSHOT_DATE,
    ) -> dict[str, Any]:
        """Return the auditable components used by integrated forecast tasks."""

        with self._connect() as connection:
            invoice = connection.execute(
                """
                SELECT SUM(l.amount) AS revenue,
                       SUM(l.quantity*(
                           s.standard_material_cost+
                           s.standard_labor_cost+
                           s.standard_variable_overhead+
                           s.standard_fixed_overhead+
                           s.standard_outside_processing)) AS standard_cost,
                       SUM(l.quantity*s.standard_material_cost)
                           AS material_cost
                FROM customer_invoices ci
                JOIN customer_invoice_lines l ON l.invoice_id=ci.id
                JOIN item_sites s ON s.item_id=l.item_id
                                 AND s.site_code=l.site_code
                WHERE ci.invoice_date BETWEEN ? AND ?
                """,
                (date_from, date_to),
            ).fetchone()
            operating_expense = connection.execute(
                """
                SELECT SUM(l.debit-l.credit)
                FROM journal_headers h
                JOIN journal_lines l ON l.journal_id=h.id
                JOIN accounts a ON a.code=l.account_code
                WHERE h.status='Posted'
                  AND h.posting_date BETWEEN ? AND ?
                  AND a.code>='6000' AND a.code<'7000'
                """,
                (date_from, date_to),
            ).fetchone()[0]
            balances = {
                row["code"]: float(row["balance"] or 0)
                for row in connection.execute(
                    """
                    SELECT a.code,
                           SUM(CASE WHEN a.normal_balance='Debit'
                                    THEN l.debit-l.credit
                                    ELSE l.credit-l.debit END) AS balance
                    FROM journal_headers h
                    JOIN journal_lines l ON l.journal_id=h.id
                    JOIN accounts a ON a.code=l.account_code
                    WHERE h.status='Posted' AND h.posting_date<=?
                      AND a.code IN ('1100','1210','2000','2010')
                    GROUP BY a.code,a.normal_balance
                    """,
                    (as_of,),
                )
            }
            inventory = connection.execute(
                """
                SELECT SUM(on_hand_quantity*standard_unit_cost)
                FROM inventory_balances
                """
            ).fetchone()[0]
            revenue = float(invoice["revenue"] or 0)
            standard_cost = float(invoice["standard_cost"] or 0)
            material_cost = float(invoice["material_cost"] or 0)
            working_capital = (
                balances.get("1100", 0.0)
                + float(inventory or 0)
                + balances.get("1210", 0.0)
                - balances.get("2000", 0.0)
                - balances.get("2010", 0.0)
            )
            return {
                "date_from": date_from,
                "date_to": date_to,
                "as_of": as_of,
                "baseline_revenue": _money(revenue),
                "baseline_standard_cost": _money(standard_cost),
                "baseline_material_cost": _money(material_cost),
                "baseline_nonmaterial_cost": _money(
                    standard_cost - material_cost
                ),
                "baseline_gross_profit": _money(
                    revenue - standard_cost
                ),
                "baseline_gross_margin": _ratio(
                    revenue - standard_cost, revenue
                ),
                "baseline_operating_expense": _money(operating_expense),
                "operating_working_capital": _money(working_capital),
                "operating_working_capital_rate": (
                    round(working_capital / revenue, 12)
                    if revenue
                    else 0.0
                ),
            }

    def get_balance_sheet(self, as_of: str = SNAPSHOT_DATE) -> dict[str, Any]:
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT a.code AS account_code,a.name,a.account_type,
                           a.normal_balance,
                           ROUND(SUM(CASE WHEN a.account_type='Asset'
                                          THEN l.debit-l.credit
                                          ELSE l.credit-l.debit END),2) AS balance
                    FROM accounts a JOIN journal_lines l ON l.account_code=a.code
                    JOIN journal_headers h ON h.id=l.journal_id
                    WHERE h.status='Posted' AND h.posting_date<=?
                      AND a.financial_statement='Balance Sheet'
                    GROUP BY a.code,a.name,a.account_type,a.normal_balance
                    ORDER BY a.code
                    """,
                    (as_of,),
                )
            ]
            unclosed_earnings = _money(
                -connection.execute(
                    """
                    SELECT COALESCE(SUM(l.debit-l.credit),0)
                    FROM journal_lines l JOIN journal_headers h ON h.id=l.journal_id
                    JOIN accounts a ON a.code=l.account_code
                    WHERE h.status='Posted' AND h.posting_date<=?
                      AND a.financial_statement='Income Statement'
                    """,
                    (as_of,),
                ).fetchone()[0]
            )
            rows.append(
                {
                    "account_code": "3990",
                    "name": "Current and Unclosed Earnings",
                    "account_type": "Equity",
                    "normal_balance": "Credit",
                    "balance": unclosed_earnings,
                }
            )
            totals = {
                kind: _money(
                    sum(row["balance"] for row in rows if row["account_type"] == kind)
                )
                for kind in ("Asset", "Liability", "Equity")
            }
            return {"as_of": as_of, "rows": rows, "totals": totals}

    def get_gl_detail(
        self,
        date_from: str,
        date_to: str,
        account_code: str | None = None,
        source: str | None = None,
        reference: str | None = None,
        site_code: str | None = None,
        status: str = "Posted",
        limit: int = 500,
    ) -> list[dict[str, Any]]:
        clauses = ["h.posting_date BETWEEN ? AND ?", "h.status=?"]
        params: list[Any] = [date_from, date_to, status]
        for column, value in (
            ("l.account_code", account_code),
            ("h.source", source),
            ("h.reference", reference),
            ("l.site_code", site_code),
        ):
            if value:
                clauses.append(f"{column}=?")
                params.append(value)
        params.append(self._limit(limit, 2000))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT h.id AS journal_id,h.posting_date,h.period,h.source,
                           h.reference,h.memo,h.status,l.id AS line_id,
                           l.account_code,a.name AS account_name,l.department_code,
                           l.site_code,l.item_id,l.production_order_id,
                           l.customer_id,l.vendor_id,l.debit,l.credit,l.line_memo
                    FROM journal_headers h JOIN journal_lines l ON l.journal_id=h.id
                    JOIN accounts a ON a.code=l.account_code
                    WHERE {' AND '.join(clauses)}
                    ORDER BY h.posting_date DESC,h.id DESC,l.id LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def get_subledger_reconciliation(
        self, as_of: str = SNAPSHOT_DATE
    ) -> dict[str, Any]:
        with self._connect() as connection:
            ar = _money(
                connection.execute(
                    "SELECT COALESCE(SUM(open_amount),0) FROM customer_invoices "
                    "WHERE invoice_date<=?",
                    (as_of,),
                ).fetchone()[0]
            )
            ap = _money(
                connection.execute(
                    "SELECT COALESCE(SUM(open_amount),0) FROM vendor_invoices "
                    "WHERE invoice_date<=?",
                    (as_of,),
                ).fetchone()[0]
            )
            inventory = _money(
                connection.execute(
                    """
                    SELECT COALESCE(SUM(on_hand_quantity*standard_unit_cost),0)
                    FROM inventory_balances
                    """
                ).fetchone()[0]
            )
            wip = _money(
                connection.execute(
                    """
                    SELECT COALESCE(SUM(
                        actual_material_cost+actual_labor_cost+
                        actual_variable_overhead+actual_fixed_overhead+
                        actual_outside_processing-
                        completed_quantity*standard_unit_cost
                    ),0) FROM production_orders WHERE status<>'Ended'
                    """
                ).fetchone()[0]
            )

            def gl_balance(accounts: tuple[str, ...], credit_normal: bool = False) -> float:
                placeholders = ",".join("?" for _ in accounts)
                value = connection.execute(
                    f"""
                    SELECT COALESCE(SUM(l.debit-l.credit),0)
                    FROM journal_lines l JOIN journal_headers h ON h.id=l.journal_id
                    WHERE h.status='Posted' AND h.posting_date<=?
                      AND l.account_code IN ({placeholders})
                    """,
                    (as_of, *accounts),
                ).fetchone()[0]
                return _money(-value if credit_normal else value)

            rows = [
                ("Accounts Receivable", ar, gl_balance(("1100",))),
                ("Accounts Payable", ap, gl_balance(("2000",), True)),
                (
                    "Discrete Inventory",
                    inventory,
                    gl_balance(("1200", "1220", "1230")),
                ),
                ("Work in Process", wip, gl_balance(("1210",))),
            ]
            return {
                "as_of": as_of,
                "rows": [
                    {
                        "control": name,
                        "subledger": subledger,
                        "general_ledger": general_ledger,
                        "difference": _money(subledger - general_ledger),
                    }
                    for name, subledger, general_ledger in rows
                ],
            }

    def get_ar_aging(self, as_of: str = SNAPSHOT_DATE) -> dict[str, Any]:
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT CASE
                               WHEN due_date>=? THEN 'Current'
                               WHEN julianday(?) - julianday(due_date)<=30 THEN '1-30'
                               WHEN julianday(?) - julianday(due_date)<=60 THEN '31-60'
                               WHEN julianday(?) - julianday(due_date)<=90 THEN '61-90'
                               ELSE 'Over 90'
                           END AS aging_bucket,
                           COUNT(*) AS invoice_count,
                           ROUND(SUM(open_amount),2) AS open_amount
                    FROM customer_invoices
                    WHERE invoice_date<=? AND open_amount<>0
                    GROUP BY aging_bucket
                    ORDER BY MIN(julianday(?) - julianday(due_date))
                    """,
                    (as_of, as_of, as_of, as_of, as_of, as_of),
                )
            ]
            return {
                "as_of": as_of,
                "rows": rows,
                "total_open_ar": _money(sum(row["open_amount"] for row in rows)),
            }

    def get_ap_aging(self, as_of: str = SNAPSHOT_DATE) -> dict[str, Any]:
        with self._connect() as connection:
            rows = [
                dict(row)
                for row in connection.execute(
                    """
                    SELECT CASE
                               WHEN due_date>=? THEN 'Current'
                               WHEN julianday(?) - julianday(due_date)<=30 THEN '1-30'
                               WHEN julianday(?) - julianday(due_date)<=60 THEN '31-60'
                               WHEN julianday(?) - julianday(due_date)<=90 THEN '61-90'
                               ELSE 'Over 90'
                           END AS aging_bucket,
                           COUNT(*) AS invoice_count,
                           ROUND(SUM(open_amount),2) AS open_amount
                    FROM vendor_invoices
                    WHERE invoice_date<=? AND open_amount<>0
                    GROUP BY aging_bucket
                    ORDER BY MIN(julianday(?) - julianday(due_date))
                    """,
                    (as_of, as_of, as_of, as_of, as_of, as_of),
                )
            ]
            return {
                "as_of": as_of,
                "rows": rows,
                "total_open_ap": _money(sum(row["open_amount"] for row in rows)),
            }

    def get_cash_and_debt(self, as_of: str = SNAPSHOT_DATE) -> dict[str, Any]:
        with self._connect() as connection:
            cash = _money(
                connection.execute(
                    """
                    SELECT COALESCE(SUM(l.debit-l.credit),0)
                    FROM journal_lines l JOIN journal_headers h ON h.id=l.journal_id
                    WHERE h.status='Posted' AND h.posting_date<=?
                      AND l.account_code IN ('1000','1010','1020')
                    """,
                    (as_of,),
                ).fetchone()[0]
            )
            debt = [
                dict(row)
                for row in connection.execute(
                    "SELECT * FROM debt_instruments WHERE status='Active' "
                    "ORDER BY maturity_date,id"
                )
            ]
            total_debt = _money(sum(row["outstanding_principal"] for row in debt))
            liquidity = _money(
                cash
                + sum(
                    row["commitment"] - row["outstanding_principal"]
                    for row in debt
                    if row["instrument_type"] == "Revolver"
                )
            )
            return {
                "as_of": as_of,
                "cash": cash,
                "debt_instruments": debt,
                "total_debt": total_debt,
                "net_debt": _money(total_debt - cash),
                "available_liquidity": liquidity,
            }

    def get_fixed_asset_register(
        self,
        site_code: str | None = None,
        asset_class: str | None = None,
        status: str | None = "Active",
        limit: int = 500,
    ) -> list[dict[str, Any]]:
        clauses = ["1=1"]
        params: list[Any] = []
        for column, value in (
            ("site_code", site_code),
            ("asset_class", asset_class),
            ("status", status),
        ):
            if value:
                clauses.append(f"{column}=?")
                params.append(value)
        params.append(self._limit(limit, 2000))
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT *,
                           ROUND(gross_cost-accumulated_depreciation,2)
                               AS net_book_value
                    FROM fixed_assets WHERE {' AND '.join(clauses)}
                    ORDER BY net_book_value DESC,id LIMIT ?
                    """,
                    tuple(params),
                )
            ]

    def get_payroll_summary(
        self,
        date_from: str,
        date_to: str,
        site_code: str | None = None,
        department_code: str | None = None,
    ) -> list[dict[str, Any]]:
        clauses = ["r.check_date BETWEEN ? AND ?"]
        params: list[Any] = [date_from, date_to]
        if site_code:
            clauses.append("e.site_code=?")
            params.append(site_code)
        if department_code:
            clauses.append("e.department_code=?")
            params.append(department_code)
        with self._connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    f"""
                    SELECT e.site_code,e.department_code,d.name AS department_name,
                           COUNT(DISTINCT l.employee_id) AS employee_count,
                           ROUND(SUM(l.regular_hours),2) AS regular_hours,
                           ROUND(SUM(l.overtime_hours),2) AS overtime_hours,
                           ROUND(SUM(l.gross_pay),2) AS gross_pay,
                           ROUND(SUM(l.employer_taxes),2) AS employer_taxes,
                           ROUND(SUM(l.benefits),2) AS benefits,
                           ROUND(SUM(l.gross_pay+l.employer_taxes+l.benefits),2)
                               AS total_employer_cost
                    FROM payroll_runs r JOIN payroll_lines l ON l.payroll_run_id=r.id
                    JOIN employees e ON e.id=l.employee_id
                    JOIN departments d ON d.code=e.department_code
                    WHERE {' AND '.join(clauses)}
                    GROUP BY e.site_code,e.department_code,d.name
                    ORDER BY total_employer_cost DESC
                    """,
                    tuple(params),
                )
            ]

    # ------------------------------------------------------------------
    # Controlled, transactional ERP workflows
    # ------------------------------------------------------------------

    def create_draft_journal(
        self,
        posting_date: str,
        reference: str,
        memo: str,
        lines: list[dict[str, Any]],
    ) -> dict[str, Any]:
        if not reference.strip() or not memo.strip():
            raise ERPError("Reference and memo are required")
        if len(lines) < 2:
            raise ERPError("A journal requires at least two lines")
        posting_date = _iso_date(posting_date, label="Posting date")
        period = posting_date[:7]
        with self._write() as connection:
            fiscal = self._require_row(
                connection,
                "SELECT * FROM fiscal_periods WHERE period=?",
                (period,),
                "Fiscal period",
            )
            if fiscal["status"] != "Open":
                raise ERPError(f"Fiscal period {period} is not open")
            if connection.execute(
                "SELECT 1 FROM journal_headers WHERE reference=?", (reference,)
            ).fetchone():
                raise ERPError(f"Journal reference {reference} already exists")
            normalized = []
            debit_cents = credit_cents = 0
            for position, line in enumerate(lines, 1):
                account = str(line.get("account_code", ""))
                self._require_row(
                    connection,
                    "SELECT code FROM accounts WHERE code=?",
                    (account,),
                    f"Account on line {position}",
                )
                debit = _money(line.get("debit", 0))
                credit = _money(line.get("credit", 0))
                if debit < 0 or credit < 0 or (debit and credit) or not (debit or credit):
                    raise ERPError(
                        f"Line {position} must contain one nonnegative debit or credit"
                    )
                for table, key, value in (
                    ("sites", "code", line.get("site_code")),
                    ("departments", "code", line.get("department_code")),
                    ("items", "id", line.get("item_id")),
                    ("customers", "id", line.get("customer_id")),
                    ("vendors", "id", line.get("vendor_id")),
                ):
                    if value and not connection.execute(
                        f"SELECT 1 FROM {table} WHERE {key}=?", (value,)
                    ).fetchone():
                        raise ERPError(f"Invalid {key} {value} on line {position}")
                debit_cents += int(round(debit * 100))
                credit_cents += int(round(credit * 100))
                normalized.append({**line, "debit": debit, "credit": credit})
            if debit_cents != credit_cents:
                raise ERPError(
                    f"Journal is not balanced: debits={debit_cents/100:.2f}, "
                    f"credits={credit_cents/100:.2f}"
                )
            cursor = connection.execute(
                """
                INSERT INTO journal_headers(
                    posting_date,period,source,reference,memo,status,created_at
                ) VALUES(?,?,'Agent Journal',?,?, 'Draft',?)
                """,
                (
                    posting_date,
                    period,
                    reference,
                    memo,
                    datetime.now(timezone.utc).isoformat(),
                ),
            )
            journal_id = int(cursor.lastrowid)
            for line in normalized:
                connection.execute(
                    """
                    INSERT INTO journal_lines(
                        journal_id,account_code,department_code,site_code,item_id,
                        production_order_id,customer_id,vendor_id,debit,credit,line_memo
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?)
                    """,
                    (
                        journal_id,
                        line["account_code"],
                        line.get("department_code"),
                        line.get("site_code"),
                        line.get("item_id"),
                        line.get("production_order_id"),
                        line.get("customer_id"),
                        line.get("vendor_id"),
                        line["debit"],
                        line["credit"],
                        line.get("line_memo", memo),
                    ),
                )
            self._audit(
                connection,
                self.actor,
                "create_draft",
                "journal",
                reference,
                {"journal_id": journal_id, "posting_date": posting_date, "lines": len(lines)},
            )
            return {
                "journal_id": journal_id,
                "reference": reference,
                "status": "Draft",
                "debits": debit_cents / 100,
                "credits": credit_cents / 100,
            }

    def post_draft_journal(self, reference: str) -> dict[str, Any]:
        with self._write() as connection:
            journal = self._require_row(
                connection,
                "SELECT * FROM journal_headers WHERE reference=?",
                (reference,),
                "Journal",
            )
            if journal["status"] != "Draft":
                raise ERPError(f"Journal {reference} is not a draft")
            period = self._require_row(
                connection,
                "SELECT * FROM fiscal_periods WHERE period=?",
                (journal["period"],),
                "Fiscal period",
            )
            if period["status"] != "Open":
                raise ERPError(f"Fiscal period {journal['period']} is not open")
            totals = connection.execute(
                """
                SELECT ROUND(SUM(debit),2),ROUND(SUM(credit),2)
                FROM journal_lines WHERE journal_id=?
                """,
                (journal["id"],),
            ).fetchone()
            if int(round(float(totals[0]) * 100)) != int(round(float(totals[1]) * 100)):
                raise ERPError("Draft journal is not balanced")
            connection.execute(
                "UPDATE journal_headers SET status='Posted' WHERE id=?",
                (journal["id"],),
            )
            self._audit(
                connection,
                self.actor,
                "post",
                "journal",
                reference,
                {"journal_id": journal["id"], "period": journal["period"]},
            )
            return {
                "journal_id": journal["id"],
                "reference": reference,
                "status": "Posted",
            }

    def create_purchase_order(
        self,
        vendor_id: str,
        site_code: str,
        order_date: str,
        expected_date: str,
        buyer_employee_id: str,
        lines: list[dict[str, Any]],
    ) -> dict[str, Any]:
        if not lines:
            raise ERPError("A purchase order requires at least one line")
        order_date = _iso_date(order_date, label="Order date")
        expected_date = _iso_date(expected_date, label="Expected date")
        if expected_date < order_date:
            raise ERPError("Expected date cannot precede order date")
        with self._write() as connection:
            vendor = self._require_row(
                connection,
                "SELECT * FROM vendors WHERE id=?",
                (vendor_id,),
                "Vendor",
            )
            if not vendor["active"]:
                raise ERPError("Vendor is inactive")
            self._require_row(
                connection, "SELECT code FROM sites WHERE code=?", (site_code,), "Site"
            )
            buyer = self._require_row(
                connection,
                "SELECT * FROM employees WHERE id=?",
                (buyer_employee_id,),
                "Buyer",
            )
            if not buyer["active"]:
                raise ERPError("Buyer is inactive")
            if (
                buyer["site_code"] != site_code
                or buyer["department_code"] != "SCM"
                or not any(
                    token in str(buyer["title"]).casefold()
                    for token in ("buyer", "planner")
                )
            ):
                raise ERPError(
                    "Buyer must be an active site-matched Supply Chain "
                    "planner or buyer"
                )
            po_id = int(
                connection.execute(
                    "SELECT COALESCE(MAX(id),0)+1 FROM purchase_orders"
                ).fetchone()[0]
            )
            po_number = f"PO-A{po_id:06d}"
            connection.execute(
                """
                INSERT INTO purchase_orders(
                    id,po_number,vendor_id,site_code,order_date,expected_date,
                    buyer_employee_id,status,currency
                ) VALUES(?,?,?,?,?,?,?,'Open','USD')
                """,
                (
                    po_id,
                    po_number,
                    vendor_id,
                    site_code,
                    order_date,
                    expected_date,
                    buyer_employee_id,
                ),
            )
            total = 0.0
            next_line_id = int(
                connection.execute(
                    "SELECT COALESCE(MAX(id),0)+1 FROM purchase_order_lines"
                ).fetchone()[0]
            )
            for position, line in enumerate(lines, 1):
                item_id = str(line.get("item_id", ""))
                site_item = self._require_row(
                    connection,
                    """
                    SELECT s.*,i.primary_vendor_id
                    FROM item_sites s
                    JOIN items i ON i.id=s.item_id
                    WHERE s.item_id=? AND s.site_code=?
                    """,
                    (item_id, site_code),
                    f"Item/site on line {position}",
                )
                if site_item["primary_vendor_id"] != vendor_id:
                    raise ERPError(
                        f"Vendor is not authorized as the primary vendor "
                        f"for item {item_id} on line {position}"
                    )
                quantity = _qty(line.get("quantity"))
                unit_price = _money(line.get("unit_price"))
                if quantity <= 0 or unit_price <= 0:
                    raise ERPError(f"Line {position} quantity and price must be positive")
                standard = _money(site_item["standard_material_cost"])
                connection.execute(
                    """
                    INSERT INTO purchase_order_lines(
                        id,po_id,line_number,item_id,order_quantity,
                        received_quantity,invoiced_quantity,unit_price,
                        standard_unit_cost,status
                    ) VALUES(?,?,?,?,?,0,0,?,?, 'Open')
                    """,
                    (
                        next_line_id,
                        po_id,
                        position * 10,
                        item_id,
                        quantity,
                        unit_price,
                        standard,
                    ),
                )
                next_line_id += 1
                total += quantity * unit_price
            self._audit(
                connection,
                self.actor,
                "create",
                "purchase_order",
                po_number,
                {
                    "vendor_id": vendor_id,
                    "site_code": site_code,
                    "line_count": len(lines),
                    "order_value": _money(total),
                },
            )
            return {
                "po_number": po_number,
                "status": "Open",
                "line_count": len(lines),
                "order_value": _money(total),
            }

    def create_production_order(
        self,
        item_id: str,
        site_code: str,
        order_quantity: float,
        scheduled_start: str,
        scheduled_finish: str,
        source_reference: str | None = None,
    ) -> dict[str, Any]:
        quantity = _qty(order_quantity)
        if quantity <= 0:
            raise ERPError("Order quantity must be positive")
        scheduled_start = _iso_date(
            scheduled_start,
            label="Scheduled start",
        )
        scheduled_finish = _iso_date(
            scheduled_finish,
            label="Scheduled finish",
        )
        if scheduled_finish < scheduled_start:
            raise ERPError("Scheduled finish cannot precede scheduled start")
        with self._write() as connection:
            source_type = "Manual"
            if (
                isinstance(source_reference, str)
                and source_reference.startswith("PLAN-")
            ):
                planned_order = self._require_row(
                    connection,
                    """
                    SELECT * FROM planned_orders
                    WHERE id=? AND plan_version='2026-06-SOP'
                    """,
                    (source_reference,),
                    "Approved planned order",
                )
                if (
                    planned_order["order_type"] != "Production"
                    or planned_order["status"] != "Planned"
                    or planned_order["item_id"] != item_id
                    or planned_order["site_code"] != site_code
                    or _qty(planned_order["quantity"]) != quantity
                    or planned_order["release_date"] != scheduled_start
                    or planned_order["required_date"] != scheduled_finish
                ):
                    raise ERPError(
                        "Production order does not exactly match the approved "
                        "planned order"
                    )
                source_type = "Planned Order"
            item = self._require_row(
                connection, "SELECT * FROM items WHERE id=?", (item_id,), "Item"
            )
            if item["make_buy"] not in {"Make", "Either"}:
                raise ERPError("Item is not configured for production")
            site_item = self._require_row(
                connection,
                "SELECT * FROM item_sites WHERE item_id=? AND site_code=?",
                (item_id, site_code),
                "Item/site",
            )
            bom = self._require_row(
                connection,
                """
                SELECT * FROM bom_headers
                WHERE item_id=? AND site_code=? AND status='Active'
                  AND effective_from<=?
                  AND (effective_to IS NULL OR effective_to>=?)
                ORDER BY effective_from DESC LIMIT 1
                """,
                (item_id, site_code, scheduled_start, scheduled_start),
                "Effective BOM",
            )
            routing = self._require_row(
                connection,
                """
                SELECT * FROM routing_headers
                WHERE item_id=? AND site_code=? AND status='Active'
                  AND effective_from<=?
                  AND (effective_to IS NULL OR effective_to>=?)
                ORDER BY effective_from DESC LIMIT 1
                """,
                (item_id, site_code, scheduled_start, scheduled_start),
                "Effective routing",
            )
            order_id = int(
                connection.execute(
                    "SELECT COALESCE(MAX(id),0)+1 FROM production_orders"
                ).fetchone()[0]
            )
            order_number = f"MO-A{order_id:06d}"
            standard_unit_cost = _money(
                sum(
                    float(site_item[column])
                    for column in (
                        "standard_material_cost",
                        "standard_labor_cost",
                        "standard_variable_overhead",
                        "standard_fixed_overhead",
                        "standard_outside_processing",
                    )
                )
            )
            connection.execute(
                """
                INSERT INTO production_orders(
                    id,order_number,item_id,site_code,bom_id,routing_id,
                    source_type,source_reference,created_date,scheduled_start,
                    scheduled_finish,actual_start,actual_finish,order_quantity,
                    completed_quantity,scrapped_quantity,status,standard_unit_cost,
                    actual_material_cost,actual_labor_cost,actual_variable_overhead,
                    actual_fixed_overhead,actual_outside_processing
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,NULL,NULL,?,0,0,
                         'Scheduled',?,0,0,0,0,0)
                """,
                (
                    order_id,
                    order_number,
                    item_id,
                    site_code,
                    bom["id"],
                    routing["id"],
                    source_type,
                    source_reference,
                    SNAPSHOT_DATE,
                    scheduled_start,
                    scheduled_finish,
                    quantity,
                    standard_unit_cost,
                ),
            )
            material_id = int(
                connection.execute(
                    "SELECT COALESCE(MAX(id),0)+1 FROM production_order_materials"
                ).fetchone()[0]
            )
            for line in connection.execute(
                "SELECT * FROM bom_lines WHERE bom_id=? ORDER BY line_number",
                (bom["id"],),
            ):
                cost = self._require_row(
                    connection,
                    """
                    SELECT standard_material_cost FROM item_sites
                    WHERE item_id=? AND site_code=?
                    """,
                    (line["component_item_id"], site_code),
                    "Component item/site",
                )
                connection.execute(
                    """
                    INSERT INTO production_order_materials(
                        id,production_order_id,line_number,component_item_id,
                        planned_quantity,issued_quantity,standard_unit_cost,
                        actual_unit_cost,issue_date,substitution
                    ) VALUES(?,?,?,?,?,0,?,?,NULL,0)
                    """,
                    (
                        material_id,
                        order_id,
                        line["line_number"],
                        line["component_item_id"],
                        _qty(
                            quantity
                            * line["quantity_per"]
                            * (1 + line["scrap_factor"])
                        ),
                        _money(cost["standard_material_cost"]),
                        _money(cost["standard_material_cost"]),
                    ),
                )
                material_id += 1
            operation_id = int(
                connection.execute(
                    "SELECT COALESCE(MAX(id),0)+1 FROM production_order_operations"
                ).fetchone()[0]
            )
            for operation in connection.execute(
                """
                SELECT o.*,w.labor_rate,w.machine_rate
                FROM routing_operations o JOIN work_centers w
                  ON w.code=o.work_center_code
                WHERE o.routing_id=? ORDER BY o.operation_number
                """,
                (routing["id"],),
            ):
                connection.execute(
                    """
                    INSERT INTO production_order_operations(
                        id,production_order_id,operation_number,work_center_code,
                        planned_setup_hours,planned_run_hours,actual_setup_hours,
                        actual_run_hours,labor_rate_standard,labor_rate_actual,
                        machine_rate_standard,completed_date,status
                    ) VALUES(?,?,?,?,?,?,0,0,?,?,?,NULL,'Scheduled')
                    """,
                    (
                        operation_id,
                        order_id,
                        operation["operation_number"],
                        operation["work_center_code"],
                        operation["setup_hours"],
                        _qty(operation["run_hours_per_unit"] * quantity),
                        operation["labor_rate"],
                        operation["labor_rate"],
                        operation["machine_rate"],
                    ),
                )
                operation_id += 1
            connection.execute(
                """
                INSERT INTO production_variances(
                    production_order_id,material_price_variance,
                    material_usage_variance,labor_rate_variance,
                    labor_efficiency_variance,variable_overhead_variance,
                    fixed_overhead_volume_variance,scrap_variance,
                    substitution_variance,total_variance,settled_date
                ) VALUES(?,0,0,0,0,0,0,0,0,0,NULL)
                """,
                (order_id,),
            )
            self._audit(
                connection,
                self.actor,
                "create",
                "production_order",
                order_number,
                {
                    "item_id": item_id,
                    "site_code": site_code,
                    "order_quantity": quantity,
                    "source_type": source_type,
                    "source_reference": source_reference,
                },
            )
            return {
                "order_number": order_number,
                "status": "Scheduled",
                "item_id": item_id,
                "site_code": site_code,
                "order_quantity": quantity,
                "standard_unit_cost": standard_unit_cost,
            }

    def update_production_order_status(
        self, order_number: str, new_status: str
    ) -> dict[str, Any]:
        transitions = {
            "Scheduled": {"Released"},
            "Released": {"Started", "Cancelled"},
            "Started": {"Reported Finished"},
        }
        with self._write() as connection:
            order = self._require_row(
                connection,
                "SELECT * FROM production_orders WHERE order_number=?",
                (order_number,),
                "Production order",
            )
            allowed = transitions.get(order["status"], set())
            if new_status not in allowed:
                raise ERPError(
                    f"Invalid production transition {order['status']} -> {new_status}"
                )
            updates = ["status=?"]
            params: list[Any] = [new_status]
            if new_status == "Started":
                updates.append("actual_start=?")
                params.append(max(SNAPSHOT_DATE, order["scheduled_start"]))
            params.append(order["id"])
            connection.execute(
                f"UPDATE production_orders SET {','.join(updates)} WHERE id=?",
                tuple(params),
            )
            self._audit(
                connection,
                self.actor,
                "status_change",
                "production_order",
                order_number,
                {"from": order["status"], "to": new_status},
            )
            return {
                "order_number": order_number,
                "previous_status": order["status"],
                "status": new_status,
            }

    def place_inventory_quality_hold(
        self,
        item_id: str,
        site_code: str,
        warehouse_code: str,
        location_code: str,
        quantity: float,
        reason: str,
        lot_number: str = "",
    ) -> dict[str, Any]:
        hold_quantity = _qty(quantity)
        if hold_quantity <= 0 or not reason.strip():
            raise ERPError("A positive quantity and reason are required")
        with self._write() as connection:
            balance = self._require_row(
                connection,
                """
                SELECT * FROM inventory_balances
                WHERE item_id=? AND site_code=? AND warehouse_code=?
                  AND location_code=? AND lot_number=?
                """,
                (item_id, site_code, warehouse_code, location_code, lot_number),
                "Inventory balance",
            )
            available = (
                balance["on_hand_quantity"]
                - balance["reserved_quantity"]
                - balance["quality_hold_quantity"]
            )
            if hold_quantity > available + 0.0001:
                raise ERPError(
                    f"Hold quantity exceeds available inventory ({available:.4f})"
                )
            connection.execute(
                """
                UPDATE inventory_balances
                SET quality_hold_quantity=quality_hold_quantity+?
                WHERE item_id=? AND site_code=? AND warehouse_code=?
                  AND location_code=? AND lot_number=?
                """,
                (
                    hold_quantity,
                    item_id,
                    site_code,
                    warehouse_code,
                    location_code,
                    lot_number,
                ),
            )
            sequence = int(
                connection.execute(
                    "SELECT COUNT(*)+1 FROM quality_orders WHERE id LIKE 'QO-A%'"
                ).fetchone()[0]
            )
            quality_order_id = f"QO-A{sequence:05d}"
            exposure = _money(hold_quantity * balance["standard_unit_cost"])
            connection.execute(
                """
                INSERT INTO quality_orders(
                    id,reference_type,reference_id,item_id,site_code,lot_number,
                    opened_date,closed_date,status,disposition,quantity_inspected,
                    quantity_failed,estimated_financial_exposure
                ) VALUES(?,'Inventory Hold',?,?,?,?,?,NULL,'Open',?, ?,?,?)
                """,
                (
                    quality_order_id,
                    f"{warehouse_code}/{location_code}",
                    item_id,
                    site_code,
                    lot_number,
                    SNAPSHOT_DATE,
                    reason,
                    hold_quantity,
                    0,
                    exposure,
                ),
            )
            self._audit(
                connection,
                self.actor,
                "place_hold",
                "inventory",
                quality_order_id,
                {
                    "item_id": item_id,
                    "site_code": site_code,
                    "warehouse_code": warehouse_code,
                    "location_code": location_code,
                    "lot_number": lot_number,
                    "quantity": hold_quantity,
                    "reason": reason,
                    "quality_hold_before": _qty(
                        balance["quality_hold_quantity"]
                    ),
                    "quality_hold_after": _qty(
                        balance["quality_hold_quantity"] + hold_quantity
                    ),
                },
            )
            return {
                "quality_order_id": quality_order_id,
                "status": "Open",
                "held_quantity": hold_quantity,
                "estimated_financial_exposure": exposure,
            }

    def release_inventory_quality_hold(
        self, quality_order_id: str, disposition: str
    ) -> dict[str, Any]:
        if not disposition.strip():
            raise ERPError("Disposition is required")
        with self._write() as connection:
            order = self._require_row(
                connection,
                "SELECT * FROM quality_orders WHERE id=?",
                (quality_order_id,),
                "Quality order",
            )
            if order["status"] != "Open" or order["reference_type"] != "Inventory Hold":
                raise ERPError("Quality order is not an open agent inventory hold")
            warehouse_code, location_code = order["reference_id"].split("/", 1)
            balance = self._require_row(
                connection,
                """
                SELECT * FROM inventory_balances
                WHERE item_id=? AND site_code=? AND warehouse_code=?
                  AND location_code=? AND lot_number=?
                """,
                (
                    order["item_id"],
                    order["site_code"],
                    warehouse_code,
                    location_code,
                    order["lot_number"] or "",
                ),
                "Inventory balance",
            )
            release_quantity = min(
                float(order["quantity_inspected"]),
                float(balance["quality_hold_quantity"]),
            )
            connection.execute(
                """
                UPDATE inventory_balances
                SET quality_hold_quantity=quality_hold_quantity-?
                WHERE item_id=? AND site_code=? AND warehouse_code=?
                  AND location_code=? AND lot_number=?
                """,
                (
                    release_quantity,
                    order["item_id"],
                    order["site_code"],
                    warehouse_code,
                    location_code,
                    order["lot_number"] or "",
                ),
            )
            connection.execute(
                """
                UPDATE quality_orders
                SET status='Closed',closed_date=?,disposition=?
                WHERE id=?
                """,
                (SNAPSHOT_DATE, disposition, quality_order_id),
            )
            self._audit(
                connection,
                self.actor,
                "release_hold",
                "inventory",
                quality_order_id,
                {
                    "released_quantity": _qty(release_quantity),
                    "disposition": disposition,
                    "quality_hold_before": _qty(
                        balance["quality_hold_quantity"]
                    ),
                    "quality_hold_after": _qty(
                        balance["quality_hold_quantity"] - release_quantity
                    ),
                },
            )
            return {
                "quality_order_id": quality_order_id,
                "status": "Closed",
                "released_quantity": _qty(release_quantity),
                "disposition": disposition,
            }
