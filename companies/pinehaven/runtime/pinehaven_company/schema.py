from __future__ import annotations

import sqlite3


DDL = """
PRAGMA foreign_keys=ON;

CREATE TABLE company (
    id TEXT PRIMARY KEY,
    legal_name TEXT NOT NULL,
    trade_name TEXT NOT NULL,
    founded_date TEXT NOT NULL,
    headquarters TEXT NOT NULL,
    currency TEXT NOT NULL,
    fiscal_year_end TEXT NOT NULL,
    reporting_cutoff TEXT NOT NULL,
    closed_through TEXT NOT NULL,
    accounting_basis TEXT NOT NULL,
    inventory_method TEXT NOT NULL
);

CREATE TABLE sites (
    code TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    city TEXT NOT NULL,
    state TEXT NOT NULL,
    site_type TEXT NOT NULL,
    profit_center TEXT NOT NULL,
    active INTEGER NOT NULL CHECK(active IN (0,1))
);

CREATE TABLE departments (
    code TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    site_code TEXT,
    cost_center TEXT NOT NULL,
    function TEXT NOT NULL,
    FOREIGN KEY(site_code) REFERENCES sites(code)
);

CREATE TABLE fiscal_periods (
    period TEXT PRIMARY KEY,
    start_date TEXT NOT NULL,
    end_date TEXT NOT NULL,
    status TEXT NOT NULL CHECK(status IN ('Open','Closed','Future'))
);

CREATE TABLE accounts (
    code TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    account_type TEXT NOT NULL,
    normal_balance TEXT NOT NULL CHECK(normal_balance IN ('Debit','Credit')),
    financial_statement TEXT NOT NULL,
    is_cash INTEGER NOT NULL DEFAULT 0 CHECK(is_cash IN (0,1)),
    is_control INTEGER NOT NULL DEFAULT 0 CHECK(is_control IN (0,1))
);

CREATE TABLE customers (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    end_market TEXT NOT NULL,
    region TEXT NOT NULL,
    payment_terms_days INTEGER NOT NULL,
    credit_limit REAL NOT NULL,
    active INTEGER NOT NULL CHECK(active IN (0,1))
);

CREATE TABLE vendors (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    category TEXT NOT NULL,
    country TEXT NOT NULL,
    payment_terms_days INTEGER NOT NULL,
    preferred INTEGER NOT NULL CHECK(preferred IN (0,1)),
    active INTEGER NOT NULL CHECK(active IN (0,1))
);

CREATE TABLE employees (
    id TEXT PRIMARY KEY,
    full_name TEXT NOT NULL,
    title TEXT NOT NULL,
    department_code TEXT NOT NULL,
    site_code TEXT NOT NULL,
    hire_date TEXT NOT NULL,
    annual_salary REAL,
    hourly_rate REAL,
    active INTEGER NOT NULL CHECK(active IN (0,1)),
    email TEXT NOT NULL UNIQUE,
    FOREIGN KEY(department_code) REFERENCES departments(code),
    FOREIGN KEY(site_code) REFERENCES sites(code)
);

CREATE TABLE warehouses (
    code TEXT PRIMARY KEY,
    site_code TEXT NOT NULL,
    name TEXT NOT NULL,
    warehouse_type TEXT NOT NULL,
    FOREIGN KEY(site_code) REFERENCES sites(code)
);

CREATE TABLE locations (
    code TEXT PRIMARY KEY,
    warehouse_code TEXT NOT NULL,
    name TEXT NOT NULL,
    location_type TEXT NOT NULL,
    quality_blocked INTEGER NOT NULL DEFAULT 0 CHECK(quality_blocked IN (0,1)),
    FOREIGN KEY(warehouse_code) REFERENCES warehouses(code)
);

CREATE TABLE items (
    id TEXT PRIMARY KEY,
    description TEXT NOT NULL,
    item_type TEXT NOT NULL,
    product_family TEXT,
    inventory_uom TEXT NOT NULL,
    make_buy TEXT NOT NULL CHECK(make_buy IN ('Make','Buy','Either')),
    lot_controlled INTEGER NOT NULL CHECK(lot_controlled IN (0,1)),
    serial_controlled INTEGER NOT NULL CHECK(serial_controlled IN (0,1)),
    lifecycle_status TEXT NOT NULL,
    standard_order_quantity REAL NOT NULL,
    primary_vendor_id TEXT,
    commodity TEXT,
    FOREIGN KEY(primary_vendor_id) REFERENCES vendors(id)
);

CREATE TABLE item_sites (
    item_id TEXT NOT NULL,
    site_code TEXT NOT NULL,
    planner_code TEXT NOT NULL,
    buyer_code TEXT,
    lead_time_days INTEGER NOT NULL,
    safety_stock REAL NOT NULL,
    reorder_point REAL NOT NULL,
    min_order_qty REAL NOT NULL,
    order_multiple REAL NOT NULL,
    standard_material_cost REAL NOT NULL,
    standard_labor_cost REAL NOT NULL,
    standard_variable_overhead REAL NOT NULL,
    standard_fixed_overhead REAL NOT NULL,
    standard_outside_processing REAL NOT NULL,
    standard_cost_effective_date TEXT NOT NULL,
    PRIMARY KEY(item_id,site_code),
    FOREIGN KEY(item_id) REFERENCES items(id),
    FOREIGN KEY(site_code) REFERENCES sites(code)
);

CREATE TABLE work_centers (
    code TEXT PRIMARY KEY,
    site_code TEXT NOT NULL,
    name TEXT NOT NULL,
    cost_center TEXT NOT NULL,
    capacity_hours_per_day REAL NOT NULL,
    efficiency_target REAL NOT NULL,
    labor_rate REAL NOT NULL,
    machine_rate REAL NOT NULL,
    variable_overhead_rate REAL NOT NULL,
    fixed_overhead_rate REAL NOT NULL,
    active INTEGER NOT NULL CHECK(active IN (0,1)),
    FOREIGN KEY(site_code) REFERENCES sites(code)
);

CREATE TABLE work_center_calendar (
    work_center_code TEXT NOT NULL,
    work_date TEXT NOT NULL,
    available_hours REAL NOT NULL,
    planned_downtime_hours REAL NOT NULL,
    PRIMARY KEY(work_center_code,work_date),
    FOREIGN KEY(work_center_code) REFERENCES work_centers(code)
);

CREATE TABLE bom_headers (
    id INTEGER PRIMARY KEY,
    item_id TEXT NOT NULL,
    site_code TEXT NOT NULL,
    version TEXT NOT NULL,
    effective_from TEXT NOT NULL,
    effective_to TEXT,
    status TEXT NOT NULL,
    approved_by TEXT,
    approved_at TEXT,
    UNIQUE(item_id,site_code,version),
    FOREIGN KEY(item_id) REFERENCES items(id),
    FOREIGN KEY(site_code) REFERENCES sites(code)
);

CREATE TABLE bom_lines (
    id INTEGER PRIMARY KEY,
    bom_id INTEGER NOT NULL,
    line_number INTEGER NOT NULL,
    component_item_id TEXT NOT NULL,
    quantity_per REAL NOT NULL,
    scrap_factor REAL NOT NULL,
    issue_operation INTEGER NOT NULL,
    cost_relevant INTEGER NOT NULL CHECK(cost_relevant IN (0,1)),
    UNIQUE(bom_id,line_number),
    FOREIGN KEY(bom_id) REFERENCES bom_headers(id),
    FOREIGN KEY(component_item_id) REFERENCES items(id)
);

CREATE TABLE routing_headers (
    id INTEGER PRIMARY KEY,
    item_id TEXT NOT NULL,
    site_code TEXT NOT NULL,
    version TEXT NOT NULL,
    effective_from TEXT NOT NULL,
    effective_to TEXT,
    status TEXT NOT NULL,
    approved_by TEXT,
    approved_at TEXT,
    UNIQUE(item_id,site_code,version),
    FOREIGN KEY(item_id) REFERENCES items(id),
    FOREIGN KEY(site_code) REFERENCES sites(code)
);

CREATE TABLE routing_operations (
    id INTEGER PRIMARY KEY,
    routing_id INTEGER NOT NULL,
    operation_number INTEGER NOT NULL,
    work_center_code TEXT NOT NULL,
    setup_hours REAL NOT NULL,
    run_hours_per_unit REAL NOT NULL,
    queue_hours REAL NOT NULL,
    outside_processing INTEGER NOT NULL CHECK(outside_processing IN (0,1)),
    yield_percent REAL NOT NULL,
    UNIQUE(routing_id,operation_number),
    FOREIGN KEY(routing_id) REFERENCES routing_headers(id),
    FOREIGN KEY(work_center_code) REFERENCES work_centers(code)
);

CREATE TABLE engineering_changes (
    id TEXT PRIMARY KEY,
    item_id TEXT NOT NULL,
    change_type TEXT NOT NULL,
    description TEXT NOT NULL,
    requested_date TEXT NOT NULL,
    effective_date TEXT,
    status TEXT NOT NULL,
    financial_impact REAL NOT NULL,
    owner_employee_id TEXT NOT NULL,
    FOREIGN KEY(item_id) REFERENCES items(id),
    FOREIGN KEY(owner_employee_id) REFERENCES employees(id)
);

CREATE TABLE sales_orders (
    id INTEGER PRIMARY KEY,
    order_number TEXT NOT NULL UNIQUE,
    customer_id TEXT NOT NULL,
    order_date TEXT NOT NULL,
    requested_ship_date TEXT NOT NULL,
    promised_ship_date TEXT NOT NULL,
    status TEXT NOT NULL,
    currency TEXT NOT NULL,
    sales_region TEXT NOT NULL,
    customer_po TEXT,
    FOREIGN KEY(customer_id) REFERENCES customers(id)
);

CREATE TABLE sales_order_lines (
    id INTEGER PRIMARY KEY,
    sales_order_id INTEGER NOT NULL,
    line_number INTEGER NOT NULL,
    item_id TEXT NOT NULL,
    site_code TEXT NOT NULL,
    order_quantity REAL NOT NULL,
    shipped_quantity REAL NOT NULL,
    unit_price REAL NOT NULL,
    discount_percent REAL NOT NULL,
    standard_unit_cost REAL NOT NULL,
    status TEXT NOT NULL,
    UNIQUE(sales_order_id,line_number),
    FOREIGN KEY(sales_order_id) REFERENCES sales_orders(id),
    FOREIGN KEY(item_id) REFERENCES items(id),
    FOREIGN KEY(site_code) REFERENCES sites(code)
);

CREATE TABLE shipments (
    id INTEGER PRIMARY KEY,
    shipment_number TEXT NOT NULL UNIQUE,
    sales_order_id INTEGER NOT NULL,
    ship_date TEXT NOT NULL,
    status TEXT NOT NULL,
    freight_amount REAL NOT NULL,
    carrier TEXT NOT NULL,
    FOREIGN KEY(sales_order_id) REFERENCES sales_orders(id)
);

CREATE TABLE shipment_lines (
    id INTEGER PRIMARY KEY,
    shipment_id INTEGER NOT NULL,
    sales_order_line_id INTEGER NOT NULL,
    item_id TEXT NOT NULL,
    site_code TEXT NOT NULL,
    quantity REAL NOT NULL,
    unit_standard_cost REAL NOT NULL,
    FOREIGN KEY(shipment_id) REFERENCES shipments(id),
    FOREIGN KEY(sales_order_line_id) REFERENCES sales_order_lines(id),
    FOREIGN KEY(item_id) REFERENCES items(id),
    FOREIGN KEY(site_code) REFERENCES sites(code)
);

CREATE TABLE customer_invoices (
    id INTEGER PRIMARY KEY,
    invoice_number TEXT NOT NULL UNIQUE,
    customer_id TEXT NOT NULL,
    sales_order_id INTEGER,
    invoice_date TEXT NOT NULL,
    due_date TEXT NOT NULL,
    status TEXT NOT NULL,
    amount REAL NOT NULL,
    open_amount REAL NOT NULL,
    journal_id INTEGER,
    FOREIGN KEY(customer_id) REFERENCES customers(id),
    FOREIGN KEY(sales_order_id) REFERENCES sales_orders(id)
);

CREATE TABLE customer_invoice_lines (
    id INTEGER PRIMARY KEY,
    invoice_id INTEGER NOT NULL,
    line_number INTEGER NOT NULL,
    item_id TEXT,
    quantity REAL NOT NULL,
    unit_price REAL NOT NULL,
    amount REAL NOT NULL,
    revenue_account TEXT NOT NULL,
    site_code TEXT,
    product_family TEXT,
    FOREIGN KEY(invoice_id) REFERENCES customer_invoices(id),
    FOREIGN KEY(item_id) REFERENCES items(id)
);

CREATE TABLE purchase_orders (
    id INTEGER PRIMARY KEY,
    po_number TEXT NOT NULL UNIQUE,
    vendor_id TEXT NOT NULL,
    site_code TEXT NOT NULL,
    order_date TEXT NOT NULL,
    expected_date TEXT NOT NULL,
    buyer_employee_id TEXT NOT NULL,
    status TEXT NOT NULL,
    currency TEXT NOT NULL,
    FOREIGN KEY(vendor_id) REFERENCES vendors(id),
    FOREIGN KEY(site_code) REFERENCES sites(code),
    FOREIGN KEY(buyer_employee_id) REFERENCES employees(id)
);

CREATE TABLE purchase_order_lines (
    id INTEGER PRIMARY KEY,
    po_id INTEGER NOT NULL,
    line_number INTEGER NOT NULL,
    item_id TEXT NOT NULL,
    order_quantity REAL NOT NULL,
    received_quantity REAL NOT NULL,
    invoiced_quantity REAL NOT NULL,
    unit_price REAL NOT NULL,
    standard_unit_cost REAL NOT NULL,
    status TEXT NOT NULL,
    UNIQUE(po_id,line_number),
    FOREIGN KEY(po_id) REFERENCES purchase_orders(id),
    FOREIGN KEY(item_id) REFERENCES items(id)
);

CREATE TABLE receipts (
    id INTEGER PRIMARY KEY,
    receipt_number TEXT NOT NULL UNIQUE,
    po_id INTEGER NOT NULL,
    receipt_date TEXT NOT NULL,
    site_code TEXT NOT NULL,
    status TEXT NOT NULL,
    FOREIGN KEY(po_id) REFERENCES purchase_orders(id),
    FOREIGN KEY(site_code) REFERENCES sites(code)
);

CREATE TABLE receipt_lines (
    id INTEGER PRIMARY KEY,
    receipt_id INTEGER NOT NULL,
    po_line_id INTEGER NOT NULL,
    item_id TEXT NOT NULL,
    quantity REAL NOT NULL,
    unit_price REAL NOT NULL,
    standard_unit_cost REAL NOT NULL,
    lot_number TEXT,
    quality_status TEXT NOT NULL,
    FOREIGN KEY(receipt_id) REFERENCES receipts(id),
    FOREIGN KEY(po_line_id) REFERENCES purchase_order_lines(id),
    FOREIGN KEY(item_id) REFERENCES items(id)
);

CREATE TABLE vendor_invoices (
    id INTEGER PRIMARY KEY,
    invoice_number TEXT NOT NULL,
    vendor_id TEXT NOT NULL,
    po_id INTEGER,
    invoice_date TEXT NOT NULL,
    due_date TEXT NOT NULL,
    status TEXT NOT NULL,
    amount REAL NOT NULL,
    open_amount REAL NOT NULL,
    match_status TEXT NOT NULL,
    journal_id INTEGER,
    UNIQUE(vendor_id,invoice_number),
    FOREIGN KEY(vendor_id) REFERENCES vendors(id),
    FOREIGN KEY(po_id) REFERENCES purchase_orders(id)
);

CREATE TABLE vendor_invoice_lines (
    id INTEGER PRIMARY KEY,
    vendor_invoice_id INTEGER NOT NULL,
    po_line_id INTEGER,
    item_id TEXT,
    quantity REAL NOT NULL,
    unit_price REAL NOT NULL,
    amount REAL NOT NULL,
    account_code TEXT NOT NULL,
    site_code TEXT,
    FOREIGN KEY(vendor_invoice_id) REFERENCES vendor_invoices(id),
    FOREIGN KEY(po_line_id) REFERENCES purchase_order_lines(id),
    FOREIGN KEY(item_id) REFERENCES items(id)
);

CREATE TABLE production_orders (
    id INTEGER PRIMARY KEY,
    order_number TEXT NOT NULL UNIQUE,
    item_id TEXT NOT NULL,
    site_code TEXT NOT NULL,
    bom_id INTEGER NOT NULL,
    routing_id INTEGER NOT NULL,
    source_type TEXT NOT NULL,
    source_reference TEXT,
    created_date TEXT NOT NULL,
    scheduled_start TEXT NOT NULL,
    scheduled_finish TEXT NOT NULL,
    actual_start TEXT,
    actual_finish TEXT,
    order_quantity REAL NOT NULL,
    completed_quantity REAL NOT NULL,
    scrapped_quantity REAL NOT NULL,
    status TEXT NOT NULL,
    standard_unit_cost REAL NOT NULL,
    actual_material_cost REAL NOT NULL,
    actual_labor_cost REAL NOT NULL,
    actual_variable_overhead REAL NOT NULL,
    actual_fixed_overhead REAL NOT NULL,
    actual_outside_processing REAL NOT NULL,
    FOREIGN KEY(item_id) REFERENCES items(id),
    FOREIGN KEY(site_code) REFERENCES sites(code),
    FOREIGN KEY(bom_id) REFERENCES bom_headers(id),
    FOREIGN KEY(routing_id) REFERENCES routing_headers(id)
);

CREATE TABLE production_order_materials (
    id INTEGER PRIMARY KEY,
    production_order_id INTEGER NOT NULL,
    line_number INTEGER NOT NULL,
    component_item_id TEXT NOT NULL,
    planned_quantity REAL NOT NULL,
    issued_quantity REAL NOT NULL,
    standard_unit_cost REAL NOT NULL,
    actual_unit_cost REAL NOT NULL,
    issue_date TEXT,
    substitution INTEGER NOT NULL CHECK(substitution IN (0,1)),
    FOREIGN KEY(production_order_id) REFERENCES production_orders(id),
    FOREIGN KEY(component_item_id) REFERENCES items(id)
);

CREATE TABLE production_order_operations (
    id INTEGER PRIMARY KEY,
    production_order_id INTEGER NOT NULL,
    operation_number INTEGER NOT NULL,
    work_center_code TEXT NOT NULL,
    planned_setup_hours REAL NOT NULL,
    planned_run_hours REAL NOT NULL,
    actual_setup_hours REAL NOT NULL,
    actual_run_hours REAL NOT NULL,
    labor_rate_standard REAL NOT NULL,
    labor_rate_actual REAL NOT NULL,
    machine_rate_standard REAL NOT NULL,
    completed_date TEXT,
    status TEXT NOT NULL,
    FOREIGN KEY(production_order_id) REFERENCES production_orders(id),
    FOREIGN KEY(work_center_code) REFERENCES work_centers(code)
);

CREATE TABLE production_variances (
    production_order_id INTEGER PRIMARY KEY,
    material_price_variance REAL NOT NULL,
    material_usage_variance REAL NOT NULL,
    labor_rate_variance REAL NOT NULL,
    labor_efficiency_variance REAL NOT NULL,
    variable_overhead_variance REAL NOT NULL,
    fixed_overhead_volume_variance REAL NOT NULL,
    scrap_variance REAL NOT NULL,
    substitution_variance REAL NOT NULL,
    total_variance REAL NOT NULL,
    settled_date TEXT,
    FOREIGN KEY(production_order_id) REFERENCES production_orders(id)
);

CREATE TABLE inventory_transactions (
    id INTEGER PRIMARY KEY,
    transaction_date TEXT NOT NULL,
    transaction_type TEXT NOT NULL,
    item_id TEXT NOT NULL,
    site_code TEXT NOT NULL,
    warehouse_code TEXT NOT NULL,
    location_code TEXT NOT NULL,
    lot_number TEXT,
    quantity REAL NOT NULL,
    unit_cost REAL NOT NULL,
    reference_type TEXT NOT NULL,
    reference_id TEXT NOT NULL,
    quality_status TEXT NOT NULL,
    journal_id INTEGER,
    FOREIGN KEY(item_id) REFERENCES items(id),
    FOREIGN KEY(site_code) REFERENCES sites(code),
    FOREIGN KEY(warehouse_code) REFERENCES warehouses(code),
    FOREIGN KEY(location_code) REFERENCES locations(code)
);

CREATE TABLE inventory_balances (
    item_id TEXT NOT NULL,
    site_code TEXT NOT NULL,
    warehouse_code TEXT NOT NULL,
    location_code TEXT NOT NULL,
    lot_number TEXT NOT NULL DEFAULT '',
    on_hand_quantity REAL NOT NULL,
    reserved_quantity REAL NOT NULL,
    quality_hold_quantity REAL NOT NULL,
    standard_unit_cost REAL NOT NULL,
    last_movement_date TEXT NOT NULL,
    PRIMARY KEY(item_id,site_code,warehouse_code,location_code,lot_number),
    FOREIGN KEY(item_id) REFERENCES items(id),
    FOREIGN KEY(site_code) REFERENCES sites(code),
    FOREIGN KEY(warehouse_code) REFERENCES warehouses(code),
    FOREIGN KEY(location_code) REFERENCES locations(code)
);

CREATE TABLE quality_orders (
    id TEXT PRIMARY KEY,
    reference_type TEXT NOT NULL,
    reference_id TEXT NOT NULL,
    item_id TEXT NOT NULL,
    site_code TEXT NOT NULL,
    lot_number TEXT,
    opened_date TEXT NOT NULL,
    closed_date TEXT,
    status TEXT NOT NULL,
    disposition TEXT,
    quantity_inspected REAL NOT NULL,
    quantity_failed REAL NOT NULL,
    estimated_financial_exposure REAL NOT NULL,
    FOREIGN KEY(item_id) REFERENCES items(id),
    FOREIGN KEY(site_code) REFERENCES sites(code)
);

CREATE TABLE quality_tests (
    id INTEGER PRIMARY KEY,
    quality_order_id TEXT NOT NULL,
    test_name TEXT NOT NULL,
    specification TEXT NOT NULL,
    result_value TEXT,
    result_status TEXT NOT NULL,
    tested_at TEXT,
    FOREIGN KEY(quality_order_id) REFERENCES quality_orders(id)
);

CREATE TABLE nonconformances (
    id TEXT PRIMARY KEY,
    quality_order_id TEXT,
    opened_date TEXT NOT NULL,
    site_code TEXT NOT NULL,
    item_id TEXT,
    root_cause_category TEXT,
    status TEXT NOT NULL,
    disposition TEXT,
    estimated_cost REAL NOT NULL,
    owner_employee_id TEXT NOT NULL,
    FOREIGN KEY(quality_order_id) REFERENCES quality_orders(id),
    FOREIGN KEY(site_code) REFERENCES sites(code),
    FOREIGN KEY(item_id) REFERENCES items(id),
    FOREIGN KEY(owner_employee_id) REFERENCES employees(id)
);

CREATE TABLE maintenance_assets (
    id TEXT PRIMARY KEY,
    site_code TEXT NOT NULL,
    work_center_code TEXT,
    description TEXT NOT NULL,
    asset_class TEXT NOT NULL,
    in_service_date TEXT NOT NULL,
    criticality TEXT NOT NULL,
    replacement_value REAL NOT NULL,
    status TEXT NOT NULL,
    FOREIGN KEY(site_code) REFERENCES sites(code),
    FOREIGN KEY(work_center_code) REFERENCES work_centers(code)
);

CREATE TABLE maintenance_work_orders (
    id TEXT PRIMARY KEY,
    asset_id TEXT NOT NULL,
    request_date TEXT NOT NULL,
    scheduled_date TEXT,
    completed_date TEXT,
    work_type TEXT NOT NULL,
    priority TEXT NOT NULL,
    status TEXT NOT NULL,
    downtime_hours REAL NOT NULL,
    labor_cost REAL NOT NULL,
    material_cost REAL NOT NULL,
    outside_service_cost REAL NOT NULL,
    description TEXT NOT NULL,
    FOREIGN KEY(asset_id) REFERENCES maintenance_assets(id)
);

CREATE TABLE demand_forecast (
    id INTEGER PRIMARY KEY,
    forecast_version TEXT NOT NULL,
    forecast_month TEXT NOT NULL,
    item_id TEXT NOT NULL,
    site_code TEXT NOT NULL,
    customer_id TEXT,
    quantity REAL NOT NULL,
    unit_price REAL NOT NULL,
    status TEXT NOT NULL,
    FOREIGN KEY(item_id) REFERENCES items(id),
    FOREIGN KEY(site_code) REFERENCES sites(code),
    FOREIGN KEY(customer_id) REFERENCES customers(id)
);

CREATE TABLE planned_orders (
    id TEXT PRIMARY KEY,
    plan_version TEXT NOT NULL,
    order_type TEXT NOT NULL,
    item_id TEXT NOT NULL,
    site_code TEXT NOT NULL,
    quantity REAL NOT NULL,
    required_date TEXT NOT NULL,
    release_date TEXT NOT NULL,
    action_message TEXT,
    priority INTEGER NOT NULL,
    source_demand_type TEXT NOT NULL,
    source_demand_reference TEXT,
    status TEXT NOT NULL,
    FOREIGN KEY(item_id) REFERENCES items(id),
    FOREIGN KEY(site_code) REFERENCES sites(code)
);

CREATE TABLE payroll_runs (
    id INTEGER PRIMARY KEY,
    period_start TEXT NOT NULL,
    period_end TEXT NOT NULL,
    check_date TEXT NOT NULL,
    status TEXT NOT NULL
);

CREATE TABLE payroll_lines (
    id INTEGER PRIMARY KEY,
    payroll_run_id INTEGER NOT NULL,
    employee_id TEXT NOT NULL,
    site_code TEXT NOT NULL,
    department_code TEXT NOT NULL,
    regular_hours REAL NOT NULL,
    overtime_hours REAL NOT NULL,
    gross_pay REAL NOT NULL,
    employer_taxes REAL NOT NULL,
    benefits REAL NOT NULL,
    FOREIGN KEY(payroll_run_id) REFERENCES payroll_runs(id),
    FOREIGN KEY(employee_id) REFERENCES employees(id),
    FOREIGN KEY(site_code) REFERENCES sites(code),
    FOREIGN KEY(department_code) REFERENCES departments(code)
);

CREATE TABLE fixed_assets (
    id TEXT PRIMARY KEY,
    description TEXT NOT NULL,
    site_code TEXT NOT NULL,
    asset_class TEXT NOT NULL,
    in_service_date TEXT NOT NULL,
    gross_cost REAL NOT NULL,
    accumulated_depreciation REAL NOT NULL,
    useful_life_months INTEGER NOT NULL,
    salvage_value REAL NOT NULL,
    account_code TEXT NOT NULL,
    accumulated_depreciation_account TEXT NOT NULL,
    depreciation_expense_account TEXT NOT NULL,
    status TEXT NOT NULL,
    FOREIGN KEY(site_code) REFERENCES sites(code)
);

CREATE TABLE debt_instruments (
    id TEXT PRIMARY KEY,
    description TEXT NOT NULL,
    instrument_type TEXT NOT NULL,
    lender TEXT NOT NULL,
    original_principal REAL NOT NULL,
    outstanding_principal REAL NOT NULL,
    commitment REAL NOT NULL,
    interest_rate REAL NOT NULL,
    maturity_date TEXT NOT NULL,
    current_portion REAL NOT NULL,
    status TEXT NOT NULL
);

CREATE TABLE journal_headers (
    id INTEGER PRIMARY KEY,
    posting_date TEXT NOT NULL,
    period TEXT NOT NULL,
    source TEXT NOT NULL,
    reference TEXT NOT NULL,
    memo TEXT NOT NULL,
    status TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE TABLE journal_lines (
    id INTEGER PRIMARY KEY,
    journal_id INTEGER NOT NULL,
    account_code TEXT NOT NULL,
    department_code TEXT,
    site_code TEXT,
    item_id TEXT,
    production_order_id INTEGER,
    customer_id TEXT,
    vendor_id TEXT,
    debit REAL NOT NULL,
    credit REAL NOT NULL,
    line_memo TEXT NOT NULL,
    FOREIGN KEY(journal_id) REFERENCES journal_headers(id),
    FOREIGN KEY(account_code) REFERENCES accounts(code)
);

CREATE TABLE audit_events (
    id INTEGER PRIMARY KEY,
    event_at TEXT NOT NULL,
    actor TEXT NOT NULL,
    action TEXT NOT NULL,
    entity TEXT NOT NULL,
    identifier TEXT NOT NULL,
    details_json TEXT NOT NULL
);

CREATE INDEX idx_item_sites_site ON item_sites(site_code,item_id);
CREATE INDEX idx_bom_item_site ON bom_headers(item_id,site_code,effective_from);
CREATE INDEX idx_routing_item_site ON routing_headers(item_id,site_code,effective_from);
CREATE INDEX idx_sales_order_date ON sales_orders(order_date,status);
CREATE INDEX idx_sales_lines_item ON sales_order_lines(item_id,site_code);
CREATE INDEX idx_po_date ON purchase_orders(order_date,status);
CREATE INDEX idx_po_lines_item ON purchase_order_lines(item_id);
CREATE INDEX idx_prod_date ON production_orders(actual_finish,status,site_code);
CREATE INDEX idx_prod_item ON production_orders(item_id,site_code);
CREATE INDEX idx_prod_material_order ON production_order_materials(production_order_id);
CREATE INDEX idx_prod_operation_order ON production_order_operations(production_order_id);
CREATE INDEX idx_inventory_date ON inventory_transactions(transaction_date,item_id,site_code);
CREATE INDEX idx_inventory_balance_item ON inventory_balances(item_id,site_code);
CREATE INDEX idx_quality_item ON quality_orders(item_id,site_code,status);
CREATE INDEX idx_maintenance_date ON maintenance_work_orders(request_date,status);
CREATE INDEX idx_forecast_version ON demand_forecast(forecast_version,forecast_month);
CREATE INDEX idx_planned_version ON planned_orders(plan_version,required_date);
CREATE INDEX idx_journal_date ON journal_headers(posting_date,status);
CREATE INDEX idx_journal_account ON journal_lines(account_code,site_code,item_id);
"""


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(DDL)

