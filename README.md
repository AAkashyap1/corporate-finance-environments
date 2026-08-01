# Corporate Finance Environments

## Overview

This repository contains two simulated companies designed for evaluating agents on corporate finance work:

- **Alder Ridge Mechanical** is a specialty mechanical contractor with project-based accounting, percentage-of-completion revenue recognition, WIP, change orders, service operations, treasury, and contractor-accounting records.
- **Pinehaven Motion Systems** is a multi-plant industrial manufacturer with production, inventory, procurement, sales, quality, fixed-asset, treasury, and general-ledger activity maintained through a manufacturing ERP.

The tasks require agents to use company accounting and ERP tools along with spreadsheets, documents, presentations, PDFs, emails, policies, and operating extracts. They test controllership, FP&A, treasury, working capital, manufacturing and project accounting, tax, corporate development, capital allocation, and lender and board reporting. Assignments range from reconciliations and analyses to creating/editing full finance deliverables.

## Rights, provenance, and privacy

All tasks and source files were authored by domain experts and informed by Sureform's partner engagements with a specialty mechanical contractor and a multi-plant industrial manufacturer. The represented work includes monthly close, budgeting and forecasting, project WIP and revenue recognition, manufacturing cost and inventory accounting, treasury and liquidity, working capital, procurement, capital planning, tax, management reporting, and board and lender deliverables.

Alder Ridge Mechanical, Pinehaven Motion Systems, their personnel, counterparties, communications, transactions, and records are simulated. The package does not contain identifiable client, partner, employee, insurance, banking, or customer data; email namespaces and business identifiers are non-production examples.

The included code, task definitions, evaluator logic, and company records are controlled by the repository owner for commercial licensing. Distribution and use are governed by [LICENSE](LICENSE) and the applicable order form. The materials are not accounting, investment, tax, legal, or operational advice.

## Container entrypoints

Each company is self-contained. Use the company directory as the build context:

```bash
docker build -t alder-ridge-finance .
```

Run that command from `companies/alder-ridge`; substitute `companies/pinehaven` for the manufacturing package. `Dockerfile` starts the HUD server on port 8765, and each company directory can be passed directly to `hud deploy`. Deployment credentials must be supplied by the target platform and must never be placed in the agent workspace.
