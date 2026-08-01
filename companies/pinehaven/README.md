# Pinehaven Motion Systems

Pinehaven Motion Systems is a multi-plant discrete manufacturer represented at a June 30, 2026 pre-close snapshot. The company package combines a finance and manufacturing file system, a controlled ERP application, and 100 senior corporate-finance tasks.

## What the tasks test

The task set covers:

- manufacturing close, production variance, WIP, subledger control, and pre-close P&L;
- inventory value, aging, reserve exposure, turns, transfers, and working-capital decisions;
- S&OP demand, backlog, capacity, production planning, and forecast scenarios;
- customer and product profitability, pricing, margin sensitivities, and commercial finance;
- purchase-price variance, supplier performance, commitments, invoice holds, and procurement controls;
- scrap, quality, cost of poor quality, maintenance, standard cost, and operations finance;
- capital projects, NPV, portfolio selection, fixed assets, treasury, liquidity, FX, hedging, tax, and corporate development;
- executive spreadsheets, documents, presentations, and concise structured answers; and
- eight tightly scoped ERP write tasks with explicit authorization and audit requirements.

The 100-task bank contains 52 console assignments, 24 spreadsheet assignments, 10 document assignments, 6 presentation assignments, and 8 ERP workflow assignments. Difficulty labels comprise 78 advanced and 22 expert tasks. The complete bank is indexed in [tasks/index.json](tasks/index.json), and [tasks/hud-taskset.json](tasks/hud-taskset.json) includes all 100 tasks.

## Company sources and tools

`source-files/` contains 164 files: 62 spreadsheets, 30 emails, 20 PDFs, 20 documents, 20 CSV extracts, 10 presentations, and 2 structured text files. Files are arranged by finance, manufacturing, policy, communication, and reporting function.

The `pinehaven_manufacturing_erp` MCP provides controlled reporting across general ledger, production, inventory, sales, procurement, payroll, quality, fixed assets, treasury, tax, and corporate-development records. Write operations are limited to the eight tasks that explicitly authorize them.

## Package layout

```text
company.json              Structured company, task, source, and runtime metadata
tasks/index.json          Compact list of all 100 tasks
tasks/hud-taskset.json    HUD submission taskset with all 100 tasks
tasks/task_###.json       Prompt, workflow, output mode, and evaluator pointer
source-files/             Agent-visible company files
runtime/                  HUD environment, task registration, and ERP application
evaluator/                Private scoring code and expected-result records
data/                     Verified ERP transport, manifests, and source controls
Dockerfile                Container entrypoint
```

The evaluator and `data/` directories are included for reviewer reproducibility but are never mounted into the agent shell. On reset, only `source-files/` is copied to `/workspace`; ERP state is reconstructed into a private mutable database.

## Build

Use this directory as the container context:

```bash
docker build -t pinehaven-finance .
```

The image starts `pinehaven-manufacturing-finance-v1` through HUD on port 8765.
