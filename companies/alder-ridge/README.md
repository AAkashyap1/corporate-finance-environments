# Alder Ridge Mechanical

Alder Ridge Mechanical is a specialty mechanical contractor represented at a June 30, 2026 pre-close snapshot. The company package combines a realistic finance file system with a contractor-accounting application and 100 senior-finance tasks.

## What the tasks test

The task set covers:

- project accounting, percentage-of-completion WIP, change-order evidence, and proposed close entries;
- 13-week liquidity, borrowing-base availability, cash controls, debt, covenant, and surety analysis;
- actual-versus-plan reporting, branch economics, labor productivity, service KPIs, backlog, and capital planning;
- acquisition valuation, quality of earnings, financing capacity, integration planning, and portfolio decisions;
- tax provision, deferred tax, transaction tax, and controllership review;
- investor-relations, lender, board, and CFO decision materials; and
- spreadsheet, presentation, and document creation or editing in addition to structured console answers.

The 100 tasks contain 69 console assignments, 23 spreadsheet assignments, 5 document assignments, and 3 presentation assignments. Difficulty labels comprise 25 advanced tasks and 75 long-horizon advanced or expert tasks. The complete index is [tasks/index.json](tasks/index.json), and [tasks/hud-taskset.json](tasks/hud-taskset.json) includes all 100 tasks.

## Company sources and tools

`source-files/` contains 122 files: 52 spreadsheets, 20 emails, 19 PDFs, 15 documents, 9 presentations, 6 CSV extracts, and 1 text note. The directory preserves the company’s natural `Requests/`, `Shared/Finance/`, and `Shared/Operations/` organization so reviewers can trace each task through the same materials available to the agent.

The `contractor_accounting` MCP exposes controlled accounting and operating records. It is read-only for analytical assignments. File-editing tasks write only to the reset workspace.

## Package layout

```text
company.json              Structured company, task, source, and runtime metadata
tasks/index.json          Compact list of all 100 tasks
tasks/hud-taskset.json    HUD submission taskset with all 100 tasks
tasks/task_###.json       Prompt, workflow, output mode, and evaluator pointer
source-files/             Agent-visible company files
runtime/                  HUD environment, task registration, and accounting MCP
evaluator/                Private scoring code and expected-result records
data/                     Immutable accounting seed and source controls
Dockerfile                Container entrypoint
```

The evaluator and `data/` directories are included for reviewer reproducibility but are never mounted into the agent shell. On reset, only `source-files/` is copied to `/workspace`; the accounting database is copied to private mutable state.

## Build

Use this directory as the container context:

```bash
docker build -t alder-ridge-finance .
```

The image starts `alder-ridge-corporate-finance` through HUD on port 8765.
