from __future__ import annotations

from env import TASK_TEMPLATES
from task_catalog import TASKS


tasks = []
for spec in TASKS:
    task = TASK_TEMPLATES[spec.task_id]()
    task.slug = spec.slug
    task.columns = {
        "company": "alder-ridge-mechanical",
        "workflow": spec.workflow,
        "level": "senior-finance-analyst",
        "output_mode": spec.output_mode,
        "difficulty": spec.difficulty,
        "snapshot": "2026-06-30-pre-close",
        "grading": "weighted-atomic-hybrid-v3",
    }
    tasks.append(task)

# HUD discovers Task objects from globals and from ``tasks``. Avoid exposing the
# loop variable as a duplicate task object.
del task, spec
