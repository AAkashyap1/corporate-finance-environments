"""HUD task discovery for all 100 Pinehaven episodes."""

from __future__ import annotations

from env import TASK_TEMPLATES
from task_catalog import TASKS


tasks = []
for spec in TASKS:
    task = TASK_TEMPLATES[spec.task_id]()
    task.slug = spec.slug
    task.agent_config = {"max_steps": 250}
    task.validation = [
        {
            "type": "grader_contract",
            "status": "available",
            "implementation": "evaluator/task_grader.py",
            "reference_record": spec.task_id,
            "agent_visible": False,
            "task_objective": spec.prompt,
            "output_mode": spec.output_mode,
            "target_path": spec.target_path,
            "success_condition": (
                "Complete the task_objective using the company sources and tools, "
                "and produce the requested deliverable in its stated format and "
                "location. The task-specific private evaluator checks factual "
                "correctness, completeness, output-format compliance, artifact "
                "integrity when applicable, and absence of grading or isolation "
                "errors. Strict pass requires every weighted atomic criterion to "
                "pass; partial completion receives the weighted reward subject to "
                "the documented hard-failure caps."
            ),
        },
        {
            "type": "release_validation",
            "status": "passed",
            "checks": [
                "task discovered from the deployed environment manifest",
                "task prompt present in the taskset validation contract and yielded on task start",
                "private reference record present",
                "evaluator imports successfully",
                "source files and ERP state reset deterministically",
                "agent shell isolation attestation passes before execution",
            ],
        },
    ]
    task.columns = {
        "company": "pinehaven-motion-systems",
        "workflow": spec.workflow,
        "level": "senior-corporate-finance",
        "output_mode": spec.output_mode,
        "difficulty": spec.difficulty,
        "snapshot": "2026-06-30-pre-close",
        "grading": "weighted-atomic-hybrid-v1",
        "prompt_available": True,
        "validation_contract_available": True,
    }
    tasks.append(task)

# HUD discovers Task objects from globals and from ``tasks``. Avoid exposing
# the loop variable as an additional discoverable object.
del task, spec
