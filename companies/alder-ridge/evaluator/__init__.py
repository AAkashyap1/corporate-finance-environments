"""TASK-style deterministic evaluator for the Alder Ridge finance task set."""

from .task_grader import grade_task, load_reference

__all__ = ["grade_task", "load_reference"]
