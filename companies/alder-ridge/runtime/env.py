from __future__ import annotations

import asyncio
import contextlib
import os
import socket
from pathlib import Path

from hud.capabilities import Capability
from hud.environment import Environment

from alder_ridge_company.mcp_server import server
from alder_ridge_company.reset import PROJECT_ROOT, reset_company
from alder_ridge_company.workspace_security import (
    IsolatedWorkspace,
    agent_runtime_mounts,
    safe_agent_environment,
    verify_workspace_isolation,
)
from task_templates import register_task_templates


RUNTIME_ROOT = Path(os.environ.get("COMPANY_RUNTIME_ROOT", PROJECT_ROOT / ".runtime" / "workspace")).resolve()
STATE_ROOT = Path(os.environ.get("COMPANY_STATE_ROOT", PROJECT_ROOT / ".runtime" / "state")).resolve()

# Capture grader credentials in the environment process, then remove them
# before the agent-facing shell starts. The verifier can still use the
# in-memory values, while `env`, `/proc`, and child shell processes cannot.
_SEMANTIC_CREDENTIALS = {
    "hud_api_key": os.environ.pop("HUD_API_KEY", None),
    "openai_api_key": os.environ.pop("OPENAI_API_KEY", None),
}

env = Environment(name="alder-ridge-corporate-finance-v1")
WORKSPACE = IsolatedWorkspace(
    RUNTIME_ROOT,
    mounts=agent_runtime_mounts(),
    env=safe_agent_environment(),
)


@env.initialize
async def start_workspace() -> None:
    await WORKSPACE.start()
    env.add_capability(WORKSPACE.capability("shell"))


@env.shutdown
async def stop_workspace() -> None:
    await WORKSPACE.stop()

_server_task: asyncio.Task | None = None


@env.initialize
async def initialize_company() -> None:
    global _server_task
    reset_company()
    verify_workspace_isolation(WORKSPACE, project_root=PROJECT_ROOT)
    if _server_task is None:
        sock = socket.socket()
        sock.bind(("", 0))
        port = sock.getsockname()[1]
        sock.close()
        _server_task = asyncio.create_task(
            server.run_async(
                transport="http",
                host="127.0.0.1",
                port=port,
                show_banner=False,
            )
        )
        await asyncio.sleep(.35)
        env.add_capability(
            Capability.mcp(
                name="contractor_accounting",
                url=f"http://127.0.0.1:{port}/mcp",
            )
        )


@env.shutdown
async def shutdown_company() -> None:
    global _server_task
    if _server_task is not None:
        _server_task.cancel()
        with contextlib.suppress(asyncio.CancelledError):
            await _server_task
        _server_task = None


TASK_TEMPLATES = register_task_templates(
    env,
    RUNTIME_ROOT,
    STATE_ROOT,
    PROJECT_ROOT,
    semantic_credentials=_SEMANTIC_CREDENTIALS,
)
