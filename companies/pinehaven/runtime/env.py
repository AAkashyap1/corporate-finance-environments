from __future__ import annotations

import asyncio
import contextlib
import os
import socket
from pathlib import Path

from hud.capabilities import Capability
from hud.environment import Environment

from pinehaven_company.mcp_server import DATABASE_ENV, configure_database, server
from pinehaven_company.reset import PROJECT_ROOT, reset_company, runtime_roots
from pinehaven_company.workspace_security import (
    IsolatedWorkspace,
    agent_runtime_mounts,
    safe_agent_environment,
    verify_workspace_isolation,
)
from task_templates import register_task_templates


RUNTIME_ROOT, STATE_ROOT = runtime_roots()

# The agent shell never inherits semantic-grader credentials.
_SEMANTIC_CREDENTIALS = {
    "hud_api_key": os.environ.pop("HUD_API_KEY", None),
    "openai_api_key": os.environ.pop("OPENAI_API_KEY", None),
}

env = Environment(name="pinehaven-manufacturing-finance-v1")
WORKSPACE = IsolatedWorkspace(
    RUNTIME_ROOT,
    mounts=agent_runtime_mounts(),
    env=safe_agent_environment(),
    shell_uid=int(os.environ["COMPANY_SHELL_UID"])
    if os.environ.get("COMPANY_SHELL_UID")
    else None,
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
    _, database_path = reset_company()
    configure_database(database_path)
    os.environ.pop(DATABASE_ENV, None)
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
        await asyncio.sleep(0.35)
        env.add_capability(
            Capability.mcp(
                name="pinehaven_manufacturing_erp",
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
