FROM python:3.12.13-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PINEHAVEN_REQUIRE_SHELL_ISOLATION=1 \
    COMPANY_SHELL_UID=1000 \
    COMPANY_RUNTIME_ROOT=/workspace \
    COMPANY_STATE_ROOT=/state

RUN apt-get update && apt-get install -y --no-install-recommends \
      bubblewrap \
      libreoffice-calc \
      libreoffice-impress \
      libreoffice-writer \
      poppler-utils \
      util-linux \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /environment
COPY runtime/ /environment/
COPY evaluator/ /environment/evaluator/

RUN pip install --no-cache-dir uv==0.11.29 \
    && uv lock --check \
    && uv sync --frozen --no-dev --no-editable \
    && uv run python scripts/prepare_agent_runtime.py \
    && rm -rf /root/.cache

COPY data/ /environment/company/seed/
COPY source-files/ /environment/company/seed/workspace/

ENV PATH="/environment/.venv/bin:${PATH}"

RUN cp -a /environment/.build/agent-runtime /opt/pinehaven-agent-venv \
    && chmod -R a+rX,go-w /opt/pinehaven-agent-venv \
    && chmod -R go-rwx /environment \
    && chmod 0711 /environment \
    && mkdir -p /workspace /state \
    && chmod 0777 /workspace \
    && chmod 0700 /state

EXPOSE 8765
CMD ["hud", "serve", "env.py", "--host", "0.0.0.0", "--port", "8765"]
