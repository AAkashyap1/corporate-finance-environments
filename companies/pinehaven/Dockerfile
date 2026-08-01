FROM python:3.12.13-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PINEHAVEN_REQUIRE_SHELL_ISOLATION=1 \
    COMPANY_RUNTIME_ROOT=/workspace \
    COMPANY_STATE_ROOT=/state

RUN apt-get update && apt-get install -y --no-install-recommends \
      bubblewrap \
      libreoffice-calc \
      libreoffice-impress \
      libreoffice-writer \
      poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /environment
COPY runtime/ /environment/
COPY evaluator/ /environment/evaluator/
COPY data/ /environment/company/seed/
COPY source-files/ /environment/company/seed/workspace/

RUN pip install --no-cache-dir uv==0.11.29 \
    && uv lock --check \
    && uv sync --frozen --no-dev --no-editable \
    && uv run python scripts/prepare_agent_runtime.py \
    && rm -rf /root/.cache

ENV PATH="/environment/.venv/bin:${PATH}"

RUN mkdir -p /workspace /state && chmod 0777 /workspace /state

EXPOSE 8765
CMD ["hud", "serve", "env.py", "--host", "0.0.0.0", "--port", "8765"]
