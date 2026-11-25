# syntax=docker/dockerfile:1
#
# A clean Linux machine holding nothing but this repository and its locked dependencies.
# Running it checks that kyralaw-redline and its benchmark scorer work on a platform other
# than the one they were written on, depend on no other codebase, and need no network once
# built:
#
#   docker build -t kyralaw-redline .
#   docker run --rm --network none kyralaw-redline
#   docker run --rm --network none -v /path/to/kyralaw-redline-bench:/data:ro kyralaw-redline
#
# The second run also scores kyralaw-redline-bench with every engine and compares each
# score with bench/results.md. Building the benchmark itself needs macOS with Microsoft
# Word (bench/), so it is not part of this image. See scripts/verify.sh.

FROM python:3.12-slim-bookworm

COPY --from=ghcr.io/astral-sh/uv:0.11.20 /uv /bin/uv

ENV UV_PROJECT_ENVIRONMENT=/opt/venv \
    UV_LINK_MODE=copy \
    UV_COMPILE_BYTECODE=1 \
    PATH=/opt/venv/bin:$PATH \
    PYTHONDONTWRITEBYTECODE=1 \
    # python-redlines runs a self-contained .NET binary; the slim image has no ICU.
    DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=1

RUN useradd --create-home --uid 1000 checker
WORKDIR /repo

# Dependencies first, from the lock file alone, so a code change does not reinstall them.
COPY pyproject.toml uv.lock README.md LICENSE ./
RUN uv sync --frozen --no-install-project --all-groups

COPY . .
RUN uv sync --frozen --all-groups \
    && rm /bin/uv \
    && chown -R checker:checker /repo /opt/venv

USER checker
ENTRYPOINT ["scripts/verify.sh"]
