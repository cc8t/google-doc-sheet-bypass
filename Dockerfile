FROM python:3.12-slim

WORKDIR /app

# Install uv
COPY --from=ghcr.io/astral-sh/uv:latest /uv /uvx /bin/

# Copy dependency files
COPY pyproject.toml ./

# Generate lock file and install dependencies
RUN uv lock && uv sync --frozen --no-dev

# Copy application files
COPY main.py ./
COPY templates/ ./templates/
COPY static/ ./static/

EXPOSE 8000

CMD ["uv", "run", "uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
