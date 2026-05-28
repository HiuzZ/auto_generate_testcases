# ── Stage 1: build React frontend ────────────────────────────────────────────
FROM node:20-alpine AS frontend-builder

WORKDIR /app/frontend
COPY frontend/package*.json ./
RUN npm ci
COPY frontend/ .
RUN npm run build          # outputs to /app/frontend/dist

# ── Stage 2: Python runtime ───────────────────────────────────────────────────
FROM python:3.11-slim

WORKDIR /app

# System deps (for torch CPU and numpy)
RUN apt-get update && apt-get install -y --no-install-recommends \
        gcc g++ libgomp1 \
    && rm -rf /var/lib/apt/lists/*

# Install torch CPU-only first (much smaller than GPU build)
RUN pip install --no-cache-dir \
    torch --index-url https://download.pytorch.org/whl/cpu

# Install remaining Python dependencies (merge root + backend requirements)
COPY requirements.txt ./requirements.txt
COPY backend/requirements.txt ./backend/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt -r backend/requirements.txt

# Copy application code
COPY backend/     ./backend/
COPY prompts/     ./prompts/
COPY prompt_templates.py ./
COPY *.py         ./

# Copy built frontend from stage 1
COPY --from=frontend-builder /app/frontend/dist ./frontend/dist

# Writable runtime directories
RUN mkdir -p backend/uploads backend/outputs backend/data/versions

EXPOSE 8000

CMD ["uvicorn", "backend.app:app", "--host", "0.0.0.0", "--port", "8000"]
