# ---- Frontend build ----
FROM node:20 AS fe-build
WORKDIR /app/frontend
COPY frontend/package.json frontend/package-lock.json* ./
RUN if [ -f package-lock.json ]; then npm ci; else npm install; fi
COPY frontend/ .
RUN npm run build

# ---- Backend runtime ----
FROM python:3.12-slim AS app
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
WORKDIR /app

# libs for Pillow
RUN apt-get update \
 && apt-get install -y --no-install-recommends build-essential libjpeg62-turbo-dev zlib1g-dev \
 && rm -rf /var/lib/apt/lists/*

# deps
COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# backend code
COPY backend/ .

# built frontend -> served by FastAPI at "/"
COPY --from=fe-build /app/frontend/dist ./static

EXPOSE 8000
CMD ["sh","-c","uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
