FROM python:3.11-slim

ENV PIP_NO_CACHE_DIR=1 \
    PIP_ONLY_BINARY=:all: \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app
COPY requirements.txt .
RUN pip install --upgrade pip setuptools wheel \
    && pip install -r requirements.txt

COPY . .

# Koyeb inyecta PORT (normalmente 8080).
ENV PORT=8080
# OJO: en forma "shell" para expandir ${PORT}
CMD exec sh -lc "gunicorn main:app --bind 0.0.0.0:${PORT} --workers 2 --threads 8 --timeout 120"
