FROM python:3.11-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --upgrade pip setuptools wheel \
    && PIP_ONLY_BINARY=:all: pip install -r requirements.txt

COPY . .
ENV PORT=8080
CMD ["gunicorn","main:app","--bind","0.0.0.0:${PORT}","--workers","2","--threads","8","--timeout","120"]
