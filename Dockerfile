FROM python:3.11-slim
RUN apt-get update && apt-get install -y \
    chromium \
    chromium-driver \
    libfreetype6 \
    procps \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
ENV SELENIUM_CACHE_DIR=/tmp/.selenium
ENV WEB_CONCURRENCY=1
ENV PYTHONUNBUFFERED=1
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 8000
CMD sh -c "uvicorn student_app:app --host 0.0.0.0 --port ${PORT:-8000} --workers 1 --limit-concurrency 10"
