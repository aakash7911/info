FROM python:3.11-slim

# Install system Chromium & ChromeDriver for headless browser automation + procps for memory/process management
RUN apt-get update && apt-get install -y \
    chromium \
    chromium-driver \
    libfreetype6 \
    procps \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Set Selenium cache directory to writable /tmp
ENV SELENIUM_CACHE_DIR=/tmp/.selenium

# Strict memory safety: enforce single worker and unbuffered output on Render (512MB RAM cap)
ENV WEB_CONCURRENCY=1
ENV PYTHONUNBUFFERED=1

# Install Python requirements
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source
COPY . .

# Expose web port
EXPOSE 8000

# Start UMS Student Dossier Portal with single worker
CMD sh -c "uvicorn student_app:app --host 0.0.0.0 --port ${PORT:-8000} --workers 1 --limit-concurrency 10"
