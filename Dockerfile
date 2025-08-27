FROM python:3.11-slim

# Prevent Python from writing .pyc files; ensure unbuffered logs
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# Install minimal OS packages (tzdata for correct timezones)
RUN apt-get update \
    && apt-get install -y --no-install-recommends tzdata \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies first (leverage Docker layer caching)
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Create non-root user and prepare app directory
RUN useradd -m botuser \
    && mkdir -p /app/data \
    && chown -R botuser:botuser /app

# Copy application code
COPY bot.py ./
COPY README.md ./
# COPY Sample.xlsx ./Sample.xlsx

# Use non-root user for security
USER botuser

# Persist chat data and uploaded Excel files
VOLUME ["/app/data"]

# Default envs (override BOT_TOKEN at runtime or mount .env)
ENV TEST_MODE=0

# Run the bot
CMD ["python", "bot.py"]


