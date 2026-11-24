FROM python:3.12-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=5000 \
    HOST=0.0.0.0

WORKDIR /app

# Install system dependencies if required
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY . .

# Create temp directories
RUN mkdir -p temp_uploads temp_outputs

EXPOSE 5000

# Run with Gunicorn WSGI for production
CMD ["gunicorn", "-w", "2", "--threads", "4", "-b", "0.0.0.0:5000", "server:app", "--timeout", "180"]
