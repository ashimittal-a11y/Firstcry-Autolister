FROM python:3.11-slim

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements and install
COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy backend files
COPY backend/ .

# Expose port
EXPOSE 8000

# Start server using dynamic port mapping for Railway hosting
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}
