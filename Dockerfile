FROM python:3.12-slim

# Install LibreOffice
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy backend files
COPY . .

# Start FastAPI
CMD uvicorn main:app --host 0.0.0.0 --port $PORT