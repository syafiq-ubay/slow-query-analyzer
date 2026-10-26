# Use lightweight Python 3.11 base image
FROM python:3.11-slim

# Prevent Python from writing .pyc files and enable unbuffered logging
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PYTHONPATH="/app"

# Set work directory for app code
WORKDIR /app

# Copy requirements and install dependencies
COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source code
COPY exporter.py elastic_client.py correlator.py main.py /app/

# Set working directory to /workspace where host files are mounted
WORKDIR /workspace

# Default entrypoint
ENTRYPOINT ["python", "/app/main.py"]

# Default command arguments
CMD ["mysql", "--output", "output/mysql_slow_queries.xlsx"]
