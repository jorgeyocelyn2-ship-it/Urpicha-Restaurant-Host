FROM python:3.13-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

RUN apt-get update \
    && apt-get install -y --no-install-recommends tesseract-ocr tesseract-ocr-spa \
    && rm -rf /var/lib/apt/lists/*

# Render SSH/SCP injects the account's authorized key at runtime. Ensure the
# Docker image has a secure SSH directory for the root runtime user.
RUN mkdir -p /root/.ssh && chmod 700 /root/.ssh

WORKDIR /opt/render/project/src
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

CMD ["/bin/sh", "-c", "gunicorn --workers 1 --threads 8 --timeout 180 --bind 0.0.0.0:${PORT:-8080} app:app"]
