FROM python:3.13-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

RUN apt-get update \
    && apt-get install -y --no-install-recommends tesseract-ocr tesseract-ocr-spa \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /opt/render/project/src
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

CMD ["/bin/sh", "-c", "python app.py ${PORT:-8080}"]
