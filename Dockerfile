FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    poppler-utils \
    ghostscript \
    tesseract-ocr \
    libtiff5 \
    libjpeg62-turbo \
    libxrender1 \
    libxext6 \
    libsm6 \
    fontconfig \
    fonts-dejavu \
    fonts-liberation \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*
    
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

CMD ["gunicorn", "-b", "0.0.0.0:10000", "app:app"]
