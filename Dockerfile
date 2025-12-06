FROM python:3.11-slim

# Avoid interruptions
ENV DEBIAN_FRONTEND=noninteractive

# Install system deps
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    poppler-utils \
    ghostscript \
    tesseract-ocr \
    libtiff5 \
    libjpeg62 \
    libpng16-16 \
    libxrender1 \
    libxext6 \
    libsm6 \
    fonts-dejavu-core \
    fonts-noto \
    fonts-noto-cjk \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Install pip deps
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App code
COPY . .

# Run server
CMD ["gunicorn", "-b", "0.0.0.0:10000", "app:app"]
