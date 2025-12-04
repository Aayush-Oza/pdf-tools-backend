FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive
WORKDIR /app

# ------------------------------------------------------
# Install system dependencies for LibreOffice, Poppler,
# Ghostscript, Tesseract OCR, and fonts.
# ------------------------------------------------------
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice \
        uno-libs-private \
        ure \
        libglu1-mesa \
        libxinerama1 \
        libxrandr2 \
        libxcursor1 \
        libxrender1 \
        libfontconfig1 \
        poppler-utils \
        ghostscript \
        tesseract-ocr \
        fonts-dejavu-core \
        fonts-liberation \
        fonts-noto \
        fonts-noto-cjk \
        fonts-freefont-ttf \
        && apt-get clean && rm -rf /var/lib/apt/lists/*

# ------------------------------------------------------
# Copy project files
# ------------------------------------------------------
COPY . .

# ------------------------------------------------------
# Install Python packages
# ------------------------------------------------------
RUN pip install --no-cache-dir -r requirements.txt

# ------------------------------------------------------
# Expose API port + run server
# ------------------------------------------------------
EXPOSE 5000
CMD ["gunicorn", "app:app", "-b", "0.0.0.0:5000", "--workers", "2", "--timeout", "200"]
