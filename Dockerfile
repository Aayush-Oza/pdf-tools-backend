FROM ubuntu:22.04

ENV DEBIAN_FRONTEND=noninteractive
ENV HOME=/tmp

# -------------------------------------------------
# Install ONLY the required dependencies
# -------------------------------------------------
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 \
    python3-pip \
    python3-distutils \
    # LibreOffice core only
    libreoffice-core \
    libreoffice-writer \
    libreoffice-impress \
    # PDF + OCR tools
    poppler-utils \
    ghostscript \
    tesseract-ocr \
    # Image libs
    libjpeg-turbo8 \
    libtiff5 \
    libxrender1 \
    libxext6 \
    libsm6 \
    # Basic fonts
    fonts-dejavu-core \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# -------------------------------------------------
# App folder
# -------------------------------------------------
WORKDIR /app

# Install Python packages
COPY requirements.txt /app/
RUN pip3 install --no-cache-dir -r requirements.txt

# Copy application
COPY . .

EXPOSE 5000

# -------------------------------------------------
# Run Flask app
# -------------------------------------------------
CMD ["python3", "app.py"]
