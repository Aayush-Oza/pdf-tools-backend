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
    libreoffice-core \
    libreoffice-writer \
    libreoffice-impress \
    libreoffice-calc \
    default-jre \
    poppler-utils \
    ghostscript \
    tesseract-ocr \
    tesseract-ocr-hin \
    tesseract-ocr-mar \
    tesseract-ocr-guj \
    libjpeg-turbo8 \
    libtiff5 \
    libxrender1 \
    libxext6 \
    libsm6 \
    fonts-dejavu-core \
    fonts-noto-core \
    fonts-noto-ui-extra \
    fonts-noto-devanagari \
    fonts-noto-gujarati \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*
    
# -------------------------------------------------
# Set working directory
# -------------------------------------------------
WORKDIR /app

# Install Python packages early to leverage caching
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

# Copy the application
COPY . .

# Port for Flask app
EXPOSE 5000

# -------------------------------------------------
# Entry Point: Run Flask app
# -------------------------------------------------
CMD ["gunicorn", "--timeout", "180", "-w", "2", "-k", "sync", "-b", "0.0.0.0:5000", "app:app"]



