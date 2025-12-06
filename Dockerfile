FROM ubuntu:22.04

ENV DEBIAN_FRONTEND=noninteractive

# Speed boost: Reduce apt cache + fewer packages
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 \
    python3-pip \
    python3-distutils \
    python3-uno \
    libreoffice \
    poppler-utils \
    ghostscript \
    tesseract-ocr \
    libjpeg-turbo8 \
    libtiff5 \
    libxrender1 \
    libxext6 \
    libsm6 \
    fontconfig \
    fonts-dejavu-core \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Create app folder
WORKDIR /app

# Install dependencies early to use Docker cache
COPY requirements.txt /app/
RUN pip3 install --no-cache-dir -r requirements.txt

# Copy the app
COPY . .

EXPOSE 5000

# Prevent LibreOffice crash on Render
ENV HOME=/tmp

# Start Flask
CMD ["python3", "app.py"]
