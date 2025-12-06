FROM ubuntu:22.04

ENV DEBIAN_FRONTEND=noninteractive

# -------------------------------------------------------
# Install system dependencies
# -------------------------------------------------------
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 \
    python3-pip \
    python3-distutils \
    libreoffice \
    poppler-utils \
    ghostscript \
    tesseract-ocr \
    libtiff5 \
    libjpeg-turbo8 \
    libxrender1 \
    libxext6 \
    libsm6 \
    fontconfig \
    fonts-dejavu-core \
    fonts-liberation \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# -------------------------------------------------------
# App folder
# -------------------------------------------------------
WORKDIR /app

COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5000

CMD ["python3", "app.py"]
