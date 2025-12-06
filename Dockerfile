FROM debian:bookworm

ENV DEBIAN_FRONTEND=noninteractive

# -------------------------------------------------------
# System Dependencies
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
    libjpeg62-turbo \
    libxrender1 \
    libxext6 \
    libsm6 \
    libxinerama1 \
    libxrandr2 \
    libfontconfig1 \
    fonts-dejavu \
    fonts-liberation \
    fonts-noto \
    fonts-noto-cjk \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5000

CMD ["python3", "app.py"]
