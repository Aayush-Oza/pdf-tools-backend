FROM python:3.11-slim

# Avoid tzdata interactive prompt
ENV DEBIAN_FRONTEND=noninteractive

# -------------------------------------------------------
# Install LibreOffice + Poppler + Ghostscript + Tesseract
# -------------------------------------------------------
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    uno-libs-private \
    ure \
    default-jre \
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

# -------------------------------------------------------
# Set working directory
# -------------------------------------------------------
WORKDIR /app

# -------------------------------------------------------
# Copy requirements and install Python deps
# -------------------------------------------------------
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# -------------------------------------------------------
# Copy backend source code
# -------------------------------------------------------
COPY . .

# -------------------------------------------------------
# Expose port used by Flask
# -------------------------------------------------------
EXPOSE 5000

# -------------------------------------------------------
# Start Flask app
# -------------------------------------------------------
CMD ["python", "app.py"]
