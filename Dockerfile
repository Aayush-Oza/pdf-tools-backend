FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive
WORKDIR /app

# ------------------------------------------------------
# Install system dependencies
# ------------------------------------------------------
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice \
        uno-libs-private \
        ure \
        default-jre \
        poppler-utils \
        ghostscript \
        tesseract-ocr \
        libtiff5 \
        libjpeg62-turbo \
        libpng16-16 \
        libxrender1 \
        libxext6 \
        libsm6 \
        libxinerama1 \
        libxrandr2 \
        libfontconfig1 \
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
# Install Python dependencies
# ------------------------------------------------------
RUN pip install --no-cache-dir -r requirements.txt

# ------------------------------------------------------
# Expose + Start Gunicorn
# ------------------------------------------------------
EXPOSE 5000
CMD ["gunicorn", "app:app", "-b", "0.0.0.0:5000", "--workers", "2", "--timeout", "200"]
