FROM python:3.11-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV REALESRGAN_MODEL_DIR=/opt/realesrgan/models

COPY requirements.txt ./
RUN apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates wget unzip libvulkan1 && \
    rm -rf /var/lib/apt/lists/*
RUN pip install --no-cache-dir -r requirements.txt

COPY realesrgan /opt/realesrgan
RUN chmod +x /opt/realesrgan/realesrgan-ncnn-vulkan

COPY . .

CMD ["sh", "-c", "gunicorn -w 2 -k gthread --threads 4 -b 0.0.0.0:${PORT:-10000} app:app"]
