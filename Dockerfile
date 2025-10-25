# syntax=docker/dockerfile:1
FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive \
    PIP_NO_CACHE_DIR=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# System packages: LibreOffice, ImageMagick, fonts
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      libreoffice \
      libreoffice-script-provider-python \
      python3-uno \
      imagemagick \
      ghostscript \
      fonts-noto-cjk \
      fonts-ipafont \
      curl ca-certificates && \
    # ImageMagick の PDF 書き込み許可（必要な環境でのみ適用）
    sed -i 's/<policy domain="coder" rights="none" pattern="PDF"\/>/<policy domain="coder" rights="read|write" pattern="PDF"\/>/' /etc/ImageMagick-6/policy.xml || true && \
    rm -rf /var/lib/apt/lists/*

# App
WORKDIR /app
COPY app/ /app/app/

# Python deps
RUN pip install --upgrade pip && \
    pip install pydantic

# CLI 実行をデフォルトに（ヘルプ表示）
CMD ["python", "-m", "app.cli.candytune_cli", "--help"]
