# syntax=docker/dockerfile:1
FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive \
    PIP_NO_CACHE_DIR=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# System packages: LibreOffice, ImageMagick, fonts, xvfb
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      libreoffice \
      libreoffice-script-provider-python \
      python3-uno \
      imagemagick \
      ghostscript \
      fonts-noto-cjk \
      fonts-ipafont \
      xvfb \
      curl ca-certificates && \
    # ImageMagick の PDF 書き込み許可（必要な環境でのみ適用）
    sed -i 's/<policy domain="coder" rights="none" pattern="PDF"\/>/<policy domain="coder" rights="read|write" pattern="PDF"\/>/' /etc/ImageMagick-6/policy.xml || true && \
    rm -rf /var/lib/apt/lists/*

# 仮想ディスプレイを設定（LibreOfficeのGUI機能を使用可能にする）
ENV DISPLAY=:99

# システムのpython3-unoをPython 3.11から使えるようにシンボリックリンクを作成
RUN SYSTEM_UNO_PATH="/usr/lib/python3/dist-packages" && \
    TARGET_UNO_PATH="/usr/local/lib/python3.11/site-packages" && \
    if [ -d "$SYSTEM_UNO_PATH" ]; then \
        mkdir -p "$TARGET_UNO_PATH" && \
        ln -sf "$SYSTEM_UNO_PATH"/uno* "$TARGET_UNO_PATH/" && \
        ln -sf "$SYSTEM_UNO_PATH"/com "$TARGET_UNO_PATH/" 2>/dev/null || true; \
    fi

# App
WORKDIR /app
COPY app/ /app/app/

# Python deps
RUN pip install --upgrade pip && \
    pip install pydantic pikepdf

# CLI 実行をデフォルトに（ヘルプ表示）
CMD ["python", "-m", "app.cli.candytune_cli", "--help"]
