FROM python:3.13-slim

# 環境（ログ即時出力）
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=10000 \
    TZ=Asia/Tokyo

# OSパッケージ：LibreOffice + 日本語フォント（Noto CJK と IPA）
RUN apt-get update && \
    DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
      libreoffice libreoffice-calc \
      fonts-noto-cjk fonts-ipafont-gothic fonts-ipafont-mincho \
      tzdata && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 依存ライブラリ
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# アプリ本体
COPY . .

# Render はポート自動検出だが明示しておく
EXPOSE 10000

# 本番は gunicorn 推奨
CMD ["gunicorn", "-b", "0.0.0.0:10000", "app:app"]
