# 選擇 Python 基底映像
FROM python:3.11-slim

# 安裝 Chromium 與相關依賴套件（Debian/Ubuntu 系統需補充依賴）
RUN apt-get update && \
    apt-get install -y \
        chromium \
        chromium-driver \
        fonts-liberation \
        libasound2 \
        libatk-bridge2.0-0 \
        libatk1.0-0 \
        libc6 \
        libcairo2 \
        libcups2 \
        libdbus-1-3 \
        libdrm2 \
        libexpat1 \
        libfontconfig1 \
        libgbm1 \
        libgcc1 \
        libglib2.0-0 \
        libgtk-3-0 \
        libnspr4 \
        libnss3 \
        libpango-1.0-0 \
        libpangocairo-1.0-0 \
        libstdc++6 \
        libx11-6 \
        libx11-xcb1 \
        libxcb1 \
        libxcomposite1 \
        libxcursor1 \
        libxdamage1 \
        libxext6 \
        libxfixes3 \
        libxi6 \
        libxrandr2 \
        libxrender1 \
        libxss1 \
        libxtst6 \
        lsb-release \
        xdg-utils \
        --no-install-recommends && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# 設定環境變數讓 Selenium 找到 Chromium 路徑
ENV CHROME_BIN=/usr/bin/chromium
ENV PATH="$PATH:/usr/bin/chromium"

# 複製專案程式碼
COPY . /app
WORKDIR /app

# 安裝 Python 套件
RUN pip install --no-cache-dir -r requirements.txt

# 啟動 Flask（可根據 deploy 平台調整）
CMD ["gunicorn", "app:app", "-b", "0.0.0.0:5000"]
