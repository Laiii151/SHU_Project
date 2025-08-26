# 選擇基底映像
FROM python:3.11-slim

# 安裝 Chromium 和驅動程式
RUN apt-get update && apt-get install -y chromium chromium-driver

# 複製程式碼到容器內
COPY . /app
WORKDIR /app

# 安裝 Python 套件
RUN pip install --no-cache-dir -r requirements.txt

# 設定環境變數讓 Selenium 找到瀏覽器
ENV CHROME_BIN=/usr/bin/chromium

# 啟動 Flask (根據你的procfile或實際執行命令)
CMD ["gunicorn", "app:app", "-b", "0.0.0.0:5000"]
