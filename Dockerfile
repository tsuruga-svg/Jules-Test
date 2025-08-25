# Pythonの公式イメージをベースにする
FROM python:3.9-slim

# 作業ディレクトリを設定
WORKDIR /app

# 依存関係ファイルをコピーしてインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# アプリケーションコードをコピー
COPY ./app .

# アプリケーションがリッスンするポートを公開
EXPOSE 5000

# アプリケーションを起動するコマンド
CMD ["flask", "run", "--host=0.0.0.0"]
