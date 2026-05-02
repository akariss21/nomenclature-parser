# Базовый образ Python
FROM python:3.10-slim

# Установка необходимых системных пакетов и antiword
RUN apt-get update && apt-get install -y --no-install-recommends \
    antiword \
    libreoffice \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Установка зависимостей Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копирование приложения
WORKDIR /app
COPY . /app

# Запуск сервера
EXPOSE 9999
CMD ["python", "app.py"]
