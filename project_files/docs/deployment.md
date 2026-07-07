# Легкое веб-развертывание

## Рекомендуемый вариант

Для поддомена IntBiS самый простой рабочий вариант - Docker-контейнер за reverse proxy:

```text
subdomain.intbis.ru
  -> nginx / traefik
  -> Docker container :8000
  -> Flask + gunicorn
  -> temporary XLSX processing
```

Почему так:

- не нужна база данных;
- не нужно постоянное файловое хранилище;
- приложение легко перенести на любой сервер;
- DevOps может привязать поддомен стандартным reverse proxy;
- правила расчета остаются в `rules.toml`.

## Что хранится

В продуктовой схеме XLSX-файлы не хранятся.

Во время запроса приложение:

1. принимает файл;
2. кладет его во временную директорию;
3. создает выходной XLSX;
4. читает результат в память;
5. удаляет временную директорию;
6. отдает файл пользователю.

## Локальный запуск

```bash
cd project_files/app
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
flask --app sap_xlsx_transformer.web run --host 0.0.0.0 --port 8000
```

Открыть:

```text
http://127.0.0.1:8000
```

## Docker

Сборка:

```bash
cd project_files/app
docker build -t us-sap-xlsx-calculator .
```

Запуск:

```bash
docker run --rm -p 8000:8000 \
  -e FLASK_SECRET_KEY="replace-with-random-secret" \
  us-sap-xlsx-calculator
```

Проверка:

```bash
curl http://127.0.0.1:8000/health
```

## Reverse proxy

Пример nginx location:

```nginx
server {
    server_name sap-kk.intbis.ru;

    client_max_body_size 32m;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

## Переменные окружения

```text
FLASK_SECRET_KEY  секрет для flash-сообщений Flask
```

Для текущей реализации база данных, S3/MinIO и постоянный volume не нужны.

## Serverless

Serverless возможен, но для этой задачи хуже как первый шаг:

- XLSX-файлы могут упираться в лимиты размера запроса/ответа;
- холодный старт Python с `openpyxl` может быть заметным;
- сложнее отлаживать временную файловую обработку;
- все равно понадобится аккуратная настройка домена и лимитов.

Если позже будет требование именно serverless, ядро `transform_workbook` можно переиспользовать в функции без изменения расчетной логики.
