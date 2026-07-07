# MVP XLSX Transformer

Минимальный обработчик SAP-выгрузки.

## Установка

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Для локальной проверки можно использовать системный Python, если в нем уже есть `openpyxl`.

## CLI-запуск

```bash
python3 -m sap_xlsx_transformer input.xlsx output.xlsx
```

С фиксированной датой расчета:

```bash
python3 -m sap_xlsx_transformer input.xlsx output.xlsx --calc-date 2026-07-01
```

С кастомным конфигом правил:

```bash
python3 -m sap_xlsx_transformer input.xlsx output.xlsx --rules rules.toml
```

## Web-запуск

```bash
flask --app sap_xlsx_transformer.web run --host 0.0.0.0 --port 8000
```

Production-like запуск:

```bash
gunicorn --bind 0.0.0.0:8000 --workers 2 --threads 4 --timeout 120 sap_xlsx_transformer.web:app
```
