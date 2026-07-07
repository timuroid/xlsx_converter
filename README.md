# УС SAP XLSX Калькулятор КК

Проект для простого преобразования SAP-выгрузки в XLSX: на вход подается исходная таблица с серыми колонками SAP, на выходе формируется новый XLSX с добавленными расчетными колонками и Excel-формулами.

## Цель

- Не хранить пользовательские входные и выходные файлы в будущей продуктовой реализации.
- Держать расчетные правила отдельно от кода, чтобы их можно было быстро менять.
- Сохранить результат в формате XLSX, пригодном для проверки и ручной доработки в Excel.

## Текущий MVP

Есть две реализации:

- `project_files/browser_app/` - основная browser-only версия без backend;
- `project_files/app/` - Python/Flask версия, оставлена как эталон и fallback.

Browser-only запуск:

```bash
cd project_files/browser_app
pnpm install
pnpm dev
```

Открыть:

```text
http://127.0.0.1:8002
```

Browser-only проверка:

```bash
cd project_files/browser_app
pnpm test:transform
```

Python MVP находится в `project_files/app/`.

Пример запуска:

```bash
cd project_files/app
python3 -m sap_xlsx_transformer \
  "../../Источники/inbox/100019128.XLSX" \
  "../exports/100019128_transformed.xlsx"
```

Проверка:

```bash
cd project_files/app
python3 -m unittest discover -s tests
```

Локальный web-запуск:

```bash
cd project_files/app
flask --app sap_xlsx_transformer.web run --host 0.0.0.0 --port 8000
```

## Вход и выход

Вход:

- XLSX-выгрузка из SAP с колонками `A:S`.
- В файле могут быть строки с разными видами документа.

Выход:

- исходные колонки `A:S`;
- расчетные колонки `T:W`;
- служебная дата расчета в `AB1`;
- формулы проставляются только для строк, где `Вид документа = DR`.
- в browser-only версии формулы пишутся без кэшированных значений; Excel пересчитает их при открытии.

Колонки `X:AA` из примера отложены до подтверждения правил заказчиком и в актуальной browser-only версии не считаются.

## GitLab / публикация

Проект подготовлен как локальный git-репозиторий без подключенного remote. До ответа DevOps код можно держать локально и при необходимости передать им следующие вводные:

- основная поставляемая часть: `project_files/browser_app/`;
- тип приложения: статический browser-only frontend на Vite/TypeScript;
- backend для основной версии не нужен;
- сборка:

```bash
cd project_files/browser_app
pnpm install --frozen-lockfile
pnpm build
```

- публиковать нужно содержимое `project_files/browser_app/dist`;
- пользовательские XLSX обрабатываются в браузере и не отправляются на сервер.

В `.gitignore` намеренно исключены:

- `node_modules/`;
- `dist/` и `dist-test/`;
- Python `.venv/` и `__pycache__/`;
- локальные результаты `project_files/exports/`;
- входные XLSX из `Источники/inbox/`, потому что они могут содержать рабочие данные.

Если понадобятся тестовые XLSX в GitLab, сначала нужно добавить обезличенные fixtures и явно разрешить их в `.gitignore`.

## Структура

```text
Источники/
  inbox/       # исходные примеры от пользователя
  reference/   # справочные материалы
  research/    # исследовательские заметки
project_files/
  browser_app/ # основная browser-only версия
  app/         # Python/Flask fallback и эталон
  docs/        # проектная документация
  deliverables/
  exports/     # локальные тестовые выгрузки
  management/
```
