# УС SAP XLSX Калькулятор КК

Статическое browser-only приложение для преобразования SAP XLSX-выгрузки: пользователь выбирает один или несколько файлов, приложение добавляет расчетные колонки `T:AA` с Excel-формулами и отдает готовые XLSX обратно на скачивание.

Файлы обрабатываются в браузере. Backend для основной версии не нужен.

## Локальный запуск

```bash
cd project_files/browser_app
pnpm install
pnpm dev
```

Открыть:

```text
http://127.0.0.1:8002
```

## Проверка

```bash
cd project_files/browser_app
pnpm test:transform
```

## Production build

```bash
cd project_files/browser_app
pnpm install --frozen-lockfile
pnpm build
```

Результат сборки:

```text
project_files/browser_app/dist/
```

Для развертывания на поддомене достаточно отдавать содержимое папки `dist/` как обычную статику через nginx, GitLab Pages или другой static hosting.

## Docker

Локально из корня проекта:

```bash
docker compose up --build
```

Открыть:

```text
http://127.0.0.1:8088
```

Контейнер собирает `project_files/browser_app` и отдает production-сборку через nginx.

## Что публиковать

В Git должны попадать только:

- `project_files/browser_app/` - исходники браузерного приложения;
- `project_files/docs/` - проектная документация;
- корневые `README.md`, `CHANGE_HISTORY.md`, `AGENTS.md`, `.gitignore`.

Локальные исходные XLSX, справочные материалы, исследовательские заметки, Python fallback, результаты выгрузок и сборочные артефакты не публикуются.
