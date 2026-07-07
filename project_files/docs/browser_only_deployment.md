# Browser-only развертывание

## Рекомендуемый вариант для поддомена

Основной вариант после browser-only POC - статическое приложение:

```text
sap-kk.intbis.ru
  -> nginx / CDN
  -> browser_app/dist/
  -> обработка XLSX в браузере пользователя
```

Backend, база данных, Python runtime и файловое хранилище не нужны.

## Что сказать DevOps

Нужна раздача статической папки `dist/` на поддомене. Приложение не требует backend API и не принимает файлы на сервер. Все XLSX обрабатываются локально в браузере пользователя.

Сборка:

```bash
cd project_files/browser_app
pnpm install
pnpm build
```

Результат:

```text
project_files/browser_app/dist/
```

Эту папку нужно положить за nginx/CDN.

## Пример nginx

```nginx
server {
    server_name sap-kk.intbis.ru;

    root /var/www/sap-kk/dist;
    index index.html;

    location / {
        try_files $uri $uri/ /index.html;
    }
}
```

## Безопасность данных

Файлы не уходят на сервер:

1. пользователь выбирает или перетаскивает `.xlsx`;
2. браузер читает файл через File API;
3. браузер добавляет расчетные колонки `T:W` и формулы;
4. браузер создает ссылку на скачивание результата;
5. при очистке списка временные object URL освобождаются.

## Поддержка браузеров

Целевые современные браузеры:

- Google Chrome;
- Microsoft Edge;
- Safari;
- Firefox;
- Opera.

Нужны стандартные возможности современного браузера: File API, Blob, object URL, WebAssembly/JS runtime для ExcelJS-бандла.

## Проверка

```bash
cd project_files/browser_app
pnpm test:transform
```

Тест проверяет browser-only трансформер на двух реальных XLSX-файлах из `Источники/inbox/`.

## Ограничения

- Формулы записываются без расчета кэшированных значений.
- Excel пересчитывает формулы при открытии файла.
- Очень большие XLSX могут временно нагружать память браузера пользователя.
- Колонки `X:AA` сейчас не считаются: они отложены до подтверждения правил заказчиком.
