from __future__ import annotations

from datetime import date
from io import BytesIO
import os
from pathlib import Path
from tempfile import TemporaryDirectory

from flask import Flask, Response, flash, redirect, render_template_string, request, send_file, url_for
from werkzeug.utils import secure_filename

from .rules import load_rules
from .transformer import transform_workbook


MAX_UPLOAD_BYTES = 32 * 1024 * 1024
ALLOWED_SUFFIXES = {".xlsx"}


PAGE_TEMPLATE = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Калькулятор КК SAP XLSX</title>
  <style>
    :root {
      color-scheme: light;
      --text: #172033;
      --muted: #5b6475;
      --line: #d8dee8;
      --surface: #f6f8fb;
      --accent: #2457a6;
      --accent-strong: #163f80;
      --error: #9f1d24;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      color: var(--text);
      background: #ffffff;
    }
    main {
      width: min(920px, calc(100vw - 32px));
      margin: 0 auto;
      padding: 40px 0;
    }
    header {
      border-bottom: 1px solid var(--line);
      padding-bottom: 18px;
      margin-bottom: 28px;
    }
    h1 {
      font-size: 28px;
      line-height: 1.2;
      margin: 0 0 8px;
      font-weight: 650;
    }
    p {
      margin: 0;
      color: var(--muted);
      line-height: 1.5;
    }
    form {
      display: grid;
      gap: 18px;
      padding: 24px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
    }
    label {
      display: grid;
      gap: 8px;
      font-weight: 600;
    }
    input[type="file"],
    input[type="date"] {
      width: 100%;
      min-height: 44px;
      padding: 9px 10px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #ffffff;
      color: var(--text);
      font: inherit;
    }
    .hint {
      font-size: 14px;
      color: var(--muted);
      font-weight: 400;
    }
    .actions {
      display: flex;
      gap: 12px;
      align-items: center;
      flex-wrap: wrap;
    }
    button {
      min-height: 44px;
      border: 0;
      border-radius: 6px;
      padding: 0 18px;
      background: var(--accent);
      color: #ffffff;
      font: inherit;
      font-weight: 650;
      cursor: pointer;
    }
    button:hover { background: var(--accent-strong); }
    .notice {
      margin-bottom: 18px;
      padding: 12px 14px;
      border-radius: 6px;
      border: 1px solid #efc7cb;
      background: #fff1f2;
      color: var(--error);
    }
    .details {
      margin-top: 24px;
      display: grid;
      gap: 8px;
      color: var(--muted);
      font-size: 14px;
    }
    code {
      color: var(--text);
      background: #eef2f7;
      padding: 1px 5px;
      border-radius: 4px;
    }
  </style>
</head>
<body>
  <main>
    <header>
      <h1>Калькулятор КК SAP XLSX</h1>
      <p>Загрузите SAP-выгрузку XLSX. Сервис добавит расчетные колонки с формулами и сразу вернет новый файл.</p>
    </header>

    {% with messages = get_flashed_messages() %}
      {% if messages %}
        {% for message in messages %}
          <div class="notice">{{ message }}</div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    <form method="post" action="{{ url_for('convert') }}" enctype="multipart/form-data">
      <label>
        XLSX-файл
        <input type="file" name="file" accept=".xlsx" required>
        <span class="hint">Файл обрабатывается временно и не сохраняется после ответа.</span>
      </label>

      <label>
        Дата расчета
        <input type="date" name="calc_date" value="{{ today }}">
        <span class="hint">Попадает в <code>AB1</code> и используется в формулах для открытых позиций.</span>
      </label>

      <div class="actions">
        <button type="submit">Сконвертировать XLSX</button>
        <span class="hint">Максимальный размер файла: 32 МБ.</span>
      </div>
    </form>

    <section class="details">
      <div>Формулы применяются к строкам, где <code>Вид документа = DR</code>.</div>
      <div>Исходные колонки SAP: <code>A:S</code>. Расчетные колонки: <code>T:W</code>.</div>
    </section>
  </main>
</body>
</html>
"""


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES
    app.config["SECRET_KEY"] = os.environ.get("FLASK_SECRET_KEY", "dev-only-secret")

    @app.get("/")
    def index() -> str:
        return render_template_string(PAGE_TEMPLATE, today=date.today().isoformat())

    @app.get("/health")
    def health() -> Response:
        return Response("ok\n", mimetype="text/plain")

    @app.post("/convert")
    def convert() -> Response:
        uploaded = request.files.get("file")
        if uploaded is None or uploaded.filename == "":
            flash("Выберите XLSX-файл.")
            return redirect(url_for("index"))

        source_name = secure_filename(uploaded.filename) or "upload.xlsx"
        if Path(source_name).suffix.lower() not in ALLOWED_SUFFIXES:
            flash("Поддерживаются только файлы .xlsx.")
            return redirect(url_for("index"))

        try:
            calculation_date = _parse_calc_date(request.form.get("calc_date"))
        except ValueError:
            flash("Дата расчета должна быть в формате YYYY-MM-DD.")
            return redirect(url_for("index"))

        rules = load_rules()
        with TemporaryDirectory(prefix="sap-xlsx-") as tmp_dir_name:
            tmp_dir = Path(tmp_dir_name)
            input_path = tmp_dir / source_name
            output_path = tmp_dir / _output_filename(source_name)
            uploaded.save(input_path)

            try:
                transform_workbook(input_path, output_path, rules, calculation_date)
            except Exception as error:
                app.logger.exception("XLSX conversion failed")
                flash(f"Не удалось обработать файл: {error}")
                return redirect(url_for("index"))

            output_bytes = BytesIO(output_path.read_bytes())

        return send_file(
            output_bytes,
            as_attachment=True,
            download_name=_output_filename(source_name),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @app.errorhandler(413)
    def request_entity_too_large(_: Exception) -> tuple[str, int]:
        return "Файл слишком большой. Максимальный размер: 32 МБ.\n", 413

    return app


def _parse_calc_date(raw_value: str | None) -> date | None:
    if raw_value in (None, ""):
        return None
    return date.fromisoformat(raw_value)


def _output_filename(source_name: str) -> str:
    path = Path(source_name)
    return f"{path.stem}_converted.xlsx"


app = create_app()
