"""
xlsx_pivot_to_grid.py

Парсер №2 из трёх задуманных ("для сводной").
Берёт .xlsx с классической сводной таблицей (на листе, не Data Model)
и генерирует РОВНО ДВА файла:

  1) <name>.data.js  — const DEMO_DATA = [...]; const DEMO_CONFIG = {...};
                       (то, что подгружается в кэш грида)
  2) <name>.html     — тот же шаблон, что и example_simple.html из пакета
                       pivotgrid-js, но с:
                         - поправленными путями к pivotgrid.css/pivotgrid.js/
                           pivot-widget.js (относительно реального node_modules,
                           без единой копии файлов);
                         - подключённым <name>.data.js вместо demo-data.js/
                           demo-config.js;
                         - без data-config/data-server (сервер не нужен).

Использует xlsx_to_pivotgrid.py для разбора структуры (не дублирует
логику парсинга).
"""

import json
import os
from pathlib import Path

from xlsx_to_pivotgrid import convert
from report_messages import rt, report_lang

# По умолчанию — не буквально "без лимита" (это дорого по времени/памяти на
# больших файлах, см. реальные замеры: ~40 сек и ~700 МБ на 1 млн строк),
# а разумный потолок. Переопределяется через переменную окружения (сервер
# может прокинуть значение из server/.env), при чистом CLI-использовании
# просто используется указанный здесь фолбэк.
DEFAULT_ROW_LIMIT = int(os.environ.get("PIVOT_DEFAULT_ROW_LIMIT", 500_000))


def find_pivotgrid_pkg():
    """
    Ищет папку пакета pivotgrid-js:
      1) переменная окружения PIVOTGRID_PKG (если задана явно)
      2) node_modules/pivotgrid-js рядом со скриптом
      3) node_modules/pivotgrid-js в текущей рабочей папке
    """
    env_path = os.environ.get("PIVOTGRID_PKG")
    if env_path:
        p = Path(env_path)
        if (p / "dist" / "pivotgrid.js").exists():
            return p
        raise FileNotFoundError(f"PIVOTGRID_PKG указывает на {p}, но там нет dist/pivotgrid.js")

    candidates = [
        Path(__file__).resolve().parent.parent.parent / "node_modules" / "pivotgrid-js",
        Path.cwd() / "node_modules" / "pivotgrid-js",
        Path.cwd().parent / "node_modules" / "pivotgrid-js",
        # старое расположение (на случай, если кто-то ставил раньше) — фолбэк
        Path(__file__).resolve().parent.parent.parent / "client" / "node_modules" / "pivotgrid-js",
    ]
    for c in candidates:
        if (c / "dist" / "pivotgrid.js").exists():
            return c

    raise FileNotFoundError(
        "Не нашёл пакет pivotgrid-js ни в одном из мест:\n"
        + "\n".join(f"  - {c}" for c in candidates)
        + "\n\nУстанови пакет в папке со скриптом командой:\n"
        "  npm install pivotgrid-js\n"
        "Или укажи путь явно через переменную окружения PIVOTGRID_PKG."
    )


def relhref(target_path: Path, from_dir: Path) -> str:
    """Относительный путь для href/src в html, всегда с прямыми слэшами."""
    rel = os.path.relpath(target_path, start=from_dir)
    return rel.replace(os.sep, "/")


# Шаблон — прямая копия структуры example_simple.html,
# только пути и подключаемые скрипты подставляются под конкретный отчёт.
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title}</title>
  <link rel="stylesheet" href="{css_href}">
  <style>
    body {{
      margin: 0;
      font-family: -apple-system, 'Segoe UI', sans-serif;
      background: #f4f5f7;
      color: #1a1a1a;
      padding: 12px;

      height: 100vh;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      box-sizing: border-box;
    }}

    #my-pivot {{
      flex: 1;
      min-height: 0;
    }}
    {nav_css}
  </style>
</head>
<body>

  {nav_html}

  <div id="my-pivot"
    data-disable-cache="true"
    data-disable-constructor-checkbox="true"
    data-demo="true"
    data-disable-drillthrough-panel="true"
    data-lang="{lang}">
  </div>

  <!-- Single bundle instead of many script tags -->
  <script src="{pivotgrid_js_href}"></script>

  <!-- Данные и конфиг, сгенерированные из {source_name} -->
  <script src="{data_js_href}"></script>

  <!-- Widget -->
  <script src="{widget_js_href}"></script>

  <script>{nav_js}</script>

  <footer style="text-align:center;padding:12px;font-size:12px;color:#999;">
    {pivot_label}: {pivot_table_name} · {rows_label}: {row_count}{limit_note}
  </footer>
</body>
</html>
"""

DATA_JS_TEMPLATE = """// Сгенерировано из {source_name}, сводная: {pivot_table_name}, строк: {row_count}
const DEMO_DATA = {data_json};

const DEMO_CONFIG = {config_json};
"""


def generate(xlsx_path, output_html_path, row_limit=None, nav=("", "", ""),
             target_sheet_name=None, occurrence_index=0, precomputed_results=None):
    nav_css, nav_html, nav_js = nav

    effective_limit = row_limit if row_limit is not None else DEFAULT_ROW_LIMIT
    if row_limit is not None and row_limit > DEFAULT_ROW_LIMIT:
        print(f"  [!] Сводная: запрошено {row_limit} строк, это выше дефолтного "
              f"потолка в {DEFAULT_ROW_LIMIT} — обработка может занять заметное время и память.")

    # Если данные уже посчитаны снаружи (сборщик отчёта считает их один раз
    # на файл, а не по разу на каждую вкладку) — переиспользуем, не парсим
    # заново тот же xlsx.
    results = precomputed_results if precomputed_results is not None else convert(xlsx_path, row_limit=effective_limit)
    if not results:
        raise ValueError("Не найдено ни одной поддерживаемой (worksheet) сводной")

    if target_sheet_name is not None:
        matches = [x for x in results if x["sheet_name"] == target_sheet_name]
        if occurrence_index >= len(matches):
            raise ValueError(
                f"На листе '{target_sheet_name}' не найдена сводная №{occurrence_index + 1}"
            )
        r = matches[occurrence_index]
    else:
        # Одна сводная по правилу проекта — берём первую найденную.
        r = results[0]

    html_path = Path(output_html_path).resolve()
    out_dir = html_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    data_js_path = html_path.with_suffix("").with_suffix(".data.js")

    pkg = find_pivotgrid_pkg()

    limit_note = ""
    if r.get("truncated"):
        limit_note = rt("shown_first_note", limit=effective_limit)

    # Файл №1 — данные и конфиг для кэша грида
    data_js = DATA_JS_TEMPLATE.format(
        source_name=Path(xlsx_path).name,
        pivot_table_name=r["pivot_table"],
        row_count=r["row_count"],
        data_json=json.dumps(r["data_sample"], ensure_ascii=False).replace("</", "<\\/"),
        config_json=json.dumps(r["config"], ensure_ascii=False, indent=2).replace("</", "<\\/"),
    )
    data_js_path.write_text(data_js, encoding="utf-8")

    # Файл №2 — html по образцу example_simple.html, пути — реальные
    # относительные пути до pivotgrid-js в node_modules, без копирования.
    html = HTML_TEMPLATE.format(
        title="Отчёт: " + Path(xlsx_path).name,
        css_href=relhref(pkg / "dist" / "pivotgrid.css", out_dir),
        pivotgrid_js_href=relhref(pkg / "dist" / "pivotgrid.js", out_dir),
        widget_js_href=relhref(pkg / "widget" / "pivot-widget.js", out_dir),
        data_js_href=data_js_path.name,
        source_name=Path(xlsx_path).name,
        pivot_table_name=r["pivot_table"],
        row_count=r["row_count"],
        limit_note=limit_note,
        pivot_label=rt("pivot_label"),
        rows_label=rt("rows_label"),
        lang=report_lang(),
        nav_css=nav_css,
        nav_html=nav_html,
        nav_js=nav_js,
    )
    html_path.write_text(html, encoding="utf-8")

    return r, html_path, data_js_path


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Генерирует .data.js + .html из xlsx со сводной для PivotGrid JS"
    )
    parser.add_argument("xlsx_path", nargs="?", default="test_dataS.xlsx")
    parser.add_argument("output_html", nargs="?", default="report.html")
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Максимум строк для декодирования. Не указано — грузятся все строки.",
    )
    args = parser.parse_args()

    r, html_path, data_js_path = generate(args.xlsx_path, args.output_html, row_limit=args.limit)
    print(f"Файл 1 (данные+конфиг): {data_js_path}")
    print(f"Файл 2 (html):          {html_path}")
    limit_desc = args.limit if args.limit is not None else "без лимита"
    print(f"Строк встроено: {r['row_count']} (лимит={limit_desc})")
