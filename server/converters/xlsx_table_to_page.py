"""
xlsx_table_to_page.py

Генератор страницы для обычной таблицы (лист без сводной/графика).

Рендер и интерактив (сортировка, Excel-style фильтр по столбцу) отданы
внешней библиотеке table-tools/ (table-tools.js/.css) — она подключается
как файл, не копируется текстом в html (тот же принцип, что и с
pivotgrid-js). Генератор только готовит данные в формате, который эта
библиотека ожидает, и тонкий html-шаблон.

ВАЖНО — потолок в 5000 строк по умолчанию:
Сортировка/фильтр тут работают напрямую по DOM (пересобирают <tbody>
из уже отфильтрованного/отсортированного JS-массива). Это надёжно и
без тормозов примерно до 5–10 тысяч строк; выше — операции на глаз
подвисают (сотни тысяч строк, как в нашем тестовом файле, точно не
годятся без виртуализации, которой здесь нет). Поэтому строки > 5000
по умолчанию просто не подгружаются — если это осознанный выбор,
можно явно передать --limit больше, но тогда таблица предсказуемо
станет менее отзывчивой при сортировке/фильтрации.

Использует xlsx_to_table.py для разбора (не дублирует парсинг).
"""

import html as html_lib
import json
import os
from pathlib import Path

from xlsx_to_table import convert_tables
from xlsx_pivot_to_grid import relhref

DEFAULT_ROW_LIMIT = int(os.environ.get("TABLE_DEFAULT_ROW_LIMIT", 5000))

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title}</title>
  <link rel="stylesheet" href="{table_tools_css_href}">
  <style>
    body {{
      margin: 0;
      font-family: -apple-system, 'Segoe UI', sans-serif;
      background: #f4f5f7;
      color: #1a1a1a;
      padding: 12px;
      box-sizing: border-box;

      height: 100vh;
      overflow: hidden;
      display: flex;
      flex-direction: column;
    }}
    .table-wrap {{
      flex: 1;
      min-height: 0;
      overflow: auto;
      background: #fff;
      border: 1px solid #e0e0e0;
      border-radius: 8px;
    }}
    {nav_css}
  </style>
</head>
<body>

  {nav_html}

  <div id="table-tools-root" class="table-wrap"></div>

  <script src="{table_tools_js_href}"></script>
  <script src="{data_js_href}"></script>
  <script>
    TableTools.init(document.getElementById('table-tools-root'), TABLE_COLUMNS, TABLE_ROWS);
  </script>

  <script>{nav_js}</script>

  <footer style="text-align:center;padding:12px;font-size:12px;color:#999;">
    Лист: {sheet_name} · строк: {row_count}{limit_note}
  </footer>
</body>
</html>
"""

DATA_JS_TEMPLATE = """// Сгенерировано из {source_name}, лист: {sheet_name}, строк: {row_count}
const TABLE_COLUMNS = {columns_json};

const TABLE_ROWS = {rows_json};
"""


# Числовой формат соответствует lang="ru" в шаблоне страницы: запятая как
# десятичный разделитель, неразрывный пробел как разделитель тысяч (принято
# в русской типографике). Разделители разрядов ставятся только для дробных
# чисел — целые (id, Year, Units и т.п.) показываются как есть.
NUMBER_DECIMAL_SEP = ','
NUMBER_THOUSANDS_SEP = '\u00A0'


def _is_numeric(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _format_number(value):
    if isinstance(value, int) or (isinstance(value, float) and value.is_integer()):
        return str(int(value))
    s = f"{value:,.2f}"
    int_part, frac_part = s.split(".")
    int_part = int_part.replace(",", NUMBER_THOUSANDS_SEP)
    return f"{int_part}{NUMBER_DECIMAL_SEP}{frac_part}"


def _fmt_scalar(value):
    if value is None:
        return ""
    if hasattr(value, "strftime"):  # datetime
        if value.hour == 0 and value.minute == 0 and value.second == 0:
            return value.strftime("%d.%m.%Y")
        return value.strftime("%d.%m.%Y %H:%M")
    if isinstance(value, bool):
        return "Да" if value else "Нет"
    if _is_numeric(value):
        return _format_number(value)
    return str(value)


def _wrap(text, bold, italic, underline):
    if bold:
        text = f"<b>{text}</b>"
    if italic:
        text = f"<i>{text}</i>"
    if underline:
        text = f"<u>{text}</u>"
    return text


def _render_cell_html(cell):
    """cell — CellValue|None. Возвращает уже готовый (экранированный) HTML для ячейки."""
    if cell is None or cell.value is None:
        return ""
    if cell.runs is not None:
        return "".join(
            _wrap(html_lib.escape(text), bold, italic, underline)
            for text, bold, italic, underline in cell.runs
        )
    escaped = html_lib.escape(_fmt_scalar(cell.value))
    return _wrap(escaped, cell.bold, cell.italic, cell.underline)


def _infer_column_type(rows, col_index):
    """Тип столбца по первому непустому значению — для сортировки/фильтра
    (число сравнивается численно, дата — по ISO-строке, текст — локально)."""
    for row in rows:
        if col_index >= len(row):
            continue
        cell = row[col_index]
        if cell is None or cell.value is None:
            continue
        v = cell.value
        if hasattr(v, "strftime"):
            return "date"
        if isinstance(v, bool):
            return "bool"
        if isinstance(v, (int, float)):
            return "number"
        return "text"
    return "text"


def _cell_json(cell):
    if cell is None or cell.value is None:
        return {"v": None, "h": ""}
    v = cell.value
    v_json = v.isoformat() if hasattr(v, "isoformat") else v
    return {"v": v_json, "h": _render_cell_html(cell)}


def find_table_tools_pkg():
    """
    Ищет папку table-tools/ (наша собственная библиотека, не npm-пакет):
      1) переменная окружения TABLE_TOOLS_PKG
      2) table-tools/ рядом со скриптом
      3) table-tools/ в текущей рабочей папке
    """
    env_path = os.environ.get("TABLE_TOOLS_PKG")
    if env_path:
        p = Path(env_path)
        if (p / "table-tools.js").exists():
            return p
        raise FileNotFoundError(f"TABLE_TOOLS_PKG указывает на {p}, но там нет table-tools.js")

    candidates = [
        Path(__file__).resolve().parent.parent.parent / "client" / "table-tools",
        Path.cwd() / "client" / "table-tools",
        Path.cwd().parent / "client" / "table-tools",
    ]
    for c in candidates:
        if (c / "table-tools.js").exists():
            return c

    raise FileNotFoundError(
        "Не нашёл папку table-tools/ ни в одном из мест:\n"
        + "\n".join(f"  - {c}" for c in candidates)
        + "\n\nПоложи папку table-tools/ рядом со скриптом (там должны быть "
        "table-tools.js и table-tools.css), либо укажи путь через TABLE_TOOLS_PKG."
    )


def generate(xlsx_path, output_html_path, include_connected_tables=False,
             row_limit=None, nav=("", "", ""), target_sheet_name=None, precomputed_results=None):
    nav_css, nav_html, nav_js = nav

    effective_limit = row_limit if row_limit is not None else DEFAULT_ROW_LIMIT
    if row_limit is not None and row_limit > DEFAULT_ROW_LIMIT:
        print(f"  [!] Таблица: запрошено {row_limit} строк, это выше рекомендованного "
              f"потолка в {DEFAULT_ROW_LIMIT} — сортировка/фильтр в браузере могут тормозить.")

    if precomputed_results is not None:
        results = precomputed_results
    else:
        results = convert_tables(xlsx_path, include_connected_tables=include_connected_tables,
                                  row_limit=effective_limit)
    if not results:
        raise ValueError("Не найдено ни одной самостоятельной таблицы")

    if target_sheet_name is not None:
        matches = [x for x in results if x["sheet_name"] == target_sheet_name]
        if not matches:
            raise ValueError(f"Лист '{target_sheet_name}' не найден среди таблиц")
        r = matches[0]
    else:
        # Один лист — одна таблица по правилу проекта, несколько таблиц
        # сразу тут в принципе не встречается — берём первую найденную.
        r = results[0]

    columns = [
        {
            "key": f"c{i}",
            "html": _render_cell_html(header_cell),
            "type": _infer_column_type(r["rows"], i),
        }
        for i, header_cell in enumerate(r["headers"])
    ]

    rows_json_ready = [
        {col["key"]: _cell_json(row[i] if i < len(row) else None) for i, col in enumerate(columns)}
        for row in r["rows"]
    ]

    html_path = Path(output_html_path).resolve()
    out_dir = html_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    data_js_path = html_path.with_suffix("").with_suffix(".data.js")

    tools_pkg = find_table_tools_pkg()

    data_js = DATA_JS_TEMPLATE.format(
        source_name=Path(xlsx_path).name,
        sheet_name=r["sheet_name"],
        row_count=r["row_count"],
        columns_json=json.dumps(columns, ensure_ascii=False).replace("</", "<\\/"),
        rows_json=json.dumps(rows_json_ready, ensure_ascii=False).replace("</", "<\\/"),
    )
    data_js_path.write_text(data_js, encoding="utf-8")

    limit_note = ""
    if r["row_count"] >= effective_limit - 1:
        limit_note = f" (показаны первые {effective_limit}, в исходнике может быть больше)"

    page = HTML_TEMPLATE.format(
        title="Таблица: " + Path(xlsx_path).name,
        table_tools_css_href=relhref(tools_pkg / "table-tools.css", out_dir),
        table_tools_js_href=relhref(tools_pkg / "table-tools.js", out_dir),
        data_js_href=data_js_path.name,
        sheet_name=r["sheet_name"],
        row_count=r["row_count"],
        limit_note=limit_note,
        nav_css=nav_css,
        nav_html=nav_html,
        nav_js=nav_js,
    )
    html_path.write_text(page, encoding="utf-8")
    return r, html_path


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="xlsx (таблица) -> html-страница c сортировкой/фильтром")
    parser.add_argument("xlsx_path", nargs="?", default="test_data.xlsx")
    parser.add_argument("output_html", nargs="?", default="table.html")
    parser.add_argument("--limit", type=int, default=None,
                         help=f"По умолчанию {DEFAULT_ROW_LIMIT} (см. комментарий в начале файла)")
    parser.add_argument("--include-connected", action="store_true")
    args = parser.parse_args()

    r, html_path = generate(
        args.xlsx_path, args.output_html,
        include_connected_tables=args.include_connected,
        row_limit=args.limit,
    )
    print(f"Файл: {html_path}")
    print(f"Лист: {r['sheet_name']}, строк: {r['row_count']}")
