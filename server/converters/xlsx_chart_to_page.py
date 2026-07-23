"""
xlsx_chart_to_page.py

Генератор страницы для графика. Рендер отдан внешней библиотеке
chart-tools/ (chart-tools.js/.css) — подключается как файл, не
копируется текстом (тот же принцип, что pivotgrid-js и table-tools).

Использует xlsx_to_chart.py для разбора (не дублирует парсинг).
"""

import json
from pathlib import Path

from xlsx_to_chart import convert_charts
from report_messages import rt
from xlsx_pivot_to_grid import relhref

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title}</title>
  <link rel="stylesheet" href="{chart_tools_css_href}">
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
    #chart-root {{
      flex: 1;
      min-height: 0;
    }}
    {nav_css}
  </style>
</head>
<body>

  {nav_html}

  <div id="chart-root" class="chart-tools-root"></div>

  <script src="{chart_tools_js_href}"></script>
  <script src="{chart_tools_pie_js_href}"></script>
  <script src="{data_js_href}"></script>
  <script>
    ChartTools.init(document.getElementById('chart-root'), CHART_CONFIG);
  </script>

  <script>{nav_js}</script>

  <footer style="text-align:center;padding:12px;font-size:12px;color:#999;">
    {sheet_label}: {sheet_name} · {type_label}: {chart_type} · {series_label}: {series_count}
  </footer>
</body>
</html>
"""

DATA_JS_TEMPLATE = """// Сгенерировано из {source_name}, лист: {sheet_name}, тип: {chart_type}
const CHART_CONFIG = {config_json};
"""


def find_chart_tools_pkg():
    """
    Ищет папку chart-tools/ (собственная библиотека, не npm-пакет):
      1) переменная окружения CHART_TOOLS_PKG
      2) chart-tools/ рядом со скриптом
      3) chart-tools/ в текущей рабочей папке
    """
    import os

    env_path = os.environ.get("CHART_TOOLS_PKG")
    if env_path:
        p = Path(env_path)
        if (p / "chart-tools.js").exists():
            return p
        raise FileNotFoundError(f"CHART_TOOLS_PKG указывает на {p}, но там нет chart-tools.js")

    candidates = [
        Path(__file__).resolve().parent.parent.parent / "client" / "chart-tools",
        Path.cwd() / "client" / "chart-tools",
        Path.cwd().parent / "client" / "chart-tools",
    ]
    for c in candidates:
        if (c / "chart-tools.js").exists():
            return c

    raise FileNotFoundError(
        "Не нашёл папку chart-tools/ ни в одном из мест:\n"
        + "\n".join(f"  - {c}" for c in candidates)
        + "\n\nПоложи папку chart-tools/ рядом со скриптом (там должны быть "
        "chart-tools.js и chart-tools.css), либо укажи путь через CHART_TOOLS_PKG."
    )


def generate(xlsx_path, output_html_path, nav=("", "", ""),
             target_sheet_name=None, occurrence_index=0, precomputed_results=None):
    nav_css, nav_html, nav_js = nav

    results = precomputed_results if precomputed_results is not None else convert_charts(xlsx_path)
    if not results:
        raise ValueError("Не найдено ни одного поддерживаемого графика (line/bar)")

    if target_sheet_name is not None:
        matches = [x for x in results if x["sheet_name"] == target_sheet_name]
        if occurrence_index >= len(matches):
            raise ValueError(
                f"На листе '{target_sheet_name}' не найден график №{occurrence_index + 1}"
            )
        r = matches[occurrence_index]
    else:
        # Один график по правилу проекта — берём первый найденный.
        r = results[0]

    html_path = Path(output_html_path).resolve()
    out_dir = html_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    data_js_path = html_path.with_suffix("").with_suffix(".data.js")

    tools_pkg = find_chart_tools_pkg()

    config = {
        "chart_type": r["chart_type"],
        "bar_dir": r["bar_dir"],
        "title": r["title"],
        "has_legend": r["has_legend"],
        "categories": r["categories"],
        "point_colors": r.get("point_colors", {}),
        "axis_min": r["axis_min"],
        "axis_max": r["axis_max"],
        "series": r["series"],
    }

    data_js = DATA_JS_TEMPLATE.format(
        source_name=Path(xlsx_path).name,
        sheet_name=r["sheet_name"],
        chart_type=r["chart_type"],
        config_json=json.dumps(config, ensure_ascii=False).replace("</", "<\\/"),
    )
    data_js_path.write_text(data_js, encoding="utf-8")

    page = HTML_TEMPLATE.format(
        title="График: " + Path(xlsx_path).name,
        chart_tools_css_href=relhref(tools_pkg / "chart-tools.css", out_dir),
        chart_tools_js_href=relhref(tools_pkg / "chart-tools.js", out_dir),
        chart_tools_pie_js_href=relhref(tools_pkg / "chart-tools-pie.js", out_dir),
        data_js_href=data_js_path.name,
        sheet_name=r["sheet_name"],
        chart_type=r["chart_type"],
        series_count=len(r["series"]),
        sheet_label=rt("sheet_label"),
        type_label=rt("type_label"),
        series_label=rt("series_label"),
        nav_css=nav_css,
        nav_html=nav_html,
        nav_js=nav_js,
    )
    html_path.write_text(page, encoding="utf-8")
    return r, html_path


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="xlsx (график) -> html-страница")
    parser.add_argument("xlsx_path", nargs="?", default="test_data.xlsx")
    parser.add_argument("output_html", nargs="?", default="chart.html")
    args = parser.parse_args()

    r, html_path = generate(args.xlsx_path, args.output_html)
    print(f"Файл: {html_path}")
    print(f"Лист: {r['sheet_name']}, тип: {r['chart_type']}, серий: {len(r['series'])}")
