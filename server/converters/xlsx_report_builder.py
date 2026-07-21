"""
xlsx_report_builder.py

Собирает многостраничный отчёт из xlsx: по одной html-странице на каждую
единицу контента (таблица / сводная / график), с общим бургер-меню для
навигации между страницами (report_nav.py).

Правила именования и порядка:
  - Название пункта меню = имя листа в Excel (как есть).
  - Порядок пунктов меню = порядок листов в книге Excel.
  - Один лист = один элемент (таблица ИЛИ сводная ИЛИ график) — но
    сводных или графиков на одном листе может быть несколько (таблиц —
    нет, по правилам xlsx_to_table.py). Если на одном листе несколько
    сводных/графиков — к имени добавляется _2, _3 и т.д.

Никакого JS-роутинга/аккордеона — обычные статичные страницы со ссылками
друг на друга, плюс index.html как точка входа (открывает первую
доступную страницу).
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import xlsx_pivot_to_grid
import xlsx_table_to_page
import xlsx_chart_to_page
from xlsx_to_table import convert_tables
from xlsx_to_pivotgrid import convert as convert_pivot
from xlsx_to_chart import convert_charts
from report_nav import render_nav

NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

INDEX_TEMPLATE = """<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="refresh" content="0; url={first_file}">
</head>
<body style="font-family: -apple-system, 'Segoe UI', sans-serif; padding: 24px;">
  Открываю отчёт... <a href="{first_file}">{first_label}</a>
</body>
</html>
"""


def _get_sheet_order_map(xlsx_path):
    """{имя_листа: порядковый номер в книге} — для сортировки пунктов меню."""
    with zipfile.ZipFile(xlsx_path) as z:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        sheets_el = wb.find(f'{{{NS_MAIN}}}sheets')
        names = [s.get('name') for s in sheets_el.findall(f'{{{NS_MAIN}}}sheet')]
    return {name: i for i, name in enumerate(names)}


KIND_FILE_PREFIX = {'table': 'table', 'pivot': 'grid', 'chart': 'graph'}


def _assign_labels_and_filenames(items):
    """
    items: [(kind, sheet_name), ...] уже в нужном порядке.
    Возвращает [(kind, sheet_name, label, filename), ...]:
      - label — имя листа Excel, с суффиксом _2, _3... если на одном листе
        несколько сводных/графиков (метка в меню всегда осмысленная);
      - filename — простое сквозное имя (table1.html, grid1.html, graph1.html...),
        никак не завязанное на имя листа — чтобы длинное/неудобное имя листа
        не превращалось в громоздкий путь на диске.
    """
    label_seen = {}
    kind_counter = {}
    result = []
    for kind, sheet_name in items:
        key = (kind, sheet_name)
        label_seen[key] = label_seen.get(key, 0) + 1
        n = label_seen[key]
        suffix = '' if n == 1 else f'_{n}'
        label = f'{sheet_name}{suffix}'

        kind_counter[kind] = kind_counter.get(kind, 0) + 1
        filename = f'{KIND_FILE_PREFIX[kind]}{kind_counter[kind]}.html'

        result.append((kind, sheet_name, label, filename))
    return result


def build_report(xlsx_path, output_dir, include_connected_tables=False,
                  pivot_row_limit=None, table_row_limit=None):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    sheet_order = _get_sheet_order_map(xlsx_path)

    # Данные разбираем РОВНО ОДИН РАЗ на файл целиком (не по разу на
    # каждую вкладку) — иначе при нескольких сводных/графиках в одном
    # файле пришлось бы пересчитывать одно и то же по нескольку раз
    # (convert()/convert_tables()/convert_charts() всегда разбирают ВСЕ
    # найденные сводные/таблицы/графики за один проход).
    # Подставляем дефолт САМИ, до вызова convert_tables()/convert() —
    # иначе None означало бы "без лимита вообще" (эта подстановка раньше
    # жила только внутри generate(), которую мы теперь не вызываем первой).
    effective_table_limit = table_row_limit if table_row_limit is not None else xlsx_table_to_page.DEFAULT_ROW_LIMIT
    effective_pivot_limit = pivot_row_limit if pivot_row_limit is not None else xlsx_pivot_to_grid.DEFAULT_ROW_LIMIT

    try:
        table_results = convert_tables(xlsx_path, include_connected_tables=include_connected_tables,
                                        row_limit=effective_table_limit)
    except ValueError:
        table_results = []
    try:
        pivot_results = convert_pivot(xlsx_path, row_limit=effective_pivot_limit)
    except ValueError:
        pivot_results = []
    try:
        chart_results = convert_charts(xlsx_path)
    except ValueError:
        chart_results = []

    items = []
    for r in table_results:
        items.append(('table', r['sheet_name']))
    for r in pivot_results:
        items.append(('pivot', r['sheet_name']))
    for r in chart_results:
        items.append(('chart', r['sheet_name']))

    # Сортировка по порядковому номеру листа в книге. sort стабильна —
    # относительный порядок нескольких элементов на одном листе (уже
    # выставленный парсерами) сохраняется.
    items.sort(key=lambda it: sheet_order.get(it[1], 10**9))

    labeled = _assign_labels_and_filenames(items)
    if not labeled:
        raise ValueError("В файле не нашлось ни таблицы, ни сводной, ни графика для отчёта")

    pages = [(label, filename) for _, _, label, filename in labeled]

    generated = []
    pivot_seen = {}
    chart_seen = {}

    for kind, sheet_name, label, filename in labeled:
        nav = render_nav(pages, filename)
        try:
            if kind == 'table':
                r, path = xlsx_table_to_page.generate(
                    xlsx_path, output_dir / filename,
                    include_connected_tables=include_connected_tables,
                    row_limit=table_row_limit, nav=nav,
                    target_sheet_name=sheet_name,
                    precomputed_results=table_results,
                )
            elif kind == 'pivot':
                idx = pivot_seen.get(sheet_name, 0)
                pivot_seen[sheet_name] = idx + 1
                r, path, data_js_path = xlsx_pivot_to_grid.generate(
                    xlsx_path, output_dir / filename,
                    row_limit=pivot_row_limit, nav=nav,
                    target_sheet_name=sheet_name, occurrence_index=idx,
                    precomputed_results=pivot_results,
                )
            else:  # chart
                idx = chart_seen.get(sheet_name, 0)
                chart_seen[sheet_name] = idx + 1
                r, path = xlsx_chart_to_page.generate(
                    xlsx_path, output_dir / filename, nav=nav,
                    target_sheet_name=sheet_name, occurrence_index=idx,
                    precomputed_results=chart_results,
                )
            generated.append(path)
        except Exception as e:
            print(f"  [!] {label}: не удалось собрать страницу ({e}), пропускаю")

    if generated:
        first_label, first_file = pages[0]
        index_path = output_dir / "index.html"
        index_path.write_text(
            INDEX_TEMPLATE.format(first_file=first_file, first_label=first_label),
            encoding="utf-8",
        )
        generated.append(index_path)

    return generated


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Собрать многостраничный html-отчёт из xlsx")
    parser.add_argument("xlsx_path", nargs="?", default="test_data.xlsx")
    parser.add_argument("output_dir", nargs="?", default="report_output")
    parser.add_argument("--pivot-limit", type=int, default=None,
                         help="Строк для сводной. По умолчанию — без лимита, все строки "
                              "(агрегация, DOM-ограничение не бьёт).")
    parser.add_argument("--table-limit", type=int, default=None,
                         help="Строк для таблицы. По умолчанию 5000 (см. комментарий "
                              "в xlsx_table_to_page.py — сортировка/фильтр в браузере).")
    parser.add_argument("--include-connected", action="store_true",
                         help="Показывать и таблицы, использующиеся как источник сводной")
    args = parser.parse_args()

    files = build_report(
        args.xlsx_path, args.output_dir,
        include_connected_tables=args.include_connected,
        pivot_row_limit=args.pivot_limit,
        table_row_limit=args.table_limit,
    )
    print("Сгенерированы файлы:")
    for f in files:
        print(" ", f)
