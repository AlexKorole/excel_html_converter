"""
xlsx_to_table.py

Парсер №1 из трёх ("для таблицы"). Разбирает обычные листы xlsx
(не сводные, не графики) в структуру:
  {"sheet_name": str, "headers": [...], "rows": [[...], ...], "row_count": int}

Правила (зафиксированы по ходу обсуждения, до кода):
  - Лист пропускается, если на нём есть pivotTable или chart (drawing) —
    это уже "занято" другими парсерами (один лист = одна единица контента).
  - Лист, использующийся как cacheSource какой-то сводной, по умолчанию
    ТОЖЕ пропускается (include_connected_tables=False) — это сырой
    источник, уже посчитанный сводной. Включается явно флагом.
  - Строка выкидывается целиком, только если ВСЕ ячейки в границах
    таблицы (dimension) пустые. Если хоть одна не пустая — строка
    остаётся целиком, включая её собственные пустые ячейки
    (это же правило само по себе покрывает строки итогов).
  - Формулы: берём только <v> (вычисленное значение), <f> не трогаем.
  - Типы ячеек:
      без t         -> число (возможно дата, см. ниже)
      t="s"         -> shared string
      t="str"       -> inline-строка (результат текстовой формулы)
      t="inlineStr" -> <is><t>...</t></is>
      t="b"         -> булево (0/1 -> False/True)
      t="e"         -> 'ERROR' (без разбора конкретного кода ошибки;
                       так же для формул с ошибкой — тот же путь, t="e")
  - Дата определяется ТОЛЬКО если у ячейки явно задан стиль (s есть)
    и numFmtId соответствующего xf != 0 и формат распознан как дата/время
    (встроенный ID из фиксированного набора, либо кастомный формат,
    разобранный по токенам y/m/d/h/s вне кавычек/экранирования/цветовых
    кодов). Иначе — обычное число, как показал бы и сам Excel.
  - Эпоха дат (1900 vs 1904) читается один раз из workbook.xml.

Зависимости: только стандартная библиотека Python + переиспользование
общих XML-хелперов из xlsx_to_pivotgrid.py (не дублируем).
"""

import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta

from xlsx_to_pivotgrid import qn, local, read_rels, NS_REL


# ── Даты и числовые форматы ──────────────────────────────────────────────────

BUILTIN_DATE_FMT_IDS = {14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47}

_QUOTED_RE = re.compile(r'"[^"]*"')
_ESCAPED_RE = re.compile(r'\\.')
_COLOR_BRACKET_RE = re.compile(r'\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\d+)\]', re.IGNORECASE)
_DATE_TOKEN_RE = re.compile(r'[ymdhs]', re.IGNORECASE)

_EXCEL_EPOCH_1900 = datetime(1899, 12, 30)  # с поправкой на баг Excel "1900 — високосный"
_EXCEL_EPOCH_1904 = datetime(1904, 1, 1)


def is_date_format_code(format_code):
    """Формат — дата/время, если после чистки кавычек/экранирования/цветовых
    кодов в нём остаются буквы y/m/d/h/s (иначе — обычное число)."""
    if not format_code or format_code == 'General':
        return False
    cleaned = _QUOTED_RE.sub('', format_code)
    cleaned = _ESCAPED_RE.sub('', cleaned)
    cleaned = _COLOR_BRACKET_RE.sub('', cleaned)
    return bool(_DATE_TOKEN_RE.search(cleaned))


def parse_fonts(z):
    """Список {'bold':bool,'italic':bool,'underline':bool} по индексу шрифта."""
    root = ET.fromstring(z.read('xl/styles.xml'))
    fonts_el = root.find(qn('fonts'))
    fonts = []
    if fonts_el is not None:
        for font in fonts_el.findall(qn('font')):
            fonts.append({
                'bold': font.find(qn('b')) is not None,
                'italic': font.find(qn('i')) is not None,
                'underline': font.find(qn('u')) is not None,
            })
    return fonts


def parse_styles(z):
    """
    numfmt_by_style_index[s] -> numFmtId
    fontid_by_style_index[s] -> fontId
    custom_formats[numFmtId] -> formatCode
    """
    root = ET.fromstring(z.read('xl/styles.xml'))

    custom_formats = {}
    numfmts_el = root.find(qn('numFmts'))
    if numfmts_el is not None:
        for nf in numfmts_el.findall(qn('numFmt')):
            custom_formats[int(nf.get('numFmtId'))] = nf.get('formatCode')

    numfmt_by_style_index = []
    fontid_by_style_index = []
    cellxfs_el = root.find(qn('cellXfs'))
    if cellxfs_el is not None:
        for xf in cellxfs_el.findall(qn('xf')):
            numfmt_by_style_index.append(int(xf.get('numFmtId', '0')))
            fontid_by_style_index.append(int(xf.get('fontId', '0')))

    return numfmt_by_style_index, fontid_by_style_index, custom_formats


def style_is_date(style_index, numfmt_by_style_index, custom_formats):
    if style_index is None or style_index >= len(numfmt_by_style_index):
        return False
    numfmt_id = numfmt_by_style_index[style_index]
    if numfmt_id == 0:
        return False
    if numfmt_id in BUILTIN_DATE_FMT_IDS:
        return True
    if numfmt_id in custom_formats:
        return is_date_format_code(custom_formats[numfmt_id])
    return False


def get_date1904(z):
    root = ET.fromstring(z.read('xl/workbook.xml'))
    props = root.find(qn('workbookPr'))
    return bool(props is not None and props.get('date1904') == '1')


def excel_serial_to_datetime(value, date1904):
    epoch = _EXCEL_EPOCH_1904 if date1904 else _EXCEL_EPOCH_1900
    return epoch + timedelta(days=value)


# ── Строки и ячейки ───────────────────────────────────────────────────────────

def parse_shared_strings(z, fonts):
    """
    Список записей: {'text': str, 'runs': None | [(text,bold,italic,underline), ...]}
    'runs' заполнен, только если строка реально состоит из нескольких
    <r> (значит, форматирование может отличаться внутри одной ячейки —
    например, половина слова жирная). Если строка простая (один <t>
    или один <r>) — 'runs' = None, форматирование берётся из стиля ячейки.
    """
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    root = ET.fromstring(z.read('xl/sharedStrings.xml'))
    result = []
    for si in root.findall(qn('si')):
        t_direct = si.find(qn('t'))
        if t_direct is not None:
            result.append({'text': t_direct.text or '', 'runs': None})
            continue

        r_elements = si.findall(qn('r'))
        if len(r_elements) <= 1:
            text = ''
            if r_elements:
                t = r_elements[0].find(qn('t'))
                text = t.text if t is not None and t.text else ''
            result.append({'text': text, 'runs': None})
            continue

        runs = []
        for r in r_elements:
            t = r.find(qn('t'))
            text = t.text if t is not None and t.text else ''
            rpr = r.find(qn('rPr'))
            bold = rpr is not None and rpr.find(qn('b')) is not None
            italic = rpr is not None and rpr.find(qn('i')) is not None
            underline = rpr is not None and rpr.find(qn('u')) is not None
            runs.append((text, bold, italic, underline))
        result.append({'text': ''.join(r[0] for r in runs), 'runs': runs})
    return result


def col_letters_to_index(ref):
    """'C7' -> 2 (0-based индекс колонки)."""
    letters = ''.join(ch for ch in ref if ch.isalpha())
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
    return idx - 1


def parse_dimension_ref(ref):
    """'A1:O200036' -> (min_col_idx, max_col_idx), 0-based."""
    start, _, end = ref.partition(':')
    if not end:
        end = start
    return col_letters_to_index(start), col_letters_to_index(end)


class CellValue:
    """Значение ячейки + форматирование.
    runs заполнен только для multi-run текстовых ячеек (см. parse_shared_strings) —
    тогда каждый кусок текста рендерится со своим собственным bold/italic/underline,
    а bold/italic/underline самой CellValue в этом случае не используются."""
    __slots__ = ('value', 'bold', 'italic', 'underline', 'runs')

    def __init__(self, value=None, bold=False, italic=False, underline=False, runs=None):
        self.value = value
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.runs = runs


def decode_cell(c_el, shared_strings, numfmt_by_style_index, fontid_by_style_index,
                 fonts, custom_formats, date1904):
    t = c_el.get('t')
    style_index = c_el.get('s')
    style_index = int(style_index) if style_index is not None else 0

    font = fonts[fontid_by_style_index[style_index]] if (
        style_index < len(fontid_by_style_index) and fontid_by_style_index[style_index] < len(fonts)
    ) else {'bold': False, 'italic': False, 'underline': False}

    if t == 'e':
        return CellValue('ERROR', **font)  # тот же путь и для формул с ошибкой (t="e" одинаков)

    if t == 's':
        v_el = c_el.find(qn('v'))
        if v_el is None or v_el.text is None:
            return CellValue(None, **font)
        entry = shared_strings[int(v_el.text)]
        if entry['runs'] is not None:
            return CellValue(entry['text'], runs=entry['runs'])
        return CellValue(entry['text'], **font)

    if t == 'str':
        v_el = c_el.find(qn('v'))
        return CellValue(v_el.text if v_el is not None else None, **font)

    if t == 'inlineStr':
        is_el = c_el.find(qn('is'))
        if is_el is None:
            return CellValue(None, **font)
        t_el = is_el.find(qn('t'))
        return CellValue(t_el.text if t_el is not None else None, **font)

    if t == 'b':
        v_el = c_el.find(qn('v'))
        if v_el is None or v_el.text is None:
            return CellValue(None, **font)
        return CellValue(v_el.text == '1', **font)

    # без t -> число, возможно дата
    v_el = c_el.find(qn('v'))
    if v_el is None or v_el.text is None:
        return CellValue(None, **font)
    num = float(v_el.text)

    if style_is_date(style_index, numfmt_by_style_index, custom_formats):
        return CellValue(excel_serial_to_datetime(num, date1904), **font)

    return CellValue(num, **font)


# ── Определение "листов-таблиц" ──────────────────────────────────────────────

def get_sheet_map(z):
    """workbook.xml -> [(имя_листа, путь_к_sheetN.xml), ...] в порядке книги."""
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    wb_rels = read_rels(z, 'xl/workbook.xml')
    sheets_el = wb.find(qn('sheets'))
    result = []
    for sheet in sheets_el.findall(qn('sheet')):
        name = sheet.get('name')
        rid = sheet.get(f'{{{NS_REL}}}id')
        path = wb_rels.get(rid)
        if path:
            result.append((name, path))
    return result


def get_sheets_with_pivot_or_chart(z, sheet_map):
    """Имена листов, у которых в собственном _rels есть pivotTable или drawing —
    такие листы уже обрабатываются парсером сводной/графика, не таблицей."""
    claimed = set()
    for name, path in sheet_map:
        rels = read_rels(z, path)
        for target in rels.values():
            if 'pivotTables/' in target or 'drawings/' in target:
                claimed.add(name)
                break
    return claimed


def get_cache_source_sheets(z):
    """Имена листов, использующихся как cacheSource какой-либо сводной."""
    sources = set()
    for name in z.namelist():
        if name.startswith('xl/pivotCache/pivotCacheDefinition') and name.endswith('.xml'):
            root = ET.fromstring(z.read(name))
            cs = root.find(qn('cacheSource'))
            if cs is not None:
                ws = cs.find(qn('worksheetSource'))
                if ws is not None and ws.get('sheet'):
                    sources.add(ws.get('sheet'))
    return sources


def find_table_sheets(z, include_connected_tables=False):
    """Список (имя, путь) листов-кандидатов на роль самостоятельной таблицы."""
    sheet_map = get_sheet_map(z)
    claimed = get_sheets_with_pivot_or_chart(z, sheet_map)
    cache_sources = set() if include_connected_tables else get_cache_source_sheets(z)

    return [
        (name, path) for name, path in sheet_map
        if name not in claimed and name not in cache_sources
    ]


# ── Разбор содержимого листа (стримингом, листы могут быть огромными) ───────

def stream_parse_table(z, sheet_path, shared_strings, numfmt_by_style_index,
                        fontid_by_style_index, fonts, custom_formats, date1904, row_limit=None):
    col_min, col_max = None, None
    rows_out = []
    count = 0

    with z.open(sheet_path) as fh:
        for event, el in ET.iterparse(fh, events=('end',)):
            tag = local(el.tag)

            if tag == 'dimension' and col_min is None:
                ref = el.get('ref')
                if ref:
                    col_min, col_max = parse_dimension_ref(ref)
                el.clear()
                continue

            if tag != 'row':
                continue

            # Строка, скрытая в Excel (автофильтром или вручную, Формат ->
            # Скрыть строки) — пользователь её не видит, мы тоже не показываем.
            if el.get('hidden') == '1':
                el.clear()
                continue

            cells_by_col = {}
            for c_el in el.findall(qn('c')):
                ref = c_el.get('r')
                if not ref:
                    continue
                col_idx = col_letters_to_index(ref)
                cells_by_col[col_idx] = decode_cell(
                    c_el, shared_strings, numfmt_by_style_index,
                    fontid_by_style_index, fonts, custom_formats, date1904,
                )

            if col_min is not None:
                lo, hi = col_min, col_max
            elif cells_by_col:
                lo, hi = min(cells_by_col), max(cells_by_col)
            else:
                lo, hi = 0, -1

            row_values = [cells_by_col.get(i) for i in range(lo, hi + 1)]

            if any(v is not None and v.value is not None for v in row_values):
                rows_out.append(row_values)
                count += 1

            el.clear()
            if row_limit is not None and count >= row_limit:
                break

    return rows_out, count


# ── Точка входа ───────────────────────────────────────────────────────────────

def convert_tables(xlsx_path, include_connected_tables=False, row_limit=None):
    with zipfile.ZipFile(xlsx_path) as z:
        fonts = parse_fonts(z)
        shared_strings = parse_shared_strings(z, fonts)
        numfmt_by_style_index, fontid_by_style_index, custom_formats = parse_styles(z)
        date1904 = get_date1904(z)

        candidates = find_table_sheets(z, include_connected_tables=include_connected_tables)

        results = []
        for name, path in candidates:
            rows, count = stream_parse_table(
                z, path, shared_strings, numfmt_by_style_index,
                fontid_by_style_index, fonts, custom_formats, date1904,
                row_limit=row_limit,
            )
            headers = rows[0] if rows else []
            data_rows = rows[1:] if rows else []
            results.append({
                'sheet_name': name,
                'headers': headers,
                'rows': data_rows,
                'row_count': len(data_rows),
            })
        return results


if __name__ == '__main__':
    import argparse
    import time

    parser = argparse.ArgumentParser(description='Разбор обычных листов (таблиц) из xlsx')
    parser.add_argument('xlsx_path', nargs='?', default='test_data.xlsx')
    parser.add_argument('--limit', type=int, default=10, help='Строк на лист для просмотра (по умолчанию 10)')
    parser.add_argument('--include-connected', action='store_true',
                         help='Показывать и листы, использующиеся как источник сводной')
    args = parser.parse_args()

    t0 = time.time()
    results = convert_tables(args.xlsx_path, include_connected_tables=args.include_connected, row_limit=args.limit)
    elapsed = time.time() - t0

    def _preview(cell):
        if cell is None:
            return None
        flags = ''.join(f for f, on in (('b', cell.bold), ('i', cell.italic), ('u', cell.underline)) if on)
        return f'{cell.value!r}[{flags}]' if flags or cell.runs else repr(cell.value)

    if not results:
        print('Самостоятельных таблиц не найдено (все листы — источники сводных/графики/сводные).')
    for r in results:
        print(f"\n=== Лист: {r['sheet_name']} ===")
        print('Заголовки:', [_preview(h) for h in r['headers']])
        print(f"Строк (лимит={args.limit}): {r['row_count']}")
        for row in r['rows'][:5]:
            print(' ', [_preview(c) for c in row])

    print(f'\nВремя выполнения: {elapsed:.2f} сек')
