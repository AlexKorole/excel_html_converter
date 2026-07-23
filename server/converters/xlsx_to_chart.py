"""
xlsx_to_chart.py

Парсер №3 из трёх ("для графика"). Находит график на листе (через цепочку
sheet -> _rels -> drawingN.xml -> _rels -> chartN.xml) и читает из
chart.xml УЖЕ ЗАКЭШИРОВАННЫЕ данные (strCache/numCache/multiLvlStrCache).

Не важно, построен график по сводной (PivotChart, есть <c:pivotSource>)
или прямо по диапазону обычной таблицы — в обоих случаях реальные числа
для отрисовки лежат в кэше внутри самого chart.xml, читаем одинаково,
без обращения к живой сводной/листу.

Поддерживаются: lineChart, barChart (оба типа — сам график и
столбиковая диаграмма). Без дополнительного функционала — статичный
снимок того, что уже посчитано и сохранено в файле.

Использует xlsx_to_table.get_sheet_map и xlsx_to_pivotgrid.read_rels/qn/local
(не дублирует).
"""

import zipfile
import xml.etree.ElementTree as ET

from xlsx_to_pivotgrid import read_rels, local
from xlsx_to_table import get_sheet_map

NS_CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
NS_DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/main'


def qnc(tag):
    return f'{{{NS_CHART}}}{tag}'


def qna(tag):
    return f'{{{NS_DRAWING}}}{tag}'


CHART_TYPE_TAGS = {
    'lineChart': 'line',
    'line3DChart': 'line',
    'barChart': 'bar',
    'bar3DChart': 'bar',
    'pieChart': 'pie',
    'pie3DChart': 'pie',
}


def find_chart_sheets(z):
    """[(имя_листа, путь_к_chartN.xml), ...] — для листов, где есть drawing->chart.
    Один лист может ссылаться на несколько графиков — собираем все."""
    result = []
    for name, sheet_path in get_sheet_map(z):
        sheet_rels = read_rels(z, sheet_path)
        drawing_path = next((t for t in sheet_rels.values() if 'drawings/' in t), None)
        if drawing_path is None:
            continue
        drawing_rels = read_rels(z, drawing_path)
        chart_paths = sorted(t for t in drawing_rels.values() if 'charts/' in t)
        for chart_path in chart_paths:
            result.append((name, chart_path))
    return result


def _extract_sparse_pts(container_el):
    """<c:pt idx=".."><c:v>..</c:v></c:pt>... -> {idx: значение}"""
    result = {}
    for pt in container_el.findall(qnc('pt')):
        idx = int(pt.get('idx'))
        v = pt.find(qnc('v'))
        result[idx] = v.text if v is not None else None
    return result


def _extract_flat_cache(cache_el):
    """Одноуровневый strCache/numCache -> список значений по порядку (с учётом пропусков)."""
    sparse = _extract_sparse_pts(cache_el)
    count_el = cache_el.find(qnc('ptCount'))
    count = int(count_el.get('val')) if count_el is not None else (max(sparse) + 1 if sparse else 0)
    return [sparse.get(i) for i in range(count)]


def _extract_categories(cat_el):
    """
    Возвращает список подписей категорий (по одной строке на точку).
    Многоуровневые категории (multiLvlStrCache) схлопываются в одну подпись:
    "самый детальный уровень (остальные уровни через запятую)".
    """
    if cat_el is None:
        return []

    str_ref = cat_el.find(qnc('strRef'))
    if str_ref is not None:
        cache = str_ref.find(qnc('strCache'))
        return _extract_flat_cache(cache) if cache is not None else []

    num_ref = cat_el.find(qnc('numRef'))
    if num_ref is not None:
        cache = num_ref.find(qnc('numCache'))
        return _extract_flat_cache(cache) if cache is not None else []

    multi_ref = cat_el.find(qnc('multiLvlStrRef'))
    if multi_ref is not None:
        cache = multi_ref.find(qnc('multiLvlStrCache'))
        if cache is None:
            return []
        count_el = cache.find(qnc('ptCount'))
        count = int(count_el.get('val')) if count_el is not None else 0

        # Уровни хранятся от самого детального (0) к самому общему; менее
        # детальные уровни перечисляют значение, только когда оно меняется —
        # протягиваем последнее известное значение вперёд по индексам.
        filled_levels = []
        for lvl in cache.findall(qnc('lvl')):
            sparse = _extract_sparse_pts(lvl)
            filled, last = [], ''
            for i in range(count):
                if sparse.get(i) is not None:
                    last = sparse[i]
                filled.append(last)
            filled_levels.append(filled)

        labels = []
        for i in range(count):
            parts = [lvl[i] for lvl in filled_levels]
            if len(parts) > 1:
                extra = ', '.join(p for p in parts[1:] if p)
                labels.append(f'{parts[0]} ({extra})' if extra else parts[0])
            else:
                labels.append(parts[0] if parts else '')
        return labels

    return []


def _extract_series_name(ser_el):
    tx = ser_el.find(qnc('tx'))
    if tx is None:
        return None
    v = tx.find(qnc('v'))
    if v is not None:
        return v.text
    str_ref = tx.find(qnc('strRef'))
    if str_ref is not None:
        cache = str_ref.find(qnc('strCache'))
        if cache is not None:
            pt = cache.find(qnc('pt'))
            if pt is not None:
                v2 = pt.find(qnc('v'))
                return v2.text if v2 is not None else None
    return None


def _resolve_theme_colors(z):
    """{'accent1': '#RRGGBB', 'dk1': '#...', ...} — реальные цвета темы книги."""
    if 'xl/_rels/workbook.xml.rels' not in z.namelist():
        return {}
    wb_rels_root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    theme_path = None
    for rel in wb_rels_root:
        if rel.get('Type', '').endswith('/theme'):
            target = rel.get('Target').lstrip('/')
            theme_path = target if target.startswith('xl/') else f'xl/{target}'
            break
    if theme_path is None or theme_path not in z.namelist():
        return {}

    troot = ET.fromstring(z.read(theme_path))
    theme_elements_el = troot.find(qna('themeElements'))
    scheme_el = theme_elements_el.find(qna('clrScheme')) if theme_elements_el is not None else None
    if scheme_el is None:
        return {}

    colors = {}
    for child in scheme_el:
        name = local(child.tag)
        srgb = child.find(qna('srgbClr'))
        sys_clr = child.find(qna('sysClr'))
        if srgb is not None:
            colors[name] = '#' + srgb.get('val')
        elif sys_clr is not None:
            colors[name] = '#' + sys_clr.get('lastClr', '000000')
    return colors


def _find_solid_fill(sp_pr_el):
    """
    Ищет <a:solidFill>. Приоритет — ПРЯМАЯ заливка (это и есть цвет
    сектора/бара; у pie/bar рядом почти всегда есть ЕЩЁ одна заливка
    внутри <a:ln> — это просто цвет тонкой обводки между элементами,
    не основной цвет). <a:ln> проверяем только как фолбэк — это тот
    случай, когда своей заливки нет вообще (line-график: там "цвет" —
    это буквально цвет самой линии, других заливок в spPr не бывает).
    """
    direct = sp_pr_el.find(qna('solidFill'))
    if direct is not None:
        return direct
    ln = sp_pr_el.find(qna('ln'))
    if ln is not None:
        return ln.find(qna('solidFill'))
    return None


def _resolve_fill_color(sp_pr_el, theme_colors):
    """Общая часть: <a:solidFill> (srgbClr или schemeClr) -> '#RRGGBB' | None."""
    if sp_pr_el is None:
        return None
    solid_fill = _find_solid_fill(sp_pr_el)
    if solid_fill is None:
        return None
    srgb = solid_fill.find(qna('srgbClr'))
    if srgb is not None:
        return '#' + srgb.get('val')
    scheme_clr = solid_fill.find(qna('schemeClr'))
    if scheme_clr is not None:
        return theme_colors.get(scheme_clr.get('val'))
    return None


def _extract_series_color(ser_el, theme_colors):
    return _resolve_fill_color(ser_el.find(qnc('spPr')), theme_colors)


def _extract_data_point_colors(ser_el, theme_colors):
    """
    Только для pie/pie3D — там цвет обычно не на всю серию (она одна),
    а на каждый КУСОЧЕК отдельно: <c:dPt><c:idx val="N"/><c:spPr>...
    Возвращает {индекс_кусочка: '#RRGGBB'}.
    """
    colors = {}
    for dpt in ser_el.findall(qnc('dPt')):
        idx_el = dpt.find(qnc('idx'))
        if idx_el is None:
            continue
        color = _resolve_fill_color(dpt.find(qnc('spPr')), theme_colors)
        if color is not None:
            colors[int(idx_el.get('val'))] = color
    return colors


def _extract_axis_bounds(root):
    """Явные min/max оси значений, если заданы в файле (иначе None -> авто)."""
    val_ax = root.find(qnc('chart'))
    if val_ax is None:
        return None, None
    val_ax = val_ax.find(qnc('plotArea'))
    if val_ax is None:
        return None, None
    for ax in val_ax.findall(qnc('valAx')):
        scaling = ax.find(qnc('scaling'))
        if scaling is None:
            continue
        min_el = scaling.find(qnc('min'))
        max_el = scaling.find(qnc('max'))
        axis_min = float(min_el.get('val')) if min_el is not None else None
        axis_max = float(max_el.get('val')) if max_el is not None else None
        if axis_min is not None or axis_max is not None:
            return axis_min, axis_max
    return None, None
    chart_el = root.find(qnc('chart'))
    if chart_el is None:
        return None
    title_el = chart_el.find(qnc('title'))
    if title_el is None:
        return None
    texts = [t.text for t in title_el.iter(qna('t')) if t.text]
    return ''.join(texts) if texts else None


def _extract_title(root):
    chart_el = root.find(qnc('chart'))
    if chart_el is None:
        return None
    title_el = chart_el.find(qnc('title'))
    if title_el is None:
        return None
    texts = [t.text for t in title_el.iter(qna('t')) if t.text]
    return ''.join(texts) if texts else None


def _detect_chart_type_element(root):
    """Ищет первый известный тип графика (lineChart/barChart) где угодно
    в дереве и возвращает (тип, barDir, сам_элемент)."""
    for el in root.iter():
        tag = local(el.tag)
        if tag in CHART_TYPE_TAGS:
            bar_dir = None
            if tag.startswith('bar'):
                bd = el.find(qnc('barDir'))
                bar_dir = bd.get('val') if bd is not None else 'col'
            return CHART_TYPE_TAGS[tag], bar_dir, el
    return None, None, None


def parse_chart(chart_xml_bytes, theme_colors=None):
    theme_colors = theme_colors or {}
    root = ET.fromstring(chart_xml_bytes)
    chart_type, bar_dir, type_el = _detect_chart_type_element(root)
    if chart_type is None:
        return None  # неизвестный/неподдерживаемый тип — вызывающий код решает, что делать

    series = []
    for ser_el in type_el.findall(qnc('ser')):
        name = _extract_series_name(ser_el)
        categories = _extract_categories(ser_el.find(qnc('cat')))
        color = _extract_series_color(ser_el, theme_colors)

        values = []
        val_el = ser_el.find(qnc('val'))
        if val_el is not None:
            num_ref = val_el.find(qnc('numRef'))
            if num_ref is not None:
                cache = num_ref.find(qnc('numCache'))
                if cache is not None:
                    raw = _extract_flat_cache(cache)
                    values = [float(v) if v is not None else None for v in raw]

        series.append({'name': name, 'categories': categories, 'values': values, 'color': color})

    categories = next((s['categories'] for s in series if s['categories']), [])
    chart_el = root.find(qnc('chart'))
    axis_min, axis_max = _extract_axis_bounds(root)

    # У pie/pie3D цвет обычно задан не на серию (она одна), а на каждый
    # кусочек отдельно (<c:dPt>). Берём из первой серии — у pie их и так одна.
    point_colors = {}
    if chart_type == 'pie' and type_el.find(qnc('ser')) is not None:
        point_colors = _extract_data_point_colors(type_el.find(qnc('ser')), theme_colors)

    return {
        'chart_type': chart_type,
        'bar_dir': bar_dir,
        'title': _extract_title(root),
        'has_legend': chart_el is not None and chart_el.find(qnc('legend')) is not None,
        'categories': categories,
        'point_colors': point_colors,
        'axis_min': axis_min,
        'axis_max': axis_max,
        'series': [{'name': s['name'], 'values': s['values'], 'color': s['color']} for s in series],
    }


def convert_charts(xlsx_path):
    with zipfile.ZipFile(xlsx_path) as z:
        theme_colors = _resolve_theme_colors(z)
        results = []
        for sheet_name, chart_path in find_chart_sheets(z):
            try:
                chart = parse_chart(z.read(chart_path), theme_colors)
            except Exception as e:
                print(f'  [!] {sheet_name}: не удалось разобрать график ({e}), пропускаю')
                continue
            if chart is None:
                print(f'  [!] {sheet_name}: неподдерживаемый тип графика, пропускаю')
                continue
            chart['sheet_name'] = sheet_name
            results.append(chart)
        return results


if __name__ == '__main__':
    import sys
    import json

    path = sys.argv[1] if len(sys.argv) > 1 else 'test_data.xlsx'
    results = convert_charts(path)

    if not results:
        print('Графиков не найдено (или тип не поддерживается).')
    for r in results:
        print(f"\n=== Лист: {r['sheet_name']} ===")
        print('Тип:', r['chart_type'], '| barDir:', r['bar_dir'])
        print('Заголовок:', r['title'])
        print('Легенда:', r['has_legend'])
        print('Категорий:', len(r['categories']), '->', r['categories'][:10])
        print('Серий:', len(r['series']))
        for s in r['series']:
            print(f"  {s['name']}: {s['values'][:5]}{'...' if len(s['values']) > 5 else ''}")
