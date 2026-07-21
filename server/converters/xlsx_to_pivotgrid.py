"""
xlsx_to_pivotgrid.py

Разбирает классическую (на листе, не Data Model) сводную таблицу
из .xlsx и превращает её в:
  - data:   список плоских строк (то же, что было на исходном листе)
  - config: dimensions/measures/funcs/fields/rows/columns/measure/func

Ничего не хардкодится по именам полей — всё берётся из структуры
самого файла (pivotCacheDefinition + pivotTableDefinition).

Зависимости: только стандартная библиотека Python.
"""

import zipfile
import posixpath
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import iterparse
from datetime import datetime

NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def qn(tag):
    return f'{{{NS_MAIN}}}{tag}'


def local(tag):
    return tag.split('}')[-1]


def read_rels(z, part_path):
    """rels-файл для part_path -> {rId: target_path_absolute_in_zip}"""
    base_dir = posixpath.dirname(part_path)
    rels_path = posixpath.join(base_dir, '_rels', posixpath.basename(part_path) + '.rels')
    if rels_path not in z.namelist():
        return {}
    root = ET.fromstring(z.read(rels_path))
    result = {}
    for rel in root:
        rid = rel.get('Id')
        target = rel.get('Target')
        if target.startswith('/'):
            resolved = target.lstrip('/')
        else:
            resolved = posixpath.normpath(posixpath.join(base_dir, target))
        result[rid] = resolved
    return result


def find_pivot_tables(z):
    """Возвращает список всех pivotTable*.xml, найденных в архиве."""
    return sorted(n for n in z.namelist() if n.startswith('xl/pivotTables/pivotTable') and n.endswith('.xml'))


def _get_sheet_order(z):
    """[(имя_листа, путь_к_sheetN.xml), ...] в порядке листов книги."""
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


def find_pivot_tables_by_sheet(z):
    """
    [(имя_листа, путь_к_pivotTableN.xml), ...] в порядке листов книги.
    Если на одном листе несколько сводных — идут подряд в стабильном порядке.
    """
    result = []
    for sheet_name, sheet_path in _get_sheet_order(z):
        sheet_rels = read_rels(z, sheet_path)
        pivot_paths = sorted(t for t in sheet_rels.values() if 'pivotTables/' in t)
        for p in pivot_paths:
            result.append((sheet_name, p))
    return result


def resolve_cache_definition_for_pivot_table(z, pivot_table_path, workbook_rels, cache_id_to_rid):
    """По cacheId из pivotTableDefinition находит путь к pivotCacheDefinitionN.xml."""
    root = ET.fromstring(z.read(pivot_table_path))
    cache_id = root.get('cacheId')
    rid = cache_id_to_rid.get(cache_id)
    if rid is None:
        raise ValueError(f'Не найден cacheId={cache_id} в workbook.xml pivotCaches')
    return workbook_rels[rid], root


def get_cache_id_map(z):
    """workbook.xml -> {cacheId: rId} из <pivotCaches>"""
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    mapping = {}
    pcs = wb.find(qn('pivotCaches'))
    if pcs is not None:
        for pc in pcs.findall(qn('pivotCache')):
            cache_id = pc.get('cacheId')
            rid = pc.get(f'{{{NS_REL}}}id')
            mapping[cache_id] = rid
    return mapping


def parse_field_group(cache_field_el):
    """
    Если cacheField — это группировка по другому полю (Excel: правой кнопкой
    по дате -> Группировать -> Годы/Кварталы/Месяцы), возвращает:
      {'base_index': int, 'group_by': 'years'|'quarters'|'months', 'start': str, 'end': str, 'items': [str,...]}
    Такие поля НЕ хранят значение в pivotCacheRecords — Excel вычисляет его
    на лету из поля-источника (base_index), поэтому декодировать их надо отдельно.
    """
    fg = cache_field_el.find(qn('fieldGroup'))
    if fg is None:
        return None
    base = fg.get('base')
    range_pr = fg.find(qn('rangePr'))
    if base is None or range_pr is None:
        return None
    group_by = range_pr.get('groupBy')
    group_items_el = fg.find(qn('groupItems'))
    items = []
    if group_items_el is not None:
        for it in group_items_el:
            items.append(it.get('v'))
    return {
        'base_index': int(base),
        'group_by': group_by,
        'start': range_pr.get('startDate'),
        'end': range_pr.get('endDate'),
        'items': items,
    }


def parse_cache_fields(cache_def_xml_bytes):
    """
    Возвращает список dict: {"name": str, "shared_items": [values] | None, "group": dict | None}
    Индекс в списке == индекс поля (используется в pivotTableDefinition как x="N"/fld="N").
    """
    root = ET.fromstring(cache_def_xml_bytes)
    fields_el = root.find(qn('cacheFields'))
    fields = []
    for cf in fields_el.findall(qn('cacheField')):
        name = cf.get('name')
        shared_items = None
        si = cf.find(qn('sharedItems'))
        if si is not None and len(list(si)) > 0:
            values = []
            for item in si:
                tag = local(item.tag)
                if tag == 's':
                    values.append(item.get('v'))
                elif tag == 'n':
                    v = item.get('v')
                    values.append(float(v) if v is not None else None)
                elif tag == 'd':
                    values.append(item.get('v'))
                elif tag == 'b':
                    values.append(item.get('v') == '1')
                elif tag == 'm':
                    values.append(None)
                else:
                    values.append(None)  # group items и т.п. — не поддерживаем в MVP
            shared_items = values

        # fieldGroup смотрим, только если у поля нет собственных значений.
        # У "базового" поля (например Date) fieldGroup может присутствовать
        # ради ВНУТРЕННЕЙ иерархии (base указывает само на себя) — но раз
        # свои значения (sharedItems) есть, это не "чисто производное" поле,
        # и его собственные значения используем как обычно.
        group = parse_field_group(cf) if shared_items is None else None

        fields.append({'name': name, 'shared_items': shared_items, 'group': group})
    return fields


def _parse_iso_datetime(s):
    if s is None:
        return None
    try:
        return datetime.fromisoformat(s)
    except (ValueError, TypeError):
        return None


def compute_group_value(group, base_value):
    """
    Вычисляет подпись группы (например "2023" или "Кв-л2") для значения
    базового поля (обычно даты). Возвращает None, если тип группировки
    не поддержан или дата не распознана.
    """
    items = group['items']
    if not items:
        return None

    base_dt = base_value if isinstance(base_value, datetime) else _parse_iso_datetime(base_value)
    if base_dt is None:
        return None

    start = _parse_iso_datetime(group['start'])
    end = _parse_iso_datetime(group['end'])
    if start is None or end is None:
        return None

    group_by = group['group_by']

    if group_by == 'years':
        if base_dt.year < start.year:
            return items[0]
        if base_dt.year > end.year:
            return items[-1]
        idx = 1 + (base_dt.year - start.year)
        return items[idx] if idx < len(items) else None

    if group_by == 'quarters':
        if base_dt < start:
            return items[0]
        if base_dt > end:
            return items[-1]
        idx = 1 + (base_dt.month - 1) // 3
        return items[idx] if idx < len(items) else None

    if group_by == 'months':
        if base_dt < start:
            return items[0]
        if base_dt > end:
            return items[-1]
        idx = base_dt.month  # items[0]=до диапазона, items[1..12]=месяцы
        return items[idx] if idx < len(items) else None

    return None  # неизвестный groupBy (days/seconds и т.п.) — не поддержано в MVP


def decode_record_value(el, cache_fields, field_index):
    tag = local(el.tag)
    if tag == 'x':
        idx = int(el.get('v'))
        shared = cache_fields[field_index]['shared_items']
        return shared[idx] if shared is not None and idx < len(shared) else None
    elif tag == 'n':
        v = el.get('v')
        return float(v) if v is not None else None
    elif tag == 's':
        return el.get('v')
    elif tag == 'd':
        return el.get('v')
    elif tag == 'b':
        return el.get('v') == '1'
    elif tag == 'm':
        return None
    else:
        return None


def stream_decode_records(z, records_path, cache_fields, limit=None, row_filter=None):
    """
    Стримингом (iterparse) разбирает pivotCacheRecords*.xml — важно,
    т.к. файл может весить сотни МБ при 1 млн строк.

    row_filter — необязательная функция row_dict -> row_dict|None.
    Если задана и лимит тоже задан, ЛИМИТ СЧИТАЕТСЯ ПО СТРОКАМ, ПРОШЕДШИМ
    ФИЛЬТР, а не по количеству прочитанных сырых записей — иначе при
    активном фильтре сводной можно было бы прочитать только первые N сырых
    записей файлового порядка, ни разу не отфильтровав их, и получить
    неполный/неверный результат, если реально подходящие записи не
    сосредоточены в начале файла.

    Возвращает (rows, kept_count, raw_count).
    """
    field_names = [f['name'] for f in cache_fields]
    rows = []
    raw_count = 0
    kept_count = 0
    with z.open(records_path) as fh:
        context = iterparse(fh, events=('end',))
        for event, el in context:
            if local(el.tag) == 'r' and el.tag != qn('pivotCacheRecords'):
                values = [decode_record_value(child, cache_fields, i) for i, child in enumerate(el)]
                row = dict(zip(field_names, values))
                raw_count += 1
                el.clear()

                if row_filter is not None:
                    row = row_filter(row)
                    if row is None:
                        continue

                rows.append(row)
                kept_count += 1
                if limit is not None and kept_count >= limit:
                    break
    return rows, kept_count, raw_count


SUBTOTAL_TO_FUNC = {
    'sum': 'sum',
    'average': 'avg',
    'count': 'count',
    'countNums': 'count',
    'max': 'max',
    'min': 'min',
    # product/stdDev/stdDevp/var/varp -> не поддерживаются в MVP
}


def build_config(pivot_table_root, cache_fields):
    field_names = [f['name'] for f in cache_fields]

    row_fields_el = pivot_table_root.find(qn('rowFields'))
    col_fields_el = pivot_table_root.find(qn('colFields'))
    data_fields_el = pivot_table_root.find(qn('dataFields'))

    def field_indices(container_el):
        if container_el is None:
            return []
        result = []
        for f in container_el.findall(qn('field')):
            x = f.get('x')
            if x is not None and int(x) >= 0:  # x=-2 (0xFFFFFFFE) означает "Values", пропускаем
                result.append(int(x))
        return result

    rows = [field_names[i] for i in field_indices(row_fields_el)]
    columns = [field_names[i] for i in field_indices(col_fields_el)]

    measures = []
    funcs_used = set()
    measure_func_pairs = []
    if data_fields_el is not None:
        for df in data_fields_el.findall(qn('dataField')):
            fld = int(df.get('fld'))
            subtotal = df.get('subtotal', 'sum')
            name = field_names[fld]
            func = SUBTOTAL_TO_FUNC.get(subtotal)
            if func is None:
                print(f'  [!] Агрегация "{subtotal}" для поля "{name}" не поддерживается в MVP, пропускаю')
                continue
            measures.append(name)
            funcs_used.add(func)
            measure_func_pairs.append((name, func))

    dimensions = rows + [c for c in columns if c not in rows]
    fields = {name: {'label': name, 'title': name} for name in field_names}

    config = {
        'dimensions': dimensions,
        'measures': measures,
        'funcs': sorted(funcs_used) if funcs_used else ['sum'],
        'fields': fields,
        'rows': rows,
        'columns': columns,
        'measure': measure_func_pairs[0][0] if measure_func_pairs else None,
        'func': measure_func_pairs[0][1] if measure_func_pairs else None,
    }
    return config


def compute_hidden_values(pivot_table_root, cache_fields):
    """
    Возвращает {field_name: set(скрытых значений)} — только для полей,
    где в самой сводной сняты галочки в выпадающем фильтре значений поля
    (<item h="1"/> в <pivotFields><pivotField><items>).

    Это не наша фильтрация "по смыслу" — это ровно то состояние, которое
    аналитик сам выставил в Excel, мы просто его читаем и применяем один раз
    при разборе. Никакого интерактивного фильтра в html не появляется —
    набор данных уже такой, каким он должен быть.
    """
    hidden_by_field = {}
    pivot_fields_el = pivot_table_root.find(qn('pivotFields'))
    if pivot_fields_el is None:
        return hidden_by_field

    for field_index, pf in enumerate(pivot_fields_el.findall(qn('pivotField'))):
        items_el = pf.find(qn('items'))
        if items_el is None:
            continue
        shared = cache_fields[field_index]['shared_items']
        if shared is None:
            group = cache_fields[field_index].get('group')
            shared = group['items'] if group else None
        if shared is None:
            continue

        hidden_values = set()
        for item in items_el.findall(qn('item')):
            if item.get('t') is not None:
                continue  # служебные item (default/subtotal), не значение поля
            if item.get('h') != '1':
                continue  # видимое значение — не трогаем
            x = item.get('x')
            if x is None:
                continue
            idx = int(x)
            if idx < len(shared):
                hidden_values.add(shared[idx])

        if hidden_values:
            hidden_by_field[cache_fields[field_index]['name']] = hidden_values

    return hidden_by_field


def compute_page_filter_values(pivot_table_root, cache_fields):
    """
    Отдельная зона "Фильтры отчёта" (Report Filter) — если там выбрано
    конкретное значение, а не "(Все)". <pageFields><pageField fld=".." item=".."/>
    Возвращает {field_name: единственное разрешённое значение}.
    """
    result = {}
    page_fields_el = pivot_table_root.find(qn('pageFields'))
    if page_fields_el is None:
        return result

    for pf in page_fields_el.findall(qn('pageField')):
        fld = pf.get('fld')
        item = pf.get('item')
        if fld is None or item is None:
            continue
        fld = int(fld)
        shared = cache_fields[fld]['shared_items']
        if shared is None:
            group = cache_fields[fld].get('group')
            shared = group['items'] if group else None
        if shared is None:
            continue
        idx = int(item)
        if idx < len(shared):
            result[cache_fields[fld]['name']] = shared[idx]

    return result


def convert(xlsx_path, row_limit=None):
    with zipfile.ZipFile(xlsx_path) as z:
        pivot_table_sheets = find_pivot_tables_by_sheet(z)
        if not pivot_table_sheets:
            raise ValueError('В файле не найдено ни одной сводной таблицы')

        workbook_rels = read_rels(z, 'xl/workbook.xml')
        cache_id_map = get_cache_id_map(z)

        results = []
        for sheet_name, pt_path in pivot_table_sheets:
            cache_def_path, pt_root = resolve_cache_definition_for_pivot_table(
                z, pt_path, workbook_rels, cache_id_map
            )

            cache_def_root_bytes = z.read(cache_def_path)
            cache_source_root = ET.fromstring(cache_def_root_bytes)
            source_type = cache_source_root.find(qn('cacheSource')).get('type')
            if source_type != 'worksheet':
                print(f'  [!] {pt_path}: источник "{source_type}" (не worksheet) — пропускаю, это Data Model')
                continue

            cache_fields = parse_cache_fields(cache_def_root_bytes)
            cache_def_rels = read_rels(z, cache_def_path)
            records_path = next(iter(cache_def_rels.values()))  # единственная связь — records

            config = build_config(pt_root, cache_fields)

            # Фильтры/группы вычисляем ЗАРАНЕЕ (не зависят от самих данных),
            # чтобы применить их построчно во время чтения, а не после —
            # иначе лимит обрезал бы СЫРЫЕ записи до фильтрации, что при
            # активном фильтре может дать неполный/неверный результат.
            group_fields = [(f['name'], f['group']) for f in cache_fields if f.get('group')]
            hidden_by_field = compute_hidden_values(pt_root, cache_fields)
            page_filter = compute_page_filter_values(pt_root, cache_fields)

            dropped_blank = 0
            dropped_filtered = 0

            def _row_filter(row):
                nonlocal dropped_blank, dropped_filtered
                # 1) Годы/Кварталы/Месяцы — вычисляются из поля-источника,
                #    их нет как raw-значений в записи.
                for field_name, group in group_fields:
                    base_name = cache_fields[group['base_index']]['name']
                    row[field_name] = compute_group_value(group, row.get(base_name))

                # 2) Полностью пустая запись — артефакт завышенного диапазона
                #    листа-источника (не запись).
                if not any(v is not None for v in row.values()):
                    dropped_blank += 1
                    return None

                # 3) Фильтры, реально выставленные в самой сводной (снятые
                #    галочки в списке значений поля + "Фильтры отчёта").
                for fname, hidden_vals in hidden_by_field.items():
                    if row.get(fname) in hidden_vals:
                        dropped_filtered += 1
                        return None
                for fname, allowed_val in page_filter.items():
                    if row.get(fname) != allowed_val:
                        dropped_filtered += 1
                        return None

                return row

            data, count, raw_count = stream_decode_records(
                z, records_path, cache_fields, limit=row_limit, row_filter=_row_filter
            )

            # Обрезка теперь фиксируется по количеству строк, ПРОШЕДШИХ
            # фильтр (не по сырым записям) — именно они отражают, останется
            # ли за пределами прочитанного ещё реальных подходящих строк.
            truncated = row_limit is not None and count >= row_limit

            if dropped_blank or dropped_filtered:
                print(f'  [i] {pt_path}: убрано {dropped_blank} пустых записей, '
                      f'{dropped_filtered} отфильтровано сводной '
                      f'(скрытые значения: {hidden_by_field}, page filter: {page_filter})')

            results.append({
                'sheet_name': sheet_name,
                'pivot_table': pt_path,
                'row_count': count,
                'truncated': truncated,
                'config': config,
                'data_sample': data,
            })
        return results


if __name__ == '__main__':
    import sys
    import json
    import time

    path = sys.argv[1] if len(sys.argv) > 1 else 'test_dataS.xlsx'
    limit = int(sys.argv[2]) if len(sys.argv) > 2 else 5

    t0 = time.time()
    results = convert(path, row_limit=limit)
    elapsed = time.time() - t0

    for r in results:
        print(f'\n=== {r["pivot_table"]} ===')
        print(f'Строк декодировано (лимит={limit}): {r["row_count"]}')
        print('Config (сгенерирован из структуры файла, без хардкода имён):')
        print(json.dumps(r['config'], ensure_ascii=False, indent=2))
        print('Первые декодированные строки:')
        for row in r['data_sample'][:5]:
            print(' ', row)

    print(f'\nВремя выполнения: {elapsed:.2f} сек (при лимите {limit} строк)')
