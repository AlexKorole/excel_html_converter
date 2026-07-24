"""
report_messages.py

Двуязычные (ru/en) подписи для сгенерированных страниц отчёта (футеры,
пункт меню "назад"). Язык — через переменную окружения LANGUAGE (сервер
устанавливает её из server/.env; при чистом CLI-использовании без
сервера, по умолчанию — ru).
"""

import os

MESSAGES = {
    "back_to_reports": {
        "ru": "\u2190 Ко всем отчётам",
        "en": "\u2190 Back to all reports",
    },
    "sheet_label": {"ru": "Лист", "en": "Sheet"},
    "pivot_label": {"ru": "Сводная", "en": "Pivot"},
    "rows_label": {"ru": "строк", "en": "rows"},
    "type_label": {"ru": "тип", "en": "type"},
    "series_label": {"ru": "серий", "en": "series"},
    "shown_first_note": {
        "ru": " (показаны первые {limit}, в исходнике может быть больше)",
        "en": " (showing first {limit}, source may have more)",
    },
}


def report_lang():
    return os.environ.get("LANGUAGE", "en")


def rt(key, **kwargs):
    lang = report_lang()
    entry = MESSAGES.get(key, {})
    text = entry.get(lang) or entry.get("ru") or key
    return text.format(**kwargs) if kwargs else text
