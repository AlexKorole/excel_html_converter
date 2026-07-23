"""
messages.py

Двуязычные (ru/en) сообщения сервера. Язык задаётся через LANGUAGE
в server/.env (по умолчанию — ru). Тот же принцип, что и в NoTimeoutCSV.
"""

MESSAGES = {
    "invalid_request": {
        "ru": "Некорректный запрос",
        "en": "Invalid request",
    },
    "name_and_file_required": {
        "ru": "Укажите имя отчёта и выберите xlsx-файл",
        "en": "Please provide a report name and select an xlsx file",
    },
    "name_empty": {
        "ru": "Имя отчёта не может быть пустым",
        "en": "Report name cannot be empty",
    },
    "pivot_limit_label": {
        "ru": "лимит строк для сводной",
        "en": "pivot row limit",
    },
    "table_limit_label": {
        "ru": "лимит строк для таблицы",
        "en": "table row limit",
    },
    "must_be_number": {
        "ru": "{label} должен быть числом",
        "en": "{label} must be a number",
    },
    "must_be_positive": {
        "ru": "{label} должен быть больше 0",
        "en": "{label} must be greater than 0",
    },
    "process_died_unexpectedly": {
        "ru": "Процесс сборки завершился неожиданно (код {code}), вероятная причина "
              "— нехватка памяти на большом файле. Попробуйте задать лимит строк.",
        "en": "The build process terminated unexpectedly (exit code {code}), most "
              "likely due to insufficient memory on a large file. Try setting a row limit.",
    },
    "server_running": {
        "ru": "Сервер:   http://{host}:{port}",
        "en": "Server:   http://{host}:{port}",
    },
    "project_label": {
        "ru": "Проект:   {path}",
        "en": "Project:  {path}",
    },
    "reports_label": {
        "ru": "Отчёты:   {path}",
        "en": "Reports:  {path}",
    },
}


def t(key, lang, **kwargs):
    entry = MESSAGES.get(key, {})
    text = entry.get(lang) or entry.get("ru") or key
    return text.format(**kwargs) if kwargs else text
