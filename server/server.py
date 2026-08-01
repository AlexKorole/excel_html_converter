"""
server.py

REST API + раздача статики на стандартной библиотеке Python (http.server),
без внешних веб-фреймворков. Список/форма — статичные страницы в client/,
получают данные через fetch к этому API (тот же паттерн, что в NoTimeoutCSV).

Структура проекта:
  project/
    server/
      server.py          <- этот файл
      converters/         <- парсеры и генераторы (xlsx_to_*.py и т.д.)
      reports/             <- runtime, создаётся автоматически
    client/
      index.html           <- список отчётов (статичный, JS дорисовывает)
      new.html             <- форма создания отчёта
      js/                  api.js, reports-list.js, new-report.js
      css/style.css
      node_modules/pivotgrid-js/, table-tools/, chart-tools/  <- для сгенерированных отчётов

Сервер раздаёт статику от КОРНЯ проекта (на уровень выше server/), чтобы
и client/, и server/reports/ были достижимы обычными относительными
путями браузера.

Роуты:
  GET    /                    - client/index.html
  GET    /new                 - client/new.html
  GET    /api/reports         - JSON-список отчётов
  POST   /api/reports         - создание (multipart), запуск фоновой сборки
  PUT    /api/reports/<id>    - пересборка существующего отчёта (тот же id и
                                 имя, старое содержимое папки удаляется ПЕРЕД
                                 стартом новой сборки — без атомарной подмены)
  DELETE /api/reports/<id>    - удаление отчёта целиком
  GET    /<любой путь>        - статика от корня проекта (client/js/..., server/reports/...)
"""

import argparse
import json
from messages import t
import multiprocessing
import os
import re
import shutil
import sys
import threading
import traceback
import uuid
from datetime import datetime, timezone
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import unquote

SERVER_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SERVER_DIR.parent
CONVERTERS_DIR = SERVER_DIR / "converters"
CLIENT_DIR = PROJECT_ROOT / "client"

# --server-config / --client-config — путь к внешним конфигам, лежащим
# ВНЕ node_modules (сами бандловые server/.env и client/js/config.js
# слетают при npm install/переустановке пакета).
_arg_parser = argparse.ArgumentParser(description="excel-html-converter server")
_arg_parser.add_argument("--server-config", dest="server_config", default=None,
                          help="Путь к внешнему .env (вместо бандлового server/.env)")
_arg_parser.add_argument("--client-config", dest="client_config", default=None,
                          help="Путь к внешнему config.js (вместо бандлового client/js/config.js)")
_args, _ = _arg_parser.parse_known_args()

CLIENT_CONFIG_OVERRIDE = None
if _args.client_config:
    _client_config_candidate = Path(_args.client_config).resolve()
    if _client_config_candidate.exists():
        CLIENT_CONFIG_OVERRIDE = _client_config_candidate
    else:
        print(f"[!] --client-config points to {_client_config_candidate}, but that file "
              f"doesn't exist — using the bundled client/js/config.js instead")


def _load_env(path):
    """Простой .env-парсер (KEY=VALUE построчно) — без python-dotenv,
    та же логика, что и в NoTimeoutCSV, без лишних зависимостей."""
    env = {}
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            env[key.strip()] = value.strip()
    return env


# ВАЖНО: .env должен быть загружен и прокинут в os.environ ДО импорта
# конвертеров — они читают свои дефолты (лимиты строк и т.п.) через
# os.environ.get(...) на уровне модуля, в момент самого импорта.
_env_path = SERVER_DIR / ".env"
if _args.server_config:
    _server_config_candidate = Path(_args.server_config).resolve()
    if _server_config_candidate.exists():
        _env_path = _server_config_candidate
    else:
        print(f"[!] --server-config points to {_server_config_candidate}, but that file "
              f"doesn't exist — using the bundled server/.env instead (if present)")
_env = _load_env(_env_path)
for _key, _value in _env.items():
    os.environ.setdefault(_key, _value)

PORT = int(os.environ.get("PORT", 8000))
HOST = os.environ.get("HOST", "localhost")
LANGUAGE = os.environ.get("LANGUAGE", "en")

# REPORTS_DIR — НЕ внутри пакета по умолчанию. При установке через
# npm install пакет целиком лежит в node_modules/, а node_modules — папка
# одноразовая (rm -rf node_modules && npm install — обычное дело). Если
# отчёты (загруженные xlsx + сгенерированные страницы) хранить внутри
# пакета, они пропадут при переустановке. По умолчанию — папка "reports"
# там, откуда реально запущена команда (текущая рабочая директория), не
# рядом со server.py. Переопределяется через REPORTS_DIR в server/.env.
_reports_dir_setting = os.environ.get("REPORTS_DIR")
REPORTS_DIR = Path(_reports_dir_setting).resolve() if _reports_dir_setting else (Path.cwd() / "reports").resolve()
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

sys.path.insert(0, str(CONVERTERS_DIR))
import xlsx_report_builder  # noqa: E402 (после правки sys.path и .env)
import xlsx_pivot_to_grid  # noqa: E402 (для find_pivotgrid_pkg - см. Handler.translate_path)

# Фоновая сборка отчёта идёт в ОТДЕЛЬНОМ ПРОЦЕССЕ (см. multiprocessing.Process
# ниже). На Windows такие процессы всегда создаются через spawn (не fork, как
# на Linux) — рабочая папка дочернего процесса не всегда совпадает с тем, что
# ожидается, из-за чего поиск pivotgrid-js по Path.cwd() внутри дочернего
# процесса может не сработать, даже если в этом (основном) процессе всё
# нашлось нормально. Решаем один раз ЗДЕСЬ (тут cwd точно верный — это и
# есть та папка, откуда реально запущена команда) и кладём результат в
# переменную окружения: её дочерний процесс унаследует надёжно на любой
# платформе, в отличие от самой рабочей папки.
if not os.environ.get("PIVOTGRID_PKG"):
    try:
        os.environ["PIVOTGRID_PKG"] = str(xlsx_pivot_to_grid.find_pivotgrid_pkg())
    except FileNotFoundError:
        pass  # не нашли — конкретная ошибка всплывёт при сборке отчёта со сводной


# ── Хранилище (meta.json на отчёт, без общего файла-маппинга) ──────────────

def _meta_path(report_id):
    return REPORTS_DIR / report_id / "meta.json"


def read_meta(report_id):
    path = _meta_path(report_id)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None


def write_meta(report_id, meta):
    _meta_path(report_id).write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def list_reports():
    """Все отчёты (читаем каждую подпапку отдельно), новые сверху."""
    reports = []
    for d in REPORTS_DIR.iterdir():
        if not d.is_dir():
            continue
        meta = read_meta(d.name)
        if meta:
            reports.append(meta)
    reports.sort(key=lambda m: m.get("created_at", ""), reverse=True)
    return reports


# ── Фоновая сборка отчёта (отдельный процесс) ───────────────────────────────

def build_report_background(report_id, xlsx_path, include_connected, pivot_limit, table_limit):
    report_dir = REPORTS_DIR / report_id
    try:
        xlsx_report_builder.build_report(
            str(xlsx_path), str(report_dir),
            include_connected_tables=include_connected,
            pivot_row_limit=pivot_limit,
            table_row_limit=table_limit,
        )
        meta = read_meta(report_id) or {}
        meta["status"] = "ready"
        write_meta(report_id, meta)
    except Exception as e:
        meta = read_meta(report_id) or {}
        meta["status"] = "error"
        meta["error"] = str(e)
        write_meta(report_id, meta)
        traceback.print_exc()
    finally:
        try:
            xlsx_path.unlink(missing_ok=True)
        except OSError:
            pass


def _start_build(report_id, name, created_at, xlsx_bytes, include_connected, pivot_limit, table_limit):
    """Создаёт (или пересоздаёт — папка к этому моменту уже должна быть
    удалена/создана заново вызывающим кодом) содержимое отчёта: пишет
    исходный xlsx, meta.json со статусом "processing" и запускает фоновую
    сборку в отдельном процессе + поток-наблюдатель. Общий код для
    _handle_create и _handle_update — отличаются только тем, что кладут
    в report_id/created_at (новые для create, сохранённые для update).

    created_at и build_started_at — разные вещи: created_at не меняется
    при update (иначе отчёт скакнёт в начало списка), а build_started_at
    всегда "сейчас" — момент старта именно ТЕКУЩЕЙ сборки, по нему
    считается таймер "обрабатывается... (N сек)" в reports-list.js. Если
    таймер считать от created_at, при update он будет тикать от времени
    самого первого создания отчёта, а не от начала пересборки.
    """
    report_dir = REPORTS_DIR / report_id
    report_dir.mkdir(parents=True, exist_ok=True)

    xlsx_path = report_dir / "_source.xlsx"
    xlsx_path.write_bytes(xlsx_bytes)

    meta = {
        "id": report_id,
        "name": name,
        "created_at": created_at,
        "build_started_at": datetime.now(timezone.utc).isoformat(),
        "status": "processing",
    }
    write_meta(report_id, meta)

    # Сборка — в отдельном ПРОЦЕССЕ (не потоке): падение изолировано,
    # не заденет сам сервер.
    process = multiprocessing.Process(
        target=build_report_background,
        args=(report_id, xlsx_path, include_connected, pivot_limit, table_limit),
        daemon=True,
    )
    process.start()

    # Лёгкий поток-наблюдатель — если процесс умрёт без записи статуса
    # (например, OOM-kill), meta.json иначе навсегда останется в "processing".
    def _watch(proc, rid):
        proc.join()
        m = read_meta(rid)
        if m and m.get("status") == "processing":
            m["status"] = "error"
            m["error"] = t("process_died_unexpectedly", LANGUAGE, code=proc.exitcode)
            write_meta(rid, m)

    threading.Thread(target=_watch, args=(process, report_id), daemon=True).start()

    return meta


# ── multipart/form-data (свой разбор — cgi устарел и убран в Python 3.13) ───

def parse_multipart(body, boundary):
    """Возвращает {имя_поля: {'value': bytes, 'filename': str|None}}."""
    boundary_bytes = ("--" + boundary).encode("utf-8")
    parts = body.split(boundary_bytes)
    fields = {}
    for part in parts:
        part = part.strip(b"\r\n")
        if not part or part == b"--":
            continue
        if b"\r\n\r\n" not in part:
            continue
        headers_raw, content = part.split(b"\r\n\r\n", 1)
        if content.endswith(b"\r\n"):
            content = content[:-2]
        headers_text = headers_raw.decode("utf-8", errors="replace")
        name_match = re.search(r'name="([^"]*)"', headers_text)
        if not name_match:
            continue
        name = name_match.group(1)
        filename_match = re.search(r'filename="([^"]*)"', headers_text)
        filename = filename_match.group(1) if filename_match else None
        fields[name] = {"value": content, "filename": filename}
    return fields


# ── HTTP-обработчик ───────────────────────────────────────────────────────────

class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        # Раздаём статику от корня ПРОЕКТА (на уровень выше server/), чтобы
        # и client/, и server/reports/ были достижимы обычными путями.
        super().__init__(*args, directory=str(PROJECT_ROOT), **kwargs)

    def translate_path(self, path):
        # /server/reports/... — это REPORTS_DIR, который теперь МОЖЕТ
        # лежать за пределами PROJECT_ROOT (см. комментарий у REPORTS_DIR
        # выше) — маппим явно, не полагаясь на вложенность в PROJECT_ROOT.
        path_only = path.split("?", 1)[0].split("#", 1)[0]
        prefix = "/server/reports/"
        if path_only.startswith(prefix):
            rel = unquote(path_only[len(prefix):])
            return str(REPORTS_DIR / rel)

        candidate = super().translate_path(path)
        if os.path.exists(candidate):
            return candidate

        # Не нашли внутри PROJECT_ROOT — вероятно, относительная ссылка на
        # одну из клиентских библиотек была посчитана исходя из РЕАЛЬНОГО
        # расположения на диске (может быть где угодно — hoisting в
        # node_modules, отдельный REPORTS_DIR и т.п.), а не совпадает с
        # тем, что подразумевает URL. Ищем характерный фрагмент пути и
        # мапим то, что после него, на реальное расположение библиотеки —
        # независимо от того, сколько "../" сюда привело.
        markers = (
            ("pivotgrid-js" + os.sep,
             lambda: xlsx_pivot_to_grid.find_pivotgrid_pkg()),
            (os.path.join("client", "table-tools") + os.sep,
             lambda: CLIENT_DIR / "table-tools"),
            (os.path.join("client", "chart-tools") + os.sep,
             lambda: CLIENT_DIR / "chart-tools"),
        )
        for marker, get_base_dir in markers:
            if marker not in candidate:
                continue
            try:
                base_dir = get_base_dir()
            except FileNotFoundError:
                continue
            suffix = candidate.split(marker, 1)[1]
            alt_candidate = str(base_dir / suffix)
            if os.path.exists(alt_candidate):
                return alt_candidate

        return candidate

    def do_GET(self):
        path_only = self.path.split("?", 1)[0]
        if path_only == "/" or path_only == "/index.html":
            self._send_file(CLIENT_DIR / "index.html", "text/html; charset=utf-8")
        elif path_only == "/new" or path_only == "/new.html":
            self._send_file(CLIENT_DIR / "new.html", "text/html; charset=utf-8")
        elif path_only == "/client/js/config.js" and CLIENT_CONFIG_OVERRIDE:
            self._send_file(CLIENT_CONFIG_OVERRIDE, "application/javascript; charset=utf-8")
        elif path_only == "/api/reports":
            self._send_json(list_reports())
        else:
            super().do_GET()

    def do_POST(self):
        if self.path == "/api/reports":
            self._handle_create()
        else:
            self.send_error(404)

    def do_PUT(self):
        m = re.match(r"^/api/reports/([^/]+)$", self.path)
        if not m:
            self.send_error(404)
            return
        self._handle_update(unquote(m.group(1)))

    def do_DELETE(self):
        m = re.match(r"^/api/reports/([^/]+)$", self.path)
        if not m:
            self.send_error(404)
            return
        report_id = unquote(m.group(1))
        report_dir = REPORTS_DIR / report_id
        # защита от выхода за пределы reports/ через ../ в id
        if report_dir.exists() and report_dir.is_dir() and report_dir.resolve().parent == REPORTS_DIR.resolve():
            shutil.rmtree(report_dir)
            self._send_json({"deleted": True})
        else:
            self.send_error(404)

    def _parse_report_form(self, require_name):
        """Разбирает multipart-тело формы (общей для создания и обновления
        отчёта) и валидирует поля. При require_name=False поле "name" не
        обязательно — так и есть при обновлении, поскольку в форме оно
        заблокировано (disabled) и браузер его вообще не отправляет; имя в
        этом случае берётся из уже сохранённого meta.json, не из формы.

        Возвращает (fields_dict, None) при успехе или (None, (message, status))
        при ошибке валидации.
        """
        content_type = self.headers.get("Content-Type", "")
        boundary_match = re.search(r"boundary=(.+)", content_type)
        if "multipart/form-data" not in content_type or not boundary_match:
            return None, (t("invalid_request", LANGUAGE), 400)

        boundary = boundary_match.group(1).strip('"')
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)
        fields = parse_multipart(body, boundary)

        name_field = fields.get("name")
        file_field = fields.get("file")
        if (require_name and not name_field) or not file_field or not file_field.get("filename"):
            key = "name_and_file_required" if require_name else "file_required"
            return None, (t(key, LANGUAGE), 400)

        name = None
        if name_field:
            name = name_field["value"].decode("utf-8", errors="replace").strip()
            if require_name and not name:
                return None, (t("name_empty", LANGUAGE), 400)

        include_connected = "include_connected" in fields

        def _optional_int_field(field_name, label_key):
            label = t(label_key, LANGUAGE)
            f = fields.get(field_name)
            text = f["value"].decode("utf-8", errors="replace").strip() if f else ""
            if not text:
                return None, None  # не указано - используется дефолт конвертера
            try:
                value = int(text)
            except ValueError:
                return None, t("must_be_number", LANGUAGE, label=label)
            if value < 1:
                return None, t("must_be_positive", LANGUAGE, label=label)
            return value, None

        pivot_limit, pivot_error = _optional_int_field("pivot_limit", "pivot_limit_label")
        if pivot_error:
            return None, (pivot_error, 400)

        table_limit, table_error = _optional_int_field("table_limit", "table_limit_label")
        if table_error:
            return None, (table_error, 400)

        return {
            "name": name,
            "file_bytes": file_field["value"],
            "include_connected": include_connected,
            "pivot_limit": pivot_limit,
            "table_limit": table_limit,
        }, None

    def _handle_create(self):
        parsed, error = self._parse_report_form(require_name=True)
        if error:
            message, status = error
            self._send_json({"error": message}, status=status)
            return

        report_id = uuid.uuid4().hex[:8]
        meta = _start_build(
            report_id=report_id,
            name=parsed["name"],
            created_at=datetime.now(timezone.utc).isoformat(),
            xlsx_bytes=parsed["file_bytes"],
            include_connected=parsed["include_connected"],
            pivot_limit=parsed["pivot_limit"],
            table_limit=parsed["table_limit"],
        )
        self._send_json(meta, status=201)

    def _handle_update(self, report_id):
        report_dir = REPORTS_DIR / report_id
        # защита от выхода за пределы reports/ через ../ в id - та же
        # проверка, что и в do_DELETE
        if not (report_dir.exists() and report_dir.is_dir()
                and report_dir.resolve().parent == REPORTS_DIR.resolve()):
            self._send_json({"error": t("report_not_found", LANGUAGE)}, status=404)
            return

        existing_meta = read_meta(report_id)
        if not existing_meta:
            self._send_json({"error": t("report_not_found", LANGUAGE)}, status=404)
            return

        parsed, error = self._parse_report_form(require_name=False)
        if error:
            message, status = error
            self._send_json({"error": message}, status=status)
            return

        # Имя и id сохраняются от старого отчёта независимо от формы -
        # переименование через Update не предусмотрено.
        # Сервер сам удаляет старое содержимое ДО старта новой сборки
        # (без атомарной подмены - окно гонки принято сознательно).
        shutil.rmtree(report_dir)
        meta = _start_build(
            report_id=report_id,
            name=existing_meta["name"],
            created_at=existing_meta.get("created_at", datetime.now(timezone.utc).isoformat()),
            xlsx_bytes=parsed["file_bytes"],
            include_connected=parsed["include_connected"],
            pivot_limit=parsed["pivot_limit"],
            table_limit=parsed["table_limit"],
        )
        self._send_json(meta, status=200)

    def _send_file(self, path, content_type):
        try:
            data = path.read_bytes()
        except OSError:
            self.send_error(404)
            return
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        print(f"[{self.log_date_time_string()}] {format % args}")


if __name__ == "__main__":
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(t("server_running", LANGUAGE, host=HOST, port=PORT))
    print(t("project_label", LANGUAGE, path=PROJECT_ROOT))
    print(t("reports_label", LANGUAGE, path=REPORTS_DIR))
    server.serve_forever()
