import os
import time
import threading
import requests
import openpyxl
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

# Новый рабочий endpoint ФГИС «Аршин»
ARSHIN_URL = "https://fgis.gost.ru/fundmetrology/cm/xcdb/vri/select"
ARSHIN_MAIN = "https://fgis.gost.ru/fundmetrology/cm/"

# Поля, которые запрашиваем у API
FIELDS = [
    "vri_id",
    "org_title",
    "mi.mitnumber",
    "mi.mititle",
    "mi.mitype",
    "mi.modification",
    "mi.number",
    "verification_date",
    "valid_date",
    "applicability",
    "result_docnum",
    "sticker_num",
]

# Настройки повторных попыток
MAX_RETRIES = 10
RETRY_DELAY = 5

# Параллельные запросы
MAX_WORKERS = 1  # последовательные запросы, чтобы избежать 429

# Сессия для запросов к Аршину (с cookies)
_session = None
_session_lock = threading.Lock()


def get_arshin_session():
    """Создаёт сессию с cookies, полученными с главной страницы Аршина."""
    global _session
    if _session is not None:
        return _session

    with _session_lock:
        if _session is not None:
            return _session

        session = requests.Session()
        session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "ru,en;q=0.9",
        })

        try:
            resp = session.get(ARSHIN_MAIN, timeout=30)
            resp.raise_for_status()
            print(f"[Сессия] Статус: {resp.status_code}")
            print(f"[Сессия] Cookies от сервера: {dict(session.cookies)}")
        except Exception as e:
            print(f"[Сессия] Ошибка получения cookies: {e}")

        _session = session
        return _session


def search_arshin(mi_number):
    """Поиск сведений о поверке по номеру СИ."""
    params = {
        "fq": [f"*{mi_number}*"],
        "q": "*",
        "fl": ",".join(FIELDS),
        "sort": "verification_date desc,org_title asc",
        "rows": 20,
        "start": 0,
    }

    headers = {
        "Referer": "https://fgis.gost.ru/fundmetrology/cm/results",
    }

    session = get_arshin_session()
    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.get(ARSHIN_URL, params=params, headers=headers, timeout=60)
            print(f"[API] {mi_number}: попытка {attempt}, статус {resp.status_code}")

            resp.raise_for_status()
            data = resp.json()
            docs = data.get("response", {}).get("docs", [])
            print(f"[API] {mi_number}: найдено {len(docs)} документов")
            return docs
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else 0
            print(f"[API] {mi_number}: HTTP ошибка {status}")

            if status == 429 and attempt < MAX_RETRIES:
                wait = RETRY_DELAY * attempt * 2
                print(f"[API] {mi_number}: 429, пауза {wait}с, попытка {attempt}/{MAX_RETRIES}")
                time.sleep(wait)
                continue

            if 500 <= status < 600 and attempt < MAX_RETRIES:
                wait = RETRY_DELAY * attempt
                print(f"[API] {mi_number}: 5xx, пауза {wait}с, попытка {attempt}/{MAX_RETRIES}")
                time.sleep(wait)
                continue

            last_error = str(e)
            break
        except requests.exceptions.Timeout:
            print(f"[API] {mi_number}: таймаут, попытка {attempt}")
            if attempt < MAX_RETRIES:
                wait = RETRY_DELAY * attempt
                time.sleep(wait)
                continue
            last_error = "Сервер не ответил за отведённое время"
        except requests.exceptions.ConnectionError:
            print(f"[API] {mi_number}: ошибка соединения, попытка {attempt}")
            if attempt < MAX_RETRIES:
                wait = RETRY_DELAY * attempt
                time.sleep(wait)
                continue
            last_error = "Не удалось подключиться к серверу"
        except requests.exceptions.RequestException as e:
            print(f"[API] {mi_number}: ошибка {e}")
            last_error = str(e)
            break

    return {"error": last_error}


def format_date(date_str):
    if not date_str:
        return ""
    return date_str[:10].replace("-", ".")


def format_applicability(val):
    if val is True:
        return "ГОДЕН"
    elif val is False:
        return "НЕ ГОДЕН"
    return str(val)


def process_item(item):
    """Обработать один элемент (выполняется в потоке)."""
    num = item["number"]
    mi_type = item["type"] if item["type"] else ""

    docs = search_arshin(num)
    if isinstance(docs, dict) and "error" in docs:
        return {"error": {"number": num, "error": docs["error"]}}

    if not docs:
        return {"not_found": {"number": num, "input_type": mi_type}}

    # Локальная фильтрация по типу (LIKE, без учёта регистра)
    if mi_type:
        mi_type_lower = mi_type.lower()
        docs = [
            doc for doc in docs
            if mi_type_lower in doc.get("mi.mitype", "").lower()
        ]
        print(f"[App] {num}: после фильтрации по типу '{mi_type}': {len(docs)} записей")

    records = []
    for doc in docs:
        records.append({
            "number": num,
            "input_type": mi_type,
            "mi_number": doc.get("mi.number", ""),
            "title": doc.get("mi.mititle", ""),
            "type": doc.get("mi.mitype", ""),
            "modification": doc.get("mi.modification", ""),
            "verification_date": format_date(doc.get("verification_date", "")),
            "valid_date": format_date(doc.get("valid_date", "")),
            "applicability": format_applicability(doc.get("applicability")),
            "org_title": doc.get("org_title", ""),
            "result_docnum": doc.get("result_docnum", ""),
        })

    return {"records": records}


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            return render_template("index.html", error="Файл не выбран")

        try:
            wb = openpyxl.load_workbook(file, read_only=True)
            ws = wb.active
        except Exception as e:
            return render_template("index.html", error=f"Ошибка чтения Excel: {e}")

        # Собираем номера и типы из первого и второго столбцов
        rows_data = []
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            num = row[0]
            if num is not None:
                num = str(num).strip()
                type_val = str(row[1]).strip() if row[1] is not None else ""
                rows_data.append({"number": num, "type": type_val})

        print(f"[App] Прочитано строк: {len(rows_data)}")

        if not rows_data:
            return render_template("index.html", error="Нет номеров в первом столбце")

        # Последовательный опрос API (1 поток)
        print(f"[App] Запуск последовательных запросов...")
        results = []
        errors = []

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(process_item, item): item for item in rows_data}

            for future in as_completed(futures):
                item = futures[future]
                try:
                    result = future.result()
                    if "error" in result:
                        errors.append(result["error"])
                    elif "not_found" in result:
                        nf = result["not_found"]
                        results.append({
                            "number": nf["number"],
                            "input_type": nf["input_type"],
                            "mi_number": "",
                            "title": "Не найдено",
                            "type": "",
                            "modification": "",
                            "verification_date": "",
                            "valid_date": "",
                            "applicability": "",
                            "org_title": "",
                            "result_docnum": "",
                        })
                    elif "records" in result:
                        results.extend(result["records"])
                except Exception as e:
                    errors.append({"number": item["number"], "error": str(e)})

        # Сортируем результаты в порядке исходных номеров
        order = {item["number"]: idx for idx, item in enumerate(rows_data)}
        results.sort(key=lambda r: order.get(r["number"], 999))

        not_found_count = sum(1 for r in results if r["title"] == "Не найдено")

        # Пагинация: по 50 записей на страницу
        per_page = 50
        total_pages = max(1, (len(results) + per_page - 1) // per_page)
        page_results = results[:per_page]

        return render_template(
            "result.html", results=page_results, errors=errors, total=len(rows_data),
            not_found_count=not_found_count,
            page=1, total_pages=total_pages, per_page=per_page,
            all_results=results, total_records=len(results)
        )

    return render_template("index.html")


@app.route("/page/", methods=["POST"])
def page():
    import json

    results = json.loads(request.form.get("results", "[]"))
    page = int(request.form.get("page", 1))
    per_page = 50
    total_pages = max(1, (len(results) + per_page - 1) // per_page)

    if page < 1:
        page = 1
    if page > total_pages:
        page = total_pages

    start_idx = (page - 1) * per_page
    end_idx = start_idx + per_page
    page_results = results[start_idx:end_idx]

    return render_template(
        "result.html", results=page_results, errors=[], total=0,
        not_found_count=0,
        page=page, total_pages=total_pages, per_page=per_page,
        all_results=results, total_records=len(results)
    )


@app.route("/download/", methods=["POST"])
def download():
    import json

    results = json.loads(request.form.get("results", "[]"))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Результаты"

    headers = [
        "Искомый номер",
        "Заданный тип",
        "Заводской номер СИ",
        "Наименование СИ",
        "Тип СИ",
        "Модификация",
        "Дата поверки",
        "Действителен до",
        "Результат",
        "Организация",
        "Номер документа",
    ]
    ws.append(headers)

    for r in results:
        ws.append([
            r.get("number", ""),
            r.get("input_type", ""),
            r.get("mi_number", ""),
            r.get("title", ""),
            r.get("type", ""),
            r.get("modification", ""),
            r.get("verification_date", ""),
            r.get("valid_date", ""),
            r.get("applicability", ""),
            r.get("org_title", ""),
            r.get("result_docnum", ""),
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="results.xlsx",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)