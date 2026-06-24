import os
import time
import threading
import requests
import openpyxl
from io import BytesIO
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


def search_arshin(mi_number, mi_type=None):
    """Один запрос к API (без retry). Возвращает список docs или {"error": ...}."""
    fq_conditions = [f"*{mi_number}*"]
    if mi_type:
        fq_conditions.append(f"mi.mitype:*{mi_type}*")

    params = {
        "fq": fq_conditions,
        "q": "*",
        "fl": ",".join(FIELDS),
        "sort": "verification_date desc,org_title asc",
        "rows": 250,
        "start": 0,
    }

    headers = {
        "Referer": "https://fgis.gost.ru/fundmetrology/cm/results",
    }

    session = get_arshin_session()

    try:
        resp = session.get(ARSHIN_URL, params=params, headers=headers, timeout=60)
        print(f"[API] {mi_number}: статус {resp.status_code}")
        resp.raise_for_status()
        data = resp.json()
        docs = data.get("response", {}).get("docs", [])
        print(f"[API] {mi_number}: найдено {len(docs)} документов")
        return docs
    except requests.exceptions.RequestException as e:
        print(f"[API] {mi_number}: ошибка {e}")
        return {"error": str(e)}


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


def process_docs(docs, num, mi_type):
    """Обработать список документов: локальная фильтрация по типу, форматирование."""
    if not docs:
        return []

    # Локальная фильтрация по типу (LIKE, без учёта регистра)
    if mi_type:
        mi_type_lower = mi_type.lower()
        filtered_docs = [
            doc for doc in docs
            if mi_type_lower in doc.get("mi.mitype", "").lower()
        ]
        print(f"[App] {num}: после локальной фильтрации по типу '{mi_type}': {len(filtered_docs)} из {len(docs)} записей")

        if not filtered_docs:
            print(f"[App] {num}: тип '{mi_type}' не совпал локально, показываем все записи без фильтра")
            docs_to_use = docs
            type_mismatch = True
        else:
            docs_to_use = filtered_docs
            type_mismatch = False
    else:
        docs_to_use = docs
        type_mismatch = False

    records = []
    for doc in docs_to_use:
        record = {
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
        }
        if type_mismatch:
            record["type_mismatch"] = True
            record["title"] = f"⚠️ {record['title']} (тип не совпал)"
        records.append(record)

    return records


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

        # Итеративный поиск: паузы 10, 15, 20, 25, ... 55 сек
        delays = [10, 15, 20, 25, 30, 35, 40, 45, 50, 55]
        pending = list(rows_data)
        results = []
        errors = []

        for iteration, delay in enumerate(delays, 1):
            if not pending:
                print(f"[App] Все номера найдены, досрочное завершение")
                break

            print(f"\n[App] === Итерация {iteration}, пауза {delay}с, осталось {len(pending)} номеров ===")

            found_this_iter = []
            still_pending = []

            for idx, item in enumerate(pending):
                num = item["number"]
                mi_type = item["type"] if item["type"] else ""

                # Сначала ищем с типом (если задан)
                docs = search_arshin(num, mi_type if mi_type else None)

                # Если ошибка — остаётся для следующей итерации
                if isinstance(docs, dict) and "error" in docs:
                    still_pending.append(item)
                    # Пауза перед следующим запросом (кроме последнего элемента)
                    if idx < len(pending) - 1:
                        time.sleep(delay)
                    continue

                # Если с типом ничего не найдено — пробуем без типа
                if not docs and mi_type:
                    print(f"[App] {num}: с типом ничего не найдено, пробуем без типа")
                    docs = search_arshin(num, None)
                    if isinstance(docs, dict) and "error" in docs:
                        still_pending.append(item)
                        if idx < len(pending) - 1:
                            time.sleep(delay)
                        continue

                # Если всё равно ничего не найдено — остаётся на следующую итерацию
                if not docs:
                    print(f"[App] {num}: не найден, остаётся для следующей итерации")
                    still_pending.append(item)
                    if idx < len(pending) - 1:
                        time.sleep(delay)
                    continue

                # Нашли! Обрабатываем результаты
                records = process_docs(docs, num, mi_type)
                found_this_iter.extend(records)

                # Пауза перед следующим запросом (кроме последнего элемента)
                if idx < len(pending) - 1:
                    time.sleep(delay)

            results.extend(found_this_iter)
            pending = still_pending
            print(f"[App] Итерация {iteration}: найдено {len(found_this_iter)} записей, осталось {len(pending)} номеров")

        # Оставшиеся не найденными — добавляем как "Не найдено"
        for item in pending:
            results.append({
                "number": item["number"],
                "input_type": item["type"] if item["type"] else "",
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