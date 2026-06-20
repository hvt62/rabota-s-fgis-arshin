import os
import time
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

# Настройки повторных попыток
MAX_RETRIES = 5
RETRY_DELAY = 3  # секунд, с каждым разом увеличивается

# Задержка между запросами к разным номерам (чтобы не превышать лимит)
REQUEST_DELAY = 2

# Сессия для запросов к Аршину (с cookies)
_session = None


def get_arshin_session():
    """Создаёт сессию с cookies, полученными с главной страницы Аршина."""
    global _session
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

    # Сначала заходим на главную страницу, чтобы получить cookies сессии
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
    """Поиск сведений о поверке по номеру СИ и опционально по типу."""
    params = {
        "fq": [f"*{mi_number}*"],
        "q": "*",
        "fl": ",".join(FIELDS),
        "sort": "verification_date desc,org_title asc",
        "rows": 20,
        "start": 0,
    }
    # Если указан тип — добавляем фильтр по нему
    if mi_type:
        params["fq"].append(f"mi.mitype:*{mi_type}*")

    headers = {
        "Referer": "https://fgis.gost.ru/fundmetrology/cm/results",
    }

    session = get_arshin_session()
    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.get(ARSHIN_URL, params=params, headers=headers, timeout=60)
            print(f"[API] Попытка {attempt}: статус {resp.status_code}")

            resp.raise_for_status()
            data = resp.json()
            docs = data.get("response", {}).get("docs", [])
            print(f"[API] Найдено документов: {len(docs)}")
            if docs:
                print(f"[API] Первый результат: {docs[0]}")
            return docs
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else 0
            print(f"[API] HTTP ошибка {status}: {e}")

            # 429 Too Many Requests — ждём дольше и пробуем снова
            if status == 429 and attempt < MAX_RETRIES:
                wait = RETRY_DELAY * attempt * 2  # 6, 12, 18, 24 сек
                last_error = f"Слишком много запросов (429), пауза {wait}с, попытка {attempt}/{MAX_RETRIES}"
                print(f"[API] {last_error}")
                time.sleep(wait)
                continue

            # 5xx — временная ошибка сервера
            if 500 <= status < 600 and attempt < MAX_RETRIES:
                wait = RETRY_DELAY * attempt
                last_error = f"Сервер временно недоступен ({status}), пауза {wait}с, попытка {attempt}/{MAX_RETRIES}"
                print(f"[API] {last_error}")
                time.sleep(wait)
                continue

            last_error = str(e)
            break
        except requests.exceptions.Timeout:
            print(f"[API] Таймаут, попытка {attempt}")
            if attempt < MAX_RETRIES:
                wait = RETRY_DELAY * attempt
                last_error = f"Таймаут, пауза {wait}с, попытка {attempt}/{MAX_RETRIES}"
                time.sleep(wait)
                continue
            last_error = "Сервер не ответил за отведённое время"
        except requests.exceptions.ConnectionError:
            print(f"[API] Ошибка соединения, попытка {attempt}")
            if attempt < MAX_RETRIES:
                wait = RETRY_DELAY * attempt
                last_error = f"Ошибка соединения, пауза {wait}с, попытка {attempt}/{MAX_RETRIES}"
                time.sleep(wait)
                continue
            last_error = "Не удалось подключиться к серверу"
        except requests.exceptions.RequestException as e:
            print(f"[API] Ошибка: {e}")
            last_error = str(e)
            break

    return {"error": last_error}


def format_date(date_str):
    """Преобразует дату из ISO в ДД.ММ.ГГГГ."""
    if not date_str:
        return ""
    return date_str[:10].replace("-", ".")


def format_applicability(val):
    """Преобразует результат поверки в читаемый вид."""
    if val is True:
        return "ГОДЕН"
    elif val is False:
        return "НЕ ГОДЕН"
    return str(val)


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

        # Собираем номера и типы из первого и второго столбцов (начиная со второй строки)
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

        # Опрашиваем API для каждой строки с задержкой между запросами
        results = []
        errors = []
        for idx, item in enumerate(rows_data):
            num = item["number"]
            mi_type = item["type"] if item["type"] else None
            print(f"[App] Ищем номер {idx+1}/{len(rows_data)}: {num}, тип: {mi_type or 'не указан'}")

            docs = search_arshin(num, mi_type)
            if isinstance(docs, dict) and "error" in docs:
                print(f"[App] Ошибка для {num}: {docs['error']}")
                errors.append({"number": num, "error": docs["error"]})
            elif docs:
                print(f"[App] Найдено {len(docs)} записей для {num}")
                for doc in docs:
                    results.append(
                        {
                            "number": num,
                            "input_type": mi_type or "",
                            "mi_number": doc.get("mi.number", ""),
                            "title": doc.get("mi.mititle", ""),
                            "type": doc.get("mi.mitype", ""),
                            "modification": doc.get("mi.modification", ""),
                            "verification_date": format_date(
                                doc.get("verification_date", "")
                            ),
                            "valid_date": format_date(doc.get("valid_date", "")),
                            "applicability": format_applicability(
                                doc.get("applicability")
                            ),
                            "org_title": doc.get("org_title", ""),
                            "result_docnum": doc.get("result_docnum", ""),
                        }
                    )
            else:
                print(f"[App] Нет данных для {num}")
                results.append(
                    {
                        "number": num,
                        "input_type": mi_type or "",
                        "mi_number": "",
                        "title": "Не найдено",
                        "type": "",
                        "modification": "",
                        "verification_date": "",
                        "valid_date": "",
                        "applicability": "",
                        "org_title": "",
                        "result_docnum": "",
                    }
                )

            # Задержка между запросами, чтобы не превысить лимит API
            if idx < len(rows_data) - 1:
                print(f"[App] Пауза {REQUEST_DELAY}с перед следующим запросом...")
                time.sleep(REQUEST_DELAY)

        # Считаем, сколько номеров не найдено
        not_found_count = sum(1 for r in results if r["title"] == "Не найдено")

        return render_template(
            "result.html", results=results, errors=errors, total=len(rows_data),
            not_found_count=not_found_count
        )

    return render_template("index.html")


@app.route("/download/", methods=["POST"])
def download():
    """Скачать результаты в Excel."""
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
        ws.append(
            [
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
            ]
        )

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