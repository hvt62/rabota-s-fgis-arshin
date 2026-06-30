import os
import time
import json
import threading
import requests
import openpyxl
from io import BytesIO
from flask import Flask, render_template, request, flash, send_file, redirect, url_for, jsonify
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.urandom(24).hex()

UPLOAD_FOLDER = '/tmp/uploads'
RESULT_FOLDER = '/tmp/results'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

FGIS_BASE_URL = "https://fgis.gost.ru/fundmetrology"
FGIS_API_URL = f"{FGIS_BASE_URL}/api/verify_results"

# Хранилище прогресса {task_id: {...}}
progress_store = {}
lock = threading.Lock()


def create_fgis_session() -> requests.Session:
    """Создаёт сессию с браузерными заголовками."""
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
        "Referer": FGIS_BASE_URL,
    })
    return session


def search_fgis(session: requests.Session, mi_number: str, mi_type: str = None) -> dict | None:
    """
    Поиск данных о средстве измерений во ФГИС «Аршин».
    3 попытки при 500 ошибке с паузой 7.5 сек.
    rows=1000, фильтр по типу через mi.modification.
    """
    params = {
        "mi_number": f"*{mi_number.strip()}*",
        "rows": 1000,
    }
    if mi_type and mi_type.strip():
        params["mi_modification"] = f"*{mi_type.strip()}*"

    print(f"[search_fgis] Запрос: {mi_number}, тип: {mi_type}")
    for attempt in range(1, 4):
        try:
            resp = session.get(FGIS_API_URL, params=params, timeout=30)
            print(f"[search_fgis] Ответ {mi_number}: статус {resp.status_code}")
            if resp.status_code == 200:
                data = resp.json()
                results = data.get("results", [])
                if not results:
                    print(f"[search_fgis] {mi_number}: результатов нет")
                    return None

                # Если тип указан — проверяем совпадение
                type_mismatch = False
                if mi_type and mi_type.strip():
                    type_found = False
                    for r in results:
                        mod = r.get("mi_modification", "") or ""
                        if mi_type.strip().lower() in mod.lower():
                            type_found = True
                            break
                    if not type_found:
                        type_mismatch = True

                # Берём последнюю (самую свежую) поверку
                r = results[-1]
                result = {
                    "mi_name": r.get("mi_name", ""),
                    "serial_number": r.get("serial_number", ""),
                    "verification_date": r.get("verification_date", ""),
                    "valid_until": r.get("valid_until", ""),
                    "result": r.get("result", ""),
                    "organization": r.get("organization", ""),
                    "type_mismatch": type_mismatch,
                }
                print(f"[search_fgis] {mi_number}: найдено")
                return result

            elif resp.status_code == 500:
                print(f"[search_fgis] {mi_number}: 500 ошибка, попытка {attempt}/3")
                if attempt < 3:
                    time.sleep(7.5)
                continue
            else:
                print(f"[search_fgis] {mi_number}: неожиданный статус {resp.status_code}")
                return None
        except requests.exceptions.Timeout:
            print(f"[search_fgis] {mi_number}: timeout, попытка {attempt}/3")
            if attempt < 3:
                time.sleep(7.5)
            continue
        except Exception as e:
            print(f"[search_fgis] {mi_number}: ошибка {e}, попытка {attempt}/3")
            if attempt < 3:
                time.sleep(7.5)
            continue

    print(f"[search_fgis] {mi_number}: не найден после 3 попыток")
    return None


def process_numbers(numbers, task_id, mi_type=None):
    """
    Обработка списка номеров:
    - Создаёт отдельную сессию для задачи
    - 5 итераций с паузами [3, 7, 12, 18, 25] сек между номерами
    - Обновление прогресса
    """
    print(f"[process_numbers] Старт задачи {task_id}, номеров: {len(numbers)}")
    session = create_fgis_session()

    PAUSES = [3, 7, 12, 18, 25]
    ITERATIONS = 5

    all_rows = []
    found_count = 0
    not_found_count = 0
    type_mismatch_count = 0

    with lock:
        progress_store[task_id]["total"] = len(numbers)
        progress_store[task_id]["status"] = "processing"
        progress_store[task_id]["current_iteration"] = 0
        progress_store[task_id]["current_number"] = ""
        progress_store[task_id]["progress_pct"] = 0

    for iteration in range(1, ITERATIONS + 1):
        with lock:
            progress_store[task_id]["current_iteration"] = iteration
            progress_store[task_id]["iteration_phase"] = f"Итерация {iteration}/{ITERATIONS}"

        for idx, number in enumerate(numbers, start=1):
            # Проверяем, не отменена ли задача
            with lock:
                if progress_store[task_id].get("cancel"):
                    progress_store[task_id]["status"] = "cancelled"
                    return

            with lock:
                progress_store[task_id]["current_number"] = number
                overall_progress = ((iteration - 1) * len(numbers) + idx) / (ITERATIONS * len(numbers)) * 100
                progress_store[task_id]["progress_pct"] = round(overall_progress, 1)

            data = search_fgis(session, number, mi_type)

            if data:
                found_count += 1
                if data.get("type_mismatch"):
                    type_mismatch_count += 1
                all_rows.append({
                    "mi_number": number,
                    "found": True,
                    "type_mismatch": data.get("type_mismatch", False),
                    "mi_name": data["mi_name"],
                    "serial_number": data["serial_number"],
                    "verification_date": data["verification_date"],
                    "valid_until": data["valid_until"],
                    "result": data["result"],
                    "organization": data["organization"],
                })
            else:
                not_found_count += 1
                all_rows.append({
                    "mi_number": number,
                    "found": False,
                    "type_mismatch": False,
                    "mi_name": "",
                    "serial_number": "",
                    "verification_date": "",
                    "valid_until": "",
                    "result": "",
                    "organization": "",
                })

            # Пауза между номерами (кроме последнего номера в последней итерации)
            if idx < len(numbers) or iteration < ITERATIONS:
                pause = PAUSES[min(iteration - 1, len(PAUSES) - 1)]
                time.sleep(pause)

    # Сохраняем результат
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Результаты"

    headers = [
        "№ п/п", "Номер СИ", "Статус",
        "Наименование СИ", "Заводской номер",
        "Дата поверки", "Действителен до",
        "Результат поверки", "Организация",
    ]
    ws_out.append(headers)

    for idx, row in enumerate(all_rows, start=1):
        if row["found"]:
            status = "Найдено"
            if row["type_mismatch"]:
                status = "Найдено ⚠️ тип не совпал"
            ws_out.append([
                idx, row["mi_number"], status,
                row["mi_name"], row["serial_number"],
                row["verification_date"], row["valid_until"],
                row["result"], row["organization"],
            ])
        else:
            ws_out.append([idx, row["mi_number"], "Не найдено", "—", "—", "—", "—", "—", "—"])

    # Автоширина колонок
    for col in ws_out.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws_out.column_dimensions[col_letter].width = min(max_len + 4, 50)

    result_filename = f"result_{task_id}.xlsx"
    output_path = os.path.join(RESULT_FOLDER, result_filename)
    wb_out.save(output_path)

    stats = {
        "total": len(numbers),
        "found": found_count,
        "not_found": not_found_count,
        "type_mismatch": type_mismatch_count,
    }

    with lock:
        progress_store[task_id]["status"] = "complete"
        progress_store[task_id]["progress_pct"] = 100
        progress_store[task_id]["rows"] = all_rows
        progress_store[task_id]["stats"] = stats
        progress_store[task_id]["result_filename"] = result_filename

    print(f"[process_numbers] Задача {task_id} завершена. Найдено: {found_count}, не найдено: {not_found_count}")


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/", methods=["POST"])
def upload():
    # Ручной ввод
    mode = request.form.get("mode", "")
    if mode == "manual":
        manual_number = request.form.get("manual_number", "").strip()
        if not manual_number:
            flash("Введите номер средства измерений", "error")
            return render_template("index.html")
        manual_type = request.form.get("manual_type", "").strip()

        import uuid
        task_id = str(uuid.uuid4())

        numbers = [manual_number]

        with lock:
            progress_store[task_id] = {
                "status": "starting",
                "total": 1,
                "current_number": "",
                "current_iteration": 0,
                "progress_pct": 0,
                "cancel": False,
            }

        thread = threading.Thread(
            target=process_numbers,
            args=(numbers, task_id, manual_type if manual_type else None),
            daemon=True
        )
        thread.start()

        return redirect(url_for("results_page", task_id=task_id))

    # Загрузка файла
    if "file" not in request.files:
        flash("Файл не выбран", "error")
        return render_template("index.html")

    file = request.files["file"]
    if file.filename == "":
        flash("Файл не выбран", "error")
        return render_template("index.html")

    if not file.filename.endswith(".xlsx"):
        flash("Поддерживаются только файлы формата .xlsx", "error")
        return render_template("index.html")

    filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(input_path)

    # Читаем номера из Excel
    try:
        wb = openpyxl.load_workbook(input_path)
        ws = wb.active
        numbers = []
        for row in ws.iter_rows(min_row=1, max_col=1, values_only=True):
            val = row[0]
            if val is not None:
                s = str(val).strip()
                if s:
                    numbers.append(s)
        if not numbers:
            flash("В файле не найдено ни одного номера", "error")
            return render_template("index.html")
    except Exception:
        flash("Ошибка при чтении файла", "error")
        return render_template("index.html")

    import uuid
    task_id = str(uuid.uuid4())

    with lock:
        progress_store[task_id] = {
            "status": "starting",
            "total": len(numbers),
            "current_number": "",
            "current_iteration": 0,
            "progress_pct": 0,
            "cancel": False,
        }

    thread = threading.Thread(
        target=process_numbers,
        args=(numbers, task_id),
        daemon=True
    )
    thread.start()

    return redirect(url_for("progress", task_id=task_id))


@app.route("/progress/<task_id>")
def progress(task_id):
    with lock:
        if task_id not in progress_store:
            flash("Задача не найдена", "error")
            return redirect(url_for("index"))
        task = dict(progress_store[task_id])

    return render_template("progress.html", task_id=task_id, task=task)


@app.route("/progress_data/<task_id>")
def progress_data(task_id):
    with lock:
        if task_id not in progress_store:
            return jsonify({"status": "not_found"})
        task = dict(progress_store[task_id])

    return jsonify({
        "status": task.get("status", "unknown"),
        "progress_pct": task.get("progress_pct", 0),
        "current_number": task.get("current_number", ""),
        "current_iteration": task.get("current_iteration", 0),
        "total": task.get("total", 0),
    })


@app.route("/cancel/<task_id>", methods=["POST"])
def cancel_task(task_id):
    with lock:
        if task_id in progress_store:
            progress_store[task_id]["cancel"] = True
    flash("Поиск отменён", "success")
    return redirect(url_for("index"))


@app.route("/results/<task_id>")
def results_page(task_id):
    with lock:
        if task_id not in progress_store:
            flash("Задача не найдена", "error")
            return redirect(url_for("index"))
        task = dict(progress_store[task_id])

    status = task.get("status", "unknown")

    if status == "processing" or status == "starting":
        return render_template("result.html", task_id=task_id, waiting=True, rows=None, stats=None, filename=None)

    if status == "complete":
        rows = task.get("rows", [])
        stats = task.get("stats", {})
        filename = task.get("result_filename", "")
        return render_template("result.html", task_id=task_id, waiting=False, rows=rows, stats=stats, filename=filename)

    if status == "cancelled":
        flash("Поиск был отменён", "error")
        return redirect(url_for("index"))

    flash("Неизвестный статус задачи", "error")
    return redirect(url_for("index"))


@app.route("/download/<filename>")
def download(filename):
    path = os.path.join(RESULT_FOLDER, filename)
    if not os.path.exists(path):
        flash("Файл не найден. Загрузите файл заново.", "error")
        return redirect(url_for("index"))
    return send_file(path, as_attachment=True, download_name="fgis_arshin_result.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)