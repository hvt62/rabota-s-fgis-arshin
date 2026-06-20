import os
import requests
import openpyxl
from io import BytesIO
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

# Новый рабочий endpoint ФГИС «Аршин»
ARSHIN_URL = "https://fgis.gost.ru/fundmetrology/cm/xcdb/vri/select"

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


def search_arshin(mi_number):
    """Поиск сведений о поверке по номеру СИ через API Аршин."""
    params = {
        "fq": f"*{mi_number}*",
        "q": "*",
        "fl": ",".join(FIELDS),
        "sort": "verification_date desc,org_title asc",
        "rows": 20,
        "start": 0,
    }
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json, text/plain, */*",
        "Referer": "https://fgis.gost.ru/fundmetrology/cm/",
    }

    try:
        resp = requests.get(ARSHIN_URL, params=params, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        docs = data.get("response", {}).get("docs", [])
        return docs
    except requests.exceptions.RequestException as e:
        return {"error": str(e)}


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

        # Собираем номера из первого столбца (начиная со второй строки)
        numbers = []
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            val = row[0]
            if val is not None:
                numbers.append(str(val).strip())

        if not numbers:
            return render_template("index.html", error="Нет номеров в первом столбце")

        # Опрашиваем API для каждого номера
        results = []
        errors = []
        for num in numbers:
            docs = search_arshin(num)
            if isinstance(docs, dict) and "error" in docs:
                errors.append({"number": num, "error": docs["error"]})
            elif docs:
                for doc in docs:
                    results.append(
                        {
                            "number": num,
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
                results.append(
                    {
                        "number": num,
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

        return render_template(
            "result.html", results=results, errors=errors, total=len(numbers)
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