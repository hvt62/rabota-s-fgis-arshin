def search_arshin(mi_number):
    """Запрос к API с 3 попытками при 500 ошибке.
       Сначала ищет точное совпадение mi.number, затем по подстроке.
       Возвращает список docs или {"error": ...}."""
    session = get_arshin_session()
    headers = {"Referer": "https://fgis.gost.ru/fundmetrology/cm/results"}

    # Этап 1: точное совпадение mi.number
    params_exact = {
        "fq": [f'mi.number:"{mi_number}"'],
        "q": "*",
        "fl": ",".join(FIELDS),
        "sort": "verification_date desc,org_title asc",
        "rows": 1000,
        "start": 0,
    }

    for attempt in range(1, 4):
        try:
            resp = session.get(ARSHIN_URL, params=params_exact, headers=headers, timeout=60)
            print(f"[API] {mi_number}: точный поиск, попытка {attempt}, статус {resp.status_code}")
            resp.raise_for_status()
            data = resp.json()
            docs = data.get("response", {}).get("docs", [])
            print(f"[API] {mi_number}: точный поиск: найдено {len(docs)} документов")
            if docs:
                return docs
            break  # 200, но пусто — выходим из retry
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else 0
            if status == 500 and attempt < 3:
                print(f"[API] {mi_number}: 500, пауза 5с, попытка {attempt+1}/3")
                time.sleep(5)
                continue
            print(f"[API] {mi_number}: HTTP ошибка {status}")
            return {"error": str(e)}
        except requests.exceptions.RequestException as e:
            print(f"[API] {mi_number}: ошибка {e}")
            if attempt < 3:
                time.sleep(5)
                continue
            return {"error": str(e)}

    # Этап 2: поиск по подстроке mi.number
    params_substr = {
        "fq": [f"mi.number:*{mi_number}*"],
        "q": "*",
        "fl": ",".join(FIELDS),
        "sort": "verification_date desc,org_title asc",
        "rows": 1000,
        "start": 0,
    }

    for attempt in range(1, 4):
        try:
            resp = session.get(ARSHIN_URL, params=params_substr, headers=headers, timeout=60)
            print(f"[API] {mi_number}: поиск по подстроке, попытка {attempt}, статус {resp.status_code}")
            resp.raise_for_status()
            data = resp.json()
            docs = data.get("response", {}).get("docs", [])
            print(f"[API] {mi_number}: поиск по подстроке: найдено {len(docs)} документов")
            return docs
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else 0
            if status == 500 and attempt < 3:
                print(f"[API] {mi_number}: 500, пауза 5с, попытка {attempt+1}/3")
                time.sleep(5)
                continue
            print(f"[API] {mi_number}: HTTP ошибка {status}")
            return {"error": str(e)}
        except requests.exceptions.RequestException as e:
            print(f"[API] {mi_number}: ошибка {e}")
            if attempt < 3:
                time.sleep(5)
                continue
            return {"error": str(e)}

    return {"error": "3 попытки не удались"}