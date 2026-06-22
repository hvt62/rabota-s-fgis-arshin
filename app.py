if 500 <= status < 600 and attempt < MAX_RETRIES:
                # Для 500 ошибок ждём: 7.5, 15, 22.5, ... 75 сек
                wait = 7.5 * attempt
                print(f"[API] {mi_number}: 5xx, пауза {wait}с, попытка {attempt}/{MAX_RETRIES}")
                time.sleep(wait)
                continue