for iteration, delay in enumerate(delays, 1):
        if not pending:
            break

        # Сбрасываем счётчики сбоев и не найденных для текущей итерации
        with progress_lock:
            progress_store[task_id]["errors"] = 0
            progress_store[task_id]["not_found"] = 0

        print(f"\n[App] === Итерация {iteration}, пауза {delay}с, осталось {len(pending)} номеров ===")

        found_this_iter = []
        still_pending = []

        for idx, item in enumerate(pending):
            # Проверка отмены
            with progress_lock:
                if progress_store[task_id].get("cancel"):
                    print(f"[App] {task_id}: получен сигнал отмены")
                    finalize_results(task_id, rows_data, results, errors_list)
                    return

            num = item["number"]
            mi_type = item["type"] if item["type"] else ""

            docs = search_arshin(num, mi_type if mi_type else None)

            if isinstance(docs, dict) and "error" in docs:
                still_pending.append(item)
                with progress_lock:
                    progress_store[task_id]["errors"] += 1
                if idx < len(pending) - 1:
                    time.sleep(delay)
                continue

            if not docs:
                still_pending.append(item)
                with progress_lock:
                    progress_store[task_id]["not_found"] += 1
                if idx < len(pending) - 1:
                    time.sleep(delay)
                continue

            records = process_docs(docs, num, mi_type)
            found_this_iter.extend(records)

            with progress_lock:
                progress_store[task_id]["processed"] += 1

            if idx < len(pending) - 1:
                time.sleep(delay)

        results.extend(found_this_iter)
        pending = still_pending

        with progress_lock:
            progress_store[task_id]["pending_items"] = pending