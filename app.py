if isinstance(docs, dict) and "error" in docs:
                still_pending.append(item)
                with progress_lock:
                    progress_store[task_id]["errors"] += 1
                if idx < len(pending) - 1:
                    if cancellable_sleep(task_id, delay):
                        return
                continue

            if not docs:
                still_pending.append(item)
                with progress_lock:
                    progress_store[task_id]["not_found"] += 1
                if idx < len(pending) - 1:
                    if cancellable_sleep(task_id, delay):
                        return
                continue

            records = process_docs(docs, num, mi_type)
            found_this_iter.extend(records)

            with progress_lock:
                progress_store[task_id]["processed"] += 1

            if idx < len(pending) - 1:
                if cancellable_sleep(task_id, delay):
                    return