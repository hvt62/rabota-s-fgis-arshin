def cancellable_sleep(task_id, seconds):
    """Сон с проверкой отмены каждую секунду."""
    for _ in range(int(seconds)):
        with progress_lock:
            if progress_store[task_id].get("cancel"):
                return True
        time.sleep(1)
    return False