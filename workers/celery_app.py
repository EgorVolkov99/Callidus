from celery import Celery
import os

# Используем Redis в качестве брокера сообщений.
# В продакшене (Docker) адрес будет браться из переменных окружения.
REDIS_URL = os.getenv("REDIS_URL", "redis://localhost:6379/0")

# Создаем инстанс Celery
celery = Celery(
    "callidus_workers",
    broker=REDIS_URL,
    backend=REDIS_URL,
    include=["workers.tasks"]  # Указываем, где искать задачи
)

# Базовые настройки (чтобы задачи не висели вечно)
celery.conf.update(
    task_track_started=True,
    task_time_limit=600,         # Жесткий лимит на задачу (10 минут)
    worker_concurrency=4         # Сколько документов генерировать параллельно
)