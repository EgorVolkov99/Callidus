import logging
from pathlib import Path
from typing import Dict, Any

from .celery_app import celery
from engine.pipeline import DocumentPipeline

# Предположим, что у нас есть адаптер для подготовки данных
# from app.adapters.passport_adapter import PassportAdapter

# Задаем пути к нужным директориям
BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = BASE_DIR / "templates"
STORAGE_DIR = BASE_DIR / "storage"


@celery.task(bind=True, name="generate_passport_task")
def generate_passport_task(self, raw_json_data: Dict[str, Any], project_name: str, file_name: str,
                           use_word: bool = True):
    """
    Фоновая задача для генерации паспорта.
    self.update_state позволяет отправлять на фронтенд прогресс (например, "20% - Компиляция PDF").
    """
    try:
        self.update_state(state='PROGRESS', meta={'message': 'Инициализация движка...'})

        # 1. Инициализируем конвейер
        pipeline = DocumentPipeline(templates_base_dir=TEMPLATES_DIR, output_base_dir=STORAGE_DIR)

        # 2. Здесь должен отработать адаптер, который валидирует Pydantic модель
        # и готовит Subdoc-таблицы. Для примера считаем, что data уже готова.
        data = raw_json_data

        self.update_state(state='PROGRESS', meta={'message': 'Сборка исходного файла...'})

        # 3. Выбор движка генерации
        if use_word:
            source_file = pipeline.build_word(
                data=data,
                doc_type="passport",
                template_name="passport_full_rus-eng.docx",
                out_filename=file_name
            )
        else:
            source_file = pipeline.build_latex(
                data=data,
                doc_type="passport",
                template_name="main.tex.j2",
                out_filename=file_name
            )

        self.update_state(state='PROGRESS', meta={'message': 'Конвертация в PDF...'})

        # 4. Конвертация
        final_pdf = pipeline.convert_to_pdf(source_path=source_file)

        # Успешное завершение: возвращаем путь к файлу (строкой) для выдачи пользователю
        return {
            "status": "success",
            "pdf_url": f"/api/downloads/{final_pdf.name}",
            "docx_url": f"/api/downloads/{source_file.name}" if use_word else None
        }

    except Exception as e:
        logging.error(f"Ошибка в задаче генерации: {str(e)}")
        # При ошибке статус задачи автоматически станет FAILURE,
        # а этот текст ошибки можно будет показать пользователю.
        raise Exception(f"Сбой генерации: {str(e)}")