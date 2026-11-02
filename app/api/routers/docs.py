import json
import uuid
import zipfile
import traceback
from typing import List
from fastapi import APIRouter, UploadFile, File, Form, BackgroundTasks, HTTPException, Depends

from document_types.passport import PassportData
from document_types.manual import ManualData
from engine.pipeline import DocumentPipeline
from app.api.dependencies import get_pipeline, STORAGE_DIR
from app.core.logger import logger

# Подключаем адаптеры
from app.adapters.passport_adapter import prepare_passport_context
from app.adapters.manual_adapter import prepare_manual_context

router = APIRouter()
tasks_status = {}


def process_batch(
        batch_id: str,
        files_data: list,
        need_docx: bool,
        need_pdf: bool,
        need_zip: bool,
        pipeline: DocumentPipeline
):
    tasks_status[batch_id]["status"] = "processing"
    generated_files = []

    logger.info(f"Начало обработки batch_id: {batch_id}. Файлов: {len(files_data)}")

    for idx, f_data in enumerate(files_data):
        raw_data = f_data["content"]
        file_id = f"file_{idx}"
        tasks_status[batch_id]["files"][file_id]["status"] = "В процессе..."

        try:
            tasks_status[batch_id]["files"][file_id]["log"] = "Анализ данных..."
            logger.debug(f"[{batch_id}] Обработка файла {idx + 1}...")

            unit_info = raw_data.get("unit", {})
            kks = unit_info.get("kks_code", "NO_KKS")
            if isinstance(kks, list) and kks:
                kks = kks[0]

            # Теперь тип документа и мастер-шаблон всегда берутся из самого JSON
            actual_doc_type = raw_data.get("metadata", {}).get("type")
            master_title = raw_data.get("metadata", {}).get("master_title", "base")

            doc_name = raw_data.get("metadata", {}).get("document_name_rus")
            doc_designation = raw_data.get("metadata", {}).get("document_designation").replace("/", "_")
            product = unit_info.get("product_name_rus").replace("/", "_")
            brand = unit_info.get("product_brand").replace("/", "_")
            out_name = f"{doc_name} {doc_designation}. {product} {brand}"

            # Чистый роутинг адаптеров (поддержка ВЭ и будущих форматов)
            if actual_doc_type == "manual":
                prepared_data = prepare_manual_context(raw_data)
            elif actual_doc_type == "passport":
                prepared_data = prepare_passport_context(raw_data)
            else:
                # Если это ВЭ ("ve") или другой тип, для которого еще нет адаптера, передаем как есть
                prepared_data = raw_data

            def clean_nulls(d):
                if isinstance(d, dict):
                    return {k: clean_nulls(v) for k, v in d.items()}
                elif isinstance(d, list):
                    return [clean_nulls(v) for v in d]
                elif d is None:
                    return ""
                return d

            # Очищаем весь контекст от null перед генерацией
            prepared_data = clean_nulls(prepared_data)

            if "unit" not in prepared_data:
                prepared_data["unit"] = {}

            kks_val = prepared_data["unit"].get("kks_code")
            if not kks_val:
                # Если кода нет, передаем пустой список (join отработает без ошибок и вернет пустоту)
                prepared_data["unit"]["kks_code"] = []
            elif isinstance(kks_val, str):
                # Если код - это строка, оборачиваем ее в список, чтобы join не разбил ее по буквам
                prepared_data["unit"]["kks_code"] = [kks_val]

            if need_docx or need_pdf:
                tasks_status[batch_id]["files"][file_id]["log"] = "Генерация DOCX..."
                logger.debug(f"[{batch_id}] Старт генерации Word для {out_name}")

                # В pipeline передаем master_title вместо старого template_name
                docx_path = pipeline.build_word(
                    data=prepared_data,
                    doc_type=actual_doc_type,
                    template_name=master_title,
                    out_filename=out_name
                )
                if need_docx:
                    generated_files.append(docx_path)

            if need_pdf:
                tasks_status[batch_id]["files"][file_id]["log"] = "Конвертация в PDF..."
                logger.debug(f"[{batch_id}] Конвертация в PDF для {out_name}")

                pdf_path = pipeline.convert_to_pdf(source_path=docx_path)
                generated_files.append(pdf_path)

                if not need_docx and docx_path.exists():
                    docx_path.unlink()

            tasks_status[batch_id]["files"][file_id]["status"] = "✅ Готово"
            tasks_status[batch_id]["files"][file_id]["log"] = "Успешно завершено."
            tasks_status[batch_id]["completed"] += 1

            logger.info(f"[{batch_id}] Файл {out_name} успешно сгенерирован.")

        except Exception as e:
            error_trace = traceback.format_exc()
            logger.error(f"[{batch_id}] ОШИБКА генерации файла {idx + 1}:\n{error_trace}")

            tasks_status[batch_id]["files"][file_id]["status"] = "❌ Ошибка"
            tasks_status[batch_id]["files"][file_id]["log"] = str(e)
            tasks_status[batch_id]["errors"] += 1

    if len(generated_files) == 0:
        logger.warning(f"[{batch_id}] Ни одного файла не сгенерировано. Отмена создания ZIP.")
        tasks_status[batch_id]["status"] = "completed"
        tasks_status[batch_id]["global_log"] = "Генерация прервана из-за ошибок. Проверьте логи."
        return

    if need_zip or len(files_data) > 1:
        tasks_status[batch_id]["global_log"] = "Упаковка в ZIP архив..."
        zip_dir = STORAGE_DIR / "archives"
        zip_dir.mkdir(parents=True, exist_ok=True)
        zip_path = zip_dir / f"Batch_{batch_id[:8]}.zip"

        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for f in generated_files:
                    if f.exists():
                        zipf.write(f, f.name)

            tasks_status[batch_id]["archive_url"] = f"/archives/{zip_path.name}"
            logger.info(f"[{batch_id}] ZIP архив успешно создан: {zip_path.name}")
        except Exception as e:
            logger.error(f"[{batch_id}] Ошибка при создании ZIP архива: {e}")
            tasks_status[batch_id]["global_log"] = f"Ошибка упаковки ZIP: {e}"
    else:
        for f in generated_files:
            if f.suffix == '.pdf':
                tasks_status[batch_id]["pdf_url"] = f"/downloads/{batch_id}/{f.name}"
            elif f.suffix == '.docx':
                tasks_status[batch_id]["docx_url"] = f"/downloads/{batch_id}/{f.name}"

    tasks_status[batch_id]["status"] = "completed"
    tasks_status[batch_id]["global_log"] = "Все задачи завершены."


@router.post("/generate")
async def generate_batch(
        background_tasks: BackgroundTasks,
        project_name: str = Form("NIKIMT_KIP"),
        master_title: str = Form("nikimt"),  # Новое поле из UI (Шаблон)
        language: str = Form("rus-eng"),  # Новое поле из UI (Язык)
        need_docx: bool = Form(False),  # Form(False) спасает от 422 ошибки
        need_pdf: bool = Form(False),
        need_zip: bool = Form(False),
        files: List[UploadFile] = File(...),
        pipeline: DocumentPipeline = Depends(get_pipeline)
):
    batch_id = str(uuid.uuid4())
    logger.info(f"=== НОВЫЙ ЗАПРОС НА ГЕНЕРАЦИЮ (Batch: {batch_id}) ===")
    logger.info(f"UI Параметры: master_title={master_title}, language={language}, DOCX={need_docx}, PDF={need_pdf}")

    files_data = []
    file_statuses = {}

    for idx, file in enumerate(files):
        file_id = f"file_{idx}"
        try:
            content = await file.read()
            raw_data = json.loads(content)

            # 1. ПРИНУДИТЕЛЬНОЕ ПЕРЕОПРЕДЕЛЕНИЕ ПАРАМЕТРОВ ИЗ ИНТЕРФЕЙСА
            if "metadata" not in raw_data:
                raw_data["metadata"] = {}
            raw_data["metadata"]["master_title"] = master_title
            raw_data["metadata"]["language"] = language

            # 2. ДИНАМИЧЕСКОЕ ОПРЕДЕЛЕНИЕ ТИПА ИЗ JSON
            doc_type = raw_data["metadata"].get("type", "passport")

            # 3. УМНЫЙ ВЫБОР МОДЕЛИ ВАЛИДАЦИИ
            if doc_type == "passport":
                ModelClass = PassportData
            elif doc_type == "manual":
                ModelClass = ManualData
            else:
                ModelClass = None  # Для новых типов (например, "ve"), пока для них нет Pydantic-модели

            if ModelClass:
                if hasattr(ModelClass, 'model_validate'):
                    validated_data = ModelClass.model_validate(raw_data)
                    dumped_data = validated_data.model_dump()
                else:
                    validated_data = ModelClass.parse_obj(raw_data)
                    dumped_data = validated_data.dict()
                logger.info(f"Файл {file.filename} успешно прошел Pydantic-валидацию как {doc_type}.")
            else:
                dumped_data = raw_data
                logger.info(f"Файл {file.filename} загружен без Pydantic-валидации (тип: {doc_type}).")

            files_data.append({"filename": file.filename, "content": dumped_data})
            file_statuses[file_id] = {"name": file.filename, "status": "В очереди", "log": "Ожидание..."}

        except Exception as e:
            error_trace = traceback.format_exc()
            logger.error(f"Ошибка чтения/валидации файла {file.filename}:\n{error_trace}")
            file_statuses[file_id] = {"name": file.filename, "status": "❌ Ошибка валидации", "log": str(e)}

    tasks_status[batch_id] = {
        "status": "starting",
        "total": len(files_data),
        "completed": 0,
        "errors": len(files) - len(files_data),
        "files": file_statuses,
        "global_log": "Инициализация...",
        "archive_url": None,
        "pdf_url": None,
        "docx_url": None
    }

    if files_data:
        logger.info(f"Передача Batch {batch_id} в фоновую обработку BackgroundTasks...")
        background_tasks.add_task(
            process_batch, batch_id, files_data, need_docx, need_pdf, need_zip, pipeline
        )
    else:
        logger.warning(f"Batch {batch_id} отменен: нет валидных файлов.")
        tasks_status[batch_id]["status"] = "completed"
        tasks_status[batch_id]["global_log"] = "Нет валидных файлов для обработки."

    return {"batch_id": batch_id}


@router.get("/status/{batch_id}")
async def get_status(batch_id: str):
    if batch_id not in tasks_status:
        raise HTTPException(status_code=404, detail="Batch not found")
    return tasks_status[batch_id]