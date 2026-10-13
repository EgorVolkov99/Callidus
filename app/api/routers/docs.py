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

# Подключаем оба адаптера
from app.adapters.passport_adapter import prepare_passport_context
from app.adapters.manual_adapter import prepare_manual_context

router = APIRouter()
tasks_status = {}

def process_batch(
        batch_id: str,
        files_data: list,
        doc_type: str,
        template_name: str,
        need_docx: bool,
        need_pdf: bool,
        need_zip: bool,
        pipeline: DocumentPipeline
):
    tasks_status[batch_id]["status"] = "processing"
    generated_files = []

    logger.info(f"Начало обработки batch_id: {batch_id}. Файлов: {len(files_data)}, Тип: {doc_type}")

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

            # Определяем тип документа из метаданных (с фолбеком на запрос)
            actual_doc_type = raw_data.get("metadata", {}).get("type", doc_type)

            brand = unit_info.get("product_brand", "Brand").replace("/", "_")
            out_name = f"{actual_doc_type}_{kks}_{brand}"

            # Чистый роутинг адаптеров
            if actual_doc_type == "manual":
                prepared_data = prepare_manual_context(raw_data)
            elif actual_doc_type == "passport":
                prepared_data = prepare_passport_context(raw_data)
            else:
                prepared_data = raw_data

            if need_docx or need_pdf:
                tasks_status[batch_id]["files"][file_id]["log"] = "Генерация DOCX..."
                logger.debug(f"[{batch_id}] Старт генерации Word для {out_name}")

                docx_path = pipeline.build_word(
                    data=prepared_data,
                    doc_type=actual_doc_type,
                    template_name=template_name,
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
        doc_type: str = Form(...),
        template_name: str = Form(...),
        need_docx: bool = Form(False),
        need_pdf: bool = Form(False),
        need_zip: bool = Form(False),
        files: List[UploadFile] = File(...),
        pipeline: DocumentPipeline = Depends(get_pipeline)
):
    batch_id = str(uuid.uuid4())
    logger.info(f"=== НОВЫЙ ЗАПРОС НА ГЕНЕРАЦИЮ (Batch: {batch_id}) ===")
    logger.info(f"Параметры: doc_type={doc_type}, DOCX={need_docx}, PDF={need_pdf}, ZIP={need_zip}")

    files_data = []
    file_statuses = {}

    try:
        ModelClass = PassportData if doc_type == "passport" else ManualData
        logger.debug(f"Выбрана модель валидации: {ModelClass.__name__}")
    except Exception as e:
        logger.error(f"Ошибка при выборе модели: {e}")
        raise HTTPException(status_code=500, detail="Внутренняя ошибка выбора модели")

    for idx, file in enumerate(files):
        file_id = f"file_{idx}"
        try:
            content = await file.read()
            logger.debug(f"Файл {file.filename} прочитан. Размер: {len(content)} байт.")

            raw_data = json.loads(content)

            if hasattr(ModelClass, 'model_validate'):
                validated_data = ModelClass.model_validate(raw_data)
                dumped_data = validated_data.model_dump()
            else:
                validated_data = ModelClass.parse_obj(raw_data)
                dumped_data = validated_data.dict()

            files_data.append({"filename": file.filename, "content": dumped_data})
            file_statuses[file_id] = {"name": file.filename, "status": "В очереди", "log": "Ожидание..."}

            logger.info(f"Файл {file.filename} успешно прошел Pydantic-валидацию.")

        except Exception as e:
            error_trace = traceback.format_exc()
            logger.error(f"Ошибка валидации файла {file.filename}:\n{error_trace}")
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
            process_batch, batch_id, files_data, doc_type, template_name, need_docx, need_pdf, need_zip, pipeline
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