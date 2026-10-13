import json
import shutil
from pathlib import Path
from fastapi import APIRouter, HTTPException, UploadFile, File

router = APIRouter()

# Путь к папке с пользовательскими данными
WORKSPACE_DIR = Path(__file__).resolve().parent.parent.parent.parent / "workspace"


@router.get("/")
async def list_projects():
    """Возвращает список всех существующих проектов."""
    if not WORKSPACE_DIR.exists():
        WORKSPACE_DIR.mkdir(parents=True)
        return {"projects": []}

    projects = [d.name for d in WORKSPACE_DIR.iterdir() if d.is_dir()]
    return {"projects": projects}


@router.post("/{project_name}")
async def create_project(project_name: str):
    """Создает новое рабочее пространство для проекта."""
    project_path = WORKSPACE_DIR / project_name

    if project_path.exists():
        raise HTTPException(status_code=400, detail="Проект с таким именем уже существует.")

    # Создаем базовую структуру проекта
    (project_path / "inputs").mkdir(parents=True)
    (project_path / "assets").mkdir(parents=True)

    # Создаем дефолтный конфиг
    config_path = project_path / "config.json"
    default_config = {
        "company_name": "",
        "default_designer": "",
        "logo_filename": None
    }
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(default_config, f, ensure_ascii=False, indent=4)

    return {"message": f"Проект {project_name} успешно создан."}


@router.get("/{project_name}/config")
async def get_project_config(project_name: str):
    """Получает специфичные настройки проекта (штампы, авторы, реквизиты)."""
    config_path = WORKSPACE_DIR / project_name / "config.json"

    if not config_path.exists():
        raise HTTPException(status_code=404, detail="Конфигурация проекта не найдена.")

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    return config


@router.post("/{project_name}/upload-image")
async def upload_project_image(project_name: str, file: UploadFile = File(...)):
    """Загружает картинку (схему, логотип) в ассеты конкретного проекта."""
    assets_dir = WORKSPACE_DIR / project_name / "assets"

    if not assets_dir.exists():
        raise HTTPException(status_code=404, detail="Проект не найден.")

    file_path = assets_dir / file.filename

    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    return {"message": "Изображение загружено", "filename": file.filename}