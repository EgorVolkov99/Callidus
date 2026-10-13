from pathlib import Path
from engine.pipeline import DocumentPipeline

# Определяем абсолютные пути корня проекта
BASE_DIR = Path(__file__).resolve().parent.parent.parent
TEMPLATES_DIR = BASE_DIR / "templates"
STORAGE_DIR = BASE_DIR / "storage"

def get_pipeline() -> DocumentPipeline:
    """Зависимость (Dependency) для получения настроенного движка генерации."""
    return DocumentPipeline(
        templates_base_dir=TEMPLATES_DIR,
        output_base_dir=STORAGE_DIR
    )