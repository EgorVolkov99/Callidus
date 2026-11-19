import logging
from pathlib import Path
from typing import Dict, Any

from .builders.tex_builder import TexBuilder
from engine.builders.modular_builder import ModularDocxBuilder
from .converters.tex2pdf import TexToPdfConverter
from .converters.docx2pdf import PdfConverter

logger = logging.getLogger("callidus")

class DocumentPipeline:
    """Единый конвейер сборки документов."""

    def __init__(self, templates_base_dir: Path, output_base_dir: Path):
        self.templates_dir = templates_base_dir
        self.output_dir = output_base_dir

    def build_word(self, data: Dict[str, Any], doc_type: str, template_name: str, out_filename: str) -> Path:
        """Сборка Word документа."""# Получаем корневую директорию проекта
        BASE_DIR = Path(__file__).resolve().parent.parent
        TEMPLATE_DIR = BASE_DIR / "templates"
        output_path = self.output_dir / "docx" / f"{out_filename}.docx"

        builder = ModularDocxBuilder(TEMPLATE_DIR)
        # Примечание: Здесь данные уже должны содержать готовые subdocs,
        # либо мы передаем инстанс builder-а в адаптер (ViewModel), чтобы он встроил таблицы.
        return builder.render_and_save(data, output_path)

    def build_latex(self, data: Dict[str, Any], doc_type: str, template_name: str, out_filename: str) -> Path:
        """Сборка LaTeX документа."""
        template_dir = self.templates_dir / doc_type / "latex"
        output_path = self.output_dir / "tex" / f"{out_filename}.tex"

        builder = TexBuilder(template_dir)
        return builder.build(template_name, data, output_path)

    def convert_to_pdf(self, source_path: Path) -> Path:
        """Автоматически определяет формат и конвертирует в PDF."""
        output_pdf = self.output_dir / "pdf" / f"{source_path.stem}.pdf"

        if source_path.suffix == ".tex":
            converter = TexToPdfConverter()
        elif source_path.suffix == ".docx":
            converter = PdfConverter()
        else:
            raise ValueError("Неподдерживаемый формат для конвертации.")

        return converter.convert(source_path, output_pdf)