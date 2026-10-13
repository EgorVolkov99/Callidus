import logging
from pathlib import Path
from pydantic import BaseModel
from typing import Dict, Any

from .builders.tex_builder import TexBuilder
from .builders.docx_builder import DocxBuilder
from .converters.tex2pdf import TexToPdfConverter
from .converters.docx2pdf import DocxToPdfConverter


class DocumentPipeline:
    """Единый конвейер сборки документов."""

    def __init__(self, templates_base_dir: Path, output_base_dir: Path):
        self.templates_dir = templates_base_dir
        self.output_dir = output_base_dir

    def build_word(self, data: Dict[str, Any], doc_type: str, template_name: str, out_filename: str) -> Path:
        """Сборка Word документа."""
        template_path = self.templates_dir / doc_type / "word" / template_name
        output_path = self.output_dir / "docx" / f"{out_filename}.docx"

        builder = DocxBuilder(template_path)
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
            converter = DocxToPdfConverter()
        else:
            raise ValueError("Неподдерживаемый формат для конвертации.")

        return converter.convert(source_path, output_pdf)