import sys
import subprocess
from pathlib import Path


class DocxToPdfConverter:
    def convert(self, docx_path: Path, output_pdf: Path) -> Path:
        output_pdf.parent.mkdir(parents=True, exist_ok=True)

        # Если Windows - используем нативный MS Word[cite: 10]
        if sys.platform == "win32":
            try:
                from docx2pdf import convert
                convert(str(docx_path.resolve()), str(output_pdf.resolve()))
                return output_pdf
            except ImportError:
                raise RuntimeError("Библиотека docx2pdf не установлена.")

        # Для Linux / Docker используем LibreOffice
        else:
            cmd = [
                "soffice", "--headless", "--convert-to", "pdf",
                "--outdir", str(output_pdf.parent), str(docx_path)
            ]
            subprocess.run(cmd, check=True)

            # LibreOffice сохраняет с тем же именем, переименовываем, если нужно
            generated_pdf = output_pdf.parent / docx_path.with_suffix('.pdf').name
            if generated_pdf != output_pdf:
                generated_pdf.rename(output_pdf)

            return output_pdf