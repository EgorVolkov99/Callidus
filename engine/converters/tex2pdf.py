import subprocess
import logging
import shutil
from pathlib import Path

class TexToPdfConverter:
    def __init__(self, timeout: int = 120):
        self.timeout = timeout
        self.cmd = "xelatex" #[cite: 9]

    def convert(self, tex_file: Path, output_pdf: Path) -> Path:
        working_dir = tex_file.parent
        output_pdf.parent.mkdir(parents=True, exist_ok=True)

        cmd = [self.cmd, "-interaction=nonstopmode", f"-output-directory={working_dir}", str(tex_file.name)]

        # 2 прохода для оглавления[cite: 7]
        for run_num in range(1, 3):
            try:
                result = subprocess.run(
                    cmd, cwd=working_dir, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    text=True, encoding="utf-8", errors="replace", timeout=self.timeout
                )
                if result.returncode != 0:
                    logging.error(f"Ошибка XeLaTeX:\n{result.stdout}\n{result.stderr}")
                    raise RuntimeError("Сбой компиляции LaTeX.")
            except subprocess.TimeoutExpired:
                raise RuntimeError("Тайм-аут компиляции XeLaTeX.")

        generated_pdf = working_dir / tex_file.with_suffix(".pdf").name
        if generated_pdf.exists():
            shutil.move(str(generated_pdf), str(output_pdf))
            return output_pdf
        raise FileNotFoundError("PDF файл не найден после компиляции.")