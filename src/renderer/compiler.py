"""
compiler.py

Компиляция комплекта КТД.

Назначение:

    *.tex
        ↓
    main.tex
        ↓
    XeLaTeX
        ↓
    PDF

Compiler не знает ничего о формах КТД.
Он работает только с готовыми tex-файлами.
"""

from pathlib import Path
import shutil
import subprocess
import datetime

from src.paths import (
    ProjectPaths,
    TEMPLATES_DIR
)

from jinja2 import (
    Environment,
    FileSystemLoader
)


class Compiler:
    """
    Компиляция проекта LaTeX.
    """

    def __init__(self, paths: ProjectPaths):

        self.paths = paths

        self.environment = Environment(

            loader=FileSystemLoader(TEMPLATES_DIR),

            block_start_string="((*",
            block_end_string="*))",

            variable_start_string="(((",
            variable_end_string=")))",

            comment_start_string="((=",
            comment_end_string="=))",

            autoescape=False,
            trim_blocks=True,
            lstrip_blocks=True

        )

    # ======================================================
    # Создание main.tex
    # ======================================================

    def create_main(self, tex_files: list[str]):

        """
        Создает главный файл main.tex.
        """

        template = self.environment.get_template(

            "main.tex.j2"

        )

        text = template.render(

            tex_files=tex_files

        )

        output = self.paths.build / "main.tex"

        output.write_text(

            text,

            encoding="utf8"

        )

    # ======================================================
    # Копирование общих файлов
    # ======================================================

    def copy_common_files(self):

        """
        Копирует общие tex-файлы
        в build.
        """

        common = TEMPLATES_DIR / "common"

        build_common = self.paths.build / "common"

        build_common.mkdir(

            parents=True,

            exist_ok=True

        )

        shutil.copy(

            common / "preamble.tex",

            self.paths.build / "preamble.tex"

        )

        shutil.copy(

            common / "macros.tex",

            self.paths.build / "macros.tex"

        )

        shutil.copy(

            common / "table_styles.tex",

            self.paths.build / "table_styles.tex"

        )

        shutil.copy(

            common / "header_first.tex.j2",

            build_common / "header_first.tex"

        )

        shutil.copy(

            common / "header_next.tex.j2",

            build_common / "header_next.tex"

        )

        shutil.copy(

            common / "footer.tex.j2",

            build_common / "footer.tex"

        )

    # ======================================================
    # Компиляция
    # ======================================================

    def compile(self):

        """
        Компиляция main.tex.
        """

        command = [

            "xelatex",

            "-interaction=nonstopmode",

            "main.tex"

        ]

        result = subprocess.run(

            command,

            cwd=self.paths.build,

            capture_output=True,

            text=True,

            encoding="utf-8",

            errors="replace"

        )

        self._write_log(result)

        if result.returncode != 0:

            raise RuntimeError(

                "Ошибка компиляции XeLaTeX.\n"
                "Подробности см. в generation.log"

            )

    # ======================================================
    # Копирование результатов
    # ======================================================

    def copy_results(self):

        """
        Переносит результаты
        в output.
        """

        pdf = self.paths.build / "main.pdf"

        if pdf.exists():

            shutil.copy(

                pdf,

                self.paths.output_pdf / "main.pdf"

            )

        for tex in self.paths.build.glob("*.tex"):

            shutil.copy(

                tex,

                self.paths.output_tex / tex.name

            )

    # ======================================================
    # Создание лога
    # ======================================================

    def _write_log(self, result):

        log = self.paths.output_reports / "generation.log"

        text = []

        text.append(

            "===== GENERATION LOG ====="

        )

        text.append(

            f"Date: {datetime.datetime.now()}"

        )

        text.append("")

        text.append(

            "----- STDOUT -----"

        )

        text.append(

            result.stdout or ""

        )

        text.append("")

        text.append(

            "----- STDERR -----"

        )

        text.append(

            result.stderr or ""

        )

        log.write_text(

            "\n".join(text),

            encoding="utf8"

        )

    # ======================================================
    # Полная сборка
    # ======================================================

    def build(self, tex_files: list[str]):

        """
        Полный цикл сборки.

            main.tex

                ↓

            XeLaTeX

                ↓

            output/
        """

        self.copy_common_files()

        self.create_main(tex_files)

        self.compile()

        self.copy_results()
