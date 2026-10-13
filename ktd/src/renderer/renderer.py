"""
renderer.py

Универсальный движок генерации LaTeX.

Назначение:

    JSON
        ↓
    Page
        ↓
    Jinja2
        ↓
    *.tex

Renderer не знает ничего о КТД.

Он получает:

    • шаблон;
    • данные;

и создает tex-файл.

Подходит абсолютно для любых форм
(ТЛ, ВТД, МК, ВМ, КТП, ВОК и т.д.).
"""

from pathlib import Path

from jinja2 import Environment
from jinja2 import FileSystemLoader

from paths import TEMPLATES_DIR
from paths import ProjectPaths
from renderer.latex import latex_escape
from renderer.latex import tex_newline


# ==========================================================
# Renderer
# ==========================================================

class Renderer:

    def __init__(self, paths: ProjectPaths):

        self.paths = paths

        self.environment = Environment(

            loader=FileSystemLoader(

                TEMPLATES_DIR

            ),

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

        self.environment.filters["latex"] = latex_escape

        self.environment.filters["newline"] = tex_newline

    # ======================================================

    def render_template(

            self,

            template_name: str,

            output_name: str,

            **context

    ) -> Path:

        """
        Рендер одного шаблона.
        """

        template = self.environment.get_template(

            template_name

        )

        text = template.render(

            **context

        )

        output_path = self.paths.get_build_tex(

            output_name

        )

        output_path.write_text(

            text,

            encoding="utf8"

        )

        return output_path

    # ======================================================

    def render_pages(

            self,

            pages,

            template_name: str,

            **context

    ) -> list[str]:

        """
        Генерация набора страниц одной формы.

        Возвращает список имен tex-файлов,
        который затем используется compiler.py.
        """

        tex_files = []

        form_prefix = pages[0].form_code.lower()

        for page in pages:

            filename = (

                f"{form_prefix}_"

                f"{page.page_number:03d}.tex"

            )

            page_context = {

                **context,

                **page.extra

            }

            self.render_template(

                template_name=template_name,

                output_name=filename,

                page=page,

                **page_context

            )

            tex_files.append(

                filename

            )

        return tex_files

    # ======================================================

    def render_single(

            self,

            template_name: str,

            output_name: str,

            **context

    ) -> str:

        """
        Генерация одного tex-файла.

        Используется для:

            main.tex

            preamble.tex

            Любых одиночных документов.
        """

        self.render_template(

            template_name,

            output_name,

            **context

        )

        return output_name
