"""
main.py

Главный модуль генерации комплекта технологической документации.

====================================================================

Общий алгоритм работы

    Проект КТД
            │
            ▼
    ktd_structure.json
            │
            ▼
    Определение состава комплекта документации
            │
            ▼
    Последовательная обработка каждой формы
            │
            ▼
        JSON формы
            │
            ▼
        Form.paginate()
            │
            ▼
        Renderer.render_pages()
            │
            ▼
        *.tex
            │
            ▼
    Compiler.build()
            │
            ▼
          PDF

====================================================================

Файл main.py не знает особенностей конкретных форм.

Он только:

    • читает структуру проекта;
    • загружает соответствующий JSON;
    • вызывает нужный класс формы;
    • запускает Renderer;
    • запускает Compiler.
"""

from pathlib import Path
import json
import sys


if hasattr(sys.stdout, "reconfigure"):

    sys.stdout.reconfigure(encoding="utf-8")

if hasattr(sys.stderr, "reconfigure"):

    sys.stderr.reconfigure(encoding="utf-8")


if __package__ in (None, ""):

    sys.path.insert(

        0,

        str(Path(__file__).resolve().parent.parent)

    )

from src.paths import ProjectPaths

from src.models import (
    KTDStructure,
    TitleDocument,
    VTDDocument
)

from src.forms.title import TitleForm
from src.forms.vtd import VTDForm

from src.renderer.renderer import Renderer
from src.renderer.compiler import Compiler


# ==========================================================
# Соответствие кодов форм их обработчикам
# ==========================================================

FORM_REGISTRY = {

    "TL": {

        "model": TitleDocument,

        "form": TitleForm,

        "template": "tl/tl.tex.j2"

    },

    "VTD": {

        "model": VTDDocument,

        "form": VTDForm,

        "template": "vtd/vtd.tex.j2"

    }

}


# ==========================================================
# Загрузка JSON
# ==========================================================

def load_json(path: Path):

    """
    Загружает JSON-файл.

    Parameters
    ----------
    path : Path

        Путь к JSON.

    Returns
    -------
    dict
    """

    with open(

            path,

            encoding="utf8"

    ) as f:

        return json.load(f)


# ==========================================================
# Главная функция
# ==========================================================

def main():

    # ------------------------------------------------------
    # Папка проекта КТД
    # ------------------------------------------------------

    project_directory = Path(

        "projects/ASOS Desalination Plant"

    )

    paths = ProjectPaths(

        project_directory

    )

    paths.create()

    paths.clear_build()

    # ------------------------------------------------------
    # Загрузка структуры комплекта документации
    # ------------------------------------------------------

    structure = KTDStructure(

        **load_json(

            paths.get_temp_json(

                "ktd_structure.json"

            )

        )

    )

    # ------------------------------------------------------
    # Инициализация Renderer
    # ------------------------------------------------------

    renderer = Renderer(

        paths

    )

    generated_tex_files = []

    # ------------------------------------------------------
    # Последовательная генерация всех форм
    # ------------------------------------------------------

    for item in structure.render_order:

        form_code = item.form.upper()

        equipment = item.equipment

        print(

            f"Формируется {form_code} ({equipment})..."

        )

        registry = FORM_REGISTRY[

            form_code

        ] if form_code in FORM_REGISTRY else None

        if registry is None:

            available_forms = ", ".join(

                sorted(FORM_REGISTRY)

            )

            raise ValueError(

                f"Форма {form_code} указана в ktd_structure.json, "
                f"но не зарегистрирована. Доступные формы: {available_forms}."

            )

        # ----------------------------------------------
        # Имя JSON-файла формы
        # ----------------------------------------------

        json_filename = (

            f"{form_code.lower()}_"

            f"{equipment}.json"

        )

        json_path = paths.get_temp_json(

            json_filename

        )

        # ----------------------------------------------
        # Загрузка данных формы
        # ----------------------------------------------

        model = registry["model"](

            **load_json(

                json_path

            )

        )

        # ----------------------------------------------
        # Постраничное разбиение
        # ----------------------------------------------

        form = registry["form"](model)

        pages = form.paginate()

        # ----------------------------------------------
        # Генерация TeX
        # ----------------------------------------------

        tex_files = renderer.render_pages(

            pages=pages,

            template_name=registry["template"],

            document=model

        )

        generated_tex_files.extend(

            tex_files

        )

    # ------------------------------------------------------
    # Компиляция комплекта документации
    # ------------------------------------------------------

    compiler = Compiler(

        paths

    )

    compiler.build(

        generated_tex_files

    )

    print()

    print("----------------------------------------")

    print("Комплект КТД успешно сформирован.")

    print()

    print(

        f"PDF: {paths.output_pdf}"

    )

    print(

        f"TEX: {paths.output_tex}"

    )

    print(

        f"LOG: {paths.output_reports}"

    )

    print("----------------------------------------")


# ==========================================================
# Точка входа
# ==========================================================

if __name__ == "__main__":

    main()
