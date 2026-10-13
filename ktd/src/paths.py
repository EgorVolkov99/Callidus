"""
paths.py

Работа с путями проекта.

Назначение:
    • формирование структуры каталогов проекта;
    • создание отсутствующих директорий;
    • предоставление единых путей всем модулям программы.

Структура проекта КТД:

project/
│
├── input/
├── temp/
├── build/
│
└── output/
    ├── pdf/
    ├── tex/
    └── reports/

Весь проект строится относительно папки,
переданной пользователем в main.py.
"""

from pathlib import Path


# ==========================================================
# Пути программы
# ==========================================================

ROOT_DIR = Path(__file__).resolve().parent.parent

SRC_DIR = ROOT_DIR / "src"

TEMPLATES_DIR = ROOT_DIR / "templates"

PROMPTS_DIR = ROOT_DIR / "prompts"

STANDARDS_DIR = ROOT_DIR / "standards"


# ==========================================================
# Класс проекта
# ==========================================================

class ProjectPaths:
    """
    Хранит все пути конкретного проекта.

    Один экземпляр класса соответствует
    одному комплекту КТД.
    """

    def __init__(self, project_directory: str | Path):

        project_directory = Path(project_directory)

        if not project_directory.is_absolute():

            project_directory = ROOT_DIR / project_directory

        self.project_dir = project_directory

        # -----------------------
        # Исходные данные
        # -----------------------

        self.input = self.project_dir / "input"

        # -----------------------
        # JSON после работы LLM
        # -----------------------

        self.temp = self.project_dir / "temp"

        # -----------------------
        # Рабочая папка LaTeX
        # -----------------------

        self.build = self.project_dir / "build"

        # -----------------------
        # Финальные результаты
        # -----------------------

        self.output = self.project_dir / "output"

        self.output_pdf = self.output / "pdf"

        self.output_tex = self.output / "tex"

        self.output_reports = self.output / "reports"

    # ======================================================

    def create(self):

        """
        Создает структуру каталогов проекта.
        """

        folders = [

            self.input,

            self.temp,

            self.build,

            self.output,

            self.output_pdf,

            self.output_tex,

            self.output_reports

        ]

        for folder in folders:

            folder.mkdir(

                parents=True,

                exist_ok=True

            )

    # ======================================================

    def clear_build(self):

        """
        Полностью очищает build.

        Используется перед новой сборкой.
        """

        if not self.build.exists():

            return

        for file in self.build.iterdir():

            if file.is_file():

                file.unlink()

    # ======================================================

    def get_temp_json(self, filename: str) -> Path:
        """
        Возвращает путь к JSON в temp/.
        """

        return self.temp / filename

    # ======================================================

    def get_build_tex(self, filename: str) -> Path:
        """
        Возвращает путь к TeX-файлу в build/.
        """

        return self.build / filename

    # ======================================================

    def get_output_pdf(self, filename: str) -> Path:
        """
        Возвращает путь к PDF.
        """

        return self.output_pdf / filename

    # ======================================================

    def get_output_tex(self, filename: str) -> Path:
        """
        Возвращает путь к итоговому TeX.
        """

        return self.output_tex / filename
