from jinja2 import Environment, FileSystemLoader
from pathlib import Path
from engine.utils.sanitizer import sanitize_latex_string

def get_latex_env(templates_dir: Path) -> Environment:
    """Среда Jinja2 с синтаксисом под LaTeX."""
    env = Environment(
        loader=FileSystemLoader(str(templates_dir)),
        variable_start_string="\\VAR{",
        variable_end_string="}",
        block_start_string="\\BLOCK{",
        block_end_string="}",
        comment_start_string="\\COMMENT{",
        comment_end_string="}",
        trim_blocks=True,
        lstrip_blocks=True,
    )
    env.filters["latex"] = sanitize_latex_string
    return env

def get_word_env(templates_dir: Path) -> Environment:
    """Стандартная среда Jinja2 для Word (опционально, docxtpl использует свою, но полезна для рендера кусков)."""
    return Environment(loader=FileSystemLoader(str(templates_dir)))