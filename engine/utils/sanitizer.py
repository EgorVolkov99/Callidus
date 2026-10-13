from typing import Any

LATEX_ESCAPES = {
    '\\': r'\textbackslash{}',
    '%': r'\%',
    '$': r'\$',
    '#': r'\#',
    '_': r'\_',
    '&': r'\&',
    '{': r'\{',
    '}': r'\}',
    '~': r'\textasciitilde{}',
    '^': r'\textasciicircum{}',
}

def sanitize_latex_string(value: Any) -> str:
    """
    Фильтр Jinja2 для безопасного экранирования спецсимволов LaTeX.
    Принимает любой тип данных, принудительно приводит к строке и экранирует.
    """
    if value is None:
        return ""
    text = str(value)
    for char, escape in LATEX_ESCAPES.items():
        text = text.replace(char, escape)
    return text