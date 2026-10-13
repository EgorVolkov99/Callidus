from __future__ import annotations

LATEX_SPECIALS = {
    "&": r"\&",
    "%": r"\%",
    "$": r"\$",
    "#": r"\#",
    "_": r"\_",
    "{": r"\{",
    "}": r"\}",
    "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
}


def latex_escape(text: str | None) -> str:
    """
    Экранирование специальных символов LaTeX.
    """

    if text is None:
        return ""

    result = str(text)

    for key, value in LATEX_SPECIALS.items():
        result = result.replace(key, value)

    return result


def tex_newline(text: str | None) -> str:
    """
    Заменяет переносы строк на LaTeX-перенос.
    """

    if text is None:
        return ""

    return latex_escape(text).replace("\n", r"\\")