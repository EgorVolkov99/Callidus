from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Any


@dataclass
class FormPage:
    """
    Одна страница формы после пагинации.
    """

    form_code: str
    template_name: str

    page_number: int
    total_pages: int

    is_first: bool
    is_last: bool
    has_next: bool

    context: dict[str, Any]


class BaseForm(ABC):
    """
    Базовый класс любой формы КТД.
    """

    form_code: str = ""
    template_name: str = ""

    @abstractmethod
    def paginate(self) -> list[FormPage]:
        """
        Возвращает список страниц формы.
        """
        pass