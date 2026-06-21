"""
title.py

Разбиение титульного листа на страницы.

Назначение:

    TitleDocument
            │
            ▼
    list[Page]

Данный модуль не занимается генерацией TeX.

Его задача — подготовить данные
для шаблона tl.tex.j2.
"""

from src.models import (
    TitleDocument,
    Page
)

# ==========================================================
# Константы
# ==========================================================

# Количество подписей,
# помещающихся на первом листе ТЛ

FIRST_PAGE_SIGNATURES = 3

# Количество подписей
# на каждом последующем листе

NEXT_PAGE_SIGNATURES = 4


# ==========================================================
# TitleForm
# ==========================================================

class TitleForm:

    """
    Формирует страницы титульного листа.
    """

    def __init__(self, document: TitleDocument):

        self.document = document

    # ======================================================

    def paginate(self) -> list[Page]:

        """
        Разбивает титульный лист
        на необходимое количество страниц.
        """

        approvals = self.document.approvals

        pages = []

        index = 0

        page_number = 1

        while index < len(approvals):

            if page_number == 1:

                limit = FIRST_PAGE_SIGNATURES

            else:

                limit = NEXT_PAGE_SIGNATURES

            current = approvals[index:index + limit]

            pages.append(

                Page(

                    form_name="ТЛ-1г" if page_number == 1 else "ТЛ-2г",

                    form_code="TL",

                    page_number=page_number,

                    total_pages=0,

                    is_first=(page_number == 1),

                    has_next=False,

                    rows=[],

                    extra={

                        "info": self.document.info,

                        "approvals": current

                    }

                )

            )

            index += limit

            page_number += 1

        if not pages:

            pages.append(

                Page(

                    form_name="ТЛ-1г",

                    form_code="TL",

                    page_number=1,

                    total_pages=1,

                    is_first=True,

                    has_next=False,

                    rows=[],

                    extra={

                        "info": self.document.info,

                        "approvals": []

                    }

                )

            )

        total = len(pages)

        for i, page in enumerate(pages):

            page.total_pages = total

            page.has_next = i < total - 1

        return pages