"""
vtd.py

Разбиение ведомости технологических документов.

Назначение:

    VTDDocument
            │
            ▼
    list[Page]

Модуль не занимается генерацией TeX.

Он только делит документ
на страницы согласно СТО.
"""

from src.models import (

    VTDDocument,

    Page

)

# ==========================================================
# Константы
# ==========================================================

FIRST_PAGE_ROWS = 19

NEXT_PAGE_ROWS = 22


# ==========================================================
# VTDForm
# ==========================================================

class VTDForm:

    """
    Формирование страниц ВТД.
    """

    def __init__(self, document: VTDDocument):

        self.document = document

    # ======================================================

    def paginate(self) -> list[Page]:

        """
        Разбивает ВТД
        на страницы.
        """

        documents = self.document.documents

        pages = []

        index = 0

        page_number = 1

        while index < len(documents):

            if page_number == 1:

                limit = FIRST_PAGE_ROWS

            else:

                limit = NEXT_PAGE_ROWS

            current = documents[index:index + limit]

            pages.append(

                Page(

                    form_name="ВТД-1" if page_number == 1 else "ВТД-2",

                    form_code="ВТД",

                    page_number=page_number,

                    total_pages=0,

                    is_first=(page_number == 1),

                    has_next=False,

                    rows=current,

                    extra={

                        "info": self.document.info

                    }

                )

            )

            index += limit

            page_number += 1

        if not pages:

            pages.append(

                Page(

                    form_name="ВТД-1",

                    form_code="VTD",

                    page_number=1,

                    total_pages=1,

                    is_first=True,

                    has_next=False,

                    rows=[],

                    extra={

                        "info": self.document.info

                    }

                )

            )

        total = len(pages)

        for i, page in enumerate(pages):

            page.total_pages = total

            page.has_next = i < total - 1

        return pages