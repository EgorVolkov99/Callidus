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

НОВАЯ СТРУКТУРА:
- content: список единиц оборудования/СЕ/Д
- Каждая СЕД содержит sets: список комплектов
- Каждый комплект содержит docs: список документов
- Иерархия: СЕД → комплект → документы
"""

from models import (

    VTDDocument,

    Page,

    ContentItem

)

from typing import Tuple

# ==========================================================
# Константы
# ==========================================================

FIRST_PAGE_ROWS = 19

NEXT_PAGE_ROWS = 22

# Размеры ячеек в мм
CELL_WIDTHS = {
    "graph": 14,
    "npp": 16,
    "v_designation": 22,
    "code": 31.5,
    "name": 161.5,
    "sheets": 20,
    "remark": 36
}

# ==========================================================
# Вспомогательные функции
# ==========================================================

def calculate_line_count(text: str, max_width_mm: float, font_size: int = 8) -> int:
    """
    Рассчитывает количество строк, которые займет текст.
    
    text - исходный текст
    max_width_mm - максимальная ширина ячейки в мм
    font_size - размер шрифта в pts
    
    Возвращает количество строк
    """
    if not text:
        return 1
    
    # Упрощенная оценка: примерно 0.35 мм на символ для шрифта 8pt
    chars_per_line = int(max_width_mm / 0.35)
    
    # Разбиваем текст на слова
    words = text.split()
    if not words:
        return 1
    
    lines = 1
    current_line_len = 0
    
    for word in words:
        word_len = len(word) + 1  # +1 для пробела
        if current_line_len + word_len > chars_per_line:
            lines += 1
            current_line_len = word_len
        else:
            current_line_len += word_len
    
    return lines


def calculate_row_height(graph: str, designation: str, code: str, 
                         name: str, sheet_count: int, remark: str) -> int:
    """
    Рассчитывает высоту строки в ячейках.
    
    Для СЕД: graph, designation, name (3 ячейки)
    Для комплекта: graph, designation, name, sheet_count, remark (5 ячеек)
    Для документа: graph, designation, code, name, sheet_count, remark (6 ячеек)
    """
    max_lines = 1
    
    # Определяем, какие тексты обрабатывать в зависимости от типа строки
    if graph.startswith("В"):
        # СЕД: 3 ячейки
        texts = [graph, "_", designation, name]
        widths = [CELL_WIDTHS["graph"], CELL_WIDTHS["npp"], CELL_WIDTHS["v_designation"], CELL_WIDTHS["name"]]
    elif graph.startswith("Ф"):
        # Комплект: 5 ячеек
        texts = [graph, designation, name, str(sheet_count), remark]
        widths = [CELL_WIDTHS["graph"], CELL_WIDTHS["f_designation"], CELL_WIDTHS["name"],
                  CELL_WIDTHS["sheets"], CELL_WIDTHS["remark"]]
    elif graph.startswith("Г"):
        # Документ: 6 ячеек
        texts = [graph, designation, code, name, str(sheet_count), remark]
        widths = [CELL_WIDTHS["graph"], CELL_WIDTHS["g_designation"], CELL_WIDTHS["code"],
                  CELL_WIDTHS["name"], CELL_WIDTHS["sheets"], CELL_WIDTHS["remark"]]
    else:
        return 1
    
    # Для каждой ячейки рассчитываем количество строк
    for text, width in zip(texts, widths):
        if text:
            lines = calculate_line_count(text, width)
            max_lines = max(max_lines, lines)
    
    return max_lines


def create_vtd_row(graph: str, designation: str, code: str, 
                   name: str, sheet_count: int, first_sheet: int, remark: str) -> dict:
    """
    Создает словарь для одной строки ВТД.
    """
    return {
        "graph": graph,
        "designation": designation,
        "code": code,
        "name": name,
        "sheet_count": sheet_count,
        "first_sheet": first_sheet,
        "remark": remark,
        "height": calculate_row_height(graph, designation, code, name, sheet_count, remark)
    }


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
        
        Иерархия: content (СЕД) -> sets (комплекты) -> docs (документы)
        """

        pages = []
        current_rows = []
        
        remaining_lines = FIRST_PAGE_ROWS  # 19 для первой страницы
        page_number = 1
        
        npp = 1  # НПП для СЕД
        
        # Обходим все единицы оборудования/СЕД
        for equipment in self.document.content:
            # Генерируем строки для одной СЕД
            equipment_rows, equipment_height = self._process_equipment(equipment, npp)
            
            # Проверяем, помещается ли СЕД на текущую страницу
            if equipment_height > remaining_lines:
                # Создаем страницу и начинаем новую
                page = self._create_page(current_rows, page_number, page_number == 1)
                pages.append(page)
                current_rows = []
                remaining_lines = NEXT_PAGE_ROWS  # 22 для следующих страниц
                page_number += 1
            
            current_rows.extend(equipment_rows)
            remaining_lines -= equipment_height
            npp += 1
        
        # Добавляем последнюю страницу
        if current_rows:
            page = self._create_page(current_rows, page_number, page_number == 1)
            pages.append(page)
        
        # Если страниц нет, создаем пустую
        if not pages:
            pages.append(
                Page(
                    form_name="ВТД-1",
                    form_code="ВТД",
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

    # ======================================================

    def _process_equipment(self, equipment: ContentItem, npp: int) -> Tuple[list, int]:
        """
        Генерирует строки для одной единицы оборудования/СЕД.
        
        Возвращает:
        - список строк
        - общую высоту в ячейках
        """
        rows = []
        total_height = 0
        
        # Строка СЕД (graph="В")
        row = create_vtd_row(
            graph=f"В{npp:02d}",
            designation=equipment.designation,
            code="",
            name=equipment.name,
            sheet_count=0,
            first_sheet=1,
            remark=""
        )
        rows.append(row)
        total_height += row["height"]
        
        # Обработка комплектов
        for set_idx, set_item in enumerate(equipment.sets, 1):
            # Рассчитываем общее количество листов в комплекте
            total_sheets = sum(doc.sheet_count for doc in set_item.docs)
            
            # Строка комплекта (graph="Ф")
            row = create_vtd_row(
                graph=f"Ф{set_idx}",
                designation=set_item.designation,
                code="",
                name=set_item.name,
                sheet_count=total_sheets,
                first_sheet=1,
                remark=set_item.remark
            )
            rows.append(row)
            total_height += row["height"]
            
            # Обработка документов
            for doc_idx, doc_item in enumerate(set_item.docs, 1):
                row = create_vtd_row(
                    graph=f"Г{doc_idx}",
                    designation=doc_item.designation,
                    code=doc_item.code,
                    name=doc_item.name,
                    sheet_count=doc_item.sheet_count,
                    first_sheet=1,
                    remark=doc_item.remark
                )
                rows.append(row)
                total_height += row["height"]
        
        return rows, total_height

    # ======================================================

    def _create_page(self, rows: list, page_number: int, is_first: bool) -> Page:
        """
        Создает объект Page из списка строк.
        """
        return Page(
            form_name="ВТД-1" if is_first else "ВТД-2",
            form_code="ВТД",
            page_number=page_number,
            total_pages=0,
            is_first=is_first,
            has_next=False,
            rows=rows,
            extra={
                "info": self.document.info
            }
        )