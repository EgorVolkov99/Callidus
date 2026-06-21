"""
models.py

Структуры данных проекта генерации КТД.

Назначение:
    • описание структуры всех промежуточных JSON;
    • валидация данных, полученных от LLM;
    • единый интерфейс для всех модулей программы.

Все модели основаны на Pydantic.
"""

from typing import List, Optional
from pydantic import BaseModel, Field


# ==========================================================
# Общая информация о документе
# ==========================================================

class DocumentInfo(BaseModel):
    """
    Общая информация о форме.
    Используется практически всеми документами.
    """

    developer_organization: str

    equipment_name: str

    document_title: str

    document_designation: str

    document_set_designation: str

    revision: str = ""


# ==========================================================
# Структура проекта КТД
# ==========================================================

class ProjectInfo(BaseModel):
    """
    Информация о проекте.
    """

    project_id: str

    project_name: str

    generated_at: str

    standard: str


class DocumentSet(BaseModel):
    """
    Информация о комплекте документации.
    """

    designation: str

    title: str


class EquipmentStructure(BaseModel):
    """
    Описание единицы оборудования,
    входящей в комплект КТД.
    """

    id: str

    designation: str

    name: str

    documents: List[str]


class RenderItem(BaseModel):
    """
    Один пункт очереди генерации.
    """

    equipment: str

    form: str


class KTDStructure(BaseModel):
    """
    Полное содержимое файла ktd_structure.json
    """

    project: ProjectInfo

    document_set: DocumentSet

    equipment: List[EquipmentStructure]

    render_order: List[RenderItem]


# ==========================================================
# Титульный лист
# ==========================================================

class Approval(BaseModel):

    role: str

    organization: str = ""

    position: str

    fio: str

    date: str = ""

    signature: str = ""


class TitleInfo(DocumentInfo):
    """
    Информация, используемая только титульным листом.
    """

    department_name: str

    developer_full_name: str

    documentation_name: str

    equipment_designation: str

    technological_processes: str

    effective_date: str


class TitleDocument(BaseModel):
    """
    Полное содержимое tl_xxxxx.json
    """

    equipment_id: str

    form: str

    info: TitleInfo

    approvals: List[Approval]


# ==========================================================
# ВТД
# ==========================================================

class VTDRow(BaseModel):
    """
    Одна строка ведомости технологических документов.
    """

    graph: str

    designation: str

    code: str

    name: str

    sheet_count: int

    first_sheet: int

    remark: str = ""


class VTDDocument(BaseModel):
    """
    Полное содержимое файла vtd_xxxxx.json
    """

    equipment_id: str

    form: str

    info: DocumentInfo

    documents: List[VTDRow]


# ==========================================================
# Универсальная страница после пагинации
# ==========================================================

class Page(BaseModel):
    """
    Универсальная модель страницы.

    Создается Python после разбиения документа
    на отдельные листы.

    Именно она передается в Jinja2.
    """

    form_name: str

    form_code: str

    page_number: int

    total_pages: int

    is_first: bool

    has_next: bool

    rows: list = Field(default_factory=list)

    extra: dict = Field(default_factory=dict)