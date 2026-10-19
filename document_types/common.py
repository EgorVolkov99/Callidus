from pydantic import BaseModel, Field, ConfigDict
from typing import Optional, List, Union

class BaseDocModel(BaseModel):
    # Настройка защиты от падения, если в JSON придут лишние ключи
    model_config = ConfigDict(extra="ignore", populate_by_name=True)

class DocumentMetadata(BaseDocModel):
    # Базовые данные документа
    type: str = Field(default="")
    document_name_rus: str = Field(default="")
    document_name_eng: str = Field(default="")
    document_designation: str = Field(default="")
    document_kks: str = Field(default="")
    master_title: Optional[str] = "base"
    language: Optional[str] = "rus-eng"

class CompanyInfo(BaseDocModel):
    # Данные организации
    okpd_code: Optional[str] = ""
    company_full_name_rus: Optional[str] = ""
    company_full_name_eng: Optional[str] = ""
    company_short_name_rus: Optional[str] = ""
    company_short_name_eng: Optional[str] = ""
    # Даты и подписи для штампов
    designer_rus: Optional[str] = None
    designer_eng: Optional[str] = None
    design_date: Optional[str] = None
    inspector_rus: Optional[str] = None
    inspector_eng: Optional[str] = None
    inspection_date: Optional[str] = None
    approver_rus: Optional[str] = None
    approver_eng: Optional[str] = None
    approve_date: Optional[str] = None
    registration_date: Optional[str] = None
    inv_num: Optional[str] = None

class UnitParameters(BaseDocModel):
    # Обязательно заполняемые параметры
    kks_code: Union[List[str], str] = Field(default="!!! НЕ НАЗНАЧЕН ККС КОД !!!")
    product_name_rus: str = Field(default="!!! НЕ НАЗНАЧЕНО НАИМЕНОВАНИЕ ОБОРУДОВАНИЯ/ИЗДЕЛИЯ !!!")
    product_name_eng: str = Field(default="!!! NO NAME OF THE EQUIPMENT/PRODUCT ASSIGNED !!!")
    # Дополнительные параметры
    product_system_rus: Optional[str] = ""
    product_system_eng: Optional[str] = ""
    product_building_rus: Optional[str] = ""
    product_building_eng: Optional[str] = ""
    product_name_2_rus: Optional[str] = ""
    product_name_2_eng: Optional[str] = ""
    product_brand: Union[List[str], str, None] = ""
    product_model: Union[List[str], str, None] = ""
    climate: Optional[str] = ""
    placement: Optional[str] = ""
    manual_code: Optional[str] = ""
    rr1_code: Optional[str] = ""
    rr2_code: Optional[str] = ""
    sb_code: Optional[str] = ""
    tz_code: Optional[str] = ""
    warranty_text_rus: Optional[str] = ""
    warranty_text_eng: Optional[str] = ""