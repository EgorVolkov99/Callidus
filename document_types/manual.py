from typing import Optional, Dict, Any
from .common import BaseDocModel, DocumentMetadata, CompanyInfo, UnitParameters


class ManualData(BaseDocModel):
    metadata: DocumentMetadata
    company: CompanyInfo
    unit: UnitParameters

    # Динамически заполняемые блоки РЭ
    contents: Optional[bool] = True
    introduction: Optional[Dict[str, Any]] = None
    description_and_operation: Optional[Dict[str, Any]] = None
    usage: Optional[Dict[str, Any]] = None
    maintenance: Optional[Dict[str, Any]] = None
    current_repair: Optional[Dict[str, Any]] = None
    storage: Optional[Dict[str, Any]] = None
    transportation: Optional[Dict[str, Any]] = None
    disposal: Optional[Dict[str, Any]] = None
    revision_sheet: Optional[bool] = True