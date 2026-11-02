from typing import Dict, Any, Optional, Union
from .common import BaseDocModel, DocumentMetadata, CompanyInfo, UnitParameters

class PassportData(BaseDocModel):
    metadata: DocumentMetadata
    company: CompanyInfo
    unit: UnitParameters

    # Динамические разделы-словари
    basic_info: Optional[Dict[str, Any]] = None
    technical_data: Optional[Dict[str, Any]] = None
    main_parts: Optional[Dict[str, Any]] = None
    approval_tests: Optional[Dict[str, Any]] = None
    scope_of_delivery: Optional[Dict[str, Any]] = None

    # Разделы, которые могут быть словарем, флагом bool (true/false) или null
    warranties: Optional[Union[Dict[str, Any], bool]] = None
    preservation: Optional[Union[Dict[str, Any], bool]] = None
    packing: Optional[Union[Dict[str, Any], bool]] = None
    deviations: Optional[Union[Dict[str, Any], bool]] = None
    conclusion: Optional[Union[Dict[str, Any], bool]] = None
    disposal: Optional[Union[Dict[str, Any], bool]] = None
    terms_list: Optional[Union[Dict[str, Any], bool]] = None
    documents_list: Optional[Union[Dict[str, Any], bool]] = None
    revision_sheet: Optional[bool] = True