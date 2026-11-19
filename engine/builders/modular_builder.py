import io
import logging
from pathlib import Path
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from engine.builders.table_factory import TableFactory

logger = logging.getLogger("callidus")


class ModularDocxBuilder:
    def __init__(self, templates_dir: Path):
        self.templates_dir = templates_dir

    def render_and_save(self, context: dict, output_path: Path):
        metadata = context.get("metadata", {})
        doc_type = metadata.get("type", {})

        # 1. КОНФИГУРАЦИЯ СБОРКИ (Читаем из metadata)
        master_name = metadata.get("master_title", "base")
        language_option = metadata.get("language", "rus-eng")

        # Строим пути согласно новой структуре ms_word/...
        ms_word_dir = self.templates_dir / "ms_word"
        master_path = ms_word_dir / "masters" / f"{master_name}_{language_option}.docx"
        body_path = ms_word_dir / "bodies" / f"{doc_type}_{language_option}.docx"
        sections_dir = ms_word_dir / "sections"

        # Защита: если указанный мастер-шаблон не найден, откатываемся к базовому
        if not master_path.exists():
            logger.warning(f"Мастер-шаблон {master_path.name} не найден. Используем base_{language_option}.docx")
            master_path = ms_word_dir / "masters" / f"base_{language_option}.docx"

        # Полный перечень всех возможных секций для сборки
        all_sections = [
            "contents", "basic_info", "technical_data", "main_parts", "approval_tests",
            "scope_of_delivery", "warranties", "preservation", "packing",
            "deviations", "conclusion", "disposal", "terms_list", "documents_list", "revision_sheet",
            "introduction", "description_and_operation", "usage", "maintenance",
            "current_repair", "storage", "transportation", "ve",
            "repair_docs", "zik_repair", "zk_repair", "mk_repair",
            "repair_introduction", "repair_organization", "repair_safety", "repair_requirements",
            "repair_process", "repair_replacement", "repair_assembly", "repair_testing",
            "repair_installation", "repair_coatings", "repair_marking", "repair_packaging",
            "tk_introduction", "tk_technical_requirements", "tk_repair_cycle",
            "tk_special_requirements", "tk_quality_control", "tk_testing_and_acceptance",
            "tk_coatings_and_lubrication", "tk_completeness", "tk_marking_and_packaging", "tk_warranties",
            "tk_appendix_a", "tk_appendix_b", "tk_appendix_v", "tk_appendix_g", "tk_appendix_d",
            "tk_appendix_e", "tk_appendix_zh", "tk_appendix_i",
            "toir_vd", "toir_op", "toir_po", "toir_ksc", "toir_ko", "toir_vr",
            "kttd_vtd", "kttd_kti", "kttd_kttp", "kttd_vd", "kttd_vok", "kttd_kk", "kttd_vo", "kttd_vm"
        ]

        logger.info(f"[ModularBuilder] Инициализация скелета ({master_path.name})...")

        # ==============================================================
        # ЭТАП 1: РЕНДЕРИНГ ОТДЕЛЬНЫХ РАЗДЕЛОВ (SECTIONS) В ПАМЯТИ
        # ==============================================================
        body_tpl = DocxTemplate(body_path)
        rendered_sections = {}

        # Активируем сборку файлов Приложений, если для них есть данные в ТТ
        tk_reqs = context.get("tk_technical_requirements", {})
        if tk_reqs.get("defectation_card_data"):
            context["tk_appendix_a"] = True
        if tk_reqs.get("replaced_parts_data"):
            context["tk_appendix_b"] = True
        if tk_reqs.get("clearances_data"):
            context["tk_appendix_v"] = True
        if tk_reqs.get("torques_data"):
            context["tk_appendix_g"] = True
        if tk_reqs.get("materials_data"):
            context["tk_appendix_d"] = True
        if tk_reqs.get("measuring_instruments_data"):
            context["tk_appendix_e"] = True

        for sec_key in all_sections:
            if context.get(sec_key):

                # 1. Пытаемся найти специфичный для шаблона файл (например, nikimt_...)
                sec_file = sections_dir / doc_type / f"{master_name}_{sec_key}_{language_option}.docx"
                if not sec_file.exists():
                    sec_file = sections_dir / "repair" / f"{master_name}_{sec_key}_{language_option}.docx"
                if not sec_file.exists():
                    sec_file = sections_dir / "common" / f"{master_name}_{sec_key}_{language_option}.docx"

                # 2. ФОЛБЭК: Если кастомного файла нет, берем базовый (base_...)
                if not sec_file.exists():
                    logger.debug(f"Кастомный {master_name}_{sec_key} не найден. Используем base_...")
                    sec_file = sections_dir / doc_type / f"base_{sec_key}_{language_option}.docx"
                    if not sec_file.exists():
                        sec_file = sections_dir / "repair" / f"base_{sec_key}_{language_option}.docx"
                    if not sec_file.exists():
                        sec_file = sections_dir / "common" / f"base_{sec_key}_{language_option}.docx"

                if sec_file.exists():
                    logger.info(f"[ModularBuilder] Сборка раздела: {sec_key}")
                    sec_tpl = DocxTemplate(sec_file)
                    sec_factory = TableFactory(sec_tpl, master_name=master_name)

                    sec_context = context.copy()

                    self._inject_signatures(sec_context, sec_tpl)

                    # Инъекция таблиц строго в нужный раздел
                    self._inject_tables_for_section(sec_key, sec_context, sec_factory, doc_type)

                    sec_tpl.render(sec_context)

                    # Сохраняем раздел в поток
                    stream = io.BytesIO()
                    sec_tpl.save(stream)
                    stream.seek(0)

                    rendered_sections[f"sec_{sec_key}"] = body_tpl.new_subdoc(stream)
                else:
                    logger.warning(f"[ModularBuilder] Файл раздела не найден: {sec_key}_{language_option}.docx")

        # ==============================================================
        # ЭТАП 2: СБОРКА ТЕЛА С ОГЛАВЛЕНИЕМ (BODY)
        # ==============================================================
        logger.info(f"[ModularBuilder] Сборка тела документа ({body_path.name})...")
        body_context = context.copy()
        body_context.update(rendered_sections)

        body_tpl.render(body_context)

        body_stream = io.BytesIO()
        body_tpl.save(body_stream)
        body_stream.seek(0)

        # ==============================================================
        # ЭТАП 3: ФИНАЛЬНАЯ ВСТАВКА В ТИТУЛЬНИК (MASTER)
        # ==============================================================
        logger.info(f"[ModularBuilder] Вставка в титульный лист ({master_path.name})...")
        master_tpl = DocxTemplate(master_path)
        master_context = context.copy()
        master_context["main_body"] = master_tpl.new_subdoc(body_stream)

        self._inject_signatures(master_context, master_tpl)

        master_tpl.render(master_context)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        master_tpl.save(output_path)
        logger.info(f"[ModularBuilder] Модульная сборка успешно завершена: {output_path.name}")
        return output_path

    def _inject_tables_for_section(self, sec_key, context, factory, doc_type):
        """Инъекция таблиц строго в тот раздел, где они нужны"""

        if sec_key == "basic_info":
            context["basic_info_table_rus"] = factory.build_basic_info_table(
                context.get("basic_info", {}).get("rus", []))
            context["basic_info_table_eng"] = factory.build_basic_info_table(
                context.get("basic_info", {}).get("eng", []))

        elif sec_key in ["technical_data", "description_and_operation"]:
            tech_data = context.get("technical_data")
            if not tech_data:
                tech_data = context.get("description_and_operation", {}).get("unit", {}).get("specifications", {})

            context["technical_data_table_rus"] = factory.build_grouped_table(
                tech_data.get("rus", {}), ["Наименование параметра", "Значение параметра"])
            context["technical_data_table_eng"] = factory.build_grouped_table(
                tech_data.get("eng", {}), ["Parameter name", "Parameter value"])

        elif sec_key == "main_parts":
            mp = context.get("main_parts", {})
            context["materials_table_rus"] = factory.build_materials_table(mp.get("materials", {}).get("rus", {}),
                                                                           is_eng=False)
            context["materials_table_eng"] = factory.build_materials_table(mp.get("materials", {}).get("eng", {}),
                                                                           is_eng=True)
            context["welds_table_rus"] = factory.build_welds_table(mp.get("welds", {}).get("rus", {}), is_eng=False)
            context["welds_table_eng"] = factory.build_welds_table(mp.get("welds", {}).get("eng", {}), is_eng=True)
            context["valves_table_rus"] = factory.build_valves_table(mp.get("valves", {}).get("rus", {}), is_eng=False)
            context["valves_table_eng"] = factory.build_valves_table(mp.get("valves", {}).get("eng", {}), is_eng=True)

        elif sec_key == "approval_tests":
            scope = context.get("approval_tests", {})
            context["approval_tests_table_rus"] = factory.build_approval_tests_table(
                scope.get("rus_word_rows", []), context.get("unit", {}), scope.get("has_extra_columns", False),
                is_eng=False)
            context["approval_tests_table_eng"] = factory.build_approval_tests_table(
                scope.get("eng_word_rows", []), context.get("unit", {}), scope.get("has_extra_columns", False),
                is_eng=True)

        elif sec_key == "scope_of_delivery":
            scope = context.get("scope_of_delivery", {})
            context["delivery_set_table_rus"] = factory.build_generic_table(
                ["Обозначение", "Наименование", "Серийный номер", "Код KKS", "Кол-во"], scope.get("units_word_rus", []),
                [3.5, 5.0, 4.0, 3.5, 2.0])
            context["delivery_set_table_eng"] = factory.build_generic_table(
                ["Designation", "Description", "Serial No", "KKS code", "Qty"], scope.get("units_word_eng", []),
                [3.5, 5.0, 4.0, 3.5, 2.0])
            context["documentation_set_table_rus"] = factory.build_generic_table(
                ["Обозначение", "Наименование", "Кол-во", "Примечание"], scope.get("docs_word_rus", []),
                [3.5, 7.5, 2.0, 5.0])
            context["documentation_set_table_eng"] = factory.build_generic_table(
                ["Designation", "Description", "Qty", "Note"], scope.get("docs_word_eng", []), [3.5, 7.5, 2.0, 5.0])

        elif sec_key == "usage":
            usage = context.get("usage", {})
            context["manual_limitations_table_rus"] = factory.build_generic_table(
                ["Параметр", "Предельное значение", "Последствия несоблюдения"], usage.get("limitations_flat_rus", []),
                [5.5, 4.5, 8.0])
            context["manual_limitations_table_eng"] = factory.build_generic_table(
                ["Parameter", "Limit value", "Consequences"], usage.get("limitations_flat_eng", []), [5.5, 4.5, 8.0])
            context["manual_prep_troubleshooting_table_rus"] = factory.build_generic_table(
                ["Описание неисправности", "Действия по устранению"], usage.get("prep_troubleshoot_flat_rus", []),
                [6.0, 12.0])
            context["manual_prep_troubleshooting_table_eng"] = factory.build_generic_table(
                ["Malfunction", "Actions to eliminate"], usage.get("prep_troubleshoot_flat_eng", []), [6.0, 12.0])
            context["manual_op_troubleshooting_table_rus"] = factory.build_generic_table(
                ["Описание неисправности", "Действия по устранению"], usage.get("op_troubleshoot_flat_rus", []),
                [6.0, 12.0])
            context["manual_op_troubleshooting_table_eng"] = factory.build_generic_table(
                ["Malfunction", "Actions to eliminate"], usage.get("op_troubleshoot_flat_eng", []), [6.0, 12.0])

        elif sec_key == "maintenance":
            maint = context.get("maintenance", {})
            context["manual_fuel_table_rus"] = factory.build_generic_table(
                ["Наименование применяемых ГСМ"], maint.get("fuel_flat_rus", []), [18.0])
            context["manual_fuel_table_eng"] = factory.build_generic_table(
                ["Description of applied fuels and lubricants"], maint.get("fuel_flat_eng", []), [18.0])
            context["manual_maint_procedure_table_rus"] = factory.build_generic_table(
                ["Наименование объекта ТО и работы", "Виды ТО", "Примечание"], maint.get("procedure_flat_rus", []),
                [7.0, 5.0, 6.0])
            context["manual_maint_procedure_table_eng"] = factory.build_generic_table(
                ["Name of the maintenance object and work", "Maintenance types", "Note"],
                maint.get("procedure_flat_eng", []), [7.0, 5.0, 6.0])
            context["manual_operability_check_table_rus"] = factory.build_generic_table(
                ["Наименование работы", "Кто выполняет", "Средства измерений и материалы",
                 "Контрольные значения параметров"], maint.get("check_flat_rus", []), [5.0, 3.0, 5.0, 5.0])
            context["manual_operability_check_table_eng"] = factory.build_generic_table(
                ["Work name", "Executor", "Measuring instruments and materials", "Control values"],
                maint.get("check_flat_eng", []), [5.0, 3.0, 5.0, 5.0])

        elif sec_key == "current_repair":
            current_repair = context.get("current_repair", {})
            context["manual_current_repair_table_rus"] = factory.build_generic_table(
                ["Описание отказа", "Последствия", "Возможные причины", "Способ обнаружения", "Способ устранения"],
                current_repair.get("repair_flat_rus", []), [3.5, 3.5, 3.5, 3.5, 4.0])
            context["manual_current_repair_table_eng"] = factory.build_generic_table(
                ["Failure description", "Consequences", "Possible causes", "Detection method", "Elimination method"],
                current_repair.get("repair_flat_eng", []), [3.5, 3.5, 3.5, 3.5, 4.0])

        elif sec_key == "storage":
            storage = context.get("storage", {})
            context["manual_limited_shelf_life_table_rus"] = factory.build_generic_table(
                ["Наименование детали", "Срок годности", "Примечание"], storage.get("shelf_life_flat_rus", []),
                [6.0, 4.0, 8.0])
            context["manual_limited_shelf_life_table_eng"] = factory.build_generic_table(
                ["Part name", "Shelf life", "Note"], storage.get("shelf_life_flat_eng", []), [6.0, 4.0, 8.0])

        elif sec_key == "disposal":
            disposal = context.get("disposal", {})
            context["manual_calculated_disposal_table_rus"] = factory.build_generic_table(
                ["Наименование детали", "Материал", "Масса, кг", "Вид переработки"],
                disposal.get("calculated_flat_rus", []), [5.0, 5.0, 3.0, 5.0])
            context["manual_calculated_disposal_table_eng"] = factory.build_generic_table(
                ["Part name", "Material", "Weight, kg", "Recycling type"], disposal.get("calculated_flat_eng", []),
                [5.0, 5.0, 3.0, 5.0])
            context["manual_operational_disposal_table_rus"] = factory.build_generic_table(
                ["Наименование детали", "Причина утилизации", "Вид переработки"],
                disposal.get("operational_flat_rus", []), [6.0, 6.0, 6.0])
            context["manual_operational_disposal_table_eng"] = factory.build_generic_table(
                ["Part name", "Reason for disposal", "Recycling type"], disposal.get("operational_flat_eng", []),
                [6.0, 6.0, 6.0])

        elif sec_key == "ve":
            ve_data = context.get("ve_data", {})
            context["ve_table"] = factory.build_ve_table(ve_data, language=context.get("metadata", {}).get("language", "rus-eng"))

        elif sec_key == "repair_docs":
            docs_data = context.get("repair_docs_data", {})
            context["repair_docs_table"] = factory.build_repair_docs_table(
                docs_data,
                language=context.get("metadata", {}).get("language", "rus-eng")
            )

        elif sec_key == "zik_repair":
            zik_data = context.get("zik_repair_data", {})
            context["zik_repair_table"] = factory.build_zik_repair_table(
                zik_data,
                language=context.get("metadata", {}).get("language", "rus-eng")
            )

        elif sec_key == "zk_repair":
            zk_data = context.get("zk_repair_data", {})
            context["zk_repair_table"] = factory.build_zk_repair_table(
                zk_data,
                language=context.get("metadata", {}).get("language", "rus-eng")
            )

        elif sec_key == "mk_repair":
            mk_data = context.get("mk_repair_data", {})
            context["mk_repair_table"] = factory.build_mk_repair_table(
                mk_data,
                language=context.get("metadata", {}).get("language", "rus-eng")
            )

        elif sec_key == "repair_organization":
            ro = context.get("repair_organization", {})
            context["repair_equipment_table_rus"] = factory.build_generic_table(
                ["Наименование и назначение", "Обозначение", "Основная характеристика", "Примечание"],
                ro.get("equipment_flat_rus", []), [4.5, 4.0, 5.0, 3.5])
            context["repair_equipment_table_eng"] = factory.build_generic_table(
                ["Name and purpose", "Designation", "Main characteristic", "Note"],
                ro.get("equipment_flat_eng", []), [4.5, 4.0, 5.0, 3.5])

        elif sec_key == "repair_requirements":
            rr = context.get("repair_requirements", {})
            context["repair_clearances_table_rus"] = factory.build_generic_table(
                ["Обозначение сопряжения", "Номера позиций", "Размер по чертежу", "Зазор/натяг предельный, мм"],
                rr.get("clearances_flat_rus", []), [4.5, 3.5, 4.5, 4.5])
            context["repair_clearances_table_eng"] = factory.build_generic_table(
                ["Connection designation", "Item Nos.", "Nominal size", "Limit clearance/interference, mm"],
                rr.get("clearances_flat_eng", []), [4.5, 3.5, 4.5, 4.5])

            context["repair_torques_table_rus"] = factory.build_generic_table(
                ["Обозначение соединения", "Номер позиции", "Момент затяжки, Н·м", "Примечание"],
                rr.get("torques_flat_rus", []), [4.5, 3.5, 4.5, 4.5])
            context["repair_torques_table_eng"] = factory.build_generic_table(
                ["Connection designation", "Item No.", "Tightening torque, N·m", "Note"],
                rr.get("torques_flat_eng", []), [4.5, 3.5, 4.5, 4.5])

        # ==============================================================
        # ГЕНЕРАЦИЯ ТАБЛИЦ ДЛЯ ТЕХНИЧЕСКИХ УСЛОВИЙ (ТУ на ремонт)
        # ==============================================================

        elif sec_key == "tk_introduction":
            docs_data = context.get("tk_introduction", {}).get("documents_list", {})
            context["tk_introduction_documents_list_rus"] = factory.build_generic_table(
                ["Наименование документа", "Обозначение документа", "Номер пункта"],
                docs_data.get("rus", []),
                [7.0, 7.0, 3.0]
            )
            context["tk_introduction_documents_list_eng"] = factory.build_generic_table(
                ["Document name", "Document designation", "Item number"],
                docs_data.get("eng", []),
                [7.0, 7.0, 3.0]
            )

        elif sec_key == "tk_repair_cycle":
            rc = context.get("tk_repair_cycle", {})
            context["tk_cycle_structure_rus"] = factory.build_generic_table(
                ["Годы ремонтного цикла", "Категория ТО/ремонта"],
                rc.get("cycle_structure_data", {}).get("rus", []),
                [8.5, 8.5]
            )
            context["tk_cycle_structure_eng"] = factory.build_generic_table(
                ["Years of repair cycle", "Maintenance/repair category"],
                rc.get("cycle_structure_data", {}).get("eng", []),
                [8.5, 8.5]
            )
            context["tk_scope_of_work_rus"] = factory.build_generic_table(
                ["Наименование оборудования и сборочных единиц. Объем (состав) работ",
                 "Обозначение НД, эксплуатационной, ремонтной документации", "Примечание"],
                rc.get("scope_of_work_data", {}).get("rus", []),
                [7.0, 6.0, 4.0]
            )
            context["tk_scope_of_work_eng"] = factory.build_generic_table(
                ["Name of equipment and assembly units. Scope (composition) of works",
                 "Designation of standard, operational, repair documentation", "Note"],
                rc.get("scope_of_work_data", {}).get("eng", []),
                [7.0, 6.0, 4.0]
            )

        elif sec_key == "tk_appendix_a":
            reqs = context.get("tk_technical_requirements", {})
            context["tk_defectation_card_rus"] = factory.build_generic_table(
                ["Обозначение", "Возможный дефект", "Метод установления дефекта", "Средство измерения",
                 "Заключение и рекомендуемые методы ремонта", "Требования после ремонта"],
                reqs.get("defectation_card_data", {}).get("rus", []),
                [2.5, 3.5, 3.0, 3.0, 3.0, 2.0]
            )
            context["tk_defectation_card_eng"] = factory.build_generic_table(
                ["Designation", "Possible defect", "Defect detection method", "Measuring instrument",
                 "Conclusion and recommended repair methods", "Requirements after repair"],
                reqs.get("defectation_card_data", {}).get("eng", []),
                [2.5, 3.5, 3.0, 3.0, 3.0, 2.0]
            )

        elif sec_key == "tk_appendix_b":
            reqs = context.get("tk_technical_requirements", {})
            context["tk_replaced_parts_rus"] = factory.build_generic_table(
                ["Обозначение иллюстрации", "Номер позиции", "Наименование", "Обозначение детали", "Кол-во на ед."],
                reqs.get("replaced_parts_data", {}).get("rus", []),
                [3.0, 2.5, 5.0, 4.5, 2.0]
            )
            context["tk_replaced_parts_eng"] = factory.build_generic_table(
                ["Illustration designation", "Item No.", "Part name", "Part designation", "Qty per unit"],
                reqs.get("replaced_parts_data", {}).get("eng", []),
                [3.0, 2.5, 5.0, 4.5, 2.0]
            )

        elif sec_key == "tk_appendix_v":
            reqs = context.get("tk_technical_requirements", {})
            context["tk_clearances_rus"] = factory.build_generic_table(
                ["Обозначение иллюстрации", "Позиции сопрягаемых деталей", "Обозначение сопряжения",
                 "Номинальный размер", "Размеры с допусками", "Зазор/натяг по КД", "Зазор/натяг после ремонта"],
                reqs.get("clearances_data", {}).get("rus", []),
                [2.5, 2.5, 3.0, 2.0, 2.5, 2.5, 2.0]
            )
            context["tk_clearances_eng"] = factory.build_generic_table(
                ["Illustration designation", "Positions of mating parts", "Mating designation", "Nominal size",
                 "Dimensions with tolerances", "Clearance/interference per design",
                 "Clearance/interference after repair"],
                reqs.get("clearances_data", {}).get("eng", []),
                [2.5, 2.5, 3.0, 2.0, 2.5, 2.5, 2.0]
            )

        elif sec_key == "tk_appendix_g":
            reqs = context.get("tk_technical_requirements", {})
            context["tk_torques_rus"] = factory.build_generic_table(
                ["Обозначение иллюстрации", "Позиции деталей фланцевых соединений", "Позиции крепежных деталей",
                 "Момент затяжки, Н·м", "Примечание"],
                reqs.get("torques_data", {}).get("rus", []),
                [3.0, 4.5, 3.5, 3.0, 3.0]
            )
            context["tk_torques_eng"] = factory.build_generic_table(
                ["Illustration designation", "Positions of flange connection parts", "Positions of fasteners",
                 "Tightening torque, N·m", "Note"],
                reqs.get("torques_data", {}).get("eng", []),
                [3.0, 4.5, 3.5, 3.0, 3.0]
            )

        elif sec_key == "tk_appendix_d":
            reqs = context.get("tk_technical_requirements", {})
            context["tk_materials_rus"] = factory.build_generic_table(
                ["Обозначение иллюстрации", "Номер позиции", "Наименование и обозначение", "Материал по КД",
                 "Материал заменитель"],
                reqs.get("materials_data", {}).get("rus", []),
                [2.5, 2.5, 4.0, 4.0, 4.0]
            )
            context["tk_materials_eng"] = factory.build_generic_table(
                ["Illustration designation", "Item No.", "Name and designation", "Material per design",
                 "Substitute material"],
                reqs.get("materials_data", {}).get("eng", []),
                [2.5, 2.5, 4.0, 4.0, 4.0]
            )

        elif sec_key == "tk_appendix_e":
            reqs = context.get("tk_technical_requirements", {})
            context["tk_measuring_instruments_rus"] = factory.build_generic_table(
                ["Наименование, обозначение, НД (ТУ)", "Номера пунктов"],
                reqs.get("measuring_instruments_data", {}).get("rus", []),
                [12.0, 5.0]
            )
            context["tk_measuring_instruments_eng"] = factory.build_generic_table(
                ["Name, designation, standard", "Item numbers"],
                reqs.get("measuring_instruments_data", {}).get("eng", []),
                [12.0, 5.0]
            )

        # ==============================================================
        # ГЕНЕРАЦИЯ ТАБЛИЦ ДЛЯ РЕГЛАМЕНТА ТОиР
        # ==============================================================

        elif sec_key == "toir_vd":
            vd = context.get("toir_vd", {})
            context["toir_vd_table_rus"] = factory.build_generic_table(
                ["Обозначение документа", "Наименование документа", "Объем документа, листов"],
                vd.get("vd_data", {}).get("rus", []),
                [6.0, 8.0, 3.0]
            )
            context["toir_vd_table_eng"] = factory.build_generic_table(
                ["Document designation", "Document name", "Volume, sheets"],
                vd.get("vd_data", {}).get("eng", []),
                [6.0, 8.0, 3.0]
            )

        elif sec_key == "toir_po":
            po = context.get("toir_po", {})
            context["toir_po_table_rus"] = factory.build_generic_table(
                ["Наименование оборудования", "Обозначение оборудования по КД", "Количество единиц в системе",
                 "Технологическое обозначение"],
                po.get("po_data", {}).get("rus", []),
                [6.0, 5.0, 3.0, 3.0]
            )
            context["toir_po_table_eng"] = factory.build_generic_table(
                ["Equipment name", "Equipment designation by DD", "Number of units in system",
                 "Technological designation"],
                po.get("po_data", {}).get("eng", []),
                [6.0, 5.0, 3.0, 3.0]
            )

        elif sec_key == "toir_ksc":
            ksc = context.get("toir_ksc", {})
            context["toir_ksc_table_rus"] = factory.build_generic_table(
                ["Годы ремонтного цикла", "Категории ТО/ремонта", "Эксплуатационное состояние энергоблока(ов)"],
                ksc.get("ksc_data", {}).get("rus", []),
                [5.0, 6.0, 6.0]
            )
            context["toir_ksc_table_eng"] = factory.build_generic_table(
                ["Years of repair cycle", "Categories of maintenance/repair", "Operational state of the power unit(s)"],
                ksc.get("ksc_data", {}).get("eng", []),
                [5.0, 6.0, 6.0]
            )

        elif sec_key == "toir_ko":
            ko = context.get("toir_ko", {})
            context["toir_ko_table_rus"] = factory.build_generic_table(
                ["Наименование оборудования и сборочных единиц. Объем (состав) работ",
                 "Обозначение НД, эксплуатационной, ремонтной документации", "Примечания"],
                ko.get("ko_data", {}).get("rus", []),
                [7.0, 6.0, 4.0]
            )
            context["toir_ko_table_eng"] = factory.build_generic_table(
                ["Name of equipment and assembly units. Scope (composition) of works",
                 "Designation of standard, operational, repair documentation", "Notes"],
                ko.get("ko_data", {}).get("eng", []),
                [7.0, 6.0, 4.0]
            )

        elif sec_key == "toir_vr":
            vr = context.get("toir_vr", {})
            context["toir_vr_table_rus"] = factory.build_generic_table(
                ["Наименование оборудования и сборочных единиц. Объем (состав) работ",
                 "Обозначение НД, эксплуатационной, ремонтной документации", "Примечания"],
                vr.get("vr_data", {}).get("rus", []),
                [7.0, 6.0, 4.0]
            )
            context["toir_vr_table_eng"] = factory.build_generic_table(
                ["Name of equipment and assembly units. Scope (composition) of works",
                 "Designation of standard, operational, repair documentation", "Notes"],
                vr.get("vr_data", {}).get("eng", []),
                [7.0, 6.0, 4.0]
            )

        # ==============================================================
        # ГЕНЕРАЦИЯ ТАБЛИЦ ДЛЯ КОМПЛЕКТА ТИПОВОЙ ТЕХНОЛОГИЧЕСКОЙ ДОКУМЕНТАЦИИ (КТТД)
        # ==============================================================

        elif sec_key == "kttd_vtd":
            vtd = context.get("kttd_vtd", {})
            context["kttd_vtd_table_rus"] = factory.build_generic_table(
                ["Обозначение документа", "Наименование документа", "Объем документа, листов"],
                vtd.get("vtd_data", {}).get("rus", []),
                [6.0, 8.0, 3.0]
            )
            context["kttd_vtd_table_eng"] = factory.build_generic_table(
                ["Document designation", "Document name", "Volume, sheets"],
                vtd.get("vtd_data", {}).get("eng", []),
                [6.0, 8.0, 3.0]
            )

        elif sec_key == "kttd_kttp":
            kttp = context.get("kttd_kttp", {})
            context["kttd_kttp_table_rus"] = factory.build_generic_table(
                ["Цех, Номер, наименование операции", "Описание переходов", "Материалы и оснастка",
                 "Трудозатраты (Проф, Р, КР)"],
                kttp.get("kttp_data", {}).get("rus", []),
                [4.0, 7.0, 4.0, 2.0]
            )
            context["kttd_kttp_table_eng"] = factory.build_generic_table(
                ["Shop, No., operation name", "Description of transitions", "Materials and tooling",
                 "Labor costs (Prof, G, Qty)"],
                kttp.get("kttp_data", {}).get("eng", []),
                [4.0, 7.0, 4.0, 2.0]
            )

        elif sec_key == "kttd_vd":
            vd = context.get("kttd_vd", {})
            context["kttd_vd_table_rus"] = factory.build_generic_table(
                ["Обозначение СЕ/Д", "Наименование дефекта", "Допустимое значение", "Средство измерения",
                 "Особые указания"],
                vd.get("vd_data", {}).get("rus", []),
                [3.0, 4.0, 3.0, 3.0, 4.0]
            )
            context["kttd_vd_table_eng"] = factory.build_generic_table(
                ["Designation", "Defect name", "Permissible value", "Measuring instrument", "Special instructions"],
                vd.get("vd_data", {}).get("eng", []),
                [3.0, 4.0, 3.0, 3.0, 4.0]
            )

        elif sec_key == "kttd_vok":
            vok = context.get("kttd_vok", {})
            context["kttd_vok_table_rus"] = factory.build_generic_table(
                ["Номер, наименование операции", "Обозначение документа", "Объем контроля",
                 "Наименование, марка средства", "Кол-во"],
                vok.get("vok_data", {}).get("rus", []),
                [4.0, 4.0, 3.0, 4.0, 2.0]
            )
            context["kttd_vok_table_eng"] = factory.build_generic_table(
                ["No., operation name", "Document designation", "Control volume", "Name, brand of instrument", "Qty"],
                vok.get("vok_data", {}).get("eng", []),
                [4.0, 4.0, 3.0, 4.0, 2.0]
            )

        elif sec_key == "kttd_kk":
            kk = context.get("kttd_kk", {})
            context["kttd_kk_table_rus"] = factory.build_generic_table(
                ["Наименование детали/СЕ", "ЕН", "ЕВ", "Кол-во"],
                kk.get("kk_data", {}).get("rus", []),
                [8.0, 3.0, 3.0, 3.0]
            )
            context["kttd_kk_table_eng"] = factory.build_generic_table(
                ["Name of part/assembly unit", "NU", "UM", "Qty"],
                kk.get("kk_data", {}).get("eng", []),
                [8.0, 3.0, 3.0, 3.0]
            )

        elif sec_key == "kttd_vo":
            vo = context.get("kttd_vo", {})
            context["kttd_vo_table_rus"] = factory.build_generic_table(
                ["Наименование, марка средства оснащения", "ЕВ", "Кол-во"],
                vo.get("vo_data", {}).get("rus", []),
                [10.0, 3.0, 4.0]
            )
            context["kttd_vo_table_eng"] = factory.build_generic_table(
                ["Name, brand of tooling", "UM", "Qty"],
                vo.get("vo_data", {}).get("eng", []),
                [10.0, 3.0, 4.0]
            )

        elif sec_key == "kttd_vm":
            vm = context.get("kttd_vm", {})
            context["kttd_vm_table_rus"] = factory.build_generic_table(
                ["Наименование, марка материала", "ЕВ", "Кол-во"],
                vm.get("vm_data", {}).get("rus", []),
                [10.0, 3.0, 4.0]
            )
            context["kttd_vm_table_eng"] = factory.build_generic_table(
                ["Name, brand of material", "UM", "Qty"],
                vm.get("vm_data", {}).get("eng", []),
                [10.0, 3.0, 4.0]
            )

    def _inject_signatures(self, context: dict, tpl: DocxTemplate):
        """
        Умная загрузка подписей. Ищет .png файлы по английским фамилиям.
        Ширина 3.0 см оптимальна для ГОСТовских штампов.
        """
        company = context.get("company", {})
        signatures_dir = self.templates_dir / "ms_word" / "signatures"

        # Сопоставляем ключи контекста с полями фамилий из JSON
        roles = {
            "designer_sign": company.get("designer_eng"),
            "inspector_sign": company.get("inspector_eng"),
            "approver_sign": company.get("approver_eng")
        }

        for sign_key, name in roles.items():
            context[sign_key] = ""  # Если нет файла или фамилии - просто пустота
            if name:
                clean_name = str(name).strip()
                img_path = signatures_dir / f"{clean_name}.png"
                if img_path.exists():
                    try:
                        # Вставляем картинку с жестко заданной шириной (защита от разрыва ячеек)
                        context[sign_key] = InlineImage(tpl, str(img_path), width=Cm(1.0))
                    except Exception as e:
                        logger.warning(f"[ModularBuilder] Ошибка вставки подписи {clean_name}: {e}")
                else:
                    logger.debug(f"[ModularBuilder] Подпись не найдена: {img_path}")