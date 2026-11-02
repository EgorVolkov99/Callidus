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
        doc_type = metadata.get("type", "passport")

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
            "toc", "basic_info", "technical_data", "main_parts", "approval_tests",
            "scope_of_delivery", "warranties", "preservation", "packing",
            "deviations", "conclusion", "disposal", "terms_list", "documents_list", "revision_sheet",
            "introduction", "description_and_operation", "usage", "maintenance",
            "current_repair", "storage", "transportation", "ve"
        ]

        logger.info(f"[ModularBuilder] Инициализация скелета ({master_path.name})...")

        # ==============================================================
        # ЭТАП 1: РЕНДЕРИНГ ОТДЕЛЬНЫХ РАЗДЕЛОВ (SECTIONS) В ПАМЯТИ
        # ==============================================================
        body_tpl = DocxTemplate(body_path)
        rendered_sections = {}

        for sec_key in all_sections:
            if context.get(sec_key):

                # 1. Пытаемся найти специфичный для шаблона файл (например, nikimt_...)
                sec_file = sections_dir / doc_type / f"{master_name}_{sec_key}_{language_option}.docx"
                if not sec_file.exists():
                    sec_file = sections_dir / "common" / f"{master_name}_{sec_key}_{language_option}.docx"

                # 2. ФОЛБЭК: Если кастомного файла нет, берем базовый (base_...)
                if not sec_file.exists():
                    logger.debug(f"Кастомный {master_name}_{sec_key} не найден. Используем base_...")
                    sec_file = sections_dir / doc_type / f"base_{sec_key}_{language_option}.docx"
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

        # Принудительное обновление оглавления и полей (NumPages)
        settings = master_tpl.docx.settings._element
        update_fields = settings.find(qn('w:updateFields'))
        if update_fields is None:
            update_fields = OxmlElement('w:updateFields')
            settings.append(update_fields)
        update_fields.set(qn('w:val'), 'true')

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