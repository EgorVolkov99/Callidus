from pathlib import Path
from docxtpl import DocxTemplate
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from engine.builders.table_factory import TableFactory


class DocxBuilder:
    def __init__(self, template_path: Path):
        self.tpl = DocxTemplate(template_path)
        self.factory = TableFactory(self.tpl)

    def render_and_save(self, context: dict, output_path: Path):
        doc_type = context.get("metadata", {}).get("type", "passport")

        # На всякий случай задаем пустые значения для всех плейсхолдеров,
        # чтобы Jinja2 не крашился, если раздел выключен
        self._set_default_placeholders(context)

        # ==========================================
        # УНИКАЛЬНЫЕ ТАБЛИЦЫ (В ОСНОВНОМ ПАСПОРТ)
        # ==========================================
        if context.get("basic_info"):
            context["basic_info_table_rus"] = self.factory.build_basic_info_table(context["basic_info"].get("rus", []))
            context["basic_info_table_eng"] = self.factory.build_basic_info_table(context["basic_info"].get("eng", []))

        if context.get("main_parts"):
            mp = context["main_parts"]
            context["materials_table_rus"] = self.factory.build_materials_table(mp.get("materials_rus", {}),
                                                                                is_eng=False)
            context["materials_table_eng"] = self.factory.build_materials_table(mp.get("materials_eng", {}),
                                                                                is_eng=True)
            context["welds_table_rus"] = self.factory.build_welds_table(mp.get("welds_rus", {}), is_eng=False)
            context["welds_table_eng"] = self.factory.build_welds_table(mp.get("welds_eng", {}), is_eng=True)
            context["valves_table_rus"] = self.factory.build_valves_table(mp.get("valves_rus", {}), is_eng=False)
            context["valves_table_eng"] = self.factory.build_valves_table(mp.get("valves_eng", {}), is_eng=True)

        if context.get("approval_tests"):
            scope = context["approval_tests"]
            context["approval_tests_table_rus"] = self.factory.build_approval_tests_table(
                scope.get("rus_word_rows", []), context.get("unit", {}), scope.get("has_extra_columns", False),
                is_eng=False)
            context["approval_tests_table_eng"] = self.factory.build_approval_tests_table(
                scope.get("eng_word_rows", []), context.get("unit", {}), scope.get("has_extra_columns", False),
                is_eng=True)

        # ==========================================
        # БАЗОВЫЕ И ГРУППИРОВАННЫЕ ТАБЛИЦЫ
        # ==========================================

        # --- Технические характеристики (сгруппированные) ---
        if doc_type == "passport" and context.get("technical_data"):
            context["technical_data_table_rus"] = self.factory.build_grouped_table(
                context["technical_data"].get("rus", {}), ["Наименование параметра", "Значение параметра"])
            context["technical_data_table_eng"] = self.factory.build_grouped_table(
                context["technical_data"].get("eng", {}), ["Parameter name", "Parameter value"])
        elif doc_type == "manual":
            desc_op = context.get("description_and_operation", {})
            if isinstance(desc_op, dict) and desc_op.get("unit", {}).get("specifications"):
                specs = desc_op["unit"]["specifications"]
                context["technical_data_table_rus"] = self.factory.build_grouped_table(
                    specs.get("rus", {}), ["Наименование параметра", "Значение параметра"])
                context["technical_data_table_eng"] = self.factory.build_grouped_table(
                    specs.get("eng", {}), ["Parameter name", "Parameter value"])

        # --- Комплектность (плоские) ---
        if context.get("scope_of_delivery"):
            scope = context["scope_of_delivery"]
            context["delivery_set_table_rus"] = self.factory.build_generic_table(
                ["Обозначение", "Наименование", "Серийный номер", "Код KKS", "Кол-во"], scope.get("units_word_rus", []))
            context["delivery_set_table_eng"] = self.factory.build_generic_table(
                ["Designation", "Description", "Serial No", "KKS code", "Qty"], scope.get("units_word_eng", []))
            context["documentation_set_table_rus"] = self.factory.build_generic_table(
                ["Обозначение", "Наименование", "Кол-во", "Примечание"], scope.get("docs_word_rus", []))
            context["documentation_set_table_eng"] = self.factory.build_generic_table(
                ["Designation", "Description", "Qty", "Note"], scope.get("docs_word_eng", []))

        # ==========================================
        # ТАБЛИЦЫ ТОЛЬКО ДЛЯ РУКОВОДСТВА (MANUAL)
        # ==========================================
        if doc_type == "manual":
            # 1. Раздел: Использование по назначению
            usage = context.get("usage", {})
            if isinstance(usage, dict):
                # Ограничения
                context["manual_limitations_table_rus"] = self.factory.build_generic_table(
                    ["Параметр", "Предельное значение", "Последствия несоблюдения"],
                    usage.get("limitations_flat_rus", []))
                context["manual_limitations_table_eng"] = self.factory.build_generic_table(
                    ["Parameter", "Limit value", "Consequences"], usage.get("limitations_flat_eng", []))

                # Возможные неисправности (Подготовка)
                context["manual_prep_troubleshooting_table_rus"] = self.factory.build_generic_table(
                    ["Описание неисправности", "Действия по устранению"], usage.get("prep_troubleshoot_flat_rus", []))
                context["manual_prep_troubleshooting_table_eng"] = self.factory.build_generic_table(
                    ["Malfunction", "Actions to eliminate"], usage.get("prep_troubleshoot_flat_eng", []))

                # Возможные неисправности (Использование)
                context["manual_op_troubleshooting_table_rus"] = self.factory.build_generic_table(
                    ["Описание неисправности", "Действия по устранению"], usage.get("op_troubleshoot_flat_rus", []))
                context["manual_op_troubleshooting_table_eng"] = self.factory.build_generic_table(
                    ["Malfunction", "Actions to eliminate"], usage.get("op_troubleshoot_flat_eng", []))

            # 2. Раздел: Техническое обслуживание
            maint = context.get("maintenance", {})
            if isinstance(maint, dict):
                # ГСМ
                context["manual_fuel_table_rus"] = self.factory.build_generic_table(
                    ["Наименование"], maint.get("fuel_flat_rus", []))
                context["manual_fuel_table_eng"] = self.factory.build_generic_table(
                    ["Description"], maint.get("fuel_flat_eng", []))

                # Порядок ТО
                context["manual_maint_procedure_table_rus"] = self.factory.build_generic_table(
                    ["Наименование объекта ТО и работы", "Виды ТО", "Примечание"], maint.get("procedure_flat_rus", []))
                context["manual_maint_procedure_table_eng"] = self.factory.build_generic_table(
                    ["Name of the maintenance object and work", "Maintenance types", "Note"],
                    maint.get("procedure_flat_eng", []))

                # Проверка работоспособности
                context["manual_operability_check_table_rus"] = self.factory.build_generic_table(
                    ["Наименование работы", "Кто выполняет", "Средства измерений и материалы",
                     "Контрольные значения параметров"], maint.get("check_flat_rus", []))
                context["manual_operability_check_table_eng"] = self.factory.build_generic_table(
                    ["Work name", "Executor", "Measuring instruments and materials", "Control values"],
                    maint.get("check_flat_eng", []))

            # 3. Раздел: Текущий ремонт (Таблицы формируются на уровне составных частей)
            # Внимание: для частей логика чуть сложнее, мы прокидываем пока пустые заглушки.
            # Если требуется единая таблица ремонта на все изделие:
            current_repair = context.get("current_repair", {})
            if isinstance(current_repair, dict):
                context["manual_current_repair_table_rus"] = self.factory.build_generic_table(
                    ["Описание отказа", "Последствия", "Возможные причины", "Способ обнаружения", "Способ устранения"],
                    current_repair.get("repair_flat_rus", []))
                context["manual_current_repair_table_eng"] = self.factory.build_generic_table(
                    ["Failure description", "Consequences", "Possible causes", "Detection method",
                     "Elimination method"], current_repair.get("repair_flat_eng", []))

            # 4. Раздел: Хранение
            storage = context.get("storage", {})
            if isinstance(storage, dict):
                context["manual_limited_shelf_life_table_rus"] = self.factory.build_generic_table(
                    ["Наименование детали", "Срок годности", "Примечание"], storage.get("shelf_life_flat_rus", []))
                context["manual_limited_shelf_life_table_eng"] = self.factory.build_generic_table(
                    ["Part name", "Shelf life", "Note"], storage.get("shelf_life_flat_eng", []))

            # 5. Раздел: Утилизация
            disposal = context.get("disposal", {})
            if isinstance(disposal, dict):
                context["manual_calculated_disposal_table_rus"] = self.factory.build_generic_table(
                    ["Наименование детали", "Материал", "Масса, кг", "Вид переработки"],
                    disposal.get("calculated_flat_rus", []))
                context["manual_calculated_disposal_table_eng"] = self.factory.build_generic_table(
                    ["Part name", "Material", "Weight, kg", "Recycling type"], disposal.get("calculated_flat_eng", []))

                context["manual_operational_disposal_table_rus"] = self.factory.build_generic_table(
                    ["Наименование детали", "Причина утилизации", "Вид переработки"],
                    disposal.get("operational_flat_rus", []))
                context["manual_operational_disposal_table_eng"] = self.factory.build_generic_table(
                    ["Part name", "Reason for disposal", "Recycling type"], disposal.get("operational_flat_eng", []))

        # ==========================================
        # ФИНАЛЬНЫЙ РЕНДЕР И СОХРАНЕНИЕ
        # ==========================================
        self.tpl.render(context)

        settings = self.tpl.docx.settings._element
        update_fields = settings.find(qn('w:updateFields'))
        if update_fields is None:
            update_fields = OxmlElement('w:updateFields')
            settings.append(update_fields)
        update_fields.set(qn('w:val'), 'true')

        output_path.parent.mkdir(parents=True, exist_ok=True)
        self.tpl.save(output_path)
        return output_path

    def _set_default_placeholders(self, context: dict):
        """Заглушки для всех таблиц, чтобы Jinja2 не падал, если данных нет."""
        keys = [
            "basic_info_table_rus", "basic_info_table_eng",
            "technical_data_table_rus", "technical_data_table_eng",
            "materials_table_rus", "materials_table_eng",
            "welds_table_rus", "welds_table_eng",
            "valves_table_rus", "valves_table_eng",
            "approval_tests_table_rus", "approval_tests_table_eng",
            "delivery_set_table_rus", "delivery_set_table_eng",
            "documentation_set_table_rus", "documentation_set_table_eng",
            "manual_limitations_table_rus", "manual_limitations_table_eng",
            "manual_prep_troubleshooting_table_rus", "manual_prep_troubleshooting_table_eng",
            "manual_op_troubleshooting_table_rus", "manual_op_troubleshooting_table_eng",
            "manual_fuel_table_rus", "manual_fuel_table_eng",
            "manual_maint_procedure_table_rus", "manual_maint_procedure_table_eng",
            "manual_operability_check_table_rus", "manual_operability_check_table_eng",
            "manual_current_repair_table_rus", "manual_current_repair_table_eng",
            "manual_limited_shelf_life_table_rus", "manual_limited_shelf_life_table_eng",
            "manual_calculated_disposal_table_rus", "manual_calculated_disposal_table_eng",
            "manual_operational_disposal_table_rus", "manual_operational_disposal_table_eng"
        ]
        for key in keys:
            context.setdefault(key, "")