from typing import Dict, Any


def dict_list_to_flat_rows(data_list: list, keys: list) -> list:
    """
    Вытаскивает значения из словарей по ключам и превращает в плоский список.
    """
    if not data_list or not isinstance(data_list, list):
        return []

    flat_rows = []
    for item in data_list:
        if isinstance(item, dict):
            row = [str(item.get(k) or "") for k in keys]
            flat_rows.append(row)
        else:
            # Для простых списков (например, строк)
            flat_rows.append([str(item)])
    return flat_rows


def prepare_manual_context(raw_data: Dict[str, Any]) -> Dict[str, Any]:
    """Главный адаптер Руководства по эксплуатации (РЭ)."""
    context = raw_data.copy()

    # 1. Защита KKS-кода
    unit = context.get("unit", {})
    kks = unit.get("kks_code", [])
    if isinstance(kks, str):
        unit["kks_code"] = [kks] if kks else []
    context["unit"] = unit

    context["variables"] = context.get("company", {})

    # 2. Динамическая нумерация и подготовка таблиц
    sec_counter = 1
    sections_config = [
        "description_and_operation", "usage", "maintenance",
        "current_repair", "storage", "transportation", "disposal"
    ]

    for sec_key in sections_config:
        sec_data = context.get(sec_key)

        if sec_data is None or sec_data is False:
            context[sec_key] = False
            continue

        if sec_data is True:
            sec_data = {}

        # --- РАЗДЕЛ 2: Использование по назначению ---
        if sec_key == "usage":
            # Ограничения
            lim = sec_data.get("limitations", {})
            if isinstance(lim, dict):
                sec_data["limitations_flat_rus"] = dict_list_to_flat_rows(
                    lim.get("rus", {}).get("limitations_list", []), ["parameter", "value", "consequence"])
                sec_data["limitations_flat_eng"] = dict_list_to_flat_rows(
                    lim.get("eng", {}).get("limitations_list", []), ["parameter", "value", "consequence"])

            # Неисправности (Подготовка)
            prep = sec_data.get("preparation", {})
            if isinstance(prep, dict):
                sec_data["prep_troubleshoot_flat_rus"] = dict_list_to_flat_rows(
                    prep.get("rus", {}).get("troubleshooting", []), ["problem", "action"])
                sec_data["prep_troubleshoot_flat_eng"] = dict_list_to_flat_rows(
                    prep.get("eng", {}).get("troubleshooting", []), ["problem", "action"])

            # Неисправности (Работа)
            op = sec_data.get("operation", {})
            if isinstance(op, dict):
                sec_data["op_troubleshoot_flat_rus"] = dict_list_to_flat_rows(
                    op.get("rus", {}).get("troubleshooting", []), ["problem", "action"])
                sec_data["op_troubleshoot_flat_eng"] = dict_list_to_flat_rows(
                    op.get("eng", {}).get("troubleshooting", []), ["problem", "action"])

        # --- РАЗДЕЛ 3: Техническое обслуживание ---
        if sec_key == "maintenance":
            unit_maint = sec_data.get("unit", {})
            if isinstance(unit_maint, dict):
                # Порядок ТО
                proc = unit_maint.get("maintenance_procedure", {})
                sec_data["procedure_flat_rus"] = dict_list_to_flat_rows(proc.get("rus", {}).get("procedure_table", []),
                                                                        ["item_and_work", "maintenance_type", "note"])
                sec_data["procedure_flat_eng"] = dict_list_to_flat_rows(proc.get("eng", {}).get("procedure_table", []),
                                                                        ["item_and_work", "maintenance_type", "note"])

                # Проверка работоспособности
                chk = unit_maint.get("operability_check", {})
                sec_data["check_flat_rus"] = dict_list_to_flat_rows(chk.get("rus", {}).get("check_table", []),
                                                                    ["work_name", "executor", "tools",
                                                                     "control_values"])
                sec_data["check_flat_eng"] = dict_list_to_flat_rows(chk.get("eng", {}).get("check_table", []),
                                                                    ["work_name", "executor", "tools",
                                                                     "control_values"])

        # --- РАЗДЕЛ 4: Текущий ремонт ---
        if sec_key == "current_repair":
            parts = sec_data.get("parts", [])
            # Берем таблицу ремонта из первой составной части для вывода в общую таблицу
            if isinstance(parts, list) and len(parts) > 0:
                first_part = parts[0]
                sec_data["repair_flat_rus"] = dict_list_to_flat_rows(first_part.get("rus", {}).get("repair_table", []),
                                                                     ["failure_description", "consequences",
                                                                      "possible_causes", "detection_methods",
                                                                      "elimination_methods"])
                sec_data["repair_flat_eng"] = dict_list_to_flat_rows(first_part.get("eng", {}).get("repair_table", []),
                                                                     ["failure_description", "consequences",
                                                                      "possible_causes", "detection_methods",
                                                                      "elimination_methods"])

        # --- РАЗДЕЛ 5: Хранение ---
        if sec_key == "storage":
            sec_data["shelf_life_flat_rus"] = dict_list_to_flat_rows(
                sec_data.get("rus", {}).get("limited_shelf_life_parts", []), ["part_name", "shelf_life", "note"])
            sec_data["shelf_life_flat_eng"] = dict_list_to_flat_rows(
                sec_data.get("eng", {}).get("limited_shelf_life_parts", []), ["part_name", "shelf_life", "note"])

        # --- РАЗДЕЛ 7: Утилизация ---
        if sec_key == "disposal":
            sec_data["calculated_flat_rus"] = dict_list_to_flat_rows(
                sec_data.get("rus", {}).get("calculated_disposed_parts", []),
                ["part_name", "material", "weight_kg", "recycling_type"])
            sec_data["calculated_flat_eng"] = dict_list_to_flat_rows(
                sec_data.get("eng", {}).get("calculated_disposed_parts", []),
                ["part_name", "material", "weight_kg", "recycling_type"])

            sec_data["operational_flat_rus"] = dict_list_to_flat_rows(
                sec_data.get("rus", {}).get("operational_disposed_parts", []),
                ["part_name", "reason", "recycling_type"])
            sec_data["operational_flat_eng"] = dict_list_to_flat_rows(
                sec_data.get("eng", {}).get("operational_disposed_parts", []),
                ["part_name", "reason", "recycling_type"])

        sec_data["sec_num"] = sec_counter
        context[sec_key] = sec_data
        sec_counter += 1

    return context