from typing import Dict, Any


def build_word_rows(raw_dict: dict, is_eng=False) -> list:
    """Сборка плоского списка для ПСИ."""
    tests = raw_dict.get("Type of test" if is_eng else "Вид испытаний", [])
    if not tests:
        return []
    mediums = raw_dict.get("Test medium" if is_eng else "Среда испытательная") or [None] * len(tests)
    pressures = raw_dict.get("Test pressure, MPa" if is_eng else "Давление испытаний, МПа") or [None] * len(tests)
    temperatures = raw_dict.get("Test temperature, not less than, °C" if is_eng else "Температура испытаний, °C") or [
        None] * len(tests)
    results = raw_dict.get(
        "Test results\naccording to documentation\nactual:" if is_eng else "Результаты испытаний\nпо документации\nфактические:",
        []) or [None] * len(tests)

    word_rows = []
    for i in range(len(tests)):
        word_rows.append({
            "test_type": tests[i],
            "medium": mediums[i] if i < len(mediums) else None,
            "pressure": pressures[i] if i < len(pressures) else None,
            "temperature": temperatures[i] if i < len(temperatures) else None,
            "result": results[i] if i < len(results) else None,
        })
    return word_rows


def build_flat_rows(data_dict: dict) -> list:
    """Конвертация dict массивов в списки списков [[col1, col2], ...]."""
    if not data_dict or not isinstance(data_dict, dict):
        return []
    keys = list(data_dict.keys())
    if not keys or not isinstance(data_dict[keys[0]], list):
        return []
    num_rows = len(data_dict[keys[0]])
    flat_rows = []
    for i in range(num_rows):
        row_vals = [data_dict[k][i] if i < len(data_dict[k]) else "" for k in keys]
        flat_rows.append(row_vals)
    return flat_rows


def process_materials_table(data_dict: dict) -> dict:
    """
    Умный парсер для новой структуры таблиц.
    Автоматически определяет базовые колонки и колонки испытаний.
    """
    if not data_dict or not isinstance(data_dict, dict):
        return {"base_cols_count": 0, "headers": [], "rows": []}

    # 1. Определяем базовые колонки (добавлен "Operation names" для английской версии)
    ignore_keys = ["Наименования операций", "Type of test", "Operation names"]
    base_keys = [k for k in data_dict.keys() if k not in ignore_keys]

    # 2. Извлекаем динамические заголовки испытаний (вертикальные колонки)
    dynamic_headers = (data_dict.get("Наименования операций") or
                       data_dict.get("Operation names") or
                       data_dict.get("Type of test") or [])
    if not isinstance(dynamic_headers, list):
        dynamic_headers = []

    headers = base_keys + dynamic_headers
    base_cols_count = len(base_keys)

    # 3. Определяем количество строк по самой длинной базовой колонке
    num_rows = 0
    for k in base_keys:
        val = data_dict[k]
        if isinstance(val, list):
            num_rows = max(num_rows, len(val))

    # 4. Собираем плоские строки
    rows = []
    for i in range(num_rows):
        row = []
        # Заполняем базовые колонки
        for k in base_keys:
            val_list = data_dict.get(k)
            if isinstance(val_list, list) and i < len(val_list) and val_list[i] is not None:
                row.append(str(val_list[i]))
            else:
                row.append("")  # Пустота, если пришел null

        # Заполняем колонки испытаний
        for dyn_k in dynamic_headers:
            val_list = data_dict.get(dyn_k)
            if isinstance(val_list, list) and i < len(val_list) and val_list[i] is not None:
                row.append(str(val_list[i]))
            else:
                row.append("")  # Оставляем ячейку пустой под заполнение

        rows.append(row)

    return {
        "base_cols_count": base_cols_count,
        "headers": headers,
        "rows": rows
    }


def prepare_passport_context(raw_data: Dict[str, Any]) -> Dict[str, Any]:
    """Главный адаптер Паспорта: готовит данные, считает разделы, формирует таблицы."""
    context = raw_data.copy()

    # 1. Жесткая защита KKS-кода (в Паспорт идет только строка)
    unit = context.get("unit", {})
    kks = unit.get("kks_code", "")
    if isinstance(kks, list):
        unit["kks_code"] = kks[0] if kks else ""
    context["unit"] = unit

    # Обратная совместимость для старых шаблонов, если где-то остался variables
    context["variables"] = context.get("company", {})

    # 2. Нумерация разделов и фильтрация таблиц
    sec_counter = 1
    sections_config = [
        "basic_info", "technical_data", "main_parts", "approval_tests",
        "scope_of_delivery", "warranties", "preservation", "packing",
        "deviations", "conclusion", "disposal"
    ]

    for sec_key in sections_config:
        sec_data = context.get(sec_key)

        if sec_data is None or sec_data is False:
            context[sec_key] = False
            continue

        if sec_data is True:
            sec_data = {}

        if sec_key == "basic_info":
            sec_data["rus"] = [i for i in sec_data.get("rus", []) if i.get("value")]
            sec_data["eng"] = [i for i in sec_data.get("eng", []) if i.get("value")]


        if sec_key == "main_parts":
            materials = sec_data.get("materials", {})
            welds = sec_data.get("welds", {})
            valves = sec_data.get("valves", {})

            materials = materials if isinstance(materials, dict) else {}
            welds = welds if isinstance(welds, dict) else {}
            valves = valves if isinstance(valves, dict) else {}

            sec_data["materials_rus"] = process_materials_table(materials.get("rus", {}))
            sec_data["materials_eng"] = process_materials_table(materials.get("eng", {}))

            sec_data["welds_rus"] = process_materials_table(welds.get("rus", {}))
            sec_data["welds_eng"] = process_materials_table(welds.get("eng", {}))

            sec_data["valves_rus"] = process_materials_table(valves.get("rus", {}))
            sec_data["valves_eng"] = process_materials_table(valves.get("eng", {}))


        if sec_key == "approval_tests":
            rus_raw = sec_data.get("rus", {})
            eng_raw = sec_data.get("eng", {})

            def check_extra(d):
                keys = ["Среда испытательная", "Давление испытаний, МПа", "Температура испытаний, °C", "Test medium",
                        "Test pressure, MPa", "Test temperature, not less than, °C"]
                for k in keys:
                    if d.get(k) and any(v is not None for v in d[k]): return True
                return False

            sec_data["has_extra_columns"] = check_extra(rus_raw) or check_extra(eng_raw)
            sec_data["rus_word_rows"] = build_word_rows(rus_raw, is_eng=False)
            sec_data["eng_word_rows"] = build_word_rows(eng_raw, is_eng=True)

        if sec_key == "scope_of_delivery":
            units = sec_data.get("units_set", {})
            docs = sec_data.get("document_set", {})
            sec_data["units_word_rus"] = build_flat_rows(units.get("rus", {}))
            sec_data["units_word_eng"] = build_flat_rows(units.get("eng", {}))
            sec_data["docs_word_rus"] = build_flat_rows(docs.get("rus", {}))
            sec_data["docs_word_eng"] = build_flat_rows(docs.get("eng", {}))

        # Присваиваем номер раздела (ТЕПЕРЬ sec_num БУДЕТ РАБОТАТЬ!)
        sec_data["sec_num"] = sec_counter
        context[sec_key] = sec_data
        sec_counter += 1

    return context