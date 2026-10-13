import json
import argparse
from pathlib import Path
from collections import defaultdict


def extract_equipment_data(data: dict) -> tuple:
    """
    Извлекает Наименование, Бренд, Модель и KKS из JSON файла.
    Ищет сначала в 'unit_parameters' (как в вашей модели Pydantic),
    затем делает fallback на 'basic_info' (на всякий случай).
    """
    name, brand, model, kks = "Неизвестно", "Неизвестно", "Неизвестно", "Неизвестно"

    # 1. Попытка извлечь из корневого узла unit_parameters (самый надежный вариант)
    unit = data.get("unit_parameters", {})
    if unit:
        name = unit.get("product_name_rus", name)
        brand = unit.get("product_brand", brand)
        model = unit.get("product_model", model)
        kks = unit.get("kks_code", kks)
        return name, brand, model, kks

    # 2. Резервный поиск внутри таблицы basic_info.rus (если структура другая)
    basic_info_rus = data.get("basic_info", {}).get("rus", [])
    for item in basic_info_rus:
        if not isinstance(item, dict):
            continue

        item_name = str(item.get("name", "")).lower()
        item_value = str(item.get("value", ""))

        if not item_value or item_value == "None":
            continue

        if item_name == "наименование":
            name = item_value
        elif "марка" in item_name or "бренд" in item_name:
            brand = item_value
        elif "модель" in item_name:
            model = item_value
        elif "kks" in item_name:
            kks = item_value

    return name, brand, model, kks


def main():
    parser = argparse.ArgumentParser(description="Сборщик сводной ведомости оборудования из JSON-паспортов")
    parser.add_argument(
        "--folder",
        type=str,
        default="../passport/projects/NIKIMT_KIP/input",
        help="Путь к папке с исходными .json файлами"
    )
    parser.add_argument(
        "--output",
        type=str,
        default="equipment_summary.json",
        help="Имя итогового файла"
    )
    args = parser.parse_args()

    target_dir = Path(args.folder)
    output_file = Path(args.output)

    if not target_dir.exists() or not target_dir.is_dir():
        print(f"❌ Ошибка: Папка не найдена: {target_dir.resolve()}")
        return

    json_files = list(target_dir.glob("*.json"))
    if not json_files:
        print(f"⚠️ В папке '{target_dir}' не найдено ни одного .json файла.")
        return

    print(f"🔍 Сканирование {len(json_files)} файлов в директории: {target_dir.name}...")

    # Структура для группировки: dict[name][brand][model] = [kks1, kks2, ...]
    grouped_data = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))

    # 1. Читаем и группируем данные
    processed_count = 0
    for file_path in json_files:
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            name, brand, model, kks = extract_equipment_data(data)

            # Добавляем KKS в список соответствующей модели
            if kks not in grouped_data[name][brand][model]:
                grouped_data[name][brand][model].append(kks)

            processed_count += 1

        except Exception as e:
            print(f"  ❌ Ошибка чтения {file_path.name}: {e}")

    # 2. Формируем итоговый массив в нужном формате
    final_output = []

    for name, brands in grouped_data.items():
        for brand, models in brands.items():
            model_list = list(models.keys())
            # Получаем параллельный массив списков KKS кодов
            kks_list = [models[m] for m in model_list]

            final_output.append({
                "Наименование оборудования": name,
                "Бренд оборудования": brand,
                "Модели оборудования": model_list,
                "KKS коды": kks_list
            })

    # 3. Сохраняем результат
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(final_output, f, ensure_ascii=False, indent=4)

    print(f"✅ Успешно обработано файлов: {processed_count}/{len(json_files)}")
    print(f"💾 Сводная ведомость сохранена в файл: {output_file.resolve()}")


if __name__ == "__main__":
    main()