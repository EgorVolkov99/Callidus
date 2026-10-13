import json
import argparse
from pathlib import Path

# Значения для подстановки
RUS_SUPPLIER_NAME = "Поставщик, адрес"
RUS_SUPPLIER_VALUE = "NUKEM Technologies Engineering Services GmbH, Германия, 63755, Альценау, Индустриштрассе, д. 13"

ENG_SUPPLIER_NAME = "Supplier, address"
ENG_SUPPLIER_VALUE = "NUKEM Technologies Engineering Services GmbH, 13 Industriestrasse, 63755 Alzenau, Germany"


def update_supplier_in_list(items: list, target_name: str, target_val: str, alt_names: list) -> bool:
    """
    Ищет в списке словарей элемент по названию (или синонимам) и обновляет значение.
    Если элемент не найден — добавляет его в конец.
    """
    all_names_to_check = [target_name.lower()] + [n.lower() for n in alt_names]

    # 1. Поиск существующего элемента
    for item in items:
        if isinstance(item, dict):
            item_name = str(item.get("name", "")).strip().lower()
            if item_name in all_names_to_check:
                item["name"] = target_name  # Приводим к эталонному написанию
                item["value"] = target_val
                return True

    # 2. Если поле не было найдено — добавляем новое
    items.append({
        "name": target_name,
        "value": target_val
    })
    return True


def process_json_file(file_path: Path) -> bool:
    """Обрабатывает один JSON файл."""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        basic_info = data.get("basic_info")
        if not isinstance(basic_info, dict):
            print(f"⚠️ Пропуск: {file_path.name} (нет секции basic_info)")
            return False

        # Обновляем русскую секцию
        if "rus" in basic_info and isinstance(basic_info["rus"], list):
            update_supplier_in_list(
                items=basic_info["rus"],
                target_name=RUS_SUPPLIER_NAME,
                target_val=RUS_SUPPLIER_VALUE,
                alt_names=["Supplier, address", "Поставщик"]
            )

        # Обновляем английскую секцию
        if "eng" in basic_info and isinstance(basic_info["eng"], list):
            update_supplier_in_list(
                items=basic_info["eng"],
                target_name=ENG_SUPPLIER_NAME,
                target_val=ENG_SUPPLIER_VALUE,
                alt_names=["Поставщик, адрес", "Supplier"]
            )

        # Сохраняем обновленный JSON с сохранением форматирования и UTF-8
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        return True

    except Exception as e:
        print(f"❌ Ошибка в файле {file_path.name}: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description="Пакетное обновление поставщика в JSON-файлах")
    parser.add_argument(
        "--folder",
        type=str,
        default="passport/projects/NIKIMT_KIP/input",
        help="Путь к папке с .json файлами"
    )
    args = parser.parse_args()

    target_dir = Path(args.folder)

    if not target_dir.exists() or not target_dir.is_dir():
        print(f"❌ Ошибка: Папка не найдена: {target_dir.resolve()}")
        return

    json_files = list(target_dir.glob("*.json"))
    if not json_files:
        print(f"⚠️ В папке '{target_dir}' не найдено ни одного .json файла.")
        return

    print(f"🔄 Обработка {len(json_files)} файлов в директории: {target_dir.resolve()}\n")

    updated_count = 0
    for file_path in json_files:
        if process_json_file(file_path):
            print(f"  ✓ Обновлен: {file_path.name}")
            updated_count += 1

    print(f"\n✅ Завершено! Успешно обновлено файлов: {updated_count}/{len(json_files)}")


if __name__ == "__main__":
    main()