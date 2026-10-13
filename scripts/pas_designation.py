import json
from pathlib import Path
import sys

def process_json_files(folder_path: str) -> None:
    folder = Path(folder_path)
    if not folder.is_dir():
        print(f"Ошибка: папка '{folder_path}' не найдена.")
        return

    for json_file in folder.glob("pas_*.json"):
        try:
            with open(json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            if "company" in data and "passport_code" in data["company"]:
                # Создаём metadata, если его нет
                if "metadata" not in data:
                    data["metadata"] = {}

                # Переносим ключ и удаляем из company
                data["metadata"]["document_designation"] = data["company"].pop("passport_code")

            with open(json_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            print(f"✓ Обработан: {json_file.name}")

        except json.JSONDecodeError as e:
            print(f"✗ Ошибка чтения JSON в {json_file.name}: {e}")
        except Exception as e:
            print(f"✗ Ошибка обработки {json_file.name}: {e}")

if __name__ == "__main__":
    target_dir = sys.argv[1] if len(sys.argv) > 1 else "."
    process_json_files(target_dir)