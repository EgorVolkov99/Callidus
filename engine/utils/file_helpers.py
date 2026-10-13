import re

def generate_safe_filename(base_name: str, extension: str) -> str:
    """Удаляет запрещенные символы из имени файла."""
    safe_name = re.sub(r'[\\/*?:"<>|]', "_", base_name)
    safe_name = " ".join(safe_name.split()) # Убираем двойные пробелы
    return f"{safe_name}.{extension.strip('.')}"