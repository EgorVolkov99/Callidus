from typing import Dict, List, Any


def flatten_parallel_arrays(data_dict: Dict[str, Any]) -> List[List[Any]]:
    """
    Превращает словарь параллельных массивов в плоский список списков.
    Пример:
    {"Name": ["A", "B"], "Qty": [1, 2]} -> [["A", 1], ["B", 2]]
    """
    if not data_dict or not isinstance(data_dict, dict):
        return []

    keys = list(data_dict.keys())
    if not keys or not isinstance(data_dict[keys[0]], list):
        return []

    num_rows = len(data_dict[keys[0]])
    flat_rows = []

    for i in range(num_rows):
        # Если в каком-то массиве не хватает элементов, подставляем пустую строку
        row_vals = [data_dict[k][i] if i < len(data_dict[k]) else "" for k in keys]
        flat_rows.append(row_vals)

    return flat_rows