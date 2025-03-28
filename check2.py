import json
import os
from pathlib import Path
from collections.abc import Mapping, Sequence

# Коды, которые мы ищем
target_codes = {
"025-0130","025-0138","020-0777","025-0140","025-0201","025-0141","025-0149","025-0144",
"025-0174","019-0452","022-0701","020-0178","025-0200","025-0137","025-0133","025-0132",
"025-0134","025-0139","024-0837","014-2246","023-0073","025-0202","018-0201","020-0127",
"025-0146","015-0719","025-0147","025-0172","024-0847","022-0645","017-0075","025-0173",
"019-0094","025-0170","022-0132","016-0763","018-0288","019-0441","022-0124","024-0296",
"022-0505","021-0453","020-0134","023-0428","020-0144","013-0345","020-0694","025-0136",
"023-0771","025-0204","016-1513","021-0153","023-0098","023-0099","024-0348","025-0053",
"025-0155","023-0095","024-0361","025-0189","019-0450","018-0278","013-1344","021-0452",
"025-0052","023-0403","018-0262","024-0121","025-0192","024-0283","025-0171","025-0169",
"017-0204","023-0565","023-0067","022-0601","025-0194","023-0213","025-0117","022-0315",
"023-0064","025-0153","018-0132","023-0177","023-0070","021-0190","024-0349","021-0522",
"025-0055","022-0054","019-0132","025-0050","024-0726","023-0176","023-0750","023-0180",
"025-0156","017-0499","024-0650","023-0272","023-0405","022-0504","023-0913","023-0579",
"024-0685","024-0659","024-0689","025-0166","024-0274","017-0539","025-0051","023-0732",
"024-0866","022-0530","025-0152","018-0342","019-0451","023-0074","022-0401","018-0241",
"025-0120","022-0400","023-0562","023-0404","024-0297","023-1137","023-0068","021-0520",
"022-0403","019-0133","023-0591","022-0715","025-0119","024-0352","025-0154","023-0896",
"023-0495","023-0434","023-0059","025-0162","023-0065","025-0157","025-0054","024-0657",
"024-0684","024-0662","024-0656","024-0580","024-0584","024-0593","023-0589","023-0550",
"023-0578","023-0952","024-0602","024-0604","024-0607","024-0609","024-0588","024-0589",
"024-0596","024-0599","024-0645","023-0250",
}

# Ключи, которые должны присутствовать в результате
required_keys = {
    "full_name", "sost_object", "zamestitel", "building_address", "raion", "prefecture",
    "developer", "developer_inn", "builder", "builder_inn", "aip_max_cost", "aip_source_finance",
    "entry_year_approved", "rv_prognoz", "title_number", "title_doc_num", "title_doc_date",
    "appr_flag", "atr_total_area", "atr_liv_total_area", "atr_length_road_km", "kolvo_plan"
}

# Список файлов для обработки
files_to_process = [
    "zadachi.json",
    "vypolnenie.json",
    "vvod.json",
    "titles.json",
    "teps.json",
    "osvoenie.json",
    "organizations.json",
    "objects-mgz.json",
    "finance.json",
    "chislennost.json",
    "aip.json"
]

def find_items_with_codes(data):
    """Рекурсивно ищет элементы с целевыми кодами в структуре данных"""
    found_items = []

    # Если это словарь и содержит нужный код
    if isinstance(data, Mapping):
        if "code" in data and data["code"] in target_codes:
            found_items.append(data)
        
        # Рекурсивно проверяем все значения словаря
        for value in data.values():
            found_items.extend(find_items_with_codes(value))

    # Если это список или другая последовательность
    elif isinstance(data, Sequence) and not isinstance(data, str):
        for item in data:
            found_items.extend(find_items_with_codes(item))

    return found_items

def merge_data_items(items):
    """Объединяет данные для одного кода из разных файлов"""
    merged = {}
    for item in items:
        # Объединяем только разрешенные ключи
        for key in required_keys:
            if key in item and item[key] is not None:
                # Если значение уже есть, но новое не None, обновляем
                if key not in merged or merged[key] is None:
                    merged[key] = item[key]
        # Сохраняем код
        merged["code"] = item["code"]
    return merged

def process_directory(directory_path):
    """Обрабатывает все файлы в указанной директории"""
    directory_path = Path(directory_path).absolute()
    print(f"Обрабатываю файлы в директории: {directory_path}")

    # Словарь для хранения данных по кодам: {code: [items]}
    code_data = {code: [] for code in target_codes}

    processed_files = 0

    for file_name in files_to_process:
        file_path = directory_path / file_name
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
                processed_files += 1
                
                # Ищем все элементы с целевыми кодами
                found_items = find_items_with_codes(data)
                
                for item in found_items:
                    code = item["code"]
                    if code in code_data and item.get("appr_flag") != "0":
                        code_data[code].append(item)

                print(f"Обработан {file_name} - найдено {len(found_items)} совпадений")

        except FileNotFoundError:
            print(f"Файл {file_name} не найден, пропускаем...")
        except json.JSONDecodeError:
            print(f"Ошибка декодирования JSON в файле {file_name}, пропускаем...")
        except Exception as e:
            print(f"Неизвестная ошибка при обработке файла {file_name}: {e}")

    # Объединяем данные для каждого кода
    result_data = []
    for code, items in code_data.items():
        if items:  # Если есть данные для этого кода
            merged_item = merge_data_items(items)
            result_data.append(merged_item)

    # Сохраняем результат
    output_path = directory_path / "merged_result.json"
    with open(output_path, 'w', encoding='utf-8') as out_file:
        json.dump(result_data, out_file, ensure_ascii=False, indent=4)

    print(f"\nИтоги:")
    print(f"- Обработано файлов: {processed_files}")
    print(f"- Найдено уникальных кодов с данными: {len(result_data)}")
    print(f"- Результат сохранен в {output_path}")

if __name__ == "__main__":
    # Введите путь к директории
    directory = "E://Загрузки//Telegram Desktop"
    
    if not os.path.isdir(directory):
        print(f"Ошибка: директория '{directory}' не существует!")
    else:
        process_directory(directory)