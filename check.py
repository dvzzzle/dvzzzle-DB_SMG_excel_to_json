import json
import os
from pathlib import Path
from collections.abc import Mapping, Sequence

# Коды, которые мы ищем
target_codes = {"018-0281", "018-0279"}

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

def find_items_with_codes(data, current_path=None):
    """Рекурсивно ищет элементы с целевыми кодами в структуре данных"""
    if current_path is None:
        current_path = []

    found_items = []

    # Если это словарь и содержит нужный код
    if isinstance(data, Mapping):
        if "code" in data and data["code"] in target_codes:
            found_items.append((data, current_path))
        
        # Рекурсивно проверяем все значения словаря
        for key, value in data.items():
            found_items.extend(find_items_with_codes(value, current_path + [key]))

    # Если это список или другая последовательность
    elif isinstance(data, Sequence) and not isinstance(data, str):
        for index, item in enumerate(data):
            found_items.extend(find_items_with_codes(item, current_path + [index]))

    return found_items

def process_directory(directory_path):
    """Обрабатывает все файлы в указанной директории"""
    directory_path = Path(directory_path).absolute()
    print(f"Обрабатываю файлы в директории: {directory_path}")

    result_data = []
    processed_files = 0
    found_items_count = 0

    for file_name in files_to_process:
        file_path = directory_path / file_name
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
                processed_files += 1
                
                # Ищем все элементы с целевыми кодами
                found_items = find_items_with_codes(data)
                
                for item, path in found_items:
                    # Проверяем appr_flag != "0"
                    if item.get("appr_flag") != "0":
                        # Создаем отфильтрованный элемент
                        filtered_item = {
                            key: item.get(key) 
                            for key in required_keys 
                            if key in item
                        }
                        # Добавляем код и путь для отладки
                        filtered_item["code"] = item["code"]
                        filtered_item["source_path"] = {
                            "file": file_name,
                            "path": path
                        }
                        result_data.append(filtered_item)
                        found_items_count += 1

                print(f"Обработан {file_name} - найдено {len(found_items)} совпадений")

        except FileNotFoundError:
            print(f"Файл {file_name} не найден, пропускаем...")
        except json.JSONDecodeError:
            print(f"Ошибка декодирования JSON в файле {file_name}, пропускаем...")
        except Exception as e:
            print(f"Неизвестная ошибка при обработке файла {file_name}: {e}")

    # Сохраняем результат
    output_path = directory_path / "result.json"
    with open(output_path, 'w', encoding='utf-8') as out_file:
        json.dump(result_data, out_file, ensure_ascii=False, indent=4)

    print(f"\nИтоги:")
    print(f"- Обработано файлов: {processed_files}")
    print(f"- Найдено подходящих записей: {found_items_count}")
    print(f"- Результат сохранен в {output_path}")

if __name__ == "__main__":
    # Введите путь к директории
    directory = "E://Загрузки//Telegram Desktop"
    
    # Проверяем, существует ли директория
    if not os.path.isdir(directory):
        print(f"Ошибка: директория '{directory}' не существует!")
    else:
        process_directory(directory)