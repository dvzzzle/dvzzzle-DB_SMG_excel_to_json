import json

def count_top_level_datasets_from_file(filepath):
    print("Начало обработки файла")  # Отладочное сообщение

    try:
        with open(filepath, 'r', encoding='utf-8') as f: # Укажите кодировку явно
            data = json.load(f)

        if isinstance(data, dict):
            result = len(data)
        elif isinstance(data, list):
            result = len(data)
        else:
            result = 0
        print(f"Количество наборов данных: {result}") # Отладочное сообщение
        return result


    except FileNotFoundError:
        print("Ошибка: Файл не найден.")
        return -1
    except json.JSONDecodeError as e:
        print(f"Ошибка: Некорректный формат JSON: {e}")
        return -1
    except UnicodeDecodeError as e:
        print(f"Ошибка: Проблема с кодировкой: {e}")
        return -1
    except Exception as e:
        print(f"Произошла непредвиденная ошибка: {e}")
        return -1
    finally:
        print("Функция завершена") # Отладочное сообщение

# Пример использования:
filepath = "E://Загрузки//1 - СМГ ежедневный.json"  # Замените на реальный путь к вашему файлу
print(f"Путь к файлу: {filepath}") # Отладочное сообщение

count = count_top_level_datasets_from_file(filepath)
print(f"count = {count}") # Отладочное сообщение

if count != -1:
    print(f"Количество наборов данных в файле: {count}")
else:
    print("Не удалось обработать файл JSON.")
