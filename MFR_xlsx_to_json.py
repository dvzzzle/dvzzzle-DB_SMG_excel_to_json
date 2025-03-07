import pandas as pd
import json
import os
from datetime import datetime, time

def convert_excel_to_json(excel_file_path):
    # Чтение Excel-файла
    xls = pd.ExcelFile(excel_file_path)

    # Получение директории исходного файла
    directory = os.path.dirname(excel_file_path)

    # Проверка наличия папки ДСТИИ и создание, если она не существует
    file_directory = os.path.join(directory, 'ДГП')
    if not os.path.exists(file_directory):
        os.makedirs(file_directory)

    # Список листов, которые нужно обработать
    sheets_to_process = [
        "1 - СМГ ежедневный",
        "2.1 - СМГ срывы и действия",
        "2.2 - СМГ культура произв-ва",
        "3 - ОИВ ресурсы план (мес.)",
        "4 - ОИВ план",
        "5 - ОИВ факт",
        "6 - ОИВ КТ"
    ]

    # Обработка каждого листа
    for sheet_name in sheets_to_process:
        if sheet_name in xls.sheet_names:
            if sheet_name == "6 - ОИВ КТ":
                # Обработка листа "6 - ОИВ КТ"
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

                # Получение заголовков из первых трех строк
                headers = df.iloc[0:3].ffill(axis=1).fillna('').values.tolist()

                # Пропуск четвертой строки
                data_rows = df.iloc[4:].values.tolist()

                # Удаление колонок, содержащих слово "комментарий" в любом регистре
                columns_to_drop = [col_idx for col_idx in range(len(headers[0])) 
                                  if 'комментарий||титул||год титула' in headers[0][col_idx].lower() or 
                                     'комментарий||титул||год титула' in headers[1][col_idx].lower() or 
                                     'комментарий||титул||год титула' in headers[2][col_idx].lower()]
                headers = [[header for col_idx, header in enumerate(row) if col_idx not in columns_to_drop] 
                           for row in headers]
                data_rows = [[value for col_idx, value in enumerate(row) if col_idx not in columns_to_drop] 
                             for row in data_rows]

                # Создание структуры JSON
                data_dict = []
                for row in data_rows:
                    record = {}
                    current_level = record
                    for col_idx in range(len(row)):
                        # Получение уровней вложенности
                        level1 = headers[0][col_idx].replace('\n', ' ')
                        level2 = headers[1][col_idx].replace('\n', ' ')
                        level3 = headers[2][col_idx].replace('\n', ' ')

                        # Если первая строка не пустая, создаем первый уровень вложенности
                        if level1:
                            level1 = f"ОБЩКОНТРТОЧКА {level2}"
                            if level1 not in current_level:
                                current_level[level1] = {}
                            current_level = current_level[level1]

                        # Если вторая строка не пустая, создаем второй уровень вложенности
                        if level2:
                            level2 = f"КОНТРТОЧКА {level2}"  # Добавляем "КОНТРТОЧКА" к level2
                            if level2 not in current_level:
                                current_level[level2] = {}
                            current_level = current_level[level2]

                        # Добавляем значение в третий уровень вложенности
                        if level3:
                            value = row[col_idx]
                            if isinstance(value, (datetime, pd.Timestamp)):
                                value = value.strftime('%d.%m.%Y')
                            elif isinstance(value, time):
                                value = value.strftime('%H:%M:%S')
                            elif pd.isna(value) or value == '':
                                value = None

                            # Добавляем значение в current_level
                            current_level[level3] = value

                            # Если значение равно "не требуется", удаляем ключ
                            if value == "не требуется":
                                del current_level[level3]

                            current_level = record  # Возвращаемся к корневому уровню

                    data_dict.append(record)

                # Формирование пути для JSON-файла
                json_file_path = os.path.join(file_directory, f'{sheet_name}.json')

                # Запись данных в JSON-файл
                with open(json_file_path, 'w', encoding='utf-8') as json_file:
                    json.dump(data_dict, json_file, ensure_ascii=False, indent=4)

                print(f'Лист "{sheet_name}" успешно конвертирован в файл "{json_file_path}".')
            else:
                if sheet_name == "2.1 - СМГ срывы и действия" or sheet_name == "2.2 - СМГ культура произв-ва" or sheet_name == "3 - ОИВ ресурсы план (мес.)":
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=0, skiprows=[1])
                else:
                    # Обработка остальных листов
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=1, skiprows=[2])

                # Обработка листа "4 - ОИВ план"
                if sheet_name == "4 - ОИВ план":
                    # Проверка наличия нужных колонок
                    if 'Плановый ввод по директивному графику' in df.columns and 'Плановый ввод по договору' in df.columns:
                        # Заполнение пустых значений в колонке "Плановый ввод по директивному графику"
                        df['Плановый ввод по директивному графику'] = df.apply(
                            lambda row: row['Плановый ввод по договору'] if pd.isna(row['Плановый ввод по директивному графику']) else row['Плановый ввод по директивному графику'],
                            axis=1
                        )

                # Удаление колонок с комментариями, титулами и годами титулов
                df = df.loc[:, ~df.columns.str.contains('комментарий|титул|год титула', case=False)]

                # Удаление указанных колонок в листах "1 - СМГ ежедневный", "4 - ОИВ план" и "5 - ОИВ факт"
                if sheet_name in ["1 - СМГ ежедневный", "4 - ОИВ план", "5 - ОИВ факт"]:
                    columns_to_drop = ["АИП (да/нет)", "Дата включения в АИП", "Сумма по АИП, млрд руб", "Аванс, млрд руб"]
                    df.drop(columns=[col for col in columns_to_drop if col in df.columns], inplace=True)

                # Удаление строк, где все три колонки 'УИН', 'Мастер код ФР', 'Мастер код ДС' пустые
                if sheet_name in ["1 - СМГ ежедневный", "4 - ОИВ план", "5 - ОИВ факт"]:
                                # Удаление строк, где все указанные колонки пустые
                                df.dropna(subset='УИН', how='all', inplace=True)


                # Конвертация DataFrame в словарь
                data_dict = df.to_dict(orient='records')

                # Список названий ключей, к которым нужно добавить "СТРЭТАП"
                keys_to_modify = [
                    'СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (план)', 'СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (факт)',
                    'КОНСТРУКТИВ (план)', 'КОНСТРУКТИВ (факт)', 'КОЛ-ВО КОРПУСОВ (план)',
                    'КОЛ-ВО КОРПУСОВ (факт)', 'КОЛ-ВО ЭТАЖЕЙ (план)', 'КОЛ-ВО ЭТАЖЕЙ (факт)',
                    'СТЕНЫ И ПЕРЕГОРОДКИ (план)', 'СТЕНЫ И ПЕРЕГОРОДКИ (факт)', 'ФАСАД (план)',
                    'ФАСАД (факт)', 'УТЕПЛИТЕЛЬ (план)', 'УТЕПЛИТЕЛЬ (факт)',
                    'ФАСАДНАЯ СИСТЕМА (план)', 'ФАСАДНАЯ СИСТЕМА (факт)',
                    'ВНУТРЕННЯЯ ОТДЕЛКА (план)', 'ВНУТРЕННЯЯ ОТДЕЛКА (факт)',
                    'ЧЕРНОВАЯ ОТДЕЛКА (план)', 'ЧЕРНОВАЯ ОТДЕЛКА (факт)',
                    'ЧИСТОВАЯ ОТДЕЛКА (план)', 'ЧИСТОВАЯ ОТДЕЛКА (факт)',
                    'ВНУТРЕННИЕ СЕТИ (план)', 'ВНУТРЕННИЕ СЕТИ (факт)',
                    'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (план)',
                    'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (факт)',
                    'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (план)',
                    'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (факт)',
                    'ВЕНТИЛЯЦИЯ (план)', 'ВЕНТИЛЯЦИЯ (факт)',
                    'ЭЛЕКТРОСНАБЖЕНИЕ И СКС (план)', 'ЭЛЕКТРОСНАБЖЕНИЕ И СКС (факт)',
                    'НАРУЖНЫЕ СЕТИ (план)', 'НАРУЖНЫЕ СЕТИ (факт)',
                    'БЛАГОУСТРОЙСТВО (план)', 'БЛАГОУСТРОЙСТВО (факт)',
                    'ТВЕРДОЕ ПОКРЫТИЕ (план)', 'ТВЕРДОЕ ПОКРЫТИЕ (факт)',
                    'ОЗЕЛЕНЕНИЕ (план)', 'ОЗЕЛЕНЕНИЕ (факт)', 'МАФ (план)', 'МАФ (факт)'
                ]

                # Преобразование объектов datetime и time в строки и удаление колонок со значением "не требуется"
                for record in data_dict:
                    new_record = {}
                    for key, value in record.items():
                        if pd.notna(value) and isinstance(value, (datetime, pd.Timestamp)):
                            value = value.strftime('%d.%m.%Y')
                        elif pd.notna(value) and isinstance(value, time):
                            value = value.strftime('%H:%M:%S')
                        elif pd.isna(value):  # Проверка на NaT или NaN
                            value = None
                        elif isinstance(value, str) and value.strip().lower() == "не требуется":
                            continue  # Пропустить добавление этого ключа в new_record

                        # Добавление "СТРЭТАП" к нужным ключам
                        if key.strip() in keys_to_modify:
                            new_record[f'СТРЭТАП {key.strip()}'] = value
                        else:
                            new_record[key] = value

                    # Замена старого словаря новым
                    record.clear()
                    record.update(new_record)

                # Формирование пути для JSON-файла
                json_file_path = os.path.join(file_directory, f'{sheet_name}.json')

                # Запись данных в JSON-файл
                with open(json_file_path, 'w', encoding='utf-8') as json_file:
                    json.dump(data_dict, json_file, ensure_ascii=False, indent=4)

                print(f'Лист "{sheet_name}" успешно конвертирован в файл "{json_file_path}".')

# Пример использования
convert_excel_to_json('E://Загрузки//Telegram Desktop//МФР_итоговый_перечень_для_ДБ_04_03_2025_с_культурой.xlsx')