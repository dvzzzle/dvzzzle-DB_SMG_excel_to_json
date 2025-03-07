import pandas as pd
import json
import os
from datetime import datetime, time

def convert_excel_to_json(excel_file_path):
    # Чтение Excel-файла
    xls = pd.ExcelFile(excel_file_path)

    # Получение директории исходного файла
    directory = os.path.dirname(excel_file_path)

    # Проверка наличия папки ДРНТ и создание, если она не существует
    file_directory = os.path.join(directory, 'ДГС')
    if not os.path.exists(file_directory):
        os.makedirs(file_directory)

    # Чтение данных с листа "4 - ОИВ план"
    df_plan = pd.read_excel(xls, sheet_name="4 - ОИВ план", header=1, skiprows=[2])

    # Список листов, которые нужно обработать
    sheets_to_process = [
        "1 - СМГ ежедневный",
        "2.1 - СМГ срывы и действия",
        "2.2 - СМГ культура произв-ва",
        "3 - ОИВ ресурсы план (мес.)",
        "4 - ОИВ план",
        "5 - ОИВ факт"
    ]

    # Список названий ключей, к которым нужно добавить "СТРЭТАП"
    keys_to_modify = [
        'СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (фактическая дата начала)', 'СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (фактическая дата завершения)',
        'СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (% выполнения, план)', 'СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (% выполнения, факт)',
        'КОНСТРУКТИВ (фактическая дата начала)', 'КОНСТРУКТИВ (фактическая дата завершения)',
        'КОНСТРУКТИВ (% выполнения, план)', 'КОНСТРУКТИВ (% выполнения, факт)',
        'КОЛ-ВО КОРПУСОВ (фактическая дата начала)', 'КОЛ-ВО КОРПУСОВ (фактическая дата завершения)',
        'КОЛ-ВО КОРПУСОВ (% выполнения, план)', 'КОЛ-ВО КОРПУСОВ (% выполнения, факт)',
        'КОЛ-ВО ЭТАЖЕЙ (фактическая дата начала)', 'КОЛ-ВО ЭТАЖЕЙ (фактическая дата завершения)',
        'КОЛ-ВО ЭТАЖЕЙ (% выполнения, план)', 'КОЛ-ВО ЭТАЖЕЙ (% выполнения, факт)',
        'СТЕНЫ И ПЕРЕГОРОДКИ (фактическая дата начала)', 'СТЕНЫ И ПЕРЕГОРОДКИ (фактическая дата завершения)',
        'СТЕНЫ И ПЕРЕГОРОДКИ (% выполнения, план)', 'СТЕНЫ И ПЕРЕГОРОДКИ (% выполнения, факт)',
        'ФАСАД (фактическая дата начала)', 'ФАСАД (фактическая дата завершения)',
        'ФАСАД (% выполнения, план)', 'ФАСАД (% выполнения, факт)',
        'УТЕПЛИТЕЛЬ (фактическая дата начала)', 'УТЕПЛИТЕЛЬ (фактическая дата завершения)',
        'УТЕПЛИТЕЛЬ (% выполнения, план)', 'УТЕПЛИТЕЛЬ (% выполнения, факт)',
        'ФАСАДНАЯ СИСТЕМА (фактическая дата начала)', 'ФАСАДНАЯ СИСТЕМА (фактическая дата завершения)',
        'ФАСАДНАЯ СИСТЕМА (% выполнения, план)', 'ФАСАДНАЯ СИСТЕМА (% выполнения, факт)',
        'ВНУТРЕННЯЯ ОТДЕЛКА (фактическая дата начала)', 'ВНУТРЕННЯЯ ОТДЕЛКА (фактическая дата завершения)',
        'ВНУТРЕННЯЯ ОТДЕЛКА (% выполнения, план)', 'ВНУТРЕННЯЯ ОТДЕЛКА (% выполнения, факт)',
        'ЧЕРНОВАЯ ОТДЕЛКА (фактическая дата начала)', 'ЧЕРНОВАЯ ОТДЕЛКА (фактическая дата завершения)',
        'ЧЕРНОВАЯ ОТДЕЛКА (% выполнения, план)', 'ЧЕРНОВАЯ ОТДЕЛКА (% выполнения, факт)',
        'ЧИСТОВАЯ ОТДЕЛКА (фактическая дата начала)', 'ЧИСТОВАЯ ОТДЕЛКА (фактическая дата завершения)',
        'ЧИСТОВАЯ ОТДЕЛКА (% выполнения, план)', 'ЧИСТОВАЯ ОТДЕЛКА (% выполнения, факт)',
        'ВНУТРЕННИЕ СЕТИ (фактическая дата начала)', 'ВНУТРЕННИЕ СЕТИ (фактическая дата завершения)',
        'ВНУТРЕННИЕ СЕТИ (% выполнения, план)', 'ВНУТРЕННИЕ СЕТИ (% выполнения, факт)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (фактическая дата начала)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (фактическая дата завершения)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (% выполнения, план)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (% выполнения, факт)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (фактическая дата начала)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (фактическая дата завершения)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (% выполнения, план)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (% выполнения, факт)',
        'ВЕНТИЛЯЦИЯ (фактическая дата начала)', 'ВЕНТИЛЯЦИЯ (фактическая дата завершения)',
        'ВЕНТИЛЯЦИЯ (% выполнения, план)', 'ВЕНТИЛЯЦИЯ (% выполнения, факт)',
        'ЭЛЕКТРОСНАБЖЕНИЕ И СКС (фактическая дата начала)', 'ЭЛЕКТРОСНАБЖЕНИЕ И СКС (фактическая дата завершения)',
        'ЭЛЕКТРОСНАБЖЕНИЕ И СКС (% выполнения, план)', 'ЭЛЕКТРОСНАБЖЕНИЕ И СКС (% выполнения, факт)',
        'НАРУЖНЫЕ СЕТИ (фактическая дата начала)', 'НАРУЖНЫЕ СЕТИ (фактическая дата завершения)',
        'НАРУЖНЫЕ СЕТИ (% выполнения, план)', 'НАРУЖНЫЕ СЕТИ (% выполнения, факт)',
        'БЛАГОУСТРОЙСТВО (фактическая дата начала)', 'БЛАГОУСТРОЙСТВО (фактическая дата завершения)',
        'БЛАГОУСТРОЙСТВО (% выполнения, план)', 'БЛАГОУСТРОЙСТВО (% выполнения, факт)',
        'ТВЕРДОЕ ПОКРЫТИЕ (фактическая дата начала)', 'ТВЕРДОЕ ПОКРЫТИЕ (фактическая дата завершения)',
        'ТВЕРДОЕ ПОКРЫТИЕ (% выполнения, план)', 'ТВЕРДОЕ ПОКРЫТИЕ (% выполнения, факт)',
        'ОЗЕЛЕНЕНИЕ (фактическая дата начала)', 'ОЗЕЛЕНЕНИЕ (фактическая дата завершения)',
        'ОЗЕЛЕНЕНИЕ (% выполнения, план)', 'ОЗЕЛЕНЕНИЕ (% выполнения, факт)',
        'МАФ (фактическая дата начала)', 'МАФ (фактическая дата завершения)',
        'МАФ (% выполнения, план)', 'МАФ (% выполнения, факт)',
        'ОБОРУДОВАНИЕ ПО ТХЗ (фактическая дата начала)', 'ОБОРУДОВАНИЕ ПО ТХЗ (фактическая дата завершения)',
        'ОБОРУДОВАНИЕ ПО ТХЗ (% выполнения, план)', 'ОБОРУДОВАНИЕ ПО ТХЗ (% выполнения, факт)',
        'ПОСТАВЛЕНО НА ОБЪЕКТ (фактическая дата начала)', 'ПОСТАВЛЕНО НА ОБЪЕКТ (фактическая дата завершения)',
        'ПОСТАВЛЕНО НА ОБЪЕКТ (% выполнения, план)', 'ПОСТАВЛЕНО НА ОБЪЕКТ (% выполнения, факт)',
        'СМОНТИРОВАНО (фактическая дата начала)', 'СМОНТИРОВАНО (фактическая дата завершения)',
        'СМОНТИРОВАНО (% выполнения, план)', 'СМОНТИРОВАНО (% выполнения, факт)',
        'МЕБЕЛЬ (фактическая дата начала)', 'МЕБЕЛЬ (фактическая дата завершения)',
        'МЕБЕЛЬ (% выполнения, план)', 'МЕБЕЛЬ (% выполнения, факт)',
        'КЛИНИНГ (фактическая дата начала)', 'КЛИНИНГ (фактическая дата завершения)',
        'КЛИНИНГ (% выполнения, план)', 'КЛИНИНГ (% выполнения, факт)',
        'Дорожное покрытие (фактическая дата начала)', 'Дорожное покрытие (фактическая дата завершения)',
        'Дорожное покрытие (% выполнения, план)', 'Дорожное покрытие (% выполнения, факт)',
        'Искусственные сооружения (ИССО) (фактическая дата начала)',
        'Искусственные сооружения (ИССО) (фактическая дата завершения)',
        'Искусственные сооружения (ИССО) (% выполнения, план)',
        'Искусственные сооружения (ИССО) (% выполнения, факт)',
        'Наружные инженерные сети (фактическая дата начала)',
        'Наружные инженерные сети (фактическая дата завершения)',
        'Наружные инженерные сети (% выполнения, план)',
        'Наружные инженерные сети (% выполнения, факт)',
        'Средства организации дорожного движения (фактическая дата начала)',
        'Средства организации дорожного движения (фактическая дата завершения)',
        'Средства организации дорожного движения (% выполнения, план)',
        'Средства организации дорожного движения (% выполнения, факт)',
        'Дорожная разметка (фактическая дата начала)', 'Дорожная разметка (фактическая дата завершения)',
        'Дорожная разметка (% выполнения, план)', 'Дорожная разметка (% выполнения, факт)',
        'Благоустройство прилегающей территории (фактическая дата начала)',
        'Благоустройство прилегающей территории (фактическая дата завершения)',
        'Благоустройство прилегающей территории (% выполнения, план)',
        'Благоустройство прилегающей территории (% выполнения, факт)',
        'ОЦЕНКА СТРОИТЕЛЬНОГО ГОРОДКА ', 'ПАСПОРТ ОБЪЕКТА (наличие)',
        'ПАСПОРТ ОБЪЕКТА (оценка)', 'ПЕРИМЕТРАЛЬНОЕ ОГРАЖДЕНИЕ (наличие)',
        'ПЕРИМЕТРАЛЬНОЕ ОГРАЖДЕНИЕ (оценка)', 'СИГНАЛЬНОЕ ОГРАЖДЕНИЕ "ГИРЛЯНДА" (наличие)',
        'СИГНАЛЬНОЕ ОГРАЖДЕНИЕ "ГИРЛЯНДА" (оценка)', 'ПОСТ ОХРАНЫ (наличие)',
        'ПОСТ ОХРАНЫ (оценка)', 'ПУНКТ МОЙКИ КОЛЕС (наличие)', 'ПУНКТ МОЙКИ КОЛЕС (оценка)',
        'ПОКРЫТИЕ СТРОИТЕЛЬНОГО ГОРОДКА (наличие)', 'ПОКРЫТИЕ СТРОИТЕЛЬНОГО ГОРОДКА (оценка)',
        'ПОДЪЕЗДНЫХ ПУТЕЙ (наличие)', 'ПОДЪЕЗДНЫХ ПУТЕЙ (оценка)',
        'НАЛИЧИЕ И СОСТОЯНИЕ  САНИТАРНО-БЫТОВЫХ, ПРОИЗВОДСТВЕННЫХ И АДМИНИСТРАТИВНЫХ ПОМЕЩЕНИЙ И СООРУЖЕНИЙ (наличие)',
        'НАЛИЧИЕ И СОСТОЯНИЕ  САНИТАРНО-БЫТОВЫХ, ПРОИЗВОДСТВЕННЫХ И АДМИНИСТРАТИВНЫХ ПОМЕЩЕНИЙ И СООРУЖЕНИЙ (оценка)',
        'ЧИСТОТА И ПОРЯДОК НА ТЕРРИТОРИИ СТРОИТЕЛЬНОГО ГОРОДКА \n(наличие)',
        'ЧИСТОТА И ПОРЯДОК НА ТЕРРИТОРИИ СТРОИТЕЛЬНОГО ГОРОДКА\n (оценка)',
        'СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (плановая дата начала)',
        'СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (плановая дата завершения)',
        'КОНСТРУКТИВ (плановая дата начала)',
        'КОНСТРУКТИВ (плановая дата завершения)',
        'КОЛ-ВО КОРПУСОВ (плановая дата начала)',
        'КОЛ-ВО КОРПУСОВ (плановая дата завершения)',
        'КОЛ-ВО ЭТАЖЕЙ (плановая дата начала)',
        'КОЛ-ВО ЭТАЖЕЙ (плановая дата завершения)',
        'СТЕНЫ И ПЕРЕГОРОДКИ (плановая дата начала)',
        'СТЕНЫ И ПЕРЕГОРОДКИ (плановая дата завершения)',
        'ФАСАД (плановая дата начала)',
        'ФАСАД (плановая дата завершения)',
        'УТЕПЛИТЕЛЬ (плановая дата начала)',
        'УТЕПЛИТЕЛЬ (плановая дата завершения)',
        'ФАСАДНАЯ СИСТЕМА (плановая дата начала)',
        'ФАСАДНАЯ СИСТЕМА (плановая дата завершения)',
        'ВНУТРЕННЯЯ ОТДЕЛКА (плановая дата начала)',
        'ВНУТРЕННЯЯ ОТДЕЛКА (плановая дата завершения)',
        'ЧЕРНОВАЯ ОТДЕЛКА (плановая дата начала)',
        'ЧЕРНОВАЯ ОТДЕЛКА (плановая дата завершения)',
        'ЧИСТОВАЯ ОТДЕЛКА (плановая дата начала)',
        'ЧИСТОВАЯ ОТДЕЛКА (плановая дата завершения)',
        'ВНУТРЕННИЕ СЕТИ (плановая дата начала)',
        'ВНУТРЕННИЕ СЕТИ (плановая дата завершения)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (плановая дата начала)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (плановая дата завершения)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (плановая дата начала)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (плановая дата завершения)',
        'ВЕНТИЛЯЦИЯ (плановая дата начала)',
        'ВЕНТИЛЯЦИЯ (плановая дата завершения)',
        'ЭЛЕКТРОСНАБЖЕНИЕ И СКС (плановая дата начала)',
        'ЭЛЕКТРОСНАБЖЕНИЕ И СКС (плановая дата завершения)',
        'НАРУЖНЫЕ СЕТИ (плановая дата начала)',
        'НАРУЖНЫЕ СЕТИ (плановая дата завершения)',
        'БЛАГОУСТРОЙСТВО (плановая дата начала)',
        'БЛАГОУСТРОЙСТВО (плановая дата завершения)',
        'ТВЕРДОЕ ПОКРЫТИЕ (плановая дата начала)',
        'ТВЕРДОЕ ПОКРЫТИЕ (плановая дата завершения)',
        'ОЗЕЛЕНЕНИЕ (плановая дата начала)',
        'ОЗЕЛЕНЕНИЕ (плановая дата завершения)',
        'МАФ (плановая дата начала)',
        'МАФ (плановая дата завершения)',
        'ОБОРУДОВАНИЕ ПО ТХЗ (плановая дата начала)',
        'ОБОРУДОВАНИЕ ПО ТХЗ (плановая дата завершения)',
        'ПОСТАВЛЕНО НА ОБЪЕКТ (плановая дата начала)',
        'ПОСТАВЛЕНО НА ОБЪЕКТ (плановая дата завершения)',
        'СМОНТИРОВАНО (плановая дата начала)',
        'СМОНТИРОВАНО (плановая дата завершения)',
        'МЕБЕЛЬ (плановая дата начала)',
        'МЕБЕЛЬ (плановая дата завершения)',
        'КЛИНИНГ (плановая дата начала)',
        'КЛИНИНГ (плановая дата завершения)',
        'Дорожное покрытие (плановая дата начала)',
        'Дорожное покрытие (плановая дата завершения)',
        'Искусственные сооружения (ИССО) (плановая дата начала)',
        'Искусственные сооружения (ИССО) (плановая дата завершения)',
        'Наружные инженерные сети (плановая дата начала)',
        'Наружные инженерные сети  (плановая дата завершения)',
        'Средства организации дорожного движения (плановая дата начала)',
        'Средства организации дорожного движения (плановая дата завершения)',
        'Дорожная разметка (плановая дата начала)',
        'Дорожная разметка (плановая дата завершения)',
        'Благоустройство прилегающей территории (плановая дата начала)',
        'Благоустройство прилегающей территории (плановая дата завершения)',
        'СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (план)',
        'СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (факт)',
        'КОНСТРУКТИВ (план)',
        'КОНСТРУКТИВ (факт)',
        'КОЛ-ВО КОРПУСОВ (план)',
        'КОЛ-ВО КОРПУСОВ (факт)',
        'КОЛ-ВО ЭТАЖЕЙ (план)',
        'КОЛ-ВО ЭТАЖЕЙ (факт)',
        'СТЕНЫ И ПЕРЕГОРОДКИ (план)',
        'СТЕНЫ И ПЕРЕГОРОДКИ (факт)',
        'ФАСАД (план)',
        'ФАСАД (факт)',
        'УТЕПЛИТЕЛЬ (план)',
        'УТЕПЛИТЕЛЬ (факт)',
        'ФАСАДНАЯ СИСТЕМА (план)',
        'ФАСАДНАЯ СИСТЕМА (факт)',
        'ВНУТРЕННЯЯ ОТДЕЛКА (план)',
        'ВНУТРЕННЯЯ ОТДЕЛКА (факт)',
        'ЧЕРНОВАЯ ОТДЕЛКА (план)',
        'ЧЕРНОВАЯ ОТДЕЛКА (факт)',
        'ЧИСТОВАЯ ОТДЕЛКА (план)',
        'ЧИСТОВАЯ ОТДЕЛКА (факт)',
        'ВНУТРЕННИЕ СЕТИ (план)',
        'ВНУТРЕННИЕ СЕТИ (факт)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (план)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (факт)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (план)',
        'ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (факт)',
        'ВЕНТИЛЯЦИЯ (план)',
        'ВЕНТИЛЯЦИЯ (факт)',
        'ЭЛЕКТРОСНАБЖЕНИЕ И СКС (план)',
        'ЭЛЕКТРОСНАБЖЕНИЕ И СКС (факт)',
        'НАРУЖНЫЕ СЕТИ (план)',
        'НАРУЖНЫЕ СЕТИ (факт)',
        'БЛАГОУСТРОЙСТВО (план)',
        'БЛАГОУСТРОЙСТВО (факт)',
        'ТВЕРДОЕ ПОКРЫТИЕ (план)',
        'ТВЕРДОЕ ПОКРЫТИЕ (факт)',
        'ОЗЕЛЕНЕНИЕ (план)',
        'ОЗЕЛЕНЕНИЕ (факт)',
        'МАФ (план)',
        'МАФ (факт)',
        'ОБОРУДОВАНИЕ ПО ТХЗ (план)',
        'ОБОРУДОВАНИЕ ПО ТХЗ (факт)',
        'ПОСТАВЛЕНО НА ОБЪЕКТ (план)',
        'ПОСТАВЛЕНО НА ОБЪЕКТ (факт)',
        'СМОНТИРОВАНО (план)',
        'СМОНТИРОВАНО (факт)',
        'МЕБЕЛЬ (план)',
        'МЕБЕЛЬ (факт)',
        'КЛИНИНГ (план)',
        'КЛИНИНГ (факт)',
        'Дорожное покрытие (план)',
        'Дорожное покрытие (факт)',
        'Искусственные сооружения (ИССО) (план)',
        'Искусственные сооружения (ИССО) (факт)',
        'Наружные инженерные сети (план)',
        'Наружные инженерные сети  (факт)',
        'Средства организации дорожного движения (план)',
        'Средства организации дорожного движения (факт)',
        'Дорожная разметка (план)',
        'Дорожная разметка (факт)',
        'Благоустройство прилегающей территории (план)',
        'Благоустройство прилегающей территории (факт)'
    ]

    # Обработка каждого листа
    for sheet_name in sheets_to_process:
        if sheet_name in xls.sheet_names:
            if sheet_name == "2.1 - СМГ срывы и действия" or sheet_name == "2.2 - СМГ культура произв-ва" or sheet_name == "3 - ОИВ ресурсы план (мес.)":
                df = pd.read_excel(xls, sheet_name=sheet_name, header=0, skiprows=[1])
            else:
                # Обработка остальных листов
                df = pd.read_excel(xls, sheet_name=sheet_name, header=1, skiprows=[2])

            # Обновление данных на листах "1 - СМГ ежедневный" и "5 - ОИВ факт"
            if sheet_name in ["1 - СМГ ежедневный", "5 - ОИВ факт"]:
                columns_to_update = [
                    'УИН', 'Мастер код ФР', 'Мастер код ДС', 'link', 'ФНО', 'Наименование',
                    'Наименование ЖК (коммерческое наименование)', 'Краткое наименование',
                    'ФНО для аналитики', 'Округ', 'Район', 'Источник финансирования',
                    'Статус ОКС', 'АИП (да/нет)', 'Дата включения в АИП', 'Сумма по АИП, млрд руб',
                    'Аванс, млрд руб', 'Адрес объекта', 'Признак реновации (да/нет)',
                    'ИНН Застройщика', 'Застройщик', 'ИНН Заказчика', 'Заказчик',
                    'ИНН Генподрядчика', 'Генподрядчик', 'Общая площадь, м2', 'Протяженность'
                ]
                for column in columns_to_update:
                    if column in df.columns and column in df_plan.columns:
                        df[column] = df_plan[column]

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

            # Преобразование значений в колонке "АИП (да/нет)" (если она существует)
            if 'АИП (да/нет)' in df.columns:
                df['АИП (да/нет)'] = df['АИП (да/нет)'].astype(str).str.lower().replace({
                    'true': 'да',
                    'false': 'нет',
                    'да': 'да',
                    'нет': 'нет'
                })

            # Удаление строк, где все указанные колонки пустые
            df.dropna(subset='УИН', how='all', inplace=True)

            # Конвертация DataFrame в словарь
            data_dict = df.to_dict(orient='records')

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
convert_excel_to_json('E://Загрузки//Telegram Desktop//тест.xlsx')
