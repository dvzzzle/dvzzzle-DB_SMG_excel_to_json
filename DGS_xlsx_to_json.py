import pandas as pd
import json
import os
from datetime import datetime, time

def validate_data_type(value, expected_type):
    if expected_type == 'text':
        return value if isinstance(value, str) else None
    elif expected_type == 'numeric':
        return value if isinstance(value, (int, float)) and not pd.isna(value) else None
    elif expected_type == 'ДД.ММ.ГГГГ':
        if isinstance(value, (datetime, pd.Timestamp)):
            return value.strftime('%d.%m.%Y')
        elif isinstance(value, str):
            try:
                datetime.strptime(value, '%d.%m.%Y')
                return value
            except ValueError:
                return None
        else:
            return None
    elif expected_type == 'int':
        return int(value) if isinstance(value, (int, float)) and not pd.isna(value) else None
    else:
        return value  # Если тип не указан, оставляем значение как есть
    
column_types = {
    "дата редактирования строки": "ДД.ММ.ГГГГ",
    "На какую дату актуализировано состояние ОКС": "ДД.ММ.ГГГГ",
    "На какую дату актуализировано состояник ОКС": "ДД.ММ.ГГГГ",
    "Дата включения в АИП": "ДД.ММ.ГГГГ",
    "Получение ГПЗУ факт": "ДД.ММ.ГГГГ",
    "Получение ТУ от ресурсоснабжающих организаций (факт)": "ДД.ММ.ГГГГ",
    "Разработка и согласование АГР (факт)": "ДД.ММ.ГГГГ",
    "Разработка и получение положительного заключения экспертизы ПСД (факт)": "ДД.ММ.ГГГГ",
    "Получение РНС (факт)": "ДД.ММ.ГГГГ",
    "Производство СМР (факт)": "ДД.ММ.ГГГГ",
    "Технологическое присоединение (факт)": "ДД.ММ.ГГГГ",
    "Получение ЗОС (факт)": "ДД.ММ.ГГГГ",
    "Получение РВ (факт)": "ДД.ММ.ГГГГ",
    "СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "КОНСТРУКТИВ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "КОНСТРУКТИВ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "КОЛ-ВО КОРПУСОВ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "КОЛ-ВО КОРПУСОВ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "КОЛ-ВО ЭТАЖЕЙ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "КОЛ-ВО ЭТАЖЕЙ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "СТЕНЫ И ПЕРЕГОРОДКИ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "СТЕНЫ И ПЕРЕГОРОДКИ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ФАСАД (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ФАСАД (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "УТЕПЛИТЕЛЬ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "УТЕПЛИТЕЛЬ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ФАСАДНАЯ СИСТЕМА (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ФАСАДНАЯ СИСТЕМА (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ВНУТРЕННЯЯ ОТДЕЛКА (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ВНУТРЕННЯЯ ОТДЕЛКА (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ЧЕРНОВАЯ ОТДЕЛКА (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ЧЕРНОВАЯ ОТДЕЛКА (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ЧИСТОВАЯ ОТДЕЛКА (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ЧИСТОВАЯ ОТДЕЛКА (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ВНУТРЕННИЕ СЕТИ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ВНУТРЕННИЕ СЕТИ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ВЕНТИЛЯЦИЯ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ВЕНТИЛЯЦИЯ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ЭЛЕКТРОСНАБЖЕНИЕ И СКС (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ЭЛЕКТРОСНАБЖЕНИЕ И СКС (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "НАРУЖНЫЕ СЕТИ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "НАРУЖНЫЕ СЕТИ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "БЛАГОУСТРОЙСТВО (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "БЛАГОУСТРОЙСТВО (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ТВЕРДОЕ ПОКРЫТИЕ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ТВЕРДОЕ ПОКРЫТИЕ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ОЗЕЛЕНЕНИЕ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ОЗЕЛЕНЕНИЕ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "МАФ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "МАФ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ОБОРУДОВАНИЕ ПО ТХЗ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ОБОРУДОВАНИЕ ПО ТХЗ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "ПОСТАВЛЕНО НА ОБЪЕКТ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "ПОСТАВЛЕНО НА ОБЪЕКТ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "СМОНТИРОВАНО (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "СМОНТИРОВАНО (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "МЕБЕЛЬ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "МЕБЕЛЬ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "КЛИНИНГ (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "КЛИНИНГ (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "Дорожное покрытие (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "Дорожное покрытие (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "Искусственные сооружения (ИССО) (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "Искусственные сооружения (ИССО) (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "Наружные инженерные сети (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "Наружные инженерные сети (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "Средства организации дорожного движения (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "Средства организации дорожного движения (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "Дорожная разметка (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "Дорожная разметка (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "Благоустройство прилегающей территории (фактическая дата начала)": "ДД.ММ.ГГГГ",
    "Благоустройство прилегающей территории (фактическая дата завершения)": "ДД.ММ.ГГГГ",
    "Дата включения в АИП": "ДД.ММ.ГГГГ",
    "Плановый ввод по директивному графику": "ДД.ММ.ГГГГ",
    "Дата совещания - изменение даты ввода по директивному графику": "ДД.ММ.ГГГГ",
    "Плановый ввод по договору": "ДД.ММ.ГГГГ",
    "Прогнозируемый срок ввода": "ДД.ММ.ГГГГ",
    "Получение ГПЗУ план": "ДД.ММ.ГГГГ",
    "Получение ТУ от ресурсоснабжающих организаций (план)": "ДД.ММ.ГГГГ",
    "Разработка и согласование АГР (план)": "ДД.ММ.ГГГГ",
    "Разработка и получение положительного заключения экспертизы ПСД (план)": "ДД.ММ.ГГГГ",
    "Получение РНС (план)": "ДД.ММ.ГГГГ",
    "Производство СМР (план)": "ДД.ММ.ГГГГ",
    "Технологическое присоединение (план)": "ДД.ММ.ГГГГ",
    "Получение ЗОС (план)": "ДД.ММ.ГГГГ",
    "Получение РВ (план)": "ДД.ММ.ГГГГ",
    "СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "СТРОИТЕЛЬНАЯ ГОТОВНОСТЬ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "КОНСТРУКТИВ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "КОНСТРУКТИВ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "КОЛ-ВО КОРПУСОВ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "КОЛ-ВО КОРПУСОВ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "КОЛ-ВО ЭТАЖЕЙ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "КОЛ-ВО ЭТАЖЕЙ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "СТЕНЫ И ПЕРЕГОРОДКИ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "СТЕНЫ И ПЕРЕГОРОДКИ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ФАСАД (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ФАСАД (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "УТЕПЛИТЕЛЬ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "УТЕПЛИТЕЛЬ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ФАСАДНАЯ СИСТЕМА (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ФАСАДНАЯ СИСТЕМА (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ВНУТРЕННЯЯ ОТДЕЛКА (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ВНУТРЕННЯЯ ОТДЕЛКА (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ЧЕРНОВАЯ ОТДЕЛКА (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ЧЕРНОВАЯ ОТДЕЛКА (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ЧИСТОВАЯ ОТДЕЛКА (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ЧИСТОВАЯ ОТДЕЛКА (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ВНУТРЕННИЕ СЕТИ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ВНУТРЕННИЕ СЕТИ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ (ВЕРТИКАЛЬ) (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ВОДОСНАБЖЕНИЕ И ОТОПЛЕНИЕ ПОЭТАЖНО (ГОРИЗОНТ) (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ВЕНТИЛЯЦИЯ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ВЕНТИЛЯЦИЯ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ЭЛЕКТРОСНАБЖЕНИЕ И СКС (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ЭЛЕКТРОСНАБЖЕНИЕ И СКС (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "НАРУЖНЫЕ СЕТИ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "НАРУЖНЫЕ СЕТИ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "БЛАГОУСТРОЙСТВО (плановая дата начала)": "ДД.ММ.ГГГГ",
    "БЛАГОУСТРОЙСТВО (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ТВЕРДОЕ ПОКРЫТИЕ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ТВЕРДОЕ ПОКРЫТИЕ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ОЗЕЛЕНЕНИЕ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ОЗЕЛЕНЕНИЕ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "МАФ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "МАФ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ОБОРУДОВАНИЕ ПО ТХЗ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ОБОРУДОВАНИЕ ПО ТХЗ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "ПОСТАВЛЕНО НА ОБЪЕКТ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "ПОСТАВЛЕНО НА ОБЪЕКТ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "СМОНТИРОВАНО (плановая дата начала)": "ДД.ММ.ГГГГ",
    "СМОНТИРОВАНО (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "МЕБЕЛЬ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "МЕБЕЛЬ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "КЛИНИНГ (плановая дата начала)": "ДД.ММ.ГГГГ",
    "КЛИНИНГ (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "Дорожное покрытие (плановая дата начала)": "ДД.ММ.ГГГГ",
    "Дорожное покрытие (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "Искусственные сооружения (ИССО) (плановая дата начала)": "ДД.ММ.ГГГГ",
    "Искусственные сооружения (ИССО) (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "Наружные инженерные сети (плановая дата начала)": "ДД.ММ.ГГГГ",
    "Наружные инженерные сети (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "Средства организации дорожного движения (плановая дата начала)": "ДД.ММ.ГГГГ",
    "Средства организации дорожного движения (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "Дорожная разметка (плановая дата начала)": "ДД.ММ.ГГГГ",
    "Дорожная разметка (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "Благоустройство прилегающей территории (плановая дата начала)": "ДД.ММ.ГГГГ",
    "Благоустройство прилегающей территории (плановая дата завершения)": "ДД.ММ.ГГГГ",
    "Наружные инженерные сети  (плановая дата завершения)": "ДД.ММ.ГГГГ"
}


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
        'ОЦЕНКА СТРОИТЕЛЬНОГО ГОРОДКА', 'ПАСПОРТ ОБЪЕКТА (наличие)',
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
        'Благоустройство прилегающей территории (факт)',
        
    ]

    # Обработка каждого листа
    for sheet_name in sheets_to_process:
        if sheet_name in xls.sheet_names:
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

                # Добавление колонок из листа "5 - ОИВ факт"
                df_fact = pd.read_excel(xls, sheet_name="5 - ОИВ факт", header=1, skiprows=[2], decimal=',')
                copy_columns = ["Дата контрактации", "Сумма по контракту, млрд руб"]

                # Убедитесь, что колонки существуют в листе "5 - ОИВ факт"
                if all(column in df_fact.columns for column in copy_columns):
                    # Добавление колонок в DataFrame "4 - ОИВ план"
                    for column in copy_columns:
                        df[column] = df_fact[column]

                    # Перемещение колонок после "Этажность"
                    cols = list(df.columns)
                    idx_etazhnost = cols.index("Этажность")
                    for column in copy_columns:
                        cols.insert(idx_etazhnost + 1, cols.pop(cols.index(column)))
                    df = df[cols]

            # Удаление колонок с комментариями, титулами и годами титулов
            df = df.loc[:, ~df.columns.str.contains('комментарий|титул|год титула', case=False)]

            # Удаление указанных колонок в листах "1 - СМГ ежедневный", "4 - ОИВ план" и "5 - ОИВ факт"
            if sheet_name in ["1 - СМГ ежедневный", "4 - ОИВ план", "5 - ОИВ факт"]:
                columns_to_drop = [
                    "КОЛ-ВО КОРПУСОВ (плановая дата начала)",
                    "КОЛ-ВО КОРПУСОВ (плановая дата завершения)",
                    "КОЛ-ВО ЭТАЖЕЙ (плановая дата начала)",
                    "КОЛ-ВО ЭТАЖЕЙ (плановая дата завершения)",
                    "КОЛ-ВО КОРПУСОВ (план)",
                    "КОЛ-ВО КОРПУСОВ (факт)",
                    "КОЛ-ВО ЭТАЖЕЙ (план)",
                    "КОЛ-ВО ЭТАЖЕЙ (факт)",
                    "КОЛ-ВО КОРПУСОВ (фактическая дата начала)",
                    "КОЛ-ВО КОРПУСОВ (фактическая дата завершения)",
                    "КОЛ-ВО КОРПУСОВ (% выполнения, план)",
                    "КОЛ-ВО КОРПУСОВ (% выполнения, факт)",
                    "КОЛ-ВО ЭТАЖЕЙ (фактическая дата начала)",
                    "КОЛ-ВО ЭТАЖЕЙ (фактическая дата завершения)",
                    "КОЛ-ВО ЭТАЖЕЙ (% выполнения, план)",
                    "КОЛ-ВО ЭТАЖЕЙ (% выполнения, факт)"]
                df.drop(columns=[col for col in columns_to_drop if col in df.columns], inplace=True)


            # Удаление указанных колонок для листа "1 - СМГ ежедневный"
            if sheet_name == "1 - СМГ ежедневный":
                columns_to_drop = [
                    "ОЦЕНКА СТРОИТЕЛЬНОГО ГОРОДКА",
                    "ПАСПОРТ ОБЪЕКТА (наличие)",
                    "ПАСПОРТ ОБЪЕКТА (оценка)",
                    "ПЕРИМЕТРАЛЬНОЕ ОГРАЖДЕНИЕ (наличие)",
                    "ПЕРИМЕТРАЛЬНОЕ ОГРАЖДЕНИЕ (оценка)",
                    "СИГНАЛЬНОЕ ОГРАЖДЕНИЕ \"ГИРЛЯНДА\" (наличие)",
                    "СИГНАЛЬНОЕ ОГРАЖДЕНИЕ \"ГИРЛЯНДА\" (оценка)",
                    "ПОСТ ОХРАНЫ (наличие)",
                    "ПОСТ ОХРАНЫ (оценка)",
                    "ПУНКТ МОЙКИ КОЛЕС (наличие)",
                    "ПУНКТ МОЙКИ КОЛЕС (оценка)",
                    "ПОКРЫТИЕ СТРОИТЕЛЬНОГО ГОРОДКА (наличие)",
                    "ПОКРЫТИЕ СТРОИТЕЛЬНОГО ГОРОДКА (оценка)",
                    "ПОДЪЕЗДНЫХ ПУТЕЙ (наличие)",
                    "ПОДЪЕЗДНЫХ ПУТЕЙ (оценка)",
                    "НАЛИЧИЕ И СОСТОЯНИЕ  САНИТАРНО-БЫТОВЫХ, ПРОИЗВОДСТВЕННЫХ И АДМИНИСТРАТИВНЫХ ПОМЕЩЕНИЙ И СООРУЖЕНИЙ (наличие)",
                    "НАЛИЧИЕ И СОСТОЯНИЕ  САНИТАРНО-БЫТОВЫХ, ПРОИЗВОДСТВЕННЫХ И АДМИНИСТРАТИВНЫХ ПОМЕЩЕНИЙ И СООРУЖЕНИЙ (оценка)",
                    "ЧИСТОТА И ПОРЯДОК НА ТЕРРИТОРИИ СТРОИТЕЛЬНОГО ГОРОДКА \n(наличие)",
                    "ЧИСТОТА И ПОРЯДОК НА ТЕРРИТОРИИ СТРОИТЕЛЬНОГО ГОРОДКА\n (оценка)"
                ]
                # Удаляем лишние пробелы в названиях колонок
                df.columns = df.columns.str.strip()
                
                # Удаляем только те колонки, которые существуют в DataFrame
                columns_to_drop = [col.strip() for col in columns_to_drop if col.strip() in df.columns]
                df.drop(columns=columns_to_drop, inplace=True)

            # Преобразование значений в колонке "АИП (да/нет)" (если она существует)
            if 'АИП (да/нет)' in df.columns:
                df['АИП (да/нет)'] = df['АИП (да/нет)'].astype(str).str.lower().replace({
                    'true': 'да',
                    'false': 'нет',
                    'да': 'да',
                    'нет': 'нет',
                    'nan': None
                })

            # Удаление строк, где все указанные колонки пустые
            df.dropna(subset='УИН', how='all', inplace=True)
            df.drop(df[df['УИН'] == 0].index, inplace=True)

            # Обработка листа "5 - ОИВ факт"
            if sheet_name == "5 - ОИВ факт" or sheet_name == '1 - СМГ ежедневный':
                df['ИНН Генподрядчика'] = df['ИНН Генподрядчика'].fillna(0).astype('Int64').replace(0, None)
                df = df.replace(r'^\s*$', None, regex=True)
                # Список колонок, которые нужно переименовать
                columns_to_rename = {
                    "Получение ГПЗУ (факт)": "ЭТАПРЕАЛИЗАЦИИ ГПЗУ (факт)",
                    "Получение ГПЗУ факт": "ЭТАПРЕАЛИЗАЦИИ ГПЗУ (факт)",
                    "Получение ТУ от ресурсоснабжающих организаций (факт)": "ЭТАПРЕАЛИЗАЦИИ ТУ от РСО (факт)",
                    "Разработка и согласование АГР (факт)": "ЭТАПРЕАЛИЗАЦИИ АГР (факт)",
                    "Разработка и получение положительного заключения экспертизы ПСД (факт)": "ЭТАПРЕАЛИЗАЦИИ Экспертиза ПСД (факт)",
                    "Получение РНС (факт)": "ЭТАПРЕАЛИЗАЦИИ РНС (факт)",
                    "Начато производство СМР": "ЭТАПРЕАЛИЗАЦИИ СМР (факт)",
                    "Технологическое присоединение (факт)": "ЭТАПРЕАЛИЗАЦИИ Техприс (факт)",
                    "Получение ЗОС (факт)": "ЭТАПРЕАЛИЗАЦИИ ЗОС (факт)",
                    "Получение РВ (факт)": "ЭТАПРЕАЛИЗАЦИИ РВ (факт)"
                }
                
                # Переименование колонок
                df.rename(columns=columns_to_rename, inplace=True)

                if sheet_name == "5 - ОИВ факт":
                    columns_to_drop = ["Дата контрактации", "Сумма по контракту, млрд руб"]
                    df.drop(columns=columns_to_drop, inplace=True)
            
            # Обработка листа "4 - ОИВ план"
            if sheet_name == "4 - ОИВ план":
                df['ИНН Генподрядчика'] = df['ИНН Генподрядчика'].fillna(0).astype('Int64').replace(0, None)
                df = df.replace(r'^\s*$', None, regex=True)
                df["Год ввода"] = df["Год ввода (АИП)"]
                # Список колонок, которые нужно переименовать
                columns_to_rename = {
                    "Получение ГПЗУ план": "ЭТАПРЕАЛИЗАЦИИ ГПЗУ (план)",
                    "Получение ТУ от ресурсоснабжающих организаций (план)": "ЭТАПРЕАЛИЗАЦИИ ТУ от РСО (план)",
                    "Разработка и согласование АГР (план)": "ЭТАПРЕАЛИЗАЦИИ АГР (план)",
                    "Разработка и получение положительного заключения экспертизы ПСД (план)": "ЭТАПРЕАЛИЗАЦИИ Экспертиза ПСД (план)",
                    "Получение РНС (план)": "ЭТАПРЕАЛИЗАЦИИ РНС (план)",
                    "Производство СМР (план)": "ЭТАПРЕАЛИЗАЦИИ СМР (план)",
                    "Технологическое присоединение (план)": "ЭТАПРЕАЛИЗАЦИИ Техприс (план)",
                    "Получение ЗОС (план)": "ЭТАПРЕАЛИЗАЦИИ ЗОС (план)",
                    "Получение РВ (план)": "ЭТАПРЕАЛИЗАЦИИ РВ (план)"
                }
                
                # Переименование колонок
                df.rename(columns=columns_to_rename, inplace=True)

            # Удаление указанных колонок на листах "5 - ОИВ факт" и "1 - СМГ ежедневный"
            if sheet_name in ["5 - ОИВ факт", "1 - СМГ ежедневный"]:
                columns_to_exclude = [
                    "Мастер код ФР", "Мастер код ДС", "link", "ФНО", "Наименование",
                    "Наименование ЖК (коммерческое наименование)", "Краткое наименование",
                    "ФНО для аналитики", "Округ", "Район", "Источник финансирования",
                    "Статус ОКС", "Адрес объекта", "Признак реновации (да/нет)",
                    "ИНН Застройщика", "Застройщик", "ИНН Заказчика", "Заказчик",
                    "ИНН Генподрядчика", "Генподрядчик", "Общая площадь, м2",
                    "Протяженность", "Этажность",  "АИП (да/нет)", "Дата включения в АИП",
                    "Сумма по АИП, млрд руб", "Аванс, млрд руб"
                ]
                # Удаляем только те колонки, которые существуют в DataFrame
                columns_to_exclude = [col for col in columns_to_exclude if col in df.columns]
                df.drop(columns=columns_to_exclude, inplace=True)

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
                    elif isinstance(value, str) and value.strip().lower() == "данных не будет":
                        continue
                    elif key in column_types:
                        value = validate_data_type(value, column_types[key])

                    # Если значение равно "не требуется" или null, удаляем ключ
                    if "корректировка" in key.lower() and (value == "не требуется" or value is None):
                        del new_record[key]


                    # Добавление "СТРЭТАП" к нужным ключам
                    if key.strip() in keys_to_modify:
                        new_record[f'СТРЭТАП {key.strip()}'] = value
                    else:
                        # Добавление "ЭТАПРЕАЛИЗАЦИИ" к нужным ключам и удаление оригинальных ключей
                        if any(phrase in key.strip() for phrase in [
                            'Получение ГПЗУ', 'Получение ТУ от РСО',
                            'Разработка и согласование АГР', 'Разработка и экспертиза ПСД',
                            'Получение РНС', 'Начато производство СМР', 'Технологическое присоединение',
                            'Получение ЗОС', 'Получение РВ', 'Производство СМР'
                        ]):
                            new_record[f'ЭТАПРЕАЛИЗАЦИИ {key.strip()}'] = value
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
convert_excel_to_json('E://Загрузки//Telegram Desktop//Текущая обработка/ДГС_для_ДБ_2024_2027_год_тест_пролив_СМГ_с_наименованиями.xlsx')
