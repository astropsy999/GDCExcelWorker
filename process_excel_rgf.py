import logging
import os
from openpyxl import load_workbook

def process_excel_file_rgf(file_path, installation_name, control_date):
    # Загрузите файл
    workbook = load_workbook(file_path)

    file_name = os.path.basename(file_path)

    # Лист 'Характеристики'
    characteristics_sheet = workbook['Характеристики']

    # Установка значения в C2
    characteristics_sheet['C2'].value = installation_name
    logging.info(f'----- НАЧАЛО ОБРАБОТКИ ФАЙЛА: {file_name} -----"')
    logging.info(f'Установлено значение {installation_name} в ячейку C2 на листе "Характеристики"')

    # Проверка ячейки A1 и установка значения в C4
    a1_value = characteristics_sheet['A1'].value
    c4_value = None  # Значение по умолчанию для C4

    if a1_value:
        c4_value = a1_value
        logging.info(f'Значение {a1_value} из ячейки A1 скопировано в ячейку C4 на листе "Характеристики"')
    else:
        logging.info(f'Ячейка A1 пуста, ячейка C4 будет установлена в None')

    characteristics_sheet['C4'].value = c4_value

    # Лист 'Диагностическая карта'
    diagnostics_sheet = workbook['Диагностическая карта']

    # Установка даты контроля в L5
    diagnostics_sheet['L5'].value = control_date
    logging.info(f'Установлено значение {control_date} в ячейку L5 на листе "Диагностическая карта"')

    # Счетчик измененных строк в столбце A
    changed_rows_count = 0

    # Добавление формулы в столбец A
    # Начинаем с первой строки данных (возможно, первая строка - это заголовок)
    for row in range(8, diagnostics_sheet.max_row + 1):
        # Формула для конкатенации B и C с точкой
        formula = f'=CONCATENATE(B{row}, ".", C{row}, ".")'

        # Проверка, что ячейки B{row} и C{row} не пусты перед установкой формулы в A{row}
        if diagnostics_sheet[f'B{row}'].value is not None and diagnostics_sheet[f'C{row}'].value is not None:
            diagnostics_sheet[f'A{row}'].value = formula
            changed_rows_count += 1  # Увеличиваем счетчик для каждой измененной строки

    # Логирование количества измененных строк
    logging.info(f'Количество строк, где была установлена формула в столбец A: {changed_rows_count}')

    # Сохранение файла
    workbook.save(file_path)

    # Логирование успешной обработки файла
    logging.info(f'ФАЙЛ {file_name} ОБРАБОТАН УСПЕШНО!.')

    # Загрузите файл
    workbook = load_workbook(file_path)

    # Лист 'Характеристики'
    characteristics_sheet = workbook['Характеристики']

    # Установка значения в C2
    characteristics_sheet['C2'].value = installation_name

    # Проверка ячейки A1 и установка значения в C4
    if characteristics_sheet['A1'].value:
        characteristics_sheet['C4'].value = characteristics_sheet['A1'].value
    else:
        characteristics_sheet['C4'].value = None

    # Лист 'Диагностическая карта'
    diagnostics_sheet = workbook['Диагностическая карта']

    # Установка даты контроля в L5
    diagnostics_sheet['L5'].value = control_date

    # Сохранение файла
    workbook.save(file_path)

    # Логирование успешной обработки файла
    # logging.info(f'Файл {file_path} обработан успешно.')