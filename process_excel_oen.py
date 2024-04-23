import logging
import os
import re
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from datetime import datetime

def is_date_time_format(cell_value):
    """
    Определяет, является ли значение ячейки строкой в формате даты 'YYYY-MM-DD HH:MM:SS'.

    Параметры:
        cell_value: Значение ячейки, которое необходимо проверить.

    Возвращает:
        True, если строка соответствует формату 'YYYY-MM-DD HH:MM:SS', иначе False.
    """
    date_time_format = '%Y-%m-%d %H:%M:%S'

    try:
        # Попытка разобрать строку как дату и время в указанном формате
        datetime.strptime(cell_value, date_time_format)
        return True
    except ValueError:
        # Если возникло исключение, значит строка не соответствует формату
        return False

def replace_formulas_with_values(sheet):
    """
    Функция заменяет формулы в ячейках их текущими значениями.
    """
    for row in sheet.iter_rows():

        for cell in row:
            cell_address = cell.coordinate

            formula = cell.value
            # value = cell.calculate(formula)
            if isinstance(formula, str) and formula.startswith('='):
                # Заменяем формулу текущим значением ячейки

                print('Формула', formula)
                print('Тип адреса', type(cell_address))

def process_excel_file_oen(file_path, installation_name, control_date):
    # Открываем файл Excel
    workbook = load_workbook(file_path)
    logging.info(f'----- НАЧАЛО ОБРАБОТКИ ФАЙЛА: {os.path.basename(file_path)} -----')
    logging.info(f'Загружен файл Excel: {file_path}')

    # Доступ к листу "УЗТ (коркарта)"
    sheet = workbook['УЗТ (коркарта)']
    logging.info('Доступ к листу "УЗТ (коркарта)"')

    start_row = 19
    end_row = 256
    col_index = 2

     # Вставляем новый столбец в диапазоне строк
    for row in range(end_row, start_row - 1, -1):
        # Сдвигаем значения в ячейках в строке вправо
        for col in range(sheet.max_column, col_index - 1, -1):
            # Копируем значение и стиль ячейки на одну колонку вправо
            source_cell = sheet.cell(row=row, column=col)
            target_cell = sheet.cell(row=row, column=col + 1)
            target_cell.value = source_cell.value
            target_cell._style = source_cell._style

        # Удаляем значение и стиль ячейки в вставляемом столбце
        source_cell = sheet.cell(row=row, column=col_index)
        source_cell.value = None
        source_cell._style = sheet.cell(row=row, column=col_index + 1)._style

    # Сохраняем изменения
    workbook.save(file_path + '_new.xlsx')

    return
    # replace_formulas_with_values(sheet)
    # logging.info('Формулы заменены значениями')
    # return
    # Получаем ячейку L5
    cell_L5 = sheet['L5']

    # Получаем значение ячейки L5
    value_L5 = cell_L5.value
    logging.info(f'Значение ячейки L5: {value_L5}')

    if isinstance(value_L5, str) and value_L5.startswith('='):
        logging.info(f'Тип ячейки L5 {type(value_L5)}')
        replace_formulas_with_values(sheet)
        return
        sheet.insert_cols(2)
        logging.info(f'ДОБАВЛЯЕМ НОВЫЙ СТОЛБЕЦ ПЕРЕД B!')

    else:
        # Если L5 не содержит формулы или значение равно None
        # Проверяем, скрыт ли столбец A
        if sheet.column_dimensions['A'].hidden:
            logging.info(f'Cтолбец A скрыт!')
            # Показываем столбец A
            sheet.column_dimensions['A'].hidden = False
            logging.info(f'Показываем столбец A')

    # Сохраняем файл
    workbook.save(file_path)
    logging.info(f'Файл {file_path} успешно сохранен')
    return

    # if type(value_L5) == 'string':
    pattern = re.compile(r'^\d{2}\.\d{2}\.\d{4}$')
    logging.info(f'Проверка значения в ячейке F5: {value_L5}')
    if not pattern.match(value_L5):
            logging.warning(f'Значение в ячейке F5 не соответствует формату DD.MM.YYYY. Значение: {value_L5}')
            sheet['F7'].value = cell_value.replace("г.", "")

            new_value = sheet['F7'].value
            logging.info(f'Удалено "г." из ячейки F7, новое значение: {new_value}')
            return
                # Добавляем новый столбец перед B
            sheet.insert_cols(2)
            logging.info('Добавлен новый столбец перед B')

        # Добавляем формулу в столбец B
            for row in range(2, sheet.max_row + 1):
                cell = sheet.cell(row=row, column=2)
                if isinstance(cell, MergedCell):
                    top_left_cell = sheet.merged_cells.ranges.get(cell.coordinate)
                    if top_left_cell:
                        top_left_cell.value = f'=CONCATENATE(B{row}; "."; C{row}; ".")'
                else:
                    cell.value = f'=CONCATENATE(B{row}; "."; C{row}; ".")'
                logging.info(f'Установлена формула в B{row}')

    return
    # Вставляем введенное значение в ячейку F7
    sheet['F7'].value = installation_name
    logging.info(f'Установлено значение "{installation_name}" в ячейку F7')

    # Проверка значения в ячейке F7 на соответствие формату DD.MM.YYYY
    cell_value = sheet['F7'].value


    # Сохраняем файл
    workbook.save(file_path)
    logging.info(f'Файл {file_path} успешно сохранен')

    return f'Обработано успешно! Файл сохранен: {file_path}'