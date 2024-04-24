import logging
import os
import re
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from datetime import datetime
from helpers import find_table_start, find_table_end, insert_column_in_range
from insert_column_xlwings import insert_column_after

def process_excel_file_oen(file_path, installation_name):
    # Открываем файл Excel
    workbook = load_workbook(file_path)
    logging.info(f'----- НАЧАЛО ОБРАБОТКИ ФАЙЛА: {os.path.basename(file_path)} -----')
    logging.info(f'Загружен файл Excel: {file_path}')

    # Доступ к листу "УЗТ (коркарта)"
    sheet = workbook['УЗТ (коркарта)']
    logging.info('Доступ к листу "УЗТ (коркарта)"')

    start_row = find_table_start(sheet)
    logging.info(f'Первая строка таблицы Результаты контроля: {start_row}')
    end_row = find_table_end(sheet)
    logging.info(f'Последняя строка таблицы Результаты контроля: {end_row}')

    col_index = 2


    # Вставляем новый столбец в диапазоне строк
    # insert_column_in_range(sheet, col_index, start_row, end_row)
    # row = 15
    col_num = 'B'

    insert_column_after(file_path, start_row, end_row, col_num)
    # Сохраняем изменения
    # workbook.save(file_path + "_new.xlsx")
    return 
    
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