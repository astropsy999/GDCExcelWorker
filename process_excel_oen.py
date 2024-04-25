import os
import re
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from datetime import datetime
from helpers import find_table_start, find_table_end
from insert_column_xlwings import insert_column_after
from utils import copy_values_and_insert_formula

def process_excel_file_oen(file_path, installation_name):
    # Открываем файл Excel
    workbook = load_workbook(file_path)
    print(f'----- НАЧАЛО ОБРАБОТКИ ФАЙЛА: {os.path.basename(file_path)} -----')
    print(f'Загружен файл Excel: {file_path}')

    # Доступ к листу "УЗТ (коркарта)"
    sheet = workbook['УЗТ (коркарта)']
    print('Доступ к листу "УЗТ (коркарта)"')
    
    cell_L5 = sheet['L5']
    # Получаем значение ячейки L5
    value_L5 = cell_L5.value
    
    if value_L5 is not None:
        
        start_row = find_table_start(sheet)
        end_row = find_table_end(sheet)
        
        # Вставляем новый столбец в диапазоне строк
        col_num = 'B'

        insert_column_after(file_path, start_row, end_row, col_num, installation_name)
    
    else:
        date_value = sheet['M5'].value
        # Проверка значения даты на соответствие формату DD.MM.YYYY
        pattern = re.compile(r'^\d{2}\.\d{2}\.\d{4}$')
        print(f'Проверка значения даты: {date_value}')
        if not pattern.match(date_value):
            print(f'Дата не соответствует формату DD.MM.YYYY!')
            sheet['M5'].value = date_value.replace("г.", "")
            print(f'Формат даты изменен!')

        # Если L5 значение равно None
        # Проверяем, скрыт ли столбец A
        if sheet.column_dimensions['A'].hidden:
            print('Cтолбец A скрыт!')
            # Показываем столбец A
            sheet.column_dimensions['A'].hidden = False
            print(f'Показываем столбец A')
            
            sheet['F7'].value = installation_name
            print(f'Изменено::Технологическая установка (участок): {installation_name}')
            
            start_row = find_table_start(sheet)
            end_row = find_table_end(sheet)
            
            copy_values_and_insert_formula(sheet, start_row, end_row)
        
            workbook.save(file_path +'_oen.xlsx')
            print(f'Файл успешно сохранен {file_path}')

    return