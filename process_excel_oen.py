import os
import re
from openpyxl import load_workbook
from helpers import find_table_start, find_table_end
from insert_column_xlwings import insert_column_after
from utils import copy_values_and_insert_formula
from utils import bcolors

def process_excel_file_oen(file_path, installation_name):
    try:
        # Открываем файл Excel
        workbook = load_workbook(file_path)
        print(f'\n----- НАЧАЛО ОБРАБОТКИ ФАЙЛА: {os.path.basename(file_path)} -----')
        print(f'Загружен файл Excel: {file_path}')

        # Попытка получить доступ к листу "УЗТ (коркарта)"
        try:
            sheet = workbook['УЗТ (коркарта)']
        except KeyError:
            raise KeyError(f'{bcolors.FAIL} Лист "УЗТ (коркарта)" отсутствует в файле {file_path}.{bcolors.ENDC}')
        finally:
            workbook.save(file_path)
            workbook.close()
            
        # Доступ к ячейкам и обработка значений
        try:
            cell_L5 = sheet['L5']
            value_L5 = cell_L5.value
        except KeyError:
            print(f'{bcolors.FAIL}Ошибка: Не удалось получить доступ к ячейке L5 в файле {file_path}.{bcolors.ENDC}')
            return
        finally:
            workbook.save(file_path)
            workbook.close()

        if value_L5 is not None:
            start_row = find_table_start(sheet)
            end_row = find_table_end(sheet)
            
            if(start_row is None or end_row is None):
                print(f'{bcolors.FAIL}Ошибка: Не удалось обнаружить начало или конец интервала для вставки данных {file_path}.{bcolors.ENDC}')
                raise KeyError
            
            # Вставляем новый столбец в диапазоне строк
            col_num = 'B'
            workbook.save(file_path)
            workbook.close()
            insert_column_after(file_path, start_row, end_row, col_num, installation_name)
            print(f'Файл успешно сохранен: {file_path}')
        
        else:
            # Проверяем значение даты
            date_value = sheet['M5'].value
            pattern = re.compile(r'^\d{2}\.\d{2}\.\d{4}$')
            print(f'Проверка значения даты: {date_value}')
            
            if not pattern.match(date_value):
                print(f'{bcolors.WARNING}Дата не соответствует формату DD.MM.YYYY!{bcolors.ENDC}')
                sheet['M5'].value = date_value.replace("г.", "")
                print(f'Формат даты изменен!')

            # Если L5 значение равно None, проверяем столбец A
            if sheet.column_dimensions['A'].hidden:
                print('Столбец A скрыт!')
                sheet.column_dimensions['A'].hidden = False
                print(f'Показываем столбец A')

                sheet['F7'].value = installation_name
                print(f'Изменено::Технологическая установка (участок): {installation_name}')
                
                start_row = find_table_start(sheet)
                end_row = find_table_end(sheet)
                
                copy_values_and_insert_formula(sheet, start_row, end_row)
                
            else:
                sheet['F7'].value = installation_name
                print(f'Изменено::Технологическая установка (участок): {installation_name}')
                
                start_row = find_table_start(sheet)
                end_row = find_table_end(sheet)
                
                copy_values_and_insert_formula(sheet, start_row, end_row)
            
            # Сохраняем файл с обработанными данными
            workbook.save(file_path + '_oen.xlsx')
            print(f'Файл успешно сохранен: {file_path}')

    except KeyError as e:
         raise KeyError(f'{bcolors.FAIL}Ошибка: {e}{bcolors.ENDC}')
        
    finally:
        # Закрываем файл Excel в конце обработки
        workbook.save(file_path)
        workbook.close()

    return
