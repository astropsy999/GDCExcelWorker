import logging
import os
import re
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
import xlwings as xw

def process_excel_file_oen(file_path, installation_name, control_date):
    # Открываем файл Excel
    # workbook = load_workbook(file_path)
    # logging.info(f'----- НАЧАЛО ОБРАБОТКИ ФАЙЛА: {os.path.basename(file_path)} -----')
    # logging.info(f'Загружен файл Excel: {file_path}')
    
    # # Доступ к листу "УЗТ (коркарта)"
    # sheet = workbook['УЗТ (коркарта)']
    # logging.info('Доступ к листу "УЗТ (коркарта)"')

    # # Получаем ячейку L5
    # cell_L5 = sheet['L5']

    # # Получаем значение ячейки L5
    # value_L5 = cell_L5.value
    # logging.info(f'Значение ячейки L5: {value_L5}')

    # # Если значение в ячейке L5 является формулой, извлекаем адрес ячейки, на которую она ссылается
    # if isinstance(value_L5, str) and value_L5.startswith('='):
    #     sheet.insert_cols(2)
    #     logging.info(f'ДОБАВЛЯЕМ НОВЫЙ СТОЛБЕЦ ПЕРЕД B!')

    # else:
    #     # Если L5 не содержит формулы или значение равно None
    #     # Проверяем, скрыт ли столбец A
    #     if sheet.column_dimensions['A'].hidden:
    #         logging.info(f'Cтолбец A скрыт!')
    #         # Показываем столбец A
    #         sheet.column_dimensions['A'].hidden = False
    #         logging.info(f'Показываем столбец A')
    
    # # Сохраняем файл
    # workbook.save(file_path)
    # logging.info(f'Файл {file_path} успешно сохранен')        
       
      # Открываем книгу Excel с помощью xlwings
    app = xw.App(visible=False)  # Запускаем приложение Excel в фоновом режиме
    workbook = app.books.open(file_path)
    logging.info(f'----- НАЧАЛО ОБРАБОТКИ ФАЙЛА: {os.path.basename(file_path)} -----')
    logging.info(f'Загружен файл Excel: {file_path}')

    # Доступ к листу "УЗТ (коркарта)"
    sheet = workbook.sheets['УЗТ (коркарта)']
    logging.info('Доступ к листу "УЗТ (коркарта)"')

    # Получаем ячейку L5
    cell_L5 = sheet.range('L5')

    # Получаем значение ячейки L5
    value_L5 = cell_L5.value
    logging.info(f'Значение ячейки L5: {value_L5}')

    # Если значение в ячейке L5 является формулой, вставляем новый столбец перед B
    if isinstance(value_L5, str) and value_L5.startswith('='):
        sheet.api.Columns("B").Insert()
        logging.info(f'ДОБАВЛЯЕМ НОВЫЙ СТОЛБЕЦ ПЕРЕД B!')
    else:
        # Если L5 не содержит формулы или значение равно None
        # Проверяем, скрыт ли столбец A
        if sheet.range('A1').api.EntireColumn.Hidden:
            logging.info(f'Cтолбец A скрыт!')
            # Показываем столбец A
            sheet.range('A1').api.EntireColumn.Hidden = False
            logging.info(f'Показываем столбец A')

    # Обновляем формулы на листе, чтобы учесть изменения
    sheet.api.Calculate()
    logging.info('Обновили формулы на листе после изменений')

    # Сохраняем файл
    workbook.save(file_path)
    logging.info(f'Файл {file_path} успешно сохранен')

    # Закрываем книгу Excel
    workbook.close()
    app.quit()     
            
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