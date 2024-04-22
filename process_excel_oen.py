import logging
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import MergedCell
import re
def process_excel_file_oen(file_path, installation_name, control_date):
    # Открываем файл Excel
    workbook = load_workbook(file_path)

    # Доступ к листу "УЗТ (коркарта)"
    sheet = workbook['УЗТ (коркарта)']

    # Проверка ячейки L5
    cell_L5 = sheet['L5']
    if cell_L5.value:
        # Добавляем новый столбец перед B
        sheet.insert_cols(2)

        # Добавляем формулу в столбец B
        for row in range(2, sheet.max_row + 1):
            # Получаем ячейку в столбце 2
            cell = sheet.cell(row=row, column=2)
            # Проверяем, является ли ячейка объединенной
            if isinstance(cell, MergedCell):
                # Найдем верхнюю левую ячейку объединенного диапазона
                top_left_cell = sheet.merged_cells.ranges.get(cell.coordinate)
                if top_left_cell:
                    # Устанавливаем формулу в верхней левой ячейке
                    top_left_cell.value = f'=CONCATENATE(B{row}; "."; C{row}; ".")'
            else:
                # Устанавливаем значение, если ячейка не объединена
                cell.value = f'=CONCATENATE(B{row}; "."; C{row}; ".")'

    # Вставляем введенное значение в ячейку F7
    sheet['F7'].value = installation_name

    # Проверка значения в ячейке F7 на соответствие формату ХХ.ХХ.ХХХХ
    cell_value = sheet['F7'].value
    pattern = re.compile(r'^\d{2}\.\d{2}\.\d{4}$')
    if not pattern.match(cell_value):
        # Если значение не соответствует формату, удаляем "г." из даты
        sheet['F7'].value = cell_value.replace("г.", "")

    # Сохраняем файл
    workbook.save(file_path)

    return f"Обработано успешно! Файл сохранен: {file_path}"

def set_value_if_not_merged(sheet: Worksheet, row: int, col: int, value):
    cell = sheet.cell(row=row, column=col)
    merged_ranges = sheet.merged_cells.ranges

    # Проверяем, находится ли ячейка в объединенном диапазоне
    if any(range.min_row <= cell.row <= range.max_row and
           range.min_col <= cell.column <= range.max_col
           for range in merged_ranges):
        # Находим верхний левый угол объединенного диапазона
        for range in merged_ranges:
            if range.min_row <= cell.row <= range.max_row and \
               range.min_col <= cell.column <= range.max_col:
                # Устанавливаем значение в верхней левой ячейке
                top_left_cell = sheet.cell(row=range.min_row, column=range.min_col)
                top_left_cell.value = value
                break
    else:
        # Устанавливаем значение, если ячейка не объединена
        cell.value = value
