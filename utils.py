def is_row_empty(row_index):
    """
    Функция проверяет, пустая ли строка.

    :param next_row: Кортеж ячеек следующей строки после "Результаты контроля:".
    :return: True, если строка пустая (все ячейки содержат None или пустые строки), иначе False.
    """
    # Проверяем значения ячеек
    for cell in row_index:
        if cell.value not in [None, ""]:
            return False
    return True

def copy_values_and_insert_formula(sheet, start_row, end_row):
    """ Функция копирует значения из столбца B в столбец A в диапазоне строк.
        А затем в столбец B вставляет формулу."""
   
    # Диапазон строк, в котором вы хотите скопировать значения
    start = start_row + 4

    # Копируем значения из столбца B в столбец A
    for row in range(start, end_row + 1):
        # Чтение значения из ячейки в столбце B
        value_b = sheet.cell(row=row, column=2).value
        
        # Запись значения в столбец A
        sheet.cell(row=row, column=1).value = value_b

    print(f'Значения из столбца B скопированы в столбец A в диапазоне от {start} до {end_row}')

    # Вставляем формулу в столбец B
    for row in range(start, end_row + 1):
        # Создаем формулу для объединения значений из столбца A и C
        formula = f'=CONCATENATE(A{row}, ".", C{row}, ".")'
        
        # Вставляем формулу в ячейку в столбце B
        sheet.cell(row=row, column=2).value = formula

    print(f'Формула вставлена в столбец B в диапазоне от {start} до {end_row}')
    

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
