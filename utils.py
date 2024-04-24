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
