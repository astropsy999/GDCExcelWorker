import xlwings as xw

def save_excel_without_formulas(file_path, new_file_path):
    """
    Сохраняет Excel файл без формул, оставляя только текущие значения в ячейках.

    Параметры:
    file_path (str): Путь к исходному файлу Excel.
    new_file_path (str): Путь к новому файлу Excel без формул.
    """
    # Открываем приложение Excel
    app = xw.App(visible=False)  # False означает, что Excel не будет показан на экране

    # Открываем книгу Excel
    workbook = app.books.open(file_path)

    # Проходим по всем листам в книге
    for sheet in workbook.sheets:
        # Проходим по всем ячейкам в листе
        for cell in sheet.used_range:
            if cell.formula:  # Если ячейка содержит формулу
                # Заменяем формулу текущим значением ячейки
                cell.value = cell.value

    # Сохраняем книгу Excel в новом файле
    workbook.save(new_file_path)

    # Закрываем книгу и завершаем приложение Excel
    workbook.close()
    app.quit()

    print(f"Файл сохранен без формул как {new_file_path}")

# Пример использования:
if __name__ == "__main__":
    original_file = 'путь_к_исходному_файлу.xlsx'
    new_file = 'путь_к_новому_файлу.xlsx'
    save_excel_without_formulas(original_file, new_file)