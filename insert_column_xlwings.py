import xlwings as xw

def insert_column(sheet, column_index, column_name, file_path):
    """
    Вставляет новый столбец перед заданным столбцом в листе.

    Параметры:
    sheet (xlwings.Sheet): Рабочий лист, в котором вставляется столбец.
    column_index (int): Индекс столбца перед которым вставляется новый столбец (1-индексированный).
    """
    # Получаем диапазон ячеек для вставки столбца
    column_range = sheet.range((1, column_index), (sheet.api.UsedRange.Rows.Count, column_index))
    
    # Вставляем столбец перед указанным столбцом
    column_range.insert(shift='right', copy_origin='format_from_left_or_above')
    
    print(f"Новый столбец вставлен перед столбцом {column_index}")

# Пример использования:
if __name__ == "__main__":
    # Открываем приложение Excel
    app = xw.App(visible=False)  # False означает, что Excel не будет показан на экране

    # Открываем книгу Excel
    workbook = app.books.open(file_path)

    # Получаем рабочий лист по имени или индексу
    sheet = workbook.sheets['Имя_листа']  # Замените 'Имя_листа' на имя вашего листа
    
    # Вставляем столбец перед 2-м столбцом
    insert_column(sheet, 2)
    
    # Сохраняем изменения в книге
    workbook.save()

    # Закрываем книгу и завершаем приложение Excel
    workbook.close()
    app.quit()
