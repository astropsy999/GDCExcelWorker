import xlwings as xw
import os

def insert_column_after(file_path, start_range, end_range, to_insert_after):
    # Открываем Excel-файл с помощью xlwings
    app = xw.App(visible=False)
    try:
        workbook = app.books.open(file_path)
        
        # Доступ к листу "УЗТ (коркарта)"
        sheet = workbook.sheets['УЗТ (коркарта)']
        print(f'----- НАЧАЛО ОБРАБОТКИ ФАЙЛА: {os.path.basename(file_path)} -----')
        print(f'Загружен файл Excel: {file_path}')
        
        # Определяем диапазон для вставки столбца после указанного столбца
        range_to_insert = sheet.range(f'{to_insert_after}{start_range}:{to_insert_after}{end_range}')
        
        # Вставляем пустой столбец после указанного столбца
        range_to_insert.insert('right')
        
        print(f'Вставлен пустой столбец после {to_insert_after} со сдвигом всей таблицы вправо в диапазоне строк с {start_range} по {end_range}.')
        
        # Сохраняем изменения
        workbook.save(file_path + "_new.xlsx")
    except Exception as e:
        print(f'Произошла ошибка при обработке файла: {e}')
    finally:
        # Закрываем Excel-файл и приложение
        workbook.close()
        app.quit()