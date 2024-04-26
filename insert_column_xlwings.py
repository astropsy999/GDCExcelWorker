import xlwings as xw
from utils import bcolors

def insert_column_after(file_path, start_range, end_range, to_insert_after, installation_name):
    # Открываем Excel-файл с помощью xlwings
    app = xw.App(visible=False)
    try:
        workbook = app.books.open(file_path)
        
        # Доступ к листу "УЗТ (коркарта)"
        try:
            sheet = workbook.sheets['УЗТ (коркарта)']
            print('Доступ к листу "УЗТ (коркарта)"')
        except KeyError:
            raise KeyError(f'{bcolors.FAIL} Лист "УЗТ (коркарта)" отсутствует в файле {file_path}.{bcolors.ENDC}')    
        
        # Определяем диапазон для вставки столбца после указанного столбца
        range_to_insert = sheet.range(f'{to_insert_after}{start_range}:{to_insert_after}{end_range}')
        
        # Вставляем пустой столбец после указанного столбца
        range_to_insert.insert('right')
        
        print(f'Вставлен пустой столбец после {to_insert_after} со сдвигом всей таблицы вправо в диапазоне строк с {start_range} по {end_range}.')
        
        # Вставляем значение "№ Зоны" в первую строку нового столбца и объединяем 3 верхние строки этого столбца
        
        new_column_index = 2
        # Вставляем значение "№ Зоны" в первую строку нового столбца
        new_column_range = sheet.range(start_range, new_column_index)
        new_column_range.value = "№ зоны"
        
        # Объединяем три верхние строки нового столбца
        upper_range = sheet.range((start_range, new_column_index), (start_range + 2, new_column_index))
        upper_range.merge()
        
        
        # Прочитайте значения диапазона A и C в массивы
        a_values = sheet.range(f"A{start_range + 4}:A{end_range}").value
        c_values = sheet.range(f"C{start_range + 4}:C{end_range}").value

        # Создайте массив для результатов
        results = []

        for a, c in zip(a_values, c_values):
            if a is not None and c is not None:# Округляем значения a и c до ближайшего целого числа
                rounded_a = round(a) if isinstance(a, float) else a
                rounded_c = round(c) if isinstance(c, float) else c
                
                # Формируем строку результата с округленными значениями
                results.append(f"{rounded_a}.{rounded_c}.")
        
        
        # Вы должны использовать диапазон для столбца B
        sheet.range(f"B{start_range + 4}:B{end_range}").options(transpose=True).value = results

        print('Результаты вставлены в диапазон B')
        
        sheet['E7'].value = installation_name
        print(f'Изменено::Технологическая установка (участок): {installation_name}')
    
        
        # Сохраняем изменения
        workbook.save(file_path +'_oen.xlsx')
    except Exception as e:
        print(f'Произошла ошибка при обработке файла: {e}')
    finally:
        # Закрываем Excel-файл и приложение
        workbook.close()
        app.quit()