import os
import shutil
import time
import psutil

def find_process_using_file(file_path):
    """
    Функция для выявления процесса, использующего заданный файл.
    """
    # Получаем полный путь файла
    full_file_path = os.path.abspath(file_path)
    
    # Перебираем все процессы
    for process in psutil.process_iter(attrs=['pid', 'name']):
        try:
            # Получаем список файлов, открытых процессом
            open_files = process.open_files()
            for file in open_files:
                # Сравниваем полный путь файла с открытыми файлами процесса
                if full_file_path == file.path:
                    print(f"Файл {full_file_path} занят процессом: PID={process.pid}, Название процесса={process.name()}")
                    return process  # Возвращаем процесс, который использует файл
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            # Игнорируем процессы, которые не существуют или к которым нет доступа
            continue
    
    print(f"Файл {full_file_path} не используется ни одним процессом.")
    return None  # Возвращаем None, если файл не занят ни одним процессом

# Функция для получения списка всех файлов Excel в папке
def get_excel_files(directory):
    excel_files = [file for file in os.listdir(directory) if file.lower().endswith(".xlsx")]
    return excel_files

def move_files(src_directory, dst_directory, file_list, skipped_files, option):
    """
    Перемещает файлы из исходной директории в конечную директорию, учитывая список исключенных файлов.

    Параметры:
    src_directory (str): Исходная директория, из которой перемещаются файлы.
    dst_directory (str): Конечная директория, в которую перемещаются файлы.
    file_list (list): Список имен файлов для перемещения.
    skipped_files (list): Список имен файлов, которые не должны перемещаться.
    """
    for file_name in file_list:
        file_is_not_skipped = file_name not in skipped_files

        if file_is_not_skipped:
            # Определяем пути к исходному и конечному файлу
            src_file_path = os.path.join(src_directory, file_name)
            dst_file_path = os.path.join(dst_directory, file_name)

            if option == "OEN":
                if file_name.endswith('_oen.xlsx'):
                        shutil.move(src_file_path, dst_file_path)
                else:
                    time.sleep(1)
                    os.remove(src_file_path)
            else:
                # Перемещаем файл
                shutil.move(src_file_path, dst_file_path)
                print(f"Файл '{file_name}' перемещен в '{dst_directory}'.")

    print("Перемещение файлов завершено.")

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
