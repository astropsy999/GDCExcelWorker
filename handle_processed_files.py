import os
import shutil
import time
from datetime import datetime
from utils import get_excel_files
from utils import bcolors
from pathlib import Path
import pyexcel as p

def is_file_in_use(file_path):
    """Проверяет, используется ли файл каким-либо процессом, пытаясь открыть его в эксклюзивном режиме."""
    try:
        # Открываем файл в режиме эксклюзивного доступа
        file_handle = os.open(file_path, os.O_RDONLY | os.O_EXCL)
        os.close(file_handle)
        return False
    except OSError:
        # Если файл уже используется, возникает ошибка
        return True

def convert_xls_to_xlsx(directory):
    # Find all XLS files in the directory
    xls_files = [file for file in directory.glob("*.xls")]

    # Convert each XLS file to XLSX
    for xls_file in xls_files:
        xlsx_file = xls_file.with_suffix(".xlsx")
        p.save_book_as(file_name=str(xls_file), dest_file_name=str(xlsx_file))

            
def create_output_directory():
    """Создает папку для сохранения обработанных файлов."""
    current_date = datetime.now().strftime("%d-%m-%Y")
    output_directory = f"Обработано_{current_date}"
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    return output_directory

def move_files_to_directory(file_list, src_directory, dst_directory, retries=3, delay=2):
    """Перемещает файлы из исходной директории в конечную."""
    for file_name in file_list:
        src_file_path = os.path.join(src_directory, file_name)
        dst_file_path = os.path.join(dst_directory, file_name)
        if not is_file_in_use(src_file_path):
            shutil.move(src_file_path, dst_file_path)

def delete_files(file_list, src_directory, max_retries=2, delay=1):
    
    """Удаляет файлы из исходной директории с попытками и задержкой."""
    for file_name in file_list:
        src_file_path = os.path.join(src_directory, file_name)
        # if file_name in skipped_files:
        #     continue
        # Количество попыток удаления файла
        retries = 0
        
        # Пытаться удалить файл несколько раз, если возникает ошибка
        while retries < max_retries:
            try:
                # Попытка удалить файл
                os.remove(src_file_path)
                print(f'Файл {file_name} успешно удален.')
                break  # Если успешно, прерываем цикл
                
            except PermissionError:
                # Если ошибка доступа, подождите некоторое время перед следующей попыткой
                print(f'Не удается удалить файл {file_name}. Попытка повторить через {delay} секунд.')
                retries += 1
                time.sleep(delay)
        
        # Проверяем, был ли файл успешно удален после всех попыток
        if retries == max_retries:
            print(f'Не удалось удалить файл {file_name} после {max_retries} попыток.')

def move_processed_files(directory, option, skipped_files):
    """Перемещает обработанные файлы в выходную директорию."""
    output_directory = create_output_directory()
    
    if option == "RGF":
        excel_files, xls_files = get_excel_files(directory)
        file_list = [file_name for file_name in excel_files if file_name not in skipped_files]
        move_files_to_directory(file_list, directory, output_directory)
    
    elif option == "OEN":
        excel_files, xls_files = get_excel_files(directory)
        
        # Объединение списков excel_files и xls_files
        all_files = excel_files + xls_files
        
        # Фильтрация списка файлов
        file_list = [file_name for file_name in all_files if file_name not in skipped_files]
        oen_files = [file_name for file_name in file_list if file_name.endswith('_oen.xlsx')]
        other_files = [file_name for file_name in file_list if not file_name.endswith('_oen.xlsx') and file_name not in skipped_files]
        
        print('other_files: ', other_files)
        
        move_files_to_directory(oen_files, directory, output_directory)
        delete_files(other_files, directory)
    
    print(f'\n{bcolors.OKGREEN}ГОТОВЫЕ ФАЙЛЫ СОХРАНЕНЫ в папке "{output_directory}"{bcolors.ENDC}.')
    print(f'{bcolors.WARNING}Пропущены файлы: "{skipped_files}"{bcolors.ENDC}.')

