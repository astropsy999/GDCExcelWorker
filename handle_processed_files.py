import os
import shutil
import time
from datetime import datetime
from utils import get_excel_files
from utils import bcolors

def create_output_directory():
    """Создает папку для сохранения обработанных файлов."""
    current_date = datetime.now().strftime("%d-%m-%Y")
    output_directory = f"Обработано_{current_date}"
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    return output_directory

def move_files_to_directory(file_list, src_directory, dst_directory):
    """Перемещает файлы из исходной директории в конечную."""
    for file_name in file_list:
        src_file_path = os.path.join(src_directory, file_name)
        dst_file_path = os.path.join(dst_directory, file_name)
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
    output_directory = create_output_directory()
    
    if option == "RGF":
        file_list = [file_name for file_name in get_excel_files(directory) if file_name not in skipped_files]
        move_files_to_directory(file_list, directory, output_directory)
    
    elif option == "OEN":
        file_list = [file_name for file_name in get_excel_files(directory) if file_name not in skipped_files]
        oen_files = [file_name for file_name in file_list if file_name.endswith('_oen.xlsx')]
        other_files = [file_name for file_name in file_list if not file_name.endswith('_oen.xlsx') and file_name not in skipped_files]

        move_files_to_directory(oen_files, directory, output_directory)
        delete_files(other_files, directory)
    
    print(f'\n{bcolors.OKGREEN}ГОТОВЫЕ ФАЙЛЫ СОХРАНЕНЫ в папке "{output_directory}"{bcolors.ENDC}.')
    print(f'{bcolors.WARNING}Пропущены файлы: "{skipped_files}"{bcolors.ENDC}.')

