import os
import shutil
import time
from datetime import datetime
from utils import get_excel_files

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

def delete_files(file_list, src_directory):
    """Удаляет файлы из исходной директории."""
    for file_name in file_list:
        src_file_path = os.path.join(src_directory, file_name)
        time.sleep(1)
        os.remove(src_file_path)

def move_processed_files(directory, option, skipped_files):
    output_directory = create_output_directory()
    
    if option == "RGF":
        file_list = [file_name for file_name in get_excel_files(directory) if file_name not in skipped_files]
        move_files_to_directory(file_list, directory, output_directory)
    
    elif option == "OEN":
        file_list = [file_name for file_name in get_excel_files(directory) if file_name not in skipped_files]
        oen_files = [file_name for file_name in file_list if file_name.endswith('_oen.xlsx')]
        other_files = [file_name for file_name in file_list if not file_name.endswith('_oen.xlsx')]

        move_files_to_directory(oen_files, directory, output_directory)
       
    print(f'ГОТОВЫЕ ФАЙЛЫ СОХРАНЕНЫ в папке "{output_directory}".')

