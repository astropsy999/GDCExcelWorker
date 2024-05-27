import os
from zipfile import BadZipFile
from process_excel_rgf import process_excel_file_rgf
from process_excel_oen import process_excel_file_oen
from utils import bcolors
from handle_processed_files import move_processed_files, convert_xls_to_xlsx
from utils import get_excel_files
from pathlib import Path


# Объединенная функция для обработки всех файлов Excel в папке
def process_all_files(directory, installation_name, control_date, option):
    
    path_dir = Path(directory)
    convert_xls_to_xlsx(path_dir)
    
    # Получить список файлов Excel
    excel_files, xls_files = get_excel_files(directory)
    skipped_files = []
    

    # Инициализируем счетчик обработанных файлов
    processed_count = 0

    # Обрабатываем файлы
    for file_name in excel_files:
        try:
            file_path = os.path.join(directory, file_name)

            # В зависимости от выбранной опции (RGF или OEN) вызываем соответствующую функцию обработки
            if option == "RGF":
                process_excel_file_rgf(file_path, installation_name, control_date)
            elif option == "OEN":
                process_excel_file_oen(file_path, installation_name)
            else:
                continue  # Пропускаем файлы, если опция не указана
            # Увеличиваем счетчик обработанных файлов
            processed_count += 1
            
        except KeyError as e:
            skipped_files.append(file_name)
            print(f'{bcolors.FAIL}В файле {file_name} произошла какая-то ошибка и он будет пропущен!{bcolors.ENDC}')
            continue
        except BadZipFile as e:
            skipped_files.append(file_name)
            print(f'{bcolors.FAIL}Файл {file_name} не может быть открыт, возможно он повржеден и будет пропущен!{bcolors.ENDC}')
            continue

    # Логирование итогов
    print(f'\n{bcolors.OKGREEN}Обработано {processed_count} файлов.{bcolors.ENDC}')
    
    move_processed_files(directory, option, skipped_files)
    
    return processed_count, skipped_files