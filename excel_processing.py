import os
import logging
from datetime import datetime
import shutil
from process_excel_rgf import process_excel_file_rgf
from process_excel_oen import process_excel_file_oen
from utils import bcolors


# Функция для получения списка всех файлов Excel в папке
def get_excel_files(directory):
    excel_files = [file for file in os.listdir(directory) if file.lower().endswith(".xlsx")]
    print(f'Найдены файлы Excel: {excel_files}')
    return excel_files

# Объединенная функция для обработки всех файлов Excel в папке
def process_all_files(directory, installation_name, control_date, option):
    # Получить список файлов Excel
    excel_files = get_excel_files(directory)
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
            # Если лист не найден, пропускаем обработку текущего файла
            print(f"{bcolors.FAIL}Пропущен файл {file_name}: ({e}){bcolors.ENDC}")
            continue

    # Логирование итогов
    print(f'Обработано {processed_count} файлов.')
    print(f'{bcolors.WARNING}Пропущены: {skipped_files}{bcolors.ENDC}')
    # Создаем папку для сохранения обработанных файлов
    current_date = datetime.now().strftime("%d-%m-%Y")
    output_directory = f"Обработано_{current_date}"
    updated_excel_files = get_excel_files(directory)

    # Проверяем, существует ли уже папка с таким именем
    if os.path.exists(output_directory):
        # Если папка существует
        # Перемещаем обработанные файлы в существующую папку
        if option == "RGF":
            for file_name in excel_files:
                # Проверяем, что файл не пропущен
                if file_name not in skipped_files:
                    src_file_path = os.path.join(directory, file_name)
                    dst_file_path = os.path.join(output_directory, file_name)
                    shutil.move(src_file_path, dst_file_path)
        elif option == "OEN":
            for file_name in updated_excel_files:
                src_file_path = os.path.join(directory, file_name)
                dst_file_path = os.path.join(output_directory, file_name)
                if file_name.endswith('_oen.xlsx'):
                    shutil.move(src_file_path, dst_file_path)
                else:
                    os.remove(src_file_path)

        print(f'ГОТОВЫЕ ФАЙЛЫ СОХРАНЕНЫ в существующей папке "{output_directory}".')

    else:
        # Создаем новую папку
        os.makedirs(output_directory)
        if option == "RGF":
            # Перемещаем обработанные файлы в новую папку
            for file_name in excel_files:
                if file_name not in skipped_files:
                    src_file_path = os.path.join(directory, file_name)
                    dst_file_path = os.path.join(output_directory, file_name)
                    shutil.move(src_file_path, dst_file_path)
                
        elif option == "OEN":
            for file_name in updated_excel_files:
                src_file_path = os.path.join(directory, file_name)
                dst_file_path = os.path.join(output_directory, file_name)
                if file_name.endswith('_oen.xlsx'):
                    shutil.move(src_file_path, dst_file_path)
                else:
                    os.remove(src_file_path)

        print(f'ГОТОВЫЕ ФАЙЛЫ СОХРАНЕНЫ в новой папке "{output_directory}".')

    return processed_count