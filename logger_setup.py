import logging

def setup_logger():
    log_filename = 'processing.log'

    # Открытие файла в режиме 'a+' для добавления
    with open(log_filename, 'a+', encoding='utf-8') as file:
        file.seek(0)  # Перемещаемся в начало файла

        # Настройка логгера
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.StreamHandler(),
                logging.FileHandler(log_filename, mode='a')  # Открытие лог-файла в режиме 'a' для добавления
            ]
        )
    # Открытие файла в режиме чтения и записи
    with open(log_filename, 'r+', encoding='utf-8') as file:
        # Чтение существующего содержимого файла
        existing_log_content = file.read()

        # Настройка логгера для дописывания новой информации вверху файла
        logging.basicConfig(filename=log_filename, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

        # Пишем существующее содержимое файла после новых логов
        file.write(existing_log_content)
