import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os
import logging
from excel_processing import process_all_files


# Функция для запуска обработки всех файлов по нажатию кнопки
def on_process_button_click():
    # Получаем введенные пользователем данные
    installation_name = installation_name_entry.get()
    control_date = control_date_entry.get()

    # Проверяем формат даты
    try:
        datetime.strptime(control_date, '%d.%m.%Y')
    except ValueError:
        messagebox.showerror("Ошибка", "Неправильный формат даты. Используйте формат DD.MM.YYYY")
        return

    # Указываем путь к папке с файлами
    directory = os.getcwd()

    # Проверяем, какую опцию выбрал пользователь (РГФ или ОЭН)
    selected_option = option_var.get()
    print("🚀 ~ selected_option:", selected_option)

    if selected_option == "RGF":
        # Обработка с опцией РГФ
        file_count = process_all_files(directory, installation_name, control_date, selected_option)
        messagebox.showinfo("УСПЕХ!", f"Обработка завершена успешно! \n Опция РГФ была выбрана. \n Обработано {file_count} файлов.")
        root.quit()

    elif selected_option == "OEN":
        # Обработка с опцией ОЭН
        file_count = process_all_files(directory, installation_name, control_date, selected_option)
        messagebox.showinfo("УСПЕХ!", f"Обработка завершена успешно! \n Опция ОЭН была выбрана. \n Обработано {file_count} файлов.")
        root.quit()


    else:
        # Сообщение о том, что ни одна опция не выбрана
        messagebox.showinfo("Информация", "Опции РГФ и ОЭН не выбраны. Обработано 0 файлов.")
        logging.info('Опции РГФ и ОЭН не выбраны. Обработано 0 файлов.')
        root.quit()

def run_app():
    global root, installation_name_entry, control_date_entry, option_var

    root = tk.Tk()
    root.title("GDC Excel Worker")

    # Переменная для выбора опции (РГФ или ОЭН)
    option_var = tk.StringVar(value="")  # Значение по умолчанию пустое (ни одна опция не выбрана)

    # Создаем Radiobutton для РГФ
    rgf_radiobutton = ttk.Radiobutton(root, text="РГФ", variable=option_var, value="RGF")
    rgf_radiobutton.grid(row=0, column=0, padx=10, pady=10)

    # Создаем Radiobutton для ОЭН
    oen_radiobutton = ttk.Radiobutton(root, text="ОЭН", variable=option_var, value="OEN")
    oen_radiobutton.grid(row=0, column=1, padx=10, pady=10)

    # Создаем метку и поле для ввода названия установки
    ttk.Label(root, text="Введите название установки:").grid(row=1, column=0, padx=10, pady=10)
    installation_name_entry = ttk.Entry(root)
    installation_name_entry.grid(row=1, column=1, padx=10, pady=10)

    # Создаем метку и поле для ввода даты контроля
    ttk.Label(root, text="Введите дату контроля (ДД.ММ.ГГГГ):").grid(row=2, column=0, padx=10, pady=10)
    control_date_entry = ttk.Entry(root)
    control_date_entry.grid(row=2, column=1, padx=10, pady=10)

    # Создаем кнопку для запуска обработки
    process_button = ttk.Button(root, text="Запустить обработку", command=on_process_button_click)
    process_button.grid(row=3, column=0, columnspan=2, pady=10)

    # Запускаем главное окно приложения
    root.mainloop()