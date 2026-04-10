import win32print
import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
from tkinter import PhotoImage
import base64
import subprocess
import os

# Функция для получения имени компьютера
def get_computer_name():
    return os.environ['COMPUTERNAME']

# Функция для определения филиала по имени компьютера
def get_branch_by_computer_name():
    computer_name = get_computer_name()
    for branch, prefix in branch_prefixes.items():
        if computer_name.startswith(prefix):
            return branch
    return None

# Словарь для хранения принтеров по филиалам
branch_printers = {
    "Актобе": [],
    "Алматы": [],
    "Астана": [],
    "Буденовск": [],
    "Владикавказ": [],
    "Волгоград": [],
    "Воронеж": [],
    "Домодедово": [],
    "Екатеринбург": [],
    "Жаркент": [],
    "Иркутск": [],
    "Казань": [],
    "Караганда": [],
    "Коледино": [],
    "Краснодар": [],
    "Красноярск": [],
    "Москва": [],
    "Набережные Челны": [],
    "Нальчик": [],
    "Нижнекамск": [],
    "Новосибирск": [],
    "Оренбург": [],
    "Орск": [],
    "Пенза": [],
    "Пермь": [],
    "Пятигорск": [],
    "Ростов-на-Дону": [],
    "Рязань": [],
    "Санкт-Петербург": [],
    "Саратов": [],
    "Симферополь": [],
    "Сочи": [],
    "Ставрополь": [],
    "Талдыкорган": [],
    "Тольятти": [],
    "Ульяновск": [],
    "Уфа": [],
    "Челябинск": [],
    "Черкеск": [],
}

# Словарь серверов для каждого филиала
branch_servers = {
    "Актобе": [f"\\\\print.akt.dalimo.ru\\{printer}" for printer in branch_printers["Актобе"]],
    "Алматы": [f"\\\\print.alt.dalimo.ru\\{printer}" for printer in branch_printers["Алматы"]],
    "Астана": [f"\\\\print.ast.dalimo.ru\\{printer}" for printer in branch_printers["Астана"]],
    "Буденовск": [f"\\\\print.skfo.dalimo.ru\\{printer}" for printer in branch_printers["Буденовск"]],
    "Владикавказ": [f"\\\\print.skfo.dalimo.ru\\{printer}" for printer in branch_printers["Владикавказ"]],
    "Волгоград": [f"\\\\print.vlg.dalimo.ru\\{printer}" for printer in branch_printers["Волгоград"]],
    "Воронеж": [f"\\\\print.vrn.dalimo.ru\\{printer}" for printer in branch_printers["Воронеж"]],
    "Домодедово": [f"\\\\print.msk.dalimo.ru\\{printer}" for printer in branch_printers["Домодедово"]],
    "Екатеринбург": [f"\\\\print.ekb.dalimo.ru\\{printer}" for printer in branch_printers["Екатеринбург"]],
    "Жаркент": [f"\\\\print.alt.dalimo.ru\\{printer}" for printer in branch_printers["Жаркент"]],
    "Иркутск": [f"\\\\print.irk.dalimo.ru\\{printer}" for printer in branch_printers["Иркутск"]],
    "Казань": [f"\\\\print.kzn.dalimo.ru\\{printer}" for printer in branch_printers["Казань"]],
    "Караганда": [f"\\\\print.krg.dalimo.ru\\{printer}" for printer in branch_printers["Караганда"]],
    "Коледино": [f"\\\\print.msk.dalimo.ru\\{printer}" for printer in branch_printers["Коледино"]],
    "Краснодар": [f"\\\\print.krd.dalimo.ru\\{printer}" for printer in branch_printers["Краснодар"]],
    "Красноярск": [f"\\\\print.krs.dalimo.ru\\{printer}" for printer in branch_printers["Красноярск"]],
    "Москва": [f"\\\\print.msk.dalimo.ru\\{printer}" for printer in branch_printers["Москва"]],
    "Набережные Челны": [f"\\\\print.kzn.dalimo.ru\\{printer}" for printer in branch_printers["Набережные Челны"]],
    "Нальчик": [f"\\\\print.skfo.dalimo.ru\\{printer}" for printer in branch_printers["Нальчик"]],
    "Нижнекамск": [f"\\\\print.kzn.dalimo.ru\\{printer}" for printer in branch_printers["Нижнекамск"]],
    "Новосибирск": [f"\\\\print.nvs.dalimo.ru\\{printer}" for printer in branch_printers["Новосибирск"]],
    "Оренбург": [f"\\\\print.orn.dalimo.ru\\{printer}" for printer in branch_printers["Оренбург"]],
    "Орск": [f"\\\\print.orn.dalimo.ru\\{printer}" for printer in branch_printers["Орск"]],
    "Пенза": [f"\\\\print.pnz.dalimo.ru\\{printer}" for printer in branch_printers["Пенза"]],
    "Пермь": [f"\\\\print.perm.dalimo.ru\\{printer}" for printer in branch_printers["Пермь"]],
    "Пятигорск": [f"\\\\print.skfo.dalimo.ru\\{printer}" for printer in branch_printers["Пятигорск"]],
    "Ростов-на-Дону": [f"\\\\print.rnd.dalimo.ru\\{printer}" for printer in branch_printers["Ростов-на-Дону"]],
    "Рязань": [f"\\\\print.msk.dalimo.ru\\{printer}" for printer in branch_printers["Рязань"]],
    "Санкт-Петербург": [f"\\\\print.spb.dalimo.ru\\{printer}" for printer in branch_printers["Санкт-Петербург"]],
    "Саратов": [f"\\\\print.srt.dalimo.ru\\{printer}" for printer in branch_printers["Саратов"]],
    "Симферополь": [f"\\\\print.simf.dalimo.ru\\{printer}" for printer in branch_printers["Симферополь"]],
    "Сочи": [f"\\\\print.krd.dalimo.ru\\{printer}" for printer in branch_printers["Сочи"]],
    "Ставрополь": [f"\\\\print.skfo.dalimo.ru\\{printer}" for printer in branch_printers["Ставрополь"]],
    "Талдыкорган": [f"\\\\print.alt.dalimo.ru\\{printer}" for printer in branch_printers["Талдыкорган"]],
    "Тольятти": [f"\\\\print.tlt.dalimo.ru\\{printer}" for printer in branch_printers["Тольятти"]],
    "Ульяновск": [f"\\\\print.ul.dalimo.ru\\{printer}" for printer in branch_printers["Ульяновск"]],
    "Уфа": [f"\\\\print.ufa.dalimo.ru\\{printer}" for printer in branch_printers["Уфа"]],
    "Челябинск": [f"\\\\print.chel.dalimo.ru\\{printer}" for printer in branch_printers["Челябинск"]],
    "Черкеск": [f"\\\\print.skfo.dalimo.ru\\{printer}" for printer in branch_printers["Черкеск"]],
}

# Префиксы для каждого филиала
branch_prefixes = {
    "Актобе": "AKT",
    "Алматы": "ALT",
    "Астана": "AST",
    "Буденовск": "BUD",
    "Владикавказ": "VLD",
    "Волгоград": "VLG",
    "Воронеж": "VRN",
    "Домодедово": "MSK-DMD",
    "Екатеринбург": "EKB",
    "Жаркент": "JKT",
    "Иркутск": "IRK",
    "Казань": "KZN",
    "Караганда": "KRG",
    "Коледино": "KLD",
    "Краснодар": "KRD",
    "Красноярск": "KRS",
    "Москва": "MSK",
    "Набережные Челны": "NCHL",
    "Нальчик": "NALCH",
    "Нижнекамск": "NZK",
    "Новосибирск": "NSK",
    "Оренбург": "ORN",
    "Орск": "ORS",
    "Пенза": "PNZ",
    "Пермь": "PERM",
    "Пятигорск": "PTG",
    "Ростов-на-Дону": "RND",
    "Рязань": "RZN",
    "Санкт-Петербург": "SPB",
    "Саратов": "SRT",
    "Симферополь": "SIMF",
    "Сочи": "SCH",
    "Ставрополь": "STW",
    "Талдыкорган": "TLD",
    "Тольятти": "TLT",
    "Ульяновск": "ULN",
    "Уфа": "UFA",
    "Челябинск": "CHEL",
    "Черкеск": "CHERK",
}


initial_branch = get_branch_by_computer_name() 

# Список серверов, к которым нужно делать запросы
target_servers = ["\\\\cld-prn01", "\\\\cld-prn02", "\\\\ykz-prn01"]

def is_server_available(server_name):
    try:
        subprocess.run(
            ["ping", "-n", "1", server_name.strip("\\"), "-w", "1000"], 
            stdout=subprocess.DEVNULL, 
            stderr=subprocess.DEVNULL, 
            check=True, 
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        return True
    except subprocess.CalledProcessError:
        return False

def get_printers():
    for server_name in target_servers:
        if not is_server_available(server_name):
            #print(f"Сервер {server_name} недоступен. Пропускаем.")
            continue
        
        try:
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_NAME, server_name, 1)
        except Exception as e:
            #print(f"Ошибка при получении принтеров с сервера {server_name}: {e}")
            continue

        for printer in printers:
            full_printer_name = printer[2]
            printer_name = full_printer_name.split("\\")[-1]

            #print(f"Обрабатываем принтер: {printer_name}")
            for branch, prefix in branch_prefixes.items():
                if printer_name.lower().startswith(prefix.lower()):
                    branch_printers[branch].append(printer_name)

    # Отладочный вывод: выводим состояние словаря branch_printers (нужно было, когда ошибки были)
    #print("Принтеры по филиалам:", branch_printers)

def update_printer_paths():
    # Обновляем пути до сетевых принтеров на основе текущих данных по синеймам
    printer_paths = {
        "Актобе": [f"\\\\print.akt.dalimo.ru\\{printer}" for printer in branch_printers["Актобе"]],
        "Алматы": [f"\\\\print.alt.dalimo.ru\\{printer}" for printer in branch_printers["Алматы"]],
        "Астана": [f"\\\\print.ast.dalimo.ru\\{printer}" for printer in branch_printers["Астана"]],
        "Буденовск": [f"\\\\print.skfo.dalimo.ru\\{printer}" for printer in branch_printers["Буденовск"]],
        "Владикавказ": [f"\\\\print.skfo.dalimo.ru\\{printer}" for printer in branch_printers["Владикавказ"]],
        "Волгоград": [f"\\\\print.vlg.dalimo.ru\\{printer}" for printer in branch_printers["Волгоград"]],
        "Воронеж": [f"\\\\print.vrn.dalimo.ru\\{printer}" for printer in branch_printers["Воронеж"]],
        "Домодедово": [f"\\\\print.msk.dalimo.ru\\{printer}" for printer in branch_printers["Домодедово"]],
        "Екатеринбург": [f"\\\\print.ekb.dalimo.ru\\{printer}" for printer in branch_printers["Екатеринбург"]],
        "Иркутск": [f"\\\\print.irk.dalimo.ru\\{printer}" for printer in branch_printers["Иркутск"]],
        "Жаркент": [f"\\\\print.alt.dalimo.ru\\{printer}" for printer in branch_printers["Жаркент"]],
        "Казань": [f"\\\\print.kzn.dalimo.ru\\{printer}" for printer in branch_printers["Казань"]],
        "Караганда": [f"\\\\print.krg.dalimo.ru\\{printer}" for printer in branch_printers["Караганда"]],
        "Коледино": [f"\\\\print.msk.dalimo.ru\\{printer}" for printer in branch_printers["Коледино"]],
        "Краснодар": [f"\\\\print.krd.dalimo.ru\\{printer}" for printer in branch_printers["Краснодар"]],
        "Красноярск": [f"\\\\print.krs.dalimo.ru\\{printer}" for printer in branch_printers["Красноярск"]],
        "Москва": [f"\\\\print.msk.dalimo.ru\\{printer}" for printer in branch_printers["Москва"]],
        "Набережные Челны": [f"\\\\print.kzn.dalimo.ru\\{printer}" for printer in branch_printers["Набережные Челны"]],
        "Нальчик": [f"\\\\print.skfo.dalimo.ru\\{printer}" for printer in branch_printers["Нальчик"]],
        "Нижнекамск": [f"\\\\print.kzn.dalimo.ru\\{printer}" for printer in branch_printers["Нижнекамск"]],
        "Новосибирск": [f"\\\\print.nvs.dalimo.ru\\{printer}" for printer in branch_printers["Новосибирск"]],
        "Оренбург": [f"\\\\print.orn.dalimo.ru\\{printer}" for printer in branch_printers["Оренбург"]],
        "Орск": [f"\\\\print.orn.dalimo.ru\\{printer}" for printer in branch_printers["Орск"]],
        "Пенза": [f"\\\\print.pnz.dalimo.ru\\{printer}" for printer in branch_printers["Пенза"]],
        "Пермь": [f"\\\\print.perm.dalimo.ru\\{printer}" for printer in branch_printers["Пермь"]],
        "Пятигорск": [f"\\\\print.skfo.dalimo.ru\\{printer}" for printer in branch_printers["Пятигорск"]],
        "Ростов-на-Дону": [f"\\\\print.rnd.dalimo.ru\\{printer}" for printer in branch_printers["Ростов-на-Дону"]],
        "Рязань": [f"\\\\print.msk.dalimo.ru\\{printer}" for printer in branch_printers["Рязань"]],
        "Санкт-Петербург": [f"\\\\print.spb.dalimo.ru\\{printer}" for printer in branch_printers["Санкт-Петербург"]],
        "Саратов": [f"\\\\print.srt.dalimo.ru\\{printer}" for printer in branch_printers["Саратов"]],
        "Симферополь": [f"\\\\print.simf.dalimo.ru\\{printer}" for printer in branch_printers["Симферополь"]],
        "Сочи": [f"\\\\print.krd.dalimo.ru\\{printer}" for printer in branch_printers["Сочи"]],
        "Ставрополь": [f"\\\\print.skfo.dalimo.ru\\{printer}" for printer in branch_printers["Ставрополь"]],
        "Талдыкорган": [f"\\\\print.alt.dalimo.ru\\{printer}" for printer in branch_printers["Талдыкорган"]],
        "Тольятти": [f"\\\\print.tlt.dalimo.ru\\{printer}" for printer in branch_printers["Тольятти"]],
        "Ульяновск": [f"\\\\print.ul.dalimo.ru\\{printer}" for printer in branch_printers["Ульяновск"]],
        "Уфа": [f"\\\\print.ufa.dalimo.ru\\{printer}" for printer in branch_printers["Уфа"]],
        "Челябинск": [f"\\\\print.chel.dalimo.ru\\{printer}" for printer in branch_printers["Челябинск"]],
        "Черкеск": [f"\\\\print.skfo.dalimo.ru\\{printer}" for printer in branch_printers["Черкеск"]],
    }
    return printer_paths


# Получаем инфу по принтерам с серверов во время запуска приложения
get_printers() 

# Обновляем пути до принтеров
PRINTER_PATHS = update_printer_paths()

# Выводим пути до принтеров для отладки (тоже нужно было во время ошибок)
#print("Пути до принтеров:", PRINTER_PATHS)

   
# Функция для подключения принтеров
def connect_printer(printer_name):
    printer_name = printer_name.strip()
    root.grab_set()
    status_label.config(text="Идет подключение принтера...", font=('Segoe UI', 12, 'bold'), bg="#F0F0F0", fg="#0288D1")
    root.update()

    try:
        # Находим филиал, к которому относится принтер
        branch = next(branch for branch, printers in branch_printers.items() if printer_name in printers)
        
        # Получаем полный путь из PRINTER_PATHS
        full_printer_path = next((path for path in PRINTER_PATHS[branch] if printer_name == path.split("\\")[-1]), None)

        # Проверяем, есть ли такой путь в списке принтеров
        if full_printer_path not in PRINTER_PATHS[branch]:
            raise Exception(f"Принтер {printer_name} не найден на сервере.")

        # Подключаем принтер
        win32print.AddPrinterConnection(full_printer_path)
        messagebox.showinfo("Успех", f"Принтер {printer_name} успешно подключен!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось подключить принтер: {e}")
    finally:
        root.grab_release()
        status_label.config(text="Готово!", font=('Segoe UI', 12, 'bold'), bg="#F0F0F0", fg="#0288D1")

# Функция для установки принтера по умолчанию
def set_default_printer(printer_name):
    printer_name = printer_name.strip()
    root.grab_set()
    root.update()

    try:
        # Находим филиал, к которому относится принтер
        branch = next(branch for branch, printers in branch_printers.items() if printer_name in printers)
        
        # Получаем полный путь из PRINTER_PATHS
        full_printer_path = next(path for path in PRINTER_PATHS[branch] if printer_name in path)

        # Проверяем, есть ли такой путь в списке принтеров
        if full_printer_path not in PRINTER_PATHS[branch]:
            raise Exception(f"Принтер {printer_name} не найден на сервере.")
        

        # Устанавливаем принтер по умолчанию
        win32print.SetDefaultPrinter(full_printer_path)
        messagebox.showinfo("Успех", f"Принтер {printer_name} установлен по умолчанию!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось установить принтер по умолчанию: {e}")
    finally:
        root.grab_release()
        status_label.config(text="Готово!", font=('Segoe UI', 12, 'bold'), bg="#F0F0F0", fg="#0288D1")

# Удаляем принтер с компа
def delete_printer(printer_name):
    printer_name = printer_name.strip()
    root.grab_set()
    root.update()

    try:
        # Находим филиал, к которому относится принтер
        branch = next(branch for branch, printers in branch_printers.items() if printer_name in printers)
        
        # Получаем полный путь из PRINTER_PATHS
        full_printer_path = next(path for path in PRINTER_PATHS[branch] if printer_name in path)

        # Проверяем, есть ли такой путь в списке принтеров
        if full_printer_path not in PRINTER_PATHS[branch]:
            raise Exception(f"Принтер {printer_name} не найден на сервере.")

        # Удаляем принтер
        win32print.DeletePrinterConnection(full_printer_path)
        messagebox.showinfo("Успех", f"Принтер {printer_name} успешно удален!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось удалить принтер: {e}")
    finally:
        root.grab_release()
        status_label.config(text="Готово!", font=('Segoe UI', 12, 'bold'), bg="#F0F0F0", fg="#0288D1")

# Открываем настройки принтера на компе
def open_printer_settings(printer_name):
    printer_name = printer_name.strip()
    root.grab_set()
    root.update()

    try:
        # Находим филиал, к которому относится принтер
        branch = next(branch for branch, printers in branch_printers.items() if printer_name in printers)
        
        # Получаем полный путь из PRINTER_PATHS
        full_printer_path = next(path for path in PRINTER_PATHS[branch] if printer_name in path)

        # Проверяем, есть ли такой путь в списке принтеров
        if full_printer_path not in PRINTER_PATHS[branch]:
            raise Exception(f"Принтер {printer_name} не найден на сервере.")

        # Открываем настройки принтера
        subprocess.run(f'rundll32 printui.dll,PrintUIEntry /p /n "{full_printer_path}"', shell=True)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть настройки принтера: {e}")
    finally:
        root.grab_release()

def update_printers(event=None):  
    selected_branch = branch_combobox.get()

    # Очистка поля и установка плейсхолдера
    search_entry.delete(0, tk.END)
    set_placeholder()

    printers_listbox.delete(0, tk.END)  
    
    # Получаем список принтеров для выбранного филиала
    printer_names = branch_printers.get(selected_branch)
    
    if printer_names:
        printer_names.sort()
        
        max_width = 30
        for printer_name in printer_names:
            centered_printer_name = printer_name.center(max_width)  
            printers_listbox.insert(tk.END, centered_printer_name)
    
    # Показываем всё без фильтра по placeholder
    filter_printers(skip_placeholder=True)

# Преобразование base64 в нормальную картинку
def base64_to_image(base64_str):
    image_data = base64.b64decode(base64_str)
    return PhotoImage(data=image_data)

# Встраиваем строку base64 (получил ее в другом скрипте)
encoded_image = """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"""  # Вставьте вашу строку base64 здесь

# Создание окна
root = tk.Tk()
root.title("Подключение принтера")
root.geometry("950x650")
root.config(bg="#E3F2FD")  
root.resizable(True, True)

# Используем функцию для преобразования base64 в изображение
logo_image = base64_to_image(encoded_image)
logo_label = tk.Label(root, image=logo_image, bg="#E3F2FD")
logo_label.image = logo_image  # Сохраняем ссылку на изображение
logo_label.place(x=20, y=10)

# Шрифт для всего интерфейса
font_style = ('Segoe UI', 12)

# Заголовок
label = tk.Label(root, text="Выберите филиал для подключения принтера", font=('Segoe UI', 18, 'bold'), bg="#E3F2FD", fg="#0277BD")
label.pack(pady=20)

# Фрейм серенький
frame = tk.Frame(root, bg="#F0F0F0", bd=5, relief="groove")
frame.pack(padx=20, pady=20, fill="both", expand=True)

# branch_label и branch_combobox
branch_label = tk.Label(frame, text="Выберите филиал", font=('Segoe UI', 14, 'bold'), bg="#F0F0F0", fg="#0078D7")
branch_label.pack(pady=5)

branch_combobox = ttk.Combobox(
    frame,
    values=list(branch_printers.keys()),
    font=('Segoe UI', 12),
    justify="center",
    state="readonly"
)
branch_combobox.set(initial_branch if initial_branch else "")
branch_combobox.pack(pady=5)

# Поле поиска с плейсхолдером
search_var = tk.StringVar()
search_entry = tk.Entry(
    frame,
    textvariable=search_var,
    font=('Segoe UI', 12),
    width=22,
    fg='grey',
    justify='center'  # Центрируем текст
)
search_entry.pack(pady=(30, 0))

# Функции для плейсхолдера
def set_placeholder():
    if not search_var.get():
        search_entry.insert(0, "Поиск")
        search_entry.config(fg='grey')

def clear_placeholder(event):
    if search_entry.get() == "Поиск":
        search_entry.delete(0, tk.END)
        search_entry.config(fg='black')

def restore_placeholder(event):
    if not search_entry.get():
        set_placeholder()

# Привязка событий
search_entry.bind("<FocusIn>", clear_placeholder)
search_entry.bind("<FocusOut>", restore_placeholder)

# Установка начального плейсхолдера
set_placeholder() 

# Создаем Listbox и Scrollbar
listbox_frame = tk.Frame(frame, bg="#F0F0F0", padx=5)  
printers_listbox = tk.Listbox(listbox_frame, width=30, height=9, font=('Courier', 12))  
printers_scrollbar = tk.Scrollbar(listbox_frame, orient="vertical", command=printers_listbox.yview)
printers_listbox.config(yscrollcommand=printers_scrollbar.set)
printers_listbox.pack(side="left", fill="y", pady=10)
printers_scrollbar.pack(side="left", fill="y", padx=0)
listbox_frame.pack(pady=0, padx=20, expand=False)

# Создание метки с вопросительным знаком справа от заголовка
question_mark = tk.Label(frame, text="❓", font=('Segoe UI', 20), fg="#0277BD", bg="#F0F0F0", cursor="hand2")
question_mark.place(x=label.winfo_width() + 850, y=label.winfo_y())  

# Функция для отображения инструкции при нажатии на вопросительный знак
def show_instruction(event):
    messagebox.showinfo("Инструкция", "Инструкция по подключению принтеров и действиям с ними. \n\n1. Выберите филиал из списка.\n2. Выберите принтер и нажмите 'Подключить'.\n3. Для установки принтера по умолчанию выберите 'Установить по умолчанию'.\n4. Для удаления принтера с вашего ПК выберите 'Удалить'.\n5. Чтобы открыть настройки принтера выберите нужный принтер и нажмите 'Открыть настройки'.")

# Привязываем обработчик события нажатия на метку с вопросом
question_mark.bind("<Button-1>", show_instruction)

# Функция фильтрации
def filter_printers(*args, skip_placeholder=False):
    selected_branch = branch_combobox.get()
    search_term = search_var.get().lower()

    # Игнорируем фильтрацию, если плейсхолдер активен и фильтрация вызвана не вручную
    if not skip_placeholder and search_term == "поиск":
        return

    printers_listbox.delete(0, tk.END)
    printer_names = branch_printers.get(selected_branch, [])

    if printer_names:
        filtered = [p for p in printer_names if search_term in p.lower()] if not skip_placeholder else printer_names
        filtered.sort()
        max_width = printers_listbox.winfo_width() // 10

        for printer_name in filtered:
            centered_printer_name = printer_name.center(max_width)
            printers_listbox.insert(tk.END, centered_printer_name)

# Обновляем список при вводе
search_var.trace_add("write", filter_printers)

# Если филиал был определен, обновляем список принтеров с задержкой
if initial_branch:
    root.after(200, update_printers)  # 0.2 секунда задержки

# Привязка события выбора филиала в комбобоксе к функции update_printers
branch_combobox.bind("<<ComboboxSelected>>", update_printers)

# Метка для отображения статуса подключения
status_label = tk.Label(frame, text="", font=font_style, bg="#F0F0F0")  
status_label.pack(pady=5)

# Функции для обработки эффектов нажатия
def on_button_press(event, button_id):
    # Изменяем цвет кнопки при нажатии
    canvas.itemconfig(button_id, fill="#005C99")  # Цвет при нажатии
    canvas.itemconfig(button_id + "_text", fill="lightgray")  # Изменяем цвет текста при нажатии

def on_button_release(event, button_id, printer_name, action):
    # Восстанавливаем цвет после отпускания
    canvas.itemconfig(button_id, fill="#0078D7")  # Цвет после отпускания
    canvas.itemconfig(button_id + "_text", fill="white")  # Восстановление цвета текста
    # Выполнение действия
    action(printer_name)

# Функции для обработки эффектов нажатия
def on_button_press_default_sett(event, button_id):
    # Изменяем цвет кнопки и текста при нажатии
    canvas.itemconfig(button_id, fill="#005C99")  # Цвет кнопки при нажатии 
    canvas.itemconfig(button_id + "_text", fill="white")  # Цвет текста при нажатии 

def on_button_release_default_sett(event, button_id, printer_name, action):
    # Восстанавливаем цвет кнопки и текста после отпускания
    canvas.itemconfig(button_id, fill="#f0f0f0")  # Цвет кнопки после отпускания 
    canvas.itemconfig(button_id + "_text", fill="#0078D7")  # Цвет текста после отпускания 
    # Выполнение действия
    action(printer_name)

# Функции для обработки эффектов нажатия
def on_button_press_delete(event, button_id):
    # Изменяем цвет кнопки при нажатии
    canvas.itemconfig(button_id, fill="#D44040")  # Цвет при нажатии
    canvas.itemconfig(button_id + "_text", fill="lightgray")  # Изменяем цвет текста при нажатии

def on_button_release_delete(event, button_id, printer_name, action):
    # Восстанавливаем цвет после отпускания
    canvas.itemconfig(button_id, fill="#FF4C4C")  # Цвет после отпускания
    canvas.itemconfig(button_id + "_text", fill="white")  # Восстановление цвета текста
    # Выполнение действия
    action(printer_name)

# Фрейм для кнопок
buttons_frame = tk.Frame(frame, bg="#F0F0F0")
buttons_frame.pack(pady=40)

canvas = tk.Canvas(buttons_frame, width=880, height=100, bg="#F0F0F0", highlightthickness=0)
canvas.pack()

# Устанавливаем размеры и промежутки для каждой кнопки
button_height = 40  
button_gap = 20  

# Ширина каждой кнопки
connect_button_width = 200
set_default_button_width = 225
settings_button_width = 175

# Настройка координат кнопок с промежутками
connect_button = canvas.create_rectangle(10, 10, 10 + connect_button_width, 10 + button_height, fill="#0078D7", outline="", tags="connect")
canvas.create_text(10 + connect_button_width // 2, 10 + button_height // 2, text="Подключить", fill="white", font=('Segoe UI', 12, 'bold'), tags="connect_text")
canvas.tag_bind("connect", "<ButtonPress-1>", lambda e: on_button_press(e, "connect"))
canvas.tag_bind("connect", "<ButtonRelease-1>", lambda e: on_button_release(e, "connect", printers_listbox.get(tk.ACTIVE), connect_printer))
canvas.tag_bind("connect_text", "<ButtonPress-1>", lambda e: on_button_press(e, "connect"))
canvas.tag_bind("connect_text", "<ButtonRelease-1>", lambda e: on_button_release(e, "connect", printers_listbox.get(tk.ACTIVE), connect_printer))

# Кнопка "Установить по умолчанию"
set_default_button = canvas.create_rectangle(10 + connect_button_width + button_gap, 10, 10 + connect_button_width + button_gap + set_default_button_width, 10 + button_height, fill="#f0f0f0", outline="#0078D7", tags="set_default")
canvas.create_text(10 + connect_button_width + button_gap + set_default_button_width // 2, 10 + button_height // 2, text="Установить по умолчанию", fill="#0078D7", font=('Segoe UI', 12, 'bold'), tags="set_default_text")
canvas.tag_bind("set_default", "<ButtonPress-1>", lambda e: on_button_press_default_sett(e, "set_default"))
canvas.tag_bind("set_default", "<ButtonRelease-1>", lambda e: on_button_release_default_sett(e, "set_default", printers_listbox.get(tk.ACTIVE), set_default_printer))
canvas.tag_bind("set_default_text", "<ButtonPress-1>", lambda e: on_button_press_default_sett(e, "set_default"))
canvas.tag_bind("set_default_text", "<ButtonRelease-1>", lambda e: on_button_release_default_sett(e, "set_default", printers_listbox.get(tk.ACTIVE), set_default_printer))

# Кнопка "Открыть настройки"
settings_button = canvas.create_rectangle(10 + connect_button_width + button_gap + set_default_button_width + button_gap, 10, 10 + connect_button_width + button_gap + set_default_button_width + button_gap + settings_button_width, 10 + button_height, fill="#f0f0f0", outline="#0078D7", tags="settings")
canvas.create_text(10 + connect_button_width + button_gap + set_default_button_width + button_gap + settings_button_width // 2, 10 + button_height // 2, text="Открыть настройки", fill="#0078D7", font=('Segoe UI', 12, 'bold'), tags="settings_text")
canvas.tag_bind("settings", "<ButtonPress-1>", lambda e: on_button_press_default_sett(e, "settings"))
canvas.tag_bind("settings", "<ButtonRelease-1>", lambda e: on_button_release_default_sett(e, "settings", printers_listbox.get(tk.ACTIVE), open_printer_settings))
canvas.tag_bind("settings_text", "<ButtonPress-1>", lambda e: on_button_press_default_sett(e, "settings"))
canvas.tag_bind("settings_text", "<ButtonRelease-1>", lambda e: on_button_release_default_sett(e, "settings", printers_listbox.get(tk.ACTIVE), open_printer_settings))

# Кнопка "Удалить принтер"
delete_button = canvas.create_rectangle(10 + connect_button_width + button_gap + set_default_button_width + button_gap + settings_button_width + button_gap, 10, 10 + connect_button_width + button_gap + set_default_button_width + button_gap + settings_button_width + button_gap + connect_button_width, 10 + button_height, fill="#FF4C4C", outline="", tags="delete")
canvas.create_text(10 + connect_button_width + button_gap + set_default_button_width + button_gap + settings_button_width + button_gap + connect_button_width // 2, 10 + button_height // 2, text="Удалить", fill="white", font=('Segoe UI', 12, 'bold'), tags="delete_text")
canvas.tag_bind("delete", "<ButtonPress-1>", lambda e: on_button_press_delete(e, "delete"))
canvas.tag_bind("delete", "<ButtonRelease-1>", lambda e: on_button_release_delete(e, "delete", printers_listbox.get(tk.ACTIVE), delete_printer))
canvas.tag_bind("delete_text", "<ButtonPress-1>", lambda e: on_button_press_delete(e, "delete"))
canvas.tag_bind("delete_text", "<ButtonRelease-1>", lambda e: on_button_release_delete(e, "delete", printers_listbox.get(tk.ACTIVE), delete_printer))

root.mainloop()