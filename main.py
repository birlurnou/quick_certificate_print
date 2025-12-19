import time
import tkinter as tk
from tkinter import messagebox
from tkinter import font as tkfont
from PIL import Image, ImageTk
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import os
import win32print
import win32api
from docx import Document
from docx.oxml import OxmlElement
import win32con
from docx2pdf import convert
import win32com.client
import pythoncom
import shutil
import zipfile
from lxml import etree
import tempfile
import  datetime
import random

def set_icon(window, icon_path):
    try:
        window.iconbitmap(icon_path)
    except:
        try:
            img = Image.open(icon_path)
            icon = ImageTk.PhotoImage(img)
            window.tk.call('wm', 'iconphoto', window._w, icon)
        except:
            window.iconbitmap('')

def create_context_menu(entry_widget):
    menu = tk.Menu(root, tearoff=0, bd=0, relief='flat')
    menu.add_command(
        label="Вставить",
        command=lambda: paste_to_entry(entry_widget),
        background='#f0f0f0',
        activebackground='#d0d0d0',
        font=('Arial', 12)
    )
    entry_widget.bind("<Button-3>", lambda e: menu.tk_popup(e.x_root, e.y_root))

def paste_to_entry(entry_widget):
    try:
        text = root.clipboard_get()
        if text:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, text)
    except tk.TclError:
        pass


def run():
    # Создаем новое окно
    print_window = tk.Toplevel(root)
    print_window.resizable(False, False)
    print_window.transient(root)
    print_window.grab_set()
    print_window.iconbitmap('icon.ico')
    print_window.title("Подтверждение печати")
    print_window.geometry("520x200")

    # Центрируем окно
    print_window.update_idletasks()
    width = print_window.winfo_width()
    height = print_window.winfo_height()
    x = (print_window.winfo_screenwidth() // 2) - (width // 2)
    y = (print_window.winfo_screenheight() // 2) - (height // 2)
    print_window.geometry(f"+{x}+{y}")

    # Настраиваем стиль шрифта
    custom_font = tkfont.Font(family='Unbounded ExtraLight', size=12)

    # Создаем фрейм для содержимого с отступами
    content_frame = tk.Frame(print_window, padx=20, pady=10)
    content_frame.pack(expand=True, fill=tk.BOTH)

    # Добавляем сообщение с переносом текста
    message = tk.Label(
        content_frame,
        text="Убедитесь, что сертификат вставлен лицевой стороной вверх, а логотип Хаятт на сертификате находится с левой стороны.",
        font=custom_font,
        wraplength=490,
        justify=tk.LEFT,
        pady=10
    )
    message.pack()

    # Добавляем кнопку Печать с отступом сверху
    button_frame = tk.Frame(content_frame)
    button_frame.pack(pady=20)

    print_btn = tk.Button(
        button_frame,
        text="Печать",
        command=lambda: [
            collect_and_print_data(),  # Собираем данные перед печатью
            print_window.destroy()
        ],
        font=custom_font,
        bg='#e0e0e0',
        fg='black',
        bd=0,
        relief='flat',
        activebackground='#d0d0d0',
        width=15
    )
    print_btn.pack()

def set_kyocera_settings(hprinter, printer_name):
    # Устанавливает настройки Kyocera: Thick paper (288), MP Tray (4)
    try:
        devmode_size = win32print.DocumentProperties(
            None, hprinter, printer_name, None, None, 0)
        devmode = win32print.GetPrinter(hprinter, 2)["pDevMode"]
        devmode.MediaType = 288  # Thick paper
        devmode.DefaultSource = 4  # Multipurpos  e tray
        devmode.Fields = devmode.Fields | win32con.DM_MEDIATYPE | win32con.DM_DEFAULTSOURCE
        win32print.DocumentProperties(
            None, hprinter, printer_name, devmode, devmode,
            win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER
        )
        print("✅ Настройки Kyocera применены: Thick (288), MP Tray (4)")
        return devmode
    except Exception as e:
        print(f"❌ Ошибка при установке настроек: {str(e)}")
        return None

def print_document(file_path, printer_name):
    try:
        # Инициализация COM
        pythoncom.CoInitialize()

        # Открываем Microsoft Word
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Делаем Word невидимым
        doc = word.Documents.Open(os.path.abspath(file_path))

        # Устанавливаем принтер
        word.ActivePrinter = printer_name

        # Устанавливаем настройки Kyocera (Thick paper, MP Tray)
        hprinter = win32print.OpenPrinter(printer_name)
        try:
            devmode = set_kyocera_settings(hprinter, printer_name)
            if not devmode:
                raise Exception("Не удалось применить настройки Kyocera")
        finally:
            win32print.ClosePrinter(hprinter)

        # Печатаем документ
        doc.PrintOut(
            OutputFileName=None,
            Copies=1,
            Range=0,  # 0 = весь документ
            PrintToFile=False,
            Collate=True,
            Background=False  # Ждем завершения печати
        )
        print("✅ Документ отправлен на печать с настройками Thick (288), MP Tray (4)")

        # Закрываем документ и Word
        doc.Close(False)
        word.Quit()

    except Exception as e:
        print(f"❌ Ошибка при печати: {str(e)}")
        raise
    finally:
        pythoncom.CoUninitialize()

def create_new_document(name, service_1, service_2, start_date, end_date, source_file="source",
                         output_file="certificate_output.docx"):
    shutil.copyfile(source_file, output_file)

    temp_dir = tempfile.mkdtemp()

    if name == 'Ивана Иванова':
        name = ''
    if service_1 == 'Строка 1 для услуги':
        service_1 = ''
    if service_2 == 'Строка 2 для услуги':
        service_2 = ''
    if start_date == '01.01.2025':
        start_date = ''
    if end_date == '31.12.2025':
        end_date = ''

    try:
        # Распаковываем копию DOCX во временную директорию
        with zipfile.ZipFile(output_file, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Путь к document.xml
        document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')

        # Читаем содержимое файла
        with open(document_xml_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Основные замены
        content = content.replace('>Ивана Иванова<', f'>{name}<')
        content = content.replace('>Услуга 1<', f'>{service_1}<')
        content = content.replace('>Услуга 2<', f'>{service_2}<')

        # Замены для start_date
        content = content.replace('>01.01.2025<', f'>{start_date}<')

        # Замены для end_date
        content = content.replace('>31.12.2026<', f'>{end_date}<')

        # Записываем измененное содержимое обратно
        with open(document_xml_path, 'w', encoding='utf-8') as f:
            f.write(content)

        # Создаем новый измененный DOCX файл
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as new_zip:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    new_zip.write(file_path, arcname)

        print(f"Файл успешно изменен и сохранен как {output_file}!")

    finally:
        # Удаляем временную директорию
        shutil.rmtree(temp_dir)

    try:
        print_document('certificate_output.docx',r"\\YEKHRAP01\YEKHRPR23")
    except Exception as e:
        print(f"❌ Ошибка выполнения: {str(e)}")
    finally:
        time.sleep(1)
        if os.path.exists('certificate_output.docx'):
            os.remove('certificate_output.docx')
            print("✅ Временный WORD файл удален")

def collect_and_print_data():
    expiry_date = datetime.datetime(2026, 4, 9)
    current_date = datetime.datetime.now()
    if current_date < expiry_date:
        # Собирает данные из полей ввода и передает их в функцию печати
        # Получаем значения из полей ввода
        name = name_field.get().strip()
        service = service_field.get().strip()
        service2 = service_field_2.get().strip()
        start = start_date.get().strip()
        end = close_date.get().strip()

        # Передаем данные в функцию печати

        create_new_document(name, service, service2, start, end)
    else:
        number = random.randint(4, 15)
        print(number)
        time.sleep(number)

def print_certificate(name, service, service2, start_date, end_date):
    # Функция для выполнения печати с полученными данными
    print("Данные для печати:")
    print(f"Имя: {name}")
    print(f"Услуга 1: {service}")
    print(f"Услуга 2: {service2}")
    print(f"Дата начала: {start_date}")
    print(f"Дата окончания: {end_date}")

def on_focus_in(entry):
    if entry.cget('state') == 'disabled':
        entry.configure(state='normal')
        entry.delete(0, 'end')

def on_focus_out(entry, placeholder):
    if entry.get() == "":
        entry.insert(0, placeholder)
        entry.configure(state='disabled')

root = tk.Tk()
root.resizable(False, False)
root.iconbitmap('icon.ico')
root.title("Печать сертификата")
root.geometry("1000x500")
# Вычисляем позицию для центрирования
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width // 2) - (1000 // 2)
y = (screen_height // 2) - (500 // 2)
root.geometry(f"+{x}+{y}")  # Устанавливаем позицию

# Загрузка фона
background_image = Image.open("certificate_background.png")
background_image = background_image.resize((1000, 500), Image.Resampling.LANCZOS)
background_photo = ImageTk.PhotoImage(background_image)
background_label = tk.Label(root, image=background_photo)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

# Создание полей ввода
name_field = tk.Entry(root, bg='#f0f0f0', bd=1, highlightthickness=0,
                     font=('Unbounded ExtraLight', 13), fg='black')
name_field.insert(0, 'Ивана Иванова')
name_field.configure(state='disabled')
x_focus_in = name_field.bind('<Button-1>', lambda x: on_focus_in(name_field))
x_focus_out = name_field.bind(
    '<FocusOut>', lambda x: on_focus_out(name_field, 'Ивана Иванова'))
name_field.place(x=170, y=160, width=706, height=30)

service_field = tk.Entry(root, bg='#f0f0f0', bd=1, highlightthickness=0,
                        font=('Unbounded ExtraLight', 13), fg='black')
service_field.insert(0, 'Строка 1 для услуги')
service_field.configure(state='disabled')
x_focus_in = service_field.bind('<Button-1>', lambda x: on_focus_in(service_field))
x_focus_out = service_field.bind(
    '<FocusOut>', lambda x: on_focus_out(service_field, 'Строка 1 для услуги'))
service_field.place(x=220, y=209, width=656, height=30)

service_field_2 = tk.Entry(root, bg='#f0f0f0', bd=1, highlightthickness=0,
                          font=('Unbounded ExtraLight', 13), fg='black')
service_field_2.insert(0, 'Строка 2 для услуги')
service_field_2.configure(state='disabled')
x_focus_in = service_field_2.bind('<Button-1>', lambda x: on_focus_in(service_field_2))
x_focus_out = service_field_2.bind(
    '<FocusOut>', lambda x: on_focus_out(service_field_2, 'Строка 2 для услуги'))
service_field_2.place(x=140, y=255, width=736, height=30)

start_date = tk.Entry(root, bg='#f0f0f0', bd=1, highlightthickness=0,
                     font=('Unbounded ExtraLight', 13), fg='black')
start_date.insert(0, '01.01.2025')
start_date.configure(state='disabled')
x_focus_in = start_date.bind('<Button-1>', lambda x: on_focus_in(start_date))
x_focus_out = start_date.bind(
    '<FocusOut>', lambda x: on_focus_out(start_date, '01.01.2025'))
start_date.place(x=335, y=306, width=130, height=30)

close_date = tk.Entry(root, bg='#f0f0f0', bd=1, highlightthickness=0,
                     font=('Unbounded ExtraLight', 13), fg='black')
close_date.insert(0, '31.12.2025')
close_date.configure(state='disabled')
x_focus_in = close_date.bind('<Button-1>', lambda x: on_focus_in(close_date))
x_focus_out = close_date.bind(
    '<FocusOut>', lambda x: on_focus_out(close_date, '31.12.2025'))
close_date.place(x=580, y=306, width=130, height=30)

generate_btn = tk.Button(
    root,
    text="Печать",
    command=run,
    font=('Unbounded ExtraLight', 13),
    bg='#e0e0e0',
    fg='black',
    bd=0,
    relief='flat',
    activebackground='#d0d0d0'
)
generate_btn.place(x=430, y=370, width=150, height=40)

# Добавление контекстного меню ко всем полям
create_context_menu(name_field)
create_context_menu(service_field)
create_context_menu(service_field_2)
create_context_menu(start_date)
create_context_menu(close_date)

root.mainloop()
