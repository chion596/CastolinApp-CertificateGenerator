import io
import os
import shutil
import threading
from datetime import datetime

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

from config import CREDENTIALS_DATA, FOLDER_ID, SCOPES
from utils import resource_path
from certificates.type1 import certificate1
from certificates.type2 import certificate2
from certificates.type3 import certificate3

class CertificateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор сертификатов Castolin Eutectic")
        self.root.geometry("700x600")
        self.root.resizable(False, False)


        try:
            icon_path = resource_path(os.path.join("assets", "icon.ico"))
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass


        self.cert_type = tk.StringVar(value="1")
        self.batch_number = tk.StringVar()


        self.drive_service = None
        self.latest_file = None
        self.latest_file_time = None


        self.setup_styles()
        self.setup_background()
        self.create_widgets()


        self.fade_in()

    def setup_styles(self):
        """Настройка стилей для виджетов"""
        style = ttk.Style()
        style.theme_use('clam')


        style.configure('Title.TLabel', font=('Segoe UI', 20, 'bold'), foreground='#2c3e50')
        style.configure('Subtitle.TLabel', font=('Segoe UI', 10), foreground='#7f8c8d')
        style.configure('Card.TFrame', background='white', relief='flat', borderwidth=0)


        style.configure('Accent.TButton',
                       font=('Segoe UI', 11, 'bold'),
                       padding=10)


        style.configure('TRadiobutton', font=('Segoe UI', 11))

    def setup_background(self):
        """Создание фонового изображения"""

        self.canvas = tk.Canvas(self.root, width=700, height=600, highlightthickness=0)
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)


        for i in range(600):
            color_value = 240 - int(i / 10)
            color = f'#{color_value:02x}{color_value:02x}{color_value:02x}'
            self.canvas.create_line(0, i, 700, i, fill=color, width=1)


        try:
            bg_image_path = resource_path(os.path.join("assets", "background.png"))
            if os.path.exists(bg_image_path):
                bg_image = Image.open(bg_image_path)
                bg_image = bg_image.resize((700, 600), Image.Resampling.LANCZOS)
                self.bg_photo = ImageTk.PhotoImage(bg_image)
                self.canvas.create_image(0, 0, image=self.bg_photo, anchor='nw')
        except:
            pass

    def create_widgets(self):
        """Создание всех виджетов интерфейса"""

        main_frame = tk.Frame(self.root, bg='#f0f0f0', highlightthickness=0)
        main_frame.place(relx=0.5, rely=0.5, anchor='center', width=550, height=500)


        title_frame = tk.Frame(main_frame, bg='#f0f0f0', highlightthickness=0)
        title_frame.pack(pady=(0, 20))

        title_label = tk.Label(title_frame,
                              text="Генератор сертификатов",
                              font=('Segoe UI', 24, 'bold'),
                              fg='#2c3e50',
                              bg='#f0f0f0')
        title_label.pack()

        subtitle_label = tk.Label(title_frame,
                                 text="Castolin Eutectic",
                                 font=('Segoe UI', 12),
                                 fg='#7f8c8d',
                                 bg='#f0f0f0')
        subtitle_label.pack()


        type_card = tk.Frame(main_frame, bg='white', relief='flat', bd=1,
                            highlightthickness=1, highlightbackground='#e0e0e0')
        type_card.pack(pady=10, padx=20, fill='x')

        type_label = tk.Label(type_card, text="Выберите тип сертификата",
                             font=('Segoe UI', 12, 'bold'),
                             fg='#2c3e50',
                             bg='white')
        type_label.pack(pady=(15, 10))


        types_frame = tk.Frame(type_card, bg='white')
        types_frame.pack(pady=10, padx=20)

        self.type_buttons = []
        types = [
            ("Сварочные материалы", "1"),
            ("Порошки", "2"),
            ("Керамика", "3")
        ]

        for text, value in types:
            btn = tk.Button(types_frame, text=text,
                           font=('Segoe UI', 10),
                           bg='#ecf0f1',
                           fg='#2c3e50',
                           relief='flat',
                           bd=0,
                           padx=20,
                           pady=8,
                           cursor='hand2',
                           command=lambda v=value: self.select_type(v))
            btn.pack(side='left', padx=5)
            self.type_buttons.append((btn, value))


        batch_card = tk.Frame(main_frame, bg='white', relief='flat', bd=1,
                             highlightthickness=1, highlightbackground='#e0e0e0')
        batch_card.pack(pady=10, padx=20, fill='x')

        batch_label = tk.Label(batch_card, text="Введите номер партии:",
                              font=('Segoe UI', 12, 'bold'),
                              fg='#2c3e50',
                              bg='white')
        batch_label.pack(pady=(15, 5))


        entry_frame = tk.Frame(batch_card, bg='white')
        entry_frame.pack(pady=10, padx=20, fill='x')


        self.batch_entry = tk.Text(entry_frame,
                                   height=1,
                                   font=('Segoe UI', 11),
                                   bg='#f8f9fa',
                                   fg='#2c3e50',
                                   relief='solid',
                                   bd=1,
                                   highlightthickness=1,
                                   highlightcolor='#3498db',
                                   highlightbackground='#e0e0e0',
                                   wrap='none')
        self.batch_entry.pack(fill='x', ipady=8, padx=10)


        '''
        scrollbar = tk.Scrollbar(entry_frame, orient='horizontal', command=self.batch_entry.xview)
        self.batch_entry.configure(xscrollcommand=scrollbar.set)
        scrollbar.pack(fill='x', padx=10, pady=(0, 5))
        '''


        def get_text():
            return self.batch_entry.get("1.0", "end-1c").strip()


        def set_text(value):
            self.batch_entry.delete("1.0", "end")
            self.batch_entry.insert("1.0", value)


        def on_text_change(event=None):
            self.batch_number.set(get_text())

        self.batch_entry.bind("<KeyRelease>", on_text_change)
        self.batch_entry.bind("<FocusOut>", on_text_change)


        context_menu = tk.Menu(self.batch_entry, tearoff=0)
        context_menu.add_command(label="Вырезать", command=lambda: self.batch_entry.event_generate("<<Cut>>"))
        context_menu.add_command(label="Копировать", command=lambda: self.batch_entry.event_generate("<<Copy>>"))
        context_menu.add_command(label="Вставить", command=lambda: self.batch_entry.event_generate("<<Paste>>"))
        context_menu.add_separator()
        context_menu.add_command(label="Выделить всё", command=lambda: self.batch_entry.tag_add("sel", "1.0", "end"))

        def show_context_menu(event):
            context_menu.post(event.x_root, event.y_root)
            return "break"

        self.batch_entry.bind("<Button-3>", show_context_menu)


        self.generate_btn = tk.Button(main_frame,
                                      text="✨ СОЗДАТЬ СЕРТИФИКАТ ✨",
                                      font=('Segoe UI', 12, 'bold'),
                                      bg='#3498db',
                                      fg='white',
                                      relief='flat',
                                      bd=0,
                                      padx=30,
                                      pady=12,
                                      cursor='hand2',
                                      activebackground='#2980b9',
                                      activeforeground='white',
                                      command=self.generate_certificate)
        self.generate_btn.pack(pady=20)


        help_btn = tk.Button(main_frame,
                            text="❓ Помощь",
                            font=('Segoe UI', 10),
                            bg='#95a5a6',
                            fg='white',
                            relief='flat',
                            bd=0,
                            padx=20,
                            pady=5,
                            cursor='hand2',
                            activebackground='#7f8c8d',
                            command=self.show_help)
        help_btn.pack(pady=(0, 10))


        log_card = tk.Frame(main_frame, bg='white', relief='flat', bd=1,
                           highlightthickness=1, highlightbackground='#e0e0e0')
        log_card.pack(pady=10, padx=20, fill='both', expand=True)

        log_label = tk.Label(log_card, text="Статус выполнения",
                            font=('Segoe UI', 10, 'bold'),
                            fg='#2c3e50',
                            bg='white')
        log_label.pack(pady=(10, 5))


        log_text_frame = tk.Frame(log_card, bg='white')
        log_text_frame.pack(pady=5, padx=10, fill='both', expand=True)

        self.log_text = tk.Text(log_text_frame,
                                height=6,
                                state="disabled",
                                wrap="word",
                                font=('Consolas', 9),
                                bg='#f8f9fa',
                                fg='#2c3e50',
                                relief='flat',
                                bd=1)
        self.log_text.pack(side='left', fill='both', expand=True)

        scrollbar_log = tk.Scrollbar(log_text_frame, command=self.log_text.yview)
        scrollbar_log.pack(side='right', fill='y')
        self.log_text.configure(yscrollcommand=scrollbar_log.set)


        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate', length=300)


        self.select_type("1")

    def select_type(self, value):
        """Выбор типа сертификата"""
        self.cert_type.set(value)

        for btn, btn_value in self.type_buttons:
            if btn_value == value:
                btn.configure(bg='#3498db', fg='white')
            else:
                btn.configure(bg='#ecf0f1', fg='#2c3e50')

    def fade_in(self):
        """Эффект плавного появления"""
        alpha = self.root.attributes('-alpha')
        if alpha < 1:
            alpha += 0.05
            self.root.attributes('-alpha', alpha)
            self.root.after(20, self.fade_in)

    def log(self, message, is_error=False):
        """Добавить сообщение в лог"""
        def update():
            self.log_text.configure(state="normal")


            timestamp = datetime.now().strftime("%H:%M:%S")


            if is_error:
                self.log_text.insert("end", f"[{timestamp}] ❌ {message}\n", "error")
                self.log_text.tag_config("error", foreground="#e74c3c")
            else:
                self.log_text.insert("end", f"[{timestamp}] ✅ {message}\n")

            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(0, update)

    def show_help(self):
        """Показать окно помощи"""
        help_window = tk.Toplevel(self.root)
        help_window.title("Помощь")
        help_window.geometry("500x450")
        help_window.resizable(False, False)


        help_window.transient(self.root)
        help_window.grab_set()


        try:
            icon_path = resource_path(os.path.join("assets", "icon.ico"))
            if os.path.exists(icon_path):
                help_window.iconbitmap(icon_path)
        except:
            pass


        main_frame = tk.Frame(help_window, bg='white', padx=20, pady=20)
        main_frame.pack(fill='both', expand=True)


        title = tk.Label(main_frame, text="📖 Руководство пользователя",
                        font=('Segoe UI', 16, 'bold'),
                        bg='white',
                        fg='#2c3e50')
        title.pack(pady=(0, 20))


        text_frame = tk.Frame(main_frame, bg='white')
        text_frame.pack(fill='both', expand=True)

        help_text = tk.Text(text_frame, wrap='word', font=('Segoe UI', 10),
                           bg='#f8f9fa', fg='#2c3e50', relief='flat', bd=1)
        help_text.pack(side='left', fill='both', expand=True)

        scrollbar = tk.Scrollbar(text_frame, command=help_text.yview)
        scrollbar.pack(side='right', fill='y')
        help_text.configure(yscrollcommand=scrollbar.set)


        content = """
1️ Выбор одного из трёх доступных типов сертификата
   • Сварочные материалы
   • Порошки
   • Керамика

2️ Ввод номера партии
   Введите номер партии в соответствующее поле.
   Номер должен точно совпадать с номером в Excel-таблице.
   Поддерживается вставка текста через:
   • Правую кнопку мыши (контекстное меню)
   • Горячие клавиши Ctrl+V

3️ Создание сертификата
   Нажмите кнопку "СОЗДАТЬ СЕРТИФИКАТ".
   Приложение автоматически загрузит последнюю версию таблицы с Google Drive и найдет нужную партию и сгенерирует сертификат.

4️ Сохранение файла
   После генерации выберите папку для сохранения файла.
   Сертификат сохраняется в формате .docx.

📌 Важно:
   • Для работы требуется интернет-соединение
   • Excel-таблица должна находиться в указанной папке Google Drive
   • Все данные должны быть корректно заполнены в таблице

❓ Если возникли проблемы:
   В случае какой-то ошибки - должно выскочить окно с описанием того, что пошло не так. Если же после этого остались какие-то вопросы, обратитесь к администратору системы.
        """

        help_text.insert('1.0', content)
        help_text.configure(state='disabled')


        close_btn = tk.Button(main_frame, text="Закрыть",
                             command=help_window.destroy,
                             font=('Segoe UI', 10),
                             bg='#3498db',
                             fg='white',
                             relief='flat',
                             padx=20,
                             pady=5,
                             cursor='hand2')
        close_btn.pack(pady=(20, 0))

    def generate_certificate(self):
        """Генерация сертификата в отдельном потоке"""
        cert_type = self.cert_type.get()

        batch = self.batch_entry.get("1.0", "end-1c").strip()

        if not cert_type:
            messagebox.showerror("Ошибка", "Выберите тип сертификата")
            return
        if not batch:
            messagebox.showerror("Ошибка", "Введите номер партии")
            return


        self.generate_btn.config(state="disabled", text="⏳ ГЕНЕРАЦИЯ...")
        self.progress_bar.pack(pady=(0, 10))
        self.progress_bar.start()

        self.log(f"Начинаем генерацию сертификата (тип: {cert_type})...")

        thread = threading.Thread(target=self._generate_thread, args=(cert_type, batch))
        thread.daemon = True
        thread.start()

    def _generate_thread(self, cert_type, batch):
        """Фоновый процесс генерации"""
        try:
            self.log("Получение последнего файла с Google Drive...")
            latest = self.get_latest_file()
            if not latest:
                self.root.after(0, lambda: messagebox.showerror("Ошибка", "Не удалось получить файл с Google Drive (проверьте интернет-сооединение!)"))
                self.log("Ошибка: файл не найден", is_error=True)
                self._enable_button()
                return

            self.log(f"Загружаем файл: {latest['name']}")
            file_stream = self.download_excel_file(latest['id'])
            if not file_stream:
                self.root.after(0, lambda: messagebox.showerror("Ошибка", "Не удалось загрузить Excel файл (проверьте интернет-сооединение и наличие свободного места для загрузки!)"))
                self.log("Ошибка загрузки файла", is_error=True)
                self._enable_button()
                return

            df = self.process_excel_file(file_stream)
            if df is None:
                self.root.after(0, lambda: messagebox.showerror("Ошибка", "Не удалось прочитать Excel файл (что-то не так с файлом на Google Drive)"))
                self.log("Ошибка чтения Excel", is_error=True)
                self._enable_button()
                return

            column_name = 'Номер партии'
            if column_name not in df.columns:
                self.root.after(0, lambda: messagebox.showerror("Ошибка", f"Столбец '{column_name}' не найден в файле (таблица была изменена под другой формат!)"))
                self.log("Ошибка: столбец не найден", is_error=True)
                self._enable_button()
                return

            row_data = None
            for idx, row in df.iterrows():
                if str(row[column_name]) == batch:
                    row_data = row.to_dict()
                    break

            if not row_data:
                self.root.after(0, lambda: messagebox.showerror("Ошибка", f"Номер партии {batch} не найден (скорее всего была допущена опечатка/эту партию ещё не занесли в таблицу)"))
                self.log("Номер партии не найден", is_error=True)
                self._enable_button()
                return


            for key, val in row_data.items():
                if pd.isna(val):
                    row_data[key] = ' '
                else:
                    row_data[key] = str(val)

            self.log("Данные найдены, генерируем сертификат...")


            if cert_type == '1':
                filename = certificate1(row_data)
                cert_name = "Сварочные материалы"
            elif cert_type == '2':
                filename = certificate2(row_data)
                cert_name = "Порошки"
            else:
                filename = certificate3(row_data)
                cert_name = "Керамика"

            self.log(f"Сертификат '{cert_name}' успешно сгенерирован")


            default_name = f"Сертификат_{cert_type}_{batch}.docx"
            save_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word документы", "*.docx"), ("Все файлы", "*.*")],
                initialfile=default_name,
                title="Сохранить сертификат как..."
            )

            if save_path:
                shutil.copy2(filename, save_path)
                self.log(f"Сертификат сохранён: {os.path.basename(save_path)}")
                self.root.after(0, lambda: messagebox.showinfo("Успех", f"✅ Сертификат успешно сохранён!\n\n{os.path.basename(save_path)}"))
            else:
                self.log("Сохранение отменено пользователем")


            try:
                os.remove(filename)
            except:
                pass

        except Exception as e:
            self.log(f"Ошибка: {str(e)}", is_error=True)
            self.root.after(0, lambda: messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}"))
        finally:
            self._enable_button()

    def _enable_button(self):
        """Включить кнопку и скрыть прогресс"""
        def enable():
            self.generate_btn.config(state="normal", text="✨ СОЗДАТЬ СЕРТИФИКАТ ✨")
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
        self.root.after(0, enable)


    def get_drive_service(self):
        """Получить сервис Google Drive"""
        if not self.drive_service:
            creds = Credentials.from_service_account_info(CREDENTIALS_DATA, scopes=SCOPES)
            self.drive_service = build('drive', 'v3', credentials=creds)
        return self.drive_service

    def get_latest_file(self):
        """Получить последний файл из папки Google Drive"""
        try:
            service = self.get_drive_service()
            query = f"'{FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
            results = service.files().list(
                q=query,
                orderBy='modifiedTime desc',
                fields="files(id, name, modifiedTime)"
            ).execute()
            files = results.get('files', [])

            if not files:
                return None

            newest_file = files[0]
            file_time = datetime.strptime(newest_file['modifiedTime'], '%Y-%m-%dT%H:%M:%S.%fZ')

            if self.latest_file is None or file_time > self.latest_file_time:
                self.latest_file = newest_file
                self.latest_file_time = file_time
                self.log(f"Обновлен файл: {newest_file['name']}")

            return self.latest_file
        except Exception as e:
            self.log(f"Ошибка при получении файлов: {e}", is_error=True)
            return None

    def download_excel_file(self, file_id):
        """Скачать Excel файл с Google Drive"""
        try:
            service = self.get_drive_service()
            request = service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)

            done = False
            while not done:
                _, done = downloader.next_chunk()

            fh.seek(0)
            return fh
        except Exception as e:
            self.log(f"Ошибка при загрузке файла: {e}", is_error=True)
            return None

    def process_excel_file(self, file_stream):
        """Обработать Excel файл"""
        try:
            df = pd.read_excel(file_stream)
            return df
        except Exception as e:
            self.log(f"Ошибка при чтении Excel: {e}", is_error=True)
            return None

