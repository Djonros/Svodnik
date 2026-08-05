"""
GUI приложение для экспорта спецификаций из КОМПАС-3D v24.
Сводник - экспорт сводной ведомости из КОМПАС-3D.
"""

import sys
import os
import json
import itertools
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

from license import LicenseManager
from kompas_export_final import KompasExportFinal


RECENT_FILE = os.path.join(os.path.expanduser("~"), ".svodnik_recent.json")
MAX_RECENT = 10
SPINNER_FRAMES = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]


def load_recent():
    if os.path.exists(RECENT_FILE):
        try:
            with open(RECENT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []


def save_recent(files):
    with open(RECENT_FILE, "w", encoding="utf-8") as f:
        json.dump(files, f, ensure_ascii=False, indent=2)


def add_to_recent(filepath):
    files = load_recent()
    filepath = os.path.normpath(filepath)
    if filepath in files:
        files.remove(filepath)
    files.insert(0, filepath)
    files = files[:MAX_RECENT]
    save_recent(files)
    return files


class KompasExportApp:

    VERSION = "1.5.0"
    APP_NAME = "Сводник"

    def __init__(self):
        self.license = LicenseManager()
        try:
            from tkinterdnd2 import TkinterDnD
            self.root = TkinterDnD.Tk()
            self._dnd_available = True
        except Exception:
            self.root = tk.Tk()
            self._dnd_available = False
        self.root.title(f"{self.APP_NAME} v{self.VERSION}")
        self.root.geometry("620x340")
        self.root.resizable(False, True)

        self.show_recent = tk.BooleanVar(value=False)
        self.show_log = tk.BooleanVar(value=False)
        self.export_excel = tk.BooleanVar(value=True)
        self.export_word = tk.BooleanVar(value=True)
        self.export_pdf = tk.BooleanVar(value=True)
        self.export_costs = tk.BooleanVar(value=False)
        self._spinner_running = False
        self._spinner_cycle = itertools.cycle(SPINNER_FRAMES)
        self._current_template = None

        self._setup_menu()
        self._setup_ui()
        self._update_status()
        self._load_recent_list()
        self._setup_dragdrop()

    def _setup_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_checkbutton(label="Последние файлы", variable=self.show_recent,
                                  command=self._rebuild_panels)
        view_menu.add_checkbutton(label="Лог", variable=self.show_log,
                                  command=self._rebuild_panels)
        menubar.add_cascade(label="Вид", menu=view_menu)

        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="Ввести лицензию", command=self._show_license_dialog)
        settings_menu.add_command(label="Мой ID компьютера", command=self._show_machine_id)
        settings_menu.add_separator()
        settings_menu.add_command(label="Шаблоны...", command=self._show_template_dialog)
        settings_menu.add_command(label="Цены и калькулятор...", command=self._show_prices_dialog)
        settings_menu.add_separator()
        settings_menu.add_command(label="Выход", command=self.root.quit)
        menubar.add_cascade(label="Настройки", menu=settings_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="О программе", command=self._show_about)
        menubar.add_cascade(label="Справка", menu=help_menu)

    def _setup_ui(self):
        # --- Заголовок ---
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, padx=20, pady=(10, 0))

        left_header = ttk.Frame(header_frame)
        left_header.pack(side=tk.LEFT)

        ttk.Label(left_header, text=self.APP_NAME, font=("Arial", 18, "bold")).pack(anchor="w")
        ttk.Label(left_header, text="Экспорт сводной ведомости из КОМПАС-3D",
                  font=("Arial", 9), foreground="gray").pack(anchor="w")

        self.status_label = ttk.Label(header_frame, text="", font=("Arial", 9))
        self.status_label.pack(side=tk.RIGHT)

        ttk.Separator(self.root, orient="horizontal").pack(fill=tk.X, padx=20, pady=8)

        # --- Файл сборки ---
        file_frame = ttk.LabelFrame(self.root, text="Файл сборки", padding=10)
        file_frame.pack(fill=tk.X, padx=20, pady=(0, 5))

        file_row = ttk.Frame(file_frame)
        file_row.pack(fill=tk.X)

        self.file_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_row, textvariable=self.file_var, font=("Consolas", 10))
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        ttk.Button(file_row, text="Обзор...", command=self._browse_file).pack(side=tk.RIGHT)

        if self._dnd_available:
            ttk.Label(file_frame, text="или перетащите .a3d файл в окно",
                      font=("Arial", 8), foreground="gray").pack(anchor="w", pady=(5, 0))

        # --- Форматы экспорта ---
        fmt_frame = ttk.Frame(self.root)
        fmt_frame.pack(fill=tk.X, padx=20, pady=(0, 5))

        ttk.Label(fmt_frame, text="Форматы:", font=("Arial", 9)).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Checkbutton(fmt_frame, text="Excel", variable=self.export_excel).pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(fmt_frame, text="Word", variable=self.export_word).pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(fmt_frame, text="PDF", variable=self.export_pdf).pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(fmt_frame, text="Стоимость", variable=self.export_costs).pack(side=tk.LEFT, padx=8)

        # --- Кнопка ЭКСПОРТ + Спиннер ---
        action_frame = ttk.Frame(self.root)
        action_frame.pack(pady=8)

        self.export_btn = tk.Button(
            action_frame,
            text="  ЭКСПОРТ  ",
            command=self._export,
            font=("Arial", 11, "bold"),
            bg="#2F5496",
            fg="white",
            activebackground="#1F3D7A",
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            padx=30,
            pady=6,
        )
        self.export_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.spinner_label = ttk.Label(action_frame, text="", font=("Consolas", 14), width=2)
        self.spinner_label.pack(side=tk.LEFT)

        # --- Последние файлы (скрыт по умолчанию) ---
        self.recent_frame = ttk.LabelFrame(self.root, text="Последние файлы", padding=5)

        list_frame = ttk.Frame(self.recent_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.recent_listbox = tk.Listbox(list_frame, height=3, font=("Consolas", 9),
                                          yscrollcommand=scrollbar.set)
        self.recent_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.recent_listbox.yview)
        self.recent_listbox.bind("<Double-Button-1>", self._on_recent_double_click)

        recent_btn_frame = ttk.Frame(self.recent_frame)
        recent_btn_frame.pack(fill=tk.X, pady=(5, 0))

        ttk.Button(recent_btn_frame, text="Открыть", command=self._open_recent).pack(side=tk.LEFT, padx=2)
        ttk.Button(recent_btn_frame, text="Удалить", command=self._remove_recent).pack(side=tk.LEFT, padx=2)
        ttk.Button(recent_btn_frame, text="Очистить", command=self._clear_recent).pack(side=tk.LEFT, padx=2)

        # --- Лог (скрыт по умолчанию) ---
        self.log_frame = ttk.LabelFrame(self.root, text="Лог", padding=5)

        self.log_text = tk.Text(self.log_frame, height=4, state=tk.DISABLED, font=("Consolas", 9))
        self.log_text.pack(fill=tk.X)

    def _rebuild_panels(self):
        """Пересборка сворачиваемых секций."""
        self.recent_frame.pack_forget()
        self.log_frame.pack_forget()

        if self.show_recent.get():
            self.recent_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 5))
        if self.show_log.get():
            self.log_frame.pack(fill=tk.X, padx=20, pady=(0, 5))

        self.root.update_idletasks()
        self.root.geometry("")

    def _show_about(self):
        messagebox.showinfo("О программе",
            f"{self.APP_NAME} v{self.VERSION}\n\n"
            "Экспорт сводной ведомости из КОМПАС-3D v24\n"
            "в форматы Excel, Word и PDF.\n\n"
            "GitHub: github.com/Djonros/Svodnik")

    def _setup_dragdrop(self):
        if not self._dnd_available:
            return
        try:
            from tkinterdnd2 import DND_FILES
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

    def _on_drop(self, event):
        filepath = event.data.strip()
        if filepath.startswith("{") and filepath.endswith("}"):
            filepath = filepath[1:-1]
        if os.path.isfile(filepath):
            self.file_var.set(filepath)

    def _update_status(self):
        status = self.license.get_status()
        self.status_label.config(text=status)
        if not self.license.is_valid():
            self.status_label.config(foreground="red")
        else:
            self.status_label.config(foreground="green")

    def _load_recent_list(self):
        self.recent_listbox.delete(0, tk.END)
        for filepath in load_recent():
            display = filepath
            if len(display) > 60:
                display = "..." + display[-57:]
            self.recent_listbox.insert(tk.END, display)

    def _browse_file(self):
        filepath = filedialog.askopenfilename(
            title="Выберите файл сборки",
            filetypes=[("Файлы КОМПАС", "*.a3d *.a3t"), ("Все файлы", "*.*")]
        )
        if filepath:
            self.file_var.set(filepath)

    def _on_recent_double_click(self, event):
        self._open_recent()

    def _open_recent(self):
        selection = self.recent_listbox.curselection()
        if not selection:
            return
        files = load_recent()
        idx = selection[0]
        if idx < len(files):
            filepath = files[idx]
            if os.path.exists(filepath):
                self.file_var.set(filepath)
            else:
                messagebox.showwarning("Внимание", f"Файл не найден:\n{filepath}")
                files.pop(idx)
                save_recent(files)
                self._load_recent_list()

    def _remove_recent(self):
        selection = self.recent_listbox.curselection()
        if not selection:
            return
        files = load_recent()
        idx = selection[0]
        if idx < len(files):
            files.pop(idx)
            save_recent(files)
            self._load_recent_list()

    def _clear_recent(self):
        if messagebox.askyesno("Очистка", "Очистить список последних файлов?"):
            save_recent([])
            self._load_recent_list()

    def _start_spinner(self):
        self._spinner_running = True
        self._animate_spinner()

    def _stop_spinner(self):
        self._spinner_running = False
        self.spinner_label.config(text="")

    def _animate_spinner(self):
        if not self._spinner_running:
            return
        frame = next(self._spinner_cycle)
        self.spinner_label.config(text=frame)
        self.root.after(80, self._animate_spinner)

    def _log(self, message):
        if not self.show_log.get():
            self.show_log.set(True)
            self._rebuild_panels()
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _export(self):
        if not self.license.is_valid():
            remaining = self.license.MAX_EXPORTS_TRIAL - self.license.exports_count
            if remaining > 0:
                messagebox.showwarning("Пробный режим",
                    f"Осталось {remaining} экспортов.\nВведите лицензию для полной версии.")
            else:
                messagebox.showerror("Ошибка",
                    "Пробный период закончился!\nВведите лицензионный ключ.")
            return

        filepath = self.file_var.get()
        if not filepath:
            messagebox.showwarning("Внимание", "Выберите файл сборки!")
            return

        if not os.path.exists(filepath):
            messagebox.showerror("Ошибка", "Файл не найден!")
            return

        formats = set()
        if self.export_excel.get():
            formats.add("excel")
        if self.export_word.get():
            formats.add("word")
        if self.export_pdf.get():
            formats.add("pdf")
        if not formats:
            messagebox.showwarning("Внимание", "Выберите хотя бы один формат!")
            return

        self.export_btn.config(state="disabled")
        self._start_spinner()
        fmt_list = ", ".join(sorted(formats)).upper()
        self._log(f"Начало экспорта ({fmt_list})...")

        try:
            exporter = KompasExportFinal(template=self._current_template)

            import io
            old_stdout = sys.stdout
            sys.stdout = io.StringIO()

            if self.export_costs.get():
                exporter.run(filepath, formats=formats)
                exporter.calculate_costs()
                if "excel" in formats:
                    exporter.generate_excel()
            else:
                exporter.run(filepath, formats=formats)

            output = sys.stdout.getvalue()
            sys.stdout = old_stdout

            for line in output.strip().split('\n'):
                self._log(line)

            add_to_recent(filepath)
            self._load_recent_list()

            self.license.record_export()
            self._update_status()

            self._log("Экспорт завершен успешно!")
            messagebox.showinfo("Готово", "Экспорт завершен!\nФайлы сохранены в папке сборки.")

        except Exception as e:
            self._log(f"Ошибка: {str(e)}")
            messagebox.showerror("Ошибка", f"Ошибка экспорта:\n{str(e)}")
        finally:
            self._stop_spinner()
            self.export_btn.config(state="normal")

    def _show_machine_id(self):
        machine_id = self.license.get_machine_id()
        dialog = tk.Toplevel(self.root)
        dialog.title("ID компьютера")
        dialog.geometry("420x160")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="ID вашего компьютера:", font=("Arial", 10)).pack(pady=(15, 5))

        id_var = tk.StringVar(value=machine_id)
        id_entry = ttk.Entry(dialog, textvariable=id_var, width=35, state="readonly",
                             font=("Consolas", 11))
        id_entry.pack(pady=5)

        ttk.Label(dialog, text="Отправьте этот ID администратору\nдля получения лицензионного ключа",
                  font=("Arial", 9), foreground="gray").pack()

        def copy_id():
            self.root.clipboard_clear()
            self.root.clipboard_append(machine_id)
            messagebox.showinfo("Копирование", "ID скопирован в буфер обмена!")

        ttk.Button(dialog, text="Копировать", command=copy_id).pack(pady=10)

    def _show_license_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Активация лицензии")
        dialog.geometry("420x200")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Введите лицензионный ключ:", font=("Arial", 10)).pack(pady=(15, 5))

        key_var = tk.StringVar()
        key_entry = ttk.Entry(dialog, textvariable=key_var, width=35, font=("Consolas", 11))
        key_entry.pack(pady=5)

        ttk.Label(dialog, text="Формат: KEY-XXXXXXXXXXXX-YYYY-MM-DD",
                  font=("Arial", 8), foreground="gray").pack()

        def activate():
            key = key_var.get()
            if self.license.activate(key):
                messagebox.showinfo("Успех", "Лицензия активирована!")
                self._update_status()
                dialog.destroy()
            else:
                messagebox.showerror("Ошибка", "Неверный ключ лицензии!")

        ttk.Button(dialog, text="Активировать", command=activate).pack(pady=15)

    def _show_template_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Шаблоны")
        dialog.geometry("550x400")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Шаблоны экспорта:", font=("Arial", 10)).pack(pady=(15, 5))

        user_tmpl_dir = os.path.join(os.path.expanduser("~"), ".svodnik_templates")
        os.makedirs(user_tmpl_dir, exist_ok=True)

        app_dir = os.path.dirname(__file__)
        templates = {}

        for d in [app_dir, user_tmpl_dir]:
            if not os.path.exists(d):
                continue
            for f in os.listdir(d):
                if f.startswith("template_") and f.endswith(".json"):
                    path = os.path.join(d, f)
                    try:
                        import json
                        with open(path, "r", encoding="utf-8") as fh:
                            tmpl = json.load(fh)
                        name = tmpl.get("name", f)
                        templates[name] = {"path": path, "data": tmpl, "user": d == user_tmpl_dir}
                    except:
                        pass

        listbox = tk.Listbox(dialog, font=("Consolas", 10))
        listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        for name in templates:
            label = f"{name}" + (" [ваш]" if templates[name]["user"] else " [системный]")
            listbox.insert(tk.END, label)

        self._template_names = list(templates.keys())

        def apply_template():
            sel = listbox.curselection()
            if not sel:
                return
            name = self._template_names[sel[0]]
            self._current_template = templates[name]["data"]
            self._log(f"Шаблон: {name}")
            dialog.destroy()

        def create_template():
            self._edit_template_dialog(None, user_tmpl_dir, dialog)

        def edit_template():
            sel = listbox.curselection()
            if not sel:
                return
            name = self._template_names[sel[0]]
            tmpl = templates[name]
            if not tmpl["user"]:
                messagebox.showwarning("Внимание", "Системные шаблоны нельзя редактировать.\nСоздайте копию.")
                return
            self._edit_template_dialog(tmpl["data"], user_tmpl_dir, dialog)

        def delete_template():
            sel = listbox.curselection()
            if not sel:
                return
            name = self._template_names[sel[0]]
            tmpl = templates[name]
            if not tmpl["user"]:
                messagebox.showwarning("Внимание", "Системные шаблоны нельзя удалять.")
                return
            if messagebox.askyesno("Удаление", f"Удалить шаблон '{name}'?"):
                try:
                    os.remove(tmpl["path"])
                    self._log(f"Шаблон удалён: {name}")
                    dialog.destroy()
                    self._show_template_dialog()
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось удалить:\n{e}")

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Применить", command=apply_template).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="Создать", command=create_template).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="Редактировать", command=edit_template).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="Удалить", command=delete_template).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="Закрыть", command=dialog.destroy).pack(side=tk.LEFT, padx=3)

    def _edit_template_dialog(self, existing_data, save_dir, parent_dialog):
        import json

        dialog = tk.Toplevel(self.root)
        dialog.title("Редактор шаблона")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Шаблон:", font=("Arial", 10)).pack(pady=(10, 5))

        text = tk.Text(dialog, font=("Consolas", 9), width=70, height=25)
        text.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        if existing_data:
            text.insert("1.0", json.dumps(existing_data, ensure_ascii=False, indent=2))
        else:
            default = {
                "name": "Новый шаблон",
                "excel": {
                    "title": "СВОДНАЯ ВЕДОМОСТЬ",
                    "columns": [
                        {"key": "position", "header": "Позиция", "width": 8},
                        {"key": "marking", "header": "Обозначение", "width": 25},
                        {"key": "name", "header": "Наименование", "width": 40},
                        {"key": "quantity", "header": "Кол-во", "width": 12},
                        {"key": "mass_kg", "header": "Масса, кг", "width": 12},
                        {"key": "material", "header": "Материал", "width": 25},
                        {"key": "bending", "header": "Гибка", "width": 8}
                    ],
                    "header_color": "2F5496",
                    "assembly_color": "D6E4F0"
                },
                "word": {
                    "title": "СВОДНАЯ ВЕДОМОСТЬ",
                    "columns": [
                        {"key": "position", "header": "Поз."},
                        {"key": "marking", "header": "Обозначение"},
                        {"key": "name", "header": "Наименование"},
                        {"key": "quantity", "header": "Кол-во"},
                        {"key": "mass_kg", "header": "Масса, кг"},
                        {"key": "material", "header": "Материал"},
                        {"key": "bending", "header": "Гибка"}
                    ]
                }
            }
            text.insert("1.0", json.dumps(default, ensure_ascii=False, indent=2))

        ttk.Label(dialog, text="Ключи столбцов: position, marking, name, quantity, mass_kg, total_mass, material, bending\n"
                              "Цвета: hex (например 2F5496). Ширина столбцов Excel в символах.",
                  font=("Arial", 8), foreground="gray").pack(pady=(0, 5))

        def save_template():
            try:
                content = text.get("1.0", tk.END)
                data = json.loads(content)
                name = data.get("name", "template")
                filename = name.lower().replace(" ", "_").replace("/", "_")
                if not filename.startswith("template_"):
                    filename = f"template_{filename}"
                path = os.path.join(save_dir, f"{filename}.json")
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                self._log(f"Шаблон сохранён: {name}")
                dialog.destroy()
                parent_dialog.destroy()
                self._show_template_dialog()
            except json.JSONDecodeError as e:
                messagebox.showerror("Ошибка", f"Неверный формат:\n{e}")

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Сохранить", command=save_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Отмена", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def _show_prices_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Цены и калькулятор")
        dialog.geometry("450x350")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Цены:", font=("Arial", 10)).pack(pady=(15, 5))

        prices_path = os.path.join(os.path.dirname(__file__), "prices.json")
        text = tk.Text(dialog, font=("Consolas", 9), width=50, height=15)
        text.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        if os.path.exists(prices_path):
            with open(prices_path, "r", encoding="utf-8") as f:
                text.insert("1.0", f.read())

        ttk.Label(dialog, text="materials: цена за кг по названию. default_material_price: цена для неизвестных.\n"
                              "labor: стоимость работ (cutting, bending, welding, assembly) за кг.",
                  font=("Arial", 8), foreground="gray").pack(pady=(0, 5))

        def save_prices():
            try:
                import json
                content = text.get("1.0", tk.END)
                json.loads(content)
                with open(prices_path, "w", encoding="utf-8") as f:
                    f.write(content)
                self._log("Цены сохранены")
                messagebox.showinfo("Успех", "Цены сохранены!")
            except json.JSONDecodeError as e:
                messagebox.showerror("Ошибка", f"Неверный формат:\n{e}")

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Сохранить", command=save_prices).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Отмена", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def run(self):
        self.root.mainloop()


def main():
    app = KompasExportApp()
    app.run()


if __name__ == "__main__":
    main()
