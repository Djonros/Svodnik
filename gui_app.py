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

    VERSION = "1.3.0"
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
        self.root.geometry("620x310")
        self.root.resizable(False, True)

        self.show_recent = tk.BooleanVar(value=False)
        self.show_log = tk.BooleanVar(value=False)
        self._spinner_running = False
        self._spinner_cycle = itertools.cycle(SPINNER_FRAMES)

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

        # --- Кнопка ЭКСПОРТ + Спиннер ---
        action_frame = ttk.Frame(self.root)
        action_frame.pack(pady=8)

        self.export_btn = tk.Button(
            action_frame,
            text="  ЭКСПОРТ В EXCEL / WORD  ",
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
            "в форматы Excel и Word.\n\n"
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

        self.export_btn.config(state="disabled")
        self._start_spinner()
        self._log("Начало экспорта...")

        try:
            exporter = KompasExportFinal()

            import io
            old_stdout = sys.stdout
            sys.stdout = io.StringIO()

            exporter.run(filepath)

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

    def run(self):
        self.root.mainloop()


def main():
    app = KompasExportApp()
    app.run()


if __name__ == "__main__":
    main()
