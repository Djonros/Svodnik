"""
GUI приложение для экспорта спецификаций из КОМПАС-3D v24.
Сводник - экспорт сводной ведомости из КОМПАС-3D.
"""

import sys
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

from license import LicenseManager
from kompas_export_final import KompasExportFinal


class KompasExportApp:
    """Главное окно приложения."""

    VERSION = "1.0.0"
    APP_NAME = "Сводник"

    def __init__(self):
        self.license = LicenseManager()
        self.root = tk.Tk()
        self.root.title(f"{self.APP_NAME} v{self.VERSION}")
        self.root.geometry("600x500")
        self.root.resizable(False, False)

        self._setup_ui()
        self._update_status()

    def _setup_ui(self):
        """Настройка интерфейса."""
        header = ttk.Label(self.root, text=self.APP_NAME, font=("Arial", 16, "bold"))
        header.pack(pady=5)

        subtitle = ttk.Label(self.root, text="Экспорт сводной ведомости из КОМПАС-3D", font=("Arial", 10))
        subtitle.pack(pady=(0, 5))

        self.status_label = ttk.Label(self.root, text="", font=("Arial", 10))
        self.status_label.pack(pady=(0, 5))

        file_frame = ttk.LabelFrame(self.root, text="Файл сборки", padding=10)
        file_frame.pack(fill=tk.X, padx=20, pady=5)

        self.file_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_var, width=50)
        file_entry.pack(side=tk.LEFT, padx=(0, 10))

        browse_btn = ttk.Button(file_frame, text="Обзор...", command=self._browse_file)
        browse_btn.pack(side=tk.LEFT)

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)

        export_btn = ttk.Button(btn_frame, text="Экспорт в Excel/Word", command=self._export)
        export_btn.pack(side=tk.LEFT, padx=5)

        license_btn = ttk.Button(btn_frame, text="Ввести лицензию", command=self._show_license_dialog)
        license_btn.pack(side=tk.LEFT, padx=5)

        id_btn = ttk.Button(btn_frame, text="Мой ID", command=self._show_machine_id)
        id_btn.pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.pack(fill=tk.X, padx=20, pady=5)

        log_frame = ttk.LabelFrame(self.root, text="Лог", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        self.log_text = tk.Text(log_frame, height=10, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        version_label = ttk.Label(self.root, text=f"v{self.VERSION}", font=("Arial", 8))
        version_label.pack(pady=(0, 5))

    def _update_status(self):
        status = self.license.get_status()
        self.status_label.config(text=status)
        if not self.license.is_valid():
            self.status_label.config(foreground="red")
        else:
            self.status_label.config(foreground="green")

    def _browse_file(self):
        filepath = filedialog.askopenfilename(
            title="Выберите файл сборки",
            filetypes=[("Файлы КОМПАС", "*.a3d *.a3t"), ("Все файлы", "*.*")]
        )
        if filepath:
            self.file_var.set(filepath)

    def _log(self, message):
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

        self.progress.start()
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

            self.license.record_export()
            self._update_status()

            self._log("Экспорт завершен успешно!")
            messagebox.showinfo("Готово", "Экспорт завершен!\nФайлы сохранены в папке сборки.")

        except Exception as e:
            self._log(f"Ошибка: {str(e)}")
            messagebox.showerror("Ошибка", f"Ошибка экспорта:\n{str(e)}")
        finally:
            self.progress.stop()

    def _show_machine_id(self):
        """Показать ID компьютера для передачи администратору."""
        machine_id = self.license.get_machine_id()
        dialog = tk.Toplevel(self.root)
        dialog.title("ID компьютера")
        dialog.geometry("400x150")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="ID вашего компьютера:", font=("Arial", 10)).pack(pady=10)

        id_var = tk.StringVar(value=machine_id)
        id_entry = ttk.Entry(dialog, textvariable=id_var, width=40, state="readonly")
        id_entry.pack(pady=5)

        ttk.Label(dialog, text="Отправьте этот ID администратору\nдля получения лицензионного ключа",
                  font=("Arial", 9)).pack(pady=5)

        def copy_id():
            self.root.clipboard_clear()
            self.root.clipboard_append(machine_id)
            ttk.Label(dialog, text="Скопировано!", foreground="green").pack()

        ttk.Button(dialog, text="Копировать", command=copy_id).pack(pady=5)

    def _show_license_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Активация лицензии")
        dialog.geometry("400x250")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Введите лицензионный ключ:", font=("Arial", 10)).pack(pady=10)

        key_var = tk.StringVar()
        key_entry = ttk.Entry(dialog, textvariable=key_var, width=40)
        key_entry.pack(pady=5)

        ttk.Label(dialog, text="Формат: KEY-XXXXXXXXXXXX-YYYY-MM-DD", font=("Arial", 8)).pack()

        def activate():
            key = key_var.get()
            if self.license.activate(key):
                messagebox.showinfo("Успех", "Лицензия активирована!")
                self._update_status()
                dialog.destroy()
            else:
                messagebox.showerror("Ошибка", "Неверный ключ лицензии!")

        ttk.Button(dialog, text="Активировать", command=activate).pack(pady=10)

    def run(self):
        self.root.mainloop()


def main():
    app = KompasExportApp()
    app.run()


if __name__ == "__main__":
    main()
