"""
Генератор лицензионных ключей для Сводник.
GUI утилита для администратора / распространителя.
"""

import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from license import LicenseManager


class KeygenApp:
    """GUI генератор лицензионных ключей."""

    VERSION = "1.0.0"
    APP_NAME = "Сводник — Генератор ключей"

    def __init__(self):
        self.license = LicenseManager()
        self.root = tk.Tk()
        self.root.title(f"{self.APP_NAME} v{self.VERSION}")
        self.root.geometry("520x400")
        self.root.resizable(False, False)

        self._setup_ui()

    def _setup_ui(self):
        """Настройка интерфейса."""
        header = ttk.Label(self.root, text=self.APP_NAME, font=("Arial", 14, "bold"))
        header.pack(pady=10)

        # --- ID компьютера ---
        id_frame = ttk.LabelFrame(self.root, text="ID компьютера клиента", padding=10)
        id_frame.pack(fill=tk.X, padx=20, pady=5)

        self.machine_id_var = tk.StringVar()
        id_entry = ttk.Entry(id_frame, textvariable=self.machine_id_var, width=50)
        id_entry.pack(side=tk.LEFT, padx=(0, 10))

        def get_current_id():
            self.machine_id_var.set(self.license.get_machine_id())
        ttk.Button(id_frame, text="Мой ID", command=get_current_id).pack(side=tk.LEFT)

        # --- Дата окончания ---
        date_frame = ttk.LabelFrame(self.root, text="Дата окончания лицензии", padding=10)
        date_frame.pack(fill=tk.X, padx=20, pady=5)

        ttk.Label(date_frame, text="Год:").pack(side=tk.LEFT)
        self.year_var = tk.StringVar(value=str(datetime.now().year + 1))
        ttk.Entry(date_frame, textvariable=self.year_var, width=6).pack(side=tk.LEFT, padx=2)

        ttk.Label(date_frame, text="Месяц:").pack(side=tk.LEFT, padx=(10, 0))
        self.month_var = tk.StringVar(value="06")
        ttk.Entry(date_frame, textvariable=self.month_var, width=4).pack(side=tk.LEFT, padx=2)

        ttk.Label(date_frame, text="День:").pack(side=tk.LEFT, padx=(10, 0))
        self.day_var = tk.StringVar(value="30")
        ttk.Entry(date_frame, textvariable=self.day_var, width=4).pack(side=tk.LEFT, padx=2)

        # --- Кнопка генерации ---
        ttk.Button(self.root, text="Сгенерировать ключ", command=self._generate).pack(pady=15)

        # --- Результат ---
        result_frame = ttk.LabelFrame(self.root, text="Лицензионный ключ", padding=10)
        result_frame.pack(fill=tk.X, padx=20, pady=5)

        self.key_var = tk.StringVar()
        key_entry = ttk.Entry(result_frame, textvariable=self.key_var, width=40, state="readonly",
                              font=("Consolas", 12))
        key_entry.pack(side=tk.LEFT, padx=(0, 10))

        def copy_key():
            key = self.key_var.get()
            if key:
                self.root.clipboard_clear()
                self.root.clipboard_append(key)
                messagebox.showinfo("Копирование", "Ключ скопирован в буфер обмена!")
        ttk.Button(result_frame, text="Копировать", command=copy_key).pack(side=tk.LEFT)

        # --- Статус ---
        self.status_label = ttk.Label(self.root, text="", font=("Arial", 9))
        self.status_label.pack(pady=5)

        # --- Подвал ---
        ttk.Label(self.root, text=f"v{self.VERSION}", font=("Arial", 8)).pack(pady=5)

    def _generate(self):
        """Генерация ключа."""
        machine_id = self.machine_id_var.get().strip()
        if not machine_id:
            messagebox.showwarning("Внимание", "Введите ID компьютера клиента!")
            return

        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())
            day = int(self.day_var.get())
            expiry = f"{year:04d}-{month:02d}-{day:02d}"
            datetime.strptime(expiry, "%Y-%m-%d")
        except (ValueError, TypeError):
            messagebox.showerror("Ошибка", "Некорректная дата!")
            return

        key_hash = self.license.generate_key(machine_id, expiry)
        key = f"KEY-{key_hash}-{expiry}"

        self.key_var.set(key)
        self.status_label.config(text=f"Ключ сгенерирован для {machine_id}", foreground="green")

    def run(self):
        self.root.mainloop()


def main():
    app = KeygenApp()
    app.run()


if __name__ == "__main__":
    main()
