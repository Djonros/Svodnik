"""
GUI приложение для экспорта спецификаций из КОМПАС-3D v24.
Сводник - экспорт сводной ведомости из КОМПАС-3D.
"""

import sys
import os
import json
import queue
import threading
import time
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

from license import LicenseManager
from kompas_export_final import KompasExportFinal
import updater


RECENT_FILE = os.path.join(os.path.expanduser("~"), ".svodnik_recent.json")
MAX_RECENT = 10


class _QueueWriter:
    """Подмена stdout на время экспорта: печать из фонового потока
    построчно передается в очередь, окно выводит ее в лог по ходу работы."""

    def __init__(self, q):
        self._queue = q
        self._buffer = ""

    def write(self, text):
        self._buffer += text
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            self._queue.put(("out", line))
        return len(text)

    def flush(self):
        if self._buffer:
            self._queue.put(("out", self._buffer))
            self._buffer = ""


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


SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".svodnik_settings.json")

# Цвета оформления
COLOR_BG = "#F3F5F9"
COLOR_CARD = "#FFFFFF"
COLOR_BORDER = "#D9DEE7"
COLOR_ACCENT = "#2F5496"
COLOR_ACCENT_DARK = "#1F3D7A"
COLOR_HEADER = "#1F3864"
COLOR_TEXT = "#1E2430"
COLOR_MUTED = "#6B7385"
COLOR_OK = "#2E7D32"
COLOR_WARN = "#B26A00"
COLOR_ERR = "#C62828"
FONT = "Segoe UI"

# Этапы экспорта: (текст, доля прогресса). Этап определяется по строкам лога ядра.
STAGES = [
    ("Подключено к КОМПАС", "Чтение сборки...", 15),
    ("Найдено элементов", "Чтение спецификаций...", 45),
    ("Создание спецификации", "Создание недостающих спецификаций...", 55),
    ("Сверка количества", "Подготовка данных...", 70),
    ("Создание Excel", "Формирование Excel...", 80),
    ("Создание Word", "Формирование Word...", 88),
    ("Создание PDF", "Формирование PDF...", 94),
]


def load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(data):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def open_path(path):
    """Открыть файл или папку программой Windows по умолчанию."""
    try:
        os.startfile(path)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть:\n{path}\n\n{e}")


class Tooltip:
    """Всплывающая подсказка при наведении на элемент."""

    def __init__(self, widget, text, delay=500):
        self.widget, self.text, self.delay = widget, text, delay
        self._after = None
        self._tip = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._after = self.widget.after(self.delay, self._show)

    def _show(self):
        if self._tip:
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self._tip, text=self.text, justify=tk.LEFT, background="#FFFDE8",
                 foreground=COLOR_TEXT, relief="solid", borderwidth=1,
                 font=(FONT, 9), padx=6, pady=3, wraplength=360).pack()

    def _hide(self, _event=None):
        if self._after:
            self.widget.after_cancel(self._after)
            self._after = None
        if self._tip:
            self._tip.destroy()
            self._tip = None


class KompasExportApp:

    VERSION = "1.8.0"
    APP_NAME = "Сводник"

    def __init__(self):
        self.license = LicenseManager()
        # Четкий текст на мониторах с масштабированием Windows
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        try:
            from tkinterdnd2 import TkinterDnD
            self.root = TkinterDnD.Tk()
            self._dnd_available = True
        except Exception:
            self.root = tk.Tk()
            self._dnd_available = False
        self.root.title(f"{self.APP_NAME} {self.VERSION}")
        self.root.configure(background=COLOR_BG)

        self.settings = load_settings()
        s = self.settings
        self.show_log = tk.BooleanVar(value=s.get("show_log", False))
        self.export_excel = tk.BooleanVar(value=s.get("excel", True))
        self.export_word = tk.BooleanVar(value=s.get("word", False))
        self.export_pdf = tk.BooleanVar(value=s.get("pdf", False))
        self.export_costs = tk.BooleanVar(value=s.get("costs", False))
        self.edit_before_export = tk.BooleanVar(value=s.get("edit_rows", False))
        self.open_after_export = tk.BooleanVar(value=s.get("open_after", False))
        self.auto_update = tk.BooleanVar(value=s.get("auto_update", True))
        self._current_template = None
        self._template_name = "По умолчанию"
        self._busy = False
        self._export_ctx = None
        self._last_results = {}
        self._last_issues = []

        self._setup_style()
        self._set_icon()
        self._setup_menu()
        self._setup_ui()
        self._setup_shortcuts()
        self._update_status()
        self._load_recent_list()
        self._setup_dragdrop()
        self._rebuild_panels()
        scale = max(1.0, self.root.winfo_fpixels("1i") / 96.0)
        self.root.minsize(int(640 * scale), int(430 * scale))
        geometry = f"{int(780 * scale)}x{int(600 * scale)}"
        saved = s.get("geometry", "")
        try:
            w, h = (int(v) for v in saved.split("+")[0].split("x"))
            if w <= self.root.winfo_screenwidth() * 0.95 and h <= self.root.winfo_screenheight() * 0.9:
                geometry = saved
        except ValueError:
            pass
        self.root.geometry(geometry)
        if s.get("zoomed"):
            self.root.state("zoomed")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(2000, self._auto_check_updates)

    # ------------------------------------------------------------------ оформление

    def _setup_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        from tkinter import font as tkfont
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(name).configure(family=FONT, size=10)
            except tk.TclError:
                pass
        # Флажки темы clam мелкие на мониторах с масштабированием: увеличиваем
        scale = max(1.0, self.root.winfo_fpixels("1i") / 96.0)
        style.configure("TCheckbutton", indicatorsize=int(13 * scale), indicatormargin=(0, 0, 6, 0))
        style.configure(".", background=COLOR_BG, foreground=COLOR_TEXT, font=(FONT, 10))
        style.configure("TFrame", background=COLOR_BG)
        style.configure("Card.TFrame", background=COLOR_CARD)
        style.configure("Header.TFrame", background=COLOR_HEADER)
        style.configure("TLabel", background=COLOR_BG)
        style.configure("Card.TLabel", background=COLOR_CARD)
        style.configure("CardTitle.TLabel", background=COLOR_CARD, font=(FONT, 11, "bold"))
        style.configure("Muted.TLabel", background=COLOR_CARD, foreground=COLOR_MUTED, font=(FONT, 9))
        style.configure("Status.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=(FONT, 9))
        style.configure("HeaderTitle.TLabel", background=COLOR_HEADER, foreground="white",
                        font=(FONT, 18, "bold"))
        style.configure("HeaderSub.TLabel", background=COLOR_HEADER, foreground="#C9D4EA",
                        font=(FONT, 9))
        style.configure("TCheckbutton", background=COLOR_CARD)
        style.map("TCheckbutton", background=[("active", COLOR_CARD)])
        style.configure("TButton", padding=(10, 4))
        style.configure("Accent.TButton", background=COLOR_ACCENT, foreground="white",
                        font=(FONT, 11, "bold"), padding=(24, 8), borderwidth=0)
        style.map("Accent.TButton",
                  background=[("disabled", "#9AA8C4"), ("active", COLOR_ACCENT_DARK)],
                  foreground=[("disabled", "#EEF1F7")])
        style.configure("Link.TButton", background=COLOR_CARD, foreground=COLOR_ACCENT,
                        borderwidth=0, padding=(4, 2))
        style.map("Link.TButton", background=[("active", "#EEF2FA")])
        style.configure("Horizontal.TProgressbar", troughcolor="#E3E7EF",
                        background=COLOR_ACCENT, bordercolor=COLOR_BORDER,
                        lightcolor=COLOR_ACCENT, darkcolor=COLOR_ACCENT)
        style.configure("TEntry", padding=4)
        style.configure("TCombobox", padding=4)

    def _set_icon(self):
        """Значок окна: синий лист с строками ведомости (вместо пера Tk)."""
        try:
            size = 32
            img = tk.PhotoImage(width=size, height=size)
            img.put(COLOR_ACCENT, to=(0, 0, size, size))
            img.put("#FFFFFF", to=(7, 5, 25, 27))
            for y in (9, 14, 19):
                img.put(COLOR_ACCENT, to=(10, y, 22, y + 2))
            img.put(COLOR_ACCENT, to=(10, 23, 17, 25))
            self._icon = img
            self.root.iconphoto(True, img)
        except tk.TclError:
            pass

    def _card(self, parent, title=None):
        """Белая карточка с рамкой и заголовком; возвращает внутреннюю рамку."""
        outer = tk.Frame(parent, background=COLOR_CARD, highlightthickness=1,
                         highlightbackground=COLOR_BORDER)
        outer.pack(fill=tk.X, padx=16, pady=(0, 10))
        inner = ttk.Frame(outer, style="Card.TFrame", padding=(14, 10))
        inner.pack(fill=tk.BOTH, expand=True)
        if title:
            ttk.Label(inner, text=title, style="CardTitle.TLabel").pack(anchor="w", pady=(0, 6))
        return outer, inner

    # ------------------------------------------------------------------ окно

    def _on_close(self):
        if self._busy and not messagebox.askyesno(
                "Экспорт не завершен",
                "Экспорт еще выполняется. Закрыть программу?"):
            return
        self._save_settings()
        self.root.destroy()

    def _save_settings(self):
        zoomed = self.root.state() == "zoomed"
        if not zoomed:
            self.settings["geometry"] = self.root.geometry()
        self.settings.update({
            "zoomed": zoomed,
            "show_log": self.show_log.get(),
            "excel": self.export_excel.get(),
            "word": self.export_word.get(),
            "pdf": self.export_pdf.get(),
            "costs": self.export_costs.get(),
            "edit_rows": self.edit_before_export.get(),
            "open_after": self.open_after_export.get(),
            "auto_update": self.auto_update.get(),
        })
        save_settings(self.settings)

    def _setup_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Открыть сборку...", accelerator="Ctrl+O",
                              command=self._browse_file)
        file_menu.add_command(label="Взять активную сборку из КОМПАС", accelerator="Ctrl+K",
                              command=self._take_from_kompas)
        self.recent_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="Последние файлы", menu=self.recent_menu)
        file_menu.add_separator()
        file_menu.add_command(label="Сформировать ведомость", accelerator="F5",
                              command=self._export)
        file_menu.add_command(label="Открыть папку сборки", command=self._open_assembly_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self._on_close)
        menubar.add_cascade(label="Файл", menu=file_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_checkbutton(label="Журнал", variable=self.show_log, accelerator="Ctrl+L",
                                  command=self._rebuild_panels)
        menubar.add_cascade(label="Вид", menu=view_menu)

        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="Шаблоны...", command=self._show_template_dialog)
        settings_menu.add_command(label="Цены и калькулятор...", command=self._show_prices_dialog)
        settings_menu.add_separator()
        settings_menu.add_command(label="Ввести лицензию", command=self._show_license_dialog)
        settings_menu.add_command(label="Мой ID компьютера", command=self._show_machine_id)
        menubar.add_cascade(label="Настройки", menu=settings_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Проверить обновления...", command=self._check_updates)
        help_menu.add_checkbutton(label="Проверять обновления при запуске",
                                  variable=self.auto_update)
        help_menu.add_separator()
        help_menu.add_command(label="Горячие клавиши", command=self._show_shortcuts)
        help_menu.add_command(label="О программе", command=self._show_about)
        menubar.add_cascade(label="Справка", menu=help_menu)

    def _setup_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self._browse_file())
        self.root.bind("<Control-O>", lambda e: self._browse_file())
        self.root.bind("<Control-k>", lambda e: self._take_from_kompas())
        self.root.bind("<Control-K>", lambda e: self._take_from_kompas())
        self.root.bind("<F5>", lambda e: self._export())
        self.root.bind("<Control-Return>", lambda e: self._export())
        self.root.bind("<Control-l>", lambda e: self._toggle_log())
        self.root.bind("<Control-L>", lambda e: self._toggle_log())

    def _setup_ui(self):
        # --- Шапка ---
        header = ttk.Frame(self.root, style="Header.TFrame", padding=(18, 12))
        header.pack(fill=tk.X)
        left = ttk.Frame(header, style="Header.TFrame")
        left.pack(side=tk.LEFT)
        ttk.Label(left, text=self.APP_NAME, style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(left, text=f"Сводная ведомость из КОМПАС-3D  ·  версия {self.VERSION}",
                  style="HeaderSub.TLabel").pack(anchor="w")
        self.status_label = tk.Label(header, text="", font=(FONT, 9, "bold"),
                                     foreground="white", background=COLOR_OK, padx=10, pady=3,
                                     cursor="hand2")
        self.status_label.pack(side=tk.RIGHT)
        self.status_label.bind("<Button-1>", lambda e: self._show_license_dialog())
        Tooltip(self.status_label, "Состояние лицензии. Нажмите, чтобы ввести ключ.")

        body = ttk.Frame(self.root, padding=(0, 12, 0, 0))
        body.pack(fill=tk.BOTH, expand=True)
        self.body = body

        # --- Сборка ---
        _, card = self._card(body, "Сборка")
        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill=tk.X)
        self.file_var = tk.StringVar()
        self.file_combo = ttk.Combobox(row, textvariable=self.file_var, font=(FONT, 10))
        self.file_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        Tooltip(self.file_combo, "Путь к файлу сборки .a3d. Раскройте список, "
                                 "чтобы выбрать один из последних файлов.")
        btn = ttk.Button(row, text="Обзор...", command=self._browse_file)
        btn.pack(side=tk.LEFT, padx=(0, 6))
        Tooltip(btn, "Выбрать файл сборки (Ctrl+O)")
        btn = ttk.Button(row, text="Из КОМПАС", command=self._take_from_kompas)
        btn.pack(side=tk.LEFT)
        Tooltip(btn, "Взять сборку, открытую сейчас в КОМПАС-3D (Ctrl+K)")
        hint = ("Можно перетащить файл .a3d в окно. " if self._dnd_available else "") + \
            "Документы сохраняются в папку сборки."
        ttk.Label(card, text=hint, style="Muted.TLabel").pack(anchor="w", pady=(6, 0))

        # --- Что сформировать ---
        _, card = self._card(body, "Что сформировать")
        grid = ttk.Frame(card, style="Card.TFrame")
        grid.pack(fill=tk.X)
        ttk.Label(grid, text="Форматы:", style="Card.TLabel").grid(row=0, column=0, sticky="w",
                                                                  padx=(0, 12), pady=2)
        for col, (text, var, tip) in enumerate((
                ("Excel", self.export_excel, "Ведомость в Excel (.xlsx)"),
                ("Word", self.export_word, "Ведомость в Word (.docx)"),
                ("PDF", self.export_pdf, "Ведомость в PDF"),
        ), start=1):
            cb = ttk.Checkbutton(grid, text=text, variable=var)
            cb.grid(row=0, column=col, sticky="w", padx=(0, 14), pady=2)
            Tooltip(cb, tip)
        ttk.Label(grid, text="Дополнительно:", style="Card.TLabel").grid(row=1, column=0, sticky="w",
                                                                        padx=(0, 12), pady=2)
        for col, (text, var, tip) in enumerate((
                ("Расчет стоимости", self.export_costs,
                 "Добавить в Excel лист со стоимостью материалов и работ "
                 "(цены: Настройки → Цены и калькулятор)"),
                ("Проверить строки перед сохранением", self.edit_before_export,
                 "Открыть редактор строк: можно исключить строки, поправить гибку и длину"),
        ), start=1):
            cb = ttk.Checkbutton(grid, text=text, variable=var)
            cb.grid(row=1, column=col, columnspan=2 if col == 2 else 1, sticky="w",
                    padx=(0, 14), pady=2)
            Tooltip(cb, tip)
        cb = ttk.Checkbutton(grid, text="Открыть файл после экспорта", variable=self.open_after_export)
        cb.grid(row=2, column=1, columnspan=3, sticky="w", pady=2)
        tmpl_row = ttk.Frame(card, style="Card.TFrame")
        tmpl_row.pack(fill=tk.X, pady=(6, 0))
        self.template_label = ttk.Label(tmpl_row, text="", style="Muted.TLabel")
        self.template_label.pack(side=tk.LEFT)
        ttk.Button(tmpl_row, text="Изменить шаблон", style="Link.TButton",
                   command=self._show_template_dialog).pack(side=tk.LEFT, padx=(6, 0))
        self._set_template_label(self._template_name)

        # --- Запуск и прогресс ---
        action = ttk.Frame(body, padding=(16, 0, 16, 10))
        action.pack(fill=tk.X)
        self.export_btn = ttk.Button(action, text="Сформировать ведомость",
                                     style="Accent.TButton", command=self._export)
        self.export_btn.pack(side=tk.LEFT)
        Tooltip(self.export_btn, "Прочитать сборку из КОМПАС и сохранить документы (F5)")
        prog = ttk.Frame(action)
        prog.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(16, 0))
        self.stage_var = tk.StringVar(value="Готово к работе")
        ttk.Label(prog, textvariable=self.stage_var, style="Status.TLabel").pack(anchor="w")
        self.progress = ttk.Progressbar(prog, mode="determinate", maximum=100)
        self.progress.pack(fill=tk.X, pady=(3, 0))

        # --- Результат (появляется после экспорта) ---
        self.result_outer, self.result_card = self._card(body, None)
        self.result_outer.pack_forget()
        self.result_title = ttk.Label(self.result_card, text="", style="CardTitle.TLabel")
        self.result_title.pack(anchor="w")
        self.result_text = ttk.Label(self.result_card, text="", style="Muted.TLabel",
                                     justify=tk.LEFT, wraplength=680)
        self.result_text.pack(anchor="w", pady=(2, 6))
        self.result_buttons = ttk.Frame(self.result_card, style="Card.TFrame")
        self.result_buttons.pack(fill=tk.X)

        # --- Журнал ---
        self.log_frame = ttk.Frame(body, padding=(16, 0, 16, 6))
        log_box = tk.Frame(self.log_frame, background=COLOR_CARD, highlightthickness=1,
                           highlightbackground=COLOR_BORDER)
        log_box.pack(fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(log_box)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text = tk.Text(log_box, height=6, state=tk.DISABLED, font=("Consolas", 9),
                                relief="flat", background=COLOR_CARD, foreground=COLOR_TEXT,
                                yscrollcommand=scrollbar.set, padx=8, pady=6)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)
        self.log_text.tag_configure("err", foreground=COLOR_ERR)
        self.log_text.tag_configure("warn", foreground=COLOR_WARN)
        self.log_text.tag_configure("ok", foreground=COLOR_OK)

        # --- Строка состояния ---
        status = ttk.Frame(self.root, padding=(16, 4))
        status.pack(fill=tk.X, side=tk.BOTTOM, before=body)
        self.log_toggle = ttk.Button(status, text="", command=self._toggle_log)
        self.log_toggle.pack(side=tk.LEFT)
        ttk.Label(status, text="F5 — сформировать   Ctrl+O — открыть   Ctrl+K — из КОМПАС",
                  style="Status.TLabel").pack(side=tk.RIGHT)

    def _set_template_label(self, name):
        self._template_name = name
        self.template_label.config(text=f"Шаблон столбцов: {name}")

    def _toggle_log(self):
        self.show_log.set(not self.show_log.get())
        self._rebuild_panels()

    def _rebuild_panels(self):
        self.log_frame.pack_forget()
        if self.show_log.get():
            self.log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_toggle.config(text=("Скрыть журнал" if self.show_log.get() else "Показать журнал"))

    def _show_about(self):
        messagebox.showinfo("О программе",
            f"{self.APP_NAME} {self.VERSION}\n\n"
            "Сводная ведомость из сборки КОМПАС-3D v24\n"
            "в Excel, Word и PDF: позиции, количество, масса,\n"
            "материал, длина сортамента, гибка.\n\n"
            "GitHub: github.com/Djonros/Svodnik")

    # ------------------------------------------------------------------ обновления

    def _auto_check_updates(self):
        if not self.auto_update.get():
            return
        last = self.settings.get("update_checked", 0)
        if time.time() - last < 12 * 3600:
            return
        self._check_updates(manual=False)

    def _check_updates(self, manual=True):
        if getattr(self, "_update_checking", False):
            return
        self._update_checking = True
        result = {}

        def work():
            try:
                result["info"] = updater.check_latest()
            except Exception as e:
                result["error"] = e
            result["done"] = True

        def poll():
            if not result.get("done"):
                self.root.after(300, poll)
                return
            self._update_checking = False
            self._on_update_result(result.get("info"), result.get("error"), manual)

        threading.Thread(target=work, daemon=True).start()
        poll()

    def _on_update_result(self, info, error, manual):
        if error is not None:
            if manual:
                messagebox.showerror("Обновления",
                    f"Не удалось проверить обновления.\nПроверьте подключение к интернету.\n\n{error}")
            return
        self.settings["update_checked"] = time.time()
        if not info or not updater.is_newer(info["version"], self.VERSION):
            if manual:
                messagebox.showinfo("Обновления", f"У вас последняя версия ({self.VERSION}).")
            return
        if not manual and self.settings.get("skip_version") == info["version"]:
            return
        self._show_update_dialog(info)

    def _show_update_dialog(self, info):
        dialog = tk.Toplevel(self.root)
        dialog.title("Доступно обновление")
        dialog.configure(background=COLOR_BG)
        dialog.transient(self.root)
        dialog.resizable(False, False)
        body = ttk.Frame(dialog, padding=16)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, text=f"Вышла версия {info['version']}",
                  font=(FONT, 13, "bold")).pack(anchor="w")
        ttk.Label(body, text=f"У вас установлена версия {self.VERSION}.",
                  foreground=COLOR_MUTED).pack(anchor="w", pady=(2, 10))
        if info.get("notes"):
            ttk.Label(body, text="Что нового:").pack(anchor="w")
            notes = tk.Text(body, width=60, height=10, wrap="word", font=(FONT, 10),
                            relief="solid", borderwidth=1, background=COLOR_CARD)
            notes.insert("1.0", info["notes"])
            notes.config(state="disabled")
            notes.pack(fill=tk.BOTH, pady=(2, 10))
        progress = ttk.Progressbar(body, mode="determinate", maximum=100)
        progress_label = ttk.Label(body, text="", foreground=COLOR_MUTED)
        ttk.Checkbutton(body, text="Проверять обновления при запуске",
                        variable=self.auto_update).pack(anchor="w", pady=(0, 10))
        buttons = ttk.Frame(body)
        buttons.pack(fill=tk.X)

        self_update = updater.can_self_update() and info.get("asset_url")

        def skip():
            self.settings["skip_version"] = info["version"]
            dialog.destroy()

        def start():
            if not self_update:
                webbrowser.open(info["page"])
                dialog.destroy()
                return
            if self._busy:
                messagebox.showwarning("Обновление",
                    "Дождитесь окончания экспорта, потом обновите программу.", parent=dialog)
                return
            for child in buttons.winfo_children():
                child.config(state="disabled")
            progress.pack(fill=tk.X, pady=(0, 4), before=buttons)
            progress_label.pack(anchor="w", pady=(0, 10), before=buttons)
            self._download_update(info, dialog, progress, progress_label, buttons)

        ttk.Button(buttons, text="Обновить сейчас" if self_update else "Открыть страницу загрузки",
                   style="Accent.TButton", command=start).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="Позже", command=dialog.destroy).pack(side=tk.RIGHT, padx=6)
        ttk.Button(buttons, text="Пропустить версию", command=skip).pack(side=tk.LEFT)
        dialog.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dialog.winfo_height()) // 3
        dialog.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _download_update(self, info, dialog, progress, progress_label, buttons):
        dest = updater.temp_download_path()
        state = {"done": 0, "total": info.get("size") or 0}

        def report(done, total):
            state["done"], state["total"] = done, total or state["total"]

        def work():
            try:
                updater.download(info["asset_url"], dest, report)
            except Exception as e:
                state["error"] = e
            state["finished"] = True

        def poll():
            if not dialog.winfo_exists():
                return
            total = state["total"]
            mb = state["done"] / 1048576
            if total:
                progress["value"] = state["done"] * 100 / total
                progress_label.config(text=f"Загрузка: {mb:.1f} из {total / 1048576:.1f} МБ")
            else:
                progress_label.config(text=f"Загрузка: {mb:.1f} МБ")
            if not state.get("finished"):
                dialog.after(200, poll)
                return
            if "error" in state:
                progress_label.config(text="")
                for child in buttons.winfo_children():
                    child.config(state="normal")
                messagebox.showerror("Обновление",
                    f"Не удалось скачать обновление:\n{state['error']}", parent=dialog)
                return
            progress_label.config(text="Загружено. Программа перезапустится.")
            dialog.update_idletasks()
            try:
                updater.install_and_restart(dest)
            except Exception as e:
                messagebox.showerror("Обновление", f"Не удалось установить обновление:\n{e}",
                                     parent=dialog)
                return
            self._save_settings()
            self.root.destroy()

        threading.Thread(target=work, daemon=True).start()
        poll()

    def _show_shortcuts(self):
        messagebox.showinfo("Горячие клавиши",
            "F5 или Ctrl+Enter — сформировать ведомость\n"
            "Ctrl+O — открыть файл сборки\n"
            "Ctrl+K — взять активную сборку из КОМПАС\n"
            "Ctrl+L — показать или скрыть журнал")

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
        color = COLOR_OK if self.license.is_valid() else COLOR_ERR
        if self.license.is_valid() and "проб" in status.lower():
            color = COLOR_WARN
        self.status_label.config(text=status, background=color)

    # ------------------------------------------------------------------ файлы

    def _load_recent_list(self):
        files = load_recent()
        self.file_combo["values"] = files
        self.recent_menu.delete(0, tk.END)
        if not files:
            self.recent_menu.add_command(label="(пусто)", state="disabled")
            return
        for path in files:
            label = path if len(path) <= 70 else "..." + path[-67:]
            self.recent_menu.add_command(label=label, command=lambda p=path: self._select_recent(p))
        self.recent_menu.add_separator()
        self.recent_menu.add_command(label="Очистить список", command=self._clear_recent)
        if not self.file_var.get():
            self.file_var.set(files[0])

    def _select_recent(self, path):
        if os.path.exists(path):
            self.file_var.set(path)
            return
        messagebox.showwarning("Внимание", f"Файл не найден:\n{path}")
        files = [f for f in load_recent() if f != path]
        save_recent(files)
        self._load_recent_list()

    def _clear_recent(self):
        if messagebox.askyesno("Очистка", "Очистить список последних файлов?"):
            save_recent([])
            self._load_recent_list()

    def _browse_file(self):
        current = self.file_var.get()
        filepath = filedialog.askopenfilename(
            title="Выберите файл сборки",
            initialdir=os.path.dirname(current) if current else None,
            filetypes=[("Сборки КОМПАС", "*.a3d *.a3t"), ("Все файлы", "*.*")]
        )
        if filepath:
            self.file_var.set(os.path.normpath(filepath))

    def _take_from_kompas(self):
        """Путь к сборке, открытой сейчас в КОМПАС-3D."""
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            app = win32com.client.GetActiveObject("Kompas.Application.7")
            doc = app.ActiveDocument
            path = doc.PathName if doc else ""
        except Exception:
            messagebox.showwarning("КОМПАС", "КОМПАС-3D не запущен или в нем нет открытого документа.")
            return
        if not path:
            messagebox.showwarning("КОМПАС", "Активный документ КОМПАС еще не сохранен в файл.")
        elif not path.lower().endswith((".a3d", ".a3t")):
            messagebox.showwarning("КОМПАС", f"Активный документ не сборка:\n{os.path.basename(path)}")
        else:
            self.file_var.set(os.path.normpath(path))

    def _open_assembly_folder(self):
        path = self.file_var.get()
        folder = os.path.dirname(path) if path else ""
        if folder and os.path.isdir(folder):
            open_path(folder)
        else:
            messagebox.showwarning("Внимание", "Сначала выберите файл сборки.")

    # ------------------------------------------------------------------ журнал и прогресс

    def _log(self, message):
        tag = ""
        head = message.lstrip()
        if head.startswith(("[!!]", "Ошибка")):
            tag = "err"
        elif head.startswith(("[??]", "Внимание")):
            tag = "warn"
        elif head.startswith("[OK]"):
            tag = "ok"
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')}  {message}\n", tag)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self._track_stage(message)

    def _log_output(self, output):
        for line in output.strip().split("\n"):
            if line.strip():
                self._log(line)

    def _set_stage(self, text, value=None):
        self.stage_var.set(text)
        if value is not None:
            self.progress["value"] = value

    def _track_stage(self, message):
        if not self._busy:
            return
        for marker, text, value in STAGES:
            if marker in message and value >= self.progress["value"]:
                self._set_stage(text, value)
                break

    def _run_in_background(self, job, on_done):
        """Выполнение job в фоновом потоке, чтобы окно не зависало.
        Печать job построчно попадает в лог по ходу работы;
        on_done(result, error) вызывается в главном потоке."""
        q = queue.Queue()
        writer = _QueueWriter(q)
        old_stdout = sys.stdout
        sys.stdout = writer

        def worker():
            try:
                result, error = job(), None
            except Exception as e:
                result, error = None, e
            writer.flush()
            q.put(("done", result, error))

        threading.Thread(target=worker, daemon=True).start()

        def poll():
            try:
                while True:
                    kind, *payload = q.get_nowait()
                    if kind == "out":
                        if payload[0].strip():
                            self._log(payload[0])
                    else:
                        sys.stdout = old_stdout
                        on_done(*payload)
                        return
            except queue.Empty:
                pass
            self.root.after(100, poll)

        poll()

    # ------------------------------------------------------------------ экспорт

    def _export(self):
        if self._busy:
            return
        if not self.license.is_valid():
            remaining = self.license.MAX_EXPORTS_TRIAL - self.license.exports_count
            if remaining > 0:
                messagebox.showwarning("Пробный режим",
                    f"Осталось {remaining} экспортов.\nВведите лицензию для полной версии.")
            else:
                messagebox.showerror("Ошибка",
                    "Пробный период закончился!\nВведите лицензионный ключ.")
            return

        filepath = self.file_var.get().strip().strip('"')
        if not filepath:
            messagebox.showwarning("Внимание", "Выберите файл сборки: «Обзор...» или «Из КОМПАС».")
            return
        if not os.path.exists(filepath):
            messagebox.showerror("Ошибка", f"Файл не найден:\n{filepath}")
            return

        formats = set()
        if self.export_excel.get():
            formats.add("excel")
        if self.export_word.get():
            formats.add("word")
        if self.export_pdf.get():
            formats.add("pdf")
        if not formats:
            messagebox.showwarning("Внимание", "Отметьте хотя бы один формат: Excel, Word или PDF.")
            return

        self._save_settings()
        self.export_btn.config(state="disabled")
        self._busy = True
        self.result_outer.pack_forget()
        self._set_stage("Подключение к КОМПАС...", 5)
        fmt_list = ", ".join(sorted(formats)).upper()
        self._log(f"Начало экспорта ({fmt_list}): {filepath}")

        self._export_ctx = {
            "exporter": KompasExportFinal(template=self._current_template),
            "filepath": filepath,
            "formats": formats,
            "costs": self.export_costs.get() and "excel" in formats,
        }
        self._run_in_background(self._analyze_job, self._after_analyze)

    def _analyze_job(self):
        """Фоновый поток: чтение сборки и спецификаций из КОМПАС.
        Вся работа с COM (включая закрытие документов) идет в этом потоке."""
        import pythoncom
        exporter = self._export_ctx["exporter"]
        pythoncom.CoInitialize()
        try:
            return exporter.analyze(self._export_ctx["filepath"])
        finally:
            exporter.close_all_docs()
            exporter.app = exporter.doc = exporter.doc3d = exporter.top_part = None
            pythoncom.CoUninitialize()

    def _after_analyze(self, ok, error):
        if error is None and not ok:
            error = RuntimeError("Не удалось подключиться к КОМПАС "
                                 "или открыть сборку (подробности в журнале)")
        if error is not None:
            self._finish_export(error)
            return

        if self.edit_before_export.get():
            self._set_stage("Проверка строк в редакторе...")
            from row_editor import RowEditorDialog
            editor = RowEditorDialog(self.root, self._export_ctx["exporter"].all_data)
            self.root.wait_window(editor)
            if not editor.applied:
                self._log("Экспорт отменен в редакторе строк")
                self._finish_export()
                self._set_stage("Экспорт отменен", 0)
                return

        self._set_stage("Формирование документов...", 75)
        self._run_in_background(self._generate_job, self._after_generate)

    def _generate_job(self):
        """Фоновый поток: формирование документов (КОМПАС уже не нужен)."""
        exporter = self._export_ctx["exporter"]
        if self._export_ctx["costs"]:
            exporter.calculate_costs()
        return exporter.generate(self._export_ctx["formats"])

    def _after_generate(self, result, error):
        if error is not None:
            self._finish_export(error)
            return

        exporter = self._export_ctx["exporter"]
        add_to_recent(self._export_ctx["filepath"])
        self._load_recent_list()

        self.license.record_export()
        self._update_status()

        self._log("[OK] Экспорт завершен")
        self._finish_export()
        self._set_stage("Готово", 100)
        self._show_result(exporter, result or {})

        if self.open_after_export.get() and result:
            first = result.get("excel") or next(iter(result.values()), None)
            if first and os.path.exists(first):
                open_path(first)

    def _show_result(self, exporter, results):
        """Карточка с итогом: файлы, кнопки открытия, проблемы."""
        self._last_results = results
        problems = list(exporter.problems)
        warnings = list(exporter.quantity_warnings)
        generated = list(getattr(exporter, "generated_specs", []))
        parts = [it for it in exporter.all_data if not it.get("is_assembly")]
        bent = sum(1 for it in parts if it.get("is_bending"))

        issues = []
        if problems:
            issues.append(("Ошибки чтения из КОМПАС, ведомость может быть неполной", problems))
        if warnings:
            issues.append(("Количество в модели не совпадает со спецификацией", warnings))
        if generated:
            issues.append(("Спецификации не было, созданы автоматически "
                           "(папка «Генерированные спецификации»)",
                           [os.path.basename(p) for p in generated]))
        self._last_issues = issues

        n_bad = len(problems) + len(warnings)
        if n_bad:
            self.result_title.config(text=f"Готово, проверьте ведомость: замечаний {n_bad}",
                                     foreground=COLOR_WARN)
        else:
            self.result_title.config(text="Готово", foreground=COLOR_OK)
        lines = [f"Деталей и сборок: {len(exporter.all_data)}, из них гнутых деталей: {bent}."]
        if generated:
            lines.append(f"Создано спецификаций: {len(generated)}.")
        lines.append("Сохранено: " + ", ".join(os.path.basename(p) for p in results.values()))
        self.result_text.config(text="\n".join(lines))

        for w in self.result_buttons.winfo_children():
            w.destroy()
        names = {"excel": "Открыть Excel", "word": "Открыть Word", "pdf": "Открыть PDF"}
        for fmt in ("excel", "word", "pdf"):
            path = results.get(fmt)
            if path:
                ttk.Button(self.result_buttons, text=names[fmt],
                           command=lambda p=path: open_path(p)).pack(side=tk.LEFT, padx=(0, 6))
        folder = os.path.dirname(next(iter(results.values()), "") or "")
        if folder:
            ttk.Button(self.result_buttons, text="Открыть папку",
                       command=lambda: open_path(folder)).pack(side=tk.LEFT, padx=(0, 6))
        if issues:
            ttk.Button(self.result_buttons, text=f"Замечания ({n_bad + len(generated)})",
                       command=self._show_issues).pack(side=tk.LEFT, padx=(0, 6))
        if self.log_frame.winfo_manager():
            self.result_outer.pack(fill=tk.X, padx=16, pady=(0, 10), before=self.log_frame)
        else:
            self.result_outer.pack(fill=tk.X, padx=16, pady=(0, 10))
        if n_bad:
            self._show_issues()

    def _show_issues(self):
        """Окно со всеми замечаниями экспорта (с прокруткой и копированием)."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Замечания по ведомости")
        dialog.geometry("720x420")
        dialog.configure(background=COLOR_BG)
        dialog.transient(self.root)
        box = tk.Frame(dialog, background=COLOR_CARD, highlightthickness=1,
                       highlightbackground=COLOR_BORDER)
        box.pack(fill=tk.BOTH, expand=True, padx=12, pady=(12, 6))
        sb = ttk.Scrollbar(box)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        text = tk.Text(box, wrap="word", relief="flat", background=COLOR_CARD,
                       font=(FONT, 10), padx=10, pady=8, yscrollcommand=sb.set)
        text.pack(fill=tk.BOTH, expand=True)
        sb.config(command=text.yview)
        text.tag_configure("h", font=(FONT, 10, "bold"), spacing1=6, spacing3=4)
        for title, items in self._last_issues:
            text.insert(tk.END, f"{title} ({len(items)})\n", "h")
            for item in items:
                text.insert(tk.END, f"•  {item}\n")
        text.config(state=tk.DISABLED)

        def copy_all():
            self.root.clipboard_clear()
            self.root.clipboard_append(text.get("1.0", tk.END))

        btns = ttk.Frame(dialog, padding=(12, 0, 12, 12))
        btns.pack(fill=tk.X)
        ttk.Button(btns, text="Закрыть", command=dialog.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Копировать", command=copy_all).pack(side=tk.RIGHT, padx=6)

    def _finish_export(self, error=None):
        self.export_btn.config(state="normal")
        self._busy = False
        if error is not None:
            self._set_stage("Ошибка, подробности в журнале", 0)
            self._log(f"Ошибка: {error}")
            if not self.show_log.get():
                self._toggle_log()
            messagebox.showerror("Ошибка", f"Ошибка экспорта:\n{error}")

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
            self._set_template_label(name)
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
                        {"key": "stock_length", "header": "Длина сортамента", "width": 12},
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
                        {"key": "stock_length", "header": "Длина сортамента"},
                        {"key": "bending", "header": "Гибка"}
                    ]
                }
            }
            text.insert("1.0", json.dumps(default, ensure_ascii=False, indent=2))

        ttk.Label(dialog, text="Ключи столбцов: position, marking, name, quantity, mass_kg, total_mass, material, stock_length, bending\n"
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
