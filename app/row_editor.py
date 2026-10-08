"""
Редактор строк сводной ведомости.

Модальный диалог для просмотра и правки строк перед формированием
документов: изменение обозначения, наименования, количества, материала,
гибки и длины сортамента, добавление и удаление строк, поиск, фильтр
и сортировка по столбцам (сортировка меняет только вид, порядок строк
в документах остается прежним).
"""

import tkinter as tk
from tkinter import ttk, messagebox


class RowEditorDialog(tk.Toplevel):
    """Модальный редактор строк. Правки вносятся в переданный список in place."""

    COLUMNS = ("check", "num", "marking", "name", "quantity", "material", "bending", "stock_length")
    HEADERS = ("✓", "№", "Обозначение", "Наименование", "Кол-во", "Материал", "Гибка", "Длина сортамента")
    WIDTHS = (30, 36, 130, 260, 60, 150, 50, 110)

    FILTERS = ("Все строки", "Только детали", "Только сборки", "Гнутые",
               "С длиной сортамента", "Без материала")

    CHECK_ON = "☑"
    CHECK_OFF = "☐"

    def __init__(self, master, all_data):
        super().__init__(master)
        self.title("Редактор строк перед экспортом")
        self.geometry("940x560")
        self.transient(master)
        self.grab_set()

        self.all_data = all_data
        self.applied = False
        self._sort_col = None
        self._sort_desc = False

        self._setup_ui()
        self._reload()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.bind("<Escape>", lambda event: self._on_cancel())
        self.bind("<Control-f>", lambda event: self._focus_search())
        self.bind("<Control-F>", lambda event: self._focus_search())

    def _setup_ui(self):
        ttk.Label(self, text="Проверьте строки ведомости. Отмечайте строки галочками в первом "
                             "столбце или выделяйте мышью (Ctrl/Shift, протягивание). Для "
                             "выделенных применяется материал, гибка и длина сортамента.",
                  font=("Arial", 9), foreground="gray").pack(anchor="w", padx=15, pady=(10, 5))

        search_row = ttk.Frame(self)
        search_row.pack(fill=tk.X, padx=15, pady=(0, 2))
        ttk.Label(search_row, text="Поиск:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_row, textvariable=self.search_var, width=32)
        self.search_entry.pack(side=tk.LEFT, padx=(5, 4))
        ttk.Button(search_row, text="✕", width=3,
                   command=lambda: self.search_var.set("")).pack(side=tk.LEFT)
        ttk.Label(search_row, text="Показать:").pack(side=tk.LEFT, padx=(16, 5))
        self.filter_var = tk.StringVar(value=self.FILTERS[0])
        filter_box = ttk.Combobox(search_row, textvariable=self.filter_var, values=self.FILTERS,
                                  state="readonly", width=22)
        filter_box.pack(side=tk.LEFT)
        self.count_label = ttk.Label(search_row, text="", foreground="gray")
        self.count_label.pack(side=tk.RIGHT)
        self.search_var.trace_add("write", lambda *args: self._reload())
        filter_box.bind("<<ComboboxSelected>>", lambda event: self._reload())

        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        y_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        x_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        x_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(tree_frame, columns=self.COLUMNS, show="headings",
                                 selectmode="extended",
                                 yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        for col, header, width in zip(self.COLUMNS, self.HEADERS, self.WIDTHS):
            anchor = tk.W if col in ("marking", "name", "material") else tk.CENTER
            if col == "check":
                self.tree.heading(col, text=header)
            else:
                self.tree.heading(col, text=header, command=lambda c=col: self._sort_by(c))
            self.tree.column(col, width=width, anchor=anchor, stretch=(col in ("name", "material")))
        self.tree.pack(fill=tk.BOTH, expand=True)
        y_scroll.config(command=self.tree.yview)
        x_scroll.config(command=self.tree.xview)

        self.tree.tag_configure("assembly", background="#D6E4F0")
        self.tree.tag_configure("manual", background="#E2EFDA")

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Button-1>", self._on_click)

        btn_row = ttk.Frame(self)
        btn_row.pack(fill=tk.X, padx=15, pady=(5, 0))
        ttk.Button(btn_row, text="Отметить все", command=self._check_all).pack(side=tk.LEFT)
        ttk.Button(btn_row, text="Снять отметки", command=self._uncheck_all).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btn_row, text="Добавить строку", command=self._add_row).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Button(btn_row, text="Удалить выделенные", command=self._delete_rows).pack(side=tk.LEFT, padx=8)

        self.edit_frame = ttk.LabelFrame(self, text="Правка выделенной строки", padding=10)
        self.edit_frame.pack(fill=tk.X, padx=15, pady=(5, 0))

        self.marking_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.quantity_var = tk.StringVar()
        self.material_var = tk.StringVar()
        self.stock_length_var = tk.StringVar()
        self.bending_var = tk.BooleanVar(value=False)

        ttk.Label(self.edit_frame, text="Обозначение:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(self.edit_frame, textvariable=self.marking_var, width=28).grid(row=0, column=1, sticky=tk.EW, padx=(5, 15), pady=2)
        ttk.Label(self.edit_frame, text="Кол-во:").grid(row=0, column=2, sticky=tk.E, pady=2)
        ttk.Entry(self.edit_frame, textvariable=self.quantity_var, width=8).grid(row=0, column=3, sticky=tk.W, padx=(5, 0), pady=2)

        ttk.Label(self.edit_frame, text="Наименование:").grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Entry(self.edit_frame, textvariable=self.name_var, width=28).grid(row=1, column=1, sticky=tk.EW, padx=(5, 15), pady=2)
        ttk.Label(self.edit_frame, text="Гибка:").grid(row=1, column=2, sticky=tk.E, pady=2)
        ttk.Checkbutton(self.edit_frame, variable=self.bending_var).grid(row=1, column=3, sticky=tk.W, padx=(5, 0), pady=2)

        ttk.Label(self.edit_frame, text="Материал:").grid(row=2, column=0, sticky=tk.W, pady=2)
        ttk.Entry(self.edit_frame, textvariable=self.material_var, width=28).grid(row=2, column=1, sticky=tk.EW, padx=(5, 15), pady=2)
        ttk.Label(self.edit_frame, text="Длина сортамента, мм:").grid(row=2, column=2, sticky=tk.E, pady=2)
        ttk.Entry(self.edit_frame, textvariable=self.stock_length_var, width=8).grid(row=2, column=3, sticky=tk.W, padx=(5, 0), pady=2)

        self.edit_frame.columnconfigure(1, weight=1)

        ttk.Button(self.edit_frame, text="Применить", command=self._apply_form).grid(row=0, column=4, rowspan=3, padx=(15, 0))

        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill=tk.X, padx=15, pady=12)

        ttk.Button(bottom_frame, text="Продолжить экспорт", command=self._on_ok).pack(side=tk.RIGHT)
        ttk.Button(bottom_frame, text="Отмена", command=self._on_cancel).pack(side=tk.RIGHT, padx=10)

    def _reload(self, select=None):
        if select is None:
            selected = list(self.tree.selection())
        elif isinstance(select, (list, tuple)):
            selected = [str(s) for s in select]
        else:
            selected = [str(select)]

        self.tree.delete(*self.tree.get_children())
        selected = set(selected)
        indices = [i for i, item in enumerate(self.all_data) if self._matches(item)]
        if self._sort_col:
            indices.sort(key=lambda i: self._sort_key(i, self._sort_col), reverse=self._sort_desc)
        self.count_label.config(
            text=f"Показано {len(indices)} из {len(self.all_data)}"
            + (" · сортировка меняет только вид" if self._sort_col else ""))
        for idx in indices:
            item = self.all_data[idx]
            tags = []
            if item.get("manual"):
                tags.append("manual")
            if item.get("is_assembly"):
                tags.append("assembly")
            values = (
                self.CHECK_ON if str(idx) in selected else self.CHECK_OFF,
                idx + 1,
                item.get("marking", ""),
                item.get("name", ""),
                item.get("quantity", 1),
                item.get("material", ""),
                "X" if item.get("is_bending") else "",
                item.get("stock_length", ""),
            )
            self.tree.insert("", tk.END, iid=str(idx), values=values, tags=tags)

        valid = [s for s in selected if self.tree.exists(s)]
        if valid:
            self.tree.selection_set(valid)
            self.tree.see(valid[0])

    def _focus_search(self):
        self.search_entry.focus_set()
        self.search_entry.select_range(0, tk.END)

    def _matches(self, item):
        """Строка проходит поиск и фильтр."""
        text = self.search_var.get().strip().lower()
        if text:
            hay = " ".join(str(item.get(k, "") or "") for k in ("marking", "name", "material")).lower()
            if text not in hay:
                return False
        mode = self.filter_var.get()
        if mode == "Только детали":
            return not item.get("is_assembly")
        if mode == "Только сборки":
            return bool(item.get("is_assembly"))
        if mode == "Гнутые":
            return bool(item.get("is_bending"))
        if mode == "С длиной сортамента":
            return bool(str(item.get("stock_length", "") or "").strip())
        if mode == "Без материала":
            return not item.get("is_assembly") and not str(item.get("material", "") or "").strip()
        return True

    @staticmethod
    def _number(value):
        try:
            return float(str(value).replace(",", ".").strip())
        except (TypeError, ValueError):
            return None

    def _sort_key(self, idx, col):
        """Ключ сортировки; пустые значения в конце списка при любом направлении."""
        item = self.all_data[idx]
        empty = -1 if self._sort_desc else 1
        if col == "num":
            return (0, idx)
        if col == "bending":
            return (0, 1 if item.get("is_bending") else 0)
        if col in ("quantity", "stock_length"):
            n = self._number(item.get(col, ""))
            return (empty, 0) if n is None else (0, n)
        text = str(item.get(col, "") or "").strip().lower()
        return (0, text) if text else (empty, "")

    def _sort_by(self, col):
        """Сортировка по столбцу: повторный щелчок меняет направление, третий снимает."""
        if self._sort_col == col and not self._sort_desc:
            self._sort_desc = True
        elif self._sort_col == col:
            self._sort_col, self._sort_desc = None, False
        else:
            self._sort_col, self._sort_desc = col, False
        for c, header in zip(self.COLUMNS, self.HEADERS):
            arrow = ""
            if c == self._sort_col:
                arrow = " ▼" if self._sort_desc else " ▲"
            self.tree.heading(c, text=header + arrow)
        self._reload()

    def _render_checks(self):
        current = set(self.tree.selection())
        for iid in self.tree.get_children():
            mark = self.CHECK_ON if iid in current else self.CHECK_OFF
            if self.tree.set(iid, "check") != mark:
                self.tree.set(iid, "check", mark)

    def _on_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return None
        if self.tree.identify_column(event.x) != "#1":
            return None
        row = self.tree.identify_row(event.y)
        if not row:
            return None
        if row in self.tree.selection():
            self.tree.selection_remove(row)
        else:
            self.tree.selection_add(row)
        self._render_checks()
        self._on_select()
        return "break"

    def _check_all(self):
        self.tree.selection_set(self.tree.get_children())
        self._render_checks()

    def _uncheck_all(self):
        self.tree.selection_remove(*self.tree.selection())
        self._render_checks()

    def _selected_item(self):
        current = self.tree.selection()
        if not current:
            return None, None
        idx = int(current[0])
        return idx, self.all_data[idx]

    def _clear_form(self):
        self.marking_var.set("")
        self.name_var.set("")
        self.quantity_var.set("")
        self.material_var.set("")
        self.stock_length_var.set("")
        self.bending_var.set(False)

    def _update_edit_label(self, count):
        if count > 1:
            self.edit_frame.config(
                text=f"Правка выделенных строк ({count}): материал, гибка, длина сортамента")
        else:
            self.edit_frame.config(text="Правка выделенной строки")

    def _on_select(self, event=None):
        self._render_checks()
        current = self.tree.selection()
        if not current:
            self._clear_form()
            self._update_edit_label(0)
            return
        item = self.all_data[int(current[0])]
        self.marking_var.set(item.get("marking", ""))
        self.name_var.set(item.get("name", ""))
        self.quantity_var.set(str(item.get("quantity", 1)))
        self.material_var.set(item.get("material", ""))
        self.stock_length_var.set(item.get("stock_length", "") or "")
        self.bending_var.set(bool(item.get("is_bending")))
        self._update_edit_label(len(current))

    @staticmethod
    def _parse_quantity(raw, fallback):
        text = str(raw).strip().replace(",", ".")
        if not text:
            return fallback
        try:
            value = float(text)
        except ValueError:
            return fallback
        if value <= 0:
            return fallback
        return int(value) if value == int(value) else value

    def _apply_form(self):
        current = self.tree.selection()
        if not current:
            messagebox.showwarning("Внимание", "Выделите строку для правки.", parent=self)
            return

        if len(current) == 1:
            idx = int(current[0])
            item = self.all_data[idx]
            item["marking"] = self.marking_var.get().strip()
            item["name"] = self.name_var.get().strip() or item.get("name", "")
            item["quantity"] = self._parse_quantity(self.quantity_var.get(), item.get("quantity", 1))
            item["material"] = self.material_var.get().strip()
            item["stock_length"] = self.stock_length_var.get().strip()
            item["is_bending"] = bool(self.bending_var.get())
            self._reload(select=str(idx))
            return

        material = self.material_var.get().strip()
        stock_length = self.stock_length_var.get().strip()
        bending = bool(self.bending_var.get())
        for iid in current:
            item = self.all_data[int(iid)]
            item["material"] = material
            item["stock_length"] = stock_length
            if not item.get("is_assembly"):
                item["is_bending"] = bending
        self._reload(select=list(current))

    def _add_row(self):
        item = {
            "level": 1,
            "parent": "",
            "name": "Новая позиция",
            "marking": "",
            "quantity": 1,
            "is_assembly": False,
            "mass": None,
            "material": "",
            "is_bending": False,
            "stock_length": "",
            "manual": True,
        }
        self.all_data.append(item)
        self._reload(select=str(len(self.all_data) - 1))

    def _delete_rows(self):
        current = self.tree.selection()
        if not current:
            messagebox.showwarning("Внимание", "Выделите строки для удаления.", parent=self)
            return
        removable = []
        locked = 0
        for iid in current:
            item = self.all_data[int(iid)]
            if item.get("level") == 0:
                locked += 1
            else:
                removable.append(int(iid))
        if not removable:
            messagebox.showwarning("Внимание", "Главную сборку удалить нельзя.", parent=self)
            return
        message = (f"Удалить строку «{self.all_data[removable[0]].get('name', '')}»?"
                   if len(removable) == 1
                   else f"Удалить выделенные строки ({len(removable)})?")
        if locked:
            message += "\nГлавная сборка останется без изменений."
        if not messagebox.askyesno("Удаление", message, parent=self):
            return
        for idx in sorted(removable, reverse=True):
            del self.all_data[idx]
        self._reload(select=[])
        self._clear_form()
        self._update_edit_label(0)

    def _on_ok(self):
        self.applied = True
        self.destroy()

    def _on_cancel(self):
        if messagebox.askyesno("Отмена", "Прервать экспорт без сохранения изменений?", parent=self):
            self.applied = False
            self.destroy()
