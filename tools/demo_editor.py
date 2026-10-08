"""Показ редактора строк на тестовых данных с поиском и сортировкой (проверка интерфейса)."""
import os, sys, tkinter as tk
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "app"))
from row_editor import RowEditorDialog

data = [{"level": 0, "name": "Подкатная тележка", "marking": "STI 700005", "quantity": 1, "is_assembly": True}]
for i, (m, n, mat, b, L) in enumerate([
        ("700005.2801203", "Швеллер верхний", "Лист 8", True, ""),
        ("700005.2801202", "Профиль 40х40х4 L = 651 мм", "Профиль 40х40х4", False, "651"),
        ("700005.2801204", "Профиль 40х40х4 L = 391 мм", "Профиль 40х40х4", False, "391"),
        ("700005.2801205", "Планка прижимная", "Лист 10", True, ""),
        ("", "Профиль 70x70x5 L = 1330 мм", "Профиль 70x70x5", False, "1330"),
        ("", "Болт М22", "", False, "")]):
    data.append({"level": 1, "name": n, "marking": m, "quantity": i + 1, "is_assembly": False,
                 "material": mat, "is_bending": b, "stock_length": L})
root = tk.Tk()
root.geometry("200x100+0+0")
dlg = RowEditorDialog(root, data)
dlg.title("Сводник редактор демо")
dlg.search_var.set("проф")
dlg._sort_by("stock_length")
dlg._sort_by("stock_length")
dlg.update()
def show(title):
    print(title, "|", dlg.count_label.cget("text"))
    for iid in dlg.tree.get_children():
        v = dlg.tree.item(iid, "values")
        print("   ", v[1], v[2], v[3], v[4], v[6], v[7])
show("поиск 'проф', длина по убыванию")
dlg._sort_by("stock_length"); show("сортировка снята")
dlg.search_var.set(""); dlg.filter_var.set("Гнутые"); dlg._reload(); show("фильтр Гнутые")
dlg.filter_var.set("Все строки"); dlg._sort_by("name"); show("по наименованию")
dlg.tree.selection_set(dlg.tree.get_children()[0]); dlg._on_select()
dlg.quantity_var.set("7"); dlg._apply_form(); show("после правки первой строки")
