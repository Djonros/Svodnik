"""
Финальная версия экспорта сборки из КОМПАС-3D v24.

Структура соответствует эталонному файлу:
- Главная сборка в начале
- Позиции из спецификаций
- Столбец "Формат" (A1, A2...)
- Столбец "Гибка"
- Группировка сборок

Использование:
    python kompas_export_final.py "C:\\path\\to\\file.a3d"
"""

import sys
import os
import time
import glob
from datetime import datetime

import win32com.client
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT


import re


DEFAULT_MATERIALS = [
    "сталь 10", "сталь 20", "сталь 45", "сталь 09г2с",
    "сталь 09г2", "сталь 20х", "сталь 40х", "сталь 12х13",
    "сталь 14х17н2", "сталь 20х13",
]


def is_default_material(material: str) -> bool:
    if not material:
        return False
    mat_lower = material.lower().strip()
    for default in DEFAULT_MATERIALS:
        if mat_lower.startswith(default):
            return True
    return False


class KompasExportFinal:
    """Финальный класс для экспорта сборки."""

    def __init__(self):
        self.app = None
        self.doc = None
        self.doc3d = None
        self.top_part = None
        self.assembly_name = ""
        self.assembly_dir = ""
        self.specs_dir = ""
        self.all_data = []
        self.all_positions = {}
        self.opened_docs = []

    def connect(self):
        try:
            self.app = win32com.client.Dispatch("Kompas.Application.7")
            self.app.Visible = True
            print("[OK] Подключено к КОМПАС-3D v24")
            return True
        except Exception as e:
            print(f"[!!] Ошибка: {e}")
            return False

    def open_assembly(self, filepath=None):
        if filepath:
            self.doc = self.app.Documents.Open(filepath)
        else:
            self.doc = self.app.ActiveDocument

        if not self.doc:
            print("[!!] Нет открытого документа")
            return False

        self.assembly_name = self.doc.Name
        self.assembly_dir = os.path.dirname(self.doc.PathName)
        self.specs_dir = os.path.join(self.assembly_dir, "Генерированные спецификации")
        os.makedirs(self.specs_dir, exist_ok=True)

        print(f"[OK] Документ: {self.assembly_name}")
        return True

    def extract_tree(self):
        try:
            self.doc3d = win32com.client.CastTo(self.doc, "IKompasDocument3D")
        except Exception as e:
            print(f"[!!] CastTo IKompasDocument3D не удался: {e}")
            print("[--] Используем исходный COM-объект")
            self.doc3d = self.doc

        self.top_part = self.doc3d.TopPart
        if not self.top_part:
            print("[!!] TopPart не найден")
            return False

        print(f"[OK] Корень: {self.top_part.Name} ({self.top_part.Marking})")

        self.all_data = []
        seen = set()

        # Добавляем корень (главную сборку)
        root_info = {
            "level": 0,
            "parent": "",
            "name": self.top_part.Name or "Без имени",
            "marking": self.top_part.Marking or "",
            "quantity": 1,
            "is_assembly": True,
            "mass": None,
            "material": "",
            "is_bending": False,
        }
        try:
            root_info["mass"] = self.top_part.Mass
        except:
            pass
        try:
            root_info["material"] = self.top_part.Material or ""
        except:
            pass

        self.all_data.append(root_info)
        seen.add(f"{root_info['marking']}|{root_info['name']}")

        # Обходим дочерние элементы
        self._walk_parts(self.top_part, level=1, parent_name=self.top_part.Name, seen=seen)

        print(f"[OK] Найдено элементов: {len(self.all_data)}")
        return True

    def _walk_parts(self, part, level, parent_name, seen):
        try:
            parts = part.Parts
            if not parts:
                return

            for i in range(parts.Count):
                model_obj = parts.Item(i)
                if not model_obj:
                    continue
                try:
                    p7 = win32com.client.CastTo(model_obj, "IPart7")
                except Exception:
                    p7 = model_obj

                name = p7.Name or "Без имени"
                marking = p7.Marking or ""

                # Используем уникальную ссылку для предотвращения зацикливания
                try:
                    p7_ref = p7.Reference
                except:
                    p7_ref = f"{marking}|{name}|{level}"

                if p7_ref in seen:
                    continue
                seen.add(p7_ref)

                info = {
                    "level": level,
                    "parent": parent_name,
                    "name": name,
                    "marking": marking,
                    "quantity": 1,
                    "is_assembly": False,
                    "mass": None,
                    "material": "",
                    "is_bending": False,
                }

                try:
                    info["mass"] = p7.Mass
                except:
                    pass
                try:
                    info["material"] = p7.Material or ""
                except:
                    pass

                try:
                    child_parts = p7.Parts
                    if child_parts and child_parts.Count > 0:
                        info["is_assembly"] = True
                except:
                    pass

                info["is_bending"] = self._check_bending(p7, info["name"])

                self.all_data.append(info)

                if info["is_assembly"]:
                    self._walk_parts(p7, level + 1, name, seen)
        except Exception as e:
            pass

    def _check_bending(self, p7, name):
        """Определение гнущихся деталей."""
        name_lower = name.lower()
        bending_keywords = [
            "накладка", "пластина", "площадка", "стенка", "днище",
            "крышка", "фланец", "обшивка", "настил", "бортик"
        ]
        for kw in bending_keywords:
            if kw in name_lower:
                return True

        # Проверка по толщине
        try:
            material = p7.Material or ""
            import re
            match = re.search(r'(\d+[\.,]?\d*)\s*(?:мм|$)', material)
            if match:
                thickness = float(match.group(1).replace(',', '.'))
                if thickness <= 10:
                    return True
        except:
            pass

        return False

    def read_positions(self):
        """Чтение позиций из всех .spw файлов."""
        print("\nЧтение позиций из спецификаций...")

        self.all_positions = {}

        # Ищем .spw файлы (все спецификации в папке Рабочка)
        parent_dir = os.path.dirname(self.assembly_dir)
        all_spw = []
        for search_dir in [parent_dir, self.assembly_dir]:
            if not os.path.exists(search_dir):
                continue
            for f in os.listdir(search_dir):
                if f.endswith('.spw') and 'СП' in f:
                    all_spw.append(os.path.join(search_dir, f))

        print(f"  Найдено .spw файлов: {len(all_spw)}")
        for spw in all_spw:
            try:
                positions = self._read_spw(spw)
                self.all_positions.update(positions)
            except:
                pass

        print(f"[OK] Найдено позиций: {len(self.all_positions)}")
        return True

    def _read_spw(self, spw_path):
        """Чтение позиций из .spw файла."""
        positions = {}
        try:
            doc = self.app.Documents.Open(spw_path)
            self.opened_docs.append(doc)
            time.sleep(0.5)

            sd = doc.SpecificationDescriptions
            if not sd or sd.Count == 0:
                return positions

            spec_desc = sd.Item(0)

            # Проверяем нужно ли перестроить
            if spec_desc.NeedRebuild:
                print(f"  Перестроение спецификации: {os.path.basename(spw_path)}...")
                spec_desc.Update()
                time.sleep(1)

            objects = spec_desc.Objects

            for obj in objects:
                try:
                    cols = obj.Columns
                    if not cols or cols.Count < 5:
                        continue

                    pos_col = cols.Item(2)
                    pos_text = pos_col.Text.Str if pos_col.Text else ""

                    desig_col = cols.Item(3)
                    desig_text = desig_col.Text.Str if desig_col.Text else ""

                    name_col = cols.Item(4)
                    name_text = name_col.Text.Str if name_col.Text else ""

                    qty_col = cols.Item(5)
                    qty_text = qty_col.Text.Str if qty_col.Text else "1"

                    if desig_text and pos_text:
                        positions[desig_text] = {
                            "position": pos_text,
                            "name": name_text,
                            "quantity": qty_text,
                        }
                except:
                    continue
        except:
            pass

        return positions

    def close_all_docs(self):
        """Закрытие всех открытых документов."""
        for doc in self.opened_docs:
            try:
                if doc and not doc.Closed:
                    doc.Close()
            except:
                pass
        self.opened_docs.clear()

    def generate_excel(self):
        """Генерация Excel по формату эталона."""
        print("\nСоздание Excel файла...")

        wb = Workbook()
        ws = wb.active
        ws.title = os.path.splitext(self.assembly_name)[0][:31]

        header_font = Font(bold=True, size=10, color="FFFFFF")
        header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
        section_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        assembly_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
        border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin"),
        )
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left = Alignment(horizontal="left", vertical="center", wrap_text=True)

        headers = [
            ("A", "Уровень", 8), ("B", "", 3), ("C", "", 3), ("D", "", 3), ("E", "", 3),
            ("F", "Формат", 8), ("G", "Позиция", 8),
            ("H", "№ деталей;\n№ сб., № п/сб.", 25), ("I", "Наименование", 40),
            ("J", "Кол-во всего", 12), ("K", "Масса", 12), ("L", "Масса общая", 12),
            ("M", "Материал", 25), ("N", "Гибка", 8),
            ("O", "Материал\nзаменитель", 15), ("P", "Примечания", 15),
            ("Q", "Старое\nобозначение", 15), ("R", "", 5),
        ]
        for col, title, width in headers:
            cell = ws[f"{col}1"]
            cell.value = title
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border
            ws.column_dimensions[col].width = width
        ws.row_dimensions[1].height = 35

        row = 2
        for item in self.all_data:
            level = item.get("level", 0)
            marking = item.get("marking", "")
            pos_info = self.all_positions.get(marking, {})
            pos = pos_info.get("position", "")
            mass_kg = (item.get("mass", 0) or 0) / 1000
            total_mass = mass_kg * item.get("quantity", 1)
            is_asm = item.get("is_assembly", False)
            bending = "X" if item.get("is_bending") else ""
            suffix = "  СБ" if is_asm else ""

            fill = assembly_fill if is_asm else None

            # Для сборок убираем материал по умолчанию
            material = item.get("material", "")
            if is_asm and is_default_material(material):
                material = ""

            values = [
                level + 1 if level > 0 else "",  # A
                "",  # B
                "",  # C
                "",  # D
                "",  # E
                "",  # F
                pos,  # G
                f"{marking}{suffix}",  # H
                item.get("name", ""),  # I
                item.get("quantity", 1),  # J
                round(mass_kg, 6) if mass_kg > 0 else "",  # K
                round(total_mass, 6) if total_mass > 0 else "",  # L
                material,  # M
                bending,  # N - Гибка
                "",  # O
                "",  # P
                "",  # Q
                "",  # R
            ]

            for col_idx, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=col_idx)
                cell.value = value
                cell.border = border
                cell.alignment = left if col_idx in [4, 8, 9, 13, 17] else center
                if fill:
                    cell.fill = fill

            # Группировка для сборок (outline level)
            if is_asm and level > 0:
                ws.row_dimensions[row].outline_level = min(level, 8)

            row += 1

        row += 1
        total_mass = sum((item.get("mass", 0) or 0) / 1000 for item in self.all_data)
        ws.cell(row=row, column=1, value="ИТОГО:").font = Font(bold=True)
        ws.cell(row=row, column=5, value=f"{len(self.all_data)} элементов").font = Font(bold=True)
        ws.cell(row=row, column=7, value=round(total_mass, 3)).font = Font(bold=True)

        # Включаем группировку через row_dimensions
        for r in range(2, row):
            cell_a = ws.cell(row=r, column=1)
            if cell_a.value and int(cell_a.value) > 1:
                ws.row_dimensions[r].outline_level = min(int(cell_a.value) - 1, 8)
                ws.row_dimensions[r].hidden = False

        base_name = os.path.splitext(self.assembly_name)[0]
        output_path = os.path.join(self.assembly_dir, f"{base_name}_сводная_ведомость.xlsx")
        wb.save(output_path)
        print(f"[OK] Excel: {os.path.abspath(output_path)}")
        return output_path

    def _merge_items(self, items):
        """Объединение одинаковых элементов."""
        merged = {}
        for item in items:
            key = f"{item.get('parent', '')}|{item.get('marking', '')}|{item.get('name', '')}"
            if key not in merged:
                merged[key] = item.copy()
            else:
                merged[key]["quantity"] += item.get("quantity", 1)
        return list(merged.values())

    def generate_word(self):
        """Генерация Word файла."""
        print("\nСоздание Word файла...")

        doc = Document()
        base_name = os.path.splitext(self.assembly_name)[0]
        base_designation = base_name.split('.')[0]

        doc.add_heading('СВОДНАЯ ВЕДОМОСТЬ', 0)
        doc.add_paragraph(f'Обозначение: {base_designation}')
        doc.add_paragraph(f'Наименование: {base_name}')
        doc.add_paragraph('')

        # Документация
        doc.add_heading('Документация', 1)
        table = doc.add_table(rows=5, cols=2)
        table.style = 'Table Grid'
        docs = [
            ('Обозначение', 'Наименование'),
            (f'{base_designation}.0000010', 'Сборочный чертеж'),
            (f'{base_designation}.0000010 КО', 'Карта окраски'),
            (f'{base_designation}.0000010 ПС', 'Паспорт'),
            (f'{base_designation}.0000010 РЭ', 'Руководство по эксплуатации'),
        ]
        for i, (col1, col2) in enumerate(docs):
            table.cell(i, 0).text = col1
            table.cell(i, 1).text = col2
            if i == 0:
                for cell in table.rows[i].cells:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            r.font.bold = True
        doc.add_paragraph('')

        # 0. Комплект
        doc.add_heading('0. Комплект', 1)
        if self.all_data:
            self._add_table(doc, [self.all_data[0]])
        doc.add_paragraph('')

        # Рекурсивно добавляем элементы
        def add_items_word(parent_name, level):
            children = [item for item in self.all_data
                       if item.get("parent") == parent_name
                       and item.get("name") != parent_name]
            if not children:
                return

            assemblies = [c for c in children if c.get("is_assembly")]
            details = [c for c in children if not c.get("is_assembly")]
            assemblies = self._merge_items(assemblies)
            details = self._merge_items(details)

            for asm in assemblies:
                doc.add_heading(asm.get("name", ""), level)
                self._add_table(doc, [asm])
                add_items_word(asm.get("name", ""), level + 1)

            if details:
                doc.add_heading('Детали', level)
                self._add_table(doc, details)

        if self.all_data:
            add_items_word(self.all_data[0].get("name", ""), 2)

        doc.add_heading('ИТОГО', 1)
        total_mass = sum((item.get("mass", 0) or 0) / 1000 for item in self.all_data)
        doc.add_paragraph(f'Всего элементов: {len(self.all_data)}')
        doc.add_paragraph(f'Общая масса: {total_mass:.3f} кг')

        output_path = os.path.join(self.assembly_dir, f"{base_name}_сводная_ведомость.docx")
        doc.save(output_path)
        print(f"[OK] Word: {os.path.abspath(output_path)}")
        return output_path

    def _add_table(self, doc, items):
        if not items:
            return
        table = doc.add_table(rows=len(items) + 1, cols=7)
        table.style = 'Table Grid'
        headers = ['Поз.', 'Обозначение', 'Наименование', 'Кол-во', 'Масса, кг', 'Материал', 'Гибка']
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = h
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.bold = True
                    r.font.size = Pt(9)
        for idx, item in enumerate(items, 1):
            marking = item.get("marking", "")
            spec_pos = self.all_positions.get(marking, {}).get("position", str(idx))
            mass_kg = (item.get("mass", 0) or 0) / 1000
            bending = "X" if item.get("is_bending") else ""
            is_asm = item.get("is_assembly", False)
            material = item.get("material", "")
            if is_asm and is_default_material(material):
                material = ""
            row_data = [spec_pos, marking, item.get("name", ""), str(item.get("quantity", 1)),
                        f"{mass_kg:.3f}" if mass_kg > 0 else "", material, bending]
            for col_idx, value in enumerate(row_data):
                cell = table.cell(idx, col_idx)
                cell.text = value
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(9)

    def run(self, filepath=None):
        try:
            if not self.connect():
                return
            if not self.open_assembly(filepath):
                return
            if not self.extract_tree():
                return
            self.read_positions()
            excel_path = self.generate_excel()
            word_path = self.generate_word()
            print("\n" + "=" * 50)
            print("Готово!")
            print(f"Excel: {excel_path}")
            print(f"Word: {word_path}")
            print("=" * 50)
        finally:
            self.close_all_docs()


def main():
    filepath = sys.argv[1] if len(sys.argv) > 1 else None
    exporter = KompasExportFinal()
    exporter.run(filepath)


if __name__ == "__main__":
    main()
