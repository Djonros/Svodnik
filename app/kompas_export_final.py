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
from fpdf import FPDF


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

    def __init__(self, template=None):
        self.app = None
        self.doc = None
        self.doc3d = None
        self.top_part = None
        self.assembly_name = ""
        self.assembly_dir = ""
        self.specs_dir = ""
        self.all_data = []
        self.all_positions = {}
        self.spec_quantities = {}
        self.model_counts = {}
        self.quantity_warnings = []
        self.problems = []
        self.assembly_files = {}
        self.generated_specs = []
        self.bend_cache = {}
        self.length_cache = {}
        self.opened_docs = []
        self.template = template or self._default_template()
        if template:
            self._ensure_stock_length(self.template.get("excel", {}).get("columns"))
            self._ensure_stock_length(self.template.get("word", {}).get("columns"))

    def _default_template(self):
        default_path = os.path.join(os.path.dirname(__file__), "template_default.json")
        if os.path.exists(default_path):
            return self.load_template(default_path)
        return {"excel": {"columns": []}, "word": {"columns": []}}

    def load_template(self, path):
        import json
        with open(path, "r", encoding="utf-8") as f:
            template = json.load(f)
        self._ensure_stock_length(template.get("excel", {}).get("columns"))
        self._ensure_stock_length(template.get("word", {}).get("columns"))
        return template

    @staticmethod
    def _ensure_stock_length(columns):
        """Добавление столбца "Длина сортамента" в шаблоны старых версий."""
        if not columns:
            return
        if any(c.get("key") == "stock_length" for c in columns):
            return
        default = {"key": "stock_length", "header": "Длина сортамента", "width": 12}
        for i, col in enumerate(columns):
            if col.get("key") == "material":
                columns.insert(i + 1, default)
                return
        columns.append(default)

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
        self.model_counts = {}
        self.problems = []
        self.assembly_files = {}
        if self.top_part.Marking:
            self.assembly_files[self.top_part.Marking] = self.doc.PathName

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

        # Обходим дочерние элементы
        self._walk_parts(self.top_part, level=1, parent_name=self.top_part.Name,
                         parent_marking=root_info["marking"], ancestors=set())

        print(f"[OK] Найдено элементов: {len(self.all_data)}")
        return True

    def _problem(self, message):
        """Проблема, из-за которой ведомость может быть неполной или неточной:
        пишется в лог и показывается пользователю в конце экспорта."""
        self.problems.append(message)
        print(f"[!!] {message}")

    MAX_DEPTH = 30

    def _walk_parts(self, part, level, parent_name, parent_marking="", ancestors=None):
        """Обход вставок сборки. Каждая вставка компонента дает отдельную
        запись с количеством 1; одинаковые позиции суммируются при группировке.
        Повторы НЕ отсекаются: две одинаковые детали в сборке = количество 2."""
        if ancestors is None:
            ancestors = set()
        if level > self.MAX_DEPTH:
            self._problem(f"Слишком глубокая вложенность в '{parent_name}', "
                          f"дальше состав не читался")
            return

        try:
            parts = part.Parts
            count = parts.Count if parts else 0
        except Exception as e:
            self._problem(f"Не прочитан состав '{parent_name}', его детали не попали "
                          f"в ведомость: {e}")
            return

        # Количество вставок по обозначению на один экземпляр родителя (для сверки со .spw)
        counts = {}

        for i in range(count):
            try:
                model_obj = parts.Item(i)
                if not model_obj:
                    continue
                try:
                    p7 = win32com.client.CastTo(model_obj, "IPart7")
                except Exception:
                    p7 = model_obj

                # Компоновочная геометрия в состав изделия не входит
                try:
                    if p7.IsLayoutGeometry:
                        continue
                except Exception:
                    pass

                name = p7.Name or "Без имени"
                marking = p7.Marking or ""
            except Exception as e:
                self._problem(f"Пропущен компонент №{i + 1} в '{parent_name}': {e}")
                continue

            # Защита от зацикливания: компонент не может входить сам в себя
            try:
                file_key = (p7.FileName or "").lower()
            except Exception:
                file_key = ""
            if file_key and file_key in ancestors:
                continue

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

            label = f"'{name}' ({marking})" if marking else f"'{name}'"
            try:
                info["mass"] = p7.Mass
            except Exception as e:
                self._problem(f"Не прочитана масса {label} в '{parent_name}': {e}")
            try:
                info["material"] = p7.Material or ""
            except Exception as e:
                self._problem(f"Не прочитан материал {label} в '{parent_name}': {e}")

            try:
                child_parts = p7.Parts
                if child_parts and child_parts.Count > 0:
                    info["is_assembly"] = True
                    if marking and file_key:
                        self.assembly_files.setdefault(marking, p7.FileName)
            except Exception as e:
                self._problem(f"Не прочитан состав {label} в '{parent_name}', "
                              f"компонент учтен как деталь: {e}")

            info["is_bending"] = (
                self._is_bent(p7, info["material"]) if not info["is_assembly"] else False
            )
            if not info["is_assembly"]:
                info["stock_length"] = self._stock_length(p7, name, info["material"])

            self.all_data.append(info)
            if marking:
                counts[marking] = counts.get(marking, 0) + 1

            if info["is_assembly"]:
                child_ancestors = ancestors | {file_key} if file_key else ancestors
                self._walk_parts(p7, level + 1, name, marking, child_ancestors)

        if parent_marking:
            self.model_counts[parent_marking] = counts

    # Операции листового тела, которые дают гиб
    SHEET_BEND_COLLECTIONS = (
        "SheetMetalBends", "SheetMetalLineBends", "SheetMetalSketchBends",
        "SheetMetalFlangings", "SheetMetalShoulders", "SheetMetalRuledShells",
        "SheetMetalLinearRuledShells", "SheetMetalBendedStraightens",
    )
    # Материалы из листа и полосы: для них ищем гибы по геометрии
    SHEET_MATERIAL_RE = re.compile(r"лист|полос", re.IGNORECASE)
    TUBE_MATERIAL_RE = re.compile(r"труб|круг|пруток", re.IGNORECASE)
    # Толщина из обозначения материала: "Лист$d5 ...", "Лист$dБ-ПН-О-12 ...", "Полоса 4х40"
    THICKNESS_RE = re.compile(r"(?:лист|полос\w*)\D{0,15}?(\d+(?:[.,]\d+)?)", re.IGNORECASE)

    def _is_bent(self, p7, material=""):
        """Гнутая деталь или нет. Результат кэшируется по файлу детали.
        Признаки (любой):
        1. В детали есть листовые операции сгиба (сгиб, сгиб по линии, буртик,
           подсечка, обечайка и т.п.) или листовое тело разогнуто.
        2. Деталь из листа/полосы, и в ней есть пара соосных цилиндрических
           граней с одинаковым углом, радиусы которых отличаются на толщину
           листа (внутренний и наружный радиус гиба) - так находятся гибы
           у деталей, построенных обычным выдавливанием.
        3. Деталь из трубы или круга, и в ней есть тороидальные грани (гнутая труба)."""
        try:
            key = (p7.FileName or "").lower()
        except Exception:
            key = ""
        if key and key in self.bend_cache:
            return self.bend_cache[key]

        bent = self._sheet_metal_bent(p7)
        if not bent:
            bent = self._geometry_bent(p7, material)
        if key:
            self.bend_cache[key] = bent
        return bent

    # Сортовой прокат, для которого заполняется "Длина сортамента"
    STOCK_RE = re.compile(r"профил|труб|уголок|уголк|швеллер|двутавр|балк|круг|квадрат|"
                          r"шестигран|арматур|пруток|полос", re.IGNORECASE)
    LENGTH_IN_NAME_RE = re.compile(r"\bL\s*=\s*(\d+(?:[.,]\d+)?)", re.IGNORECASE)

    def _stock_length(self, p7, name, material):
        """Длина детали из сортового проката, мм (строкой), или "" для остальных.
        Прокат определяется по материалу, а если материал не задан
        (или задан по умолчанию) - по наименованию; детали из листа не считаются.
        Длина берется из "L = ... мм" в наименовании/файле (так ее пишет библиотека
        металлоконструкций), иначе - наибольший габарит детали в ее собственной
        системе координат."""
        material = material or ""
        if self.SHEET_MATERIAL_RE.search(material) and not re.search(r"полос", material, re.IGNORECASE):
            return ""
        if not (self.STOCK_RE.search(material) or
                ((not material or is_default_material(material)) and self.STOCK_RE.search(name or ""))):
            return ""
        try:
            path = p7.FileName or ""
        except Exception:
            path = ""
        key = path.lower()
        if key and key in self.length_cache:
            return self.length_cache[key]

        length = ""
        for text in (name, os.path.basename(path)):
            m = self.LENGTH_IN_NAME_RE.search(text or "")
            if m:
                length = m.group(1).replace(",", ".")
                break
        if not length and path:
            dims = self._part_own_size(path)
            if dims:
                size = max(dims)
                # Деталь может лежать в своей системе координат наискосок: тогда
                # ее длину дает самое длинное прямое ребро (если оно помещается в габарит)
                diagonal = sum(d * d for d in dims) ** 0.5
                edge = self._longest_straight_edge(p7)
                if size < edge <= diagonal + 0.5:
                    size = edge
                length = str(int(round(size)))
        if not length:
            self._problem(f"Не определена длина сортамента у '{name}'")
        if key:
            self.length_cache[key] = length
        return length

    def _part_own_size(self, path):
        """Габариты детали (x, y, z) в ее собственной системе координат
        (в сборке деталь может быть повернута, и габарит в сборке неверен)."""
        doc = None
        already_open = False
        try:
            for i in range(self.app.Documents.Count):
                d = self.app.Documents.Item(i)
                if (d.PathName or "").lower() == path.lower():
                    doc, already_open = d, True
                    break
            if doc is None:
                doc = self.app.Documents.Open(path, False, True)
            top = win32com.client.CastTo(
                win32com.client.CastTo(doc, "IKompasDocument3D").TopPart, "IPart7")
            g = top.GetGabarit(True, True)
            return [abs(g[i + 4] - g[i + 1]) for i in range(3)]
        except Exception as e:
            print(f"[--] Габарит {os.path.basename(path)} не прочитан: {e}")
            return None
        finally:
            if doc is not None and not already_open:
                try:
                    doc.Close(0)
                except Exception:
                    pass

    @staticmethod
    def _longest_straight_edge(p7):
        """Длина самого длинного прямого ребра детали, мм."""
        best = 0.0
        try:
            for e in (win32com.client.CastTo(p7, "IFeature7").ModelObjects(7) or ()):  # o3d_edge
                e = win32com.client.CastTo(e, "IEdge")
                if e.IsStraight:
                    best = max(best, e.GetLength(0) * 10.0)  # GetLength в см
        except Exception:
            pass
        return best

    def _sheet_metal_bent(self, p7):
        try:
            container = win32com.client.CastTo(p7, "ISheetMetalContainer")
        except Exception:
            return False
        for name in self.SHEET_BEND_COLLECTIONS:
            try:
                if getattr(container, name).Count > 0:
                    return True
            except Exception:
                pass
        try:
            bodies = container.SheetMetalBodies
            for i in range(bodies.Count):
                if bool(bodies.SheetMetalBody(i).Straighten):
                    return True
        except Exception:
            pass
        return False

    def _sheet_thickness(self, p7, material):
        try:
            bodies = win32com.client.CastTo(p7, "ISheetMetalContainer").SheetMetalBodies
            if bodies.Count > 0:
                return float(bodies.SheetMetalBody(0).Thickness)
        except Exception:
            pass
        m = self.THICKNESS_RE.search(material or "")
        if m:
            return float(m.group(1).replace(",", "."))
        return None

    def _geometry_bent(self, p7, material):
        is_sheet = bool(self.SHEET_MATERIAL_RE.search(material or ""))
        is_tube = bool(self.TUBE_MATERIAL_RE.search(material or ""))
        try:
            is_sheet = is_sheet or win32com.client.CastTo(
                p7, "ISheetMetalContainer").SheetMetalBodies.Count > 0
        except Exception:
            pass
        if not (is_sheet or is_tube):
            return False
        try:
            faces = win32com.client.CastTo(p7, "IFeature7").ModelObjects(6) or ()  # o3d_face
        except Exception:
            return False

        thickness = self._sheet_thickness(p7, material) if is_sheet else None
        arcs = []  # (радиус, угол, длина вдоль оси, ось)
        for f in faces:
            try:
                f = win32com.client.CastTo(f, "IFace")
                if is_tube and f.IsTorus:
                    return True
                if not (is_sheet and f.IsCylinder):
                    continue
                ms = f.MathSurface
                angle = ms.ParamUMax - ms.ParamUMin
                # Отверстия и полуцилиндры (180 и больше) гибами не считаем
                if angle <= 0 or angle >= 3.13:
                    continue
                radius = f.Radius
                if radius <= 0:
                    continue
                length = f.GetArea(0) * 100.0 / (radius * angle)  # GetArea в см2
                axis = tuple(round(abs(v), 2) for v in ms.Placement.GetVector(2)[1:4])
                arcs.append((radius, angle, length, axis))
            except Exception:
                continue

        for i, (r1, a1, l1, ax1) in enumerate(arcs):
            for r2, a2, l2, ax2 in arcs[i + 1:]:
                dr = abs(r1 - r2)
                if dr < 0.3 or abs(a1 - a2) > 0.02 or ax1 != ax2:
                    continue
                if abs(l1 - l2) > 0.05 * max(l1, l2) + 0.5:
                    continue
                if thickness:
                    if abs(dr - thickness) > 0.15 * thickness + 0.05:
                        continue
                    # Скругления углов контура имеют длину, равную толщине
                    if min(l1, l2) < 1.5 * thickness:
                        continue
                elif min(l1, l2) < 2 * dr:
                    continue
                return True
        return False

    def read_positions(self):
        """Чтение позиций из всех .spw файлов."""
        print("\nЧтение позиций из спецификаций...")

        self.all_positions = {}
        self.spec_quantities = {}

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
                self.spec_quantities[os.path.basename(spw)] = {
                    d: p["quantity"] for d, p in positions.items()
                }
            except Exception as e:
                self._problem(f"Не прочитана спецификация {os.path.basename(spw)}: {e}")

        self.create_missing_specs()

        print(f"[OK] Найдено позиций: {len(self.all_positions)}")
        self.check_quantities()
        return True

    def _has_spec(self, marking):
        key = self._norm(marking)
        return any(self._norm(f).startswith(key) for f in self.spec_quantities)

    def create_missing_specs(self):
        """Для сборок, у которых не нашлось своей .spw, спецификация создается
        по модели (из объектов спецификации ее деталей) и сохраняется в папку
        "Генерированные спецификации". Сборки при этом не меняются: связь
        со спецификацией не передается в модель. Позиции и количества из созданной
        спецификации используются так же, как из обычной."""
        self.generated_specs = []
        missing = [(m, f) for m, f in self.assembly_files.items()
                   if m and f and not self._has_spec(m)]
        if not missing:
            return
        os.makedirs(self.specs_dir, exist_ok=True)
        for marking, asm_path in missing:
            safe = re.sub(r'[\\/:*?"<>|]', "_", marking).strip()
            spw_path = os.path.join(self.specs_dir, f"{safe} СП.spw")
            print(f"  Создание спецификации {os.path.basename(spw_path)}...")
            doc = None
            try:
                doc = win32com.client.CastTo(
                    self.app.Documents.AddWithDefaultSettings(3, False),  # ksDocumentSpecification
                    "ISpecificationDocument")
                doc.AttachedDocuments.AddDocument(asm_path, False, True, "")
                positions = self._read_spec_doc(doc, os.path.basename(spw_path))
                if not positions:
                    self._problem(f"Спецификация для {marking} создана пустой: у деталей "
                                  f"сборки нет объектов спецификации")
                    continue
                doc.SaveAs(spw_path)
                self.all_positions.update(
                    {d: p for d, p in positions.items() if d not in self.all_positions})
                self.spec_quantities[os.path.basename(spw_path)] = {
                    d: p["quantity"] for d, p in positions.items()
                }
                self.generated_specs.append(spw_path)
            except Exception as e:
                self._problem(f"Не удалось создать спецификацию для {marking}: {e}")
            finally:
                if doc is not None:
                    try:
                        doc.Close(0)
                    except Exception:
                        pass
        if self.generated_specs:
            print(f"[OK] Создано спецификаций: {len(self.generated_specs)} "
                  f"(папка {self.specs_dir})")

    @staticmethod
    def _parse_qty(text):
        m = re.match(r"\s*(\d+)", str(text or ""))
        return int(m.group(1)) if m else None

    @staticmethod
    def _norm(text):
        return re.sub(r"\s+", "", str(text or "")).lower()

    def check_quantities(self):
        """Сверка количества вставок в модели с колонкой "Кол." спецификации.
        Спецификация сопоставляется со сборкой по обозначению в начале имени .spw."""
        self.quantity_warnings = []
        checked = 0
        for parent_marking, counts in self.model_counts.items():
            key = self._norm(parent_marking)
            spec_name = None
            for fname in self.spec_quantities:
                if self._norm(fname).startswith(key):
                    spec_name = fname
                    break
            if not spec_name:
                continue
            checked += 1
            spec = self.spec_quantities[spec_name]
            for designation, qty_text in spec.items():
                spec_qty = self._parse_qty(qty_text)
                if spec_qty is None:
                    continue
                model_qty = counts.get(designation)
                if model_qty is None:
                    for m, c in counts.items():
                        if self._norm(m) == self._norm(designation):
                            model_qty = c
                            break
                if model_qty is None:
                    model_qty = 0
                if model_qty != spec_qty:
                    self.quantity_warnings.append(
                        f"{parent_marking}: {designation} — в модели {model_qty}, "
                        f"в спецификации {spec_qty}"
                    )

        print(f"Сверка количества со спецификациями: проверено сборок {checked}, "
              f"расхождений {len(self.quantity_warnings)}")
        for w in self.quantity_warnings:
            print(f"  [??] {w}")

    def _read_spw(self, spw_path):
        """Чтение позиций из .spw файла."""
        try:
            doc = self.app.Documents.Open(spw_path)
            self.opened_docs.append(doc)
            time.sleep(0.5)
            return self._read_spec_doc(doc, os.path.basename(spw_path))
        except Exception as e:
            self._problem(f"Не открыта спецификация {os.path.basename(spw_path)}: {e}")
            return {}

    def _read_spec_doc(self, doc, label):
        """Позиции из открытого документа спецификации: {обозначение: {...}}."""
        positions = {}
        try:
            sd = doc.SpecificationDescriptions
            if not sd or sd.Count == 0:
                self._problem(f"В спецификации {label} нет описания, "
                              f"позиции из нее не прочитаны")
                return positions

            spec_desc = sd.Item(0)

            # Проверяем нужно ли перестроить
            if spec_desc.NeedRebuild:
                print(f"  Перестроение спецификации: {label}...")
                spec_desc.Update()
                time.sleep(1)

            objects = spec_desc.Objects
            bad_rows = 0

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
                except Exception:
                    bad_rows += 1
                    continue
            if bad_rows:
                self._problem(f"В спецификации {label} не прочитано строк: {bad_rows}")
        except Exception as e:
            self._problem(f"Не прочитана спецификация {label}: {e}")

        return positions

    def close_all_docs(self):
        """Закрытие всех открытых документов."""
        for doc in self.opened_docs:
            try:
                if doc:
                    doc.Close(0)  # спецификации только читались: без сохранения
            except:
                pass
        self.opened_docs.clear()

    def generate_excel(self):
        """Генерация Excel по шаблону."""
        print("\nСоздание Excel файла...")

        wb = Workbook()
        ws = wb.active
        ws.title = os.path.splitext(self.assembly_name)[0][:31]

        tmpl = self.template.get("excel", {})
        columns = tmpl.get("columns", [])
        if not columns:
            columns = [
                {"key": "position", "header": "Позиция", "width": 8},
                {"key": "marking", "header": "Обозначение", "width": 25},
                {"key": "name", "header": "Наименование", "width": 40},
                {"key": "quantity", "header": "Кол-во", "width": 12},
                {"key": "mass_kg", "header": "Масса, кг", "width": 12},
                {"key": "material", "header": "Материал", "width": 25},
                {"key": "stock_length", "header": "Длина сортамента", "width": 12},
                {"key": "bending", "header": "Гибка", "width": 8},
            ]

        header_color = tmpl.get("header_color", "2F5496")
        assembly_color = tmpl.get("assembly_color", "D6E4F0")

        header_font = Font(bold=True, size=10, color="FFFFFF")
        header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
        assembly_fill = PatternFill(start_color=assembly_color, end_color=assembly_color, fill_type="solid")
        border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin"),
        )
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left = Alignment(horizontal="left", vertical="center", wrap_text=True)

        for col_idx, col_def in enumerate(columns, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = col_def.get("header", "")
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border
            col_letter = chr(64 + col_idx) if col_idx <= 26 else chr(64 + (col_idx - 1) // 26) + chr(65 + (col_idx - 1) % 26)
            ws.column_dimensions[col_letter].width = col_def.get("width", 15)

        if hasattr(self, 'cost_data') and self.cost_data:
            cost_cols = [
                {"key": "material_cost", "header": "Стоимость\nмат., руб", "width": 15},
                {"key": "labor_cost", "header": "Стоимость\nработ, руб", "width": 15},
                {"key": "total_cost", "header": "Итого, руб", "width": 15},
            ]
            for i, cc in enumerate(cost_cols):
                col_idx = len(columns) + i + 1
                cell = ws.cell(row=1, column=col_idx)
                cell.value = cc["header"]
                cell.font = header_font
                cell.fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
                cell.alignment = center
                cell.border = border
                col_letter = chr(64 + col_idx) if col_idx <= 26 else chr(64 + (col_idx - 1) // 26) + chr(65 + (col_idx - 1) % 26)
                ws.column_dimensions[col_letter].width = cc["width"]
            columns = columns + cost_cols

        ws.row_dimensions[1].height = 35

        grouped_data = self._grouped_data()

        row = 2
        cost_by_marking = {}
        if hasattr(self, 'cost_data') and self.cost_data:
            cost_by_marking = {c["marking"]: c for c in self.cost_data}

        for item in grouped_data:
            marking = item.get("marking", "")
            pos_info = self.all_positions.get(marking, {})
            mass_kg = (item.get("mass", 0) or 0) / 1000
            is_asm = item.get("is_assembly", False)

            material = item.get("material", "")
            if is_asm and is_default_material(material):
                material = ""

            row_data = {
                "position": pos_info.get("position", ""),
                "marking": marking,
                "name": item.get("name", ""),
                "quantity": item.get("quantity", 1),
                "mass_kg": round(mass_kg, 6) if mass_kg > 0 else "",
                "total_mass": round(mass_kg * item.get("quantity", 1), 6) if mass_kg > 0 else "",
                "material": material,
                "stock_length": item.get("stock_length", ""),
                "bending": "X" if item.get("is_bending") else "",
                "level": item.get("level", 0) + 1 if item.get("level", 0) > 0 else "",
                "is_assembly": is_asm,
            }

            if marking in cost_by_marking:
                c = cost_by_marking[marking]
                row_data["material_cost"] = c["material_cost"]
                row_data["labor_cost"] = c["labor_cost"]
                row_data["total_cost"] = c["total_cost"]

            for col_idx, col_def in enumerate(columns, 1):
                cell = ws.cell(row=row, column=col_idx)
                cell.value = row_data.get(col_def.get("key", ""), "")
                cell.border = border
                cell.alignment = left if col_def.get("key") in ("name", "material") else center
                if is_asm:
                    cell.fill = assembly_fill

            if is_asm and item.get("level", 0) > 0:
                ws.row_dimensions[row].outline_level = min(item["level"], 8)

            row += 1

        row += 1

        if hasattr(self, 'cost_summary') and self.cost_summary:
            cost_start = len(columns) - 2
            ws.cell(row=row, column=cost_start, value="Материалы:").font = Font(bold=True)
            ws.cell(row=row, column=cost_start + 1, value=self.cost_summary["total_material"]).font = Font(bold=True)
            row += 1
            ws.cell(row=row, column=cost_start, value="Работы:").font = Font(bold=True)
            ws.cell(row=row, column=cost_start + 1, value=self.cost_summary["total_labor"]).font = Font(bold=True)
            row += 1
            ws.cell(row=row, column=cost_start, value="ИТОГО:").font = Font(bold=True, color="8B0000")
            ws.cell(row=row, column=cost_start + 1, value=self.cost_summary["total"]).font = Font(bold=True, color="8B0000")

        for r in range(2, row):
            cell_a = ws.cell(row=r, column=1)
            if cell_a.value and isinstance(cell_a.value, int) and cell_a.value > 1:
                ws.row_dimensions[r].outline_level = min(cell_a.value - 1, 8)
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
                if not merged[key].get("stock_length") and item.get("stock_length"):
                    merged[key]["stock_length"] = item["stock_length"]
        return list(merged.values())

    def _grouped_data(self):
        """Группировка всех деталей и сборок с суммированием количества."""
        grouped = {}
        for item in self.all_data:
            key = f"{item.get('marking', '')}|{item.get('name', '')}"
            if key not in grouped:
                grouped[key] = item.copy()
            else:
                grouped[key]["quantity"] += item.get("quantity", 1)
                if not grouped[key].get("stock_length") and item.get("stock_length"):
                    grouped[key]["stock_length"] = item["stock_length"]
        return list(grouped.values())

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
        grouped_data = self._grouped_data()
        doc.add_paragraph(f'Всего элементов: {len(grouped_data)}')

        output_path = os.path.join(self.assembly_dir, f"{base_name}_сводная_ведомость.docx")
        doc.save(output_path)
        print(f"[OK] Word: {os.path.abspath(output_path)}")
        return output_path

    def generate_pdf(self):
        """Генерация PDF файла."""
        print("\nСоздание PDF файла...")

        pdf = FPDF(orientation="L", unit="mm", format="A4")
        pdf.set_auto_page_break(auto=True, margin=15)

        pdf.add_font("arial", "", "C:/Windows/Fonts/arial.ttf")
        pdf.add_font("arial", "B", "C:/Windows/Fonts/arialbd.ttf")
        font_name = "arial"

        pdf.add_page()

        base_name = os.path.splitext(self.assembly_name)[0]
        base_designation = base_name.split('.')[0]

        pdf_cols = [
            {"key": "position", "header": "Поз.", "width": 15, "max_chars": 8},
            {"key": "marking", "header": "Обозначение", "width": 35, "max_chars": 20},
            {"key": "name", "header": "Наименование", "width": 80, "max_chars": 45},
            {"key": "quantity", "header": "Кол-во", "width": 15, "max_chars": 8},
            {"key": "mass_kg", "header": "Масса, кг", "width": 18, "max_chars": 10},
            {"key": "material", "header": "Материал", "width": 50, "max_chars": 28},
            {"key": "stock_length", "header": "Длина сортамента", "width": 20, "max_chars": 12},
            {"key": "bending", "header": "Гибка", "width": 12, "max_chars": 5},
        ]

        header_color = self.template.get("excel", {}).get("header_color", "2F5496")
        assembly_color = self.template.get("excel", {}).get("assembly_color", "D6E4F0")

        r, g, b = int(header_color[0:2], 16), int(header_color[2:4], 16), int(header_color[4:6], 16)

        pdf.set_font(font_name, "B", 14)
        pdf.cell(0, 10, self.template.get("excel", {}).get("title", "СВОДНАЯ ВЕДОМОСТЬ"), ln=True, align="C")
        pdf.set_font(font_name, "", 9)
        pdf.cell(0, 7, f"Обозначение: {base_designation}   Наименование: {base_name}", ln=True, align="C")
        pdf.ln(4)

        pdf.set_font(font_name, "B", 7)
        pdf.set_fill_color(r, g, b)
        pdf.set_text_color(255, 255, 255)
        for col_def in pdf_cols:
            pdf.cell(col_def["width"], 7, col_def["header"], border=1, align="C", fill=True)
        pdf.ln()

        pdf.set_text_color(0, 0, 0)
        pdf.set_font(font_name, "", 6)

        ar, ag, ab = int(assembly_color[0:2], 16), int(assembly_color[2:4], 16), int(assembly_color[4:6], 16)

        row_h = 6
        grouped_data = self._grouped_data()
        for item in grouped_data:
            marking = item.get("marking", "")
            mass_kg = (item.get("mass", 0) or 0) / 1000
            is_asm = item.get("is_assembly", False)
            material = item.get("material", "")
            if is_asm and is_default_material(material):
                material = ""

            if is_asm:
                pdf.set_fill_color(ar, ag, ab)
                fill = True
            else:
                fill = False

            raw_row = {
                "position": self.all_positions.get(marking, {}).get("position", ""),
                "marking": marking,
                "name": item.get("name", ""),
                "quantity": str(item.get("quantity", 1)),
                "mass_kg": f"{mass_kg:.3f}" if mass_kg > 0 else "",
                "total_mass": f"{mass_kg * item.get('quantity', 1):.3f}" if mass_kg > 0 else "",
                "material": material,
                "stock_length": str(item.get("stock_length", "") or ""),
                "bending": "X" if item.get("is_bending") else "",
            }
            for col_def in pdf_cols:
                val = str(raw_row.get(col_def["key"], ""))
                max_c = col_def.get("max_chars", 30)
                if len(val) > max_c:
                    val = val[:max_c - 3] + "..."
                pdf.cell(col_def["width"], row_h, val, border=1, fill=fill)
            pdf.ln()

        pdf.ln(3)
        pdf.set_font(font_name, "B", 8)
        pdf.cell(0, 7, f"ИТОГО: {len(grouped_data)} элементов", ln=True)

        output_path = os.path.join(self.assembly_dir, f"{base_name}_сводная_ведомость.pdf")
        pdf.output(output_path)
        print(f"[OK] PDF: {os.path.abspath(output_path)}")
        return output_path

    def _add_table(self, doc, items):
        if not items:
            return
        tmpl_cols = self.template.get("word", {}).get("columns", [])
        if not tmpl_cols:
            tmpl_cols = [
                {"key": "position", "header": "Поз."},
                {"key": "marking", "header": "Обозначение"},
                {"key": "name", "header": "Наименование"},
                {"key": "quantity", "header": "Кол-во"},
                {"key": "mass_kg", "header": "Масса, кг"},
                {"key": "material", "header": "Материал"},
                {"key": "stock_length", "header": "Длина сортамента"},
                {"key": "bending", "header": "Гибка"},
            ]
        num_cols = len(tmpl_cols)
        table = doc.add_table(rows=len(items) + 1, cols=num_cols)
        table.style = 'Table Grid'
        for i, col_def in enumerate(tmpl_cols):
            cell = table.cell(0, i)
            cell.text = col_def.get("header", "")
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.bold = True
                    r.font.size = Pt(9)
        for idx, item in enumerate(items, 1):
            marking = item.get("marking", "")
            mass_kg = (item.get("mass", 0) or 0) / 1000
            is_asm = item.get("is_assembly", False)
            material = item.get("material", "")
            if is_asm and is_default_material(material):
                material = ""
            row_data = {
                "position": self.all_positions.get(marking, {}).get("position", str(idx)),
                "marking": marking,
                "name": item.get("name", ""),
                "quantity": str(item.get("quantity", 1)),
                "mass_kg": f"{mass_kg:.3f}" if mass_kg > 0 else "",
                "total_mass": f"{mass_kg * item.get('quantity', 1):.3f}" if mass_kg > 0 else "",
                "material": material,
                "stock_length": str(item.get("stock_length", "") or ""),
                "bending": "X" if item.get("is_bending") else "",
            }
            for col_idx, col_def in enumerate(tmpl_cols):
                cell = table.cell(idx, col_idx)
                cell.text = row_data.get(col_def.get("key", ""), "")
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(9)

    def calculate_costs(self, prices=None):
        """Расчёт стоимости по материалам и трудозатратам."""
        if prices is None:
            prices_path = os.path.join(os.path.dirname(__file__), "prices.json")
            if os.path.exists(prices_path):
                import json
                with open(prices_path, "r", encoding="utf-8") as f:
                    prices = json.load(f)
            else:
                prices = {}

        materials = prices.get("materials", {})
        default_price = prices.get("default_material_price", 100.0)
        labor = prices.get("labor", {})

        self.cost_data = []
        total_material = 0
        total_labor = 0

        for item in self._grouped_data():
            mass_kg = (item.get("mass", 0) or 0) / 1000
            quantity = item.get("quantity", 1)
            material = item.get("material", "").lower().strip()
            is_asm = item.get("is_assembly", False)
            is_bending = item.get("is_bending", False)

            if is_asm and is_default_material(item.get("material", "")):
                material = ""

            mat_price = default_price
            for mat_name, price in materials.items():
                if material.startswith(mat_name.lower()):
                    mat_price = price
                    break

            mat_cost = mass_kg * mat_price * quantity
            labor_cost = 0
            if not is_asm:
                labor_cost += mass_kg * labor.get("cutting_per_kg", 0) * quantity
                if is_bending:
                    labor_cost += mass_kg * labor.get("bending_per_kg", 0) * quantity
                labor_cost += mass_kg * labor.get("welding_per_kg", 0) * quantity
                labor_cost += mass_kg * labor.get("assembly_per_kg", 0) * quantity

            total = mat_cost + labor_cost
            total_material += mat_cost
            total_labor += labor_cost

            self.cost_data.append({
                "marking": item.get("marking", ""),
                "name": item.get("name", ""),
                "mass_kg": mass_kg,
                "quantity": quantity,
                "material_cost": round(mat_cost, 2),
                "labor_cost": round(labor_cost, 2),
                "total_cost": round(total, 2),
            })

        self.cost_summary = {
            "total_material": round(total_material, 2),
            "total_labor": round(total_labor, 2),
            "total": round(total_material + total_labor, 2),
        }
        return self.cost_data, self.cost_summary

    def analyze(self, filepath=None):
        """Подключение к КОМПАС, открытие сборки, чтение дерева и позиций."""
        if not self.connect():
            return False
        if not self.open_assembly(filepath):
            return False
        if not self.extract_tree():
            return False
        # Вопросы КОМПАС при открытии и создании спецификаций ("Перестроить
        # спецификацию?" и т.п.) останавливали бы фоновую работу: на время чтения
        # отвечаем на них "Нет" автоматически (перестроение делаем сами).
        old_hide = None
        try:
            old_hide = self.app.HideMessage
            self.app.HideMessage = 2  # ksHideMessageNo
        except Exception:
            pass
        try:
            self.read_positions()
        finally:
            if old_hide is not None:
                try:
                    self.app.HideMessage = old_hide
                except Exception:
                    pass
        return True

    def generate(self, formats=None):
        """Формирование документов по выбранным форматам."""
        if formats is None:
            formats = {"excel", "word", "pdf"}
        results = {}
        if "excel" in formats:
            results["excel"] = self.generate_excel()
        if "word" in formats:
            results["word"] = self.generate_word()
        if "pdf" in formats:
            results["pdf"] = self.generate_pdf()
        print("\n" + "=" * 50)
        print("Готово!")
        for fmt, path in results.items():
            print(f"{fmt.capitalize()}: {path}")
        print("=" * 50)
        return results

    def run(self, filepath=None, formats=None):
        if formats is None:
            formats = {"excel", "word", "pdf"}
        try:
            if not self.analyze(filepath):
                return
            return self.generate(formats)
        finally:
            self.close_all_docs()


def main():
    filepath = sys.argv[1] if len(sys.argv) > 1 else None
    exporter = KompasExportFinal()
    exporter.run(filepath)


if __name__ == "__main__":
    main()
