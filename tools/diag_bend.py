"""
Диагностика статуса развертки листовых тел в КОМПАС-3D.

Запуск:
    python diag_bend.py [C:\\path\\to\\detail.m3d]

Без аргумента используется активный документ в КОМПАС.
Скрипт обходит дерево модели, находит листовые тела
(ISheetMetalContainer -> SheetMetalBodies) и печатает для каждого
тела свойство Straighten ("Разогнуть тело" = развертка включена).
"""

import sys

import win32com.client


def dump_part(part, indent=""):
    try:
        p7 = win32com.client.CastTo(part, "IPart7")
    except Exception:
        p7 = part

    name = getattr(p7, "Name", "?")
    marking = getattr(p7, "Marking", "")
    print(f"{indent}Деталь: {name} ({marking})")

    try:
        children = p7.Parts
        if children and children.Count > 0:
            for i in range(children.Count):
                dump_part(children.Item(i), indent + "    ")
    except Exception:
        pass

    container = None
    try:
        container = win32com.client.CastTo(p7, "ISheetMetalContainer")
    except Exception as e:
        print(f"{indent}  [--] ISheetMetalContainer недоступен: {e}")

    if container is None:
        return

    try:
        bodies = container.SheetMetalBodies
        count = bodies.Count
        print(f"{indent}  Листовых тел: {count}")
        for i in range(count):
            try:
                body = bodies.SheetMetalBody(i)
            except Exception:
                try:
                    body = bodies.Item(i)
                except Exception as e:
                    print(f"{indent}    [{i}] тело не получено: {e}")
                    continue
            body_name = getattr(body, "Name", f"тело {i}")
            try:
                straighten = bool(body.Straighten)
            except Exception as e:
                straighten = f"ошибка чтения: {e}"
            print(f"{indent}    [{i}] {body_name}: Straighten={straighten}")
    except Exception as e:
        print(f"{indent}  [!!] SheetMetalBodies недоступна: {e}")


def main():
    filepath = sys.argv[1] if len(sys.argv) > 1 else None
    app = win32com.client.Dispatch("Kompas.Application.7")
    app.Visible = True

    if filepath:
        doc = app.Documents.Open(filepath)
    else:
        doc = app.ActiveDocument

    if not doc:
        print("[!!] Нет открытого документа")
        return

    print(f"Документ: {doc.Name}")
    try:
        doc3d = win32com.client.CastTo(doc, "IKompasDocument3D")
    except Exception:
        doc3d = doc
    dump_part(doc3d.TopPart)


if __name__ == "__main__":
    main()
