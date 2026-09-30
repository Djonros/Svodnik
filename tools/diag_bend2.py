"""Расширенная диагностика: альтернативные пути доступа к листовым телам."""
import sys

import win32com.client

filepath = sys.argv[1]
app = win32com.client.Dispatch("Kompas.Application.7")
app.Visible = True
doc = app.Documents.Open(filepath)

out = []


def log(s):
    out.append(str(s))


log(f"Документ: {doc.Name}")

doc3d = win32com.client.CastTo(doc, "IKompasDocument3D")
top = doc3d.TopPart
p7 = win32com.client.CastTo(top, "IPart7")
log(f"TopPart: Name={p7.Name} Marking={p7.Marking}")

# 1. Контейнер от документа
for obj_name, obj in (("doc", doc), ("doc3d", doc3d), ("top", top), ("p7", p7)):
    try:
        c = win32com.client.CastTo(obj, "ISheetMetalContainer")
        bodies = c.SheetMetalBodies
        log(f"[{obj_name}] ISheetMetalContainer OK, Count={bodies.Count}")
    except Exception as e:
        log(f"[{obj_name}] ISheetMetalContainer FAIL: {e}")

# 2. Прямые обращения при Count=0
try:
    bodies = win32com.client.CastTo(p7, "ISheetMetalContainer").SheetMetalBodies
    for idx in (0, 1):
        try:
            b = bodies.SheetMetalBody(idx)
            log(f"SheetMetalBody({idx}) OK: Name={getattr(b, 'Name', '?')} Straighten={bool(b.Straighten)}")
        except Exception as e:
            log(f"SheetMetalBody({idx}) FAIL: {e}")
        try:
            b = bodies.Item(idx)
            log(f"Item({idx}) OK: Name={getattr(b, 'Name', '?')} Straighten={bool(b.Straighten)}")
        except Exception as e:
            log(f"Item({idx}) FAIL: {e}")
except Exception as e:
    log(f"bodies FAIL: {e}")

# 3. Дерево: компоненты и операции детали
try:
    parts = p7.Parts
    log(f"p7.Parts Count={parts.Count if parts else 0}")
except Exception as e:
    log(f"p7.Parts FAIL: {e}")

# 4. BodyContainer через IPart7 -> IBodyContainer?
for iface in ("IBodyContainer", "IModelContainer", "IPart7"):
    try:
        o = win32com.client.CastTo(p7, iface)
        log(f"CastTo {iface} OK: {type(o)}")
    except Exception as e:
        log(f"CastTo {iface} FAIL: {e}")

# 5. Все свойства/методы объекта части, содержащие sheet/metal/bend/unsheet
names = [n for n in dir(p7) if any(k in n.lower() for k in ("sheet", "metal", "bend", "unfold", "straight"))]
log(f"p7 attrs: {names}")

with open(r"C:\Users\djonr\AppData\Local\Temp\rufp\diag2_out.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("done")
