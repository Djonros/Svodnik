"""Диагностика признаков гибки: для каждой детали активной сборки печатает
число листовых тел и операций сгиба (только чтение, ничего не сохраняет)."""
import sys
import win32com.client

COLL = ["SheetMetalBodies", "SheetMetalBends", "SheetMetalLineBends", "SheetMetalSketchBends",
        "SheetMetalFlangings", "SheetMetalShoulders", "SheetMetalRuledShells",
        "SheetMetalLinearRuledShells", "SheetMetalBendedStraightens", "SheetMetalPressFormings",
        "ConvertsToSheetMetals"]

app = win32com.client.Dispatch("Kompas.Application.7")
doc = app.Documents.Open(sys.argv[1], False, True) if len(sys.argv) > 1 else app.ActiveDocument
print("Документ:", doc.Name)
doc3d = win32com.client.CastTo(doc, "IKompasDocument3D")
seen = set()
LIMIT = int(sys.argv[2]) if len(sys.argv) > 2 else 60


def counts(p7):
    res = {}
    try:
        c = win32com.client.CastTo(p7, "ISheetMetalContainer")
    except Exception as e:
        return {"err": str(e)[:60]}
    for n in COLL:
        try:
            res[n] = getattr(c, n).Count
        except Exception as e:
            res[n] = "x"
    return res


def walk(part, indent=""):
    parts = part.Parts
    for i in range(parts.Count if parts else 0):
        if len(seen) >= LIMIT:
            return
        p7 = win32com.client.CastTo(parts.Item(i), "IPart7")
        fn = (p7.FileName or "").lower()
        if p7.Parts and p7.Parts.Count > 0:
            walk(p7, indent + "  ")
            continue
        if fn in seen:
            continue
        seen.add(fn)
        c = {k: v for k, v in counts(p7).items() if v not in (0,)}
        f = {}
        try:
            d = app.Documents.Open(p7.FileName, False, True)
            top = win32com.client.CastTo(win32com.client.CastTo(d, "IKompasDocument3D").TopPart, "IPart7")
            f = {k: v for k, v in counts(top).items() if v not in (0,)}
            d.Close(0)
        except Exception as e:
            f = {"open_err": str(e)[:60]}
        print(f"{indent}{p7.Marking} {p7.Name} | {p7.Material[:30] if p7.Material else ''} | in-asm {c} | file {f}")


walk(doc3d.TopPart)

if len(sys.argv) > 1:
    doc.Close(0)
