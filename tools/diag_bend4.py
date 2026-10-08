"""Диагностика: дерево операций и цилиндрические грани деталей (только чтение)."""
import glob, os, sys
import win32com.client

app = win32com.client.Dispatch("Kompas.Application.7")
files = []
for a in sys.argv[1:]:
    files += glob.glob(os.path.join(a, "*.m3d")) if os.path.isdir(a) else [a]

for fn in files:
    d = app.Documents.Open(fn, False, True)
    try:
        top = win32com.client.CastTo(win32com.client.CastTo(d, "IKompasDocument3D").TopPart, "IPart7")
        feat = win32com.client.CastTo(top, "IFeature7")
        ops = []
        try:
            for sf in (feat.SubFeatures(1, False, False) or ()):
                sf = win32com.client.CastTo(sf, "IFeature7")
                ops.append(f"{sf.Name}")
        except Exception as e:
            ops = [f"err {e}"[:80]]
        radii = []
        try:
            for f in (feat.ModelObjects(6) or ()):
                f = win32com.client.CastTo(f, "IFace")
                if f.IsCylinder:
                    radii.append(round(f.Radius, 2) if hasattr(f, "Radius") else "?")
        except Exception as e:
            radii = [f"err {e}"[:80]]
        print(os.path.basename(fn), "|", top.Material, "| ops:", ops[:25], "| cyl:", sorted(radii)[:30])
    finally:
        d.Close(0)
