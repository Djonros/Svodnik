"""Диагностика: параметры цилиндрических граней (ось, угол) - только чтение."""
import sys
import win32com.client
app = win32com.client.Dispatch("Kompas.Application.7")
d = app.Documents.Open(sys.argv[1], False, True)
try:
    top = win32com.client.CastTo(win32com.client.CastTo(d, "IKompasDocument3D").TopPart, "IPart7")
    feat = win32com.client.CastTo(top, "IFeature7")
    for f in (feat.ModelObjects(6) or ()):
        f = win32com.client.CastTo(f, "IFace")
        if not (f.IsCylinder or f.IsTorus):
            continue
        ms = f.MathSurface
        pl = ms.Placement
        try:
            ar = f.GetArea(0)
        except Exception as e:
            ar = str(e)[:40]
        print("cyl" if f.IsCylinder else "torus", round(f.Radius, 2), "A", ar, "U", round(ms.ParamUMin, 3), round(ms.ParamUMax, 3),
              "V", round(ms.ParamVMin, 2), round(ms.ParamVMax, 2),
              "O", [round(x, 2) for x in pl.GetOrigin()],
              "X", [round(x, 3) for x in pl.GetVector(1)], "Y", [round(x, 3) for x in pl.GetVector(2)],
              "Z", [round(x, 3) for x in pl.GetVector(3)])
finally:
    d.Close(0)
