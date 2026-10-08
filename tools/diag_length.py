"""Диагностика длины сортамента: габарит детали в сборке и переменные (только чтение)."""
import os, sys
import win32com.client
app = win32com.client.Dispatch("Kompas.Application.7")
doc = app.Documents.Open(sys.argv[1], False, True)
seen = set()


def walk(part):
    parts = part.Parts
    for i in range(parts.Count if parts else 0):
        p7 = win32com.client.CastTo(parts.Item(i), "IPart7")
        if p7.Parts and p7.Parts.Count > 0:
            walk(p7)
            continue
        key = (p7.FileName or "").lower()
        if key in seen:
            continue
        seen.add(key)
        mat = p7.Material or ""
        try:
            g = p7.GetGabarit(True, True)
            dims = sorted(round(abs(g[i + 4] - g[i + 1]), 1) for i in range(3))
        except Exception as e:
            dims = str(e)[:50]
        vs = []
        try:
            vt = p7.VariableTable
            for j in range(vt.Count):
                v = vt.Variable(False, j)
                vs.append(f"{v.Name}={round(v.Value, 1)}")
        except Exception as e:
            vs = [str(e)[:40]]
        print(f"{p7.Name[:60]} | {mat[:35]} | габарит {dims} | перем {vs[:8]}")


walk(win32com.client.CastTo(doc, "IKompasDocument3D").TopPart)
doc.Close(0)
