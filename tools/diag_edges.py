"""Диагностика: самое длинное прямое ребро деталей сборки (только чтение)."""
import sys
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
        if p7.FileName in seen:
            continue
        seen.add(p7.FileName)
        best = 0
        try:
            for e in (win32com.client.CastTo(p7, "IFeature7").ModelObjects(7) or ()):  # o3d_edge
                e = win32com.client.CastTo(e, "IEdge")
                if e.IsStraight:
                    best = max(best, e.GetLength(0))
        except Exception as ex:
            best = str(ex)[:60]
        print(p7.Name[:70], "|", best)


walk(win32com.client.CastTo(doc, "IKompasDocument3D").TopPart)
doc.Close(0)
