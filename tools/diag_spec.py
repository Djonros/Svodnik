"""Опыт: создать спецификацию по сборке в отдельную папку и прочитать строки.
Сборка открывается только для чтения, исходные файлы не меняются."""
import os, sys, time
import win32com.client
asm, out_dir = sys.argv[1], sys.argv[2]
os.makedirs(out_dir, exist_ok=True)
app = win32com.client.Dispatch("Kompas.Application.7")
asm_doc = app.Documents.Open(asm, False, True)
spc = app.Documents.AddWithDefaultSettings(3, False)
print("тип документа:", spc.DocumentType)
sd = win32com.client.CastTo(spc, "ISpecificationDocument")
att = sd.AttachedDocuments
r = att.AddDocument(asm, False, True, "")
print("AddDocument:", r, "прикреплено:", att.Count)
descs = sd.SpecificationDescriptions
print("описаний:", descs.Count)
desc = descs.Active if descs.Count else None
if desc is not None:
    desc.Update()
    rows = 0
    for obj in desc.Objects:
        try:
            cols = obj.Columns
            vals = []
            for c in range(1, min(cols.Count, 7)):
                col = cols.Item(c)
                vals.append(col.Text.Str if col.Text else "")
            print("  ", vals)
            rows += 1
        except Exception as e:
            print("  err", e)
    print("строк:", rows)
path = os.path.join(out_dir, "test СП.spw")
print("SaveAs:", sd.SaveAs(path), os.path.exists(path))
sd.Close(0)
asm_doc.Close(0)

