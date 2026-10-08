"""Проверка определения гибки на сборке: печатает детали и признак гибки.
Сборка открывается только для чтения и закрывается без сохранения."""
import os, sys, time
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "app"))
import win32com.client
import kompas_export_final as kef

exp = kef.KompasExportFinal()
exp.connect()
doc = exp.app.Documents.Open(sys.argv[1], False, True)
exp.doc = doc
exp.assembly_name = doc.Name
exp.assembly_dir = os.path.dirname(doc.PathName)
_sm, _geo = exp._sheet_metal_bent, exp._geometry_bent
def sm(p7):
    r = _sm(p7)
    if r: print("  [листовые операции]", p7.Name)
    return r
def geo(p7, m):
    r = _geo(p7, m)
    if r: print("  [геометрия]", p7.Name)
    return r
exp._sheet_metal_bent, exp._geometry_bent = sm, geo
t = time.time()
exp.extract_tree()
dt = time.time() - t
seen = set()
for it in exp.all_data:
    if it["is_assembly"]:
        continue
    k = (it["marking"], it["name"])
    if k in seen:
        continue
    seen.add(k)
    print(("ГИБ " if it["is_bending"] else "    ") + f"{it['marking']} {it['name']} | {it['material'][:40]} | L={it.get('stock_length', '')}")
print(f"Деталей: {len(seen)}, гнутых: {sum(1 for it in exp.all_data if it['is_bending'] and not it['is_assembly'])} вхождений, время обхода {dt:.1f} c")
print("Проблемы:", exp.problems[:5])
doc.Close(0)
