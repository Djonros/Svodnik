"""Проверка создания недостающих спецификаций. Сборка открывается только для
чтения, созданные .spw пишутся во временную папку (второй аргумент)."""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "app"))
import kompas_export_final as kef

exp = kef.KompasExportFinal()
exp.connect()
doc = exp.app.Documents.Open(sys.argv[1], False, True)
exp.doc = doc
exp.assembly_name = doc.Name
exp.assembly_dir = os.path.dirname(doc.PathName)
exp.specs_dir = sys.argv[2]
exp.extract_tree()
print("Сборки:", exp.assembly_files)
exp.app.HideMessage = 2
try:
    exp.read_positions()
finally:
    exp.app.HideMessage = 0
print("Создано:", exp.generated_specs)
print("Позиций:", len(exp.all_positions), list(exp.all_positions.items())[:5])
print("Расхождения:", exp.quantity_warnings)
print("Проблемы:", exp.problems)
exp.close_all_docs()
doc.Close(0)
