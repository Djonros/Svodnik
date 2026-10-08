"""Список членов интерфейсов листового металла из библиотеки типов КОМПАС API7."""
import glob, sys, pythoncom
paths = glob.glob(r"C:\Program Files\ASCON\**\kAPI7.tlb", recursive=True)
print(paths)
tlb = pythoncom.LoadTypeLib(paths[0])
want = sys.argv[1:] or ["ISheetMetalContainer", "ISheetMetalBody", "ISheetMetalBend", "ISheetMetalBends", "IPart7"]
for i in range(tlb.GetTypeInfoCount()):
    name = tlb.GetDocumentation(i)[0]
    if name in want or any(w.endswith("*") and name.startswith(w[:-1]) for w in want):
        ti = tlb.GetTypeInfo(i)
        attr = ti.GetTypeAttr()
        names = set()
        for f in range(attr.cFuncs):
            fd = ti.GetFuncDesc(f)
            names.add(ti.GetNames(fd.memid)[0])
        print(name, sorted(names))
