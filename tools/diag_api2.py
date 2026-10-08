"""Типы и перечисления из библиотеки типов КОМПАС API7 (только чтение)."""
import glob, sys, pythoncom
tlb = pythoncom.LoadTypeLib(glob.glob(r"C:\Program Files\ASCON\**\Bin\kAPI7.tlb", recursive=True)[0])
for i in range(tlb.GetTypeInfoCount()):
    name = tlb.GetDocumentation(i)[0]
    if name not in sys.argv[1:]:
        continue
    ti = tlb.GetTypeInfo(i)
    a = ti.GetTypeAttr()
    if a.typekind == pythoncom.TKIND_ENUM:
        print(name, [(ti.GetNames(ti.GetVarDesc(v).memid)[0], ti.GetVarDesc(v).value) for v in range(a.cVars)])
    else:
        print(name, [(ti.GetNames(ti.GetFuncDesc(f).memid)) for f in range(a.cFuncs) if ti.GetNames(ti.GetFuncDesc(f).memid)[0] not in ("AddRef","Release","QueryInterface","GetIDsOfNames","GetTypeInfo","GetTypeInfoCount","Invoke")])
