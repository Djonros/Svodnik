"""Показ карточки результата на тестовых данных (проверка интерфейса, КОМПАС не нужен)."""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "app"))
import gui_app


class FakeExporter:
    problems = ["Не определена длина сортамента у 'Труба угловая'"]
    quantity_warnings = ["2019-2811130: 2019-2811141 — в модели 1, в спецификации 2"]
    generated_specs = [r"C:\tmp\Генерированные спецификации\2019-2811130 СП.spw"]
    all_data = [{"is_assembly": False, "is_bending": i % 3 == 0} for i in range(14)]


app = gui_app.KompasExportApp()
app.show_log.set(True)
app._rebuild_panels()
app._log("[OK] Подключено к КОМПАС-3D v24")
app._log("[!!] Не определена длина сортамента у 'Труба угловая'")
app._log("  [??] 2019-2811130: 2019-2811141 — в модели 1, в спецификации 2")
app._set_stage("Готово", 100)
app._show_issues = lambda: None
app._show_result(FakeExporter(), {"excel": r"C:\tmp\Лонжерон_сводная_ведомость.xlsx",
                                  "pdf": r"C:\tmp\Лонжерон_сводная_ведомость.pdf"})
app.run()
