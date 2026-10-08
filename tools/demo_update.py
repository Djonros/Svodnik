"""Показ окна обновления на тестовых данных (проверка интерфейса, сеть не нужна)."""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "app"))
import gui_app
import updater

updater.check_latest = lambda timeout=10: {
    "version": "1.8.1",
    "notes": "- Исправлена длина у гнутых труб\n- Ускорено чтение больших сборок\n- Мелкие исправления окна",
    "page": updater.RELEASES_PAGE,
    "asset_url": "https://example.invalid/Svodnik.exe",
    "size": 44000000,
}
updater.can_self_update = lambda: True

app = gui_app.KompasExportApp()
app._auto_check_updates = lambda: None
app._check_updates(manual=True)


def save_hwnd():
    for w in app.root.winfo_children():
        if w.winfo_class() == "Toplevel":
            with open(os.path.join(os.path.dirname(__file__), "hwnd.txt"), "w") as f:
                f.write(str(int(w.wm_frame(), 16)))


app.root.after(2500, save_hwnd)
app.run()
