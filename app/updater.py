"""
Проверка и установка обновлений Сводника.

Новые версии выкладываются как релизы GitHub в открытом репозитории
RELEASES_REPO: тег вида v1.7.2, описание релиза = список изменений,
вложение Svodnik.exe.
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import urllib.request

RELEASES_REPO = "Djonros/Svodnik-releases"
API_LATEST = f"https://api.github.com/repos/{RELEASES_REPO}/releases/latest"
RELEASES_PAGE = f"https://github.com/{RELEASES_REPO}/releases"
ASSET_NAME = "Svodnik.exe"
USER_AGENT = "Svodnik-updater"


def parse_version(text):
    """'v1.7.2' -> (1, 7, 2). Нечисловые хвосты отбрасываются."""
    nums = re.findall(r"\d+", text or "")
    return tuple(int(n) for n in nums[:4])


def is_newer(remote, local):
    return parse_version(remote) > parse_version(local)


def check_latest(timeout=10):
    """Вернуть сведения о последнем релизе или None, если релизов нет.

    Сетевые ошибки пробрасываются: вызывающий решает, показывать ли их.
    """
    req = urllib.request.Request(API_LATEST, headers={
        "User-Agent": USER_AGENT,
        "Accept": "application/vnd.github+json",
    })
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise
    asset_url = None
    size = 0
    for asset in data.get("assets", []):
        if asset.get("name", "").lower() == ASSET_NAME.lower():
            asset_url = asset.get("browser_download_url")
            size = asset.get("size", 0)
            break
    return {
        "version": (data.get("tag_name") or "").lstrip("vV"),
        "notes": (data.get("body") or "").strip(),
        "page": data.get("html_url") or RELEASES_PAGE,
        "asset_url": asset_url,
        "size": size,
    }


def download(url, dest, progress=None, timeout=30):
    """Скачать файл. progress(done_bytes, total_bytes) вызывается из этого же потока."""
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp, open(dest, "wb") as f:
        total = int(resp.headers.get("Content-Length") or 0)
        done = 0
        while True:
            chunk = resp.read(256 * 1024)
            if not chunk:
                break
            f.write(chunk)
            done += len(chunk)
            if progress:
                progress(done, total)
    if total and done != total:
        raise IOError(f"Файл скачан не полностью: {done} из {total} байт")
    with open(dest, "rb") as f:
        if f.read(2) != b"MZ":
            raise IOError("Скачанный файл не является программой Windows")


def can_self_update():
    """Самообновление возможно только у собранного exe в папке с правом записи."""
    if not getattr(sys, "frozen", False):
        return False
    folder = os.path.dirname(sys.executable)
    probe = os.path.join(folder, ".svodnik_write_test")
    try:
        with open(probe, "w") as f:
            f.write("1")
        os.remove(probe)
        return True
    except OSError:
        return False


def temp_download_path():
    return os.path.join(tempfile.gettempdir(), "Svodnik_update.exe")


def install_and_restart(new_exe):
    """Запустить скрипт, который дождется выхода программы, заменит exe и запустит новый.

    Старая версия сохраняется рядом как Svodnik_old.exe. После вызова
    программа должна сразу завершиться.
    """
    target = sys.executable
    backup = os.path.splitext(target)[0] + "_old.exe"
    script = os.path.join(tempfile.gettempdir(), "svodnik_update.cmd")
    # Загрузчик PyInstaller выходит чуть позже процесса Python, поэтому после
    # ожидания еще пауза. Пути передаются аргументами: так кириллица в пути не зависит от кодировки cmd
    lines = [
        "@echo off",
        "set /a N=0",
        ":wait",
        "tasklist /FI \"PID eq %~4\" 2>nul | find \"%~4\" >nul",
        "if not errorlevel 1 (",
        "  set /a N+=1",
        "  if %N% GEQ 60 goto fail",
        "  ping -n 2 127.0.0.1 >nul",
        "  goto wait",
        ")",
        "ping -n 4 127.0.0.1 >nul",
        "del /f /q \"%~2\" >nul 2>&1",
        "move /y \"%~1\" \"%~2\" >nul || goto fail",
        "move /y \"%~3\" \"%~1\" >nul || (move /y \"%~2\" \"%~1\" >nul & goto fail)",
        "start \"\" \"%~1\"",
        "(goto) 2>nul & del \"%~f0\"",
        ":fail",
        "start \"\" \"%~1\"",
        "(goto) 2>nul & del \"%~f0\"",
    ]
    with open(script, "w", encoding="ascii") as f:
        f.write("\r\n".join(lines) + "\r\n")
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    # Новый exe должен распаковаться заново, а не взять окружение старого
    env = dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT="1")
    subprocess.Popen(["cmd", "/c", script, target, backup, new_exe, str(os.getpid())],
                     creationflags=flags, close_fds=True, env=env,
                     stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
