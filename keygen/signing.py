"""
Подпись лицензионных ключей Сводника. Только для распространителя.

Закрытый ключ подписи хранится в файле SIGNING_KEY_FILE на компьютере
распространителя и НИКОГДА не попадает в репозиторий. Сделайте копию этого
файла: без него нельзя выдавать ключи для уже выпущенных версий программы.
Открытый ключ, парный к нему, записан в app/license.py (PUBLIC_KEY).
"""
import base64
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "app"))

import ed25519  # noqa: E402
from license import LicenseManager  # noqa: E402

SIGNING_KEY_FILE = os.environ.get(
    "SVODNIK_SIGNING_KEY", os.path.join(os.path.expanduser("~"), ".svodnik_signing_key"))


def load_secret() -> bytes:
    """Прочитать закрытый ключ и убедиться, что он парный к PUBLIC_KEY программы."""
    if not os.path.exists(SIGNING_KEY_FILE):
        raise RuntimeError(f"Не найден файл ключа подписи:\n{SIGNING_KEY_FILE}")
    with open(SIGNING_KEY_FILE, "r", encoding="ascii") as f:
        secret = bytes.fromhex(f.read().strip())
    if ed25519.public_key(secret).hex() != LicenseManager.PUBLIC_KEY:
        raise RuntimeError("Ключ подписи не подходит к этой версии программы:\n"
                           "открытый ключ в app/license.py другой.")
    return secret


def create_secret() -> str:
    """Создать новый ключ подписи. Вернуть открытый ключ для app/license.py."""
    if os.path.exists(SIGNING_KEY_FILE):
        raise RuntimeError(f"Ключ подписи уже есть, перезапись запрещена:\n{SIGNING_KEY_FILE}")
    secret = os.urandom(32)
    with open(SIGNING_KEY_FILE, "w", encoding="ascii") as f:
        f.write(secret.hex())
    return ed25519.public_key(secret).hex()


def make_key(secret: bytes, machine_id: str, expiry: str) -> str:
    """Лицензионный ключ SV-ГГГГ-ММ-ДД-ПОДПИСЬ для компьютера machine_id."""
    signature = ed25519.sign(secret, LicenseManager.signed_message(machine_id, expiry))
    return f"{LicenseManager.KEY_PREFIX}-{expiry}-" + base64.b32encode(signature).decode().rstrip("=")
