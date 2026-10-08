"""
Модуль лицензирования Сводника.
Поддержка пробного периода и полной лицензии.

Ключ подписан закрытым ключом Ed25519, который хранится только у
распространителя (генератор ключей в keygen/). В программе лежит лишь
открытый ключ PUBLIC_KEY: им можно проверить ключ, но нельзя создать новый.
Формат ключа: SV-ГГГГ-ММ-ДД-<подпись base32>.
"""

import base64
import hashlib
import json
import os
import platform
import subprocess
from datetime import datetime

import ed25519


class LicenseManager:
    """Менеджер лицензий с привязкой к оборудованию."""

    PUBLIC_KEY = "4c9aafabd94ce5daf3ea761edeb590a810555e2978a0049b1ec691707e27d1c6"
    KEY_PREFIX = "SV"
    LICENSE_FILE = os.path.join(os.path.expanduser("~"), ".svodnik_license.dat")
    TRIAL_DAYS = 14
    MAX_EXPORTS_TRIAL = 5  # Максимум экспорта в пробном режиме

    def __init__(self):
        self.license_data = {}
        self.is_licensed = False
        self.old_key = False
        self.days_left = 0
        self.exports_count = 0
        self.load_license()

    _NO_WINDOW = 0x08000000  # CREATE_NO_WINDOW: не показывать окно консоли
    _cached_ids = None

    @classmethod
    def _command_value(cls, cmd) -> str:
        """Последняя строка вывода команды (так раньше разбирался вывод wmic:
        при нескольких дисках берется серийный номер последнего)."""
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30,
                                creationflags=cls._NO_WINDOW)
        return result.stdout.strip().split("\n")[-1].strip()

    def _hardware_id(self) -> str:
        """ID по UUID компьютера и серийному номеру диска.

        Сначала wmic (как в прежних версиях, чтобы выданные ключи подходили),
        если wmic нет (Windows 11 24H2 и новее) — те же значения через
        PowerShell/CIM. Хэш получается тот же самый."""
        try:
            uuid = self._command_value(["wmic", "csproduct", "get", "uuid"])
            serial = self._command_value(["wmic", "diskdrive", "get", "serialnumber"])
        except (OSError, subprocess.SubprocessError):
            ps = ["powershell", "-NoProfile", "-NonInteractive", "-Command"]
            uuid = self._command_value(
                ps + ["(Get-CimInstance Win32_ComputerSystemProduct).UUID"])
            serial = self._command_value(
                ps + ["Get-CimInstance Win32_DiskDrive | ForEach-Object { $_.SerialNumber }"])
            if not uuid:
                raise RuntimeError("UUID компьютера не получен")
        raw = f"{uuid}|{serial}"
        return hashlib.sha256(raw.encode()).hexdigest()[:16]

    @staticmethod
    def _hostname_id() -> str:
        """ID по имени компьютера. Так считали прежние версии, когда wmic
        не было, поэтому ключи по такому ID тоже принимаются."""
        return hashlib.sha256(platform.node().encode()).hexdigest()[:16]

    def get_machine_ids(self) -> list:
        """Все ID этого компьютера, по которым принимается ключ; первый — основной."""
        if LicenseManager._cached_ids is None:
            ids = []
            try:
                ids.append(self._hardware_id())
            except Exception:
                pass
            hostname_id = self._hostname_id()
            if hostname_id not in ids:
                ids.append(hostname_id)
            LicenseManager._cached_ids = ids
        return list(LicenseManager._cached_ids)

    def get_machine_id(self) -> str:
        """Получение уникального ID компьютера."""
        return self.get_machine_ids()[0]

    @staticmethod
    def signed_message(machine_id: str, expiry_date: str) -> bytes:
        """Что подписывается ключом: ID компьютера и дата окончания."""
        return f"svodnik1|{machine_id}|{expiry_date}".encode()

    @classmethod
    def parse_key(cls, key: str):
        """Разобрать ключ SV-ГГГГ-ММ-ДД-ПОДПИСЬ. Вернуть (дата, подпись) или None."""
        parts = "".join(key.split()).upper().split("-")
        if len(parts) != 5 or parts[0] != cls.KEY_PREFIX:
            return None
        expiry = f"{parts[1]}-{parts[2]}-{parts[3]}"
        try:
            datetime.strptime(expiry, "%Y-%m-%d")
            sig = parts[4]
            signature = base64.b32decode(sig + "=" * (-len(sig) % 8))
        except (ValueError, base64.binascii.Error):
            return None
        return expiry, signature

    @staticmethod
    def is_old_format(key: str) -> bool:
        """Ключ прежнего формата KEY-XXXXXXXXXXXX-ГГГГ-ММ-ДД (до версии 1.9.0)."""
        return "".join(key.split()).upper().startswith("KEY-")

    def validate_key(self, key: str, machine_id: str) -> bool:
        """Проверка подписи ключа для указанного ID компьютера."""
        parsed = self.parse_key(key)
        if not parsed:
            return False
        expiry, signature = parsed
        return ed25519.verify(bytes.fromhex(self.PUBLIC_KEY),
                              self.signed_message(machine_id, expiry), signature)

    def save_license(self, key: str, expiry_date: str, machine_id: str):
        """Сохранение лицензии."""
        self.license_data = {
            "key": key,
            "expiry": expiry_date,
            "machine_id": machine_id,
            "activated": datetime.now().isoformat(),
            "exports": 0
        }
        with open(self.LICENSE_FILE, "w") as f:
            json.dump(self.license_data, f, indent=2)

    def load_license(self):
        """Загрузка лицензии."""
        if not os.path.exists(self.LICENSE_FILE):
            self._init_trial()
            return

        try:
            with open(self.LICENSE_FILE, "r") as f:
                self.license_data = json.load(f)

            key = self.license_data.get("key", "")
            parsed = self.parse_key(key)

            if parsed and any(self.validate_key(key, mid) for mid in self.get_machine_ids()):
                expiry = parsed[0]
                self.is_licensed = True
                expiry_dt = datetime.strptime(expiry, "%Y-%m-%d")
                self.days_left = (expiry_dt - datetime.now()).days
                self.exports_count = self.license_data.get("exports", 0)
            else:
                # Ключ прежнего формата перестал действовать: нужен новый
                self.old_key = bool(key) and not parsed
                self._init_trial()
        except Exception:
            self._init_trial()

    def _init_trial(self):
        """Инициализация пробного периода."""
        trial_file = os.path.join(os.path.expanduser("~"), ".svodnik_trial")
        if os.path.exists(trial_file):
            with open(trial_file, "r") as f:
                start_date = datetime.fromisoformat(f.read().strip())
            self.days_left = self.TRIAL_DAYS - (datetime.now() - start_date).days
        else:
            with open(trial_file, "w") as f:
                f.write(datetime.now().isoformat())
            self.days_left = self.TRIAL_DAYS

        # Загружаем счетчик экспорта
        count_file = os.path.join(os.path.expanduser("~"), ".svodnik_count")
        if os.path.exists(count_file):
            with open(count_file, "r") as f:
                self.exports_count = int(f.read().strip())
        else:
            self.exports_count = 0

        self.is_licensed = False

    def activate(self, key: str) -> bool:
        """Активация лицензии."""
        parsed = self.parse_key(key)
        if not parsed:
            return False
        key = "".join(key.split()).upper()
        for machine_id in self.get_machine_ids():
            if self.validate_key(key, machine_id):
                self.save_license(key, parsed[0], machine_id)
                self.load_license()
                return True
        return False

    def can_export(self) -> bool:
        """Проверка возможности экспорта."""
        if self.is_licensed and self.days_left > 0:
            return True
        if not self.is_licensed and self.exports_count < self.MAX_EXPORTS_TRIAL:
            return True
        return False

    def record_export(self):
        """Запись факта экспорта (для пробного режима)."""
        if not self.is_licensed:
            self.exports_count += 1
            count_file = os.path.join(os.path.expanduser("~"), ".svodnik_count")
            with open(count_file, "w") as f:
                f.write(str(self.exports_count))

    def get_status(self) -> str:
        """Получение статуса лицензии."""
        if self.is_licensed:
            if self.days_left > 0:
                return f"Лицензия: активна (осталось {self.days_left} дн.)"
            else:
                return "Лицензия: истекла"
        else:
            remaining = self.MAX_EXPORTS_TRIAL - self.exports_count
            if self.old_key:
                return f"Нужен новый ключ: {remaining}/{self.MAX_EXPORTS_TRIAL} пробных экспортов"
            return f"Пробный режим: {remaining}/{self.MAX_EXPORTS_TRIAL} экспортов (осталось {self.days_left} дн.)"

    def is_valid(self) -> bool:
        """Проверка валидности."""
        if self.is_licensed and self.days_left > 0:
            return True
        if not self.is_licensed and self.exports_count < self.MAX_EXPORTS_TRIAL and self.days_left > 0:
            return True
        return False
