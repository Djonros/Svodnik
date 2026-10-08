"""
Модуль лицензирования для KOMPAS Export Pro.
Поддержка пробного периода и полной лицензии.
"""

import hashlib
import json
import os
import platform
import subprocess
from datetime import datetime, timedelta


class LicenseManager:
    """Менеджер лицензий с привязкой к оборудованию."""

    SECRET_KEY = "Svodnik2024SecretKey!"
    LICENSE_FILE = os.path.join(os.path.expanduser("~"), ".svodnik_license.dat")
    TRIAL_DAYS = 14
    MAX_EXPORTS_TRIAL = 5  # Максимум экспорта в пробном режиме

    def __init__(self):
        self.license_data = {}
        self.is_licensed = False
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

    def generate_key(self, machine_id: str, expiry_date: str) -> str:
        """Генерация лицензионного ключа."""
        raw = f"{machine_id}|{expiry_date}|{self.SECRET_KEY}"
        return hashlib.sha256(raw.encode()).hexdigest()[:12]

    def validate_key(self, key: str, machine_id: str, expiry_date: str) -> bool:
        """Проверка лицензионного ключа."""
        expected = self.generate_key(machine_id, expiry_date)
        return key == expected

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
            expiry = self.license_data.get("expiry", "")

            if key and any(self.validate_key(key, mid, expiry)
                           for mid in self.get_machine_ids()):
                self.is_licensed = True
                expiry_dt = datetime.strptime(expiry, "%Y-%m-%d")
                self.days_left = (expiry_dt - datetime.now()).days
                self.exports_count = self.license_data.get("exports", 0)
            else:
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
        # Формат: KEY-XXXXXXXXXXXX-YYYY-MM-DD (5 частей при split)
        parts = key.split("-")
        if len(parts) != 5 or parts[0] != "KEY":
            return False

        key_hash = parts[1]
        expiry_date = f"{parts[2]}-{parts[3]}-{parts[4]}"

        for machine_id in self.get_machine_ids():
            if self.validate_key(key_hash, machine_id, expiry_date):
                self.save_license(key_hash, expiry_date, machine_id)
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
            return f"Пробный режим: {remaining}/{self.MAX_EXPORTS_TRIAL} экспортов (осталось {self.days_left} дн.)"

    def is_valid(self) -> bool:
        """Проверка валидности."""
        if self.is_licensed and self.days_left > 0:
            return True
        if not self.is_licensed and self.exports_count < self.MAX_EXPORTS_TRIAL and self.days_left > 0:
            return True
        return False

    def get_key_template(self) -> str:
        """Получение шаблона ключа для текущего компьютера."""
        machine_id = self.get_machine_id()
        expiry = (datetime.now() + timedelta(days=365)).strftime("%Y-%m-%d")
        key_hash = self.generate_key(machine_id, expiry)
        return f"KEY-{key_hash[:12]}-{expiry}"
