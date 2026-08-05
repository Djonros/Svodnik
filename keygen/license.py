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

    def get_machine_id(self) -> str:
        """Получение уникального ID компьютера."""
        try:
            result = subprocess.run(
                ["wmic", "csproduct", "get", "uuid"],
                capture_output=True, text=True, creationflags=0x08000000
            )
            uuid = result.stdout.strip().split("\n")[-1].strip()

            result = subprocess.run(
                ["wmic", "diskdrive", "get", "serialnumber"],
                capture_output=True, text=True, creationflags=0x08000000
            )
            serial = result.stdout.strip().split("\n")[-1].strip()

            raw = f"{uuid}|{serial}"
            return hashlib.sha256(raw.encode()).hexdigest()[:16]
        except Exception:
            return hashlib.sha256(platform.node().encode()).hexdigest()[:16]

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

            machine_id = self.get_machine_id()
            key = self.license_data.get("key", "")
            expiry = self.license_data.get("expiry", "")

            if key and self.validate_key(key, machine_id, expiry):
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
        machine_id = self.get_machine_id()

        # Формат: KEY-XXXXXXXXXXXX-YYYY-MM-DD (5 частей при split)
        parts = key.split("-")
        if len(parts) != 5 or parts[0] != "KEY":
            return False

        key_hash = parts[1]
        expiry_date = f"{parts[2]}-{parts[3]}-{parts[4]}"

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
