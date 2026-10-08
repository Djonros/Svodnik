"""
Утилита генерации лицензионных ключей.
Для администратора / распространителя.

Использование:
    python gen_key.py --id <MACHINE_ID>        # Ключ для указанного компьютера
    python gen_key.py                          # Ключ для текущего компьютера
    python gen_key.py --get-id                 # Показать ID текущего компьютера
    python gen_key.py --new-signing-key        # Создать ключ подписи (один раз)
"""

import sys
from datetime import datetime

import signing
from signing import LicenseManager


def main():
    lm = LicenseManager()
    args = sys.argv[1:]

    if args[:1] == ["--get-id"]:
        print(f"ID компьютера: {lm.get_machine_id()}")
        return

    if args[:1] == ["--new-signing-key"]:
        public = signing.create_secret()
        print(f"Ключ подписи сохранен: {signing.SIGNING_KEY_FILE}")
        print("Сделайте его резервную копию и не кладите в репозиторий.")
        print(f"Открытый ключ для app/license.py (PUBLIC_KEY):\n{public}")
        return

    secret = signing.load_secret()
    machine_id = args[1] if args[:1] == ["--id"] and len(args) > 1 else lm.get_machine_id()
    print(f"ID компьютера: {machine_id}")
    print()

    expiry = input("Введите дату окончания (YYYY-MM-DD): ").strip()
    datetime.strptime(expiry, "%Y-%m-%d")

    print()
    print("Лицензионный ключ:")
    print(signing.make_key(secret, machine_id, expiry))
    print()
    print("Активация: Настройки > Ввести лицензию, вставить ключ целиком")


if __name__ == "__main__":
    main()
