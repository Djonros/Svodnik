"""
Утилита генерации лицензионных ключей.
Для администратора / распространителя.

Использование:
    python gen_key.py                          # Для текущего компьютера
    python gen_key.py --id <MACHINE_ID>        # Для указанного компьютера
    python gen_key.py --get-id                 # Показать ID текущего компьютера
"""

import sys
from license import LicenseManager


def main():
    lm = LicenseManager()

    # Режим: показать ID текущего компьютера
    if len(sys.argv) > 1 and sys.argv[1] == "--get-id":
        print(f"ID компьютера: {lm.get_machine_id()}")
        return

    # Режим: указать ID вручную
    if len(sys.argv) > 2 and sys.argv[1] == "--id":
        machine_id = sys.argv[2]
    else:
        machine_id = lm.get_machine_id()

    print(f"ID компьютера: {machine_id}")
    print()

    expiry = input("Введите дату окончания (YYYY-MM-DD): ").strip()
    key_hash = lm.generate_key(machine_id, expiry)

    print()
    print(f"Лицензионный ключ:")
    print(f"KEY-{key_hash}-{expiry}")
    print()
    print("Активация: ввести ключ в программе")


if __name__ == "__main__":
    main()
