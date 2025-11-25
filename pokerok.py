# -*- coding: utf-8 -*-
"""
pokerok.py
Скрипт для чтения свойств из MSI-файла pokerok.msi
"""

import os
import sys
import argparse
import msilib


# ---------- Разбор аргументов командной строки ----------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Чтение таблицы Property из MSI-файла."
    )
    parser.add_argument(
        "msi_path",
        nargs="?",
        default="pokerok.msi",
        help="Путь к MSI-файлу (по умолчанию: pokerok.msi в текущем каталоге)."
    )
    return parser.parse_args()


# ---------- Логика работы с MSI ----------

def read_msi_properties(path: str) -> dict:
    """
    Открывает MSI-файл и читает таблицу Property.
    Возвращает словарь {имя_свойства: значение}.
    Бросает msilib.MSIError при проблемах с MSI.
    """
    db = msilib.OpenDatabase(path, msilib.MSIDBOPEN_READONLY)

    try:
        view = db.OpenView("SELECT * FROM Property")
        view.Execute(None)
    except msilib.MSIError as e:
        # Таблица Property отсутствует или файл поврежден
        raise msilib.MSIError(f"Не удалось открыть таблицу Property: {e}")

    properties = {}

    while True:
        record = view.Fetch()
        if record is None:
            break
        # 1-й столбец — имя свойства, 2-й — значение
        name = record.GetString(1)
        value = record.GetString(2)
        properties[name] = value

    return properties


# ---------- Функции вывода ----------

def print_key_properties(properties: dict) -> None:
    """
    Выводит основные свойства продукта, если они есть.
    """
    print("=== Основная информация о продукте ===")

    # Список ключевых свойств и человеко-читаемых меток
    keys = [
        ("ProductName", "ProductName"),
        ("ProductCode", "ProductCode"),
        ("ProductVersion", "ProductVersion"),
        ("Manufacturer", "Manufacturer"),
    ]

    for key, label in keys:
        value = properties.get(key)
        if value is None:
            value = "не указано"
        print(f"{label:14}: {value}")
    print()  # пустая строка для отделения блоков


def print_all_properties(properties: dict) -> None:
    """
    Выводит все свойства таблицы Property.
    """
    print("=== Все свойства MSI (таблица Property) ===")

    for name in sorted(properties.keys(), key=str.lower):
        value = properties[name]
        print(f"{name} = {value}")


# ---------- Основная функция ----------

def main() -> int:
    args = parse_args()
    path = args.msi_path

    # 1. Проверка существования файла
    if not os.path.exists(path) or not os.path.isfile(path):
        sys.stderr.write(f'Файл "{path}" не найден или не является файлом.\n')
        return 1

    try:
        # 2. Чтение свойств из MSI
        properties = read_msi_properties(path)

        if not properties:
            sys.stderr.write(
                "Таблица Property не найдена или не содержит записей.\n"
            )
            return 2

        # 3. Вывод основной информации
        print_key_properties(properties)

        # 4. Вывод всех свойств
        print_all_properties(properties)

        return 0

    except msilib.MSIError as e:
        sys.stderr.write(
            f"Ошибка при чтении файла MSI: таблица Property недоступна "
            f"или файл поврежден.\nПодробнее: {e}\n"
        )
        return 2
    except Exception as e:
        # Непредвиденная ошибка
        sys.stderr.write(f"Непредвиденная ошибка: {e}\n")
        return 3


if __name__ == "__main__":
    sys.exit(main())