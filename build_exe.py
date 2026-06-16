"""Сборка Windows `.exe` для локального запуска приложения без установленного Python."""

# Импортируется shutil, чтобы очищать старые каталоги сборки PyInstaller.
import shutil

# Импортируется subprocess, чтобы запускать PyInstaller как стандартный модуль Python.
import subprocess

# Импортируется sys, чтобы использовать текущий интерпретатор Python для запуска PyInstaller.
import sys

# Импортируется Path, чтобы надежно строить пути к файлам сборки.
from pathlib import Path


# Корень проекта нужен как базовая папка для всех относительных путей.
PROJECT_ROOT = Path(__file__).resolve().parent

# Каталог build содержит промежуточные файлы PyInstaller и может безопасно пересоздаваться.
BUILD_DIR = PROJECT_ROOT / "build"

# Имя готового `.exe` выбрано коротким и понятным пользователю.
APP_NAME = "TextTemplateGenerator"

# Готовый `.exe` лежит в корне проекта, чтобы после клонирования его было легко найти и запустить.
ROOT_EXE_PATH = PROJECT_ROOT / f"{APP_NAME}.exe"


def main() -> None:
    """Собирает приложение в один `.exe` в корне проекта.

    Входные данные не принимает.
    Ничего не возвращает.
    Побочный эффект: удаляет старые build/dist-артефакты и создает новую PyInstaller-сборку.
    """
    _remove_old_build_artifacts()
    _run_pyinstaller()
    print(f"Готовый EXE создан: {ROOT_EXE_PATH}")


def _remove_old_build_artifacts() -> None:
    """Удаляет старые артефакты сборки перед новым запуском PyInstaller.

    Входные данные не принимает.
    Ничего не возвращает.
    Побочный эффект: удаляет локальный каталог `build/` и старый корневой `.exe`, если они существуют.
    """
    if BUILD_DIR.exists():
        shutil.rmtree(BUILD_DIR)

    if ROOT_EXE_PATH.exists():
        ROOT_EXE_PATH.unlink()


def _run_pyinstaller() -> None:
    """Запускает PyInstaller с параметрами GUI-приложения.

    Входные данные не принимает.
    Ничего не возвращает.
    Побочный эффект: создает onefile-сборку приложения в корне проекта.
    """
    # Формат `источник;назначение` нужен PyInstaller в Windows для добавления data-файлов.
    assets_data = f"{PROJECT_ROOT / 'assets'};assets"

    # Команда собирает оконное приложение без консоли в один файл.
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--onefile",
        "--name",
        APP_NAME,
        "--distpath",
        str(PROJECT_ROOT),
        "--workpath",
        str(BUILD_DIR),
        "--specpath",
        str(BUILD_DIR),
        "--icon",
        str(PROJECT_ROOT / "assets" / "app_icon.ico"),
        "--add-data",
        assets_data,
        "--collect-data",
        "docxcompose",
        str(PROJECT_ROOT / "main.py"),
    ]

    subprocess.run(command, cwd=PROJECT_ROOT, check=True)


if __name__ == "__main__":
    # При прямом запуске файла выполняется полная сборка приложения.
    main()
