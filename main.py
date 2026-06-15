"""Точка входа в desktop-приложение генерации документов."""

# Импортируется sys, чтобы определить операционную систему перед Windows-специфичным скрытием консоли.
import sys

# Импортируется функция запуска GUI, чтобы в этом файле не смешивать интерфейс и бизнес-логику.
from gui import run_app


def hide_console_window() -> None:
    """Скрывает консольное окно при запуске графического приложения в Windows.

    Входные данные не принимает.
    Ничего не возвращает.
    Побочный эффект: если приложение запущено из консольного процесса Windows, скрывает его окно.
    """
    if sys.platform != "win32":
        return

    try:
        # ctypes импортируется локально, потому что нужен только на Windows для обращения к WinAPI.
        import ctypes
    except Exception:
        return

    # Если Python запущен из уже открытого PowerShell или cmd, консоль разделяется с оболочкой пользователя.
    # Такое окно скрывать нельзя: пользователь может потерять видимый терминал после запуска приложения.
    process_ids = (ctypes.c_ulong * 16)()
    attached_process_count = ctypes.windll.kernel32.GetConsoleProcessList(process_ids, len(process_ids))

    if attached_process_count > 1:
        return

    # Дескриптор текущего окна консоли нужен, чтобы скрыть именно его, не затрагивая GUI tkinter.
    console_window = ctypes.windll.kernel32.GetConsoleWindow()

    if console_window:
        # Значение 0 соответствует SW_HIDE в WinAPI и делает консоль невидимой.
        ctypes.windll.user32.ShowWindow(console_window, 0)


def main() -> None:
    """Запускает графическое приложение.

    Входные данные не принимает.
    Ничего не возвращает.
    Побочный эффект: скрывает консоль в Windows, открывает окно tkinter и запускает цикл событий.
    """
    hide_console_window()
    run_app()


if __name__ == "__main__":
    # При прямом запуске файла стартует приложение пользователя.
    main()
