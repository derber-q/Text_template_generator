# Иконка приложения

## Метаданные

- Файлы: `app_icon.png`, `app_icon.ico`
- Назначение: иконка главного окна `tkinter` и дочерних окон приложения.
- Формат: PNG для окна `tkinter`, ICO для готового `.exe`.

## Источник

Иконка загружена из открытого набора Twemoji:

```text
https://raw.githubusercontent.com/twitter/twemoji/master/assets/72x72/1f4c4.png
```

Файл сохранен в проекте, чтобы приложение показывало иконку без доступа к интернету после переноса на другой компьютер.

`app_icon.ico` создан из `app_icon.png` и используется PyInstaller при сборке `start.exe`.
