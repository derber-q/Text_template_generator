"""Графический интерфейс tkinter для генератора DOCX-документов."""

# Импортируется queue, чтобы безопасно передавать события из рабочего потока в главный поток tkinter.
import queue

# Импортируется threading, чтобы длительная генерация DOCX не блокировала графический интерфейс.
import threading

# Импортируется sys, чтобы находить ресурсы внутри собранного `.exe`.
import sys

# Импортируется tkinter, чтобы создать локальное desktop-окно без веб-сервера.
import tkinter as tk

# Импортируются стандартные диалоги tkinter для выбора файлов, сохранения и показа сообщений.
from tkinter import filedialog, font as tkfont, messagebox, ttk

# Импортируется dataclass, чтобы передавать подготовленные параметры генерации одним объектом.
from dataclasses import dataclass

# Импортируется Path, чтобы передавать пути в бизнес-логику в явном виде.
from pathlib import Path

# Импортируется логика DOCX, чтобы GUI только собирал данные и запускал обработку.
from docx_processor import DocxProcessingError, create_numbered_document, document_contains_placeholder

# Импортируются функции генерации номеров и загрузки запрещенных комбинаций.
from number_generator import (
    NumberGenerationError,
    build_full_numbers,
    generate_unique_digit_parts,
    load_forbidden_combinations,
    save_generated_combinations,
)

# Импортируются валидаторы, чтобы показывать пользователю понятные ошибки до генерации.
from validators import (
    ValidationError,
    parse_positive_int,
    validate_docx_path,
    validate_optional_txt_path,
    validate_output_docx_path,
    validate_static_part,
)


# Ширина подсказок задается один раз, чтобы длинные русские инструкции не растягивали окно.
HELP_WRAP_LENGTH = 720

def _resource_path(relative_path: str) -> Path:
    """Возвращает путь к ресурсу приложения для обычного запуска и для `.exe`.

    relative_path: относительный путь ресурса внутри проекта или PyInstaller-пакета.
    Возвращает абсолютный Path к ресурсу.
    Побочных эффектов нет.
    """
    # PyInstaller распаковывает bundled-ресурсы во временный каталог `_MEIPASS`.
    bundled_root = getattr(sys, "_MEIPASS", None)

    if bundled_root is not None:
        return Path(bundled_root) / relative_path

    # При обычном Python-запуске ресурсы лежат рядом с исходными файлами проекта.
    return Path(__file__).resolve().parent / relative_path


# Путь к иконке строится через функцию ресурсов, чтобы работать и из исходников, и из PyInstaller `.exe`.
APP_ICON_PATH = _resource_path("assets/app_icon.png")


@dataclass(frozen=True)
class GenerationRequest:
    """Подготовленные и проверенные параметры генерации документа.

    Класс хранит только данные, которые уже прошли валидацию и могут безопасно использоваться
    рабочим потоком без обращения к виджетам tkinter.
    """

    # Путь к шаблону используется рабочим потоком для чтения исходного DOCX.
    template_path: Path

    # Путь сохранения определяет, куда будет записан итоговый DOCX.
    output_path: Path

    # Полные номера содержат символ «№», постоянную часть и случайные цифры для каждой копии.
    full_numbers: list[str]

    # Случайные цифровые части нужны для необязательного TXT-файла исключений.
    digit_parts: list[str]

    # Имя шрифта применяется только к вставленным номерам.
    font_name: str

    # Размер шрифта применяется только к вставленным номерам.
    font_size: int

    # Флаг показывает, нужно ли создавать TXT-файл со сгенерированными цифровыми частями.
    export_generated_combinations: bool


class TemplateGeneratorApp:
    """Главное окно приложения и обработчики действий пользователя.

    Класс хранит состояние полей формы, строит два экрана мастера и связывает GUI с модулями
    генерации номеров и обработки DOCX.
    """

    def __init__(self, root: tk.Tk) -> None:
        """Инициализирует окно приложения.

        root: созданный объект Tk.
        Ничего не возвращает.
        Побочный эффект: настраивает окно, создает переменные формы и показывает первый экран.
        """
        # Корневое окно используется всеми виджетами приложения.
        self.root = root
        self.root.title("Генератор DOCX по шаблону")
        self.root.minsize(820, 640)

        # Путь к шаблону хранится отдельно, чтобы поле можно было обновлять после file dialog.
        self.template_path_var = tk.StringVar()

        # Количество копий по умолчанию равно 1, потому что это минимальный рабочий сценарий.
        self.copy_count_var = tk.StringVar(value="1")

        # Список системных шрифтов берется из tkinter после создания root.
        self.available_fonts = self._load_system_fonts()

        # Выбранный шрифт по умолчанию выбирается из распространенных вариантов или первого доступного.
        self.font_name_var = tk.StringVar(value=self._choose_default_font())

        # Размер шрифта по умолчанию выбран как обычный размер текста в документах Word.
        self.font_size_var = tk.StringVar(value="12")

        # Постоянная часть номера вводится без символа «№», который программа добавит сама.
        self.static_part_var = tk.StringVar()

        # Количество случайных цифр по умолчанию соответствует примеру из технического задания.
        self.digit_count_var = tk.StringVar(value="6")

        # Путь к необязательному списку запрещенных комбинаций может оставаться пустым.
        self.forbidden_path_var = tk.StringVar()

        # Путь сохранения итогового документа выбирается перед запуском генерации.
        self.output_path_var = tk.StringVar()

        # Флаг управляет созданием дополнительного TXT-файла с использованными случайными частями.
        self.export_generated_combinations_var = tk.BooleanVar(value=False)

        # Окно прогресса создается только на время генерации документа.
        self.progress_window: tk.Toplevel | None = None

        # Переменная прогресс-бара хранит количество завершенных шагов генерации.
        self.progress_value_var = tk.DoubleVar(value=0.0)

        # Текст прогресса поясняет пользователю текущую стадию создания файлов.
        self.progress_text_var = tk.StringVar(value="")

        # Очередь событий используется рабочим потоком для безопасной передачи статуса в GUI.
        self.generation_queue: queue.Queue | None = None

        # Номер текущего шага нужен для кнопок перехода между экранами.
        self.current_step = 1

        # Основная рамка заменяется содержимым текущего шага.
        self.content_frame = ttk.Frame(self.root, padding=16)
        self.content_frame.pack(fill=tk.BOTH, expand=True)

        self._show_step_one()

    def _load_system_fonts(self) -> list[str]:
        """Загружает доступные системные шрифты.

        Входные данные не принимает.
        Возвращает отсортированный список имен шрифтов.
        Побочных эффектов нет, кроме обращения к tkinter.font.
        """
        # set убирает дубликаты, которые иногда возвращают системные библиотеки шрифтов.
        font_names = sorted(set(tkfont.families()))

        return font_names or ["Arial"]

    def _choose_default_font(self) -> str:
        """Выбирает шрифт по умолчанию для вставленного номера.

        Входные данные не принимает.
        Возвращает имя шрифта.
        Побочных эффектов нет.
        """
        for preferred_font in ("Arial", "Times New Roman", "Calibri"):
            if preferred_font in self.available_fonts:
                return preferred_font

        return self.available_fonts[0]

    def _clear_content(self) -> None:
        """Удаляет виджеты текущего экрана.

        Входные данные не принимает.
        Ничего не возвращает.
        Побочный эффект: очищает content_frame перед построением следующего шага.
        """
        for child in self.content_frame.winfo_children():
            child.destroy()

    def _show_step_one(self) -> None:
        """Показывает первый экран мастера.

        Входные данные не принимает.
        Ничего не возвращает.
        Побочный эффект: создает виджеты выбора шаблона, количества копий, шрифта и размера.
        """
        self.current_step = 1
        self._clear_content()

        # Заголовок объясняет назначение программы прямо в интерфейсе.
        title_label = ttk.Label(
            self.content_frame,
            text="Шаг 1. Выберите DOCX-шаблон и параметры вставляемого номера",
            font=("Segoe UI", 14, "bold"),
        )
        title_label.pack(anchor=tk.W, pady=(0, 8))

        # Общая подсказка помогает использовать приложение без чтения README.
        intro_label = ttk.Label(
            self.content_frame,
            text=(
                "Программа создает один итоговый .docx-файл из выбранного шаблона. "
                "В шаблоне заранее должен быть текст [№] с квадратными скобками. "
                "В каждой копии этот плейсхолдер будет заменен на уникальный номер."
            ),
            wraplength=HELP_WRAP_LENGTH,
            justify=tk.LEFT,
        )
        intro_label.pack(anchor=tk.W, pady=(0, 12))

        # Группа выбора шаблона отделяет файловый ввод от числовых параметров.
        template_frame = ttk.LabelFrame(self.content_frame, text="DOCX-шаблон", padding=12)
        template_frame.pack(fill=tk.X, pady=(0, 12))

        template_help = ttk.Label(
            template_frame,
            text=(
                "Выберите .docx-файл, который будет использован как шаблон. "
                "Шаблон — это обычный документ Word, где в нужном месте указан плейсхолдер [№]."
            ),
            wraplength=HELP_WRAP_LENGTH,
            justify=tk.LEFT,
        )
        template_help.pack(anchor=tk.W, pady=(0, 8))

        template_row = ttk.Frame(template_frame)
        template_row.pack(fill=tk.X)

        template_entry = ttk.Entry(template_row, textvariable=self.template_path_var)
        template_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        template_button = ttk.Button(template_row, text="Выбрать .docx", command=self._choose_template)
        template_button.pack(side=tk.LEFT)

        # Группа параметров задает количество копий и форматирование вставленного номера.
        settings_frame = ttk.LabelFrame(self.content_frame, text="Параметры вставки", padding=12)
        settings_frame.pack(fill=tk.X, pady=(0, 12))

        settings_help = ttk.Label(
            settings_frame,
            text=(
                "Укажите, сколько копий шаблона нужно создать в итоговом документе. "
                "Шрифт и размер применяются только к номеру, который будет вставлен вместо [№]."
            ),
            wraplength=HELP_WRAP_LENGTH,
            justify=tk.LEFT,
        )
        settings_help.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 8))

        copy_label = ttk.Label(settings_frame, text="Количество копий:")
        copy_label.grid(row=1, column=0, sticky=tk.W, pady=4)

        copy_entry = ttk.Entry(settings_frame, textvariable=self.copy_count_var, width=12)
        copy_entry.grid(row=1, column=1, sticky=tk.W, pady=4)

        font_label = ttk.Label(settings_frame, text="Шрифт вставленного номера:")
        font_label.grid(row=2, column=0, sticky=tk.W, pady=4)

        font_combo = ttk.Combobox(
            settings_frame,
            textvariable=self.font_name_var,
            values=self.available_fonts,
            state="readonly",
            width=36,
        )
        font_combo.grid(row=2, column=1, sticky=tk.W, pady=4)

        size_label = ttk.Label(settings_frame, text="Размер шрифта:")
        size_label.grid(row=3, column=0, sticky=tk.W, pady=4)

        size_spinbox = tk.Spinbox(
            settings_frame,
            from_=1,
            to=200,
            textvariable=self.font_size_var,
            width=10,
        )
        size_spinbox.grid(row=3, column=1, sticky=tk.W, pady=4)

        settings_frame.columnconfigure(1, weight=1)

        navigation_frame = ttk.Frame(self.content_frame)
        navigation_frame.pack(fill=tk.X, pady=(8, 0))

        next_button = ttk.Button(navigation_frame, text="Далее", command=self._go_to_step_two)
        next_button.pack(side=tk.RIGHT)

    def _show_step_two(self) -> None:
        """Показывает второй экран мастера.

        Входные данные не принимает.
        Ничего не возвращает.
        Побочный эффект: создает виджеты ввода номера, запрещенных комбинаций и пути сохранения.
        """
        self.current_step = 2
        self._clear_content()

        title_label = ttk.Label(
            self.content_frame,
            text="Шаг 2. Настройте номер и сохранение результата",
            font=("Segoe UI", 14, "bold"),
        )
        title_label.pack(anchor=tk.W, pady=(0, 8))

        intro_label = ttk.Label(
            self.content_frame,
            text=(
                "Введите постоянную часть номера без символа «№». "
                "Программа автоматически добавит «№» в начало и затем допишет случайные цифры. "
                "Например: постоянная часть ABC-, 6 случайных цифр, результат №ABC-483920."
            ),
            wraplength=HELP_WRAP_LENGTH,
            justify=tk.LEFT,
        )
        intro_label.pack(anchor=tk.W, pady=(0, 12))

        number_frame = ttk.LabelFrame(self.content_frame, text="Формат номера", padding=12)
        number_frame.pack(fill=tk.X, pady=(0, 12))

        static_label = ttk.Label(number_frame, text="Постоянная часть без «№»:")
        static_label.grid(row=0, column=0, sticky=tk.W, pady=4)

        static_entry = ttk.Entry(number_frame, textvariable=self.static_part_var, width=36)
        static_entry.grid(row=0, column=1, sticky=tk.W, pady=4)

        digit_label = ttk.Label(number_frame, text="Количество случайных цифр:")
        digit_label.grid(row=1, column=0, sticky=tk.W, pady=4)

        digit_entry = ttk.Entry(number_frame, textvariable=self.digit_count_var, width=12)
        digit_entry.grid(row=1, column=1, sticky=tk.W, pady=4)

        number_help = ttk.Label(
            number_frame,
            text=(
                "Символ «№» вводить не нужно. Если указать 6 случайных цифр, "
                "то каждая случайная часть будет иметь ровно 6 цифр, включая возможные ведущие нули."
            ),
            wraplength=HELP_WRAP_LENGTH,
            justify=tk.LEFT,
        )
        number_help.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(8, 0))

        number_frame.columnconfigure(1, weight=1)

        forbidden_frame = ttk.LabelFrame(self.content_frame, text="Необязательный .txt-файл запретов", padding=12)
        forbidden_frame.pack(fill=tk.X, pady=(0, 12))

        forbidden_help = ttk.Label(
            forbidden_frame,
            text=(
                "Можно выбрать обычный .txt-файл со случайными цифровыми комбинациями, "
                "которые нельзя использовать. Формат: 123456,984301,000381 без пробелов. "
                "Если файл не нужен, оставьте поле пустым. "
                "Постоянные запреты из permanent_forbidden_combinations.txt учитываются всегда."
            ),
            wraplength=HELP_WRAP_LENGTH,
            justify=tk.LEFT,
        )
        forbidden_help.pack(anchor=tk.W, pady=(0, 8))

        forbidden_row = ttk.Frame(forbidden_frame)
        forbidden_row.pack(fill=tk.X)

        forbidden_entry = ttk.Entry(forbidden_row, textvariable=self.forbidden_path_var)
        forbidden_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        forbidden_button = ttk.Button(forbidden_row, text="Выбрать .txt", command=self._choose_forbidden_file)
        forbidden_button.pack(side=tk.LEFT, padx=(0, 8))

        clear_forbidden_button = ttk.Button(forbidden_row, text="Очистить", command=self._clear_forbidden_file)
        clear_forbidden_button.pack(side=tk.LEFT)

        output_frame = ttk.LabelFrame(self.content_frame, text="Итоговый файл", padding=12)
        output_frame.pack(fill=tk.X, pady=(0, 12))

        output_help = ttk.Label(
            output_frame,
            text=(
                "Перед генерацией выберите, куда сохранить итоговый .docx-документ. "
                "Внутри будет указанное количество копий шаблона, каждая новая копия начнется с новой страницы."
            ),
            wraplength=HELP_WRAP_LENGTH,
            justify=tk.LEFT,
        )
        output_help.pack(anchor=tk.W, pady=(0, 8))

        output_row = ttk.Frame(output_frame)
        output_row.pack(fill=tk.X)

        output_entry = ttk.Entry(output_row, textvariable=self.output_path_var)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        output_button = ttk.Button(output_row, text="Куда сохранить .docx", command=self._choose_output_file)
        output_button.pack(side=tk.LEFT)

        export_check = ttk.Checkbutton(
            output_frame,
            text="Создать .txt со сгенерированными цифровыми комбинациями",
            variable=self.export_generated_combinations_var,
        )
        export_check.pack(anchor=tk.W, pady=(8, 0))

        export_help = ttk.Label(
            output_frame,
            text=(
                "Если флаг включен, рядом с итоговым .docx будет создан .txt-файл со случайными "
                "цифровыми частями через запятую без пробелов. Такой файл можно позже выбрать как список исключений."
            ),
            wraplength=HELP_WRAP_LENGTH,
            justify=tk.LEFT,
        )
        export_help.pack(anchor=tk.W, pady=(4, 0))

        navigation_frame = ttk.Frame(self.content_frame)
        navigation_frame.pack(fill=tk.X, pady=(8, 0))

        back_button = ttk.Button(navigation_frame, text="Назад", command=self._show_step_one)
        back_button.pack(side=tk.LEFT)

        generate_button = ttk.Button(navigation_frame, text="Создать DOCX", command=self._generate_document)
        generate_button.pack(side=tk.RIGHT)

    def _choose_template(self) -> None:
        """Открывает диалог выбора `.docx`-шаблона.

        Входные данные не принимает.
        Ничего не возвращает.
        Побочный эффект: записывает выбранный путь в template_path_var.
        """
        # filetypes ограничивает выбор пользователя документами Word, но не заменяет полноценную валидацию.
        selected_path = filedialog.askopenfilename(
            title="Выберите DOCX-шаблон",
            filetypes=(("Документы Word", "*.docx"), ("Все файлы", "*.*")),
        )

        if selected_path:
            self.template_path_var.set(selected_path)

    def _choose_forbidden_file(self) -> None:
        """Открывает диалог выбора `.txt`-файла запрещенных комбинаций.

        Входные данные не принимает.
        Ничего не возвращает.
        Побочный эффект: записывает выбранный путь в forbidden_path_var.
        """
        selected_path = filedialog.askopenfilename(
            title="Выберите TXT-файл запрещенных комбинаций",
            filetypes=(("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")),
        )

        if selected_path:
            self.forbidden_path_var.set(selected_path)

    def _clear_forbidden_file(self) -> None:
        """Очищает путь к необязательному `.txt`-файлу.

        Входные данные не принимает.
        Ничего не возвращает.
        Побочный эффект: сбрасывает forbidden_path_var.
        """
        self.forbidden_path_var.set("")

    def _choose_output_file(self) -> None:
        """Открывает диалог выбора пути сохранения итогового `.docx`.

        Входные данные не принимает.
        Ничего не возвращает.
        Побочный эффект: записывает выбранный путь в output_path_var.
        """
        selected_path = filedialog.asksaveasfilename(
            title="Сохранить итоговый DOCX",
            defaultextension=".docx",
            filetypes=(("Документы Word", "*.docx"), ("Все файлы", "*.*")),
            initialfile="generated_document.docx",
        )

        if selected_path:
            self.output_path_var.set(selected_path)

    def _go_to_step_two(self) -> None:
        """Проверяет первый экран и переходит ко второму.

        Входные данные не принимает.
        Ничего не возвращает.
        Побочный эффект: показывает ошибку или переключает экран мастера.
        """
        try:
            self._collect_step_one_data()
        except (ValidationError, DocxProcessingError) as exc:
            messagebox.showerror("Проверьте данные", str(exc))
            return

        self._show_step_two()

    def _collect_step_one_data(self) -> tuple[Path, int, str, int]:
        """Собирает и проверяет данные первого экрана.

        Входные данные берутся из переменных формы.
        Возвращает путь к шаблону, количество копий, имя шрифта и размер шрифта.
        Побочный эффект: читает `.docx` для проверки наличия плейсхолдера.
        """
        # Путь к шаблону проверяется до чтения документа.
        template_path = validate_docx_path(self.template_path_var.get(), "DOCX-шаблон")

        if not document_contains_placeholder(template_path):
            raise ValidationError("В выбранном шаблоне не найден плейсхолдер [№].")

        # Количество копий определяет, сколько уникальных номеров понадобится сгенерировать.
        copy_count = parse_positive_int(self.copy_count_var.get(), "Количество копий")

        # Шрифт берется из combobox; пустое значение означает ошибку состояния формы.
        font_name = self.font_name_var.get().strip()

        if not font_name:
            raise ValidationError("Выберите шрифт для вставленного номера.")

        # Размер шрифта применяется только к вставленному номеру в DOCX.
        font_size = parse_positive_int(self.font_size_var.get(), "Размер шрифта")

        return template_path, copy_count, font_name, font_size

    def _generate_document(self) -> None:
        """Запускает полную генерацию итогового `.docx`.

        Входные данные берутся из переменных формы.
        Ничего не возвращает.
        Побочный эффект: проверяет поля, открывает окно прогресса и запускает рабочий поток генерации.
        """
        if not self.output_path_var.get().strip():
            self._choose_output_file()

            if not self.output_path_var.get().strip():
                return

        try:
            generation_request = self._prepare_generation_request()
        except (ValidationError, NumberGenerationError, DocxProcessingError) as exc:
            messagebox.showerror("Ошибка генерации", str(exc))
            return

        # Отдельный шаг сохранения DOCX учитывается внутри docx_processor; TXT добавляет еще один шаг.
        total_steps = len(generation_request.full_numbers) + 1

        if generation_request.export_generated_combinations:
            total_steps += 1

        self._open_progress_window(total_steps)
        self._start_generation_worker(generation_request, total_steps)

    def _prepare_generation_request(self) -> GenerationRequest:
        """Собирает, валидирует и подготавливает данные для генерации.

        Входные данные берутся из переменных формы.
        Возвращает GenerationRequest с готовыми путями, номерами и параметрами форматирования.
        Побочный эффект: читает шаблон и TXT-файлы запретов для проверки и подготовки номеров.
        """
        # Данные первого экрана проверяются повторно, потому что пользователь мог вернуться и изменить поля.
        template_path, copy_count, font_name, font_size = self._collect_step_one_data()

        # Постоянная часть номера очищается и проверяется отдельно от случайных цифр.
        static_part = validate_static_part(self.static_part_var.get())

        # Количество случайных цифр определяет длину всех сгенерированных комбинаций.
        digit_count = parse_positive_int(self.digit_count_var.get(), "Количество случайных цифр")

        # TXT-файл необязателен, но если выбран, путь и формат должны быть корректными.
        forbidden_path = validate_optional_txt_path(self.forbidden_path_var.get())

        # Итоговый путь выбирается пользователем через save dialog и должен оставаться DOCX-файлом.
        output_path = validate_output_docx_path(self.output_path_var.get())

        # Запрещенные комбинации читаются до генерации, чтобы проверить доступную емкость.
        forbidden_combinations = load_forbidden_combinations(forbidden_path, digit_count)

        # Случайные части генерируются уникальными и без совпадений с запретами.
        digit_parts = generate_unique_digit_parts(copy_count, digit_count, forbidden_combinations)

        # Полные номера получают символ «№» и постоянную часть.
        full_numbers = build_full_numbers(static_part, digit_parts)

        return GenerationRequest(
            template_path=template_path,
            output_path=output_path,
            full_numbers=full_numbers,
            digit_parts=digit_parts,
            font_name=font_name,
            font_size=font_size,
            export_generated_combinations=self.export_generated_combinations_var.get(),
        )

    def _open_progress_window(self, total_steps: int) -> None:
        """Открывает модальное окно прогресса генерации.

        total_steps: общее количество шагов, которое будет отображать progressbar.
        Ничего не возвращает.
        Побочный эффект: создает Toplevel-окно и блокирует работу с основным окном до завершения.
        """
        self._close_progress_window()

        # Новое окно прогресса привязано к главному окну, чтобы не потеряться за ним.
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title("Создание DOCX")
        self.progress_window.resizable(False, False)
        self.progress_window.transient(self.root)
        self.progress_window.protocol("WM_DELETE_WINDOW", lambda: None)

        # Начальное значение показывает, что генерация уже запущена, но первые копии еще не готовы.
        self.progress_value_var.set(0)
        self.progress_text_var.set("Подготовка создания документа...")

        frame = ttk.Frame(self.progress_window, padding=18)
        frame.pack(fill=tk.BOTH, expand=True)

        title_label = ttk.Label(frame, text="Создание документа", font=("Segoe UI", 12, "bold"))
        title_label.pack(anchor=tk.W, pady=(0, 8))

        status_label = ttk.Label(frame, textvariable=self.progress_text_var, wraplength=420, justify=tk.LEFT)
        status_label.pack(anchor=tk.W, fill=tk.X, pady=(0, 10))

        progress_bar = ttk.Progressbar(
            frame,
            orient=tk.HORIZONTAL,
            mode="determinate",
            maximum=total_steps,
            variable=self.progress_value_var,
            length=420,
        )
        progress_bar.pack(fill=tk.X)

        hint_label = ttk.Label(frame, text="Дождитесь завершения. Окно закроется автоматически.")
        hint_label.pack(anchor=tk.W, pady=(10, 0))

        # grab_set делает окно прогресса модальным и не дает запустить вторую генерацию параллельно.
        self.progress_window.grab_set()
        self.progress_window.focus_set()
        self._center_window(self.progress_window)

    def _center_window(self, window: tk.Toplevel) -> None:
        """Размещает дочернее окно по центру главного окна приложения.

        window: окно Toplevel, которое нужно центрировать.
        Ничего не возвращает.
        Побочный эффект: меняет geometry дочернего окна.
        """
        self.root.update_idletasks()
        window.update_idletasks()

        # Размеры главного и дочернего окон нужны для расчета аккуратной позиции по центру.
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        window_width = window.winfo_width()
        window_height = window.winfo_height()

        position_x = root_x + max((root_width - window_width) // 2, 0)
        position_y = root_y + max((root_height - window_height) // 2, 0)

        window.geometry(f"+{position_x}+{position_y}")

    def _close_progress_window(self) -> None:
        """Закрывает окно прогресса, если оно сейчас открыто.

        Входные данные не принимает.
        Ничего не возвращает.
        Побочный эффект: снимает modal grab и уничтожает Toplevel-окно прогресса.
        """
        if self.progress_window is None:
            return

        try:
            if self.progress_window.winfo_exists():
                self.progress_window.grab_release()
                self.progress_window.destroy()
        except tk.TclError:
            pass
        finally:
            self.progress_window = None

    def _set_progress(self, completed_steps: int, total_steps: int, message: str) -> None:
        """Обновляет текст и значение progressbar в главном потоке tkinter.

        completed_steps: количество завершенных шагов.
        total_steps: общее количество шагов.
        message: текст текущего состояния для пользователя.
        Ничего не возвращает.
        Побочный эффект: меняет переменные, связанные с виджетами окна прогресса.
        """
        # Ограничение значения защищает progressbar от некорректного состояния при неожиданных событиях.
        safe_completed_steps = max(0, min(completed_steps, total_steps))
        self.progress_value_var.set(safe_completed_steps)
        self.progress_text_var.set(message)

        if self.progress_window is not None:
            self.progress_window.update_idletasks()

    def _start_generation_worker(self, generation_request: GenerationRequest, total_steps: int) -> None:
        """Запускает рабочий поток генерации и включает опрос очереди событий.

        generation_request: подготовленные параметры генерации.
        total_steps: общее количество шагов для progressbar.
        Ничего не возвращает.
        Побочный эффект: создает очередь, запускает daemon-поток и планирует опрос очереди через tkinter.
        """
        # Очередь принадлежит конкретному запуску генерации и очищается после успеха или ошибки.
        self.generation_queue = queue.Queue()

        worker_thread = threading.Thread(
            target=self._generation_worker,
            args=(generation_request, total_steps, self.generation_queue),
            daemon=True,
        )
        worker_thread.start()

        self._poll_generation_queue()

    def _generation_worker(
        self,
        generation_request: GenerationRequest,
        total_steps: int,
        event_queue: queue.Queue,
    ) -> None:
        """Выполняет генерацию DOCX в фоновом потоке.

        generation_request: подготовленные параметры генерации.
        total_steps: общее число шагов GUI-прогресса.
        event_queue: потокобезопасная очередь событий для главного потока.
        Ничего не возвращает.
        Побочный эффект: создает DOCX и, при включенном флаге, TXT-файл с комбинациями.
        """
        try:
            # Путь к TXT остается None, если пользователь не включил экспорт комбинаций.
            generated_combinations_path: Path | None = None

            def report_docx_progress(completed_docx_steps: int, docx_total_steps: int) -> None:
                """Передает в очередь прогресс создания DOCX.

                completed_docx_steps: завершенные шаги внутри docx_processor.
                docx_total_steps: общее число шагов внутри docx_processor.
                Ничего не возвращает.
                Побочный эффект: добавляет событие progress в очередь GUI.
                """
                # Последний шаг docx_processor означает сохранение итогового DOCX-файла.
                if completed_docx_steps >= docx_total_steps:
                    progress_text = "Сохранение итогового DOCX-файла..."
                else:
                    progress_text = (
                        f"Создание копии {completed_docx_steps} "
                        f"из {len(generation_request.full_numbers)}..."
                    )

                event_queue.put(("progress", completed_docx_steps, total_steps, progress_text))

            create_numbered_document(
                generation_request.template_path,
                generation_request.output_path,
                generation_request.full_numbers,
                generation_request.font_name,
                generation_request.font_size,
                progress_callback=report_docx_progress,
            )

            if generation_request.export_generated_combinations:
                event_queue.put(("progress", total_steps - 1, total_steps, "Сохранение TXT-файла с комбинациями..."))
                generated_combinations_path = save_generated_combinations(
                    generation_request.digit_parts,
                    generation_request.output_path,
                )
                event_queue.put(("progress", total_steps, total_steps, "TXT-файл с комбинациями сохранен."))

            event_queue.put(("success", generation_request.output_path, generated_combinations_path))
        except (ValidationError, NumberGenerationError, DocxProcessingError) as exc:
            event_queue.put(("error", str(exc)))
        except Exception as exc:
            event_queue.put(("error", f"Непредвиденная ошибка генерации: {exc}"))

    def _poll_generation_queue(self) -> None:
        """Обрабатывает события фоновой генерации в главном потоке tkinter.

        Входные данные не принимает.
        Ничего не возвращает.
        Побочный эффект: обновляет progressbar, показывает сообщение успеха или ошибку.
        """
        if self.generation_queue is None:
            return

        try:
            while True:
                # События забираются пачкой, чтобы progressbar быстро догонял фактический статус.
                event = self.generation_queue.get_nowait()
                event_type = event[0]

                if event_type == "progress":
                    _, completed_steps, total_steps, message = event
                    self._set_progress(completed_steps, total_steps, message)
                    continue

                if event_type == "success":
                    _, output_path, generated_combinations_path = event
                    self.generation_queue = None
                    self._finish_generation_success(output_path, generated_combinations_path)
                    return

                if event_type == "error":
                    _, message = event
                    self.generation_queue = None
                    self._finish_generation_error(message)
                    return
        except queue.Empty:
            pass

        if self.generation_queue is not None:
            self.root.after(100, self._poll_generation_queue)

    def _finish_generation_success(self, output_path: Path, generated_combinations_path: Path | None) -> None:
        """Завершает успешную генерацию и показывает итоговое сообщение.

        output_path: путь к созданному DOCX-файлу.
        generated_combinations_path: путь к созданному TXT-файлу или None.
        Ничего не возвращает.
        Побочный эффект: закрывает окно прогресса и показывает messagebox об успешном создании.
        """
        self._close_progress_window()

        success_message = f"Итоговый DOCX-файл создан:\n{output_path}"

        if generated_combinations_path is not None:
            success_message += f"\n\nTXT-файл с комбинациями создан:\n{generated_combinations_path}"

        messagebox.showinfo("Готово", success_message)

    def _finish_generation_error(self, message: str) -> None:
        """Завершает генерацию с ошибкой и показывает понятное сообщение.

        message: текст ошибки для пользователя.
        Ничего не возвращает.
        Побочный эффект: закрывает окно прогресса и показывает messagebox с ошибкой.
        """
        self._close_progress_window()
        messagebox.showerror("Ошибка генерации", message)


def run_app() -> None:
    """Создает и запускает tkinter-приложение.

    Входные данные не принимает.
    Ничего не возвращает.
    Побочный эффект: открывает главное окно и запускает цикл событий tkinter.
    """
    # Tk должен быть создан до загрузки системных шрифтов.
    root = tk.Tk()

    # Локальная иконка подключается к окну приложения без зависимости от интернета после скачивания проекта.
    _apply_app_icon(root)

    # Объект приложения хранит состояние интерфейса на протяжении всего сеанса.
    TemplateGeneratorApp(root)

    root.mainloop()


def _apply_app_icon(root: tk.Tk) -> None:
    """Подключает иконку приложения к главному окну tkinter.

    root: главное окно Tk, которому нужно назначить иконку.
    Ничего не возвращает.
    Побочный эффект: загружает PNG-asset и задает его как иконку окна и дочерних окон.
    """
    if not APP_ICON_PATH.exists():
        return

    try:
        # PhotoImage используется, потому что tkinter умеет читать PNG без дополнительных зависимостей.
        icon_image = tk.PhotoImage(file=str(APP_ICON_PATH))
        root.iconphoto(True, icon_image)

        # Ссылка сохраняется на root, иначе сборщик мусора может удалить изображение из памяти tkinter.
        root.app_icon_image = icon_image
    except tk.TclError:
        return
