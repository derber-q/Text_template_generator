"""Генерация уникальных номеров и чтение запрещенных комбинаций."""

# Импортируется random, чтобы получать случайные цифровые части номера без внешних сервисов.
import random

# Импортируется sys, чтобы в `.exe`-сборке искать редактируемый файл исключений рядом с программой.
import sys

# Импортируется Path, чтобы читать выбранный пользователем `.txt`-файл.
from pathlib import Path

# Импортируется функция проверки текста, чтобы формат `.txt` контролировался в одном месте.
from validators import ValidationError, parse_forbidden_combinations


# Имя постоянного файла запретов используется и в исходниках, и рядом с собранным `.exe`.
PERMANENT_FORBIDDEN_COMBINATIONS_FILENAME = "permanent_forbidden_combinations.txt"

# Постоянный файл запретов лежит рядом с исходными модулями, а в `.exe`-сборке — рядом с исполняемым файлом.
PERMANENT_FORBIDDEN_COMBINATIONS_FILE = (
    Path(sys.executable).resolve().with_name(PERMANENT_FORBIDDEN_COMBINATIONS_FILENAME)
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().with_name(PERMANENT_FORBIDDEN_COMBINATIONS_FILENAME)
)


class NumberGenerationError(ValueError):
    """Ошибка генерации номеров, пригодная для показа пользователю.

    Класс принимает текст ошибки через стандартный ValueError.
    Ничего не возвращает.
    Побочный эффект: останавливает генерацию, если уникальные номера создать нельзя.
    """


def load_forbidden_combinations(
    file_path: Path | None,
    digit_count: int,
    permanent_file_path: Path | None = PERMANENT_FORBIDDEN_COMBINATIONS_FILE,
) -> set[str]:
    """Загружает постоянные и пользовательские запрещенные цифровые комбинации.

    file_path: путь к пользовательскому файлу или None, если пользователь его не выбрал.
    digit_count: ожидаемая длина каждой комбинации.
    permanent_file_path: путь к постоянному корневому файлу запретов или None для отключения.
    Возвращает объединенное множество запрещенных комбинаций.
    Побочный эффект: читает файл с диска и выбрасывает ValidationError при ошибке чтения или формата.
    """
    # Объединенное множество автоматически убирает дубликаты между постоянным и пользовательским файлами.
    forbidden_combinations: set[str] = set()

    if permanent_file_path is not None:
        forbidden_combinations.update(
            _load_forbidden_file(
                permanent_file_path,
                digit_count,
                "постоянный файл запрещенных комбинаций",
                missing_is_empty=True,
            )
        )

    if file_path is not None:
        forbidden_combinations.update(
            _load_forbidden_file(
                file_path,
                digit_count,
                "выбранный файл запрещенных комбинаций",
                missing_is_empty=False,
            )
        )

    return forbidden_combinations


def _load_forbidden_file(
    file_path: Path,
    digit_count: int,
    file_description: str,
    missing_is_empty: bool,
) -> set[str]:
    """Читает один `.txt`-файл запрещенных комбинаций.

    file_path: путь к файлу запретов.
    digit_count: ожидаемая длина каждой комбинации.
    file_description: описание файла для понятного сообщения об ошибке.
    missing_is_empty: считать ли отсутствующий файл пустым списком запретов.
    Возвращает множество запрещенных комбинаций из файла.
    Побочный эффект: читает файл с диска и выбрасывает ValidationError при ошибке чтения или формата.
    """
    if not file_path.is_file():
        if missing_is_empty:
            return set()

        raise ValidationError(f"Не найден {file_description}: {file_path}")

    try:
        # Кодировка utf-8-sig корректно обрабатывает обычный UTF-8 и возможный BOM в начале файла.
        raw_text = file_path.read_text(encoding="utf-8-sig")
    except OSError as exc:
        raise ValidationError(f"Не удалось прочитать {file_description}: {exc}") from exc

    return parse_forbidden_combinations(raw_text, digit_count)


def calculate_available_count(digit_count: int, forbidden_combinations: set[str]) -> int:
    """Считает количество доступных случайных комбинаций.

    digit_count: количество случайных цифр.
    forbidden_combinations: множество запрещенных комбинаций нужной длины.
    Возвращает число доступных вариантов.
    Побочных эффектов нет.
    """
    # Общее число вариантов для d цифр равно 10 ** d, включая варианты с ведущими нулями.
    total_count = 10**digit_count

    return total_count - len(forbidden_combinations)


def ensure_generation_capacity(
    copy_count: int,
    digit_count: int,
    forbidden_combinations: set[str],
) -> None:
    """Проверяет, можно ли создать нужное количество уникальных случайных частей.

    copy_count: количество копий документа.
    digit_count: количество случайных цифр в номере.
    forbidden_combinations: множество запрещенных комбинаций.
    Ничего не возвращает.
    Побочный эффект: выбрасывает NumberGenerationError, если доступных комбинаций недостаточно.
    """
    # Доступное количество важно проверить до генерации, чтобы не запускать долгий случайный подбор зря.
    available_count = calculate_available_count(digit_count, forbidden_combinations)

    if copy_count > available_count:
        raise NumberGenerationError(
            "Невозможно создать нужное количество уникальных номеров. "
            f"Требуется: {copy_count}, доступно: {available_count}."
        )


def generate_unique_digit_parts(
    copy_count: int,
    digit_count: int,
    forbidden_combinations: set[str] | None = None,
    rng: random.Random | None = None,
) -> list[str]:
    """Генерирует уникальные случайные цифровые части номера.

    copy_count: сколько комбинаций нужно создать.
    digit_count: длина каждой случайной цифровой части.
    forbidden_combinations: комбинации, которые нельзя использовать.
    rng: необязательный генератор случайных чисел для тестов.
    Возвращает список уникальных строк фиксированной длины.
    Побочный эффект: использует генератор случайных чисел и может выбросить NumberGenerationError.
    """
    # Пустой набор запрещенных комбинаций упрощает дальнейшую проверку принадлежности.
    forbidden_set = forbidden_combinations or set()

    ensure_generation_capacity(copy_count, digit_count, forbidden_set)

    # Генератор можно подменить в тестах, чтобы получать воспроизводимый результат.
    random_source = rng or random.Random()

    # Все созданные комбинации хранятся в множестве для быстрой проверки уникальности.
    generated_set: set[str] = set()

    # Порядковый список нужен, чтобы сохранить число комбинаций и передать их в том же порядке в DOCX.
    generated_parts: list[str] = []

    # Верхняя граница используется для randrange и сохраняет ведущие нули при форматировании.
    upper_bound = 10**digit_count

    # Ограничение попыток защищает от бесконечного цикла при очень плотном списке запретов.
    attempt_limit = max(copy_count * 100, 1000)

    # Счетчик попыток помогает вовремя перейти к гарантированному последовательному добору.
    attempts = 0

    while len(generated_parts) < copy_count and attempts < attempt_limit:
        attempts += 1

        # Форматирование с ведущими нулями сохраняет точную длину случайной части.
        candidate = f"{random_source.randrange(upper_bound):0{digit_count}d}"

        if candidate in forbidden_set or candidate in generated_set:
            continue

        generated_set.add(candidate)
        generated_parts.append(candidate)

    if len(generated_parts) == copy_count:
        return generated_parts

    # Последовательный добор нужен для редких случаев, когда случайный подбор часто попадает в запреты.
    for number in range(upper_bound):
        if len(generated_parts) == copy_count:
            break

        candidate = f"{number:0{digit_count}d}"

        if candidate in forbidden_set or candidate in generated_set:
            continue

        generated_set.add(candidate)
        generated_parts.append(candidate)

    if len(generated_parts) != copy_count:
        raise NumberGenerationError("Не удалось подобрать достаточное количество уникальных номеров.")

    return generated_parts


def build_full_numbers(static_part: str, digit_parts: list[str]) -> list[str]:
    """Собирает итоговые номера из постоянной и случайной частей.

    static_part: постоянная часть номера без символа «№».
    digit_parts: список случайных цифровых частей.
    Возвращает номера формата `№<статическая_часть><случайные_цифры>`.
    Побочных эффектов нет.
    """
    # Символ «№» добавляется централизованно, чтобы пользователь не вводил его вручную.
    return [f"№{static_part}{digit_part}" for digit_part in digit_parts]


def build_generated_combinations_path(output_docx_path: Path) -> Path:
    """Формирует путь для `.txt` со сгенерированными цифровыми комбинациями.

    output_docx_path: путь к итоговому `.docx`-файлу.
    Возвращает путь вида `<имя_docx>_generated_combinations.txt` рядом с итоговым документом.
    Побочных эффектов нет.
    """
    # Имя строится от итогового DOCX, чтобы пользователь легко связал оба результата одной генерации.
    return output_docx_path.with_name(f"{output_docx_path.stem}_generated_combinations.txt")


def save_generated_combinations(digit_parts: list[str], output_docx_path: Path) -> Path:
    """Сохраняет сгенерированные случайные цифровые части в `.txt`.

    digit_parts: список случайных цифровых частей, которые были использованы при генерации.
    output_docx_path: путь к итоговому `.docx`, рядом с которым нужно создать `.txt`.
    Возвращает путь к созданному `.txt`-файлу.
    Побочный эффект: записывает файл на диск и выбрасывает ValidationError при ошибке записи.
    """
    # TXT-файл должен быть пригоден для повторной передачи как список исключений, поэтому разделитель — запятая без пробелов.
    output_txt_path = build_generated_combinations_path(output_docx_path)

    try:
        output_txt_path.write_text(",".join(digit_parts), encoding="utf-8")
    except OSError as exc:
        raise ValidationError(f"Не удалось сохранить TXT-файл со сгенерированными комбинациями: {exc}") from exc

    return output_txt_path
