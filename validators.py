"""Проверки пользовательских данных для генератора документов."""

# Импортируется Path, чтобы безопасно работать с расширениями и родительскими каталогами файлов.
from pathlib import Path


class ValidationError(ValueError):
    """Ошибка проверки данных, которую можно показать пользователю.

    Класс не принимает специальных параметров сверх обычного текста ошибки.
    Возвращаемого значения нет, потому что исключение используется для остановки некорректного сценария.
    Побочный эффект: передает понятное сообщение вызывающему коду.
    """


def parse_positive_int(value: str, field_name: str) -> int:
    """Преобразует строку в положительное целое число.

    value: строка из поля ввода.
    field_name: человекочитаемое название поля для текста ошибки.
    Возвращает положительное целое число.
    Побочный эффект: выбрасывает ValidationError, если значение не подходит.
    """
    # Очищенное значение нужно, чтобы случайные пробелы вокруг числа не ломали корректный ввод.
    cleaned_value = str(value).strip()

    if not cleaned_value:
        raise ValidationError(f"Поле «{field_name}» обязательно для заполнения.")

    if not cleaned_value.isdigit():
        raise ValidationError(f"Поле «{field_name}» должно быть положительным целым числом.")

    # Числовое значение используется дальше для расчетов и проверок вместимости комбинаций.
    parsed_value = int(cleaned_value)

    if parsed_value <= 0:
        raise ValidationError(f"Поле «{field_name}» должно быть больше нуля.")

    return parsed_value


def validate_static_part(value: str) -> str:
    """Проверяет постоянную часть номера.

    value: строка, введенная пользователем без символа «№».
    Возвращает очищенную постоянную часть номера.
    Побочный эффект: выбрасывает ValidationError при пустом значении или наличии символа «№».
    """
    # Очищенная постоянная часть сохраняет намеренный текст, но убирает случайные края.
    static_part = str(value).strip()

    if not static_part:
        raise ValidationError("Введите постоянную часть номера без символа «№».")

    if "№" in static_part:
        raise ValidationError("Не вводите символ «№»: программа добавит его автоматически.")

    return static_part


def validate_docx_path(path_value: str, field_name: str = "DOCX-файл") -> Path:
    """Проверяет путь к существующему `.docx`-файлу.

    path_value: путь, выбранный пользователем.
    field_name: название поля для сообщения об ошибке.
    Возвращает Path к существующему файлу.
    Побочный эффект: выбрасывает ValidationError при пустом пути, неверном расширении или отсутствии файла.
    """
    # Очищенная строка нужна до создания Path, потому что Path("") превращается в текущий каталог.
    cleaned_path = str(path_value).strip()

    if not cleaned_path:
        raise ValidationError(f"Выберите файл для поля «{field_name}».")

    # Path используется, чтобы одинаково обрабатывать абсолютные и относительные пути.
    file_path = Path(cleaned_path)

    if file_path.suffix.lower() != ".docx":
        raise ValidationError(f"Файл в поле «{field_name}» должен иметь расширение .docx.")

    if not file_path.is_file():
        raise ValidationError(f"Файл в поле «{field_name}» не найден.")

    return file_path


def validate_optional_txt_path(path_value: str) -> Path | None:
    """Проверяет необязательный путь к `.txt`-файлу запрещенных комбинаций.

    path_value: путь из поля ввода, который может быть пустым.
    Возвращает Path, если файл выбран, или None, если пользователь оставил поле пустым.
    Побочный эффект: выбрасывает ValidationError при неверном расширении или отсутствии файла.
    """
    # Пустая строка означает, что внешний список запрещенных комбинаций не используется.
    cleaned_path = str(path_value).strip()

    if not cleaned_path:
        return None

    # Путь нужен для проверки расширения и существования выбранного файла.
    file_path = Path(cleaned_path)

    if file_path.suffix.lower() != ".txt":
        raise ValidationError("Файл запрещенных комбинаций должен иметь расширение .txt.")

    if not file_path.is_file():
        raise ValidationError("Файл запрещенных комбинаций не найден.")

    return file_path


def validate_output_docx_path(path_value: str) -> Path:
    """Проверяет путь сохранения итогового `.docx`.

    path_value: путь, выбранный пользователем через save dialog.
    Возвращает Path для записи итогового документа.
    Побочный эффект: выбрасывает ValidationError, если путь пустой, имеет неверное расширение или каталог не найден.
    """
    # Очищенная строка нужна до создания Path, потому что Path("") превращается в текущий каталог.
    cleaned_path = str(path_value).strip()

    if not cleaned_path:
        raise ValidationError("Выберите путь для сохранения итогового .docx-файла.")

    # Очищенный путь исключает ошибку из-за случайного пробела в начале или конце строки.
    output_path = Path(cleaned_path)

    if output_path.suffix.lower() != ".docx":
        raise ValidationError("Итоговый файл должен иметь расширение .docx.")

    if not output_path.parent.exists():
        raise ValidationError("Каталог для сохранения итогового файла не найден.")

    return output_path


def parse_forbidden_combinations(raw_text: str, digit_count: int) -> set[str]:
    """Разбирает текстовый список запрещенных цифровых комбинаций.

    raw_text: содержимое `.txt`-файла, где комбинации разделены запятыми.
    digit_count: ожидаемая длина каждой комбинации.
    Возвращает множество уникальных запрещенных комбинаций.
    Побочный эффект: выбрасывает ValidationError, если формат файла некорректен.
    """
    # Обрезка внешних переводов строк удобна для файлов, сохраненных обычным текстовым редактором.
    cleaned_text = raw_text.strip()

    if not cleaned_text:
        return set()

    # Список элементов нужен, чтобы сохранить понятную проверку каждого значения из файла.
    raw_items = cleaned_text.split(",")

    # Множество автоматически учитывает дубликаты как одну запрещенную комбинацию.
    forbidden_combinations: set[str] = set()

    for position, raw_item in enumerate(raw_items, start=1):
        # Пустые элементы допускаются как случайный лишний разделитель и просто пропускаются.
        item = raw_item

        if item == "":
            continue

        if not item.isdigit():
            raise ValidationError(
                "Файл запрещенных комбинаций должен содержать только цифры и запятые. "
                f"Ошибка в элементе {position}: «{item}»."
            )

        if len(item) != digit_count:
            raise ValidationError(
                "Длина каждой запрещенной комбинации должна совпадать с количеством случайных цифр. "
                f"Ошибка в элементе {position}: «{item}»."
            )

        forbidden_combinations.add(item)

    return forbidden_combinations
