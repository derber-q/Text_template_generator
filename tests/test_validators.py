"""Тесты проверок пользовательских данных."""

# Импортируется pytest, чтобы проверять ожидаемые ошибки валидации.
import pytest

# Импортируются валидаторы, которые используются GUI перед генерацией.
from validators import (
    ValidationError,
    parse_forbidden_combinations,
    parse_positive_int,
    validate_docx_path,
    validate_optional_txt_path,
    validate_output_docx_path,
    validate_static_part,
)


def test_parse_positive_int_accepts_positive_integer() -> None:
    """Проверяет корректное положительное целое число."""
    assert parse_positive_int(" 12 ", "Количество") == 12


def test_parse_positive_int_rejects_non_positive_or_text() -> None:
    """Проверяет отклонение нуля и нечислового текста."""
    with pytest.raises(ValidationError):
        parse_positive_int("0", "Количество")

    with pytest.raises(ValidationError):
        parse_positive_int("abc", "Количество")


def test_validate_static_part_rejects_empty_value_and_number_sign() -> None:
    """Проверяет запрет пустой постоянной части и символа «№»."""
    assert validate_static_part(" ABC- ") == "ABC-"

    with pytest.raises(ValidationError):
        validate_static_part("")

    with pytest.raises(ValidationError):
        validate_static_part("№ABC-")


def test_validate_docx_and_txt_paths(tmp_path) -> None:
    """Проверяет расширения и существование файлов."""
    # Валидация пути проверяет только путь и расширение, не содержимое Word-документа.
    docx_path = tmp_path / "template.docx"
    docx_path.write_bytes(b"placeholder")

    txt_path = tmp_path / "forbidden.txt"
    txt_path.write_text("123", encoding="utf-8")

    assert validate_docx_path(str(docx_path)) == docx_path
    assert validate_optional_txt_path(str(txt_path)) == txt_path
    assert validate_optional_txt_path("") is None

    with pytest.raises(ValidationError):
        validate_docx_path(str(txt_path))

    with pytest.raises(ValidationError):
        validate_docx_path("")


def test_validate_output_docx_path(tmp_path) -> None:
    """Проверяет путь сохранения итогового документа."""
    output_path = tmp_path / "result.docx"

    assert validate_output_docx_path(str(output_path)) == output_path

    with pytest.raises(ValidationError):
        validate_output_docx_path(str(tmp_path / "result.txt"))

    with pytest.raises(ValidationError):
        validate_output_docx_path("")


def test_parse_forbidden_combinations_accepts_valid_text_and_deduplicates() -> None:
    """Проверяет корректный формат `.txt` с дублями."""
    result = parse_forbidden_combinations("123,456,123,", 3)

    assert result == {"123", "456"}


def test_parse_forbidden_combinations_rejects_non_digits_and_wrong_length() -> None:
    """Проверяет ошибки формата `.txt`."""
    with pytest.raises(ValidationError):
        parse_forbidden_combinations("123,45a", 3)

    with pytest.raises(ValidationError):
        parse_forbidden_combinations("123,45", 3)
