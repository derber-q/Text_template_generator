"""Тесты генерации уникальных случайных частей номера."""

# Импортируется pytest, чтобы проверять ожидаемые исключения.
import pytest

# Импортируется Path, чтобы проверять построение соседнего TXT-пути.
from pathlib import Path

# Импортируются проверяемые функции генерации номеров.
from number_generator import (
    NumberGenerationError,
    build_generated_combinations_path,
    build_full_numbers,
    generate_unique_digit_parts,
    load_forbidden_combinations,
    save_generated_combinations,
)


class FixedRandom:
    """Тестовый генератор случайных чисел с заранее заданными значениями.

    Класс принимает список чисел, которые будут возвращаться по очереди.
    Возвращает значения через метод randrange.
    Побочный эффект: сдвигает внутренний индекс после каждого вызова.
    """

    def __init__(self, values: list[int]) -> None:
        """Сохраняет значения для последующих вызовов.

        values: список чисел, которые должен вернуть генератор.
        Ничего не возвращает.
        Побочный эффект: создает внутренний индекс текущей позиции.
        """
        # Значения копируются, чтобы тесты не зависели от внешнего списка.
        self.values = list(values)

        # Индекс показывает, какое значение будет возвращено следующим.
        self.index = 0

    def randrange(self, upper_bound: int) -> int:
        """Возвращает следующее тестовое значение.

        upper_bound: верхняя граница, совместимая с интерфейсом random.Random.randrange.
        Возвращает очередное число из values.
        Побочный эффект: увеличивает внутренний индекс.
        """
        # Проверка верхней границы помогает убедиться, что тестовый генератор вызван ожидаемо.
        assert upper_bound > 0

        # Текущее значение имитирует результат случайного выбора.
        value = self.values[self.index]
        self.index += 1

        return value


def test_generate_unique_digit_parts_preserves_leading_zeroes() -> None:
    """Проверяет сохранение ведущих нулей в случайной части."""
    # Генератор возвращает число 7, которое для длины 3 должно стать строкой "007".
    rng = FixedRandom([7])

    result = generate_unique_digit_parts(1, 3, rng=rng)

    assert result == ["007"]


def test_generate_unique_digit_parts_avoids_duplicates_and_forbidden_values() -> None:
    """Проверяет уникальность и учет запрещенных комбинаций."""
    # Первые два значения должны быть пропущены: одно запрещено, второе повторяется.
    rng = FixedRandom([12, 12, 13, 14, 15])

    result = generate_unique_digit_parts(3, 2, forbidden_combinations={"12"}, rng=rng)

    assert result == ["13", "14", "15"]
    assert len(result) == len(set(result))
    assert "12" not in result


def test_generate_unique_digit_parts_raises_when_capacity_is_not_enough() -> None:
    """Проверяет ошибку при нехватке доступных комбинаций."""
    # Для одной цифры всего 10 вариантов, из которых 9 запрещены.
    forbidden = {str(number) for number in range(9)}

    with pytest.raises(NumberGenerationError):
        generate_unique_digit_parts(2, 1, forbidden_combinations=forbidden)


def test_load_forbidden_combinations_reads_unique_values(tmp_path) -> None:
    """Проверяет чтение `.txt` и удаление дублей запрещенных комбинаций."""
    # Файл содержит дубль и случайную завершающую запятую.
    file_path = tmp_path / "forbidden.txt"
    file_path.write_text("123,123,045,", encoding="utf-8")

    result = load_forbidden_combinations(file_path, 3, permanent_file_path=None)

    assert result == {"123", "045"}


def test_load_forbidden_combinations_uses_permanent_file_without_user_file(tmp_path) -> None:
    """Проверяет учет постоянного файла, даже если пользовательский файл не выбран."""
    # Постоянный файл имитирует корневой список исключений проекта.
    permanent_path = tmp_path / "permanent_forbidden_combinations.txt"
    permanent_path.write_text("111,222", encoding="utf-8")

    result = load_forbidden_combinations(None, 3, permanent_file_path=permanent_path)

    assert result == {"111", "222"}


def test_load_forbidden_combinations_merges_permanent_and_user_files(tmp_path) -> None:
    """Проверяет объединение постоянных и пользовательских исключений."""
    # Дубликат "222" должен учитываться как одна запрещенная комбинация.
    permanent_path = tmp_path / "permanent_forbidden_combinations.txt"
    permanent_path.write_text("111,222", encoding="utf-8")

    user_path = tmp_path / "user_forbidden.txt"
    user_path.write_text("222,333", encoding="utf-8")

    result = load_forbidden_combinations(user_path, 3, permanent_file_path=permanent_path)

    assert result == {"111", "222", "333"}


def test_build_full_numbers_adds_number_sign_and_static_part() -> None:
    """Проверяет сборку итогового формата номера."""
    result = build_full_numbers("ABC-", ["001", "999"])

    assert result == ["№ABC-001", "№ABC-999"]


def test_save_generated_combinations_creates_txt_next_to_docx(tmp_path) -> None:
    """Проверяет сохранение использованных комбинаций рядом с итоговым DOCX."""
    # Путь к итоговому документу нужен только для построения имени соседнего TXT-файла.
    output_docx_path = tmp_path / "result.docx"

    output_txt_path = save_generated_combinations(["001", "123", "999"], output_docx_path)

    assert output_txt_path == tmp_path / "result_generated_combinations.txt"
    assert output_txt_path.read_text(encoding="utf-8") == "001,123,999"


def test_build_generated_combinations_path_uses_docx_stem() -> None:
    """Проверяет предсказуемое имя TXT-файла с комбинациями."""
    output_docx_path = Path("folder") / "generated_document.docx"

    assert build_generated_combinations_path(output_docx_path) == Path("folder") / "generated_document_generated_combinations.txt"
