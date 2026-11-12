import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from classify import CODE_MAP, classify, normalize  # noqa: E402


def codes(text):
    return {CODE_MAP[h] for h in classify(text)}


def test_normalize_strips_azerbaijani_diacritics():
    assert normalize("işləmir Əyləc Şin göz") == "islemir eylec sin goz"


@pytest.mark.parametrize(
    "text, expected",
    [
        # Azerbaijani (with and without diacritics)
        ("pedal fırlanmır", {"T02"}),
        ("zencir dusur her defe", {"T03"}),
        ("batareya tez bitir", {"T10"}),
        ("GPS yeri səhv göstərir", {"A01"}),
        ("qiymət çox bahadı", {"P01"}),
        # Russian
        ("тормоза не работают", {"T01"}),
        ("слишком дорого", {"P01"}),
        ("не могу завершить поездку в приложении", {"A02"}),
        # English
        ("The brakes are not working", {"T01"}),
        ("Seat is broken", {"T04"}),
        ("Price is too high", {"P01"}),
    ],
)
def test_single_category(text, expected):
    assert codes(text) == expected


def test_multi_label():
    assert codes("Əyləc tutmur, çox təhlükəlidir") == {"T01", "T14"}


def test_positive_feedback():
    assert codes("спасибо, всё отлично") == {"M01"}


def test_russian_price_not_confused_with_road():
    # "дорого" (expensive) must not be tagged as road infrastructure ("дорога")
    assert "I04" not in codes("слишком дорого")


def test_generic_technical_dropped_when_specific_exists():
    # T13 (general technical fault) is a catch-all and must not be added
    # when a specific technical category already explains the comment.
    assert "T13" not in codes("əyləc işləmir")


def test_unmatched_comment_returns_empty():
    assert classify("salam") == []
