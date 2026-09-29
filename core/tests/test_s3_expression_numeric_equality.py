"""S3 expression comparisons preserve decimal arithmetic in financial rules."""

import pytest

from base_audit.systems.s3_central_statistics.expression_parser import (
    evaluate_boolean,
    evaluate_expression,
)


@pytest.mark.parametrize(
    ("variables", "expected_difference"),
    [
        (
            {
                "a": -100584251.06,
                "b": -95140512.27,
                "c": -21955995.57,
                "d": -16512256.78,
            },
            -5443738.79,
        ),
        (
            {
                "a": -100084151.06,
                "b": -98828212.27,
                "c": -8468595.57,
                "d": -7212656.78,
            },
            -1255938.79,
        ),
    ],
    ids=["2026-08-rows-919-920", "2026-08-rows-1015-1016"],
)
def test_equal_decimal_differences_do_not_trigger_not_equal(
    variables: dict[str, float], expected_difference: float,
) -> None:
    # The legacy float results differ by a few binary ulps, although both
    # source pairs have the same exact decimal difference.
    assert evaluate_expression("a-b", variables) != evaluate_expression("c-d", variables)
    assert evaluate_boolean("(a-b)=(c-d)", variables) is True
    assert evaluate_boolean("(a-b)<>(c-d)", variables) is False
    assert evaluate_expression("a-b", variables) == pytest.approx(expected_difference)


def test_decimal_addition_compares_without_binary_tail() -> None:
    assert evaluate_boolean("0.1+0.2=0.3") is True
    assert evaluate_boolean("0.1+0.2<>0.3") is False


def test_comparison_keeps_small_real_decimal_difference() -> None:
    variables = {"left": 1000000.00000001, "right": 1000000.00000002}

    assert evaluate_boolean("left<>right", variables) is True
    assert evaluate_boolean("left<right", variables) is True
    assert evaluate_boolean("left=right", variables) is False
