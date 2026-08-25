"""Reject model-produced numbers that cannot be reconciled to approved source text."""

from __future__ import annotations

import math
import re
from collections.abc import Iterable

from providers.base import SourceFragment


class GroundingError(ValueError):
    pass


_NUMBER = re.compile(
    r"(?<![A-Za-z0-9])(?P<open>\()?\$?(?P<number>-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)(?P<percent>%)?(?P<close>\))?"
)


def _numbers(text: str) -> list[float]:
    values: list[float] = []
    for match in _NUMBER.finditer(text):
        try:
            value = float(match.group("number").replace(",", ""))
        except ValueError:
            continue
        if match.group("open") and match.group("close"):
            value = -abs(value)
        values.append(value)
        if match.group("percent"):
            values.append(value / 100)
    return values


def _payload_numbers(value: object, excluded_keys: frozenset[str]) -> Iterable[float]:
    if isinstance(value, bool) or value is None:
        return
    if isinstance(value, (int, float)):
        yield float(value)
        return
    if isinstance(value, str):
        yield from _numbers(value)
        return
    if isinstance(value, dict):
        for key, child in value.items():
            if str(key) not in excluded_keys:
                yield from _payload_numbers(child, excluded_keys)
        return
    if isinstance(value, (list, tuple)):
        for child in value:
            yield from _payload_numbers(child, excluded_keys)


def verify_numeric_grounding(payload: dict[str, object], sources: tuple[SourceFragment, ...],
                             *, excluded_keys: frozenset[str] = frozenset()) -> None:
    source_values = [number for fragment in sources for number in _numbers(fragment.text)]
    if not source_values:
        if any(True for _ in _payload_numbers(payload, excluded_keys)):
            raise GroundingError("The model output contains numbers but the sources contain no numeric evidence.")
        return

    unmatched = 0
    for produced in _payload_numbers(payload, excluded_keys):
        if not any(
            math.isclose(produced, source, rel_tol=1e-7, abs_tol=1e-7)
            for source in source_values
        ):
            unmatched += 1
    if unmatched:
        raise GroundingError(
            f"{unmatched} model-produced numeric value(s) could not be reconciled to the approved sources."
        )
