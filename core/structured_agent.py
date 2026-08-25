"""Validate every model response before it can reach a report generator."""

from __future__ import annotations

from dataclasses import replace
from typing import Callable

from providers.base import ProviderError, StructuredProvider, StructuredRequest


class StructuredOutputError(RuntimeError):
    pass


Validator = Callable[[dict[str, object]], dict[str, object]]


class StructuredAgent:
    def __init__(self, provider: StructuredProvider, *, attempts: int = 2):
        if attempts not in {1, 2}:
            raise ValueError("Structured generation supports one attempt and one controlled retry.")
        self.provider = provider
        self.attempts = attempts

    def run(self, request: StructuredRequest, validator: Validator) -> dict[str, object]:
        if not request.privacy_approved:
            raise StructuredOutputError("AI processing is blocked until the local privacy check passes.")
        current = request
        last_errors: tuple[str, ...] = ()
        for attempt in range(self.attempts):
            try:
                candidate = self.provider.generate_json(current)
                return validator(candidate)
            except ProviderError:
                raise
            except (KeyError, TypeError, ValueError, StructuredOutputError) as exc:
                last_errors = (str(exc) or "Structured output did not match the report contract.",)
                if attempt + 1 < self.attempts:
                    current = replace(request, correction_errors=last_errors)
        raise StructuredOutputError("AI output failed report validation after one retry: " + last_errors[0])
