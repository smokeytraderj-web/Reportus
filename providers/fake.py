"""Deterministic test provider; never enabled as the production default."""

from __future__ import annotations

from dataclasses import dataclass, field

from providers.base import ProviderError, StructuredRequest


@dataclass(slots=True)
class FixtureProvider:
    responses: list[dict[str, object]] = field(default_factory=list)
    provider_id: str = "fixture"
    external: bool = False
    requests: list[StructuredRequest] = field(default_factory=list)

    def generate_json(self, request: StructuredRequest) -> dict[str, object]:
        self.requests.append(request)
        if not self.responses:
            raise ProviderError("No fixture response is configured.")
        return self.responses.pop(0)
