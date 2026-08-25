"""Free local Ollama structured-output adapter."""

from __future__ import annotations

import json
import urllib.error
import urllib.request
from dataclasses import dataclass
from urllib.parse import urlparse

from providers.base import ProviderError, StructuredRequest


@dataclass(frozen=True, slots=True)
class OllamaProvider:
    model: str
    endpoint: str = "http://127.0.0.1:11434"
    timeout_seconds: int = 420
    provider_id: str = "ollama"

    @property
    def external(self) -> bool:
        hostname = (urlparse(self.endpoint).hostname or "").casefold()
        return hostname not in {"127.0.0.1", "localhost", "::1"}

    def generate_json(self, request: StructuredRequest) -> dict[str, object]:
        if not self.model.strip():
            raise ProviderError("An Ollama model name is required.")
        payload = {
            "model": self.model,
            "system": (
                "You are the structured-data synthesis component inside Reportus. "
                "Global privacy, accuracy, and source-faithfulness rules are mandatory."
            ),
            "prompt": request.prompt(),
            "format": request.json_schema,
            "stream": False,
            "think": False,
            "options": {"temperature": 0},
        }
        body = json.dumps(payload).encode("utf-8")
        http_request = urllib.request.Request(
            self.endpoint.rstrip("/") + "/api/generate",
            data=body,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(http_request, timeout=self.timeout_seconds) as response:
                envelope = json.loads(response.read().decode("utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError, urllib.error.URLError) as exc:
            raise ProviderError("The configured Ollama service did not return a usable response.") from exc
        raw = envelope.get("response") if isinstance(envelope, dict) else None
        if not isinstance(raw, str):
            raise ProviderError("Ollama returned no structured response.")
        try:
            result = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ProviderError("Ollama returned invalid JSON.") from exc
        if not isinstance(result, dict):
            raise ProviderError("Ollama must return one JSON object.")
        return result
