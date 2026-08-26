"""Small provider factory driven by environment configuration."""

from __future__ import annotations

import os

from providers.base import ProviderError, StructuredProvider
from providers.ollama import OllamaProvider


class DisabledProvider:
    provider_id = "disabled"
    external = False

    def generate_json(self, request):
        raise ProviderError("AI synthesis is not configured. Select a provider in the future admin settings.")


def provider_from_environment() -> StructuredProvider:
    provider_id = os.environ.get(
        "REPORTICLES_AI_PROVIDER", os.environ.get("REPORTUS_AI_PROVIDER", "disabled")
    ).strip().casefold()
    if provider_id == "ollama":
        provider = OllamaProvider(
            model=os.environ.get(
                "REPORTICLES_OLLAMA_MODEL", os.environ.get("REPORTUS_OLLAMA_MODEL", "")
            ).strip(),
            endpoint=os.environ.get(
                "REPORTICLES_OLLAMA_ENDPOINT",
                os.environ.get("REPORTUS_OLLAMA_ENDPOINT", "http://127.0.0.1:11434"),
            ).strip(),
        )
        allow_external = os.environ.get(
            "REPORTICLES_ALLOW_EXTERNAL_AI",
            os.environ.get("REPORTUS_ALLOW_EXTERNAL_AI", ""),
        ).strip().casefold() == "true"
        if provider.external and not allow_external:
            raise ProviderError(
                "Remote AI processing is disabled. Use local Ollama or explicitly enable external processing."
            )
        return provider
    if provider_id in {"", "disabled", "none"}:
        return DisabledProvider()
    raise ProviderError(f"Unsupported AI provider: {provider_id}")
