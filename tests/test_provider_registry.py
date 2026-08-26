"""Tests for safe provider defaults and remote-service opt in."""

import os
import unittest
from unittest.mock import patch

from providers.base import ProviderError
from providers.registry import DisabledProvider, provider_from_environment


class ProviderRegistryTests(unittest.TestCase):
    def test_provider_is_disabled_by_default(self) -> None:
        with patch.dict(os.environ, {}, clear=True):
            self.assertIsInstance(provider_from_environment(), DisabledProvider)

    def test_remote_ollama_requires_explicit_opt_in(self) -> None:
        with patch.dict(os.environ, {
            "REPORTICLES_AI_PROVIDER": "ollama",
            "REPORTICLES_OLLAMA_MODEL": "test-model",
            "REPORTICLES_OLLAMA_ENDPOINT": "https://models.example.com",
        }, clear=True):
            with self.assertRaisesRegex(ProviderError, "Remote AI processing"):
                provider_from_environment()

    def test_legacy_environment_names_remain_supported(self) -> None:
        with patch.dict(os.environ, {"REPORTUS_AI_PROVIDER": "disabled"}, clear=True):
            self.assertIsInstance(provider_from_environment(), DisabledProvider)


if __name__ == "__main__":
    unittest.main()
