"""Tests for provider-independent, validation-first model execution."""

import unittest

from core.structured_agent import StructuredAgent, StructuredOutputError
from providers.base import SourceFragment, StructuredRequest
from providers.fake import FixtureProvider


def _request(approved: bool = True) -> StructuredRequest:
    return StructuredRequest(
        task_name="test",
        instructions="Return the supplied value.",
        json_schema={"type": "object", "required": ["value"]},
        sources=(SourceFragment("source.csv", "row 2", "Value | 42"),),
        privacy_approved=approved,
    )


class StructuredAgentTests(unittest.TestCase):
    def test_retries_once_after_contract_failure(self) -> None:
        provider = FixtureProvider([{"wrong": 1}, {"value": 42}])

        def validate(payload):
            if "value" not in payload:
                raise ValueError("value is required")
            return payload

        result = StructuredAgent(provider).run(_request(), validate)

        self.assertEqual(result, {"value": 42})
        self.assertEqual(len(provider.requests), 2)
        self.assertEqual(provider.requests[1].correction_errors, ("value is required",))

    def test_blocks_provider_before_privacy_approval(self) -> None:
        provider = FixtureProvider([{"value": 42}])

        with self.assertRaisesRegex(StructuredOutputError, "privacy"):
            StructuredAgent(provider).run(_request(False), lambda value: value)

        self.assertEqual(provider.requests, [])


if __name__ == "__main__":
    unittest.main()
