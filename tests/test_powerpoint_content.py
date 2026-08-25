"""Tests for the strict PowerPoint JSON boundary."""

import json
import tempfile
import unittest
from pathlib import Path

from generators.powerpoint_content import (
    DeckContentError,
    deck_content_schema,
    validate_deck_content,
)


class PowerPointContentTests(unittest.TestCase):
    def test_provider_schema_requires_bounded_slides(self) -> None:
        schema = deck_content_schema()
        slides = schema["properties"]["slides"]
        self.assertEqual(slides["maxItems"], 40)
        self.assertEqual(slides["items"]["properties"]["bullets"]["maxItems"], 8)

    def test_normalizes_supported_layouts_and_adds_source_footer(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "content.json"
            source.write_text(json.dumps({
                "title": "Ignored in favor of reviewed title",
                "slides": [
                    {"title": "Markets remain resilient", "bullets": ["Earnings remain positive.", "Breadth improved."]},
                    {"title": "Selected data", "table": {"headers": ["Metric", "Value"], "rows": [["Return", "+5.20%"]]}},
                ],
            }), encoding="utf-8")

            result = validate_deck_content(source, report_title="Quarterly Review")

            self.assertEqual(result["title"], "Quarterly Review")
            self.assertEqual(len(result["slides"]), 2)
            self.assertEqual(result["slides"][0]["footer_right"], "Source: Uploaded content")

    def test_chart_requires_uploaded_image(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "content.json"
            source.write_text(json.dumps({
                "slides": [{"title": "Chart", "bullets": ["Takeaway"], "chart": True, "chart_image": "chart.png"}]
            }), encoding="utf-8")

            with self.assertRaisesRegex(DeckContentError, "uploaded chart/image"):
                validate_deck_content(source, report_title="Review")

    def test_rejects_overfilled_slide(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "content.json"
            source.write_text(json.dumps({
                "slides": [{"title": "Too much", "bullets": [f"Item {index}" for index in range(9)]}]
            }), encoding="utf-8")

            with self.assertRaisesRegex(DeckContentError, "1–8"):
                validate_deck_content(source, report_title="Review")


if __name__ == "__main__":
    unittest.main()
