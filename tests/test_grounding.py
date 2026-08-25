"""Tests for numeric source reconciliation."""

import unittest

from providers.base import SourceFragment
from quality.grounding import GroundingError, verify_numeric_grounding


class GroundingTests(unittest.TestCase):
    def test_matches_currency_and_percentage_normalization(self) -> None:
        sources = (SourceFragment("data.csv", "row 2", "$1,250.00 | 13.37%"),)

        verify_numeric_grounding(
            {"value": 1250, "return": .1337, "display": "+13.37%"}, sources
        )

    def test_rejects_number_absent_from_sources_without_echoing_it(self) -> None:
        sources = (SourceFragment("data.csv", "row 2", "Return | 5.20%"),)

        with self.assertRaises(GroundingError) as raised:
            verify_numeric_grounding({"return": .0725}, sources)

        self.assertNotIn("0.0725", str(raised.exception))


if __name__ == "__main__":
    unittest.main()
