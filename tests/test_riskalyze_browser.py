"""Tests for exact Riskalyze household matching and safe portfolio extraction."""

import tempfile
import unittest
from pathlib import Path

from services.riskalyze_browser import (
    RiskalyzeCaptureError,
    exact_client_href,
    parse_riskalyze_portfolio,
    write_riskalyze_source,
)


RISKALYZE_TEXT = """
PORTFOLIO TOTAL
$1,000,000
RISK
51
ANALYTICS
95% Historical Range (6 months)
-$120,000
-12.00%
+$180,000
+18.00%
Stocks 60.00%
Bonds 30.00%
Other 4.00%
Cash 6.00%
Riskalyze GPA 4.0
Annual Dividend 2.00%
Max Drawdown -15.00%
Annual Range Midpoint 7.00%
Portfolio Costs 0.25%
"""


class RiskalyzeBrowserTests(unittest.TestCase):
    def test_exact_match_ignores_similar_households(self) -> None:
        href = exact_client_href(
            "Sample Household",
            (
                ("Sample Person", "/client-details/1/overview"),
                ("Sample Household\nActive", "/client-details/2/overview"),
                ("Sample Householder", "/client-details/3/overview"),
            ),
        )

        self.assertEqual(href, "/client-details/2/overview")

    def test_ambiguous_exact_matches_are_blocked(self) -> None:
        with self.assertRaisesRegex(RiskalyzeCaptureError, "More than one exact"):
            exact_client_href(
                "Sample Household",
                (
                    ("Sample Household", "/client-details/1/overview"),
                    ("Sample Household", "/client-details/2/overview"),
                ),
            )

    def test_visible_portfolio_text_becomes_safe_grounded_csv(self) -> None:
        data = parse_riskalyze_portfolio(RISKALYZE_TEXT)
        with tempfile.TemporaryDirectory() as directory:
            destination = Path(directory) / "riskalyze.csv"
            write_riskalyze_source(data, destination)
            output = destination.read_text(encoding="utf-8")

        self.assertIn("Portfolio total,\"$1,000,000\"", output)
        self.assertIn("Historical loss %,-12.00%", output)
        self.assertIn("Portfolio costs,0.25%", output)
        self.assertIn("Allocation,Stocks,600000.00,60.00%", output)
        self.assertNotIn("Sample Household", output)
        self.assertNotIn("account", output.casefold())

    def test_live_dom_order_without_risk_label_is_supported(self) -> None:
        live_order = RISKALYZE_TEXT.replace("RISK\n51", "51").replace(
            "Portfolio Costs 0.25%",
            "Portfolio Costs\n0%\n1%\n0.25%\nEst. Tax Drag\n0.35%\nExpense Ratio",
        )

        data = parse_riskalyze_portfolio(live_order)

        self.assertEqual(data.risk_number, "51")
        self.assertEqual(data.portfolio_costs, "0.25%")
        self.assertEqual(data.expense_ratio, "")


if __name__ == "__main__":
    unittest.main()
