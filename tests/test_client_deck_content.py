"""Tests for Client Deck structured-output normalization."""

import unittest

from generators.client_deck_content import (
    ClientDeckContentError,
    normalize_client_deck_payload,
)


def _payload() -> dict[str, object]:
    return {
        "allocation": [{"label": "Equity", "value": 600000}, {"label": "Fixed Income", "value": 400000}],
        "risk_metrics": {
            "Portfolio total": "$1,000,000.00", "Risk": "53",
            "Historical loss": "-$105,500", "Historical loss %": "-10.55%",
            "Historical gain": "+$189,700", "Historical gain %": "+18.97%",
            "Annual dividend": "2.26%", "Max drawdown": "-13.07%",
            "Annual range midpoint": "8.60%", "Expense ratio": "0.26%",
        },
        "sector_performance": {"Technology": .14, "Energy": -.03},
        "sector_portfolio": {"Technology": 30.0, "Energy": 4.0},
        "sector_benchmark": {"Technology": 36.0, "Energy": 3.0},
        "contributors": [{"symbol": "IVV", "holding": "S&P 500 ETF", "return": .1337, "contribution": .0147}],
        "detractors": [{"symbol": "TLT", "holding": "20+ Year Treasury ETF", "return": -.0261, "contribution": -.0003}],
        "earnings_years": ["2025", "2026E"],
        "earnings_values": [300, 380],
        "earnings_notes": ["Estimates remain positive."],
        "optional_sections": {},
        "sources": {
            "allocation": "risk.csv row 2", "risk": "risk.csv row 3",
            "sector_performance": "market.pdf page 4", "sector_exposure": "risk.csv row 5",
            "attribution": "attribution.csv rows 2-8", "earnings": "market.pdf page 7",
        },
    }


class ClientDeckContentTests(unittest.TestCase):
    def test_normalizes_numeric_rows_and_formats_attribution(self) -> None:
        result = normalize_client_deck_payload(
            _payload(), client_name="Sample Household", period="August 2026", as_of="As of August 25, 2026"
        )

        self.assertEqual(result.contributors[0][2], "+13.37%")
        self.assertEqual(result.detractors[0][3], "-0.03%")
        self.assertEqual(result.client_name, "Sample Household")

    def test_rejects_mismatched_sector_labels(self) -> None:
        payload = _payload()
        payload["sector_benchmark"] = {"Technology": 36.0}

        with self.assertRaisesRegex(ClientDeckContentError, "identical"):
            normalize_client_deck_payload(
                payload, client_name="Sample", period="August", as_of="As of August 25, 2026"
            )


if __name__ == "__main__":
    unittest.main()
