"""Regression checks for the GSWM-branded desktop visual system."""

import unittest

from ui.theme import APP_STYLESHEET, CANVAS, GOLD, NAVY, NAVY_DARK


class ThemeTests(unittest.TestCase):
    def test_theme_uses_client_deck_brand_palette(self) -> None:
        self.assertEqual(NAVY, "#1B2A4A")
        self.assertEqual(NAVY_DARK, "#0A1224")
        self.assertEqual(GOLD, "#BFA054")
        self.assertEqual(CANVAS, "#F6F4EF")

    def test_signature_brand_components_are_styled(self) -> None:
        for selector in (
            "QFrame#TopBar",
            "QFrame#BrandMark",
            "QFrame#HomeHero",
            "QFrame#ReportCard",
            "QFrame#UploadBox",
            "QPushButton#PrimaryButton",
        ):
            with self.subTest(selector=selector):
                self.assertIn(selector, APP_STYLESHEET)

    def test_header_uses_gswm_navy_instead_of_near_black(self) -> None:
        header = APP_STYLESHEET.split("QFrame#TopBar", 1)[1].split("}", 1)[0]

        self.assertIn(f"background: {NAVY};", header)
        self.assertNotIn(f"background: {NAVY_DARK};", header)


if __name__ == "__main__":
    unittest.main()
