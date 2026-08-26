"""Tests for local-only portal screenshot preparation."""

import tempfile
import unittest
from pathlib import Path

from PIL import Image, ImageDraw

from security.portal_capture import PortalCaptureError, prepare_client_deck_portal_captures


class PortalCaptureTests(unittest.TestCase):
    def test_riskalyze_image_is_replaced_by_right_panel_crop(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "full-page.png"
            image = Image.new("RGB", (1800, 900), "navy")
            ImageDraw.Draw(image).rectangle((1188, 0, 1799, 899), fill=(245, 245, 245))
            image.save(source)

            prepared = prepare_client_deck_portal_captures({"risk_snapshot": (source,)})
            try:
                safe_path = prepared.selections["risk_snapshot"][0]
                self.assertNotEqual(safe_path, source)
                self.assertEqual(safe_path.name, "riskalyze_analytics_1.png")
                with Image.open(safe_path) as safe:
                    self.assertEqual(safe.size, (1224, 1800))
                    self.assertEqual(safe.getpixel((safe.width // 2, safe.height // 2)), (245, 245, 245))
            finally:
                prepared.cleanup()
            self.assertFalse(safe_path.exists())

    def test_small_non_widescreen_capture_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "small.png"
            Image.new("RGB", (640, 480), "white").save(source)

            with self.assertRaises(PortalCaptureError):
                prepare_client_deck_portal_captures({"risk_snapshot": (source,)})

    def test_already_safe_analytics_capture_is_not_cropped_again(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "riskalyze_analytics_verified.png"
            Image.new("RGB", (600, 1000), "white").save(source)

            prepared = prepare_client_deck_portal_captures({"risk_snapshot": (source,)})

            self.assertEqual(prepared.selections["risk_snapshot"], (source,))
            self.assertIsNone(prepared.temporary_directory)


if __name__ == "__main__":
    unittest.main()
