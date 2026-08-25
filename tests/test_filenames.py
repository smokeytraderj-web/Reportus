"""Tests for safe non-overwriting output paths."""

import tempfile
import unittest
from pathlib import Path

from tools.filenames import sanitize_filename, versioned_output_path


class FilenameTests(unittest.TestCase):
    def test_invalid_windows_characters_are_replaced(self) -> None:
        self.assertEqual(sanitize_filename('Client: Review?.pdf'), "Client_ Review_.pdf")

    def test_reserved_windows_name_is_prefixed(self) -> None:
        self.assertEqual(sanitize_filename("CON.pdf"), "_CON.pdf")

    def test_existing_file_gets_next_version(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "Review.pdf").touch()
            (root / "Review_v2.pdf").touch()

            result = versioned_output_path(root, "Review.pdf")

            self.assertEqual(result.name, "Review_v3.pdf")


if __name__ == "__main__":
    unittest.main()
