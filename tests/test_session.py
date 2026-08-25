"""Tests for temporary session deletion boundaries."""

import tempfile
import unittest
from pathlib import Path

from core.session import ReportSession, purge_abandoned_sessions


class ReportSessionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_cleanup_removes_all_working_data(self) -> None:
        session = ReportSession.create(self.root)
        (session.uploads / "source.csv").write_text("private working data")
        (session.working / "draft.txt").write_text("temporary draft")
        path = session.path

        session.cleanup()

        self.assertFalse(path.exists())

    def test_context_manager_cleans_after_failure(self) -> None:
        path: Path | None = None
        with self.assertRaises(RuntimeError):
            with ReportSession.create(self.root) as session:
                path = session.path
                raise RuntimeError("synthetic generation failure")

        self.assertIsNotNone(path)
        self.assertFalse(path.exists())

    def test_startup_purge_only_removes_session_shaped_directories(self) -> None:
        abandoned = self.root / ("a" * 32)
        unrelated = self.root / "keep-me"
        abandoned.mkdir()
        unrelated.mkdir()

        removed = purge_abandoned_sessions(self.root)

        self.assertEqual(removed, 1)
        self.assertFalse(abandoned.exists())
        self.assertTrue(unrelated.exists())


if __name__ == "__main__":
    unittest.main()
