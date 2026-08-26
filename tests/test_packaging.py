"""Static checks for the reproducible Windows packaging contract."""

import re
import unittest
from pathlib import Path

from config.version import __version__


ROOT = Path(__file__).resolve().parents[1]


class PackagingTests(unittest.TestCase):
    def test_installer_version_matches_runtime_version(self) -> None:
        installer = (ROOT / "packaging/windows/Reportus.iss").read_text(
            encoding="utf-8"
        )
        match = re.search(r'#define MyAppVersion "([^"]+)"', installer)

        self.assertIsNotNone(match)
        self.assertEqual(match.group(1), __version__)

    def test_spec_bundles_every_runtime_resource_root(self) -> None:
        spec = (ROOT / "packaging/windows/Reportus.spec").read_text(encoding="utf-8")

        for resource in ("skills", "report_workflows.json", "skills.json", "services", "templates"):
            with self.subTest(resource=resource):
                self.assertIn(resource, spec)

    def test_build_script_runs_tests_and_packaged_smoke_test(self) -> None:
        script = (ROOT / "scripts/build_windows.ps1").read_text(encoding="utf-8")

        self.assertIn('"unittest", "discover"', script)
        self.assertIn("--smoke-test", script)
        self.assertIn("Get-FileHash", script)

    def test_github_workflow_produces_installer_artifact(self) -> None:
        workflow = (ROOT / ".github/workflows/windows-installer.yml").read_text(
            encoding="utf-8"
        )

        self.assertIn("windows-latest", workflow)
        self.assertIn("build_windows.ps1", workflow)
        self.assertIn("Reportus-Windows-Installer", workflow)


if __name__ == "__main__":
    unittest.main()
