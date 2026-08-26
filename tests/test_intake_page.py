"""Regression tests for switching among report intake forms."""

import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

try:
    from PySide6.QtWidgets import QApplication, QLineEdit

    from core.workflows import load_workflows
    from ui.pages import IntakePage
    from ui.widgets import UploadBox
except ImportError:  # pragma: no cover - CI may run the non-GUI dependency set
    QApplication = QLineEdit = IntakePage = UploadBox = None
    load_workflows = None


@unittest.skipIf(QApplication is None, "PySide6 is unavailable")
class IntakePageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.application = QApplication.instance() or QApplication([])

    def test_switching_all_report_types_removes_previous_controls(self) -> None:
        page = IntakePage()
        workflows = load_workflows()

        for workflow in (*workflows, *reversed(workflows)):
            with self.subTest(workflow=workflow.skill_id):
                page.set_workflow(workflow)
                self.application.processEvents()

                self.assertEqual(set(page.field_inputs), {field.field_id for field in workflow.fields})
                self.assertEqual(len(page.content.findChildren(QLineEdit)), len(workflow.fields))
                expected_uploads = (
                    len(workflow.required_uploads) + len(workflow.optional_uploads) + 1
                )
                self.assertEqual(len(page.content.findChildren(UploadBox)), expected_uploads)

        page.deleteLater()
        self.application.processEvents()


if __name__ == "__main__":
    unittest.main()
