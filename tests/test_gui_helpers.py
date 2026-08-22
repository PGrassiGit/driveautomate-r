import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication

from drive_inventory_gui import DriveAutomateWindow, default_inventory_path


class GuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_default_inventory_is_xlsx(self) -> None:
        self.assertEqual(default_inventory_path().suffix, ".xlsx")

    def test_window_uses_pyside_tabs_and_safe_download_default(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir, patch.dict(
            os.environ, {"APPDATA": tmpdir}
        ):
            QSettings.setDefaultFormat(QSettings.IniFormat)
            QSettings.setPath(QSettings.IniFormat, QSettings.UserScope, tmpdir)
            window = DriveAutomateWindow()
            try:
                self.assertEqual(window.tabs.count(), 2)
                self.assertEqual(window.tabs.tabText(0), "1. Inventário")
                self.assertEqual(window.tabs.tabText(1), "2. Downloads")
                self.assertTrue(window.selection_mode.currentData())
                self.assertFalse(window.cancel_button.isEnabled())
            finally:
                window.close()


if __name__ == "__main__":
    unittest.main()

