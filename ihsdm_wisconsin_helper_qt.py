"""IHSDadaM Qt launcher -- drop-in replacement for ihsdm_wisconsin_helper.py.

Run this file directly to launch the PySide6 version of the application.
"""

import sys
import os

# Add src directory to path so ``ihsdadam`` package is importable
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from PySide6.QtWidgets import QApplication
from ihsdadam.app import IHSDadaMApp

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = IHSDadaMApp()
    window.show()
    sys.exit(app.exec())
