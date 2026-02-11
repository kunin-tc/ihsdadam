"""Entry point for IHSDadaM PySide6 application."""

import sys
import os

# Add src directory to path so ``ihsdadam`` package is importable
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from PySide6.QtWidgets import QApplication
from ihsdadam.app import IHSDadaMApp


def main():
    app = QApplication(sys.argv)
    window = IHSDadaMApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
