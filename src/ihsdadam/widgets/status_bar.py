"""Custom status bar with embedded progress indicator"""

from PySide6.QtWidgets import QWidget, QHBoxLayout, QLabel, QProgressBar
from PySide6.QtCore import Qt


class StatusBar(QWidget):
    """Status bar with message text, credits, and progress bar"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(28)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 0, 16, 0)
        layout.setSpacing(12)

        self._message = QLabel("Ready")
        self._message.setStyleSheet("color: #616161; font-size: 9pt;")
        layout.addWidget(self._message)

        self._progress = QProgressBar()
        self._progress.setFixedWidth(200)
        self._progress.setFixedHeight(4)
        self._progress.setTextVisible(False)
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        self._progress.hide()
        layout.addWidget(self._progress)

        layout.addStretch()

        self._credits = QLabel(
            "Adam Engbring \u2022 aengbring@hntb.com"
        )
        self._credits.setStyleSheet("color: #b0b0b0; font-size: 8pt;")
        layout.addWidget(self._credits)

        self.setStyleSheet("background-color: #fafafa; border-top: 1px solid #e5e5e5;")

    def set_message(self, text):
        self._message.setText(text)

    def show_progress(self, value=None, message=None):
        """Show progress bar. value=None for indeterminate."""
        if message:
            self._message.setText(message)
        if value is None:
            self._progress.setRange(0, 0)
        else:
            self._progress.setRange(0, 100)
            self._progress.setValue(value)
        self._progress.show()

    def hide_progress(self):
        self._progress.hide()
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
