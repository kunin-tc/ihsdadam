"""Reusable search/filter bar widget"""

from PySide6.QtWidgets import QWidget, QHBoxLayout, QLineEdit, QPushButton
from PySide6.QtCore import Signal, Qt


class SearchBar(QWidget):
    """Search bar with text input and clear button"""

    text_changed = Signal(str)
    cleared = Signal()

    def __init__(self, placeholder="Search...", parent=None):
        super().__init__(parent)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        self._input = QLineEdit()
        self._input.setPlaceholderText(placeholder)
        self._input.setClearButtonEnabled(True)
        self._input.textChanged.connect(self.text_changed.emit)
        layout.addWidget(self._input)

        self._clear_btn = QPushButton("Clear")
        self._clear_btn.setFixedWidth(60)
        self._clear_btn.clicked.connect(self._on_clear)
        layout.addWidget(self._clear_btn)

    def _on_clear(self):
        self._input.clear()
        self.cleared.emit()

    def text(self):
        return self._input.text()

    def set_text(self, text):
        self._input.setText(text)

    def clear(self):
        self._input.clear()
