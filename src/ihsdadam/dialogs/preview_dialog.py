"""Preview dialog for showing pending changes"""

from typing import List

from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
)
from PySide6.QtCore import Qt

from .. import theme


class PreviewDialog(QDialog):
    """Shows a list of changes for user review before applying."""

    def __init__(self, title: str, changes_list: List[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(700, 500)
        self.setMinimumSize(400, 300)
        self.setWindowFlags(
            self.windowFlags() & ~Qt.WindowContextHelpButtonHint
        )

        self._build_ui(changes_list)

    # ── UI construction ────────────────────────────────────────────────

    def _build_ui(self, changes_list: List[str]) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        # Header
        count = len(changes_list)
        header_text = f"The following {count} change{'s' if count != 1 else ''} will be made:"
        header_label = QLabel(header_text)
        header_label.setStyleSheet(
            f"font-size: 11pt; font-weight: 600; color: {theme.TEXT_PRIMARY};"
        )
        layout.addWidget(header_label)

        # Changes text area
        changes_edit = QTextEdit()
        changes_edit.setReadOnly(True)
        changes_edit.setPlainText("\n".join(changes_list))
        changes_edit.setStyleSheet(
            f"background-color: {theme.SUBTLE_FILL};"
            f" border: 1px solid {theme.BORDER};"
            " border-radius: 4px; padding: 10px;"
            ' font-family: "Cascadia Code", "Consolas", monospace;'
            " font-size: 9pt;"
        )
        layout.addWidget(changes_edit, 1)

        # Close button
        close_btn = QPushButton("Close")
        close_btn.setMinimumWidth(100)
        close_btn.clicked.connect(self.accept)

        button_layout = QVBoxLayout()
        button_layout.setAlignment(Qt.AlignRight)
        button_layout.addWidget(close_btn, 0, Qt.AlignRight)
        layout.addLayout(button_layout)
