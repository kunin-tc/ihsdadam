"""Update available notification dialog"""

import webbrowser

from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
    QFrame,
)
from PySide6.QtCore import Qt

from .. import theme


class UpdateDialog(QDialog):
    """Dialog shown when a newer version is found on GitHub."""

    def __init__(
        self,
        current_version: str,
        new_version: str,
        download_url: str,
        release_notes: str,
        parent=None,
    ):
        super().__init__(parent)
        self._download_url = download_url

        self.setWindowTitle("Update Available")
        self.setFixedSize(500, 400)
        self.setWindowFlags(
            self.windowFlags() & ~Qt.WindowContextHelpButtonHint
        )

        self._build_ui(current_version, new_version, release_notes)

    # ── UI construction ────────────────────────────────────────────────

    def _build_ui(
        self, current_version: str, new_version: str, release_notes: str
    ) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Header banner
        header = QFrame()
        header.setStyleSheet(
            f"background-color: {theme.PRIMARY}; border: none;"
        )
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(24, 16, 24, 16)

        header_label = QLabel("New Version Available!")
        header_label.setAlignment(Qt.AlignCenter)
        header_label.setStyleSheet(
            "font-size: 16px; font-weight: 700; color: #ffffff;"
            " background-color: transparent;"
        )
        header_layout.addWidget(header_label)
        layout.addWidget(header)

        # Content area
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(24, 16, 24, 20)
        content_layout.setSpacing(8)

        # Current version
        current_label = QLabel(f"Installed version:  {current_version}")
        current_label.setStyleSheet(
            f"font-size: 10pt; color: {theme.TEXT_SECONDARY};"
        )
        content_layout.addWidget(current_label)

        # New version
        new_label = QLabel(f"Latest version:  {new_version}")
        new_label.setStyleSheet(
            f"font-size: 10pt; font-weight: 700; color: {theme.PRIMARY};"
        )
        content_layout.addWidget(new_label)

        content_layout.addSpacing(8)

        # What's New section
        notes_header = QLabel("What's New:")
        notes_header.setStyleSheet(
            f"font-size: 10pt; font-weight: 600; color: {theme.TEXT_PRIMARY};"
        )
        content_layout.addWidget(notes_header)

        notes_edit = QTextEdit()
        notes_edit.setReadOnly(True)
        notes_edit.setPlainText(release_notes)
        notes_edit.setStyleSheet(
            f"background-color: {theme.SUBTLE_FILL};"
            f" border: 1px solid {theme.BORDER};"
            " border-radius: 4px; padding: 8px;"
            ' font-family: "Segoe UI", Arial, sans-serif;'
            " font-size: 9pt;"
        )
        content_layout.addWidget(notes_edit, 1)

        content_layout.addSpacing(4)

        # Buttons
        button_row = QHBoxLayout()
        button_row.setSpacing(12)

        download_btn = QPushButton("Download Update")
        download_btn.setProperty("accent", True)
        download_btn.setCursor(Qt.PointingHandCursor)
        download_btn.setMinimumWidth(150)
        download_btn.clicked.connect(self._open_download)

        later_btn = QPushButton("Remind Me Later")
        later_btn.setMinimumWidth(130)
        later_btn.clicked.connect(self.reject)

        button_row.addStretch()
        button_row.addWidget(download_btn)
        button_row.addWidget(later_btn)

        content_layout.addLayout(button_row)
        layout.addLayout(content_layout)

    # ── Helpers ────────────────────────────────────────────────────────

    def _open_download(self) -> None:
        webbrowser.open(self._download_url)
        self.accept()
