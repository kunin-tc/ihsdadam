"""About/version information dialog"""

import webbrowser

from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
)
from PySide6.QtCore import Qt

from .. import theme


class AboutDialog(QDialog):
    """Shows current version and link to check for updates."""

    def __init__(self, app_name: str, version: str, releases_url: str, parent=None):
        super().__init__(parent)
        self._releases_url = releases_url

        self.setWindowTitle("Version Info")
        self.setFixedSize(420, 250)
        self.setWindowFlags(
            self.windowFlags() & ~Qt.WindowContextHelpButtonHint
        )

        self._build_ui(app_name, version)
        self._center_on_screen()

    # ── UI construction ────────────────────────────────────────────────

    def _build_ui(self, app_name: str, version: str) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 28, 32, 24)
        layout.setSpacing(8)

        # App name
        name_label = QLabel(app_name)
        name_label.setAlignment(Qt.AlignCenter)
        name_label.setStyleSheet(
            f"font-size: 22px; font-weight: 700; color: {theme.PRIMARY};"
        )
        layout.addWidget(name_label)

        # Version
        version_label = QLabel(f"Version {version}")
        version_label.setAlignment(Qt.AlignCenter)
        version_label.setStyleSheet(
            f"font-size: 12pt; color: {theme.TEXT_SECONDARY};"
        )
        layout.addWidget(version_label)

        layout.addSpacing(12)

        # Informational message
        message_label = QLabel(
            "Visit the GitHub releases page to check for newer\n"
            "versions, view changelogs, and download updates."
        )
        message_label.setAlignment(Qt.AlignCenter)
        message_label.setWordWrap(True)
        message_label.setStyleSheet(
            f"font-size: 10pt; color: {theme.TEXT_PRIMARY}; line-height: 1.4;"
        )
        layout.addWidget(message_label)

        layout.addStretch()

        # Buttons
        button_row = QHBoxLayout()
        button_row.setSpacing(12)

        releases_btn = QPushButton("View Releases on GitHub")
        releases_btn.setProperty("accent", True)
        releases_btn.setCursor(Qt.PointingHandCursor)
        releases_btn.setMinimumWidth(180)
        releases_btn.clicked.connect(self._open_releases)

        close_btn = QPushButton("Close")
        close_btn.setMinimumWidth(90)
        close_btn.clicked.connect(self.accept)

        button_row.addStretch()
        button_row.addWidget(releases_btn)
        button_row.addWidget(close_btn)
        button_row.addStretch()

        layout.addLayout(button_row)

    # ── Helpers ────────────────────────────────────────────────────────

    def _open_releases(self) -> None:
        webbrowser.open(self._releases_url)

    def _center_on_screen(self) -> None:
        screen_geo = self.screen().availableGeometry()
        dialog_geo = self.frameGeometry()
        dialog_geo.moveCenter(screen_geo.center())
        self.move(dialog_geo.topLeft())
