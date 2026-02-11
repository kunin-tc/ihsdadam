"""Main application window for IHSDadaM (PySide6)."""

from PySide6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QStackedWidget,
    QListWidget,
    QListWidgetItem,
    QLineEdit,
    QPushButton,
    QLabel,
    QFileDialog,
    QFrame,
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QIcon, QPainter, QPixmap

from . import theme
from .widgets import StatusBar
from .tabs import (
    WarningTab,
    CompilerTab,
    AppendixTab,
    VisualTab,
    CMFTab,
    EvalYearsTab,
    AADTTab,
    ReportTab,
    AboutTab,
)
from .dialogs import UpdateDialog, AboutDialog
from .workers import UpdateCheckWorker

try:
    from version import __version__, __app_name__, GITHUB_API_URL, GITHUB_RELEASES_URL
except ImportError:
    __version__ = "1.0.0"
    __app_name__ = "IHSDadaM"
    GITHUB_API_URL = None
    GITHUB_RELEASES_URL = None


def _build_app_icon() -> QIcon:
    """Generate a simple highway-shield app icon at runtime."""
    px = QPixmap(64, 64)
    px.fill(QColor(0, 0, 0, 0))
    p = QPainter(px)
    p.setRenderHint(QPainter.Antialiasing)

    # Blue shield background
    p.setBrush(QColor(theme.PRIMARY))
    p.setPen(Qt.NoPen)
    p.drawRoundedRect(4, 4, 56, 56, 14, 14)

    # White road lines
    p.setPen(Qt.NoPen)
    p.setBrush(QColor("#ffffff"))
    # Center dashed line
    for y in range(14, 52, 10):
        p.drawRect(30, y, 4, 6)
    # Left lane edge
    p.drawRect(16, 12, 3, 40)
    # Right lane edge
    p.drawRect(45, 12, 3, 40)

    p.end()

    icon = QIcon()
    icon.addPixmap(px)
    # Also add a 32px version for taskbar
    icon.addPixmap(px.scaled(32, 32, Qt.KeepAspectRatio, Qt.SmoothTransformation))
    return icon


class IHSDadaMApp(QMainWindow):
    """Main application window -- sidebar navigation with stacked content."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{__app_name__} v{__version__}")
        self.resize(1400, 850)
        self.setMinimumSize(800, 500)
        self.setStyleSheet(theme.STYLESHEET)
        self.setWindowIcon(_build_app_icon())

        self._project_path = ""
        self._update_worker = None
        self._tabs = []

        self._setup_ui()

        # Fire-and-forget update check after the event loop settles
        QTimer.singleShot(2000, self._check_for_updates)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # -- Header banner (full width) ------------------------------------
        outer.addWidget(self._build_header())

        # -- Project path bar (full width) ---------------------------------
        outer.addWidget(self._build_project_selector())

        # -- Body: sidebar + content ---------------------------------------
        body = QHBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)

        body.addWidget(self._build_sidebar())
        body.addWidget(self._build_content_stack(), stretch=1)

        outer.addLayout(body, stretch=1)

        # -- Status bar (full width) ---------------------------------------
        outer.addWidget(self._build_status_bar())

        # Wire sidebar selection to stacked widget
        self._sidebar.currentRowChanged.connect(self._stack.setCurrentIndex)
        # Select first tool
        self._sidebar.setCurrentRow(0)

    # -- Header banner -----------------------------------------------------

    def _build_header(self):
        header = QFrame()
        header.setStyleSheet(
            f"background-color: {theme.PRIMARY};"
            " border-radius: 0px;"
        )
        header.setFixedHeight(48)

        layout = QHBoxLayout(header)
        layout.setContentsMargins(16, 0, 16, 0)

        # Left side -- app name + clickable version
        left = QHBoxLayout()
        left.setSpacing(10)

        name_label = QLabel(__app_name__)
        name_label.setStyleSheet(
            "font-size: 16pt;"
            " font-weight: 700;"
            " color: #ffffff;"
            " background-color: transparent;"
        )
        left.addWidget(name_label)

        self._version_label = QLabel(f"v{__version__}")
        self._version_label.setStyleSheet(
            "font-size: 9pt;"
            " color: rgba(255,255,255,0.80);"
            " background-color: transparent;"
        )
        self._version_label.setCursor(Qt.PointingHandCursor)
        self._version_label.mousePressEvent = lambda _event: self._show_about()
        left.addWidget(self._version_label, alignment=Qt.AlignVCenter)

        layout.addLayout(left)
        layout.addStretch()

        # Right side -- group name
        subtitle = QLabel("HNTB WI Traffic Group")
        subtitle.setStyleSheet(
            "font-size: 9pt;"
            " font-style: italic;"
            " color: rgba(255,255,255,0.75);"
            " background-color: transparent;"
        )
        layout.addWidget(subtitle, alignment=Qt.AlignRight | Qt.AlignVCenter)

        return header

    # -- Project path selector ---------------------------------------------

    def _build_project_selector(self):
        bar = QFrame()
        bar.setStyleSheet(
            f"background-color: {theme.SURFACE};"
            f" border-bottom: 1px solid {theme.BORDER};"
        )

        layout = QHBoxLayout(bar)
        layout.setContentsMargins(16, 6, 16, 6)
        layout.setSpacing(8)

        lbl = QLabel("Project:")
        lbl.setStyleSheet(
            f"font-weight: 600; color: {theme.TEXT_SECONDARY};"
            " background-color: transparent;"
        )
        layout.addWidget(lbl)

        self._path_entry = QLineEdit()
        self._path_entry.setPlaceholderText(
            r"C:\Users\...\Projects_V5\p259  (paste path or click Browse)"
        )
        self._path_entry.setStyleSheet(
            f"border: 1px solid {theme.BORDER};"
            f" border-bottom: 1px solid {theme.BORDER};"
            " border-radius: 3px;"
            " padding: 4px 8px;"
            " background-color: #ffffff;"
        )
        self._path_entry.textChanged.connect(self._update_project_path)
        layout.addWidget(self._path_entry, stretch=1)

        browse_btn = QPushButton("Browse...")
        browse_btn.setStyleSheet(
            "background-color: transparent;"
            f" border: 1px solid {theme.BORDER_STRONG};"
            " border-radius: 3px;"
            " padding: 4px 14px;"
        )
        browse_btn.clicked.connect(self._browse_project)
        layout.addWidget(browse_btn)

        return bar

    # -- Sidebar navigation ------------------------------------------------

    def _build_sidebar(self):
        self._sidebar = QListWidget()
        self._sidebar.setObjectName("sidebar")
        self._sidebar.setFixedWidth(190)

        tool_names = [
            "Warning Extractor",
            "Data Compiler",
            "Appendix Generator",
            "Visual View",
            "Evaluation Info",
            "AADT Input",
            "Evaluation Years",
            "Report Generation",
            "About",
        ]

        for name in tool_names:
            item = QListWidgetItem(name)
            self._sidebar.addItem(item)

        return self._sidebar

    # -- Stacked content ---------------------------------------------------

    def _build_content_stack(self):
        self._stack = QStackedWidget()
        self._stack.setStyleSheet(
            f"QStackedWidget {{ background-color: {theme.BACKGROUND}; }}"
        )

        # Instantiate tabs
        self._warning_tab = WarningTab()
        self._compiler_tab = CompilerTab()
        self._appendix_tab = AppendixTab()
        self._visual_tab = VisualTab()
        self._cmf_tab = CMFTab()
        self._aadt_tab = AADTTab()
        self._eval_years_tab = EvalYearsTab()
        self._report_tab = ReportTab()
        self._about_tab = AboutTab()

        self._tabs = [
            self._warning_tab,
            self._compiler_tab,
            self._appendix_tab,
            self._visual_tab,
            self._cmf_tab,
            self._aadt_tab,
            self._eval_years_tab,
            self._report_tab,
            self._about_tab,
        ]

        for tab in self._tabs:
            self._stack.addWidget(tab)
            tab.status_message.connect(self._on_status_message)
            tab.progress_update.connect(self._on_progress_update)

        return self._stack

    # -- Status bar --------------------------------------------------------

    def _build_status_bar(self):
        self._status_bar = StatusBar()
        return self._status_bar

    # ------------------------------------------------------------------
    # Project path handling
    # ------------------------------------------------------------------

    def _browse_project(self):
        directory = QFileDialog.getExistingDirectory(
            self, "Select IHSDM Project Directory"
        )
        if directory:
            self._project_path = directory
            self._path_entry.setText(directory)

    def _update_project_path(self):
        """Propagate the current path text to every tab."""
        path = self._path_entry.text()
        self._project_path = path
        for tab in self._tabs:
            tab.set_project_path(path)
        if path:
            self._status_bar.set_message(f"Project: {path}")
        else:
            self._status_bar.set_message("Ready")

    # ------------------------------------------------------------------
    # About dialog
    # ------------------------------------------------------------------

    def _show_about(self):
        releases_url = GITHUB_RELEASES_URL or ""
        dialog = AboutDialog(__app_name__, __version__, releases_url, self)
        dialog.exec()

    # ------------------------------------------------------------------
    # Update checking
    # ------------------------------------------------------------------

    def _check_for_updates(self):
        if not GITHUB_API_URL:
            return
        self._update_worker = UpdateCheckWorker(
            GITHUB_API_URL, __version__, parent=self
        )
        self._update_worker.update_available.connect(self._on_update_available)
        self._update_worker.start()

    def _on_update_available(self, version, url, notes):
        dialog = UpdateDialog(__version__, version, url, notes, self)
        dialog.exec()

    # ------------------------------------------------------------------
    # Status bar helpers
    # ------------------------------------------------------------------

    def _on_status_message(self, text):
        self._status_bar.set_message(text)

    def _on_progress_update(self, percent, message):
        if percent >= 100:
            self._status_bar.hide_progress()
            self._status_bar.set_message(message)
        else:
            self._status_bar.show_progress(percent, message)
