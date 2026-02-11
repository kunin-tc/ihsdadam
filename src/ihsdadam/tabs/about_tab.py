"""About tab -- help info and credits."""

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QLabel,
    QScrollArea,
    QFrame,
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont

from .. import theme

try:
    from version import __version__, __app_name__
except ImportError:
    __version__ = "1.0.0"
    __app_name__ = "IHSDadaM"


_HELP_TEXT = """\
IHSDadaM (IHSDM Data Manager) pulls results out of IHSDM projects \
and gets them into formats you can actually use -- Excel, PDF, etc. \
It's built around the typical workflow of running IHSDM evaluations \
and then needing to document the results for a report.

The tabs are roughly in the order you'd use them:

WARNING EXTRACTOR
  Scans your project's result XMLs and pulls out every warning, \
error, and info message. The main thing to watch for is CRITICAL \
messages -- those mean "no crash prediction supported" for a segment, \
which usually means something is wrong with your geometry inputs. \
Fix those before compiling data.

DATA COMPILER
  Reads the diagnostic CSVs from each evaluation and compiles crash \
predictions into an Excel workbook. Highways, intersections, ramp \
terminals, and site sets each get their own sheet. It applies HSM \
severity distributions (K/A/B/C) and handles deduplication. You \
pick which year(s) to extract.

APPENDIX GENERATOR
  Finds all the evaluation report PDFs in your project and merges \
them into one file. Useful for building appendices.

VISUAL VIEW
  Loads a highway XML and draws the alignment -- lanes, shoulders, \
curves, speed zones, AADT, medians, intersections. Mostly useful \
for sanity-checking your inputs before running evaluations.

EVALUATION INFO
  Pulls calibration factors from the CMF CSVs so you can document \
what was used. Shows alignment name, type, eval years, and the \
calibration value.

AADT INPUT
  Wizard for when IHSDM has forecast ID placeholders instead of \
actual AADT values. Maps the forecast IDs to traffic volumes.

EVALUATION YEARS
  Quick view of what year ranges each alignment covers. Good for \
making sure everything lines up before compiling.

REPORT GENERATION
  Builds printable HTML crash prediction reports from Data Compiler \
Excel output. Supports single-project summaries and multi-project \
comparisons. Highway segments are grouped by functional class type \
(e.g., Freeways, Ramps, Arterials). Reports include KABCO or FI/PDO \
severity breakdowns, bar charts, and summary tables. Output opens \
directly in your browser for printing or saving as PDF.
"""


class AboutTab(QWidget):
    """About tab with help text and credits."""

    status_message = Signal(str)
    progress_update = Signal(int, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._project_path = ""
        self._setup_ui()

    def set_project_path(self, path: str):
        self._project_path = path

    def _setup_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet(
            f"QScrollArea {{ background-color: {theme.BACKGROUND}; border: none; }}"
        )

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(32, 24, 32, 32)
        layout.setSpacing(16)

        # -- Title --
        title = QLabel(f"{__app_name__}  v{__version__}")
        title.setStyleSheet(
            f"font-size: 18pt; font-weight: 700; color: {theme.TEXT_PRIMARY};"
        )
        layout.addWidget(title)

        subtitle = QLabel("IHSDM Data Manager")
        subtitle.setStyleSheet(
            f"font-size: 11pt; color: {theme.TEXT_SECONDARY};"
        )
        layout.addWidget(subtitle)

        # -- Help text --
        help_label = QLabel(_HELP_TEXT.strip())
        help_label.setWordWrap(True)
        help_label.setTextFormat(Qt.PlainText)
        help_label.setFont(QFont("Cascadia Code", 9))
        help_label.setStyleSheet(
            f"color: {theme.TEXT_PRIMARY}; line-height: 1.5;"
            f" background-color: {theme.SURFACE};"
            f" border: 1px solid {theme.BORDER};"
            " border-radius: 6px;"
            " padding: 16px;"
        )
        layout.addWidget(help_label)

        # -- Credits --
        credits_text = (
            "HNTB WI Traffic Group\n"
            "HNTB Corporation, Wisconsin Office"
        )
        credits_label = QLabel(credits_text)
        credits_label.setStyleSheet(
            f"font-size: 10pt; color: {theme.TEXT_SECONDARY};"
            " padding-top: 8px;"
        )
        layout.addWidget(credits_label)

        layout.addStretch()
        scroll.setWidget(container)
        outer.addWidget(scroll)
