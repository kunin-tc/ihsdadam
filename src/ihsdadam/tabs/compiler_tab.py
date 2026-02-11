"""Data Compiler tab -- compiles IHSDM crash prediction data to Excel."""

import os
from typing import Dict, List, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTextEdit,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .. import theme
from ..widgets import ScrollableTree
from ..workers import CompileWorker, YearScanWorker


_INSTRUCTIONS = """\
IHSDM DATA COMPILER - Instructions

This tool extracts crash prediction data from IHSDM evaluation files \
and compiles them into Excel.

WHAT IT DOES:
  - Processes Highway Segments (h folders)
  - Processes Intersections (i folders)
  - Processes Ramp Terminals (r folders)
  - Processes Site Sets (ss folders) - Urban/Suburban arterial \
intersections & ramp terminals
  - Applies HSM severity distributions (K, A, B, C crashes)
  - Removes duplicates
  - Outputs to Excel workbook with separate sheets

FOLDER NAMING:
  h1, h2, h74, etc.   -> Highway alignments
  i1, i2, i100, etc.  -> Intersection alignments
  r1, r2, r25, etc.   -> Ramp terminal alignments
  ss1, ss2, etc.       -> Site sets (intersections & ramp terminals)
  c1, c2, etc.         -> Interchanges (can contain h/i/r subfolders)

ALIGNMENT NAMING BEST PRACTICES:
  - Mainline freeways: Include "Mainline" in alignment name
  - Ramps: Include "Entrance" or "Exit" in alignment name
  - Ensure single functional class per alignment

INTERSECTIONS & RAMP TERMINALS IN HIGHWAY FOLDERS:
  IHSDM allows running intersection and ramp terminal evaluations
  inside highway alignments (h folders). The compiler automatically
  detects these and routes them to the correct Intersection and
  RampTerminal sheets. Duplicates are removed by title.

OUTPUT:
  Excel file with up to 5 sheets:
    Highway, Intersection, RampTerminal, SiteSet_Int, SiteSet_Ramp
  - Site set data goes to separate SiteSet_Int and SiteSet_Ramp sheets
  - Includes crash predictions with severity breakdown
  - Extracts data for selected evaluation year(s) only

HSM SEVERITY DISTRIBUTIONS:
  Highway:          K=1.46%, A=4.48%, B=24.69%, C=69.17%
  Intersection/Ramp: K=0.26%, A=5.35%, B=27.64%, C=66.75%
  NOTE: These default distributions are ONLY applied when a \
severity distribution was not available through calibration factors.

STEPS:
  1. Select project directory above (shared with Warning Extractor)
  2. Set Excel output file path
  3. Click "Scan for Years" to find available evaluation years
  4. Select single year OR year range from dropdowns
  5. Click "Compile Data to Excel"
  6. Review output in Excel file

NOTE: Consult your state's predictive modeling practices \
for calibration factors.
"""

_SCAN_TREE_COLUMNS = ["Name", "Folder", "Available Years"]

_PREFIX_LABELS = {
    "h": "Highways",
    "i": "Intersections",
    "ra": "Intersections",
    "r": "Ramp Terminals",
    "ss": "Site Sets",
}


class CompilerTab(QWidget):
    """Data Compiler tab -- scan years and compile crash prediction data."""

    status_message = Signal(str)
    progress_update = Signal(int, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._project_path = ""
        self._available_years: List[str] = []
        self._alignment_years: Dict[tuple, List[str]] = {}
        self._year_worker: Optional[YearScanWorker] = None
        self._compile_worker: Optional[CompileWorker] = None
        self._setup_ui()

    # ── public API ───────────────────────────────────────────────────────

    def set_project_path(self, path: str):
        self._project_path = path

    # ── UI construction ──────────────────────────────────────────────────

    def _setup_ui(self):
        root_layout = QHBoxLayout(self)
        root_layout.setContentsMargins(10, 10, 10, 10)
        root_layout.setSpacing(10)

        # ── Left column ──────────────────────────────────────────────────
        left = QVBoxLayout()
        left.setSpacing(10)

        # Configuration group
        config_box = QGroupBox("Compiler Configuration")
        config_layout = QVBoxLayout(config_box)
        config_layout.setSpacing(8)

        # Target CSV
        config_layout.addWidget(QLabel("Target CSV File:"))
        target_row = QHBoxLayout()
        self._target_input = QLineEdit("evaluation.1.diag.csv")
        target_row.addWidget(self._target_input)
        hint = QLabel("(default)")
        hint.setProperty("caption", True)
        target_row.addWidget(hint)
        config_layout.addLayout(target_row)

        # Excel Output
        config_layout.addWidget(QLabel("Excel Output:"))
        excel_row = QHBoxLayout()
        self._excel_input = QLineEdit()
        self._excel_input.setPlaceholderText("Select output .xlsx file...")
        excel_row.addWidget(self._excel_input)
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self._browse_excel_output)
        excel_row.addWidget(browse_btn)
        config_layout.addLayout(excel_row)

        # Year selection group
        year_box = QGroupBox("Evaluation Years")
        year_layout = QVBoxLayout(year_box)
        year_layout.setSpacing(6)

        # Scan row
        scan_row = QHBoxLayout()
        scan_years_btn = QPushButton("Scan for Years")
        scan_years_btn.clicked.connect(self._scan_years)
        scan_row.addWidget(scan_years_btn)
        scan_row.addWidget(QLabel("Available:"))
        self._available_label = QLabel("--")
        self._available_label.setStyleSheet(f"color: {theme.PRIMARY};")
        scan_row.addWidget(self._available_label, stretch=1)
        year_layout.addLayout(scan_row)

        # Single year radio + combo
        single_row = QHBoxLayout()
        self._single_radio = QRadioButton("Single Year:")
        self._single_radio.setChecked(True)
        self._single_radio.toggled.connect(self._update_year_dropdowns)
        single_row.addWidget(self._single_radio)
        self._single_combo = QComboBox()
        self._single_combo.setFixedWidth(90)
        single_row.addWidget(self._single_combo)
        single_row.addStretch()
        year_layout.addLayout(single_row)

        # Year range radio + combos
        range_row = QHBoxLayout()
        self._range_radio = QRadioButton("Year Range:")
        self._range_radio.toggled.connect(self._update_year_dropdowns)
        range_row.addWidget(self._range_radio)
        self._start_combo = QComboBox()
        self._start_combo.setFixedWidth(90)
        self._start_combo.setEnabled(False)
        range_row.addWidget(self._start_combo)
        range_row.addWidget(QLabel("to"))
        self._end_combo = QComboBox()
        self._end_combo.setFixedWidth(90)
        self._end_combo.setEnabled(False)
        range_row.addWidget(self._end_combo)
        range_row.addStretch()
        year_layout.addLayout(range_row)

        config_layout.addWidget(year_box)

        # Debug checkbox
        self._debug_check = QCheckBox("Debug mode (verbose output)")
        config_layout.addWidget(self._debug_check)

        left.addWidget(config_box)

        # Compile button
        self._compile_btn = QPushButton("Compile Data to Excel")
        self._compile_btn.setProperty("accent", True)
        self._compile_btn.clicked.connect(self._run_compiler)
        left.addWidget(self._compile_btn)

        # Instructions panel
        instructions_box = QGroupBox("Instructions & Information")
        instructions_layout = QVBoxLayout(instructions_box)
        self._instructions_text = QTextEdit()
        self._instructions_text.setReadOnly(True)
        self._instructions_text.setFont(QFont("Cascadia Code", 9))
        self._instructions_text.setPlainText(_INSTRUCTIONS)
        instructions_layout.addWidget(self._instructions_text)
        left.addWidget(instructions_box, stretch=1)

        root_layout.addLayout(left, stretch=1)

        # ── Right column ─────────────────────────────────────────────────
        right = QVBoxLayout()
        right.setSpacing(10)

        scan_box = QGroupBox("Scan Results - Alignments & Available Years")
        scan_layout = QVBoxLayout(scan_box)

        self._scan_tree = ScrollableTree(_SCAN_TREE_COLUMNS)
        self._scan_tree.set_column_widths([240, 70, 180])
        scan_layout.addWidget(self._scan_tree)

        self._scan_status = QLabel("Click 'Scan for Years' to analyze evaluation files")
        self._scan_status.setProperty("caption", True)
        scan_layout.addWidget(self._scan_status)

        right.addWidget(scan_box, stretch=1)
        root_layout.addLayout(right, stretch=1)

        # Ensure correct initial state
        self._update_year_dropdowns()

    # ── Year dropdown state ──────────────────────────────────────────────

    def _update_year_dropdowns(self):
        is_single = self._single_radio.isChecked()
        self._single_combo.setEnabled(is_single)
        self._start_combo.setEnabled(not is_single)
        self._end_combo.setEnabled(not is_single)

    # ── Browse ───────────────────────────────────────────────────────────

    def _browse_excel_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Select Excel Output",
            "compiled_output.xlsx",
            "Excel Files (*.xlsx);;All Files (*)",
        )
        if path:
            self._excel_input.setText(path)

    # ── Year scanning ────────────────────────────────────────────────────

    def _scan_years(self):
        if not self._project_path or not os.path.isdir(self._project_path):
            QMessageBox.critical(self, "Error", "Please select a valid project directory first.")
            return

        self.status_message.emit("Scanning for available years...")

        self._year_worker = YearScanWorker(self._project_path, parent=self)
        self._year_worker.progress.connect(self._on_year_progress)
        self._year_worker.finished.connect(self._on_years_found)
        self._year_worker.error.connect(self._on_year_error)
        self._year_worker.start()

    def _on_year_progress(self, pct: int, msg: str):
        self.progress_update.emit(pct, msg)
        self.status_message.emit(msg)

    def _on_years_found(self, data: dict):
        all_years: List[str] = data.get("all_years", [])
        alignment_years: Dict[tuple, list] = data.get("alignment_years", {})

        self._available_years = list(all_years)
        self._alignment_years = dict(alignment_years)

        # Update available label
        if all_years:
            self._available_label.setText(", ".join(all_years))
        else:
            self._available_label.setText("No years found")

        # Populate year combos
        self._single_combo.clear()
        self._start_combo.clear()
        self._end_combo.clear()

        for y in all_years:
            self._single_combo.addItem(y)
            self._start_combo.addItem(y)
            self._end_combo.addItem(y)

        if all_years:
            self._single_combo.setCurrentIndex(0)
            self._start_combo.setCurrentIndex(0)
            self._end_combo.setCurrentIndex(self._end_combo.count() - 1)

        # Populate scan results tree grouped by type
        self._scan_tree.clear()

        if alignment_years:
            sorted_items = sorted(
                alignment_years.items(),
                key=lambda item: _alignment_sort_key(item[0]),
            )

            # Group by alignment type
            groups: Dict[str, list] = {}
            for (folder_id, alignment_name), years in sorted_items:
                prefix = "".join(c for c in folder_id if c.isalpha())
                group_label = _PREFIX_LABELS.get(prefix, f"Other ({prefix})")
                groups.setdefault(group_label, []).append(
                    (folder_id, alignment_name, years)
                )

            bold_font = QFont()
            bold_font.setBold(True)

            for group_label, entries in groups.items():
                parent_item = QTreeWidgetItem(
                    [f"{group_label} ({len(entries)})", "", ""]
                )
                parent_item.setFont(0, bold_font)
                from PySide6.QtGui import QColor
                parent_item.setForeground(0, QColor(theme.PRIMARY))
                self._scan_tree.tree.addTopLevelItem(parent_item)

                for folder_id, alignment_name, years in entries:
                    sorted_years = sorted(years)
                    years_str = _format_years_display(sorted_years)
                    child = QTreeWidgetItem(
                        [alignment_name, folder_id, years_str]
                    )
                    parent_item.addChild(child)

                parent_item.setExpanded(True)

            h_count = len(groups.get("Highways", []))
            i_count = len(groups.get("Intersections", []))
            r_count = len(groups.get("Ramp Terminals", []))
            ss_count = len(groups.get("Site Sets", []))
            self._scan_status.setText(
                f"H: {h_count}  |  I: {i_count}  |  "
                f"R: {r_count}  |  SS: {ss_count}  |  "
                f"{len(all_years)} unique years"
            )
            self._scan_status.setStyleSheet(f"color: {theme.ACCENT_GREEN};")
        else:
            self._scan_status.setText("No year data found in files")
            self._scan_status.setStyleSheet(f"color: {theme.DANGER};")

        if all_years:
            self.status_message.emit(
                f"Found {len(all_years)} years: {all_years[0]}-{all_years[-1]}"
            )
        else:
            self.status_message.emit("No year data found in evaluation files")

    def _on_year_error(self, msg: str):
        QMessageBox.critical(self, "Scan Error", msg)
        self._scan_status.setText(f"Scan error: {msg}")
        self._scan_status.setStyleSheet(f"color: {theme.DANGER};")
        self.status_message.emit("Error scanning for years")

    # ── Compilation ──────────────────────────────────────────────────────

    def _run_compiler(self):
        if not self._project_path or not os.path.isdir(self._project_path):
            QMessageBox.critical(self, "Error", "Please select a valid project directory.")
            return

        excel_path = self._excel_input.text().strip()
        if not excel_path:
            QMessageBox.critical(self, "Error", "Please select an Excel output file.")
            return

        target_file = self._target_input.text().strip()
        if not target_file:
            target_file = "evaluation.1.diag.csv"

        # Build target years list
        target_years = self._build_target_years()
        if target_years is None:
            return

        debug = self._debug_check.isChecked()

        self._compile_btn.setEnabled(False)
        self.status_message.emit("Compiling data to Excel...")

        self._compile_worker = CompileWorker(
            project_path=self._project_path,
            excel_path=excel_path,
            target_file=target_file,
            target_years=target_years,
            debug=debug,
            parent=self,
        )
        self._compile_worker.progress.connect(self._on_compile_progress)
        self._compile_worker.finished.connect(self._on_compile_finished)
        self._compile_worker.error.connect(self._on_compile_error)
        self._compile_worker.start()

    def _build_target_years(self) -> Optional[List[str]]:
        if not self._available_years:
            QMessageBox.warning(
                self,
                "No Years",
                "Please scan for available years first.",
            )
            return None

        if self._single_radio.isChecked():
            year = self._single_combo.currentText()
            if not year:
                QMessageBox.warning(self, "No Year", "Please select a year.")
                return None
            return [year]

        start = self._start_combo.currentText()
        end = self._end_combo.currentText()
        if not start or not end:
            QMessageBox.warning(self, "No Range", "Please select start and end years.")
            return None

        try:
            start_int = int(start)
            end_int = int(end)
        except ValueError:
            QMessageBox.warning(self, "Invalid Years", "Year values must be numeric.")
            return None

        if start_int > end_int:
            QMessageBox.warning(
                self,
                "Invalid Range",
                "Start year must be less than or equal to end year.",
            )
            return None

        return [str(y) for y in range(start_int, end_int + 1)]

    def _on_compile_progress(self, pct: int, msg: str):
        self.progress_update.emit(pct, msg)
        self.status_message.emit(msg)

    def _on_compile_finished(self, summary: str):
        self._compile_btn.setEnabled(True)
        QMessageBox.information(
            self,
            "Compilation Complete",
            f"Data compiled successfully.\n\n{summary}",
        )
        self.status_message.emit("Compilation complete")

    def _on_compile_error(self, msg: str):
        self._compile_btn.setEnabled(True)
        QMessageBox.critical(self, "Compilation Error", msg)
        self.status_message.emit("Compilation failed")


# ── Helpers ──────────────────────────────────────────────────────────────────


def _alignment_sort_key(key: tuple) -> tuple:
    """Sort alignments by type prefix then numeric id."""
    folder_id = key[0]
    prefix = "".join(c for c in folder_id if c.isalpha())
    num_str = "".join(c for c in folder_id if c.isdigit())
    num = int(num_str) if num_str else 0
    prefix_order = {"h": 0, "i": 1, "ra": 1, "r": 2, "ss": 3}.get(prefix, 4)
    return (prefix_order, num)


def _format_years_display(sorted_years: List[str]) -> str:
    """Format a sorted list of year strings for display."""
    if not sorted_years:
        return ""
    if len(sorted_years) == 1:
        return sorted_years[0]
    try:
        first = int(sorted_years[0])
        last = int(sorted_years[-1])
        if len(sorted_years) == last - first + 1:
            return f"{sorted_years[0]}-{sorted_years[-1]}"
    except ValueError:
        pass
    return ", ".join(sorted_years)
