"""Evaluation Years overview tab -- shows year ranges per alignment."""

import os
import re
from typing import Dict, List, Tuple

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTreeWidgetItem, QMessageBox,
)
from PySide6.QtCore import Qt, Signal

from ..workers import YearScanWorker
from ..widgets import ScrollableTree
from .. import theme


# Prefix-to-sort-priority mapping
_PREFIX_ORDER = {"h": 0, "i": 1, "ra": 1, "r": 2, "ss": 3}

# Regex that splits a folder ID into (alpha prefix, numeric suffix)
_FOLDER_RE = re.compile(r"^([a-z]+)(\d+)$")


def _sort_key(folder_id: str) -> Tuple[int, int]:
    """Return (prefix_priority, numeric_id) for consistent ordering."""
    m = _FOLDER_RE.match(folder_id)
    if m:
        prefix, num = m.group(1), int(m.group(2))
        return (_PREFIX_ORDER.get(prefix, 4), num)
    # Fallback: extract whatever letters and digits we can
    prefix = "".join(c for c in folder_id if c.isalpha())
    num_str = "".join(c for c in folder_id if c.isdigit())
    num = int(num_str) if num_str else 0
    return (_PREFIX_ORDER.get(prefix, 4), num)


def _format_years(years: List[str]) -> str:
    """Format a sorted list of year strings as a compact range or CSV.

    Consecutive runs collapse to "YYYY-YYYY"; gaps produce comma-separated
    values.
    """
    if not years:
        return ""
    if len(years) == 1:
        return years[0]

    int_years = [int(y) for y in years]
    # Check if the sequence is fully consecutive
    if int_years[-1] - int_years[0] + 1 == len(int_years):
        return f"{years[0]}-{years[-1]}"

    # Build mixed range/single representation
    parts: List[str] = []
    run_start = int_years[0]
    prev = int_years[0]
    for y in int_years[1:]:
        if y == prev + 1:
            prev = y
        else:
            parts.append(
                str(run_start) if run_start == prev
                else f"{run_start}-{prev}"
            )
            run_start = y
            prev = y
    parts.append(
        str(run_start) if run_start == prev
        else f"{run_start}-{prev}"
    )
    return ", ".join(parts)


class EvalYearsTab(QWidget):
    """Evaluation Years tab displaying year ranges across all alignments."""

    status_message = Signal(str)
    progress_update = Signal(int, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._project_path = ""
        self._worker = None
        self._setup_ui()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def set_project_path(self, path: str):
        self._project_path = path

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # -- top row: button + summary ----------------------------------
        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        self._scan_btn = QPushButton("Scan for Evaluation Years")
        self._scan_btn.setProperty("accent", True)
        self._scan_btn.clicked.connect(self._scan_years)
        top_row.addWidget(self._scan_btn)

        self._summary_label = QLabel("No scan performed yet")
        self._summary_label.setProperty("subheader", True)
        top_row.addWidget(self._summary_label)

        top_row.addStretch()
        layout.addLayout(top_row)

        # -- tree -------------------------------------------------------
        self._tree = ScrollableTree(
            ["Folder ID", "Alignment Name", "Available Years", "Year Count"],
        )
        self._tree.set_column_widths([100, 320, 280, 90])
        layout.addWidget(self._tree, stretch=1)

    # ------------------------------------------------------------------
    # Scan
    # ------------------------------------------------------------------

    def _scan_years(self):
        if not self._project_path or not os.path.isdir(self._project_path):
            QMessageBox.warning(
                self, "Invalid Project",
                "Please select a valid project directory first.",
            )
            return

        self._scan_btn.setEnabled(False)
        self._tree.clear()
        self._summary_label.setText("Scanning...")
        self.status_message.emit("Scanning for evaluation years...")

        self._worker = YearScanWorker(self._project_path, parent=self)
        self._worker.progress.connect(self._on_scan_progress)
        self._worker.finished.connect(self._on_years_found)
        self._worker.error.connect(self._on_scan_error)
        self._worker.start()

    def _on_scan_progress(self, pct: int, msg: str):
        self.progress_update.emit(pct, msg)

    def _on_years_found(self, data: dict):
        self._scan_btn.setEnabled(True)
        self._worker = None

        all_years: List[str] = data.get("all_years", [])
        alignment_years: Dict[tuple, list] = data.get("alignment_years", {})

        if not alignment_years:
            self._summary_label.setText("No evaluation years found")
            self.status_message.emit("No evaluation years found in project")
            return

        self._populate_tree(alignment_years)

        unique_count = len(all_years)
        alignment_count = len(alignment_years)
        self._summary_label.setText(
            f"Found {alignment_count} alignments with {unique_count} unique years"
        )
        self.status_message.emit(
            f"Found {alignment_count} alignments, {unique_count} unique years"
        )

    def _on_scan_error(self, msg: str):
        self._scan_btn.setEnabled(True)
        self._summary_label.setText("Scan failed")
        self.status_message.emit(f"Year scan failed: {msg}")
        QMessageBox.critical(
            self, "Scan Error",
            f"Error scanning evaluation years:\n{msg}",
        )
        self._worker = None

    # ------------------------------------------------------------------
    # Tree population
    # ------------------------------------------------------------------

    def _populate_tree(self, alignment_years: Dict[tuple, list]):
        self._tree.clear()

        sorted_items = sorted(
            alignment_years.items(),
            key=lambda item: _sort_key(item[0][0]),
        )

        for (folder_id, alignment_name), years in sorted_items:
            sorted_years = sorted(years)
            years_display = _format_years(sorted_years)
            year_count = str(len(sorted_years))

            item = QTreeWidgetItem([
                folder_id,
                alignment_name,
                years_display,
                year_count,
            ])
            item.setTextAlignment(0, Qt.AlignCenter)
            item.setTextAlignment(3, Qt.AlignCenter)
            self._tree.tree.addTopLevelItem(item)
