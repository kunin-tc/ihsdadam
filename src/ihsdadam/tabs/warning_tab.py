"""Warning Extractor tab -- scans IHSDM evaluation XML for ResultMessage entries."""

import csv
import os
from collections import defaultdict
from typing import Dict, List

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .. import theme
from ..models import ResultMessage
from ..widgets import ScrollableTree
from ..workers import WarningScanWorker, _format_station


# ── Status-to-colour mapping ────────────────────────────────────────────────
_STATUS_COLOURS: Dict[str, QColor] = {
    "CRITICAL": QColor(theme.CRITICAL_BG),
    "error": QColor(theme.ERROR_BG),
    "warning": QColor(theme.WARNING_BG),
}

_STATUS_FG: Dict[str, QColor] = {
    "CRITICAL": QColor(theme.CRITICAL_FG),
    "error": QColor(theme.ERROR_FG),
    "warning": QColor(theme.WARNING_FG),
}

# Labels used as separator items in the alignment filter combo
_SEPARATOR_LABELS = frozenset({
    "--- HIGHWAYS ---",
    "--- INTERSECTIONS ---",
    "--- RAMP TERMINALS ---",
})

_TYPE_NAMES: Dict[str, str] = {
    "h": "Highway",
    "i": "Intersection",
    "r": "Ramp Terminal",
}

_TREE_COLUMNS = ["Type", "Alignment", "Eval", "Status", "Start Sta", "End Sta", "Message"]


class WarningTab(QWidget):
    """Warning Extractor tab -- scan, filter, browse and export result messages."""

    status_message = Signal(str)
    progress_update = Signal(int, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._project_path = ""
        self._messages: List[ResultMessage] = []
        self._filtered_messages: List[ResultMessage] = []
        self._worker: WarningScanWorker | None = None
        self._setup_ui()

    # ── public API ───────────────────────────────────────────────────────

    def set_project_path(self, path: str):
        self._project_path = path

    # ── UI construction ──────────────────────────────────────────────────

    def _setup_ui(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(10, 10, 10, 10)
        root_layout.setSpacing(10)

        # -- Scan button row ------------------------------------------------
        scan_row = QHBoxLayout()
        self._scan_btn = QPushButton("Scan for Warning Messages")
        self._scan_btn.setProperty("accent", True)
        self._scan_btn.clicked.connect(self._scan_warnings)
        scan_row.addWidget(self._scan_btn)
        scan_row.addStretch()
        root_layout.addLayout(scan_row)

        # -- Summary panel --------------------------------------------------
        self._summary_label = QLabel("No project scanned yet")
        self._summary_label.setProperty("subheader", True)
        self._summary_label.setWordWrap(True)
        root_layout.addWidget(self._summary_label)

        # -- Body: filter panel (left) + tree (right) ----------------------
        body_layout = QHBoxLayout()
        body_layout.setSpacing(8)

        # Left: filters
        filter_box = QGroupBox("Filters")
        filter_layout = QVBoxLayout(filter_box)
        filter_layout.setSpacing(6)

        filter_layout.addWidget(QLabel("Type:"))
        self._type_combo = QComboBox()
        self._type_combo.addItems(["All", "Highways (h)", "Intersections (i)", "Ramp Terminals (r)"])
        self._type_combo.currentIndexChanged.connect(self._apply_filters)
        filter_layout.addWidget(self._type_combo)

        filter_layout.addWidget(QLabel("Alignment:"))
        self._alignment_combo = QComboBox()
        self._alignment_combo.addItem("All")
        self._alignment_combo.currentIndexChanged.connect(self._apply_filters)
        filter_layout.addWidget(self._alignment_combo)

        filter_layout.addWidget(QLabel("Message Type:"))
        self._msg_type_combo = QComboBox()
        self._msg_type_combo.addItems(["All", "CRITICAL", "error", "warning", "fault", "info"])
        self._msg_type_combo.currentIndexChanged.connect(self._apply_filters)
        filter_layout.addWidget(self._msg_type_combo)

        filter_layout.addWidget(QLabel("Search Text:"))
        self._search_input = QLineEdit()
        self._search_input.setPlaceholderText("Filter by text...")
        self._search_input.setClearButtonEnabled(True)
        self._search_input.textChanged.connect(self._apply_filters)
        filter_layout.addWidget(self._search_input)

        clear_btn = QPushButton("Clear Filters")
        clear_btn.clicked.connect(self._clear_filters)
        filter_layout.addWidget(clear_btn)

        filter_layout.addStretch()
        filter_box.setFixedWidth(260)
        body_layout.addWidget(filter_box)

        # Right: tree
        self._tree = ScrollableTree(_TREE_COLUMNS)
        self._tree.set_column_widths([40, 250, 50, 90, 110, 110, 600])
        body_layout.addWidget(self._tree, stretch=1)

        root_layout.addLayout(body_layout, stretch=1)

        # -- Bottom button row ----------------------------------------------
        btn_row = QHBoxLayout()
        export_btn = QPushButton("Export to CSV")
        export_btn.clicked.connect(self._export_csv)
        btn_row.addWidget(export_btn)

        copy_btn = QPushButton("Copy Selected")
        copy_btn.clicked.connect(self._copy_selected)
        btn_row.addWidget(copy_btn)

        btn_row.addStretch()
        root_layout.addLayout(btn_row)

    # ── Scanning ─────────────────────────────────────────────────────────

    def _scan_warnings(self):
        if not self._project_path or not os.path.isdir(self._project_path):
            QMessageBox.critical(self, "Error", "Please select a valid project directory.")
            return

        self._scan_btn.setEnabled(False)
        self.status_message.emit("Scanning project for warnings...")

        self._worker = WarningScanWorker(self._project_path, parent=self)
        self._worker.progress.connect(self._on_scan_progress)
        self._worker.finished.connect(self._on_scan_finished)
        self._worker.error.connect(self._on_scan_error)
        self._worker.start()

    def _on_scan_progress(self, pct: int, msg: str):
        self.progress_update.emit(pct, msg)
        self.status_message.emit(msg)

    def _on_scan_finished(self, messages: list):
        self._messages = list(messages)
        self._scan_btn.setEnabled(True)

        self._update_summary()
        self._update_alignment_filter()
        self._apply_filters()

        critical = sum(1 for m in self._messages if m.is_critical)
        text = f"Found {len(self._messages)} messages"
        if critical:
            text += f" ({critical} CRITICAL)"
            QMessageBox.warning(
                self,
                "Scan Complete - Critical Messages Found",
                f"{text}\n\n"
                f"WARNING: {critical} CRITICAL messages found!\n"
                "(Messages containing 'no crash prediction supported')",
            )
        else:
            QMessageBox.information(self, "Scan Complete", text)

        self.status_message.emit(f"Scan complete. {text}")

    def _on_scan_error(self, msg: str):
        self._scan_btn.setEnabled(True)
        QMessageBox.critical(self, "Scan Error", msg)
        self.status_message.emit("Error during scan")

    # ── Summary ──────────────────────────────────────────────────────────

    def _update_summary(self):
        if not self._messages:
            self._summary_label.setText("No messages found")
            return

        status_counts: Dict[str, int] = defaultdict(int)
        hw = set()
        ints = set()
        ramps = set()

        for m in self._messages:
            status_counts[m.status] += 1
            label = f"{m.alignment_id} - {m.alignment_name}"
            if m.alignment_type == "h":
                hw.add(label)
            elif m.alignment_type == "i":
                ints.add(label)
            elif m.alignment_type == "r":
                ramps.add(label)

        parts = [
            f"Total: {len(self._messages)}",
            f"Highways: {len(hw)}",
            f"Intersections: {len(ints)}",
            f"Ramp Terminals: {len(ramps)}",
        ]
        if status_counts["CRITICAL"]:
            parts.append(f"CRITICAL: {status_counts['CRITICAL']}")
        parts.append(f"Errors: {status_counts['error']}")
        parts.append(f"Warnings: {status_counts['warning']}")
        parts.append(f"Info: {status_counts['info']}")

        self._summary_label.setText(" | ".join(parts))

    # ── Alignment filter population ──────────────────────────────────────

    def _update_alignment_filter(self):
        hw: set = set()
        ints: set = set()
        ramps: set = set()

        for m in self._messages:
            label = f"{m.alignment_id} - {m.alignment_name}"
            if m.alignment_type == "h":
                hw.add(label)
            elif m.alignment_type == "i":
                ints.add(label)
            elif m.alignment_type == "r":
                ramps.add(label)

        self._alignment_combo.blockSignals(True)
        self._alignment_combo.clear()
        self._alignment_combo.addItem("All")

        if hw:
            self._alignment_combo.addItem("--- HIGHWAYS ---")
            for a in sorted(hw):
                self._alignment_combo.addItem(a)
        if ints:
            self._alignment_combo.addItem("--- INTERSECTIONS ---")
            for a in sorted(ints):
                self._alignment_combo.addItem(a)
        if ramps:
            self._alignment_combo.addItem("--- RAMP TERMINALS ---")
            for a in sorted(ramps):
                self._alignment_combo.addItem(a)

        self._alignment_combo.setCurrentIndex(0)
        self._alignment_combo.blockSignals(False)

    # ── Filtering & tree population ──────────────────────────────────────

    def _apply_filters(self):
        type_filter = self._type_combo.currentText()
        alignment_filter = self._alignment_combo.currentText()
        msg_type_filter = self._msg_type_combo.currentText()
        search_text = self._search_input.text().lower()

        filtered: List[ResultMessage] = []

        for msg in self._messages:
            # Alignment type filter
            if type_filter == "Highways (h)" and msg.alignment_type != "h":
                continue
            if type_filter == "Intersections (i)" and msg.alignment_type != "i":
                continue
            if type_filter == "Ramp Terminals (r)" and msg.alignment_type != "r":
                continue

            # Specific alignment filter
            if alignment_filter not in ("All", *_SEPARATOR_LABELS):
                label = f"{msg.alignment_id} - {msg.alignment_name}"
                if label != alignment_filter:
                    continue

            # Message type filter
            if msg_type_filter != "All" and msg.status != msg_type_filter:
                continue

            # Search text filter
            if search_text:
                searchable = f"{msg.message} {msg.alignment_name}".lower()
                if search_text not in searchable:
                    continue

            filtered.append(msg)

        self._filtered_messages = filtered

        # Group into dicts keyed by "id - name"
        highways: Dict[str, List[ResultMessage]] = defaultdict(list)
        intersections: Dict[str, List[ResultMessage]] = defaultdict(list)
        ramp_terminals: Dict[str, List[ResultMessage]] = defaultdict(list)

        for msg in filtered:
            key = f"{msg.alignment_id} - {msg.alignment_name}"
            if msg.alignment_type == "h":
                highways[key].append(msg)
            elif msg.alignment_type == "i":
                intersections[key].append(msg)
            elif msg.alignment_type == "r":
                ramp_terminals[key].append(msg)

        self._tree.clear()
        self._populate_tree_section("HIGHWAYS", highways, "h")
        self._populate_tree_section("INTERSECTIONS", intersections, "i")
        self._populate_tree_section("RAMP TERMINALS", ramp_terminals, "r")

        critical = sum(1 for m in filtered if m.is_critical)
        status = f"Showing {len(filtered)} of {len(self._messages)} messages"
        if critical:
            status += f" ({critical} CRITICAL)"
        self.status_message.emit(status)

    def _populate_tree_section(
        self,
        section_name: str,
        alignments_dict: Dict[str, List[ResultMessage]],
        type_char: str,
    ):
        if not alignments_dict:
            return

        col_count = len(_TREE_COLUMNS)

        # Section header
        section_texts = ["", f"({len(alignments_dict)} alignments)", "", "", "", "", ""]
        section_item = QTreeWidgetItem(section_texts)
        section_item.setText(0, section_name)
        section_item.setFlags(section_item.flags() & ~Qt.ItemIsSelectable)
        self._tree.tree.addTopLevelItem(section_item)
        section_item.setExpanded(True)

        for alignment, msgs in sorted(alignments_dict.items()):
            crit = sum(1 for m in msgs if m.is_critical)
            errs = sum(1 for m in msgs if m.status == "error")
            warns = sum(1 for m in msgs if m.status == "warning")

            label_parts = [f"{len(msgs)} msgs"]
            if crit:
                label_parts.append(f"{crit} CRITICAL")
            if errs:
                label_parts.append(f"{errs} errors")
            if warns:
                label_parts.append(f"{warns} warnings")

            summary = f"{alignment} ({', '.join(label_parts)})"

            align_texts = [type_char, alignment, "", "", "", "", summary]
            align_item = QTreeWidgetItem(align_texts)
            section_item.addChild(align_item)
            align_item.setExpanded(True)

            for msg in sorted(msgs, key=lambda m: (m.evaluation, m.start_sta)):
                formatted_start = _format_station(msg.start_sta)
                formatted_end = _format_station(msg.end_sta)

                row_texts = [
                    "",
                    "",
                    msg.evaluation,
                    msg.status,
                    formatted_start,
                    formatted_end,
                    msg.message,
                ]
                item = QTreeWidgetItem(row_texts)

                # Apply row colouring
                bg = _STATUS_COLOURS.get(msg.status)
                fg = _STATUS_FG.get(msg.status)
                if bg:
                    for c in range(col_count):
                        item.setBackground(c, bg)
                if fg:
                    for c in range(col_count):
                        item.setForeground(c, fg)

                align_item.addChild(item)

    # ── Filter helpers ───────────────────────────────────────────────────

    def _clear_filters(self):
        self._type_combo.setCurrentIndex(0)
        self._alignment_combo.setCurrentIndex(0)
        self._msg_type_combo.setCurrentIndex(0)
        self._search_input.clear()
        self._apply_filters()

    # ── Export / Copy ────────────────────────────────────────────────────

    def _export_csv(self):
        if not self._filtered_messages:
            QMessageBox.warning(self, "No Data", "No messages to export.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Messages to CSV",
            "warnings_export.csv",
            "CSV Files (*.csv);;All Files (*)",
        )
        if not path:
            return

        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([
                    "Type",
                    "Alignment ID",
                    "Alignment Name",
                    "Evaluation",
                    "Status",
                    "Is Critical",
                    "Start Station",
                    "End Station",
                    "Message",
                    "File Path",
                ])
                for msg in self._filtered_messages:
                    writer.writerow([
                        _TYPE_NAMES.get(msg.alignment_type, msg.alignment_type),
                        msg.alignment_id,
                        msg.alignment_name,
                        msg.evaluation,
                        msg.status,
                        "YES" if msg.is_critical else "NO",
                        msg.start_sta,
                        msg.end_sta,
                        msg.message,
                        msg.file_path,
                    ])
            QMessageBox.information(
                self,
                "Export Complete",
                f"Exported {len(self._filtered_messages)} messages to {path}",
            )
            self.status_message.emit(f"Exported to {path}")
        except Exception as exc:
            QMessageBox.critical(self, "Export Error", str(exc))

    def _copy_selected(self):
        selected = self._tree.tree.selectedItems()
        if not selected:
            QMessageBox.warning(self, "No Selection", "Please select messages to copy.")
            return

        col_count = self._tree.tree.columnCount()
        lines: List[str] = []
        for item in selected:
            values = [item.text(c) for c in range(col_count)]
            if any(values):
                lines.append("\t".join(values))

        if lines:
            clipboard = QApplication.clipboard()
            clipboard.setText("\n".join(lines))
            self.status_message.emit(f"Copied {len(lines)} items to clipboard")
        else:
            QMessageBox.information(self, "No Data", "No message data to copy.")
