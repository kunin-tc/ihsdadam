"""Appendix Generator tab -- merges IHSDM evaluation report PDFs into one."""

import os
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel,
    QPushButton, QLineEdit, QCheckBox, QTextEdit, QFileDialog,
    QMessageBox, QScrollArea, QSplitter,
)
from PySide6.QtCore import Qt, Signal

from .. import theme
from ..workers import AppendixMergeWorker, _get_alignment_name, _folder_prefix


class AppendixTab(QWidget):
    """Tab for scanning and merging IHSDM evaluation report PDFs."""

    status_message = Signal(str)
    progress_update = Signal(int, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._project_path = ""
        self._pdf_checkboxes: dict[str, QCheckBox] = {}
        self._worker = None
        self._setup_ui()

    # ── public API ──────────────────────────────────────────────────────

    def set_project_path(self, path: str):
        self._project_path = path

    # ── UI construction ─────────────────────────────────────────────────

    def _setup_ui(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(10, 10, 10, 10)

        splitter = QSplitter(Qt.Horizontal)
        root_layout.addWidget(splitter)

        # ── left panel ──────────────────────────────────────────────────
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # description
        desc_group = QGroupBox("Appendix Generator")
        desc_layout = QVBoxLayout(desc_group)
        desc_label = QLabel(
            "IMPORTANT: Before using this tool, generate PDF reports in IHSDM:\n"
            "  1. Open each alignment evaluation in IHSDM\n"
            "  2. Click \"Show PDF\" on the evaluation running\n"
            "     (or go to: Show Report > PDF > Multipage)\n"
            "  3. Save as evaluation.1.report.pdf in the evaluation folder\n\n"
            "This tool combines evaluation.*.report.pdf files into a single PDF.\n"
            "Select which alignments to include below."
        )
        desc_label.setWordWrap(True)
        desc_label.setProperty("caption", True)
        desc_layout.addWidget(desc_label)
        left_layout.addWidget(desc_group)

        # button row
        btn_row = QHBoxLayout()
        self._btn_scan = QPushButton("Scan for Reports")
        self._btn_scan.setProperty("accent", True)
        self._btn_scan.clicked.connect(self._scan_reports)

        self._btn_select_all = QPushButton("Select All")
        self._btn_select_all.clicked.connect(self._select_all)

        self._btn_deselect_all = QPushButton("Deselect All")
        self._btn_deselect_all.clicked.connect(self._deselect_all)

        btn_row.addWidget(self._btn_scan)
        btn_row.addWidget(self._btn_select_all)
        btn_row.addWidget(self._btn_deselect_all)
        btn_row.addStretch()
        left_layout.addLayout(btn_row)

        # checkbox scroll area
        reports_group = QGroupBox("Select Reports to Include")
        reports_layout = QVBoxLayout(reports_group)

        self._scroll_area = QScrollArea()
        self._scroll_area.setWidgetResizable(True)
        self._scroll_content = QWidget()
        self._scroll_layout = QVBoxLayout(self._scroll_content)
        self._scroll_layout.setAlignment(Qt.AlignTop)
        self._scroll_area.setWidget(self._scroll_content)
        reports_layout.addWidget(self._scroll_area)
        left_layout.addWidget(reports_group, stretch=1)

        splitter.addWidget(left_widget)

        # ── right panel ─────────────────────────────────────────────────
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # output file
        output_group = QGroupBox("Output PDF File")
        output_layout = QHBoxLayout(output_group)
        self._output_path_edit = QLineEdit()
        self._output_path_edit.setPlaceholderText("Select output PDF file...")
        output_layout.addWidget(self._output_path_edit, stretch=1)
        btn_browse = QPushButton("Browse...")
        btn_browse.clicked.connect(self._browse_output)
        output_layout.addWidget(btn_browse)
        right_layout.addWidget(output_group)

        # generate button
        self._btn_generate = QPushButton("Generate Appendix PDF")
        self._btn_generate.setProperty("accent", True)
        self._btn_generate.clicked.connect(self._generate_appendix)
        right_layout.addWidget(self._btn_generate)

        # status / log
        log_group = QGroupBox("Status")
        log_layout = QVBoxLayout(log_group)
        self._log_text = QTextEdit()
        self._log_text.setReadOnly(True)
        self._log_text.setStyleSheet(
            f"font-family: 'Cascadia Code', 'Consolas', monospace; font-size: 9pt;"
        )
        self._log_text.setPlainText(
            "Click 'Scan for Reports' to find available evaluation reports.\n"
        )
        log_layout.addWidget(self._log_text)
        right_layout.addWidget(log_group, stretch=1)

        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)

    # ── helpers ─────────────────────────────────────────────────────────

    def _log(self, message: str):
        self._log_text.append(message)
        scrollbar = self._log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def _clear_scroll_area(self):
        """Remove all widgets from the checkbox scroll layout."""
        while self._scroll_layout.count():
            item = self._scroll_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self._pdf_checkboxes.clear()

    def _add_section_header(self, text: str):
        header = QLabel(text)
        header.setProperty("header", True)
        header.setStyleSheet(
            f"font-size: 10pt; font-weight: 700; color: {theme.PRIMARY}; "
            f"padding-top: 8px;"
        )
        self._scroll_layout.addWidget(header)

    def _add_report_checkbox(self, pdf_path: str, display_name: str):
        cb = QCheckBox(display_name)
        cb.setChecked(True)
        self._scroll_layout.addWidget(cb)
        self._pdf_checkboxes[pdf_path] = cb

    def _get_alignment_name_from_pdf(self, pdf_path: Path) -> str:
        """Derive alignment name from XML files near the PDF."""
        try:
            alignment_dir = pdf_path.parent.parent
            return _get_alignment_name(alignment_dir)
        except Exception:
            return pdf_path.parent.parent.name

    # ── scan ────────────────────────────────────────────────────────────

    def _scan_reports(self):
        if not self._project_path or not os.path.isdir(self._project_path):
            QMessageBox.critical(self, "Error", "Please select a valid project directory.")
            return

        self.status_message.emit("Scanning for evaluation reports...")
        self._log_text.clear()
        self._clear_scroll_area()

        project_dir = Path(self._project_path)
        self._log(f"Scanning project: {project_dir}")
        self._log(f"Looking for pattern: **/e*/evaluation.*.report.pdf")
        self._log("")

        try:
            pdf_files = sorted(project_dir.glob("**/e*/evaluation.*.report.pdf"))

            self._log(f"Found {len(pdf_files)} files with e*/ pattern")

            if not pdf_files:
                self._log("Trying fallback pattern: **/evaluation.*.report.pdf")
                pdf_files = sorted(project_dir.glob("**/evaluation.*.report.pdf"))
                self._log(f"Found {len(pdf_files)} files with fallback pattern")

            if not pdf_files:
                self._log("")
                all_pdfs = sorted(project_dir.glob("**/*.pdf"))
                self._log(f"Found {len(all_pdfs)} total PDF files in project")
                if all_pdfs:
                    self._log("")
                    self._log("Sample PDF files found:")
                    for pdf in all_pdfs[:10]:
                        self._log(f"  {pdf.relative_to(project_dir)}")
                    if len(all_pdfs) > 10:
                        self._log(f"  ... and {len(all_pdfs) - 10} more")
                self._log("")
                self._log("No evaluation.*.report.pdf files found.")
                QMessageBox.warning(
                    self, "No Files",
                    "No evaluation report PDFs found in the project.",
                )
                return

            # group by type
            highways: list[dict] = []
            intersections: list[dict] = []
            ramps: list[dict] = []
            other: list[dict] = []

            for pdf_file in pdf_files:
                rel_path = pdf_file.relative_to(project_dir)
                parts = rel_path.parts

                alignment_folder = None
                for part in parts:
                    pfx = _folder_prefix(part) if part else ""
                    if pfx in ("h", "i", "r", "ra", "c"):
                        alignment_folder = part
                        break

                alignment_name = self._get_alignment_name_from_pdf(pdf_file)
                display_name = (
                    f"{'/'.join(parts[:-2])} - {alignment_name}"
                    if alignment_name
                    else str(rel_path.parent.parent)
                )

                entry = {
                    "path": str(pdf_file),
                    "display": display_name,
                }

                pfx = _folder_prefix(alignment_folder) if alignment_folder else ""
                if pfx == "h":
                    highways.append(entry)
                elif pfx in ("i", "ra"):
                    intersections.append(entry)
                elif pfx == "r":
                    ramps.append(entry)
                else:
                    other.append(entry)

            # populate scroll area
            if highways:
                self._add_section_header("HIGHWAYS")
                for entry in highways:
                    self._add_report_checkbox(entry["path"], entry["display"])
                    self._log(f"  [H] {entry['display']}")

            if intersections:
                self._add_section_header("INTERSECTIONS")
                for entry in intersections:
                    self._add_report_checkbox(entry["path"], entry["display"])
                    self._log(f"  [I] {entry['display']}")

            if ramps:
                self._add_section_header("RAMP TERMINALS")
                for entry in ramps:
                    self._add_report_checkbox(entry["path"], entry["display"])
                    self._log(f"  [R] {entry['display']}")

            if other:
                self._add_section_header("OTHER")
                for entry in other:
                    self._add_report_checkbox(entry["path"], entry["display"])
                    self._log(f"  [?] {entry['display']}")

            total = len(pdf_files)
            self._log("")
            self._log(f"All {total} reports selected by default.")
            self.status_message.emit(f"Found {total} evaluation reports")

        except Exception as exc:
            self._log(f"Error: {exc}")
            QMessageBox.critical(
                self, "Error", f"Error scanning for reports: {exc}"
            )

    # ── select / deselect ───────────────────────────────────────────────

    def _select_all(self):
        for cb in self._pdf_checkboxes.values():
            cb.setChecked(True)

    def _deselect_all(self):
        for cb in self._pdf_checkboxes.values():
            cb.setChecked(False)

    # ── browse output ───────────────────────────────────────────────────

    def _browse_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Select Output PDF File",
            "",
            "PDF files (*.pdf);;All files (*.*)",
        )
        if path:
            self._output_path_edit.setText(path)

    # ── generate ────────────────────────────────────────────────────────

    def _generate_appendix(self):
        if not self._project_path or not os.path.isdir(self._project_path):
            QMessageBox.critical(self, "Error", "Please select a valid project directory.")
            return

        output_path = self._output_path_edit.text().strip()
        if not output_path:
            QMessageBox.critical(self, "Error", "Please select an output PDF file.")
            return

        # collect checked PDF paths
        selected_paths = [
            path for path, cb in self._pdf_checkboxes.items() if cb.isChecked()
        ]

        if not selected_paths:
            QMessageBox.warning(
                self, "No Files",
                "No evaluation reports selected. Please scan and select reports first.",
            )
            return

        self._log_text.clear()
        self._log("Starting appendix generation...")
        self._log(f"Project: {self._project_path}")
        self._log(f"Output: {output_path}")
        self._log("")
        self._log(f"Selected {len(selected_paths)} PDF files to merge:")
        project_dir = Path(self._project_path)
        for pdf in sorted(selected_paths):
            try:
                rel = Path(pdf).relative_to(project_dir)
                self._log(f"  - {rel}")
            except ValueError:
                self._log(f"  - {pdf}")
        self._log("")
        self._log("Merging PDFs...")

        self._btn_generate.setEnabled(False)
        self.status_message.emit("Generating appendix...")

        self._worker = AppendixMergeWorker(selected_paths, output_path, parent=self)
        self._worker.log.connect(self._on_merge_log)
        self._worker.progress.connect(self._on_merge_progress)
        self._worker.finished.connect(self._on_merge_finished)
        self._worker.error.connect(self._on_merge_error)
        self._worker.start()

    def _on_merge_log(self, text: str):
        self._log(text)

    def _on_merge_progress(self, pct: int, msg: str):
        self.progress_update.emit(pct, msg)

    def _on_merge_finished(self, count: int):
        self._btn_generate.setEnabled(True)
        output_path = self._output_path_edit.text().strip()
        self._log("")
        self._log(f"SUCCESS! Combined PDF created: {output_path}")
        self.status_message.emit(f"Appendix generated: {count} PDFs combined")
        QMessageBox.information(
            self,
            "Success",
            f"Appendix PDF generated successfully!\n\n"
            f"Combined {count} evaluation reports into:\n{output_path}",
        )

    def _on_merge_error(self, msg: str):
        self._btn_generate.setEnabled(True)
        self._log("")
        self._log(f"ERROR: {msg}")
        self.status_message.emit("Error generating appendix")
        QMessageBox.critical(self, "Error", f"Error generating appendix: {msg}")
