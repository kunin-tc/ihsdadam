"""Visual View tab -- displays highway alignment data graphically."""

import os
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel,
    QPushButton, QComboBox, QMessageBox,
)
from PySide6.QtCore import Qt, Signal

from .. import theme
from ..workers import VisualDataWorker, _get_alignment_name
from ..widgets import HighwayCanvas


class VisualTab(QWidget):
    """Tab for visualising highway alignment geometry from XML data."""

    status_message = Signal(str)
    progress_update = Signal(int, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._project_path = ""
        self._worker = None
        self._setup_ui()

    # ── public API ──────────────────────────────────────────────────────

    def set_project_path(self, path: str):
        self._project_path = path

    # ── UI construction ─────────────────────────────────────────────────

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # description
        desc_group = QGroupBox("Visual View")
        desc_layout = QVBoxLayout(desc_group)
        desc_label = QLabel(
            "This tool displays highway alignment data visually from highway XML files.\n"
            "View lane configuration, horizontal curves, traffic, speed, and more\n"
            "for any highway alignment in your project.\n\n"
            "Scroll to zoom, click-drag to pan. Hover elements for details. "
            "Press Home to fit view."
        )
        desc_label.setWordWrap(True)
        desc_label.setProperty("caption", True)
        desc_layout.addWidget(desc_label)
        layout.addWidget(desc_group)

        # alignment selector
        selector_group = QGroupBox("Select Alignment")
        selector_layout = QVBoxLayout(selector_group)

        # first row: load button
        load_row = QHBoxLayout()
        self._btn_load = QPushButton("Load Alignments")
        self._btn_load.clicked.connect(self._refresh_alignments)
        load_row.addWidget(self._btn_load)
        load_row.addStretch()
        selector_layout.addLayout(load_row)

        # second row: combo + display button
        combo_row = QHBoxLayout()
        lbl_alignment = QLabel("Highway Alignment:")
        combo_row.addWidget(lbl_alignment)

        self._combo = QComboBox()
        self._combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self._combo.setMinimumWidth(300)
        self._combo.addItem("No alignments loaded")
        combo_row.addWidget(self._combo, stretch=1)

        self._btn_display = QPushButton("Display Alignment")
        self._btn_display.setProperty("accent", True)
        self._btn_display.clicked.connect(self._display_alignment)
        combo_row.addWidget(self._btn_display)

        selector_layout.addLayout(combo_row)
        layout.addWidget(selector_group)

        # canvas (takes remaining space)
        self._canvas = HighwayCanvas()
        layout.addWidget(self._canvas, stretch=1)

    # ── refresh alignments ──────────────────────────────────────────────

    def _refresh_alignments(self):
        if not self._project_path or not os.path.isdir(self._project_path):
            QMessageBox.critical(self, "Error", "Please select a valid project directory.")
            return

        try:
            project_dir = Path(self._project_path)
            highway_dirs: list[Path] = []

            # direct h* directories
            for item in project_dir.iterdir():
                if item.is_dir() and item.name[0].lower() == "h":
                    highway_dirs.append(item)

            # h* inside c* interchange containers
            for item in project_dir.iterdir():
                if item.is_dir() and item.name[0].lower() == "c":
                    for sub in item.iterdir():
                        if sub.is_dir() and sub.name[0].lower() == "h":
                            highway_dirs.append(sub)

            if not highway_dirs:
                self._combo.clear()
                self._combo.addItem("No highway alignments found")
                QMessageBox.information(
                    self, "No Alignments",
                    "No highway alignments found in project directory.",
                )
                return

            self._combo.clear()
            for hw_dir in sorted(highway_dirs, key=lambda d: d.name):
                name = _get_alignment_name(hw_dir)
                self._combo.addItem(f"{hw_dir.name} - {name}")

            self.status_message.emit(
                f"Found {self._combo.count()} highway alignments"
            )

        except Exception as exc:
            QMessageBox.critical(
                self, "Error", f"Error finding alignments: {exc}"
            )

    # ── display alignment ───────────────────────────────────────────────

    def _display_alignment(self):
        if not self._project_path or not os.path.isdir(self._project_path):
            QMessageBox.critical(self, "Error", "Please select a valid project directory.")
            return

        selected = self._combo.currentText()
        if not selected or selected in (
            "No alignments loaded",
            "No highway alignments found",
        ):
            QMessageBox.warning(
                self, "No Selection",
                "Please select a highway alignment first.",
            )
            return

        # extract alignment ID from "h1 - Alignment Name"
        alignment_id = selected.split(" - ", 1)[0].strip()

        try:
            project_dir = Path(self._project_path)

            # locate the highway directory
            highway_dir = project_dir / alignment_id
            if not highway_dir.exists():
                for interchange_dir in project_dir.iterdir():
                    if (
                        interchange_dir.is_dir()
                        and interchange_dir.name[0].lower() == "c"
                    ):
                        candidate = interchange_dir / alignment_id
                        if candidate.exists():
                            highway_dir = candidate
                            break

            if not highway_dir.exists():
                QMessageBox.critical(
                    self, "Error",
                    f"Could not find directory for {alignment_id}",
                )
                return

            # find highest-version highway.*.xml (fall back to highway.xml)
            versioned = sorted(highway_dir.glob("highway.*.xml"))
            if versioned:
                highway_xml = versioned[-1]
            elif (highway_dir / "highway.xml").exists():
                highway_xml = highway_dir / "highway.xml"
            else:
                QMessageBox.critical(
                    self, "Error",
                    f"Could not find highway XML in {highway_dir}",
                )
                return

            self._canvas.clear()
            self._btn_display.setEnabled(False)
            self.status_message.emit(f"Loading {alignment_id}...")

            self._worker = VisualDataWorker(
                str(highway_xml), str(project_dir), parent=self
            )
            self._worker.finished.connect(self._on_visual_data)
            self._worker.error.connect(self._on_visual_error)
            self._worker.start()

        except Exception as exc:
            QMessageBox.critical(
                self, "Error", f"Error loading alignment: {exc}"
            )

    def _on_visual_data(self, data_dict: dict):
        self._btn_display.setEnabled(True)
        self._canvas.set_data(data_dict)
        title = data_dict.get("title", "Unknown")
        self.status_message.emit(f"Displaying: {title}")

    def _on_visual_error(self, msg: str):
        self._btn_display.setEnabled(True)
        self.status_message.emit("Error loading alignment")
        QMessageBox.critical(self, "Error", f"Error loading alignment data: {msg}")
