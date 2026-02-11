"""QTreeWidget wrapper with built-in sorting and CSV export"""

import csv
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QFileDialog, QMessageBox
)
from PySide6.QtCore import Qt


class ScrollableTree(QWidget):
    """Tree widget with sortable columns and export capability"""

    def __init__(self, columns, parent=None):
        super().__init__(parent)
        self._columns = columns

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(columns)
        self.tree.setAlternatingRowColors(False)
        self.tree.setRootIsDecorated(True)
        self.tree.setSortingEnabled(False)
        self.tree.setSelectionMode(QTreeWidget.ExtendedSelection)
        self.tree.header().setStretchLastSection(True)
        self.tree.header().setSectionsClickable(True)
        self.tree.header().sectionClicked.connect(self._on_header_clicked)

        layout.addWidget(self.tree)

        self._sort_column = -1
        self._sort_order = Qt.AscendingOrder

    def _on_header_clicked(self, index):
        if self._sort_column == index:
            self._sort_order = (
                Qt.DescendingOrder if self._sort_order == Qt.AscendingOrder
                else Qt.AscendingOrder
            )
        else:
            self._sort_column = index
            self._sort_order = Qt.AscendingOrder
        self.tree.sortItems(index, self._sort_order)

    def set_column_widths(self, widths):
        for i, w in enumerate(widths):
            if i < self.tree.columnCount():
                self.tree.setColumnWidth(i, w)

    def clear(self):
        self.tree.clear()

    def add_top_level_item(self, texts, data=None):
        item = QTreeWidgetItem(texts)
        if data is not None:
            item.setData(0, Qt.UserRole, data)
        self.tree.addTopLevelItem(item)
        return item

    def export_to_csv(self, default_name="export.csv"):
        path, _ = QFileDialog.getSaveFileName(
            self, "Export to CSV", default_name,
            "CSV Files (*.csv);;All Files (*)"
        )
        if not path:
            return

        try:
            with open(path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(self._columns)
                self._write_items(writer, self.tree.invisibleRootItem())
            QMessageBox.information(self, "Export Complete", f"Data exported to {path}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", str(e))

    def _write_items(self, writer, parent):
        for i in range(parent.childCount()):
            child = parent.child(i)
            row = [child.text(c) for c in range(self.tree.columnCount())]
            if any(row):
                writer.writerow(row)
            self._write_items(writer, child)
