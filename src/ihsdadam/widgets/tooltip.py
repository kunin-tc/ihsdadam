"""Rich tooltip widget for hover information"""

from PySide6.QtWidgets import QLabel, QWidget, QVBoxLayout
from PySide6.QtCore import Qt, QPoint, QTimer
from PySide6.QtGui import QCursor


class ToolTip(QWidget):
    """Floating tooltip that follows the cursor"""

    def __init__(self, parent=None):
        super().__init__(parent, Qt.ToolTip | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground, False)
        self.setStyleSheet(
            "QWidget { background-color: #ffffd0; border: 1px solid #c4c4c4; "
            "border-radius: 4px; }"
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 4, 6, 4)

        self._label = QLabel()
        self._label.setStyleSheet(
            "QLabel { background: transparent; color: #1b1b1b; "
            "font-family: 'Cascadia Code', 'Consolas', monospace; font-size: 9pt; }"
        )
        self._label.setWordWrap(True)
        layout.addWidget(self._label)

        self.hide()
        self._hide_timer = QTimer(self)
        self._hide_timer.setSingleShot(True)
        self._hide_timer.timeout.connect(self.hide)

    def show_text(self, text, pos=None, timeout=0):
        """Show tooltip with text at given position or near cursor."""
        self._label.setText(text)
        self.adjustSize()

        if pos is None:
            pos = QCursor.pos() + QPoint(15, 10)
        self.move(pos)
        self.show()
        self.raise_()

        if timeout > 0:
            self._hide_timer.start(timeout)
        else:
            self._hide_timer.stop()

    def hide_tip(self):
        self.hide()
        self._hide_timer.stop()
