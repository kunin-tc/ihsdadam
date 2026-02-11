"""Windows 11 Fluent Design QSS stylesheet and colour constants"""

# -- Colour Palette (Fluent Light) --------------------------------------------
PRIMARY = "#005fb8"
PRIMARY_HOVER = "#0078d4"
PRIMARY_PRESS = "#004c95"
SURFACE = "#ffffff"
BACKGROUND = "#f3f3f3"
LAYER = "#fbfbfb"
BORDER = "#e5e5e5"
BORDER_STRONG = "#c4c4c4"
TEXT_PRIMARY = "#1b1b1b"
TEXT_SECONDARY = "#616161"
ACCENT_GREEN = "#0f7b0f"
DANGER = "#c42b1c"
WARNING = "#9d5d00"
SUBTLE_FILL = "#f5f5f5"

# Sidebar
SIDEBAR_BG = "#f0f0f0"
SIDEBAR_HOVER = "#e4e4e4"
SIDEBAR_SELECTED = "#dbeafe"
SIDEBAR_SELECTED_BORDER = PRIMARY

# Status-row colours (for Warning Extractor tag highlights)
CRITICAL_BG = "#fde8e8"
CRITICAL_FG = "#c42b1c"
ERROR_BG = "#fef2f2"
ERROR_FG = "#c42b1c"
WARNING_BG = "#fef9ee"
WARNING_FG = "#9d5d00"

# Lane-plan colours (for Visual View)
LANE_ASPHALT = "#2d2d2d"
LANE_LEFT_TURN = "#4a4a5a"
SHOULDER_GRAY = "#6b7280"
CENTERLINE_YELLOW = "#fbbf24"

# -- Font Family ---------------------------------------------------------------
FONT_FAMILY = '"Segoe UI", "Segoe UI Variable", Arial, sans-serif'

# -- Full QSS Stylesheet ------------------------------------------------------
STYLESHEET = f"""
/* -- Global ----------------------------------------------------------- */
QWidget {{
    font-family: {FONT_FAMILY};
    font-size: 10pt;
    color: {TEXT_PRIMARY};
    background-color: {BACKGROUND};
}}

/* -- Main Window ------------------------------------------------------ */
QMainWindow {{
    background-color: {BACKGROUND};
}}

/* -- Sidebar navigation (QListWidget) --------------------------------- */
QListWidget#sidebar {{
    background-color: {SIDEBAR_BG};
    border: none;
    border-right: 1px solid {BORDER};
    outline: none;
    font-size: 10pt;
    font-weight: 500;
    padding: 4px 0px;
}}

QListWidget#sidebar::item {{
    padding: 10px 18px;
    border-left: 3px solid transparent;
    border-radius: 0px;
    min-height: 20px;
    color: {TEXT_SECONDARY};
}}

QListWidget#sidebar::item:hover {{
    background-color: {SIDEBAR_HOVER};
    color: {TEXT_PRIMARY};
}}

QListWidget#sidebar::item:selected {{
    background-color: {SIDEBAR_SELECTED};
    border-left: 3px solid {SIDEBAR_SELECTED_BORDER};
    color: {PRIMARY};
    font-weight: 600;
}}

/* -- Group Box (replaces LabelFrame) ---------------------------------- */
QGroupBox {{
    background-color: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 8px;
    margin-top: 16px;
    padding: 16px 12px 12px 12px;
    font-weight: 600;
    font-size: 10pt;
    color: {TEXT_PRIMARY};
}}

QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 6px;
    color: {TEXT_PRIMARY};
    font-weight: 600;
}}

/* -- Push Button ------------------------------------------------------ */
QPushButton {{
    background-color: {SURFACE};
    color: {TEXT_PRIMARY};
    border: 1px solid {BORDER_STRONG};
    border-radius: 4px;
    padding: 6px 16px;
    font-size: 10pt;
    font-weight: 500;
    min-height: 24px;
}}

QPushButton:hover {{
    background-color: {SUBTLE_FILL};
    border-color: {BORDER_STRONG};
}}

QPushButton:pressed {{
    background-color: {BORDER};
}}

QPushButton:disabled {{
    color: {BORDER_STRONG};
    background-color: {SUBTLE_FILL};
    border-color: {BORDER};
}}

/* Primary button (accent) */
QPushButton[accent="true"] {{
    background-color: {PRIMARY};
    color: #ffffff;
    border: none;
    font-weight: 600;
}}

QPushButton[accent="true"]:hover {{
    background-color: {PRIMARY_HOVER};
}}

QPushButton[accent="true"]:pressed {{
    background-color: {PRIMARY_PRESS};
}}

/* Green accent button */
QPushButton[green="true"] {{
    background-color: {ACCENT_GREEN};
    color: #ffffff;
    border: none;
    font-weight: 600;
}}

QPushButton[green="true"]:hover {{
    background-color: #0a9b0a;
}}

/* -- Line Edit -------------------------------------------------------- */
QLineEdit {{
    background-color: {SURFACE};
    color: {TEXT_PRIMARY};
    border: 1px solid {BORDER_STRONG};
    border-bottom: 2px solid {BORDER_STRONG};
    border-radius: 4px;
    padding: 5px 8px;
    font-size: 10pt;
    selection-background-color: {PRIMARY_HOVER};
    selection-color: #ffffff;
}}

QLineEdit:focus {{
    border-bottom: 2px solid {PRIMARY};
}}

QLineEdit:disabled {{
    background-color: {SUBTLE_FILL};
    color: {TEXT_SECONDARY};
}}

/* -- Combo Box -------------------------------------------------------- */
QComboBox {{
    background-color: {SURFACE};
    color: {TEXT_PRIMARY};
    border: 1px solid {BORDER_STRONG};
    border-radius: 4px;
    padding: 5px 8px;
    font-size: 10pt;
    min-height: 22px;
}}

QComboBox:hover {{
    border-color: {PRIMARY_HOVER};
}}

QComboBox::drop-down {{
    border: none;
    width: 24px;
}}

QComboBox::down-arrow {{
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 5px solid {TEXT_SECONDARY};
    margin-right: 8px;
}}

QComboBox QAbstractItemView {{
    background-color: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 4px;
    selection-background-color: {SUBTLE_FILL};
    selection-color: {TEXT_PRIMARY};
    outline: none;
}}

/* -- Check Box / Radio Button ----------------------------------------- */
QCheckBox, QRadioButton {{
    spacing: 8px;
    font-size: 10pt;
}}

QCheckBox::indicator, QRadioButton::indicator {{
    width: 18px;
    height: 18px;
}}

/* -- Tree Widget ------------------------------------------------------ */
QTreeWidget, QTreeView {{
    background-color: {SURFACE};
    alternate-background-color: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 4px;
    outline: none;
    font-size: 10pt;
}}

QTreeWidget::item, QTreeView::item {{
    padding: 4px 6px;
    border: none;
}}

QTreeWidget::item:hover, QTreeView::item:hover {{
    background-color: {SUBTLE_FILL};
}}

QTreeWidget::item:selected, QTreeView::item:selected {{
    background-color: #cce4f7;
    color: {TEXT_PRIMARY};
}}

QHeaderView::section {{
    background-color: {SUBTLE_FILL};
    color: {TEXT_PRIMARY};
    font-weight: 600;
    font-size: 9pt;
    padding: 6px 8px;
    border: none;
    border-right: 1px solid {BORDER};
    border-bottom: 1px solid {BORDER};
}}

/* -- Scroll Bar (thin WinUI style) ------------------------------------ */
QScrollBar:vertical {{
    background-color: transparent;
    width: 12px;
    margin: 0;
}}

QScrollBar::handle:vertical {{
    background-color: #c1c1c1;
    min-height: 30px;
    border-radius: 3px;
    margin: 2px 3px;
}}

QScrollBar::handle:vertical:hover {{
    background-color: #a0a0a0;
}}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}

QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
    background: transparent;
}}

QScrollBar:horizontal {{
    background-color: transparent;
    height: 12px;
    margin: 0;
}}

QScrollBar::handle:horizontal {{
    background-color: #c1c1c1;
    min-width: 30px;
    border-radius: 3px;
    margin: 3px 2px;
}}

QScrollBar::handle:horizontal:hover {{
    background-color: #a0a0a0;
}}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0;
}}

QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
    background: transparent;
}}

/* -- Text Edit (ScrolledText replacement) ----------------------------- */
QTextEdit, QPlainTextEdit {{
    background-color: {SURFACE};
    color: {TEXT_PRIMARY};
    border: 1px solid {BORDER};
    border-radius: 4px;
    padding: 8px;
    font-family: "Cascadia Code", "Consolas", monospace;
    font-size: 9pt;
    selection-background-color: {PRIMARY_HOVER};
    selection-color: #ffffff;
}}

/* -- Progress Bar ----------------------------------------------------- */
QProgressBar {{
    background-color: {BORDER};
    border: none;
    border-radius: 2px;
    text-align: center;
    font-size: 8pt;
    max-height: 4px;
}}

QProgressBar::chunk {{
    background-color: {PRIMARY};
    border-radius: 2px;
}}

/* -- Status Bar area -------------------------------------------------- */
QStatusBar {{
    background-color: {LAYER};
    border-top: 1px solid {BORDER};
    font-size: 9pt;
    color: {TEXT_SECONDARY};
}}

/* -- Splitter --------------------------------------------------------- */
QSplitter::handle {{
    background-color: {BORDER};
    width: 1px;
    height: 1px;
}}

/* -- Tool Tip --------------------------------------------------------- */
QToolTip {{
    background-color: {SURFACE};
    color: {TEXT_PRIMARY};
    border: 1px solid {BORDER};
    border-radius: 4px;
    padding: 6px 10px;
    font-size: 9pt;
}}

/* -- Message Box ------------------------------------------------------ */
QMessageBox {{
    background-color: {SURFACE};
}}

QMessageBox QLabel {{
    color: {TEXT_PRIMARY};
    font-size: 10pt;
}}

/* -- Dialog ----------------------------------------------------------- */
QDialog {{
    background-color: {SURFACE};
}}

/* -- Label ------------------------------------------------------------ */
QLabel {{
    background-color: transparent;
    color: {TEXT_PRIMARY};
}}

QLabel[header="true"] {{
    font-size: 14px;
    font-weight: 600;
    color: {PRIMARY};
}}

QLabel[subheader="true"] {{
    font-size: 10pt;
    font-weight: 500;
    color: {TEXT_SECONDARY};
}}

QLabel[caption="true"] {{
    font-size: 9pt;
    color: {TEXT_SECONDARY};
}}

/* -- Separator -------------------------------------------------------- */
QFrame[frameShape="4"] {{
    color: {BORDER};
    max-height: 1px;
}}

QFrame[frameShape="5"] {{
    color: {BORDER};
    max-width: 1px;
}}
"""
