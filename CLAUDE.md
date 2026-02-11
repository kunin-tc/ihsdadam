# CLAUDE.md - IHSDadaM Project Guide

This file provides guidance to Claude Code when working with code in this repository.

## Project Overview

**IHSDadaM** (Interactive Highway Safety Design Model - Data Manager) is a Python/PySide6 desktop GUI tool for analyzing IHSDM highway safety projects. It extracts warning messages, compiles crash prediction data to Excel, merges PDF evaluation reports, visualizes highway geometry, scans calibration factors, manages AADT input, tracks evaluation years, and generates printable HTML crash prediction reports.

- **Version:** 2.0.1
- **Author:** Adam Engbring (aengbring@hntb.com)
- **Organization:** HNTB Wisconsin Office
- **Framework:** PySide6 (Qt 6) with Windows 11 Fluent Design theme
- **GitHub:** kunin-tc/ihsdadam (auto-update checks via GitHub API)

## Build & Run

```bash
# Install dependencies
pip install -r requirements.txt
pip install PySide6>=6.7,<6.8

# Run from source
python main.py

# Build standalone executable (creates dist/IHSDadaM.exe)
pyinstaller IHSDadaM.spec --clean
```

### Build Details

The build is driven entirely by `IHSDadaM.spec`. Do NOT use raw pyinstaller CLI flags -- always build from the spec.

```
Spec file:       IHSDadaM.spec
Entry point:     main.py
Output:          dist/IHSDadaM.exe (~55 MB, single file, windowed)
Icon:            ihsdadam.ico
Bundled data:    version.py, src/ihsdadam/ (full package)
Path:            src/ added to pathex so ihsdadam package is found
Hidden imports:  PySide6 (QtWidgets, QtCore, QtGui), openpyxl, PyPDF2
Collected:       openpyxl (all), PyPDF2 (all), pyxlsb (all)
Excluded:        PyQt5, PyQt6, tkinter, matplotlib, IPython, sphinx, pandas, numpy
Compression:     UPX enabled
```

Key points:
- The excludes list is critical for keeping the exe size down. Without them it balloons to ~150 MB.
- `pathex=['src']` is required so PyInstaller can find the `ihsdadam` package from the `main.py` entry point.
- Always use `--clean` flag to avoid stale cached analysis.

## Directory Structure

```
IHSDM2021/
├── main.py                          # PySide6 entry point
├── version.py                       # Version string + GitHub API config
├── requirements.txt                 # Python dependencies
├── IHSDadaM.spec                    # PyInstaller build spec
├── ihsdadam.ico                     # Application icon (multi-size)
│
├── src/ihsdadam/                    # PySide6 application package
│   ├── __init__.py
│   ├── app.py                       # Main window (IHSDadaMApp)
│   ├── models.py                    # Dataclasses: ResultMessage, AADTSection, CMFEntry
│   ├── theme.py                     # Windows 11 Fluent Design QSS + color constants
│   ├── workers.py                   # QThread workers for async operations
│   ├── report_engine.py             # HTML report generator (HNTB branding)
│   │
│   ├── tabs/                        # 9 functional tabs (sidebar navigation)
│   │   ├── __init__.py
│   │   ├── warning_tab.py           # Warning message extraction & filtering
│   │   ├── compiler_tab.py          # Excel data compilation
│   │   ├── appendix_tab.py          # PDF report merging
│   │   ├── visual_tab.py            # Highway alignment visualization
│   │   ├── cmf_tab.py               # Calibration factor scanner
│   │   ├── aadt_tab.py              # 5-step AADT input wizard
│   │   ├── eval_years_tab.py        # Evaluation years overview
│   │   ├── report_tab.py            # HTML crash prediction report generation
│   │   └── about_tab.py             # Workflow docs and credits
│   │
│   ├── dialogs/
│   │   ├── about_dialog.py          # Version info dialog
│   │   ├── update_dialog.py         # Auto-update notification
│   │   └── preview_dialog.py        # PDF preview
│   │
│   └── widgets/                     # Reusable UI components
│       ├── highway_canvas.py        # Lane visualization canvas
│       ├── scrollable_tree.py       # Tree widget with scrolling
│       ├── status_bar.py            # Status bar with progress
│       ├── search_bar.py            # Search/filter input
│       └── tooltip.py               # Tooltip helper
│
├── evaluations/                     # Sample report generation scripts
│   ├── generate_summary_report.py   # Multi-project KABCO comparison example
│   └── grant_ave_report.py          # Single-project FI/PDO example
│
└── _archived/                       # Old Tkinter files (not used in builds)
    ├── ihsdm_wisconsin_helper.py    # Legacy Tkinter monolith
    ├── ihsdm_compiler_core.py       # Legacy compiler engine
    ├── build_wisconsin_helper.bat   # Old Tkinter build script
    └── ...                          # Other legacy files
```

## Architecture

### Main Application (`src/ihsdadam/app.py`)

`IHSDadaMApp(QMainWindow)` - the main window with:
- **Header banner** (48px blue) with app name, clickable version label, email
- **Project path selector** bar with Browse button, propagates to all tabs
- **Sidebar** (190px QListWidget) for navigating between 9 tools
- **Stacked content** area (QStackedWidget) showing the active tab
- **Status bar** with progress indicator

Runtime-generated highway shield icon (blue rounded rect + white lane markings).

### Tab Pattern

Every tab follows this contract:
```python
class SomeTab(QWidget):
    status_message = Signal(str)        # Status bar updates
    progress_update = Signal(int, str)  # Progress bar (0-100, message)

    def set_project_path(self, path: str):  # Called when project path changes
        self._project_path = path
```

### The 9 Tabs

| # | Sidebar Label | Module | Purpose |
|---|--------------|--------|---------|
| 0 | Warning Extractor | `warning_tab.py` | Scan evaluation XMLs for warnings/errors/criticals |
| 1 | Data Compiler | `compiler_tab.py` | Extract crash predictions to Excel with HSM severity distributions |
| 2 | Appendix Generator | `appendix_tab.py` | Merge evaluation report PDFs into one document |
| 3 | Visual View | `visual_tab.py` | Interactive highway alignment geometry viewer |
| 4 | Evaluation Info | `cmf_tab.py` | Extract calibration factors (CMFs) from CSVs |
| 5 | AADT Input | `aadt_tab.py` | 5-step wizard for mapping forecast IDs to AADT values |
| 6 | Evaluation Years | `eval_years_tab.py` | Overview of year ranges across all alignments |
| 7 | Report Generation | `report_tab.py` | Generate HTML crash prediction reports from Data Compiler output |
| 8 | About | `about_tab.py` | Workflow docs, tab descriptions, credits |

### Report Generation (`report_tab.py` + `report_engine.py`)

Reads Data Compiler Excel output and produces printable HTML reports with HNTB branding.

- **Single project** or **multi-project comparison** modes
- **KABCO** (K, A, B, C, PDO) or **FI vs PDO** severity breakdowns
- Configurable highway segment grouping by functional class (Freeways, Arterials, Ramps, etc.)
- Report fields: title, subtitle, evaluation years, analyst name
- Optional logo, "Printed and reviewed by" footer
- Bar charts for multi-project comparison
- Metric cards and severity distribution bars

`report_engine.py` is adapted from the standalone `hntb_report.py` -- self-contained HTML generator, no external template files.

### Data Models (`models.py`)

```python
@dataclass ResultMessage     # Warning/error from evaluation XML
@dataclass AADTSection       # Station-range AADT entry from highway XML
@dataclass CMFEntry          # Calibration factor from CSV
```

**HSM Severity Distribution Constants:**
- Highway: K=1.46%, A=4.48%, B=24.69%, C=69.17%
- Intersection/Ramp: K=0.26%, A=5.35%, B=27.64%, C=66.75%

### Threading (`workers.py`)

All long-running operations use QThread workers to keep the UI responsive:
- `WarningScanWorker` - XML warning extraction
- `CompileWorker` - Excel data compilation
- `YearScanWorker` - Evaluation year discovery
- `AppendixMergeWorker` - PDF merging
- `VisualDataWorker` - Highway geometry loading
- `CMFScanWorker` - Calibration factor parsing
- `AADTScanWorker` - AADT section scanning
- `UpdateCheckWorker` - GitHub API version check

Workers emit `progress(int, str)`, `finished(data)`, and `error(str)` signals.

Helper functions in workers.py:
- `_folder_prefix(name)` - extracts alpha prefix from folder name (e.g., "h" from "h1", "ss" from "ss2")
- `_detect_eval_type_from_csv(eval_dir)` - detects i/r evaluations inside h-folders by scanning CSV for markers

### Theme System (`theme.py`)

Windows 11 Fluent Design implementation:
- Color palette: `PRIMARY=#005fb8`, `SURFACE=#ffffff`, `BACKGROUND=#f3f3f3`
- Status colors: `CRITICAL_BG/FG`, `ERROR_BG/FG`, `WARNING_BG/FG`
- Lane colors: `LANE_ASPHALT`, `CENTERLINE_YELLOW`, `SHOULDER_GRAY`
- Font: Segoe UI
- Full QSS stylesheet covering all widget types
- Accent buttons via `QPushButton[accent="true"]`

## IHSDM Project Structure

IHSDadaM operates on IHSDM project directories with this structure:

```
Project_Directory/
├── h1/, h2/, h74/          # Highway alignments (h prefix)
│   └── e1/, e2/, e20/      # Evaluation folders per alignment
│       ├── evaluation.1.result.xml     # Results (warnings, crash predictions)
│       ├── evaluation.1.cpm.cmf.csv    # Calibration factors
│       ├── evaluation.1.report.pdf     # Optional PDF report
│       └── highway.xml                 # Alignment geometry (h* only)
├── i1/, i2/, i100/          # Intersection alignments (i prefix)
├── ra1/, ra2/               # Roundabout alignments (ra prefix, treated as intersections)
├── r1/, r2/, r25/           # Ramp terminal alignments (r prefix)
├── ss1/, ss2/               # Site sets (ss prefix)
└── c1/, c2/                 # Interchange containers (c prefix, contain nested h/i/r/ra)
```

### Alignment Types
- `h*` = Highway segments (lane geometry, curves, speed, shoulders)
- `i*` = Intersections (major/minor AADT, intersection type)
- `ra*` = Roundabouts (treated as intersections in all tabs and compiler output)
- `r*` = Ramp terminals (similar to intersections)
- `ss*` = Site sets (aggregated analysis containers)
- `c*` = Interchange containers (group related h/i/r/ra alignments)

### IHSDM Highway Type Codes
Common functional class codes found in Data Compiler output:
- `2U` = Two-lane undivided, `3T` = Three-lane (TWLTL)
- `4U` = Four-lane undivided, `4D` = Four-lane divided, `4F` = Four-lane freeway
- `4SC`/`6SC` = Super-2/expressway, `6D` = Six-lane divided, `6F` = Six-lane freeway
- `8F` = Eight-lane freeway
- `1EN`/`1EX`/`2EN`/`2EX` = Ramp entrance/exit (1 or 2 lane)

### XML Namespace
IHSDM XML uses `{http://www.ihsdm.org/schema/Highway-1.0}`. The code tries both namespaced and non-namespaced lookups for compatibility.

### Lane/Shoulder Data Model
Lanes and shoulders use station-based ranges with priority stacking:
- `startStation`/`endStation` - Horizontal extent along alignment
- `sideOfRoad` - "left", "right", or "both" (mirrored)
- `priority` - Vertical stacking from centerline (P10 closest, P20 next out)
- `insideOutsideOfRoadNB` - Shoulder position ("inside" near median, "outside" at edge)

## Version Management

Update version in `version.py`, then tag and release:
```bash
git tag -a v1.x.x -m "Release message"
git push origin v1.x.x
```
The app checks GitHub API for updates on startup (2-second delayed background check).

## Dependencies

### Required
- `PySide6 >= 6.7, < 6.8` - Qt framework (UI)
- `openpyxl >= 3.0.0` - Excel export (Data Compiler, Report Generation)
- `PyPDF2 >= 3.0.0` - PDF merging (Appendix Generator)

### Build Only
- `pyinstaller >= 5.0` - Standalone executable creation
- `pyxlsb` - Legacy Excel binary format support

### Standard Library (no install)
- `xml.etree.ElementTree` - XML parsing
- `pathlib`, `os` - File system operations
- `csv` - CSV parsing
- `dataclasses` - Data models
- `re`, `collections`, `webbrowser`

## Development Notes

### Adding a New Tab

1. Create `src/ihsdadam/tabs/new_tab.py` following the tab pattern (QWidget with `status_message`, `progress_update` signals and `set_project_path` method)
2. Add import to `src/ihsdadam/tabs/__init__.py`
3. Import and instantiate in `app.py:_build_content_stack()`
4. Add sidebar entry in `app.py:_build_sidebar()`

### Key Code Patterns
- All async work goes through QThread workers in `workers.py`
- Accent buttons: `btn.setProperty("accent", True)`
- Trees use `ScrollableTree` widget wrapper
- Color constants are in `theme.py` (never hardcode colors)
- Station formatting: `_format_station()` in `workers.py` converts to "0+00.00" format
- Safe Excel cell access: `_cell(row, idx)` in `report_tab.py` prevents tuple index errors from short openpyxl rows
