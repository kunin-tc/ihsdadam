"""
IHSDadaM
Comprehensive tool for IHSDM project analysis and data compilation.

Features:
- Warning Message Extraction (supports highways, intersections, ramp terminals)
- Data Compilation to Excel (crash predictions with HSM severity distributions)

Author: Claude Code
Original IHSDM Compiler by: Adam Engbring (aengbring@hntb.com)
Date: 2025-12-17
Version: 1.0
"""

import os
import csv
import xml.etree.ElementTree as ET
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import List, Dict, Tuple
from dataclasses import dataclass, field
from collections import defaultdict, Counter
import subprocess
import time
import threading
import webbrowser
import json
from urllib.request import urlopen, Request
from urllib.error import URLError

# Import version info
try:
    from version import __version__, __app_name__, GITHUB_API_URL
except ImportError:
    __version__ = "1.0.0"
    __app_name__ = "IHSDM Wisconsin Helper"
    GITHUB_API_URL = None

try:
    from openpyxl import Workbook, load_workbook
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("Warning: openpyxl not found. Data Compiler tab will be disabled.")
    print("Install with: pip install openpyxl")


# =============================================================================
# DATA STRUCTURES
# =============================================================================

@dataclass
class ResultMessage:
    """Container for a result message from IHSDM evaluation"""
    alignment_type: str  # 'h' for highway, 'i' for intersection, 'r' for ramp terminal
    alignment_id: str    # e.g., 'h74', 'i100', 'r25'
    alignment_name: str  # e.g., 'Mainline - I-375 Centerline'
    evaluation: str      # e.g., 'e20'
    start_sta: str
    end_sta: str
    message: str
    status: str          # 'info', 'warning', 'error', 'fault', 'CRITICAL'
    file_path: str
    is_critical: bool = False  # True if contains "no crash prediction supported"


# =============================================================================
# HSM SEVERITY DISTRIBUTIONS
# =============================================================================

HSM_SEVERITY_K = 0.0146      # 1.46% fatal (highway)
HSM_SEVERITY_A = 0.044764    # 4.48% incapacitating injury (highway)
HSM_SEVERITY_B = 0.2469      # 24.69% non-incapacitating injury (highway)
HSM_SEVERITY_C = 0.69172     # 69.17% possible injury (highway)

INT_SEVERITY_K = 0.002575    # 0.26% fatal (intersection/ramp)
INT_SEVERITY_A = 0.053525    # 5.35% incapacitating injury (intersection/ramp)
INT_SEVERITY_B = 0.276415    # 27.64% non-incapacitating injury (intersection/ramp)
INT_SEVERITY_C = 0.667485    # 66.75% possible injury (intersection/ramp)


# =============================================================================
# MAIN APPLICATION CLASS
# =============================================================================

class IHSDMWisconsinHelper:
    """Main application class combining warning extraction and data compilation"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(f"{__app_name__} v{__version__}")
        self.root.geometry("1400x800")

        # Set color scheme
        self.colors = {
            'primary': '#1e3a8a',      # Deep blue
            'secondary': '#3b82f6',    # Bright blue
            'accent': '#10b981',       # Green
            'warning': '#f59e0b',      # Orange
            'danger': '#ef4444',       # Red
            'bg_light': '#f8fafc',     # Light gray
            'bg_dark': '#1e293b',      # Dark gray
            'text_light': '#ffffff',   # White
            'text_dark': '#1f2937',    # Dark text
        }

        # Configure root window
        self.root.configure(bg=self.colors['bg_light'])

        # Setup custom styles
        self.setup_styles()

        # Shared state
        self.project_path = tk.StringVar()
        self.messages: List[ResultMessage] = []
        self.filtered_messages: List[ResultMessage] = []

        # Compiler state
        self.target_file = tk.StringVar(value="evaluation.1.diag.csv")
        self.excel_output = tk.StringVar()
        self.multi_year_mode = tk.BooleanVar(value=False)  # Default to single-year mode
        self.debug_mode = tk.BooleanVar(value=False)

        self.setup_ui()

    def setup_styles(self):
        """Setup custom ttk styles for a modern look"""
        style = ttk.Style()

        # Use a modern theme as base
        try:
            style.theme_use('clam')  # Modern, customizable theme
        except:
            pass  # Fall back to default if clam not available

        # Configure notebook (tabs)
        style.configure('TNotebook', background=self.colors['bg_light'], borderwidth=0)
        style.configure('TNotebook.Tab',
                       padding=[15, 8],
                       font=('Segoe UI', 9),
                       background=self.colors['bg_dark'],
                       foreground=self.colors['text_light'])
        style.map('TNotebook.Tab',
                 background=[('selected', self.colors['primary'])],
                 foreground=[('selected', self.colors['text_light'])],
                 padding=[('selected', [20, 12])],
                 font=[('selected', ('Segoe UI', 11, 'bold'))],
                 expand=[('selected', [1, 1, 1, 0])])

        # Configure frames
        style.configure('TFrame', background=self.colors['bg_light'])
        style.configure('TLabelframe',
                       background=self.colors['bg_light'],
                       borderwidth=2,
                       relief='solid')
        style.configure('TLabelframe.Label',
                       background=self.colors['bg_light'],
                       foreground=self.colors['primary'],
                       font=('Segoe UI', 10, 'bold'))

        # Configure buttons
        style.configure('TButton',
                       font=('Segoe UI', 9),
                       borderwidth=1,
                       relief='flat',
                       padding=[10, 5])
        style.map('TButton',
                 background=[('active', self.colors['secondary'])],
                 relief=[('pressed', 'sunken')])

        # Accent button style (for primary actions)
        style.configure('Accent.TButton',
                       font=('Segoe UI', 10, 'bold'),
                       background=self.colors['accent'],
                       foreground=self.colors['text_light'],
                       borderwidth=0,
                       padding=[15, 8])
        style.map('Accent.TButton',
                 background=[('active', '#059669')],
                 foreground=[('active', self.colors['text_light'])])

        # Configure labels
        style.configure('TLabel',
                       background=self.colors['bg_light'],
                       font=('Segoe UI', 9))

        style.configure('Header.TLabel',
                       background=self.colors['bg_light'],
                       font=('Segoe UI', 12, 'bold'),
                       foreground=self.colors['primary'])

        # Configure entries
        style.configure('TEntry',
                       fieldbackground='white',
                       borderwidth=1,
                       relief='solid')

        # Configure combobox
        style.configure('TCombobox',
                       fieldbackground='white',
                       background='white',
                       borderwidth=1)

        # Configure treeview
        style.configure('Treeview',
                       background='white',
                       fieldbackground='white',
                       font=('Segoe UI', 9),
                       rowheight=45)  # Increased row height for multi-line messages
        style.configure('Treeview.Heading',
                       font=('Segoe UI', 9, 'bold'),
                       background=self.colors['primary'],
                       foreground=self.colors['text_light'],
                       borderwidth=1,
                       relief='flat')
        style.map('Treeview.Heading',
                 background=[('active', self.colors['secondary'])])

        # Configure checkbuttons
        style.configure('TCheckbutton',
                       background=self.colors['bg_light'],
                       font=('Segoe UI', 9))

    def setup_ui(self):
        """Setup the tabbed user interface"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Header banner
        header_frame = tk.Frame(main_frame, bg=self.colors['primary'], height=60)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        header_frame.grid_propagate(False)

        title_label = tk.Label(header_frame,
                               text="IHSDadaM",
                               font=('Segoe UI', 18, 'bold'),
                               bg=self.colors['primary'],
                               fg=self.colors['text_light'])
        title_label.pack(side=tk.LEFT, padx=20, pady=10)

        self.version_label = tk.Label(header_frame,
                                text=f"v{__version__}",
                                font=('Segoe UI', 10),
                                bg=self.colors['primary'],
                                fg=self.colors['text_light'],
                                cursor="hand2")
        self.version_label.pack(side=tk.LEFT, pady=10)
        self.version_label.bind("<Button-1>", lambda e: self.check_for_updates(show_current=True))

        subtitle_label = tk.Label(header_frame,
                                 text="Warning Extraction & Data Compilation",
                                 font=('Segoe UI', 10, 'italic'),
                                 bg=self.colors['primary'],
                                 fg=self.colors['text_light'])
        subtitle_label.pack(side=tk.RIGHT, padx=20, pady=10)

        # Project path selection (shared across tabs)
        path_frame = ttk.LabelFrame(main_frame, text="IHSDM Project Directory", padding="10")
        path_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        # Project path entry with placeholder
        self.path_entry = ttk.Entry(path_frame, textvariable=self.project_path, width=100)
        self.path_entry.grid(row=0, column=0, padx=(0, 10))

        # Setup placeholder text
        self.placeholder_text = r"C:\Users\aengbring\Downloads\ihsdm_17_0_0_full\IHSDM2021\users\aengbring\Projects_V5\p259"
        self.placeholder_active = True
        self.setup_placeholder()

        ttk.Button(path_frame, text="Browse...", command=self.browse_project).grid(row=0, column=1)

        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Tab 1: Warning Message Extractor
        self.warning_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.warning_tab, text="  Warning Extractor  ")
        self.setup_warning_tab()

        # Tab 2: Data Compiler
        self.compiler_tab = ttk.Frame(self.notebook)
        if OPENPYXL_AVAILABLE:
            self.notebook.add(self.compiler_tab, text="  Data Compiler  ")
            self.setup_compiler_tab()
        else:
            self.notebook.add(self.compiler_tab, text="  Data Compiler (Disabled)  ")
            self.setup_compiler_disabled_tab()

        # Tab 3: Appendix Generator
        self.appendix_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.appendix_tab, text="  Appendix Generator  ")
        self.setup_appendix_tab()

        # Tab 4: Visual View
        self.visual_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.visual_tab, text="  Visual View  ")
        self.setup_visual_tab()

        # Tab 5: CMF Scanner
        self.cmf_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.cmf_tab, text="  CMF Scanner  ")
        self.setup_cmf_tab()

        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # Status bar (shared) - enhanced styling
        status_frame = tk.Frame(self.root, bg=self.colors['bg_dark'], height=35)
        status_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))
        status_frame.grid_propagate(False)

        # Left side - status message
        self.status_var = tk.StringVar(value="Ready - Select a project directory above")
        status_bar = tk.Label(status_frame,
                             textvariable=self.status_var,
                             bg=self.colors['bg_dark'],
                             fg=self.colors['text_light'],
                             font=('Segoe UI', 9),
                             anchor=tk.W,
                             padx=15)
        status_bar.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Right side - credits and suggestions
        credits_text = "Created by Adam Engbring • aengbring@hntb.com • Message for suggestions/bugs"
        credits_label = tk.Label(status_frame,
                                text=credits_text,
                                bg=self.colors['bg_dark'],
                                fg=self.colors['text_light'],
                                font=('Segoe UI', 8),
                                anchor=tk.E,
                                padx=15)
        credits_label.pack(side=tk.RIGHT)

    # =========================================================================
    # WARNING EXTRACTOR TAB
    # =========================================================================

    def setup_warning_tab(self):
        """Setup the warning message extractor tab"""
        tab_frame = ttk.Frame(self.warning_tab, padding="10")
        tab_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.warning_tab.columnconfigure(0, weight=1)
        self.warning_tab.rowconfigure(0, weight=1)

        # Scan button
        scan_frame = ttk.Frame(tab_frame)
        scan_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Button(scan_frame, text="Scan for Warning Messages",
                  command=self.scan_warnings, style='Accent.TButton').pack(side=tk.LEFT, padx=5)

        # Summary - enhanced visual styling
        summary_frame = tk.Frame(tab_frame, bg='white', relief=tk.RIDGE, borderwidth=2)
        summary_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        summary_inner = tk.Frame(summary_frame, bg='white', padx=15, pady=10)
        summary_inner.pack(fill=tk.BOTH, expand=True)

        summary_title = tk.Label(summary_inner,
                                text="Scan Summary",
                                font=('Segoe UI', 10, 'bold'),
                                bg='white',
                                fg=self.colors['primary'])
        summary_title.pack(anchor=tk.W, pady=(0, 5))

        self.summary_text = tk.StringVar(value="No project scanned yet")
        summary_label = tk.Label(summary_inner,
                                textvariable=self.summary_text,
                                font=('Segoe UI', 9),
                                bg='white',
                                fg=self.colors['text_dark'],
                                justify=tk.LEFT)
        summary_label.pack(anchor=tk.W)

        # Filter panel
        filter_frame = ttk.LabelFrame(tab_frame, text="Filters", padding="10")
        filter_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))

        # Alignment Type filter
        ttk.Label(filter_frame, text="Type:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.alignment_type_var = tk.StringVar(value="All")
        alignment_type_combo = ttk.Combobox(filter_frame, textvariable=self.alignment_type_var,
                                           values=['All', 'Highways (h)', 'Intersections (i)', 'Ramp Terminals (r)'],
                                           state='readonly', width=30)
        alignment_type_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        alignment_type_combo.bind('<<ComboboxSelected>>', lambda e: self.apply_filters())

        # Specific Alignment filter
        ttk.Label(filter_frame, text="Alignment:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.alignment_var = tk.StringVar(value="All")
        self.alignment_combo = ttk.Combobox(filter_frame, textvariable=self.alignment_var, state='readonly', width=30)
        self.alignment_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        self.alignment_combo.bind('<<ComboboxSelected>>', lambda e: self.apply_filters())

        # Message type filter
        ttk.Label(filter_frame, text="Message Type:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.type_var = tk.StringVar(value="All")
        type_combo = ttk.Combobox(filter_frame, textvariable=self.type_var,
                                  values=['All', 'CRITICAL', 'error', 'warning', 'fault', 'info'],
                                  state='readonly', width=30)
        type_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        type_combo.bind('<<ComboboxSelected>>', lambda e: self.apply_filters())

        # Search filter
        ttk.Label(filter_frame, text="Search Text:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.search_var = tk.StringVar()
        self.search_var.trace('w', lambda *args: self.apply_filters())
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_var, width=30)
        search_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)

        ttk.Button(filter_frame, text="Clear Filters", command=self.clear_filters).grid(row=4, column=0, columnspan=2, pady=10)

        filter_frame.columnconfigure(1, weight=1)

        # Results display
        results_frame = ttk.LabelFrame(tab_frame, text="Messages", padding="10")
        results_frame.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))

        # Treeview for messages
        columns = ('Type', 'Alignment', 'Eval', 'Status', 'Start Sta', 'End Sta', 'Message')
        self.tree = ttk.Treeview(results_frame, columns=columns, show='tree headings', height=20)

        # Define column headings
        self.tree.heading('#0', text='')
        self.tree.heading('Type', text='Type')
        self.tree.heading('Alignment', text='Alignment')
        self.tree.heading('Eval', text='Eval')
        self.tree.heading('Status', text='Status')
        self.tree.heading('Start Sta', text='Start Station')
        self.tree.heading('End Sta', text='End Station')
        self.tree.heading('Message', text='Message')

        # Define column widths
        self.tree.column('#0', width=0, stretch=False)
        self.tree.column('Type', width=40, minwidth=40, stretch=False)
        self.tree.column('Alignment', width=250, minwidth=200)
        self.tree.column('Eval', width=50, minwidth=50, stretch=False)
        self.tree.column('Status', width=90, minwidth=80, stretch=False)
        self.tree.column('Start Sta', width=110, minwidth=100)
        self.tree.column('End Sta', width=110, minwidth=100)
        self.tree.column('Message', width=600, minwidth=400)

        # Scrollbars
        vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(results_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # Configure styling
        self.tree.tag_configure('critical', background='#ffcccc', foreground='#cc0000')
        self.tree.tag_configure('error', background='#ffe6e6', foreground='#cc0000')
        self.tree.tag_configure('warning', background='#fff4e6', foreground='#cc6600')

        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)

        # Action buttons
        button_frame = ttk.Frame(tab_frame)
        button_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))

        ttk.Button(button_frame, text="Export to CSV", command=self.export_warnings_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Copy Selected", command=self.copy_selected).pack(side=tk.LEFT, padx=5)

        # Configure grid weights
        tab_frame.columnconfigure(1, weight=2)
        tab_frame.columnconfigure(2, weight=1)
        tab_frame.rowconfigure(2, weight=1)

    # =========================================================================
    # DATA COMPILER TAB
    # =========================================================================

    def setup_compiler_tab(self):
        """Setup the data compiler tab"""
        tab_frame = ttk.Frame(self.compiler_tab, padding="10")
        tab_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.compiler_tab.columnconfigure(0, weight=1)
        self.compiler_tab.rowconfigure(0, weight=1)

        # Configuration section
        config_frame = ttk.LabelFrame(tab_frame, text="Compiler Configuration", padding="10")
        config_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        # Target file
        ttk.Label(config_frame, text="Target CSV File:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(config_frame, textvariable=self.target_file, width=40).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Label(config_frame, text="(default: evaluation.1.diag.csv)", font=('Arial', 8, 'italic')).grid(row=0, column=2, sticky=tk.W, padx=5)

        # Excel output
        ttk.Label(config_frame, text="Excel Output:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(config_frame, textvariable=self.excel_output, width=40).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(config_frame, text="Browse...", command=self.browse_excel_output).grid(row=1, column=2, padx=5, pady=5)

        # Options
        ttk.Checkbutton(config_frame, text="Multi-year extraction (20 years)",
                       variable=self.multi_year_mode).grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        ttk.Checkbutton(config_frame, text="Debug mode (verbose output)",
                       variable=self.debug_mode).grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        config_frame.columnconfigure(1, weight=1)

        # Compile button
        compile_frame = ttk.Frame(tab_frame)
        compile_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Button(compile_frame, text="Compile Data to Excel",
                  command=self.run_compiler, style='Accent.TButton').pack(side=tk.LEFT, padx=5)

        # Instructions
        instructions_frame = ttk.LabelFrame(tab_frame, text="Instructions & Information", padding="10")
        instructions_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        instructions_text = scrolledtext.ScrolledText(instructions_frame, wrap=tk.WORD, height=15, font=('Consolas', 9))
        instructions_text.pack(fill=tk.BOTH, expand=True)

        instructions_text.insert('1.0', """IHSDM DATA COMPILER - Instructions

This tool extracts crash prediction data from IHSDM evaluation files and compiles them into Excel.

WHAT IT DOES:
• Processes Highway Segments (h folders)
• Processes Intersections (i folders)
• Processes Ramp Terminals (r folders)
• Applies HSM severity distributions (K, A, B, C crashes)
• Removes duplicates
• Outputs to Excel workbook with separate sheets

FOLDER NAMING:
• h1, h2, h74, etc. → Highway alignments
• i1, i2, i100, etc. → Intersection alignments
• r1, r2, r25, etc. → Ramp terminal alignments
• c1, c2, etc. → Interchanges (can contain h/i/r subfolders)

ALIGNMENT NAMING BEST PRACTICES:
• Mainline freeways → Include "Mainline" in alignment name
• Ramps → Include "Entrance" or "Exit" in alignment name
• Ensure single functional class per alignment

OUTPUT:
• Excel file with 3 sheets: Highway, Intersection, RampTerminal
• Includes crash predictions with severity breakdown
• Single-year mode (default): Extracts one year of predictions
• Multi-year mode: Extracts 20 years of predictions (enable checkbox above)

HSM SEVERITY DISTRIBUTIONS:
• Highway: K=1.46%, A=4.48%, B=24.69%, C=69.17%
• Intersection/Ramp: K=0.26%, A=5.35%, B=27.64%, C=66.75%

STEPS:
1. Select project directory above (shared with Warning Extractor)
2. Set Excel output file path
3. Configure options (multi-year, debug)
4. Click "Compile Data to Excel"
5. Review output in Excel file

NOTE: Consult your state's predictive modeling practices for calibration factors.
Original script by Adam Engbring (aengbring@hntb.com)
""")
        instructions_text.config(state=tk.DISABLED)

        tab_frame.rowconfigure(2, weight=1)

    def setup_compiler_disabled_tab(self):
        """Show message when openpyxl is not available"""
        message_frame = ttk.Frame(self.compiler_tab, padding="50")
        message_frame.pack(expand=True)

        ttk.Label(message_frame, text="Data Compiler Unavailable",
                 font=('Arial', 14, 'bold')).pack(pady=10)
        ttk.Label(message_frame, text="The openpyxl library is required for the Data Compiler feature.",
                 font=('Arial', 10)).pack(pady=5)
        ttk.Label(message_frame, text="Install it using:", font=('Arial', 10)).pack(pady=5)
        ttk.Label(message_frame, text="pip install openpyxl",
                 font=('Consolas', 11, 'bold'), foreground='blue').pack(pady=10)
        ttk.Label(message_frame, text="Then restart this application.", font=('Arial', 10)).pack(pady=5)

    # =========================================================================
    # APPENDIX GENERATOR TAB
    # =========================================================================

    def setup_appendix_tab(self):
        """Setup the appendix generator tab"""
        tab_frame = ttk.Frame(self.appendix_tab, padding="10")
        tab_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.appendix_tab.columnconfigure(0, weight=1)
        self.appendix_tab.rowconfigure(0, weight=1)

        # Left panel - Alignment selection
        left_frame = ttk.Frame(tab_frame)
        left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))

        # Description
        desc_frame = ttk.LabelFrame(left_frame, text="Appendix Generator", padding="10")
        desc_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        desc_text = """Combines evaluation.1.report.pdf files into a single PDF.
Select which alignments to include below."""
        desc_label = tk.Label(desc_frame,
                             text=desc_text,
                             font=('Segoe UI', 10),
                             justify=tk.LEFT,
                             bg=self.colors['bg_light'])
        desc_label.pack(anchor=tk.W)

        # Alignment selection buttons
        btn_frame = ttk.Frame(left_frame)
        btn_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

        ttk.Button(btn_frame, text="Scan for Reports",
                  command=self.scan_appendix_alignments, style='Accent.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Select All",
                  command=self.appendix_select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Deselect All",
                  command=self.appendix_deselect_all).pack(side=tk.LEFT, padx=5)

        # Alignment selection frame with scrollbar
        alignment_frame = ttk.LabelFrame(left_frame, text="Select Reports to Include", padding="10")
        alignment_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Create canvas with scrollbar for alignment checkboxes
        appendix_canvas = tk.Canvas(alignment_frame, bg='white', highlightthickness=0, width=350)
        appendix_scrollbar = ttk.Scrollbar(alignment_frame, orient="vertical", command=appendix_canvas.yview)
        self.appendix_alignment_container = ttk.Frame(appendix_canvas)

        self.appendix_alignment_container.bind(
            "<Configure>",
            lambda e: appendix_canvas.configure(scrollregion=appendix_canvas.bbox("all"))
        )

        appendix_canvas.create_window((0, 0), window=self.appendix_alignment_container, anchor="nw")
        appendix_canvas.configure(yscrollcommand=appendix_scrollbar.set)

        appendix_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        appendix_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Bind mousewheel to scroll
        def _on_mousewheel(event):
            appendix_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        appendix_canvas.bind("<Enter>", lambda e: appendix_canvas.bind_all("<MouseWheel>", _on_mousewheel))
        appendix_canvas.bind("<Leave>", lambda e: appendix_canvas.unbind_all("<MouseWheel>"))

        # Initialize alignment tracking
        self.appendix_alignments = []  # List of discovered PDF files
        self.appendix_selected = {}  # Dict of path: BooleanVar

        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(2, weight=1)

        # Right panel - Output and log
        right_frame = ttk.Frame(tab_frame)
        right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))

        # Output file selection
        output_frame = ttk.LabelFrame(right_frame, text="Output PDF File", padding="10")
        output_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        self.appendix_output = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.appendix_output, width=50).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(output_frame, text="Browse...", command=self.browse_appendix_output).grid(row=0, column=1)

        output_frame.columnconfigure(0, weight=1)

        # Generate button
        button_frame = ttk.Frame(right_frame)
        button_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Button(button_frame, text="Generate Appendix PDF",
                  command=self.generate_appendix, style='Accent.TButton').pack(side=tk.LEFT, padx=5)

        # Status/Log area
        log_frame = ttk.LabelFrame(right_frame, text="Status", padding="10")
        log_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.appendix_log = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15, font=('Consolas', 9))
        self.appendix_log.pack(fill=tk.BOTH, expand=True)
        self.appendix_log.insert('1.0', "Click 'Scan for Reports' to find available evaluation reports.\n")
        self.appendix_log.config(state=tk.DISABLED)

        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(2, weight=1)

        # Configure main grid
        tab_frame.columnconfigure(0, weight=1)
        tab_frame.columnconfigure(1, weight=1)
        tab_frame.rowconfigure(0, weight=1)

    # =========================================================================
    # VISUAL VIEW TAB
    # =========================================================================

    def setup_visual_tab(self):
        """Setup the visual view tab"""
        tab_frame = ttk.Frame(self.visual_tab, padding="10")
        tab_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.visual_tab.columnconfigure(0, weight=1)
        self.visual_tab.rowconfigure(0, weight=1)

        # Description
        desc_frame = ttk.LabelFrame(tab_frame, text="Visual View", padding="10")
        desc_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        desc_text = """This tool displays highway alignment data visually from highway.1.xml files.
View horizontal and vertical alignment profiles for any highway alignment in your project."""
        desc_label = tk.Label(desc_frame,
                             text=desc_text,
                             font=('Segoe UI', 10),
                             justify=tk.LEFT,
                             bg=self.colors['bg_light'])
        desc_label.pack(anchor=tk.W)

        # Alignment selector
        selector_frame = ttk.LabelFrame(tab_frame, text="Select Alignment", padding="10")
        selector_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        # Load button on top row
        ttk.Button(selector_frame, text="Load Alignments",
                  command=self.refresh_visual_alignments).grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        # Alignment dropdown on second row
        ttk.Label(selector_frame, text="Highway Alignment:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.visual_alignment_var = tk.StringVar(value="No alignments found")
        self.visual_alignment_combo = ttk.Combobox(selector_frame, textvariable=self.visual_alignment_var,
                                                   state='readonly', width=60)
        self.visual_alignment_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        ttk.Button(selector_frame, text="Display Alignment",
                  command=self.display_alignment, style='Accent.TButton').grid(row=1, column=2, padx=5, pady=5)

        selector_frame.columnconfigure(1, weight=1)

        # Visualization area (placeholder for now)
        viz_frame = ttk.LabelFrame(tab_frame, text="Alignment Visualization", padding="10")
        viz_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.visual_canvas_frame = tk.Frame(viz_frame, bg='white', relief=tk.SUNKEN, borderwidth=2)
        self.visual_canvas_frame.pack(fill=tk.BOTH, expand=True)

        # Placeholder text
        placeholder_label = tk.Label(self.visual_canvas_frame,
                                     text="Select an alignment and click 'Display Alignment' to view visualization",
                                     font=('Segoe UI', 12),
                                     bg='white',
                                     fg='gray')
        placeholder_label.pack(expand=True)

        tab_frame.columnconfigure(0, weight=1)
        tab_frame.rowconfigure(2, weight=1)

    # =========================================================================
    # SHARED FUNCTIONS
    # =========================================================================

    def setup_placeholder(self):
        """Setup placeholder text for project path entry"""
        # Insert placeholder initially
        self.path_entry.insert(0, self.placeholder_text)
        self.path_entry.config(foreground='gray')

        # Bind focus events
        self.path_entry.bind('<FocusIn>', self.on_path_entry_focus_in)
        self.path_entry.bind('<FocusOut>', self.on_path_entry_focus_out)

    def on_path_entry_focus_in(self, event):
        """Handle focus in - auto-fill with default if placeholder is active"""
        if self.placeholder_active and self.project_path.get() == self.placeholder_text:
            # Keep the text but change to black (auto-fill behavior)
            self.path_entry.config(foreground='black')
            self.placeholder_active = False

    def on_path_entry_focus_out(self, event):
        """Handle focus out - restore placeholder if empty"""
        if not self.project_path.get():
            self.project_path.set(self.placeholder_text)
            self.path_entry.config(foreground='gray')
            self.placeholder_active = True

    def format_station(self, station_str: str) -> str:
        """Format station number with + notation (e.g., 400.00 → 4+00.00, 50.00 → 50.00)"""
        if not station_str or station_str.strip() == '':
            return station_str

        try:
            station_float = float(station_str)

            # If less than 100, no + is needed
            if station_float < 100:
                return station_str

            # Convert to string and split at decimal
            station_parts = str(station_float).split('.')
            integer_part = station_parts[0]
            decimal_part = station_parts[1] if len(station_parts) > 1 else '00'

            # Ensure decimal part has at least 2 digits
            decimal_part = decimal_part.ljust(2, '0')

            # Insert + two digits from the right of integer part
            if len(integer_part) >= 2:
                left_part = integer_part[:-2]
                right_part = integer_part[-2:]
                return f"{left_part}+{right_part}.{decimal_part}"
            else:
                # Edge case: integer part has only 1 digit
                return f"{integer_part}.{decimal_part}"

        except (ValueError, TypeError):
            return station_str

    def wrap_text(self, text: str, width: int = 60) -> str:
        """Wrap text to fit in treeview cell with line breaks"""
        if not text or len(text) <= width:
            return text

        words = text.split()
        lines = []
        current_line = []
        current_length = 0

        for word in words:
            word_length = len(word)
            if current_length + word_length + len(current_line) <= width:
                current_line.append(word)
                current_length += word_length
            else:
                if current_line:
                    lines.append(' '.join(current_line))
                current_line = [word]
                current_length = word_length

        if current_line:
            lines.append(' '.join(current_line))

        return '\n'.join(lines)

    def browse_project(self):
        """Open directory browser dialog for project selection"""
        # Get initial directory, avoiding placeholder text
        initial_dir = self.project_path.get() if not self.placeholder_active else os.path.expanduser("~")

        directory = filedialog.askdirectory(
            title="Select IHSDM Project Directory",
            initialdir=initial_dir
        )
        if directory:
            self.project_path.set(directory)
            self.path_entry.config(foreground='black')
            self.placeholder_active = False
            self.status_var.set(f"Project directory set: {directory}")

    def browse_excel_output(self):
        """Open file browser for Excel output"""
        file_path = filedialog.asksaveasfilename(
            title="Select Excel Output File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_output.set(file_path)

    # =========================================================================
    # WARNING EXTRACTOR FUNCTIONS
    # =========================================================================

    def scan_warnings(self):
        """Scan project for warning messages"""
        project_path = self.project_path.get()

        # Check if placeholder is active
        if self.placeholder_active or not project_path or not os.path.isdir(project_path):
            messagebox.showerror("Error", "Please select a valid project directory")
            return

        self.status_var.set("Scanning project for warnings...")
        self.root.update()

        try:
            self.messages = []
            project_dir = Path(project_path)
            alignment_dirs = []

            # Find direct alignment directories (h*, i*, r*)
            direct_alignments = [d for d in project_dir.iterdir()
                                if d.is_dir() and d.name[0].lower() in ['h', 'i', 'r']]
            alignment_dirs.extend(direct_alignments)

            # Find interchange directories (c*) and their nested alignments
            interchange_dirs = [d for d in project_dir.iterdir()
                               if d.is_dir() and d.name[0].lower() == 'c']

            for interchange_dir in interchange_dirs:
                # Inside each interchange, look for h*, i*, r* subdirectories
                nested_alignments = [d for d in interchange_dir.iterdir()
                                    if d.is_dir() and d.name[0].lower() in ['h', 'i', 'r']]
                alignment_dirs.extend(nested_alignments)

            total_dirs = len(alignment_dirs)
            processed = 0

            for alignment_dir in alignment_dirs:
                processed += 1
                self.status_var.set(f"Scanning {alignment_dir.name}... ({processed}/{total_dirs})")
                self.root.update()

                # Get alignment name
                alignment_name = self.get_alignment_name(alignment_dir)

                # Find all evaluation directories (e*)
                eval_dirs = [d for d in alignment_dir.iterdir()
                           if d.is_dir() and d.name.startswith('e')]

                for eval_dir in eval_dirs:
                    # Look for evaluation result XML files
                    result_files = list(eval_dir.glob("evaluation.*.result.xml"))

                    for result_file in result_files:
                        messages = self.parse_result_file(
                            result_file,
                            alignment_dir.name,
                            alignment_name,
                            eval_dir.name
                        )
                        self.messages.extend(messages)

            # Count critical messages
            critical_count = sum(1 for msg in self.messages if msg.is_critical)

            self.update_summary()
            self.update_alignment_filter()
            self.apply_filters()

            self.status_var.set(f"Scan complete. Found {len(self.messages)} messages ({critical_count} critical).")

            msg_text = f"Found {len(self.messages)} messages across {total_dirs} alignments"
            if critical_count > 0:
                msg_text += f"\n\nWARNING: {critical_count} CRITICAL messages found!\n(Messages containing 'no crash prediction supported')"
                messagebox.showwarning("Scan Complete - Critical Messages Found", msg_text)
            else:
                messagebox.showinfo("Scan Complete", msg_text)

        except Exception as e:
            messagebox.showerror("Error", f"Error scanning project: {str(e)}")
            self.status_var.set("Error during scan")

    def get_alignment_name(self, alignment_dir: Path) -> str:
        """Extract alignment name from XML files"""
        try:
            # Try highway.xml
            highway_file = alignment_dir / "highway.xml"
            if highway_file.exists():
                tree = ET.parse(highway_file)
                root = tree.getroot()
                return root.get('title', alignment_dir.name)

            # Try intersection.xml
            intersection_file = alignment_dir / "intersection.xml"
            if intersection_file.exists():
                tree = ET.parse(intersection_file)
                root = tree.getroot()
                return root.get('title', alignment_dir.name)

            # Try rampterminal.xml
            ramp_file = alignment_dir / "rampterminal.xml"
            if ramp_file.exists():
                tree = ET.parse(ramp_file)
                root = tree.getroot()
                return root.get('title', alignment_dir.name)

        except Exception:
            pass

        return alignment_dir.name

    def parse_result_file(self, file_path: Path, alignment_id: str,
                         alignment_name: str, evaluation: str) -> List[ResultMessage]:
        """Parse evaluation result XML file for ResultMessage elements"""
        messages = []

        try:
            tree = ET.parse(file_path)
            root = tree.getroot()

            for msg_elem in root.iter('ResultMessage'):
                start_sta = msg_elem.get('startSta', '')
                end_sta = msg_elem.get('endSta', '')
                message = msg_elem.get('message', '')
                status = msg_elem.get('ResultMessage.status', 'info')

                alignment_type = alignment_id[0].lower()  # 'h', 'i', or 'r'

                # Check for critical phrase
                is_critical = 'no crash prediction supported' in message.lower()
                if is_critical:
                    status = 'CRITICAL'

                msg = ResultMessage(
                    alignment_type=alignment_type,
                    alignment_id=alignment_id,
                    alignment_name=alignment_name,
                    evaluation=evaluation,
                    start_sta=start_sta,
                    end_sta=end_sta,
                    message=message,
                    status=status,
                    file_path=str(file_path),
                    is_critical=is_critical
                )
                messages.append(msg)

        except Exception as e:
            print(f"Error parsing {file_path}: {e}")

        return messages

    def update_summary(self):
        """Update the summary text"""
        if not self.messages:
            self.summary_text.set("No messages found")
            return

        type_counts = defaultdict(int)
        highway_alignments = set()
        intersection_alignments = set()
        ramp_alignments = set()

        for msg in self.messages:
            type_counts[msg.status] += 1
            alignment_str = f"{msg.alignment_id} - {msg.alignment_name}"

            if msg.alignment_type == 'h':
                highway_alignments.add(alignment_str)
            elif msg.alignment_type == 'i':
                intersection_alignments.add(alignment_str)
            elif msg.alignment_type == 'r':
                ramp_alignments.add(alignment_str)

        summary = f"Total: {len(self.messages)} | "
        summary += f"Highways: {len(highway_alignments)} | "
        summary += f"Intersections: {len(intersection_alignments)} | "
        summary += f"Ramp Terminals: {len(ramp_alignments)} | "

        if type_counts['CRITICAL'] > 0:
            summary += f"CRITICAL: {type_counts['CRITICAL']} | "

        summary += f"Errors: {type_counts['error']} | "
        summary += f"Warnings: {type_counts['warning']} | "
        summary += f"Info: {type_counts['info']}"

        self.summary_text.set(summary)

    def update_alignment_filter(self):
        """Update alignment filter dropdown"""
        highway_alignments = set()
        intersection_alignments = set()
        ramp_alignments = set()

        for msg in self.messages:
            alignment_str = f"{msg.alignment_id} - {msg.alignment_name}"
            if msg.alignment_type == 'h':
                highway_alignments.add(alignment_str)
            elif msg.alignment_type == 'i':
                intersection_alignments.add(alignment_str)
            elif msg.alignment_type == 'r':
                ramp_alignments.add(alignment_str)

        values = ['All']

        if highway_alignments:
            values.append('--- HIGHWAYS ---')
            values.extend(sorted(highway_alignments))

        if intersection_alignments:
            values.append('--- INTERSECTIONS ---')
            values.extend(sorted(intersection_alignments))

        if ramp_alignments:
            values.append('--- RAMP TERMINALS ---')
            values.extend(sorted(ramp_alignments))

        self.alignment_combo['values'] = values
        self.alignment_combo.set('All')

    def apply_filters(self):
        """Apply filters and update treeview"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        alignment_type_filter = self.alignment_type_var.get()
        alignment_filter = self.alignment_var.get()
        type_filter = self.type_var.get()
        search_text = self.search_var.get().lower()

        self.filtered_messages = []

        for msg in self.messages:
            # Alignment type filter
            if alignment_type_filter == 'Highways (h)' and msg.alignment_type != 'h':
                continue
            elif alignment_type_filter == 'Intersections (i)' and msg.alignment_type != 'i':
                continue
            elif alignment_type_filter == 'Ramp Terminals (r)' and msg.alignment_type != 'r':
                continue

            # Specific alignment filter
            if alignment_filter not in ['All', '--- HIGHWAYS ---', '--- INTERSECTIONS ---', '--- RAMP TERMINALS ---']:
                msg_alignment = f"{msg.alignment_id} - {msg.alignment_name}"
                if msg_alignment != alignment_filter:
                    continue

            # Type filter
            if type_filter != 'All' and msg.status != type_filter:
                continue

            # Search filter
            if search_text:
                searchable = f"{msg.message} {msg.alignment_name}".lower()
                if search_text not in searchable:
                    continue

            self.filtered_messages.append(msg)

        # Group by type
        highways = defaultdict(list)
        intersections = defaultdict(list)
        ramps = defaultdict(list)

        for msg in self.filtered_messages:
            key = f"{msg.alignment_id} - {msg.alignment_name}"
            if msg.alignment_type == 'h':
                highways[key].append(msg)
            elif msg.alignment_type == 'i':
                intersections[key].append(msg)
            elif msg.alignment_type == 'r':
                ramps[key].append(msg)

        # Populate tree
        self.populate_tree_section('HIGHWAYS', highways, 'h')
        self.populate_tree_section('INTERSECTIONS', intersections, 'i')
        self.populate_tree_section('RAMP TERMINALS', ramps, 'r')

        critical_count = sum(1 for msg in self.filtered_messages if msg.is_critical)
        status_msg = f"Showing {len(self.filtered_messages)} of {len(self.messages)} messages"
        if critical_count > 0:
            status_msg += f" ({critical_count} CRITICAL)"
        self.status_var.set(status_msg)

    def populate_tree_section(self, section_name, alignments_dict, type_char):
        """Populate a section of the tree view"""
        if not alignments_dict:
            return

        section_parent = self.tree.insert('', 'end', text=section_name,
                                         values=('', f'({len(alignments_dict)} alignments)', '', '', '', '', ''),
                                         open=True)

        for alignment, msgs in sorted(alignments_dict.items()):
            critical_count = sum(1 for m in msgs if m.is_critical)
            error_count = sum(1 for m in msgs if m.status == 'error')
            warning_count = sum(1 for m in msgs if m.status == 'warning')

            label = f"{alignment} ({len(msgs)} msgs"
            if critical_count > 0:
                label += f", {critical_count} CRITICAL"
            if error_count > 0:
                label += f", {error_count} errors"
            if warning_count > 0:
                label += f", {warning_count} warnings"
            label += ")"

            parent = self.tree.insert(section_parent, 'end', text=alignment,
                                    values=(type_char, alignment, '', '', '', '', label))

            for msg in sorted(msgs, key=lambda x: (x.evaluation, x.start_sta)):
                tags = []
                if msg.is_critical:
                    tags.append('critical')
                elif msg.status == 'error':
                    tags.append('error')
                elif msg.status == 'warning':
                    tags.append('warning')

                # Format station numbers with + notation
                formatted_start = self.format_station(msg.start_sta)
                formatted_end = self.format_station(msg.end_sta)

                # Wrap message text for multi-line display (wider wrap for readability)
                wrapped_message = self.wrap_text(msg.message, width=90)

                self.tree.insert(parent, 'end',
                               values=('', '', msg.evaluation, msg.status,
                                      formatted_start, formatted_end, wrapped_message),
                               tags=tags)

    def clear_filters(self):
        """Clear all filters"""
        self.alignment_type_var.set('All')
        self.alignment_var.set('All')
        self.type_var.set('All')
        self.search_var.set('')
        self.apply_filters()

    def export_warnings_csv(self):
        """Export filtered messages to CSV"""
        if not self.filtered_messages:
            messagebox.showwarning("No Data", "No messages to export")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Export Messages to CSV"
        )

        if not file_path:
            return

        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['Type', 'Alignment ID', 'Alignment Name', 'Evaluation',
                               'Status', 'Is Critical', 'Start Station', 'End Station', 'Message', 'File Path'])

                for msg in self.filtered_messages:
                    type_name = {'h': 'Highway', 'i': 'Intersection', 'r': 'Ramp Terminal'}[msg.alignment_type]
                    writer.writerow([
                        type_name,
                        msg.alignment_id,
                        msg.alignment_name,
                        msg.evaluation,
                        msg.status,
                        'YES' if msg.is_critical else 'NO',
                        msg.start_sta,
                        msg.end_sta,
                        msg.message,
                        msg.file_path
                    ])

            messagebox.showinfo("Export Complete", f"Exported {len(self.filtered_messages)} messages to CSV")
            self.status_var.set(f"Exported to {file_path}")

        except Exception as e:
            messagebox.showerror("Export Error", f"Error exporting to CSV: {str(e)}")

    def copy_selected(self):
        """Copy selected messages to clipboard"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select messages to copy")
            return

        text_lines = []
        for item in selection:
            values = self.tree.item(item)['values']
            if len(values) >= 6:
                text_lines.append('\t'.join(str(v) for v in values))

        if text_lines:
            self.root.clipboard_clear()
            self.root.clipboard_append('\n'.join(text_lines))
            self.status_var.set(f"Copied {len(text_lines)} messages to clipboard")
        else:
            messagebox.showinfo("No Data", "No message data to copy")

    # =========================================================================
    # DATA COMPILER FUNCTIONS (stub - will add full implementation)
    # =========================================================================

    def run_compiler(self):
        """Run the data compiler"""
        import ihsdm_compiler_core as compiler

        project_path = self.project_path.get()
        excel_path = self.excel_output.get()
        target_file = self.target_file.get()
        multi_year = self.multi_year_mode.get()
        debug = self.debug_mode.get()

        # Validation
        if self.placeholder_active or not project_path or not os.path.isdir(project_path):
            messagebox.showerror("Error", "Please select a valid project directory")
            return

        if not excel_path:
            messagebox.showerror("Error", "Please select an Excel output file")
            return

        self.status_var.set("Compiling data...")
        self.root.update()

        try:
            # Find folders
            parent_folders = compiler.find_folders_with_file(project_path, target_file)

            if not parent_folders:
                messagebox.showerror("Error", f"No folders containing '{target_file}' found in project directory")
                self.status_var.set("Error: No target files found")
                return

            # Process highways
            h_folders = [f for f in parent_folders if 'h' in os.path.basename(f).lower()]
            all_highway_rows = []

            for parent_folder in h_folders:
                for subfolder in os.listdir(parent_folder):
                    subfolder_path = os.path.join(parent_folder, subfolder)
                    if not os.path.isdir(subfolder_path):
                        continue
                    file_path = os.path.join(subfolder_path, target_file)

                    if os.path.isfile(file_path):
                        rows = compiler.extract_highway_segments_from_csv(file_path, start_index=5)
                        for row in rows:
                            if compiler.should_process_highway_row(row):
                                filtered_row = compiler.extract_highway_row_data(row, debug)
                                if filtered_row:
                                    all_highway_rows.append(filtered_row)

            # Remove duplicates and write
            unique_highway_rows, _ = compiler.remove_duplicates(all_highway_rows)
            unique_highway_rows.sort(key=lambda x: (x[0], x[1], x[2]), reverse=False)
            compiler.write_rows_to_excel(unique_highway_rows, excel_path, "Highway")
            compiler.add_header_to_excel(excel_path, "Highway", compiler.HIGHWAY_HEADER)
            compiler.fill_missing_highway_values(excel_path)

            # Process intersections
            i_folders = [f for f in parent_folders if 'i' in os.path.basename(f).lower()]
            all_intersection_rows = []
            first_file = False

            for parent_folder in i_folders:
                for subfolder in os.listdir(parent_folder):
                    subfolder_path = os.path.join(parent_folder, subfolder)
                    if not os.path.isdir(subfolder_path):
                        continue
                    file_path = os.path.join(subfolder_path, target_file)

                    if os.path.isfile(file_path):
                        rows = compiler.extract_by_headers_from_csv(
                            file_path, compiler.INTERSECTION_HEADER[:-1] + ["Fatal and Injury (FI) Crashes"],
                            first_file=not first_file, multi_year=multi_year
                        )
                        if rows:
                            all_intersection_rows.extend(rows)
                            first_file = True

            if all_intersection_rows:
                compiler.write_rows_to_excel(all_intersection_rows, excel_path, "Intersection")
                compiler.fill_missing_intersection_values(excel_path)
                compiler.scrub_duplicate_columns(excel_path, "Intersection")

            # Process ramp terminals
            r_folders = [f for f in parent_folders if 'r' in os.path.basename(f).lower()]
            all_ramp_rows = []
            first_file = False

            for parent_folder in r_folders:
                for subfolder in os.listdir(parent_folder):
                    subfolder_path = os.path.join(parent_folder, subfolder)
                    if not os.path.isdir(subfolder_path):
                        continue
                    file_path = os.path.join(subfolder_path, target_file)

                    if os.path.isfile(file_path):
                        rows = compiler.extract_by_headers_from_csv(
                            file_path, compiler.RAMP_TERMINAL_HEADER[:-1] + ["Fatal and Injury (FI) Crashes"],
                            first_file=not first_file, multi_year=multi_year
                        )
                        if rows:
                            all_ramp_rows.extend(rows)
                            first_file = True

            if all_ramp_rows:
                compiler.write_rows_to_excel(all_ramp_rows, excel_path, "RampTerminal")
                compiler.fill_missing_ramp_terminal_values(excel_path)
                compiler.scrub_duplicate_columns(excel_path, "RampTerminal")

            self.status_var.set("Compilation complete!")

            summary = f"Compilation Complete!\n\n"
            summary += f"Results written to:\n{excel_path}\n\n"
            summary += f"Summary:\n"
            summary += f"  - Highway Segments: {len(unique_highway_rows)} rows\n"
            summary += f"  - Intersections: {len(all_intersection_rows)} rows\n"
            summary += f"  - Ramp Terminals: {len(all_ramp_rows)} rows"

            messagebox.showinfo("Success", summary)

        except Exception as e:
            messagebox.showerror("Error", f"Error during compilation: {str(e)}")
            self.status_var.set("Error during compilation")
            import traceback
            traceback.print_exc()

    # =========================================================================
    # APPENDIX GENERATOR FUNCTIONS
    # =========================================================================

    def browse_appendix_output(self):
        """Open file browser for appendix PDF output"""
        file_path = filedialog.asksaveasfilename(
            title="Select Output PDF File",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.appendix_output.set(file_path)

    def log_appendix(self, message):
        """Add message to appendix log"""
        self.appendix_log.config(state=tk.NORMAL)
        self.appendix_log.insert(tk.END, message + '\n')
        self.appendix_log.see(tk.END)
        self.appendix_log.config(state=tk.DISABLED)
        self.root.update()

    def scan_appendix_alignments(self):
        """Scan project for all evaluation report PDFs"""
        project_path = self.project_path.get()

        if self.placeholder_active or not project_path or not os.path.isdir(project_path):
            messagebox.showerror("Error", "Please select a valid project directory")
            return

        self.status_var.set("Scanning for evaluation reports...")
        self.appendix_log.config(state=tk.NORMAL)
        self.appendix_log.delete('1.0', tk.END)
        self.appendix_log.config(state=tk.DISABLED)

        self.log_appendix(f"Scanning project: {project_path}")
        self.log_appendix("")

        try:
            project_dir = Path(project_path)
            pdf_files = list(project_dir.glob("**/evaluation.1.report.pdf"))

            if not pdf_files:
                self.log_appendix("No evaluation.1.report.pdf files found.")
                messagebox.showwarning("No Files", "No evaluation report PDFs found in the project.")
                return

            # Sort PDFs and organize by alignment
            self.appendix_alignments = sorted(pdf_files)

            self.log_appendix(f"Found {len(self.appendix_alignments)} evaluation reports:")

            # Clear existing checkboxes
            for widget in self.appendix_alignment_container.winfo_children():
                widget.destroy()

            self.appendix_selected = {}

            # Group by type (h=highway, i=intersection, r=ramp)
            highways = []
            intersections = []
            ramps = []
            other = []

            for pdf_file in self.appendix_alignments:
                rel_path = pdf_file.relative_to(project_dir)
                parts = rel_path.parts

                # Get alignment folder name (e.g., h1, i2, r3)
                alignment_folder = None
                for part in parts:
                    if part and part[0].lower() in ['h', 'i', 'r', 'c']:
                        alignment_folder = part
                        break

                # Get alignment name from highway.xml if possible
                alignment_name = self.get_alignment_name_from_pdf_path(pdf_file)
                display_name = f"{'/'.join(parts[:-2])} - {alignment_name}" if alignment_name else str(rel_path.parent.parent)

                entry = {'path': pdf_file, 'display': display_name, 'rel_path': rel_path}

                if alignment_folder and alignment_folder[0].lower() == 'h':
                    highways.append(entry)
                elif alignment_folder and alignment_folder[0].lower() == 'i':
                    intersections.append(entry)
                elif alignment_folder and alignment_folder[0].lower() == 'r':
                    ramps.append(entry)
                else:
                    other.append(entry)

            row = 0

            # Highways section
            if highways:
                header = ttk.Label(self.appendix_alignment_container, text="HIGHWAYS",
                                 font=('Segoe UI', 10, 'bold'), foreground=self.colors['primary'])
                header.grid(row=row, column=0, sticky=tk.W, pady=(5, 2))
                row += 1

                for entry in highways:
                    var = tk.BooleanVar(value=True)
                    self.appendix_selected[str(entry['path'])] = var
                    cb = ttk.Checkbutton(self.appendix_alignment_container,
                                        text=entry['display'],
                                        variable=var)
                    cb.grid(row=row, column=0, sticky=tk.W, padx=10)
                    row += 1
                    self.log_appendix(f"  [H] {entry['display']}")

            # Intersections section
            if intersections:
                header = ttk.Label(self.appendix_alignment_container, text="INTERSECTIONS",
                                 font=('Segoe UI', 10, 'bold'), foreground=self.colors['primary'])
                header.grid(row=row, column=0, sticky=tk.W, pady=(10, 2))
                row += 1

                for entry in intersections:
                    var = tk.BooleanVar(value=True)
                    self.appendix_selected[str(entry['path'])] = var
                    cb = ttk.Checkbutton(self.appendix_alignment_container,
                                        text=entry['display'],
                                        variable=var)
                    cb.grid(row=row, column=0, sticky=tk.W, padx=10)
                    row += 1
                    self.log_appendix(f"  [I] {entry['display']}")

            # Ramps section
            if ramps:
                header = ttk.Label(self.appendix_alignment_container, text="RAMP TERMINALS",
                                 font=('Segoe UI', 10, 'bold'), foreground=self.colors['primary'])
                header.grid(row=row, column=0, sticky=tk.W, pady=(10, 2))
                row += 1

                for entry in ramps:
                    var = tk.BooleanVar(value=True)
                    self.appendix_selected[str(entry['path'])] = var
                    cb = ttk.Checkbutton(self.appendix_alignment_container,
                                        text=entry['display'],
                                        variable=var)
                    cb.grid(row=row, column=0, sticky=tk.W, padx=10)
                    row += 1
                    self.log_appendix(f"  [R] {entry['display']}")

            # Other section
            if other:
                header = ttk.Label(self.appendix_alignment_container, text="OTHER",
                                 font=('Segoe UI', 10, 'bold'), foreground=self.colors['primary'])
                header.grid(row=row, column=0, sticky=tk.W, pady=(10, 2))
                row += 1

                for entry in other:
                    var = tk.BooleanVar(value=True)
                    self.appendix_selected[str(entry['path'])] = var
                    cb = ttk.Checkbutton(self.appendix_alignment_container,
                                        text=entry['display'],
                                        variable=var)
                    cb.grid(row=row, column=0, sticky=tk.W, padx=10)
                    row += 1
                    self.log_appendix(f"  [?] {entry['display']}")

            self.log_appendix("")
            self.log_appendix(f"All {len(self.appendix_alignments)} reports selected by default.")
            self.status_var.set(f"Found {len(self.appendix_alignments)} evaluation reports")

        except Exception as e:
            self.log_appendix(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Error scanning for reports: {str(e)}")

    def get_alignment_name_from_pdf_path(self, pdf_path):
        """Try to get alignment name from highway.xml near the PDF"""
        try:
            # PDF is in e*/evaluation.1.report.pdf, alignment is parent of e*
            alignment_dir = pdf_path.parent.parent
            # Look for highway.1.xml or similar
            for xml_file in alignment_dir.glob("*.1.xml"):
                if 'highway' in xml_file.name or 'intersection' in xml_file.name or 'ramp' in xml_file.name:
                    return self.get_alignment_name(alignment_dir)
            return self.get_alignment_name(alignment_dir)
        except:
            return None

    def appendix_select_all(self):
        """Select all alignments for appendix"""
        for var in self.appendix_selected.values():
            var.set(True)

    def appendix_deselect_all(self):
        """Deselect all alignments for appendix"""
        for var in self.appendix_selected.values():
            var.set(False)

    def generate_appendix(self):
        """Generate combined PDF from selected evaluation.1.report.pdf files"""
        project_path = self.project_path.get()
        output_path = self.appendix_output.get()

        # Validation
        if self.placeholder_active or not project_path or not os.path.isdir(project_path):
            messagebox.showerror("Error", "Please select a valid project directory")
            return

        if not output_path:
            messagebox.showerror("Error", "Please select an output PDF file")
            return

        # Check for PyPDF2
        try:
            from PyPDF2 import PdfMerger
        except ImportError:
            msg = "PyPDF2 library is required for PDF merging.\n\nInstall with: pip install PyPDF2"
            messagebox.showerror("Missing Dependency", msg)
            self.log_appendix("ERROR: PyPDF2 not installed. Run: pip install PyPDF2")
            return

        self.status_var.set("Generating appendix...")
        self.appendix_log.config(state=tk.NORMAL)
        self.appendix_log.delete('1.0', tk.END)
        self.appendix_log.config(state=tk.DISABLED)

        self.log_appendix("Starting appendix generation...")
        self.log_appendix(f"Project: {project_path}")
        self.log_appendix(f"Output: {output_path}")
        self.log_appendix("")

        try:
            project_dir = Path(project_path)

            # Get selected PDFs from checkbox selection, or scan if not done yet
            if self.appendix_selected:
                pdf_files = [Path(path) for path, var in self.appendix_selected.items() if var.get()]
            else:
                # Fallback: scan for all PDFs if user didn't scan first
                pdf_files = list(project_dir.glob("**/evaluation.1.report.pdf"))

            if not pdf_files:
                self.log_appendix("No evaluation reports selected or found.")
                messagebox.showwarning("No Files", "No evaluation reports selected. Please scan and select reports first.")
                self.status_var.set("No PDF files selected")
                return

            self.log_appendix(f"Selected {len(pdf_files)} PDF files to merge:")
            for pdf_file in sorted(pdf_files):
                try:
                    rel_path = pdf_file.relative_to(project_dir)
                    self.log_appendix(f"  - {rel_path}")
                except ValueError:
                    self.log_appendix(f"  - {pdf_file}")

            self.log_appendix("")
            self.log_appendix("Merging PDFs...")

            # Merge PDFs
            merger = PdfMerger()
            added_count = 0
            for pdf_file in sorted(pdf_files):
                try:
                    merger.append(str(pdf_file))
                    try:
                        rel_path = pdf_file.relative_to(project_dir)
                        self.log_appendix(f"  Added: {rel_path}")
                    except ValueError:
                        self.log_appendix(f"  Added: {pdf_file.name}")
                    added_count += 1
                except Exception as e:
                    self.log_appendix(f"  WARNING: Could not add {pdf_file.name}: {str(e)}")

            # Write output
            self.log_appendix("")
            self.log_appendix("Writing combined PDF...")
            merger.write(output_path)
            merger.close()

            self.log_appendix("")
            self.log_appendix(f"SUCCESS! Combined PDF created: {output_path}")
            self.status_var.set(f"Appendix generated: {added_count} PDFs combined")

            messagebox.showinfo("Success", f"Appendix PDF generated successfully!\n\nCombined {added_count} evaluation reports into:\n{output_path}")

        except Exception as e:
            self.log_appendix("")
            self.log_appendix(f"ERROR: {str(e)}")
            messagebox.showerror("Error", f"Error generating appendix: {str(e)}")
            self.status_var.set("Error generating appendix")
            import traceback
            traceback.print_exc()

    # =========================================================================
    # VISUAL VIEW FUNCTIONS
    # =========================================================================

    def refresh_visual_alignments(self):
        """Refresh the list of available highway alignments"""
        project_path = self.project_path.get()

        if self.placeholder_active or not project_path or not os.path.isdir(project_path):
            messagebox.showerror("Error", "Please select a valid project directory")
            return

        try:
            project_dir = Path(project_path)
            highway_dirs = []

            # Find direct highway directories (h*)
            direct_highways = [d for d in project_dir.iterdir()
                              if d.is_dir() and d.name[0].lower() == 'h']
            highway_dirs.extend(direct_highways)

            # Find highways inside interchanges (c*/h*)
            interchange_dirs = [d for d in project_dir.iterdir()
                               if d.is_dir() and d.name[0].lower() == 'c']

            for interchange_dir in interchange_dirs:
                nested_highways = [d for d in interchange_dir.iterdir()
                                  if d.is_dir() and d.name[0].lower() == 'h']
                highway_dirs.extend(nested_highways)

            if not highway_dirs:
                self.visual_alignment_var.set("No highway alignments found")
                self.visual_alignment_combo['values'] = []
                messagebox.showinfo("No Alignments", "No highway alignments found in project directory.")
                return

            # Get alignment names
            alignment_options = []
            for highway_dir in sorted(highway_dirs, key=lambda x: x.name):
                alignment_name = self.get_alignment_name(highway_dir)
                alignment_options.append(f"{highway_dir.name} - {alignment_name}")

            self.visual_alignment_combo['values'] = alignment_options
            if alignment_options:
                self.visual_alignment_var.set(alignment_options[0])

            self.status_var.set(f"Found {len(alignment_options)} highway alignments")

        except Exception as e:
            messagebox.showerror("Error", f"Error finding alignments: {str(e)}")

    def display_alignment(self):
        """Display the selected highway alignment visualization with rich data"""
        project_path = self.project_path.get()
        selected = self.visual_alignment_var.get()

        if self.placeholder_active or not project_path or not os.path.isdir(project_path):
            messagebox.showerror("Error", "Please select a valid project directory")
            return

        if not selected or selected == "No alignments found":
            messagebox.showwarning("No Selection", "Please select a highway alignment first")
            return

        # Extract alignment ID from selection (format: "h1 - Alignment Name")
        alignment_id = selected.split(' - ')[0]

        try:
            # Find the highway.1.xml file
            project_dir = Path(project_path)

            # Check direct highway directories
            highway_dir = project_dir / alignment_id
            if not highway_dir.exists():
                # Check inside interchanges
                for interchange_dir in project_dir.glob('c*'):
                    potential_dir = interchange_dir / alignment_id
                    if potential_dir.exists():
                        highway_dir = potential_dir
                        break

            highway_xml = highway_dir / "highway.1.xml"

            if not highway_xml.exists():
                messagebox.showerror("Error", f"Could not find highway.1.xml in {highway_dir}")
                return

            # Parse highway.1.xml
            tree = ET.parse(highway_xml)
            root = tree.getroot()

            # Find the Roadway element (actual data container)
            roadway = root.find('.//{http://www.ihsdm.org/schema/Highway-1.0}Roadway')
            if roadway is None:
                # Try without namespace
                roadway = root.find('.//Roadway')
            if roadway is None:
                messagebox.showerror("Error", "Could not find Roadway element in highway.1.xml")
                return

            # Extract alignment data from Roadway element
            min_station = roadway.get('minStation', '0.0')
            max_station = roadway.get('maxStation', '0.0')
            heading_sta = roadway.get('headingSta', '0.0')
            heading_angle = roadway.get('headingAngle', '0.0')

            # Clear canvas
            for widget in self.visual_canvas_frame.winfo_children():
                widget.destroy()

            # Create scrollable canvas
            canvas_container = tk.Frame(self.visual_canvas_frame, bg='lightblue')
            canvas_container.pack(fill=tk.BOTH, expand=True)

            scrollbar = ttk.Scrollbar(canvas_container, orient=tk.VERTICAL)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            canvas = tk.Canvas(canvas_container, bg='white', yscrollcommand=scrollbar.set, highlightthickness=0)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=canvas.yview)

            # Draw a test element to verify canvas is working
            print(f"DEBUG: Canvas created, drawing test elements...")

            # Parse alignment data
            min_sta_float = float(min_station)
            max_sta_float = float(max_station)
            length = max_sta_float - min_sta_float

            # Get canvas dimensions (wait for canvas to render)
            canvas.update()
            canvas_width = max(canvas.winfo_width(), 800)
            canvas_height = max(canvas.winfo_height(), 600)

            # Parse all data from XML (use roadway element, not root)
            lanes_data = self.parse_lanes(roadway, min_sta_float, max_sta_float)
            shoulders_data = self.parse_shoulders(roadway, min_sta_float, max_sta_float)
            ramps_data = self.parse_ramps(roadway)
            curves_data = self.parse_horizontal_curves(roadway, min_sta_float, max_sta_float)
            traffic_data = self.parse_traffic(roadway, min_sta_float, max_sta_float)
            median_data = self.parse_median(roadway, min_sta_float, max_sta_float)
            speed_data = self.parse_speed(roadway, min_sta_float, max_sta_float)
            func_class_data = self.parse_functional_class(roadway, min_sta_float, max_sta_float)

            # Debug: Print what we found
            print(f"DEBUG: Parsed data counts:")
            print(f"  Lanes: {len(lanes_data)}")
            print(f"  Shoulders: {len(shoulders_data)}")
            print(f"  Ramps: {len(ramps_data)}")
            print(f"  Curves: {len(curves_data)}")
            print(f"  Traffic: {len(traffic_data)}")
            print(f"  Median: {len(median_data)}")
            print(f"  Speed: {len(speed_data)}")
            print(f"  Functional Class: {len(func_class_data)}")

            # Calculate total height needed
            y_offset = 0
            margin_x = 40
            panel_width = canvas_width - 2 * margin_x

            # ===== HEADER SECTION =====
            header_height = 100
            print(f"DEBUG: Drawing header at canvas size: {canvas_width} x {canvas_height}")
            canvas.create_rectangle(0, y_offset, canvas_width, y_offset + header_height,
                                  fill='#f0f0f0', outline='#cccccc', width=2)
            canvas.create_text(20, y_offset + 15, anchor='nw', text=f"Highway Alignment: {selected}",
                             font=('Segoe UI', 14, 'bold'), fill='#1e3a8a')
            canvas.create_text(20, y_offset + 45, anchor='nw',
                             text=f"Station Range: {self.format_station(min_station)} to {self.format_station(max_station)} ({length:.2f} ft)",
                             font=('Segoe UI', 10), fill='black')
            canvas.create_text(20, y_offset + 65, anchor='nw',
                             text=f"Heading: {heading_angle}° at Sta {self.format_station(heading_sta)}",
                             font=('Segoe UI', 10), fill='black')
            y_offset += header_height + 20
            print(f"DEBUG: Header drawn, y_offset now: {y_offset}")

            # ===== LANE CONFIGURATION PANEL =====
            if lanes_data or shoulders_data:
                y_offset = self.draw_lane_panel(canvas, lanes_data, shoulders_data, margin_x, panel_width, y_offset,
                                              min_sta_float, max_sta_float)
                y_offset += 30

            # ===== RAMP LOCATIONS PANEL =====
            if ramps_data:
                y_offset = self.draw_ramp_panel(canvas, ramps_data, margin_x, panel_width, y_offset,
                                              min_sta_float, max_sta_float)
                y_offset += 30

            # ===== HORIZONTAL ALIGNMENT PANEL =====
            if curves_data:
                y_offset = self.draw_curves_panel(canvas, curves_data, margin_x, panel_width, y_offset,
                                                min_sta_float, max_sta_float)
                y_offset += 30

            # ===== TRAFFIC VOLUME PANEL =====
            if traffic_data:
                y_offset = self.draw_traffic_panel(canvas, traffic_data, margin_x, panel_width, y_offset,
                                                  min_sta_float, max_sta_float)
                y_offset += 30

            # ===== MEDIAN PANEL =====
            if median_data:
                y_offset = self.draw_median_panel(canvas, median_data, margin_x, panel_width, y_offset,
                                                min_sta_float, max_sta_float)
                y_offset += 30

            # ===== POSTED SPEED PANEL =====
            if speed_data:
                y_offset = self.draw_speed_panel(canvas, speed_data, margin_x, panel_width, y_offset,
                                               min_sta_float, max_sta_float)
                y_offset += 30

            # ===== FUNCTIONAL CLASS PANEL =====
            if func_class_data:
                y_offset = self.draw_func_class_panel(canvas, func_class_data, margin_x, panel_width, y_offset,
                                                    min_sta_float, max_sta_float)
                y_offset += 20

            # If no data panels were shown, display a message
            if y_offset <= header_height + 20:
                canvas.create_text(canvas_width // 2, header_height + 100,
                                 text="No detailed alignment data found in highway.1.xml",
                                 font=('Segoe UI', 12), fill='#999999', anchor='center')
                canvas.create_text(canvas_width // 2, header_height + 130,
                                 text="(Looking for LaneNS, RampConnector, HorizontalAlignment, Traffic, etc.)",
                                 font=('Segoe UI', 10), fill='#cccccc', anchor='center')
                y_offset = header_height + 200

            # Configure scroll region
            canvas.config(scrollregion=(0, 0, canvas_width, y_offset))

            # Bind mousewheel to scroll only when mouse is over the canvas
            def _on_mousewheel(event):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")

            def _bind_mousewheel(event):
                canvas.bind_all("<MouseWheel>", _on_mousewheel)

            def _unbind_mousewheel(event):
                canvas.unbind_all("<MouseWheel>")

            canvas.bind("<Enter>", _bind_mousewheel)
            canvas.bind("<Leave>", _unbind_mousewheel)

            self.status_var.set(f"Displaying alignment: {selected}")

        except Exception as e:
            messagebox.showerror("Error", f"Error displaying alignment: {str(e)}")
            import traceback
            traceback.print_exc()

    # ===== DATA PARSING FUNCTIONS =====

    def find_elements(self, parent, tag_name):
        """Find elements by tag name, handling XML namespaces"""
        # Try with IHSDM namespace
        ns = '{http://www.ihsdm.org/schema/Highway-1.0}'
        elements = parent.findall(f'.//{ns}{tag_name}')
        if elements:
            return elements

        # Try without namespace
        elements = parent.findall(f'.//{tag_name}')
        if elements:
            return elements

        # Try as direct children with namespace in tag
        return [child for child in parent if child.tag.endswith(tag_name)]

    def parse_lanes(self, root, min_sta, max_sta):
        """Parse lane configuration data from LaneNS elements"""
        lanes = []
        for lane in self.find_elements(root, 'LaneNS'):
            try:
                # IHSDM uses startStation/endStation, startWidth/endWidth
                start_sta = float(lane.get('startStation', min_sta))
                end_sta = float(lane.get('endStation', max_sta))
                start_width = float(lane.get('startWidth', '12.0'))
                end_width = float(lane.get('endWidth', start_width))
                avg_width = (start_width + end_width) / 2
                priority = int(lane.get('priority', '10'))
                side = lane.get('sideOfRoad', 'both')

                lanes.append({
                    'begin': start_sta,
                    'end': end_sta,
                    'side': side,
                    'priority': priority,
                    'lane_type': lane.get('laneType', 'thru'),
                    'width': avg_width
                })
            except (ValueError, AttributeError):
                continue
        # Sort by station first, then by side (left < both < right), then by priority
        return sorted(lanes, key=lambda x: (x['begin'], x['side'], x['priority']))

    def parse_shoulders(self, root, min_sta, max_sta):
        """Parse shoulder data from ShoulderSection elements"""
        shoulders = []
        for shoulder in self.find_elements(root, 'ShoulderSection'):
            try:
                start_sta = float(shoulder.get('startStation', min_sta))
                end_sta = float(shoulder.get('endStation', max_sta))
                start_width = float(shoulder.get('startWidth', '0.0'))
                end_width = float(shoulder.get('endWidth', start_width))
                avg_width = (start_width + end_width) / 2
                priority = int(shoulder.get('priority', '100'))
                side = shoulder.get('sideOfRoad', 'right')
                inside_outside = shoulder.get('insideOutsideOfRoadNB', 'outside')
                material = shoulder.get('material', 'paved')

                shoulders.append({
                    'begin': start_sta,
                    'end': end_sta,
                    'side': side,
                    'priority': priority,
                    'width': avg_width,
                    'position': inside_outside,  # inside or outside
                    'material': material
                })
            except (ValueError, AttributeError):
                continue
        return sorted(shoulders, key=lambda x: (x['begin'], x['side'], x['priority']))

    def parse_ramps(self, root):
        """Parse ramp connector data"""
        ramps = []
        # Look for RampConnector elements (for interchanges)
        for ramp in self.find_elements(root, 'RampConnector'):
            try:
                ramps.append({
                    'station': float(ramp.get('station', '0.0')),
                    'name': ramp.get('name', 'Ramp'),
                    'ramp_type': ramp.get('type', 'entrance')
                })
            except (ValueError, AttributeError):
                continue
        return sorted(ramps, key=lambda x: x['station'])

    def parse_horizontal_curves(self, root, min_sta, max_sta):
        """Parse horizontal alignment curves and tangents"""
        curves = []
        # IHSDM uses HorizontalElements (not HorizontalAlignment)
        h_elements_list = self.find_elements(root, 'HorizontalElements')
        if h_elements_list:
            h_elements = h_elements_list[0]

            # Look for HTangent elements
            for tangent in self.find_elements(h_elements, 'HTangent'):
                try:
                    curves.append({
                        'type': 'tangent',
                        'begin': float(tangent.get('startStation', '0.0')),
                        'end': float(tangent.get('endStation', '0.0'))
                    })
                except (ValueError, AttributeError):
                    continue

            # Look for HSimpleCurve elements
            for curve in self.find_elements(h_elements, 'HSimpleCurve'):
                try:
                    curves.append({
                        'type': 'curve',
                        'begin': float(curve.get('startStation', '0.0')),
                        'end': float(curve.get('endStation', '0.0')),
                        'radius': float(curve.get('radius', '0.0')),
                        'direction': curve.get('curveDirection', 'left')
                    })
                except (ValueError, AttributeError):
                    continue

            # Look for HSpiralCurve elements
            for spiral in self.find_elements(h_elements, 'HSpiralCurve'):
                try:
                    curves.append({
                        'type': 'spiral',
                        'begin': float(spiral.get('startStation', '0.0')),
                        'end': float(spiral.get('endStation', '0.0')),
                        'radius': float(spiral.get('radius', '0.0'))
                    })
                except (ValueError, AttributeError):
                    continue

        return sorted(curves, key=lambda x: x['begin'])

    def parse_traffic(self, root, min_sta, max_sta):
        """Parse traffic volume data"""
        traffic = []
        for aadt in self.find_elements(root, 'AnnualAveDailyTraffic'):
            try:
                # IHSDM uses startStation/endStation and adtRate
                traffic.append({
                    'begin': float(aadt.get('startStation', min_sta)),
                    'end': float(aadt.get('endStation', max_sta)),
                    'volume': int(float(aadt.get('adtRate', '0')))
                })
            except (ValueError, AttributeError):
                continue
        return sorted(traffic, key=lambda x: x['begin'])

    def parse_median(self, root, min_sta, max_sta):
        """Parse median data"""
        median = []
        for med in self.find_elements(root, 'Median'):
            try:
                # IHSDM uses startStation/endStation
                median.append({
                    'begin': float(med.get('startStation', min_sta)),
                    'end': float(med.get('endStation', max_sta)),
                    'width': float(med.get('width', '0.0')),
                    'median_type': med.get('medianType', 'none')
                })
            except (ValueError, AttributeError):
                continue
        return sorted(median, key=lambda x: x['begin'])

    def parse_speed(self, root, min_sta, max_sta):
        """Parse posted speed data"""
        speeds = []
        for speed in self.find_elements(root, 'PostedSpeed'):
            try:
                # IHSDM uses startStation/endStation and speedLimit
                speeds.append({
                    'begin': float(speed.get('startStation', min_sta)),
                    'end': float(speed.get('endStation', max_sta)),
                    'speed': int(float(speed.get('speedLimit', '0')))
                })
            except (ValueError, AttributeError):
                continue
        return sorted(speeds, key=lambda x: x['begin'])

    def parse_functional_class(self, root, min_sta, max_sta):
        """Parse functional class data"""
        func_classes = []
        for fc in self.find_elements(root, 'FunctionalClass'):
            try:
                # IHSDM uses startStation/endStation and funcClass
                func_classes.append({
                    'begin': float(fc.get('startStation', min_sta)),
                    'end': float(fc.get('endStation', max_sta)),
                    'class_type': fc.get('funcClass', 'unknown')
                })
            except (ValueError, AttributeError):
                continue
        return sorted(func_classes, key=lambda x: x['begin'])

    # ===== DRAWING FUNCTIONS =====

    def draw_lane_panel(self, canvas, lanes_data, shoulders_data, margin_x, panel_width, y_start, min_sta, max_sta):
        """Draw lane configuration as roadway plan view with lanes building from centerline

        Layout (based on user sketch):
        - Centerline (red) runs horizontally
        - LEFT side lanes appear ABOVE centerline (opposite direction)
        - RIGHT side lanes appear BELOW centerline (forward direction, increasing stationing)
        - BOTH side lanes appear mirrored on BOTH sides of centerline
        - Inside shoulders = closest to centerline/median
        - Outside shoulders = on outer edge of road
        - Priority determines stacking: P10 closest to CL, P20 next, etc.
        """
        canvas.create_text(margin_x, y_start, anchor='nw', text="Lane Configuration (Plan View)",
                         font=('Segoe UI', 11, 'bold'), fill='#1e3a8a')
        y_start += 25

        # Count max priority levels on each side for lanes
        right_priorities = set()
        left_priorities = set()

        for lane in lanes_data:
            priority_level = lane['priority'] // 10
            if lane['side'] == 'right':
                right_priorities.add(priority_level)
            elif lane['side'] == 'left':
                left_priorities.add(priority_level)
            else:  # both - appears on BOTH sides
                right_priorities.add(priority_level)
                left_priorities.add(priority_level)

        # Count inside and outside shoulders per side
        right_inside_shoulders = 0
        right_outside_shoulders = 0
        left_inside_shoulders = 0
        left_outside_shoulders = 0

        for shoulder in shoulders_data:
            position = shoulder.get('position', 'outside')
            side = shoulder['side']
            if side == 'right' or side == 'both':
                if position == 'inside':
                    right_inside_shoulders = max(right_inside_shoulders, 1)
                else:
                    right_outside_shoulders = max(right_outside_shoulders, 1)
            if side == 'left' or side == 'both':
                if position == 'inside':
                    left_inside_shoulders = max(left_inside_shoulders, 1)
                else:
                    left_outside_shoulders = max(left_outside_shoulders, 1)

        max_right_lanes = max(right_priorities) + 1 if right_priorities else 0
        max_left_lanes = max(left_priorities) + 1 if left_priorities else 0

        lane_height = 18
        shoulder_height = 12

        # Panel height calculation:
        # Above CL: left outside shoulder + left lanes + left inside shoulder
        # Below CL: right inside shoulder + right lanes + right outside shoulder
        total_above = (left_outside_shoulders * shoulder_height +
                      max_left_lanes * lane_height +
                      left_inside_shoulders * shoulder_height)
        total_below = (right_inside_shoulders * shoulder_height +
                      max_right_lanes * lane_height +
                      right_outside_shoulders * shoulder_height)
        panel_height = max(120, total_above + total_below + 50)

        canvas.create_rectangle(margin_x, y_start, margin_x + panel_width, y_start + panel_height,
                              fill='#e5e7eb', outline='#666666', width=2)

        # Centerline position - account for inside shoulders
        centerline_y = y_start + left_outside_shoulders * shoulder_height + max_left_lanes * lane_height + left_inside_shoulders * shoulder_height + 25

        # Draw centerline (red dashed)
        canvas.create_line(margin_x + 10, centerline_y, margin_x + panel_width - 10, centerline_y,
                         fill='#dc2626', width=3, dash=(10, 5))
        canvas.create_text(margin_x + 15, centerline_y, text="CL", font=('Consolas', 8, 'bold'),
                         fill='#dc2626', anchor='w')

        # Helper to draw a bar (lane or shoulder)
        def draw_bar(item, is_shoulder=False, shoulder_position='outside'):
            x_start = margin_x + 10 + (item['begin'] - min_sta) / (max_sta - min_sta) * (panel_width - 20)
            x_end = margin_x + 10 + (item['end'] - min_sta) / (max_sta - min_sta) * (panel_width - 20)

            priority_level = item['priority'] // 10
            bar_height = shoulder_height if is_shoulder else lane_height

            if is_shoulder:
                color = '#22c55e'  # Green for shoulders
                pos_label = "In" if shoulder_position == 'inside' else "Out"
                label = f"{pos_label} Shldr"
            else:
                color = self.get_lane_color(item.get('lane_type', 'thru'))
                label = f"P{item['priority']} {item.get('lane_type', '')}"

            side = item['side']

            # Determine which sides to draw on
            sides_to_draw = []
            if side == 'right':
                sides_to_draw = ['right']
            elif side == 'left':
                sides_to_draw = ['left']
            else:  # both
                sides_to_draw = ['left', 'right']

            for draw_side in sides_to_draw:
                if draw_side == 'right':
                    # Right side: BELOW centerline
                    if is_shoulder:
                        if shoulder_position == 'inside':
                            # Inside shoulder: between centerline and lanes
                            y_top = centerline_y
                            y_bottom = y_top + shoulder_height
                        else:
                            # Outside shoulder: beyond all lanes
                            y_top = centerline_y + right_inside_shoulders * shoulder_height + max_right_lanes * lane_height
                            y_bottom = y_top + shoulder_height
                    else:
                        # Lanes stack from inside shoulder outward
                        y_top = centerline_y + right_inside_shoulders * shoulder_height + priority_level * lane_height
                        y_bottom = y_top + lane_height
                else:  # left
                    # Left side: ABOVE centerline
                    if is_shoulder:
                        if shoulder_position == 'inside':
                            # Inside shoulder: between centerline and lanes
                            y_bottom = centerline_y
                            y_top = y_bottom - shoulder_height
                        else:
                            # Outside shoulder: beyond all lanes
                            y_bottom = centerline_y - left_inside_shoulders * shoulder_height - max_left_lanes * lane_height
                            y_top = y_bottom - shoulder_height
                    else:
                        # Lanes stack from inside shoulder outward
                        y_bottom = centerline_y - left_inside_shoulders * shoulder_height - priority_level * lane_height
                        y_top = y_bottom - lane_height

                # Draw the bar
                canvas.create_rectangle(x_start, y_top, x_end, y_bottom,
                                      fill=color, outline='#1e3a8a', width=1)

                # Add label on the bar
                bar_width = x_end - x_start
                bar_mid_y = (y_top + y_bottom) / 2

                if bar_width > 100:
                    canvas.create_text((x_start + x_end) / 2, bar_mid_y,
                                     text=label, font=('Consolas', 7, 'bold'), fill='white')
                elif bar_width > 50:
                    short_label = f"P{item['priority']}" if not is_shoulder else pos_label
                    canvas.create_text((x_start + x_end) / 2, bar_mid_y,
                                     text=short_label, font=('Consolas', 7), fill='white')

        # Draw inside shoulders first (closest to centerline)
        for shoulder in shoulders_data:
            if shoulder.get('position', 'outside') == 'inside':
                draw_bar(shoulder, is_shoulder=True, shoulder_position='inside')

        # Draw lanes
        for lane in lanes_data:
            draw_bar(lane, is_shoulder=False)

        # Draw outside shoulders last (furthest from centerline)
        for shoulder in shoulders_data:
            if shoulder.get('position', 'outside') == 'outside':
                draw_bar(shoulder, is_shoulder=True, shoulder_position='outside')

        # Add station markers at bottom
        marker_y = y_start + panel_height - 15
        num_markers = 5
        for i in range(num_markers):
            x = margin_x + 10 + (panel_width - 20) * i / (num_markers - 1)
            sta = min_sta + (max_sta - min_sta) * i / (num_markers - 1)
            canvas.create_line(x, marker_y - 5, x, marker_y + 5, fill='#666666', width=1)
            canvas.create_text(x, marker_y + 8, text=self.format_station(str(sta)),
                             font=('Consolas', 7), fill='black', anchor='n')

        # Direction labels
        canvas.create_text(margin_x + panel_width - 10, centerline_y + 8,
                         text="Right side (forward →)", font=('Consolas', 7),
                         fill='#666666', anchor='ne')
        canvas.create_text(margin_x + panel_width - 10, centerline_y - 8,
                         text="Left side (← opposite)", font=('Consolas', 7),
                         fill='#666666', anchor='se')

        return y_start + panel_height

    def get_lane_color(self, lane_type):
        """Get color for lane type based on IHSDM lane types"""
        colors = {
            'thru': '#3b82f6',       # Blue for through lanes
            'pass': '#8b5cf6',       # Purple for passing lanes
            'climb': '#7c3aed',      # Violet for climb lanes
            'left_turn': '#f59e0b',  # Orange for left turn lanes
            'right_turn': '#eab308', # Yellow for right turn lanes
            'taper': '#f97316',      # Dark orange for tapers
            'park': '#6b7280',       # Gray for parking lanes
            'bike': '#84cc16',       # Lime for bike lanes
            'accel': '#06b6d4',      # Cyan for accel lanes
            'decel': '#ec4899',      # Pink for decel lanes
            'other': '#9ca3af',      # Light gray for other
            'unspec': '#d1d5db'      # Lighter gray for unspecified
        }
        return colors.get(lane_type, '#3b82f6')  # Default to blue

    def draw_ramp_panel(self, canvas, ramps_data, margin_x, panel_width, y_start, min_sta, max_sta):
        """Draw ramp locations panel"""
        canvas.create_text(margin_x, y_start, anchor='nw', text="Ramp Locations",
                         font=('Segoe UI', 11, 'bold'), fill='#1e3a8a')
        y_start += 25

        panel_height = 80
        canvas.create_rectangle(margin_x, y_start, margin_x + panel_width, y_start + panel_height,
                              fill='white', outline='#666666', width=2)

        # Draw baseline
        baseline_y = y_start + panel_height // 2
        canvas.create_line(margin_x + 10, baseline_y, margin_x + panel_width - 10, baseline_y,
                         fill='#666666', width=2)

        # Draw ramp markers
        for ramp in ramps_data:
            x = margin_x + 10 + (ramp['station'] - min_sta) / (max_sta - min_sta) * (panel_width - 20)
            color = '#10b981' if ramp['ramp_type'] == 'entrance' else '#ef4444'

            # Draw ramp marker (triangle)
            if ramp['ramp_type'] == 'entrance':
                canvas.create_polygon(x, baseline_y - 15, x - 6, baseline_y - 3, x + 6, baseline_y - 3,
                                    fill=color, outline='#1e3a8a', width=1)
            else:
                canvas.create_polygon(x, baseline_y + 15, x - 6, baseline_y + 3, x + 6, baseline_y + 3,
                                    fill=color, outline='#1e3a8a', width=1)

            # Label
            canvas.create_text(x, baseline_y + (30 if ramp['ramp_type'] == 'exit' else -30),
                             text=ramp['name'], font=('Consolas', 7), fill='black', anchor='center')

        return y_start + panel_height

    def draw_curves_panel(self, canvas, curves_data, margin_x, panel_width, y_start, min_sta, max_sta):
        """Draw horizontal curves panel"""
        canvas.create_text(margin_x, y_start, anchor='nw', text="Horizontal Alignment (Curves & Tangents)",
                         font=('Segoe UI', 11, 'bold'), fill='#1e3a8a')
        y_start += 25

        panel_height = 100
        canvas.create_rectangle(margin_x, y_start, margin_x + panel_width, y_start + panel_height,
                              fill='white', outline='#666666', width=2)

        # Draw alignment elements
        baseline_y = y_start + panel_height // 2
        for curve in curves_data:
            x_start = margin_x + 10 + (curve['begin'] - min_sta) / (max_sta - min_sta) * (panel_width - 20)
            x_end = margin_x + 10 + (curve['end'] - min_sta) / (max_sta - min_sta) * (panel_width - 20)

            if curve['type'] == 'tangent':
                # Draw straight line for tangent
                canvas.create_line(x_start, baseline_y, x_end, baseline_y, fill='#3b82f6', width=4)
                canvas.create_text((x_start + x_end) / 2, baseline_y - 15, text="Tangent",
                                 font=('Consolas', 7), fill='#666666')
            else:
                # Draw arc for curve
                direction = curve.get('direction', 'left')
                radius = curve.get('radius', 0)

                # Simplified curve visualization
                y_offset = -20 if direction == 'left' else 20
                mid_x = (x_start + x_end) / 2

                # Draw curve as smooth arc (approximate with line segments)
                points = []
                for i in range(11):
                    t = i / 10
                    x = x_start + t * (x_end - x_start)
                    # Parabolic approximation of curve
                    y = baseline_y + y_offset * (4 * t * (1 - t))
                    points.extend([x, y])

                if len(points) >= 4:
                    canvas.create_line(points, fill='#f59e0b', width=4, smooth=True)

                canvas.create_text(mid_x, baseline_y + y_offset + (15 if direction == 'right' else -15),
                                 text=f"Curve R={radius:.0f}'", font=('Consolas', 7), fill='#f59e0b')

        return y_start + panel_height

    def draw_traffic_panel(self, canvas, traffic_data, margin_x, panel_width, y_start, min_sta, max_sta):
        """Draw traffic volume panel"""
        canvas.create_text(margin_x, y_start, anchor='nw', text="Annual Average Daily Traffic (AADT)",
                         font=('Segoe UI', 11, 'bold'), fill='#1e3a8a')
        y_start += 25

        panel_height = 100
        canvas.create_rectangle(margin_x, y_start, margin_x + panel_width, y_start + panel_height,
                              fill='white', outline='#666666', width=2)

        if not traffic_data:
            return y_start + panel_height

        # Find max volume for scaling
        max_volume = max(t['volume'] for t in traffic_data)

        # Draw traffic bars
        for traffic in traffic_data:
            x_start = margin_x + 10 + (traffic['begin'] - min_sta) / (max_sta - min_sta) * (panel_width - 20)
            x_end = margin_x + 10 + (traffic['end'] - min_sta) / (max_sta - min_sta) * (panel_width - 20)

            bar_height = (traffic['volume'] / max_volume) * (panel_height - 40)
            y_bar = y_start + panel_height - 20 - bar_height

            canvas.create_rectangle(x_start, y_bar, x_end, y_start + panel_height - 20,
                                  fill='#8b5cf6', outline='#5b21b6', width=1)
            canvas.create_text((x_start + x_end) / 2, y_bar - 5, text=f"{traffic['volume']:,}",
                             font=('Consolas', 8), fill='#5b21b6', anchor='s')

        return y_start + panel_height

    def draw_median_panel(self, canvas, median_data, margin_x, panel_width, y_start, min_sta, max_sta):
        """Draw median width panel"""
        canvas.create_text(margin_x, y_start, anchor='nw', text="Median Width",
                         font=('Segoe UI', 11, 'bold'), fill='#1e3a8a')
        y_start += 25

        panel_height = 70
        canvas.create_rectangle(margin_x, y_start, margin_x + panel_width, y_start + panel_height,
                              fill='white', outline='#666666', width=2)

        baseline_y = y_start + panel_height - 20
        for median in median_data:
            x_start = margin_x + 10 + (median['begin'] - min_sta) / (max_sta - min_sta) * (panel_width - 20)
            x_end = margin_x + 10 + (median['end'] - min_sta) / (max_sta - min_sta) * (panel_width - 20)

            canvas.create_rectangle(x_start, y_start + 15, x_end, baseline_y,
                                  fill='#a8a29e', outline='#78716c', width=1)
            canvas.create_text((x_start + x_end) / 2, y_start + 25,
                             text=f"{median['width']:.0f}' {median['median_type']}",
                             font=('Consolas', 8), fill='white')

        return y_start + panel_height

    def draw_speed_panel(self, canvas, speed_data, margin_x, panel_width, y_start, min_sta, max_sta):
        """Draw posted speed panel"""
        canvas.create_text(margin_x, y_start, anchor='nw', text="Posted Speed Limit",
                         font=('Segoe UI', 11, 'bold'), fill='#1e3a8a')
        y_start += 25

        panel_height = 60
        canvas.create_rectangle(margin_x, y_start, margin_x + panel_width, y_start + panel_height,
                              fill='white', outline='#666666', width=2)

        for speed in speed_data:
            x_start = margin_x + 10 + (speed['begin'] - min_sta) / (max_sta - min_sta) * (panel_width - 20)
            x_end = margin_x + 10 + (speed['end'] - min_sta) / (max_sta - min_sta) * (panel_width - 20)

            canvas.create_rectangle(x_start, y_start + 10, x_end, y_start + panel_height - 10,
                                  fill='#fef3c7', outline='#f59e0b', width=2)
            canvas.create_text((x_start + x_end) / 2, y_start + panel_height // 2,
                             text=f"{speed['speed']} mph",
                             font=('Segoe UI', 10, 'bold'), fill='#f59e0b')

        return y_start + panel_height

    def draw_func_class_panel(self, canvas, func_class_data, margin_x, panel_width, y_start, min_sta, max_sta):
        """Draw functional class panel"""
        canvas.create_text(margin_x, y_start, anchor='nw', text="Functional Classification",
                         font=('Segoe UI', 11, 'bold'), fill='#1e3a8a')
        y_start += 25

        panel_height = 60
        canvas.create_rectangle(margin_x, y_start, margin_x + panel_width, y_start + panel_height,
                              fill='white', outline='#666666', width=2)

        for fc in func_class_data:
            x_start = margin_x + 10 + (fc['begin'] - min_sta) / (max_sta - min_sta) * (panel_width - 20)
            x_end = margin_x + 10 + (fc['end'] - min_sta) / (max_sta - min_sta) * (panel_width - 20)

            color = '#dcfce7' if 'freeway' in fc['class_type'].lower() else '#e0e7ff'
            border = '#10b981' if 'freeway' in fc['class_type'].lower() else '#6366f1'

            canvas.create_rectangle(x_start, y_start + 10, x_end, y_start + panel_height - 10,
                                  fill=color, outline=border, width=2)
            canvas.create_text((x_start + x_end) / 2, y_start + panel_height // 2,
                             text=fc['class_type'],
                             font=('Consolas', 9), fill=border)

        return y_start + panel_height

    # =========================================================================
    # CMF SCANNER TAB
    # =========================================================================

    def setup_cmf_tab(self):
        """Setup the CMF calibration factor scanner tab"""
        tab_frame = ttk.Frame(self.cmf_tab, padding="10")
        tab_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.cmf_tab.columnconfigure(0, weight=1)
        self.cmf_tab.rowconfigure(0, weight=1)

        # Scan button
        scan_frame = ttk.Frame(tab_frame)
        scan_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        scan_btn = ttk.Button(scan_frame,
                             text="Scan for CMF Calibration Factors",
                             command=self.scan_cmf_files)
        scan_btn.pack(side=tk.LEFT, padx=5)

        # Export button
        export_cmf_btn = ttk.Button(scan_frame,
                                   text="Export to Excel",
                                   command=self.export_cmf_to_excel)
        export_cmf_btn.pack(side=tk.LEFT, padx=5)

        # Summary label
        self.cmf_summary_var = tk.StringVar(value="No scan performed yet")
        summary_label = ttk.Label(scan_frame,
                                 textvariable=self.cmf_summary_var,
                                 foreground='blue')
        summary_label.pack(side=tk.LEFT, padx=20)

        # Results tree
        tree_frame = ttk.Frame(tab_frame)
        tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        tab_frame.rowconfigure(1, weight=1)

        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

        self.cmf_tree = ttk.Treeview(tree_frame,
                                    columns=('Type', 'ID', 'Name', 'Evaluation', 'Calibration'),
                                    show='tree headings',
                                    yscrollcommand=vsb.set,
                                    xscrollcommand=hsb.set)

        vsb.config(command=self.cmf_tree.yview)
        hsb.config(command=self.cmf_tree.xview)

        # Column configuration
        self.cmf_tree.heading('#0', text='Item')
        self.cmf_tree.heading('Type', text='Type')
        self.cmf_tree.heading('ID', text='ID')
        self.cmf_tree.heading('Name', text='Name')
        self.cmf_tree.heading('Evaluation', text='Evaluation')
        self.cmf_tree.heading('Calibration', text='Calibration Factor')

        self.cmf_tree.column('#0', width=200, anchor='w')
        self.cmf_tree.column('Type', width=100, anchor='center')
        self.cmf_tree.column('ID', width=80, anchor='center')
        self.cmf_tree.column('Name', width=300, anchor='w')
        self.cmf_tree.column('Evaluation', width=120, anchor='w')
        self.cmf_tree.column('Calibration', width=200, anchor='w')

        # Grid layout
        self.cmf_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))

        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        # Store CMF data for export
        self.cmf_data = []

    def scan_cmf_files(self):
        """Scan project for CMF calibration files"""
        print("DEBUG: scan_cmf_files called")
        project_path = self.project_path.get()
        print(f"DEBUG: project_path = {project_path}")
        print(f"DEBUG: placeholder_active = {self.placeholder_active}")

        # Check if placeholder is active
        if self.placeholder_active or not project_path or not os.path.isdir(project_path):
            print("DEBUG: Invalid project path")
            messagebox.showerror("Error", "Please select a valid project directory")
            return

        print("DEBUG: Starting scan...")

        self.status_var.set("Scanning for CMF files...")
        self.cmf_tree.delete(*self.cmf_tree.get_children())
        self.cmf_data = []

        try:
            # Find all evaluation.1.cpm.cmf.csv files
            cmf_files = []
            for root, dirs, files in os.walk(project_path):
                for file in files:
                    if file == 'evaluation.1.cpm.cmf.csv':
                        cmf_files.append(os.path.join(root, file))

            if not cmf_files:
                self.cmf_summary_var.set("No CMF files found")
                self.status_var.set("No CMF files found in project")
                return

            # Parse each CMF file
            highway_count = 0
            intersection_count = 0
            ramp_count = 0

            for cmf_path in cmf_files:
                # Extract alignment info from path
                # Path format: .../h1/e6/evaluation.1.cpm.cmf.csv or .../i100/e1/evaluation.1.cpm.cmf.csv
                path_parts = cmf_path.replace('\\', '/').split('/')

                alignment_id = None
                alignment_type = None
                eval_folder = None

                for i, part in enumerate(path_parts):
                    if part.startswith('h') and part[1:].isdigit():
                        alignment_id = part
                        alignment_type = 'Highway'
                        if i + 1 < len(path_parts):
                            eval_folder = path_parts[i + 1]
                        break
                    elif part.startswith('i') and part[1:].isdigit():
                        alignment_id = part
                        alignment_type = 'Intersection'
                        if i + 1 < len(path_parts):
                            eval_folder = path_parts[i + 1]
                        break
                    elif part.startswith('r') and part[1:].isdigit():
                        alignment_id = part
                        alignment_type = 'Ramp Terminal'
                        if i + 1 < len(path_parts):
                            eval_folder = path_parts[i + 1]
                        break

                if not alignment_id:
                    continue

                # Read the CMF file and extract line 27, column B
                calibration = None
                alignment_name = "Unknown"

                try:
                    with open(cmf_path, 'r', encoding='utf-8') as f:
                        lines = f.readlines()

                        # Line 27 (index 26) contains calibration
                        if len(lines) >= 27:
                            line27 = lines[26].strip()
                            parts = line27.split(',')
                            if len(parts) >= 2:
                                # Remove quotes from calibration value
                                calibration = parts[1].strip('"')

                        # Line 17 contains highway/alignment name
                        if len(lines) >= 17:
                            line17 = lines[16].strip()
                            parts = line17.split(',')
                            if len(parts) >= 2:
                                alignment_name = parts[1].strip('"')

                except Exception as e:
                    calibration = f"Error: {str(e)}"

                # Store data
                cmf_entry = {
                    'type': alignment_type,
                    'id': alignment_id,
                    'name': alignment_name,
                    'evaluation': eval_folder if eval_folder else 'Unknown',
                    'calibration': calibration if calibration else 'Not found',
                    'path': cmf_path
                }
                self.cmf_data.append(cmf_entry)

                # Count by type
                if alignment_type == 'Highway':
                    highway_count += 1
                elif alignment_type == 'Intersection':
                    intersection_count += 1
                elif alignment_type == 'Ramp Terminal':
                    ramp_count += 1

            # Populate tree
            self.populate_cmf_tree()

            # Update summary
            summary = f"Total: {len(self.cmf_data)} | "
            summary += f"Highways: {highway_count} | "
            summary += f"Intersections: {intersection_count} | "
            summary += f"Ramp Terminals: {ramp_count}"
            self.cmf_summary_var.set(summary)
            self.status_var.set(f"Found {len(self.cmf_data)} CMF files")

        except Exception as e:
            messagebox.showerror("Scan Error", f"Error scanning CMF files:\n{str(e)}")
            self.status_var.set("CMF scan failed")

    def populate_cmf_tree(self):
        """Populate the CMF tree with grouped data"""
        # Group by type
        highways = [d for d in self.cmf_data if d['type'] == 'Highway']
        intersections = [d for d in self.cmf_data if d['type'] == 'Intersection']
        ramps = [d for d in self.cmf_data if d['type'] == 'Ramp Terminal']

        # Add highways
        if highways:
            hw_parent = self.cmf_tree.insert('', 'end', text=f'HIGHWAYS ({len(highways)})',
                                            values=('', '', '', '', ''))
            for hw in sorted(highways, key=lambda x: x['id']):
                self.cmf_tree.insert(hw_parent, 'end',
                                   text=hw['id'],
                                   values=(hw['type'], hw['id'], hw['name'],
                                         hw['evaluation'], hw['calibration']))

        # Add intersections
        if intersections:
            int_parent = self.cmf_tree.insert('', 'end', text=f'INTERSECTIONS ({len(intersections)})',
                                             values=('', '', '', '', ''))
            for i in sorted(intersections, key=lambda x: x['id']):
                self.cmf_tree.insert(int_parent, 'end',
                                   text=i['id'],
                                   values=(i['type'], i['id'], i['name'],
                                         i['evaluation'], i['calibration']))

        # Add ramp terminals
        if ramps:
            ramp_parent = self.cmf_tree.insert('', 'end', text=f'RAMP TERMINALS ({len(ramps)})',
                                              values=('', '', '', '', ''))
            for r in sorted(ramps, key=lambda x: x['id']):
                self.cmf_tree.insert(ramp_parent, 'end',
                                   text=r['id'],
                                   values=(r['type'], r['id'], r['name'],
                                         r['evaluation'], r['calibration']))

    def export_cmf_to_excel(self):
        """Export CMF data to Excel"""
        if not self.cmf_data:
            messagebox.showwarning("No Data", "No CMF data to export. Please scan first.")
            return

        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Missing Dependency",
                               "openpyxl is required for Excel export.\n"
                               "Install it with: pip install openpyxl")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="CMF_Calibration_Factors.xlsx"
        )

        if not filename:
            return

        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment

            wb = Workbook()
            ws = wb.active
            ws.title = "CMF Calibration Factors"

            # Headers
            headers = ['Type', 'ID', 'Name', 'Evaluation', 'Calibration Factor', 'File Path']
            ws.append(headers)

            # Style headers
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center')

            # Add data
            for entry in self.cmf_data:
                ws.append([
                    entry['type'],
                    entry['id'],
                    entry['name'],
                    entry['evaluation'],
                    entry['calibration'],
                    entry['path']
                ])

            # Adjust column widths
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 10
            ws.column_dimensions['C'].width = 50
            ws.column_dimensions['D'].width = 15
            ws.column_dimensions['E'].width = 25
            ws.column_dimensions['F'].width = 60

            wb.save(filename)
            messagebox.showinfo("Export Successful",
                              f"CMF data exported to:\n{filename}")
            self.status_var.set(f"Exported CMF data to {os.path.basename(filename)}")

        except Exception as e:
            messagebox.showerror("Export Error",
                               f"Failed to export to Excel:\n{str(e)}")

    # =========================================================================
    # UPDATE CHECKING
    # =========================================================================

    def check_for_updates(self, show_current=False):
        """Check GitHub for new releases"""
        if not GITHUB_API_URL:
            if show_current:
                messagebox.showinfo("Version",
                    f"Current version: {__version__}\n\n"
                    "GitHub repository not configured for auto-updates.")
            return

        def check_thread():
            try:
                # Make request to GitHub API
                req = Request(GITHUB_API_URL)
                req.add_header('User-Agent', f'{__app_name__}/{__version__}')

                with urlopen(req, timeout=5) as response:
                    data = json.loads(response.read().decode('utf-8'))

                    latest_version = data['tag_name'].lstrip('v')
                    download_url = data['html_url']
                    release_notes = data['body']

                    # Compare versions
                    if self._compare_versions(latest_version, __version__) > 0:
                        # New version available
                        self.root.after(0, lambda: self._show_update_dialog(
                            latest_version, download_url, release_notes))
                    elif show_current:
                        self.root.after(0, lambda: messagebox.showinfo("Up to Date",
                            f"You have the latest version ({__version__})"))

            except URLError:
                if show_current:
                    self.root.after(0, lambda: messagebox.showwarning("Update Check Failed",
                        "Could not connect to GitHub to check for updates.\n"
                        "Please check your internet connection."))
            except Exception as e:
                if show_current:
                    self.root.after(0, lambda: messagebox.showerror("Update Check Error",
                        f"Error checking for updates: {str(e)}"))

        # Run in background thread
        thread = threading.Thread(target=check_thread, daemon=True)
        thread.start()

    def _compare_versions(self, v1, v2):
        """Compare version strings (returns: 1 if v1>v2, -1 if v1<v2, 0 if equal)"""
        parts1 = [int(x) for x in v1.split('.')]
        parts2 = [int(x) for x in v2.split('.')]

        for p1, p2 in zip(parts1, parts2):
            if p1 > p2:
                return 1
            elif p1 < p2:
                return -1

        if len(parts1) > len(parts2):
            return 1
        elif len(parts1) < len(parts2):
            return -1

        return 0

    def _show_update_dialog(self, new_version, download_url, release_notes):
        """Show dialog when update is available"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Update Available")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()

        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"500x400+{x}+{y}")

        # Header
        header = tk.Frame(dialog, bg=self.colors['primary'], height=60)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        tk.Label(header, text="🎉 New Version Available!",
                font=('Segoe UI', 14, 'bold'),
                bg=self.colors['primary'],
                fg='white').pack(pady=15)

        # Content
        content = tk.Frame(dialog, padx=20, pady=20)
        content.pack(fill=tk.BOTH, expand=True)

        # Version info
        info_frame = tk.Frame(content)
        info_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(info_frame, text=f"Current version: {__version__}",
                font=('Segoe UI', 10)).pack(anchor=tk.W)
        tk.Label(info_frame, text=f"New version: {new_version}",
                font=('Segoe UI', 10, 'bold'),
                fg=self.colors['primary']).pack(anchor=tk.W)

        # Release notes
        tk.Label(content, text="What's New:",
                font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(10, 5))

        notes_text = scrolledtext.ScrolledText(content, wrap=tk.WORD, height=10,
                                               font=('Segoe UI', 9))
        notes_text.pack(fill=tk.BOTH, expand=True)
        notes_text.insert('1.0', release_notes if release_notes else "No release notes available.")
        notes_text.config(state=tk.DISABLED)

        # Buttons
        btn_frame = tk.Frame(content)
        btn_frame.pack(fill=tk.X, pady=(15, 0))

        def open_download():
            webbrowser.open(download_url)
            dialog.destroy()

        ttk.Button(btn_frame, text="Download Update", command=open_download,
                  style='Accent.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Remind Me Later", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    # =========================================================================
    # APPLICATION CONTROL
    # =========================================================================

    def run(self):
        """Start the application"""
        # Check for updates on startup (non-blocking)
        self.root.after(1000, lambda: self.check_for_updates(show_current=False))
        self.root.mainloop()


def main():
    """Entry point"""
    app = IHSDMWisconsinHelper()
    app.run()


if __name__ == "__main__":
    main()
