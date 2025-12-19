<<<<<<< HEAD
# IHSDadaM

**I**nteractive **H**ighway **S**afety **D**esign Model - **Data** **M**anager

A comprehensive tool for IHSDM project analysis and data compilation.

## Features

- **Warning Message Extraction** - Scan and filter warning messages from highways, intersections, and ramp terminals
- **Data Compilation** - Compile crash predictions to Excel with HSM severity distributions
- **Appendix Generator** - Generate detailed highway appendices with geometry visualization
- **Visual View** - Visualize highway geometry including curves, grades, traffic, and cross-sections

## Installation

### Option 1: Download Executable (Recommended)
1. Go to [Releases](https://github.com/YOUR_USERNAME/ihsdadam/releases)
2. Download the latest `IHSDadaM.exe`
3. Run the executable - no installation needed!

### Option 2: Run from Source
1. Clone this repository
2. Install Python 3.8 or later
3. Install dependencies:
   ```bash
   pip install openpyxl
   ```
4. Run the script:
   ```bash
   python ihsdm_wisconsin_helper.py
   ```

## Usage

1. **Select Project** - Browse to your IHSDM project directory (e.g., `Projects_V5/p263`)
2. **Choose Tab** - Select the tool you want to use:
   - Warning Extractor: Scan for warning messages
   - Data Compiler: Export crash predictions to Excel
   - Appendix Generator: Create highway documentation
   - Visual View: Visualize highway geometry

## Requirements

- Windows 7/8/10/11
- IHSDM 2018 or later
- For Excel export: Microsoft Excel or compatible spreadsheet software

## Auto-Updates

The application automatically checks for updates on startup. Click the version number in the header to manually check for updates.

## Development

### Building the Executable

Simply run:
```bash
build_wisconsin_helper.bat
```

Or manually:
```bash
pip install pyinstaller openpyxl PyPDF2
pyinstaller --onefile --windowed --name "IHSDadaM" --add-data "version.py;." ihsdm_wisconsin_helper.py
```

### Version Management

Update the version in `version.py`:
```python
__version__ = "1.1.0"  # Update this
```

## License

Developed by Adam Engbring (aengbring@hntb.com)

## Support

For bug reports and feature requests, please open an issue on GitHub.
=======
# ihsdadam
>>>>>>> 6cb69daeb44d9368cb1640eb3c2b988c7d01a94c
