"""QThread workers for long-running operations — keeps the UI responsive."""

import os
import csv
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Dict, Optional
from urllib.request import urlopen, Request
from urllib.error import URLError

from PySide6.QtCore import QThread, Signal

from .models import ResultMessage, AADTSection, CMFEntry

NS = "{http://www.ihsdm.org/schema/Highway-1.0}"


def _find_elements(parent, tag_name):
    """Find elements by tag name, handling XML namespaces."""
    elements = parent.findall(f'.//{NS}{tag_name}')
    if elements:
        return elements
    elements = parent.findall(f'.//{tag_name}')
    if elements:
        return elements
    return [child for child in parent if child.tag.endswith(tag_name)]


def _get_alignment_name(alignment_dir: Path) -> str:
    """Extract alignment name from XML files in the given directory."""
    for xml_name in ("highway.xml", "intersection.xml", "rampterminal.xml", "roundabout.xml"):
        xml_path = alignment_dir / xml_name
        if xml_path.exists():
            try:
                tree = ET.parse(xml_path)
                return tree.getroot().get("title", alignment_dir.name)
            except Exception:
                pass
    # Try versioned files
    for pattern in ("highway.*.xml", "intersection.*.xml", "rampterminal.*.xml", "roundabout.*.xml"):
        matches = sorted(alignment_dir.glob(pattern))
        if matches:
            try:
                tree = ET.parse(matches[-1])
                return tree.getroot().get("title", alignment_dir.name)
            except Exception:
                pass
    return alignment_dir.name


def _format_station(station_str: str) -> str:
    """Format station number with + notation."""
    if not station_str or station_str.strip() == "":
        return station_str
    try:
        val = float(station_str)
        if val < 100:
            return station_str
        parts = str(val).split(".")
        integer = parts[0]
        decimal = parts[1] if len(parts) > 1 else "00"
        decimal = decimal.ljust(2, "0")
        if len(integer) >= 2:
            return f"{integer[:-2]}+{integer[-2:]}.{decimal}"
        return f"{integer}.{decimal}"
    except (ValueError, TypeError):
        return station_str


def _folder_prefix(name: str) -> str:
    """Extract alphabetic prefix from folder name (e.g., 'h' from 'h1', 'ss' from 'ss2')."""
    prefix = ""
    for c in name.lower():
        if c.isalpha():
            prefix += c
        else:
            break
    return prefix


def _normalize_roundabout_headers(rows):
    """Rename 'Roundabout Type' to 'Intersection Type' for sheet consistency."""
    for row in rows:
        for j, val in enumerate(row):
            if val == "Roundabout Type":
                row[j] = "Intersection Type"
                break


def _detect_eval_type_from_csv(eval_dir: Path) -> str:
    """Detect evaluation type from diagnostic CSV section markers.

    IHSDM allows running intersection and ramp terminal evaluations inside
    highway folders.  Returns 'i', 'r', or 'h'.
    """
    for csv_file in eval_dir.glob("evaluation.*.diag.csv"):
        try:
            with open(csv_file, "r", encoding="utf-8") as f:
                for line in f:
                    if "USA Intersection Debug Result" in line:
                        return "i"
                    if "RML Intersection Debug Result" in line:
                        return "i"
                    if "Roundabout Debug Result" in line:
                        return "i"  # Roundabouts treated as intersections
                    if "Ramp Terminal CMF" in line:
                        return "r"
        except Exception:
            pass
    return "h"


# ─────────────────────────────────────────────────────────────────────────────
# Warning Scanner
# ─────────────────────────────────────────────────────────────────────────────


class WarningScanWorker(QThread):
    """Scan project for ResultMessage warnings in evaluation XML files."""

    progress = Signal(int, str)       # percent, message
    finished = Signal(list)           # list[ResultMessage]
    error = Signal(str)

    def __init__(self, project_path: str, parent=None):
        super().__init__(parent)
        self._project_path = project_path

    def run(self):
        try:
            messages: List[ResultMessage] = []
            project_dir = Path(self._project_path)

            # Collect alignment directories with proper prefix detection
            alignment_dirs = [
                d for d in project_dir.iterdir()
                if d.is_dir() and _folder_prefix(d.name) in ("h", "i", "r", "ra")
            ]
            for c_dir in project_dir.iterdir():
                if c_dir.is_dir() and _folder_prefix(c_dir.name) == "c":
                    alignment_dirs.extend(
                        d for d in c_dir.iterdir()
                        if d.is_dir() and _folder_prefix(d.name) in ("h", "i", "r", "ra")
                    )

            total = len(alignment_dirs)
            for idx, alignment_dir in enumerate(alignment_dirs):
                pct = int((idx / max(total, 1)) * 100)
                self.progress.emit(pct, f"Scanning {alignment_dir.name}… ({idx + 1}/{total})")

                alignment_name = _get_alignment_name(alignment_dir)
                folder_type = _folder_prefix(alignment_dir.name)
                eval_dirs = [d for d in alignment_dir.iterdir() if d.is_dir() and d.name.startswith("e")]

                for eval_dir in eval_dirs:
                    # For highway folders, detect if the evaluation is
                    # actually an intersection or ramp terminal evaluation
                    # (IHSDM allows running these inside h-folders).
                    if folder_type == "h":
                        eval_type = _detect_eval_type_from_csv(eval_dir)
                    elif folder_type == "ra":
                        eval_type = "i"  # Roundabouts treated as intersections
                    else:
                        eval_type = folder_type

                    for result_file in eval_dir.glob("evaluation.*.result.xml"):
                        try:
                            tree = ET.parse(result_file)
                            root = tree.getroot()
                            for msg_elem in root.iter("ResultMessage"):
                                message_text = msg_elem.get("message", "")
                                status = msg_elem.get("ResultMessage.status", "info")
                                is_critical = "no crash prediction supported" in message_text.lower()
                                if is_critical:
                                    status = "CRITICAL"
                                messages.append(ResultMessage(
                                    alignment_type=eval_type,
                                    alignment_id=alignment_dir.name,
                                    alignment_name=alignment_name,
                                    evaluation=eval_dir.name,
                                    start_sta=msg_elem.get("startSta", ""),
                                    end_sta=msg_elem.get("endSta", ""),
                                    message=message_text,
                                    status=status,
                                    file_path=str(result_file),
                                    is_critical=is_critical,
                                ))
                        except Exception:
                            continue

            self.progress.emit(100, "Scan complete")
            self.finished.emit(messages)
        except Exception as exc:
            self.error.emit(str(exc))


# ─────────────────────────────────────────────────────────────────────────────
# Data Compiler
# ─────────────────────────────────────────────────────────────────────────────


class CompileWorker(QThread):
    """Compile crash prediction data to Excel."""

    progress = Signal(int, str)
    finished = Signal(str)    # summary text
    error = Signal(str)

    def __init__(self, project_path: str, excel_path: str, target_file: str,
                 target_years: List[str], debug: bool = False, parent=None):
        super().__init__(parent)
        self._project = project_path
        self._excel = excel_path
        self._target = target_file
        self._years = target_years
        self._debug = debug

    def run(self):
        try:
            import ihsdm_compiler_core as compiler

            self.progress.emit(10, "Finding evaluation files…")
            parent_folders = compiler.find_folders_with_file(self._project, self._target)
            if not parent_folders:
                self.error.emit(f"No folders containing '{self._target}' found")
                return

            # Deduplicate parent folders (find_folders_with_file returns one
            # entry per evaluation subfolder, so the same alignment folder can
            # appear multiple times).
            parent_folders = list(dict.fromkeys(parent_folders))

            # ── Highways ──
            self.progress.emit(20, "Processing highway segments…")
            h_folders = [f for f in parent_folders
                         if _folder_prefix(os.path.basename(f).lower()) == "h"]
            all_highway_rows = []
            for pf in h_folders:
                for sf in os.listdir(pf):
                    sp = os.path.join(pf, sf)
                    if not os.path.isdir(sp):
                        continue
                    fp = os.path.join(sp, self._target)
                    if os.path.isfile(fp):
                        eval_name = compiler.get_evaluation_title_from_xml(sp)
                        rows = compiler.extract_highway_segments_from_csv(fp, start_index=5)
                        for row in rows:
                            if compiler.should_process_highway_row(row):
                                filtered = compiler.extract_highway_row_data(row, eval_name, self._debug)
                                if filtered:
                                    all_highway_rows.append(filtered)
            unique_hw, _ = compiler.remove_duplicates(all_highway_rows)
            # Average paired freeway configs (e.g., 6F+8F -> 7F) after dedup
            unique_hw, fw_pairs = compiler.average_freeway_pairs(unique_hw)
            unique_hw.sort(key=lambda x: (x[0], x[1], x[2]))
            compiler.write_rows_to_excel(unique_hw, self._excel, "Highway")
            compiler.add_header_to_excel(self._excel, "Highway", compiler.HIGHWAY_HEADER)
            compiler.fill_missing_highway_values(self._excel)

            # ── Intersections ──
            self.progress.emit(40, "Processing intersections…")
            i_folders = [f for f in parent_folders
                         if _folder_prefix(os.path.basename(f).lower()) == "i"]
            all_int_rows = []
            first_file = False
            for pf in i_folders:
                for sf in os.listdir(pf):
                    sp = os.path.join(pf, sf)
                    if not os.path.isdir(sp):
                        continue
                    fp = os.path.join(sp, self._target)
                    if os.path.isfile(fp):
                        eval_name = compiler.get_evaluation_title_from_xml(sp)
                        rows = compiler.extract_by_headers_from_csv(
                            fp,
                            compiler.INTERSECTION_HEADER[1:-1] + ["Fatal and Injury (FI) Crashes"],
                            first_file=not first_file,
                            target_years=self._years,
                            eval_name=eval_name,
                        )
                        if rows:
                            all_int_rows.extend(rows)
                            first_file = True
            # Also extract intersections from h-folders (USA and RML formats)
            for pf in h_folders:
                for sf in os.listdir(pf):
                    sp = os.path.join(pf, sf)
                    if not os.path.isdir(sp):
                        continue
                    fp = os.path.join(sp, self._target)
                    if os.path.isfile(fp):
                        eval_name = compiler.get_evaluation_title_from_xml(sp)
                        for marker in ("USA Intersection Debug Result",
                                       "RML Intersection Debug Result"):
                            int_rows = compiler.extract_site_set_data(
                                fp,
                                compiler.INTERSECTION_HEADER[1:-1] + ["Fatal and Injury (FI) Crashes"],
                                marker,
                                first_file=(not all_int_rows),
                                eval_name=eval_name,
                                target_years=self._years,
                            )
                            if int_rows:
                                all_int_rows.extend(int_rows)
            # Also extract roundabouts from ra-folders (treated as intersections)
            ra_folders = [f for f in parent_folders
                          if _folder_prefix(os.path.basename(f).lower()) == "ra"]
            for pf in ra_folders:
                for sf in os.listdir(pf):
                    sp = os.path.join(pf, sf)
                    if not os.path.isdir(sp):
                        continue
                    fp = os.path.join(sp, self._target)
                    if os.path.isfile(fp):
                        eval_name = compiler.get_evaluation_title_from_xml(sp)
                        ra_rows = compiler.extract_by_headers_from_csv(
                            fp,
                            compiler.ROUNDABOUT_HEADER[1:-1] + ["Fatal and Injury (FI) Crashes"],
                            first_file=(not all_int_rows),
                            target_years=self._years,
                            eval_name=eval_name,
                        )
                        if ra_rows:
                            _normalize_roundabout_headers(ra_rows)
                            all_int_rows.extend(ra_rows)
            # Also extract roundabouts from h-folders
            for pf in h_folders:
                for sf in os.listdir(pf):
                    sp = os.path.join(pf, sf)
                    if not os.path.isdir(sp):
                        continue
                    fp = os.path.join(sp, self._target)
                    if os.path.isfile(fp):
                        eval_name = compiler.get_evaluation_title_from_xml(sp)
                        ra_rows = compiler.extract_site_set_data(
                            fp,
                            compiler.ROUNDABOUT_HEADER[1:-1] + ["Fatal and Injury (FI) Crashes"],
                            "Roundabout Debug Result",
                            first_file=(not all_int_rows),
                            eval_name=eval_name,
                            target_years=self._years,
                        )
                        if ra_rows:
                            _normalize_roundabout_headers(ra_rows)
                            all_int_rows.extend(ra_rows)
            unique_int = []
            if all_int_rows:
                unique_int = compiler.deduplicate_by_title(all_int_rows, title_column_index=3)
                compiler.write_rows_to_excel(unique_int, self._excel, "Intersection")
                compiler.fill_missing_intersection_values(self._excel)
                compiler.scrub_duplicate_columns(self._excel, "Intersection")

            # ── Ramp Terminals ──
            self.progress.emit(60, "Processing ramp terminals…")
            r_folders = [f for f in parent_folders
                         if _folder_prefix(os.path.basename(f).lower()) == "r"]
            all_ramp_rows = []
            first_file = False
            for pf in r_folders:
                for sf in os.listdir(pf):
                    sp = os.path.join(pf, sf)
                    if not os.path.isdir(sp):
                        continue
                    fp = os.path.join(sp, self._target)
                    if os.path.isfile(fp):
                        eval_name = compiler.get_evaluation_title_from_xml(sp)
                        rows = compiler.extract_by_headers_from_csv(
                            fp,
                            compiler.RAMP_TERMINAL_HEADER[1:-1] + ["Fatal and Injury (FI) Crashes"],
                            first_file=not first_file,
                            target_years=self._years,
                            eval_name=eval_name,
                        )
                        if rows:
                            all_ramp_rows.extend(rows)
                            first_file = True
            # Also extract ramp terminals from h-folders (IHSDM allows
            # running intersection and ramp terminal evals inside highway
            # alignments).
            for pf in h_folders:
                for sf in os.listdir(pf):
                    sp = os.path.join(pf, sf)
                    if not os.path.isdir(sp):
                        continue
                    fp = os.path.join(sp, self._target)
                    if os.path.isfile(fp):
                        eval_name = compiler.get_evaluation_title_from_xml(sp)
                        ramp_rows = compiler.extract_site_set_data(
                            fp,
                            compiler.RAMP_TERMINAL_HEADER[1:-1] + ["Fatal and Injury (FI) Crashes"],
                            "Ramp Terminal CMF",
                            first_file=(not all_ramp_rows),
                            eval_name=eval_name,
                            target_years=self._years,
                        )
                        if ramp_rows:
                            all_ramp_rows.extend(ramp_rows)
            # Deduplicate ramp terminals by title (same as intersections)
            if all_ramp_rows:
                all_ramp_rows = compiler.deduplicate_by_title(all_ramp_rows, title_column_index=3)
                compiler.write_rows_to_excel(all_ramp_rows, self._excel, "RampTerminal")
                compiler.fill_missing_ramp_terminal_values(self._excel)
                compiler.scrub_duplicate_columns(self._excel, "RampTerminal")

            # ── Site Sets ──
            self.progress.emit(80, "Processing site sets…")
            ss_folders = [f for f in parent_folders
                          if _folder_prefix(os.path.basename(f).lower()) == "ss"]
            ss_int_rows: list = []
            ss_ramp_rows: list = []
            ss_other_rows: list = []
            ss_first_int = True
            ss_first_ramp = True
            ss_first_other = True
            unknown_sections: list = []
            known_markers = ["USA Intersection Debug Result", "RML Intersection Debug Result",
                            "Ramp Terminal CMF", "Roundabout Debug Result"]

            for pf in ss_folders:
                for sf in os.listdir(pf):
                    sp = os.path.join(pf, sf)
                    if not os.path.isdir(sp):
                        continue
                    fp = os.path.join(sp, self._target)
                    if os.path.isfile(fp):
                        eval_name = compiler.get_evaluation_title_from_xml(sp)
                        for marker in ("USA Intersection Debug Result",
                                       "RML Intersection Debug Result"):
                            int_rows = compiler.extract_site_set_data(
                                fp, compiler.SITESET_INT_HEADER[1:],
                                marker,
                                first_file=ss_first_int, eval_name=eval_name,
                                target_years=self._years,
                            )
                            if int_rows:
                                ss_int_rows.extend(int_rows)
                                ss_first_int = False
                        # Extract roundabouts from site sets (merge into intersection rows)
                        ra_ss_rows = compiler.extract_site_set_data(
                            fp, compiler.SITESET_ROUNDABOUT_HEADER[1:],
                            "Roundabout Debug Result",
                            first_file=ss_first_int, eval_name=eval_name,
                            target_years=self._years,
                        )
                        if ra_ss_rows:
                            _normalize_roundabout_headers(ra_ss_rows)
                            ss_int_rows.extend(ra_ss_rows)
                            ss_first_int = False
                        ramp_rows = compiler.extract_site_set_data(
                            fp, compiler.SITESET_RAMP_HEADER[1:],
                            "Ramp Terminal CMF",
                            first_file=ss_first_ramp, eval_name=eval_name,
                            target_years=self._years,
                        )
                        if ramp_rows:
                            ss_ramp_rows.extend(ramp_rows)
                            ss_first_ramp = False
                        unknown_data = compiler.extract_unknown_site_set_sections(
                            fp, known_markers, target_years=self._years
                        )
                        for section_name, header_row, data_rows in unknown_data:
                            if section_name not in unknown_sections:
                                unknown_sections.append(section_name)
                            if ss_first_other:
                                ss_other_rows.append(["Section Type"] + header_row)
                                ss_first_other = False
                            for dr in data_rows:
                                ss_other_rows.append([section_name] + dr)

            if ss_int_rows:
                compiler.write_rows_to_excel(ss_int_rows, self._excel, "SiteSet_Int")
                compiler.fill_missing_intersection_values(self._excel, "SiteSet_Int")
                compiler.scrub_duplicate_columns(self._excel, "SiteSet_Int")
            if ss_ramp_rows:
                compiler.write_rows_to_excel(ss_ramp_rows, self._excel, "SiteSet_Ramp")
                compiler.fill_missing_ramp_terminal_values(self._excel, "SiteSet_Ramp")
                compiler.scrub_duplicate_columns(self._excel, "SiteSet_Ramp")
            if ss_other_rows:
                compiler.write_rows_to_excel(ss_other_rows, self._excel, "SiteSet_Other")

            self.progress.emit(100, "Done")

            summary = (
                f"Highway Segments: {len(unique_hw)}\n"
                f"Intersections: {len(unique_int)}\n"
                f"Ramp Terminals: {len(all_ramp_rows)}\n"
                f"Site Set Intersections: {len(ss_int_rows)}\n"
                f"Site Set Ramp Terminals: {len(ss_ramp_rows)}"
            )
            if fw_pairs:
                summary += f"\nFreeway Pairs Averaged: {fw_pairs}"
            if unknown_sections:
                summary += f"\nSite Set Other: {len(ss_other_rows)}"
            self.finished.emit(summary)

        except Exception as exc:
            import traceback
            traceback.print_exc()
            self.error.emit(str(exc))


# ─────────────────────────────────────────────────────────────────────────────
# Appendix PDF Merge
# ─────────────────────────────────────────────────────────────────────────────


class AppendixMergeWorker(QThread):
    """Merge selected evaluation report PDFs into one."""

    progress = Signal(int, str)
    log = Signal(str)
    finished = Signal(int)   # number of PDFs merged
    error = Signal(str)

    def __init__(self, pdf_paths: List[str], output_path: str, parent=None):
        super().__init__(parent)
        self._paths = pdf_paths
        self._output = output_path

    def run(self):
        try:
            from PyPDF2 import PdfMerger
        except ImportError:
            self.error.emit("PyPDF2 is required.  Install with:  pip install PyPDF2")
            return
        try:
            merger = PdfMerger()
            total = len(self._paths)
            added = 0
            for idx, pdf in enumerate(sorted(self._paths)):
                pct = int((idx / max(total, 1)) * 100)
                self.progress.emit(pct, f"Merging {Path(pdf).name}…")
                try:
                    merger.append(pdf)
                    self.log.emit(f"  Added: {Path(pdf).name}")
                    added += 1
                except Exception as exc:
                    self.log.emit(f"  WARNING: {Path(pdf).name}: {exc}")
            self.progress.emit(95, "Writing output…")
            merger.write(self._output)
            merger.close()
            self.progress.emit(100, "Done")
            self.finished.emit(added)
        except Exception as exc:
            self.error.emit(str(exc))


# ─────────────────────────────────────────────────────────────────────────────
# CMF / Evaluation Info Scanner
# ─────────────────────────────────────────────────────────────────────────────


class CMFScanWorker(QThread):
    """Scan project for evaluation.1.cpm.cmf.csv files."""

    progress = Signal(int, str)
    finished = Signal(list)    # list[CMFEntry]
    error = Signal(str)

    def __init__(self, project_path: str, parent=None):
        super().__init__(parent)
        self._project = project_path

    def run(self):
        try:
            cmf_files = []
            for root, dirs, files in os.walk(self._project):
                if "evaluation.1.cpm.cmf.csv" in files:
                    cmf_files.append(os.path.join(root, "evaluation.1.cpm.cmf.csv"))

            if not cmf_files:
                self.finished.emit([])
                return

            entries: List[CMFEntry] = []
            total = len(cmf_files)

            for idx, cmf_path in enumerate(cmf_files):
                pct = int((idx / max(total, 1)) * 100)
                self.progress.emit(pct, f"Reading CMF {idx + 1}/{total}…")

                path_parts = cmf_path.replace("\\", "/").split("/")
                alignment_id = None
                alignment_type = None
                eval_folder = None

                for i, part in enumerate(path_parts):
                    pfx = _folder_prefix(part)
                    num = part[len(pfx):]
                    if pfx and num.isdigit():
                        if pfx == "h":
                            alignment_id, alignment_type = part, "Highway"
                        elif pfx in ("i", "ra"):
                            alignment_id, alignment_type = part, "Intersection"
                        elif pfx == "r":
                            alignment_id, alignment_type = part, "Ramp Terminal"
                    if alignment_id and i + 1 < len(path_parts):
                        eval_folder = path_parts[i + 1]
                        break

                if not alignment_id:
                    continue

                calibration = "Not found"
                alignment_name = "Unknown"
                eval_years = "Unknown"

                try:
                    with open(cmf_path, "r", encoding="utf-8") as f:
                        lines = f.readlines()
                    if len(lines) >= 27:
                        parts = lines[26].strip().split(",")
                        if len(parts) >= 2:
                            calibration = parts[1].strip('"')
                    if len(lines) >= 17:
                        parts = lines[16].strip().split(",")
                        if len(parts) >= 2:
                            alignment_name = parts[1].strip('"')
                except Exception:
                    pass

                # Parse eval years from result XML
                try:
                    result_xml = os.path.join(os.path.dirname(cmf_path), "evaluation.1.result.xml")
                    if os.path.exists(result_xml):
                        tree = ET.parse(result_xml)
                        start_year = end_year = None
                        for elem in tree.getroot().iter():
                            if "evalStartYear" in elem.attrib and not start_year:
                                start_year = elem.attrib["evalStartYear"]
                            if "evalEndYear" in elem.attrib and not end_year:
                                end_year = elem.attrib["evalEndYear"]
                            if start_year and end_year:
                                break
                        if start_year and end_year:
                            eval_years = start_year if start_year == end_year else f"{start_year}-{end_year}"
                except Exception:
                    pass

                entries.append(CMFEntry(
                    type=alignment_type,
                    id=alignment_id,
                    name=alignment_name,
                    evaluation=eval_folder or "Unknown",
                    years=eval_years,
                    calibration=calibration,
                    path=cmf_path,
                ))

            self.progress.emit(100, "Done")
            self.finished.emit(entries)
        except Exception as exc:
            self.error.emit(str(exc))


# ─────────────────────────────────────────────────────────────────────────────
# AADT Section Scanner
# ─────────────────────────────────────────────────────────────────────────────


class AADTScanWorker(QThread):
    """Scan project for AADT sections in highway XML files."""

    progress = Signal(int, str)
    finished = Signal(list)     # list[AADTSection]
    error = Signal(str)

    def __init__(self, project_path: str, parent=None):
        super().__init__(parent)
        self._project = project_path

    def run(self):
        try:
            project_dir = Path(self._project)
            highway_dirs = [
                d for d in project_dir.iterdir()
                if d.is_dir() and _folder_prefix(d.name) == "h"
            ]
            for c_dir in project_dir.iterdir():
                if c_dir.is_dir() and _folder_prefix(c_dir.name) == "c":
                    highway_dirs.extend(
                        d for d in c_dir.iterdir()
                        if d.is_dir() and _folder_prefix(d.name) == "h"
                    )

            sections: List[AADTSection] = []
            total = len(highway_dirs)

            for idx, hw_dir in enumerate(highway_dirs):
                pct = int((idx / max(total, 1)) * 100)
                self.progress.emit(pct, f"Scanning {hw_dir.name}…")

                hw_xmls = sorted(hw_dir.glob("highway.*.xml"))
                if not hw_xmls:
                    hw_xmls = sorted(hw_dir.glob("highway.xml"))
                if not hw_xmls:
                    continue

                hw_xml = hw_xmls[-1]
                try:
                    tree = ET.parse(hw_xml)
                    root = tree.getroot()
                    roadway = root.find(f".//{NS}Roadway")
                    if roadway is None:
                        roadway = root.find(".//Roadway")
                    if roadway is None:
                        continue
                    title = roadway.get("title", hw_dir.name)

                    aadt_elems = _find_elements(roadway, "AnnualAveDailyTraffic")
                    for sec_num, aadt in enumerate(aadt_elems, start=1):
                        sections.append(AADTSection(
                            roadway_title=title,
                            highway_dir=str(hw_dir),
                            xml_file=str(hw_xml),
                            section_num=sec_num,
                            start_station=aadt.get("startStation", ""),
                            end_station=aadt.get("endStation", ""),
                            year=aadt.get("adtYear", ""),
                            current_aadt=aadt.get("adtRate", "1"),
                        ))
                except Exception:
                    continue

            self.progress.emit(100, "Done")
            self.finished.emit(sections)
        except Exception as exc:
            self.error.emit(str(exc))


# ─────────────────────────────────────────────────────────────────────────────
# Year Scanner (for Data Compiler year selection)
# ─────────────────────────────────────────────────────────────────────────────


class YearScanWorker(QThread):
    """Scan evaluation CSVs to find available years with per-alignment info."""

    progress = Signal(int, str)
    finished = Signal(dict)   # {"all_years": sorted list, "alignment_years": {(folder,name): set(years)}}
    error = Signal(str)

    def __init__(self, project_path: str, parent=None):
        super().__init__(parent)
        self._project = project_path

    def run(self):
        try:
            project_dir = Path(self._project)
            csv_files = list(project_dir.glob("**/evaluation.*.diag.csv"))
            if not csv_files:
                self.finished.emit({"all_years": [], "alignment_years": {}})
                return

            all_years: set = set()
            alignment_years: Dict[tuple, set] = {}
            total = len(csv_files)

            for idx, csv_file in enumerate(csv_files):
                pct = int((idx / max(total, 1)) * 100)
                self.progress.emit(pct, f"Reading {csv_file.name}…")

                try:
                    alignment_folder = csv_file.parent.parent
                    folder_id = alignment_folder.name
                    if _folder_prefix(folder_id) not in ("h", "i", "r", "ra", "ss"):
                        continue

                    alignment_name = _get_alignment_name(alignment_folder)
                    file_years: set = set()

                    with open(csv_file, "r", encoding="utf-8") as f:
                        import csv as csv_mod
                        lines = list(csv_mod.reader(f))

                    for i, row in enumerate(lines):
                        if "Year" in row:
                            year_col = row.index("Year")
                            for data_row in lines[i + 1:]:
                                if not data_row or "*************" in str(data_row):
                                    break
                                if len(data_row) > year_col:
                                    yv = data_row[year_col].strip()
                                    if yv.isdigit() and len(yv) == 4:
                                        file_years.add(yv)
                                        all_years.add(yv)

                    key = (folder_id, alignment_name)
                    alignment_years.setdefault(key, set()).update(file_years)
                except Exception:
                    continue

            self.progress.emit(100, "Done")
            self.finished.emit({
                "all_years": sorted(all_years),
                "alignment_years": {k: sorted(v) for k, v in alignment_years.items()},
            })
        except Exception as exc:
            self.error.emit(str(exc))


# ─────────────────────────────────────────────────────────────────────────────
# Visual View data loader
# ─────────────────────────────────────────────────────────────────────────────


class VisualDataWorker(QThread):
    """Parse highway.1.xml and return all visualization data."""

    finished = Signal(dict)   # big dict of all parsed sections
    error = Signal(str)

    def __init__(self, highway_xml_path: str, project_dir: str, parent=None):
        super().__init__(parent)
        self._xml = highway_xml_path
        self._project = project_dir

    def run(self):
        try:
            tree = ET.parse(self._xml)
            root = tree.getroot()
            roadway = root.find(f".//{NS}Roadway")
            if roadway is None:
                roadway = root.find(".//Roadway")
            if roadway is None:
                self.error.emit("Could not find Roadway element")
                return

            min_sta = float(roadway.get("minStation", "0.0"))
            max_sta = float(roadway.get("maxStation", "0.0"))
            title = roadway.get("title", "Unknown")

            data = {
                "title": title,
                "min_sta": min_sta,
                "max_sta": max_sta,
                "heading_sta": roadway.get("headingSta", "0.0"),
                "heading_angle": roadway.get("headingAngle", "0.0"),
                "lanes": self._parse_lanes(roadway, min_sta, max_sta),
                "shoulders": self._parse_shoulders(roadway, min_sta, max_sta),
                "ramps": self._parse_ramps(roadway),
                "curves": self._parse_curves(roadway),
                "traffic": self._parse_traffic(roadway, min_sta, max_sta),
                "median": self._parse_median(roadway, min_sta, max_sta),
                "speed": self._parse_speed(roadway, min_sta, max_sta),
                "func_class": self._parse_func_class(roadway, min_sta, max_sta),
                "intersections": self._parse_intersections(
                    self._project, roadway.get("nodeName", "")
                ),
            }
            self.finished.emit(data)
        except Exception as exc:
            self.error.emit(str(exc))

    # ── parsers ──────────────────────────────────────────────────────────

    def _parse_lanes(self, roadway, min_sta, max_sta):
        lanes = []
        for lane in _find_elements(roadway, "LaneNS"):
            try:
                sw = float(lane.get("startWidth", "12.0"))
                ew = float(lane.get("endWidth", str(sw)))
                lanes.append({
                    "begin": float(lane.get("startStation", str(min_sta))),
                    "end": float(lane.get("endStation", str(max_sta))),
                    "side": lane.get("sideOfRoad", "both"),
                    "priority": int(lane.get("priority", "10")),
                    "lane_type": lane.get("laneType", "thru"),
                    "width": (sw + ew) / 2,
                })
            except (ValueError, AttributeError):
                continue
        return sorted(lanes, key=lambda x: (x["begin"], x["side"], x["priority"]))

    def _parse_shoulders(self, roadway, min_sta, max_sta):
        shoulders = []
        for sh in _find_elements(roadway, "ShoulderSection"):
            try:
                sw = float(sh.get("startWidth", "0.0"))
                ew = float(sh.get("endWidth", str(sw)))
                shoulders.append({
                    "begin": float(sh.get("startStation", str(min_sta))),
                    "end": float(sh.get("endStation", str(max_sta))),
                    "side": sh.get("sideOfRoad", "right"),
                    "priority": int(sh.get("priority", "100")),
                    "width": (sw + ew) / 2,
                    "position": sh.get("insideOutsideOfRoadNB", "outside"),
                    "material": sh.get("material", "paved"),
                })
            except (ValueError, AttributeError):
                continue
        return sorted(shoulders, key=lambda x: (x["begin"], x["side"], x["priority"]))

    def _parse_ramps(self, roadway):
        ramps = []
        for r in _find_elements(roadway, "RampConnector"):
            try:
                ramps.append({
                    "station": float(r.get("station", "0.0")),
                    "name": r.get("name", "Ramp"),
                    "ramp_type": r.get("type", "entrance"),
                })
            except (ValueError, AttributeError):
                continue
        return sorted(ramps, key=lambda x: x["station"])

    def _parse_curves(self, roadway):
        curves = []
        h_elements_list = _find_elements(roadway, "HorizontalElements")
        if not h_elements_list:
            return curves
        h_elem = h_elements_list[0]
        for tangent in _find_elements(h_elem, "HTangent"):
            try:
                curves.append({
                    "type": "tangent",
                    "begin": float(tangent.get("startStation", "0.0")),
                    "end": float(tangent.get("endStation", "0.0")),
                })
            except (ValueError, AttributeError):
                continue
        for curve in _find_elements(h_elem, "HSimpleCurve"):
            try:
                curves.append({
                    "type": "curve",
                    "begin": float(curve.get("startStation", "0.0")),
                    "end": float(curve.get("endStation", "0.0")),
                    "radius": float(curve.get("radius", "0.0")),
                    "direction": curve.get("curveDirection", "left"),
                })
            except (ValueError, AttributeError):
                continue
        for spiral in _find_elements(h_elem, "HSpiralCurve"):
            try:
                curves.append({
                    "type": "spiral",
                    "begin": float(spiral.get("startStation", "0.0")),
                    "end": float(spiral.get("endStation", "0.0")),
                    "radius": float(spiral.get("radius", "0.0")),
                })
            except (ValueError, AttributeError):
                continue
        return sorted(curves, key=lambda x: x["begin"])

    def _parse_traffic(self, roadway, min_sta, max_sta):
        traffic = []
        for aadt in _find_elements(roadway, "AnnualAveDailyTraffic"):
            try:
                traffic.append({
                    "begin": float(aadt.get("startStation", str(min_sta))),
                    "end": float(aadt.get("endStation", str(max_sta))),
                    "volume": int(float(aadt.get("adtRate", "0"))),
                })
            except (ValueError, AttributeError):
                continue
        return sorted(traffic, key=lambda x: x["begin"])

    def _parse_median(self, roadway, min_sta, max_sta):
        medians = []
        for med in _find_elements(roadway, "Median"):
            try:
                medians.append({
                    "begin": float(med.get("startStation", str(min_sta))),
                    "end": float(med.get("endStation", str(max_sta))),
                    "width": float(med.get("width", "0.0")),
                    "median_type": med.get("medianType", "none"),
                })
            except (ValueError, AttributeError):
                continue
        return sorted(medians, key=lambda x: x["begin"])

    def _parse_speed(self, roadway, min_sta, max_sta):
        speeds = []
        for sp in _find_elements(roadway, "PostedSpeed"):
            try:
                speeds.append({
                    "begin": float(sp.get("startStation", str(min_sta))),
                    "end": float(sp.get("endStation", str(max_sta))),
                    "speed": int(float(sp.get("speedLimit", "0"))),
                })
            except (ValueError, AttributeError):
                continue
        return sorted(speeds, key=lambda x: x["begin"])

    def _parse_func_class(self, roadway, min_sta, max_sta):
        fcs = []
        for fc in _find_elements(roadway, "FunctionalClass"):
            try:
                fcs.append({
                    "begin": float(fc.get("startStation", str(min_sta))),
                    "end": float(fc.get("endStation", str(max_sta))),
                    "class_type": fc.get("funcClass", "unknown"),
                })
            except (ValueError, AttributeError):
                continue
        return sorted(fcs, key=lambda x: x["begin"])

    def _parse_intersections(self, project_dir: str, highway_node_name: str):
        """Parse intersection connections from the project's intersection XML files."""
        connections = []
        found_stations: set = set()
        hw_suffix = highway_node_name.split(".")[-1] if "." in highway_node_name else highway_node_name

        def _search(int_dir, dir_name):
            for xml_name in ("intersection.1.xml", "intersection.xml",
                             "roundabout.1.xml", "roundabout.xml"):
                xml_path = os.path.join(int_dir, xml_name)
                if not os.path.exists(xml_path):
                    continue
                try:
                    tree = ET.parse(xml_path)
                    root = tree.getroot()
                    for ns in (
                        "{http://www.ihsdm.org/schema/Intersection-1.0}",
                        NS, "",
                    ):
                        intersection = root.find(f"{ns}Intersection")
                        if intersection is not None:
                            int_name = intersection.get("intersectionName", f"Intersection {dir_name}")
                            for ns2 in (
                                "{http://www.ihsdm.org/schema/Intersection-1.0}",
                                NS, "",
                            ):
                                for leg in intersection.findall(f"{ns2}Leg"):
                                    leg_suffix = leg.get("highwayNodeName", "").split(".")[-1]
                                    if leg_suffix == hw_suffix:
                                        station = float(leg.get("highwayStation", "0.0"))
                                        if station not in found_stations:
                                            found_stations.add(station)
                                            connections.append({
                                                "station": station,
                                                "name": int_name,
                                                "type": "intersection",
                                            })
                            break
                except Exception:
                    continue

        try:
            for item in os.listdir(project_dir):
                item_path = os.path.join(project_dir, item)
                if os.path.isdir(item_path):
                    pfx = _folder_prefix(item)
                    num = item[len(pfx):]
                    if pfx in ("i", "ra") and num.isdigit():
                        _search(item_path, item)
                    elif pfx == "c" and num.isdigit():
                        try:
                            for sub in os.listdir(item_path):
                                sub_pfx = _folder_prefix(sub)
                                sub_num = sub[len(sub_pfx):]
                                if sub_pfx in ("i", "ra") and sub_num.isdigit():
                                    _search(os.path.join(item_path, sub), f"{item}/{sub}")
                        except Exception:
                            continue
        except Exception:
            pass

        return sorted(connections, key=lambda x: x["station"])


# ─────────────────────────────────────────────────────────────────────────────
# GitHub update check (fire-and-forget on startup)
# ─────────────────────────────────────────────────────────────────────────────


class UpdateCheckWorker(QThread):
    """Check GitHub API for a newer release (non-blocking)."""

    update_available = Signal(str, str, str)  # version, download_url, notes
    no_update = Signal()
    check_failed = Signal()

    def __init__(self, api_url: str, current_version: str, parent=None):
        super().__init__(parent)
        self._url = api_url
        self._current = current_version

    def run(self):
        if not self._url:
            self.check_failed.emit()
            return
        try:
            req = Request(self._url, headers={"Accept": "application/vnd.github.v3+json"})
            with urlopen(req, timeout=5) as resp:
                import json
                data = json.loads(resp.read().decode())
            tag = data.get("tag_name", "").lstrip("v")
            if not tag:
                self.no_update.emit()
                return
            if self._compare(tag, self._current) > 0:
                dl = data.get("html_url", "")
                notes = data.get("body", "")
                self.update_available.emit(tag, dl, notes)
            else:
                self.no_update.emit()
        except Exception:
            self.check_failed.emit()

    @staticmethod
    def _compare(v1: str, v2: str) -> int:
        p1 = [int(x) for x in v1.split(".")]
        p2 = [int(x) for x in v2.split(".")]
        for a, b in zip(p1, p2):
            if a > b:
                return 1
            if a < b:
                return -1
        return (len(p1) > len(p2)) - (len(p1) < len(p2))
