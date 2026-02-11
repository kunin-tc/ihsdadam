"""Data models used across the application"""

from dataclasses import dataclass, field
from typing import List, Dict, Set


@dataclass
class ResultMessage:
    """Container for a result message from IHSDM evaluation"""
    alignment_type: str   # 'h' for highway, 'i' for intersection, 'r' for ramp terminal
    alignment_id: str     # e.g., 'h74', 'i100', 'r25'
    alignment_name: str   # e.g., 'Mainline - I-375 Centerline'
    evaluation: str       # e.g., 'e20'
    start_sta: str
    end_sta: str
    message: str
    status: str           # 'info', 'warning', 'error', 'fault', 'CRITICAL'
    file_path: str
    is_critical: bool = False


# HSM severity distributions — highway
HSM_SEVERITY_K = 0.0146
HSM_SEVERITY_A = 0.044764
HSM_SEVERITY_B = 0.2469
HSM_SEVERITY_C = 0.69172

# HSM severity distributions — intersection / ramp terminal
INT_SEVERITY_K = 0.002575
INT_SEVERITY_A = 0.053525
INT_SEVERITY_B = 0.276415
INT_SEVERITY_C = 0.667485


@dataclass
class AADTSection:
    """A single AADT station-range entry parsed from highway XML"""
    roadway_title: str
    highway_dir: str
    xml_file: str
    section_num: int
    start_station: str
    end_station: str
    year: str
    current_aadt: str
    # Forecast ID slots (up to 6) with +/- signs
    id1: str = ""
    sign1: str = "+"
    id2: str = ""
    sign2: str = "+"
    id3: str = ""
    sign3: str = "+"
    id4: str = ""
    sign4: str = "+"
    id5: str = ""
    sign5: str = "+"
    id6: str = ""
    sign6: str = "+"
    calculated_aadt: str = ""
    is_new: bool = False


@dataclass
class CMFEntry:
    """A calibration factor entry from evaluation.1.cpm.cmf.csv"""
    type: str          # 'Highway', 'Intersection', 'Ramp Terminal'
    id: str            # e.g. 'h1', 'i100'
    name: str
    evaluation: str    # e.g. 'e1'
    years: str         # e.g. '2028' or '2028-2050'
    calibration: str
    path: str
