from __future__ import annotations

import glob
import inspect
import json
import os
import re
import subprocess
import threading
import time
from collections import defaultdict
from dataclasses import asdict, dataclass, field
from datetime import datetime
from multiprocessing import Process
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    from docx import Document
    from docx.shared import RGBColor
except ImportError:
    Document = None
    RGBColor = None

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
except ImportError:
    colors = None
    letter = None
    getSampleStyleSheet = None
    Paragraph = None
    SimpleDocTemplate = None
    Spacer = None
    Table = None
    TableStyle = None

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None


APP_FONT = ("Arial", 10)
DEFAULT_TRACE_DIR = r"C:\Program Files (x86)\HAMILTON\LogFiles"
DEFAULT_RESULTS_DIR = r"C:\Results\Sequence Analysis"

ROLE_OPTIONS = ["tip pick up", "aspirate", "dispense"]
TIP_EJECT_ROLE = "tip eject"
ALL_ROLE_OPTIONS = ROLE_OPTIONS + [TIP_EJECT_ROLE]
TRANSFER_OPTIONS = ["samples", "buffer", "STD/QCs"]
REPLICATE_OPTIONS = ["single", "duplicate", "triplicate"]
REPLICATE_MAP = {"single": 1, "duplicate": 2, "triplicate": 3}
REPLICATE_NAME_BY_COUNT = {1: "single", 2: "duplicate", 3: "triplicate"}

UI_BG = "#000000"
UI_TEXT_BG = "#2b004f"
UI_TEXT_FG = "white"
UI_BUTTON_BG = "#2b004f"
UI_BUTTON_ACTIVE = "#4b0082"

PYHAMILTON_IMPORT_ERROR = ""

try:
    import pyhamilton
    import pyhamilton.interface as phi
    from pyhamilton import (
        HamiltonError,
        HamiltonStepError,
        HamiltonTimeoutError,
        LayoutManager,
        OEM_HSL_PATH,
        OEM_RUN_EXE_PATH,
        Plate96,
        ResourceType,
        initialize,
        normal_logging,
    )
except Exception as exc:
    PYHAMILTON_IMPORT_ERROR = f"{type(exc).__name__}: {exc}"
    pyhamilton = None
    phi = None
    HamiltonStepError = Exception
    HamiltonError = Exception
    HamiltonTimeoutError = Exception
    LayoutManager = None
    Plate96 = None
    ResourceType = None
    OEM_RUN_EXE_PATH = ""
    OEM_HSL_PATH = ""


def _repair_pyhamilton_exception_bindings() -> None:
    """
    Some PyHamilton builds reference HamiltonStepError (and related classes)
    from inside pyhamilton.interface without binding those names in that module.
    The robot step then fails correctly in VENUS, but Python masks the useful
    Hamilton error with ``NameError: HamiltonStepError is not defined``.

    Bind the installed exception classes back into the interface module so the
    original Hamilton exception is preserved and reported.  This is a no-op on
    PyHamilton versions that already expose the names correctly.
    """
    if pyhamilton is None or phi is None:
        return
    try:
        import pyhamilton.oemerr as _oemerr
    except Exception:
        return

    for _name in (
        "HamiltonStepError",
        "HamiltonError",
        "HamiltonTimeoutError",
        "HamiltonReturnParseError",
        "InvalidErrCodeError",
    ):
        _value = getattr(_oemerr, _name, None)
        if _value is not None:
            setattr(phi, _name, _value)
            globals()[_name] = _value


_repair_pyhamilton_exception_bindings()


@dataclass
class SequenceStepConfig:
    sequence: str
    role: str
    liquid_class: str
    transfer_type: str
    replicate: str
    volume_ul: float
    marker: str
    channel_pattern: str = ""
    hardware_mode: str = "channels"
    selected_index: int = 0
    sequence_count: int = 0
    manual: bool = False
    set_order: int = 1
    control: bool = False


@dataclass
class TransferPair:
    pair_id: str
    tip_sequence: str = ""
    source_sequence: str = ""
    destination_sequence: str = ""
    source_position: str = ""
    destination_position: str = ""
    source_sample_number: Optional[int] = None
    destination_sample_number: Optional[int] = None
    replicate_index: int = 1
    volume_ul: float = 0.0
    liquid_class: str = ""
    transfer_type: str = ""
    marker_label: str = ""
    template_marker: str = ""
    plate_cycle: int = 1


@dataclass
class ValidationIssue:
    severity: str
    category: str
    message: str
    evidence: Dict[str, Any] = field(default_factory=dict)


@dataclass
class PlateLayoutPattern:
    marker_prefix: str
    marker_numbers: List[int]
    wells_by_marker_number: Dict[int, List[str]]
    occurrences_by_marker_number: Dict[int, int]
    samples_per_plate: int
    replicate_count: int

    @property
    def replicate_name(self) -> str:
        return REPLICATE_NAME_BY_COUNT.get(self.replicate_count, f"{self.replicate_count}x")


@dataclass
class DevPalLiteOutput:
    lay_file: str
    layout_file: str
    trace_file: str
    selected_sequences: List[Dict[str, Any]]
    source_positions: List[str]
    destination_positions: List[str]
    transfer_pairs: List[Dict[str, Any]]
    execution_plan: List[Dict[str, Any]]
    marker_matches: Dict[str, List[Dict[str, Any]]]
    layout_findings: List[Dict[str, Any]]
    trace_findings: List[Dict[str, Any]]
    validation_findings: List[Dict[str, Any]]
    issues: List[Dict[str, Any]]
    limitations: List[str]
    generated_review_script: str
    report_path: str


def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except Exception:
        return default


def excel_col_name(index_zero_based: int) -> str:
    name = ""
    index = index_zero_based + 1
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def normalize_position_text(value: Any) -> str:
    return str(value or "").strip().upper().replace(" ", "").replace(":", "")


def extract_position_number(position_text: Any) -> Optional[int]:
    match = re.search(r"(\d+)$", str(position_text or "").strip())
    return int(match.group(1)) if match else None


def read_text_file_best_effort(path: str) -> str:
    for encoding in ("utf-8", "latin-1", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, errors="ignore") as handle:
                return handle.read()
        except Exception:
            continue
    return ""


def clean_control_chars(text: str) -> str:
    return re.sub(r"[\x00-\x1f]+", " ", str(text)).strip()


def unique_issue_messages(issues: List[ValidationIssue]) -> List[str]:
    seen = set()
    result = []
    for issue in issues:
        message = str(issue.message or "").strip()
        if message and message not in seen:
            seen.add(message)
            result.append(message)
    return result


def add_issue(
    issues: List[ValidationIssue],
    severity: str,
    category: str,
    message: str,
    evidence: Optional[Dict[str, Any]] = None,
) -> None:
    issues.append(ValidationIssue(severity, category, message, evidence or {}))


def normalize_assigned_channel_pattern(channel_pattern: str, hardware_mode: str = "channels") -> str:
    pattern = str(channel_pattern or "").strip()
    if str(hardware_mode).lower() == "head":
        return "1" * 96
    if not re.fullmatch(r"[01]+", pattern):
        raise ValueError(f"Invalid channel pattern: {pattern}")
    return pattern


def count_active_channels(channel_pattern: str) -> int:
    return sum(1 for char in str(channel_pattern or "") if char == "1")


def _import_optional_pyhamilton_module(module_name: str):
    try:
        return __import__(module_name, fromlist=["*"])
    except Exception:
        return None


PYHAMILTON_MODULES = []
if pyhamilton is not None:
    PYHAMILTON_MODULES.extend(
        module
        for module in (
            pyhamilton,
            phi,
            _import_optional_pyhamilton_module("pyhamilton.liquid_handling_wrappers"),
            _import_optional_pyhamilton_module("pyhamilton.defaultcmds"),
            _import_optional_pyhamilton_module("pyhamilton.oemerr"),
        )
        if module is not None
    )


def _get_optional_pyhamilton_function(*names):
    for module in PYHAMILTON_MODULES:
        for name in names:
            if hasattr(module, name):
                return getattr(module, name)
    return None


def _call_best_effort(func, *args, **kwargs):
    if func is None:
        raise RuntimeError("PyHamilton wrapper function is unavailable.")
    try:
        signature = inspect.signature(func)
    except Exception:
        signature = None
    if signature is None:
        return func(*args, **kwargs)
    allowed = {k: v for k, v in kwargs.items() if k in signature.parameters}
    return func(*args, **allowed)


def _send_command_best_effort(ham_int, command_name: str, **kwargs):
    last_error = None
    if hasattr(ham_int, "send_command"):
        for candidate in (command_name, command_name.upper(), command_name.lower()):
            try:
                return ham_int.send_command(command=candidate, **kwargs)
            except Exception as exc:
                last_error = exc
            try:
                return ham_int.send_command(candidate, **kwargs)
            except Exception as exc:
                last_error = exc
    raise RuntimeError(f"Unable to send PyHamilton command {command_name}: {last_error}")


def _wait_on_pyhamilton_command(ham_int, command_id, timeout: int = 300):
    if hasattr(ham_int, "wait_on_response"):
        return ham_int.wait_on_response(
            command_id,
            raise_first_exception=True,
            timeout=timeout,
        )
    return command_id


def _send_template_command(ham_int, template_name: str, **kwargs):
    """Send a native PyHamilton command template when the installed version exposes it."""
    template = None
    for module in (phi, pyhamilton):
        if module is not None and hasattr(module, template_name):
            template = getattr(module, template_name)
            break
    if template is None or not hasattr(ham_int, "send_command"):
        return None
    command_id = ham_int.send_command(template, **kwargs)
    return _wait_on_pyhamilton_command(ham_int, command_id)


def _clean_liquid_class(liq_class: str) -> str:
    return str(liq_class or "").strip()


def _sequence_positions_for_plan_item(
    item: Dict[str, Any],
    lay_metadata: Dict[str, Dict[str, Any]],
) -> str:
    """
    Resolve the exact LAY positions for one runtime plan item.

    Using explicit labware positions avoids relying on VENUS' mutable sequence
    cursor.  A saved/exhausted cursor can otherwise cause the very first command
    to fail with ``No more positions available in sequence ...`` even when the
    .lay still contains positions.
    """
    sequence = str(item.get("sequence") or "").strip()
    if not sequence or not lay_metadata:
        return ""

    meta = get_sequence_labware_metadata(
        sequence,
        lay_metadata,
        int(item.get("sequence_count") or 0),
    )
    items = list(meta.get("items") or [])
    if not items:
        return ""

    mode = str(item.get("hardware_mode") or "channels").lower()
    if mode == "head":
        needed = 96
    else:
        needed = max(count_active_channels(item.get("channel_pattern") or ""), 1)

    iteration = max(int(item.get("transfer_iteration") or 1), 1)
    start = 0 if bool(item.get("manual")) else (iteration - 1) * needed
    selected = items[start:start + needed]

    # Never silently wrap an exhausted sequence back to the beginning.  If the
    # LAY parser cannot provide a complete batch, fall back to the native
    # sequence reference so VENUS can report the real configuration problem.
    if len(selected) < needed:
        return ""

    positions = []
    for entry in selected:
        labware = str(entry.get("objid") or entry.get("labware_name") or "").strip()
        position = str(entry.get("position") or "").strip()
        if not labware or not position:
            return ""
        positions.append(f"{labware}, {position}")
    return ";".join(positions)


def ph_tip_pick_up_positions(ham_int, labware_positions: str, channel: str):
    channel = normalize_assigned_channel_pattern(channel, "channels")
    result = _send_template_command(
        ham_int,
        "PICKUP",
        tipSequence="",
        labwarePositions=labware_positions,
        channelVariable=channel,
        sequenceCounting=0,
        channelUse=1,
    )
    if result is not None:
        return result
    return _send_command_best_effort(
        ham_int,
        "channelTipPickUp",
        tipSequence="",
        labwarePositions=labware_positions,
        channelVariable=channel,
        sequenceCounting=0,
        channelUse=1,
    )


def ph_aspirate_positions(
    ham_int,
    labware_positions: str,
    vols: float,
    channel: str,
    liq_class: str,
):
    channel = normalize_assigned_channel_pattern(channel, "channels")
    liq_class = _clean_liquid_class(liq_class)
    result = _send_template_command(
        ham_int,
        "ASPIRATE",
        aspirateSequence="",
        labwarePositions=labware_positions,
        volumes=vols,
        channelVariable=channel,
        liquidClass=liq_class,
        sequenceCounting=0,
        channelUse=1,
    )
    if result is not None:
        return result
    return _send_command_best_effort(
        ham_int,
        "channelAspirate",
        aspirateSequence="",
        labwarePositions=labware_positions,
        volumes=vols,
        channelVariable=channel,
        liquidClass=liq_class,
        sequenceCounting=0,
        channelUse=1,
    )


def ph_dispense_positions(
    ham_int,
    labware_positions: str,
    vols: float,
    channel: str,
    liq_class: str,
):
    channel = normalize_assigned_channel_pattern(channel, "channels")
    liq_class = _clean_liquid_class(liq_class)
    result = _send_template_command(
        ham_int,
        "DISPENSE",
        dispenseSequence="",
        labwarePositions=labware_positions,
        volumes=vols,
        channelVariable=channel,
        liquidClass=liq_class,
        sequenceCounting=0,
        channelUse=1,
    )
    if result is not None:
        return result
    return _send_command_best_effort(
        ham_int,
        "channelDispense",
        dispenseSequence="",
        labwarePositions=labware_positions,
        volumes=vols,
        channelVariable=channel,
        liquidClass=liq_class,
        sequenceCounting=0,
        channelUse=1,
    )


def ph_tip_pick_up_seq(ham_int, tipseq: str, channel: str):
    channel = normalize_assigned_channel_pattern(channel, "channels")
    result = _send_template_command(
        ham_int,
        "PICKUP",
        tipSequence=tipseq,
        labwarePositions="",
        channelVariable=channel,
        sequenceCounting=1,
        channelUse=1,
    )
    if result is not None:
        return result
    func = _get_optional_pyhamilton_function(
        "tip_pick_up_seq", "tip_pickup_seq", "tip_pick_up", "tip_pickup"
    )
    if func:
        return _call_best_effort(
            func, ham_int, tipseq=tipseq, tip_seq=tipseq, sequence=tipseq,
            channel=channel, channels=channel
        )
    return _send_command_best_effort(ham_int, "channelTipPickUp", tipSequence=tipseq, labwarePositions="", channelVariable=channel, sequenceCounting=1, channelUse=1)


def ph_aspirate_seq(ham_int, asp_seq: str, vols: float, channel: str, liq_class: str):
    channel = normalize_assigned_channel_pattern(channel, "channels")
    liq_class = _clean_liquid_class(liq_class)
    result = _send_template_command(
        ham_int,
        "ASPIRATE",
        aspirateSequence=asp_seq,
        labwarePositions="",
        volumes=vols,
        channelVariable=channel,
        liquidClass=liq_class,
        sequenceCounting=1,
        channelUse=1,
    )
    if result is not None:
        return result
    func = _get_optional_pyhamilton_function("aspirate_seq", "aspirate_from_seq", "aspirate")
    if func:
        return _call_best_effort(func, ham_int, asp_seq=asp_seq, sequence=asp_seq, vols=vols, volume=vols, channel=channel, channels=channel, liq_class=liq_class, liquid_class=liq_class)
    return _send_command_best_effort(ham_int, "channelAspirate", aspirateSequence=asp_seq, labwarePositions="", volumes=vols, channelVariable=channel, liquidClass=liq_class, sequenceCounting=1, channelUse=1)


def ph_dispense_seq(ham_int, disp_seq: str, vols: float, channel: str, liq_class: str):
    channel = normalize_assigned_channel_pattern(channel, "channels")
    liq_class = _clean_liquid_class(liq_class)
    result = _send_template_command(
        ham_int,
        "DISPENSE",
        dispenseSequence=disp_seq,
        labwarePositions="",
        volumes=vols,
        channelVariable=channel,
        liquidClass=liq_class,
        sequenceCounting=1,
        channelUse=1,
    )
    if result is not None:
        return result
    func = _get_optional_pyhamilton_function("dispense_seq", "dispense_to_seq", "dispense")
    if func:
        return _call_best_effort(func, ham_int, disp_seq=disp_seq, sequence=disp_seq, vols=vols, volume=vols, channel=channel, channels=channel, liq_class=liq_class, liquid_class=liq_class)
    return _send_command_best_effort(ham_int, "channelDispense", dispenseSequence=disp_seq, labwarePositions="", volumes=vols, channelVariable=channel, liquidClass=liq_class, sequenceCounting=1, channelUse=1)


def ph_tip_eject_seq2(ham_int, waste_seq: str, channel: str):
    channel = normalize_assigned_channel_pattern(channel, "channels")
    use_default = 1 if not str(waste_seq or "").strip() or str(waste_seq).strip().lower() == "waste" else 0
    result = _send_template_command(
        ham_int,
        "EJECT",
        wasteSequence="" if use_default else waste_seq,
        labwarePositions="",
        channelVariable=channel,
        sequenceCounting=0 if use_default else 1,
        channelUse=1,
        useDefaultWaste=use_default,
        xDisplacement=0.0,
        yDisplacement=0.0,
        zDisplacement=0.0,
    )
    if result is not None:
        return result
    func = _get_optional_pyhamilton_function("tip_eject_seq2", "tip_eject_seq", "tip_eject")
    if func:
        return _call_best_effort(func, ham_int, waste_seq=waste_seq, sequence=waste_seq, channel=channel, channels=channel)
    return _send_command_best_effort(ham_int, "channelTipEject", wasteSequence=waste_seq, channelVariable=channel)


def ph_inc_sequence(ham_int, sequence: str, increment: int):
    func = _get_optional_pyhamilton_function("inc_sequence", "increment_sequence", "increment_seq")
    if func:
        return _call_best_effort(func, ham_int, sequence=sequence, increment=increment)
    return _send_command_best_effort(ham_int, "inc_sequence", sequence=sequence, increment=increment)


def ph_tip_pick_up_96_seq(ham_int, tip96_seq: str):
    func = _get_optional_pyhamilton_function("tip_pick_up_96_seq", "tip_pickup_96_seq", "tip_pick_up_96", "tip_pickup_96")
    if func:
        return _call_best_effort(func, ham_int, tip96_seq=tip96_seq, tip_seq=tip96_seq, sequence=tip96_seq)
    return _send_command_best_effort(ham_int, "tip_pick_up_96_seq", tip96_seq=tip96_seq, tip_seq=tip96_seq, sequence=tip96_seq)


def ph_aspirate_96_seq(ham_int, plate96, head_asp_seq: str, vols: float, liq_class: str):
    liq_class = _clean_liquid_class(liq_class)
    func = _get_optional_pyhamilton_function("aspirate_96_seq", "aspirate96_seq", "aspirate_96", "aspirate96")
    if func:
        return _call_best_effort(func, ham_int, plate96=plate96, plate=plate96, head_asp_seq=head_asp_seq, asp_seq=head_asp_seq, sequence=head_asp_seq, vols=vols, volume=vols, liq_class=liq_class, liquid_class=liq_class)
    return _send_command_best_effort(ham_int, "aspirate_96_seq", plate96=plate96, plate=plate96, head_asp_seq=head_asp_seq, asp_seq=head_asp_seq, sequence=head_asp_seq, vols=vols, volume=vols, liq_class=liq_class, liquid_class=liq_class)


def ph_dispense_96_seq2(ham_int, plate96, head_disp_seq: str, vols: float, liq_class: str):
    liq_class = _clean_liquid_class(liq_class)
    func = _get_optional_pyhamilton_function("dispense_96_seq2", "dispense_96_seq", "dispense96_seq", "dispense_96", "dispense96")
    if func:
        return _call_best_effort(func, ham_int, plate96=plate96, plate=plate96, head_disp_seq=head_disp_seq, disp_seq=head_disp_seq, sequence=head_disp_seq, vols=vols, volume=vols, liq_class=liq_class, liquid_class=liq_class)
    return _send_command_best_effort(ham_int, "dispense_96_seq2", plate96=plate96, plate=plate96, head_disp_seq=head_disp_seq, disp_seq=head_disp_seq, sequence=head_disp_seq, vols=vols, volume=vols, liq_class=liq_class, liquid_class=liq_class)


def ph_tip_eject_96(ham_int, sequence: str = ""):
    """
    Best-effort 96-head tip eject.

    Most PyHamilton versions do not require a sequence for 96-head eject, but
    if DevPal was given a unique tip-eject sequence we pass it through using
    common keyword names when the installed wrapper supports them.
    """
    sequence = str(sequence or "").strip()
    func = _get_optional_pyhamilton_function(
        "tip_eject_96",
        "tip_eject_96_seq",
        "tip_eject96",
    )
    if func:
        return _call_best_effort(
            func,
            ham_int,
            sequence=sequence,
            waste_seq=sequence,
            tip96_seq=sequence,
            tip_seq=sequence,
        )
    return _send_command_best_effort(
        ham_int,
        "tip_eject_96",
        sequence=sequence,
        waste_seq=sequence,
        tip96_seq=sequence,
        tip_seq=sequence,
    )


class AutoRunHamiltonInterface(phi.HamiltonInterface if phi else object):
    def start(self):
        if phi is None:
            raise RuntimeError(f"pyhamilton is not available. Import error: {PYHAMILTON_IMPORT_ERROR}")
        if self.active:
            return
        self.log("starting Hamilton interface")
        if self.simulate:
            self.server_thread = phi.HamiltonInterface.HamiltonServerThread(self.address, self.port)
            self.server_thread.start()
            time.sleep(1)
            subprocess.Popen([OEM_RUN_EXE_PATH, OEM_HSL_PATH, "-r", "-t"])
            self.log("started OEM application for simulation with auto-run")
        else:
            self.oem_process = Process(target=phi.run_hamilton_process, args=())
            self.oem_process.start()
            self.server_thread = phi.HamiltonInterface.HamiltonServerThread(self.address, self.port)
            self.server_thread.start()
        self.active = True


def infer_plate_format_from_dimensions(row_count: int, col_count: int) -> str:
    if row_count <= 0 or col_count <= 0:
        return "unknown"
    if row_count <= 4 and col_count <= 6:
        return "24-well"
    if row_count <= 6 and col_count <= 8:
        return "48-well"
    if row_count <= 8 and col_count <= 12:
        return "96-well"
    if row_count <= 16 and col_count <= 24:
        return "384-well"
    return f"unknown-{row_count}x{col_count}"

def infer_plate_format_from_position(position: str) -> str:
    pos = normalize_position_text(position)
    match = re.fullmatch(r"([A-P])(\d{1,2})", pos)
    if not match:
        return "unknown"
    row, col = match.group(1), int(match.group(2))
    if row <= "D" and col <= 6:
        return "24-well"
    if row <= "F" and col <= 8:
        return "48-well"
    if row <= "H" and col <= 12:
        return "96-well"
    if row <= "P" and col <= 24:
        return "384-well"
    return "unknown"

def infer_plate_format_from_labware_name(name: str) -> str:
    text = str(name or "").upper()
    if "384" in text:
        return "384-well"
    if "96" in text or "PLATE96" in text or "PCR96" in text:
        return "96-well"
    if "48" in text and "RACK" not in text and "TUBE" not in text:
        return "48-well"
    if "24" in text and "RACK" not in text and "TUBE" not in text:
        return "24-well"
    if "WASTE" in text:
        return "waste"
    if "TIP" in text or "SLIM" in text:
        return "tip rack"
    if any(word in text for word in ("RACK", "TUBE", "SAMPLE", "CARRIER", "HONEYCOMB")):
        return "rack/unknown capacity"
    if "PLATE" in text:
        return "plate/unknown format"
    return "unknown"

def infer_labware_format_from_name(name: str) -> str:
    return infer_plate_format_from_labware_name(name)


def parse_lay_file(lay_path: str) -> List[str]:
    text = read_text_file_best_effort(lay_path)
    if not text:
        return []
    found, seen = [], set()
    for seq_id in sorted({int(v) for v in re.findall(r"Seq\.(\d+)\.Name", text, re.I)}):
        pattern = re.compile(rf"Seq\.{seq_id}\.Name(.*?)(?=Seq\.{seq_id}\.(?:ReadOnly|Cnt|Item)|Seq\.\d+\.Name|Seq\.Cnt|$)", re.I | re.S)
        match = pattern.search(text)
        if not match:
            continue
        raw = re.sub(r"[\x00-\x1f]+", " ", match.group(1))
        raw = re.sub(r"^[^A-Za-z0-9_]+", "", raw)
        name_match = re.search(r"([A-Za-z][A-Za-z0-9_ .\-]{0,119})", raw)
        if not name_match:
            continue
        name = re.sub(r"\s+", " ", name_match.group(1)).strip()
        if name and name not in seen and not name.lower().startswith(("readonly", "cnt", "item")):
            seen.add(name)
            found.append(name)
    return found


def parse_lay_sequence_counts(lay_path: str) -> Dict[str, int]:
    text = read_text_file_best_effort(lay_path)
    if not text:
        return {}
    counts = {int(n): int(c) for n, c in re.findall(r"Seq\.(\d+)\.Cnt[^\d]*(\d+)", text, re.I)}
    names = {}
    for seq_id in sorted({int(v) for v in re.findall(r"Seq\.(\d+)\.Name", text, re.I)}):
        pattern = re.compile(rf"Seq\.{seq_id}\.Name(.*?)(?=Seq\.{seq_id}\.(?:ReadOnly|Cnt|Item)|Seq\.\d+\.Name|Seq\.Cnt|$)", re.I | re.S)
        match = pattern.search(text)
        if match:
            raw = re.sub(r"[\x00-\x1f]+", " ", match.group(1))
            raw = re.sub(r"^[^A-Za-z0-9_]+", "", raw)
            m = re.search(r"([A-Za-z][A-Za-z0-9_ .\-]{0,119})", raw)
            if m:
                names[seq_id] = re.sub(r"\s+", " ", m.group(1)).strip()
    return {name: counts.get(seq_id, 0) for seq_id, name in names.items()}


def parse_lay_sequence_metadata(lay_path: str) -> Dict[str, Dict[str, Any]]:
    """
    Preserve all sequence item labware references, not only one "most frequent"
    labware. This lets DevPal identify a dispense sequence that continues across
    plate 1, plate 2, plate 3, etc.
    """
    text = read_text_file_best_effort(lay_path)
    metadata: Dict[str, Dict[str, Any]] = {}
    if not text:
        return metadata

    def extract_value(pattern: str, default: str = "") -> str:
        match = re.search(pattern, text, re.I | re.S)
        if not match:
            return default
        value = match.group(1).replace("\x00", " ")
        value = re.sub(r"[\x01-\x1f]+", " ", value)
        return re.sub(r"\s+", " ", value).strip()

    labware_by_id: Dict[str, Dict[str, Any]] = {}

    for labware_num in sorted(set(re.findall(r"Labware\.(\d+)\.", text, re.I)), key=int):
        stop = rf"(?=Labware\.{labware_num}\.|Labware\.\d+\.|Seq\.|$)"
        labware_id = extract_value(rf"Labware\.{labware_num}\.Id(.*?){stop}")
        labware_file = extract_value(rf"Labware\.{labware_num}\.File(.*?){stop}")
        template = extract_value(rf"Labware\.{labware_num}\.Template(.*?){stop}")
        site_id = extract_value(rf"Labware\.{labware_num}\.SiteId(.*?){stop}")

        if labware_id:
            labware_by_id[labware_id.lower()] = {
                "labware_name": labware_id,
                "labware_file": labware_file,
                "template": template,
                "site_id": site_id,
                "labware_format": infer_plate_format_from_labware_name(
                    f"{labware_id} {labware_file} {template}"
                ),
            }

    for seq_id in sorted(set(re.findall(r"Seq\.(\d+)\.", text, re.I)), key=int):
        raw_name = extract_value(
            rf"Seq\.{seq_id}\.Name(.*?)(?=Seq\.{seq_id}\.(?:ReadOnly|Cnt|Item)|Seq\.\d+\.Name|Seq\.Cnt|$)"
        )
        m = re.search(r"([A-Za-z][A-Za-z0-9_ .\-]{0,119})", raw_name)
        if not m:
            continue

        seq_name = re.sub(r"\s+", " ", m.group(1)).strip()
        cnt_match = re.search(rf"Seq\.{seq_id}\.Cnt[^\d]*(\d+)", text, re.I)
        selected_position_count = int(cnt_match.group(1)) if cnt_match else 0

        item_numbers = sorted({
            int(v)
            for v in re.findall(rf"Seq\.{seq_id}\.Item\.(\d+)\.", text, re.I)
        })

        items = []
        labware_names = []
        positions = []

        for item_number in item_numbers:
            item_prefix = rf"Seq\.{seq_id}\.Item\.{item_number}\."
            item_stop = rf"(?=Seq\.{seq_id}\.Item\.\d+\.|Seq\.{seq_id}\.|Seq\.\d+\.|$)"
            raw_objid = extract_value(rf"{item_prefix}ObjId(.*?){item_stop}")
            raw_posid = extract_value(rf"{item_prefix}PosId(.*?){item_stop}")

            obj_m = re.search(r"([A-Za-z][A-Za-z0-9_ .\-]{0,150})", raw_objid)
            objid = re.sub(r"\s+", " ", obj_m.group(1)).strip() if obj_m else ""

            pos_m = re.search(r"([A-Za-z]?\d{1,3})", raw_posid)
            position = normalize_position_text(pos_m.group(1)) if pos_m else ""

            lw_meta = labware_by_id.get(objid.lower(), {})
            labware_name = lw_meta.get("labware_name", objid)

            if labware_name and labware_name not in labware_names:
                labware_names.append(labware_name)
            if position:
                positions.append(position)

            items.append({
                "item_number": item_number,
                "objid": objid,
                "labware_name": labware_name,
                "position": position,
                "labware_format": lw_meta.get(
                    "labware_format",
                    infer_plate_format_from_labware_name(labware_name),
                ),
            })

        formats = []
        for lw in labware_names:
            fmt = labware_by_id.get(lw.lower(), {}).get(
                "labware_format",
                infer_plate_format_from_labware_name(lw),
            )
            if fmt not in formats:
                formats.append(fmt)

        primary_labware = labware_names[0] if labware_names else seq_name

        metadata[seq_name] = {
            "sequence_id": seq_id,
            "sequence_name": seq_name,
            "selected_position_count": selected_position_count,
            "position_count": selected_position_count,
            "positions": positions,
            "items": items,
            "labware_name": primary_labware,
            "labware_names": labware_names,
            "labware_count": len(labware_names),
            "labware_format": (
                formats[0] if len(formats) == 1
                else "mixed" if len(formats) > 1
                else infer_plate_format_from_labware_name(primary_labware)
            ),
            "labware_formats": formats,
            "labware_format_basis": "all_sequence_items_preserved",
        }

    return metadata


def get_sequence_labware_metadata(
    sequence_name: str,
    lay_metadata: Dict[str, Dict[str, Any]],
    sequence_count: int = 0,
) -> Dict[str, Any]:
    target = str(sequence_name or "").strip()
    if target in lay_metadata:
        return lay_metadata[target]

    for name, meta in lay_metadata.items():
        if str(name).strip().lower() == target.lower():
            return meta

    inferred = infer_plate_format_from_labware_name(target)
    return {
        "sequence_name": target,
        "labware_name": target if target else "unknown",
        "labware_names": [target] if target else [],
        "labware_count": 1 if target else 0,
        "labware_format": inferred,
        "labware_formats": [inferred] if inferred != "unknown" else [],
        "selected_position_count": int(sequence_count or 0),
        "position_count": int(sequence_count or 0),
        "positions": [],
        "items": [],
    }


def summarize_layout_format(layout_matches: Dict[str, List[Dict[str, Any]]]) -> Dict[str, Any]:
    max_row_num = 0
    max_col = 0
    positions = []

    for entries in layout_matches.values():
        for entry in entries:
            row = str(entry.get("plate_row") or "").strip().upper()
            col = int(entry.get("plate_column") or 0)
            if row and "A" <= row <= "P":
                max_row_num = max(max_row_num, ord(row) - ord("A") + 1)
            max_col = max(max_col, col)
            if entry.get("dest_value"):
                positions.append(normalize_position_text(entry["dest_value"]))

    return {
        "row_count": max_row_num,
        "column_count": max_col,
        "inferred_plate_format": infer_plate_format_from_dimensions(max_row_num, max_col),
        "positions": positions,
    }


def extract_liquid_classes_from_method_files(method_path: str) -> List[str]:
    if not method_path:
        return []

    method_path = os.path.abspath(method_path)
    method_dir = os.path.dirname(method_path)
    candidates = [method_path] if os.path.isfile(method_path) else []

    for ext in ("*.hsl", "*.sub", "*.stp", "*.res", "*.med", "*.lay"):
        candidates.extend(glob.glob(os.path.join(method_dir, ext)))

    patterns = [
        re.compile(r"\b(?:liq_class|liquid_class|liquidClass|LiquidClass)\b\s*[=:,]\s*[\"']([^\"']+)[\"']", re.I),
        re.compile(r"\b(?:LiquidClassName|Liquid Class Name)\b\W+[\"']?([A-Za-z0-9_ .\-+/()]+)", re.I),
        re.compile(r"[\"']([A-Za-z0-9_]+(?:Water|Serum|Plasma|Buffer|DMSO|EtOH|Media|Dispense|Aspirate|Empty|Jet|Surface)[A-Za-z0-9_ .\-+/()]*)[\"']", re.I),
    ]

    found = set()
    for candidate in sorted(set(candidates), key=str.lower):
        text = read_text_file_best_effort(candidate)
        for pattern in patterns:
            for raw in pattern.findall(text):
                value = clean_control_chars(raw).strip("\"' ,;:=()[]{}")
                value = re.split(r"[\r\n\t;]", value)[0].strip()
                if 3 <= len(value) <= 150 and re.fullmatch(r"[A-Za-z0-9_][A-Za-z0-9_ .\-+/()]*", value):
                    found.add(value)

    return sorted(found, key=str.lower)


def load_combined_liquid_classes(method_path: str, db_path: str = "") -> List[str]:
    return extract_liquid_classes_from_method_files(method_path)


def normalize_tip_type(value: str) -> str:
    text = str(value or "").lower().replace(" ", "").replace("_", "")
    if text in {"1000ul", "1000"}:
        return "1000ul"
    if text in {"300ul", "300s", "300"}:
        return "300ul"
    if text in {"50ul", "50"}:
        return "50ul"
    return text


def infer_tip_type_from_text(text: str) -> str:
    compact = re.sub(r"[^A-Z0-9]+", "", str(text or "").upper())
    if any(t in compact for t in ("HIGHVOLUME", "1000UL", "HVT", "HTL", "HTF", "TIP1000")):
        return "1000ul"
    if any(t in compact for t in ("STANDARDVOLUME", "SLIMTIP", "300UL", "SVT", "STL", "STF", "TIP300")):
        return "300ul"
    if any(t in compact for t in ("LOWVOLUME", "50UL", "LVT", "LTL", "LTF", "TIP50")):
        return "50ul"
    return ""


def infer_tip_type_for_set(
    set_order: int,
    configs: List[SequenceStepConfig],
    lay_metadata: Dict[str, Dict[str, Any]],
) -> Tuple[str, str]:
    set_configs = [c for c in configs if int(c.set_order) == int(set_order)]
    evidence = []

    for config in [c for c in set_configs if c.role == "tip pick up"]:
        meta = get_sequence_labware_metadata(config.sequence, lay_metadata, config.sequence_count)
        text = " | ".join([
            config.sequence,
            str(meta.get("labware_name", "")),
            " ".join(meta.get("labware_names", []) or []),
            str(meta.get("labware_format", "")),
        ])
        evidence.append(text)
        tip = infer_tip_type_from_text(text)
        if tip:
            return tip, text

    for config in [c for c in set_configs if c.role in ("aspirate", "dispense")]:
        text = " | ".join([config.sequence, config.liquid_class, config.hardware_mode])
        evidence.append(text)
        tip = infer_tip_type_from_text(text)
        if tip:
            return tip, text

    return "", " | ".join(evidence) if evidence else "No tip evidence found."


def tip_volume_allowed(tip_type: str, hardware_mode: str, volume_ul: float) -> bool:
    mode = str(hardware_mode or "channels").lower()
    ranges = (
        {"1000ul": (100, 1000), "300ul": (20, 300), "50ul": (12.5, 49)}
        if mode == "head"
        else {"1000ul": (100, 1000), "300ul": (10, 300), "50ul": (10, 50)}
    )
    tip = normalize_tip_type(tip_type)
    return tip in ranges and ranges[tip][0] <= float(volume_ul) <= ranges[tip][1]


def validate_tip_volume_rules(
    configs: List[SequenceStepConfig],
    lay_metadata: Dict[str, Dict[str, Any]],
) -> List[ValidationIssue]:
    issues = []
    for set_order in sorted({int(c.set_order) for c in configs}):
        liquid_steps = [
            c for c in configs
            if int(c.set_order) == set_order and c.role in ("aspirate", "dispense")
        ]
        if not liquid_steps:
            continue

        tip_type, evidence = infer_tip_type_for_set(set_order, configs, lay_metadata)
        if not tip_type:
            issues.append(ValidationIssue(
                "minor",
                "tip volume rule",
                f"Tip type could not be confirmed for Set {set_order}; volume validation was skipped rather than blocking the run.",
                {"set_order": set_order, "evidence": evidence},
            ))
            continue

        for config in liquid_steps:
            if not tip_volume_allowed(tip_type, config.hardware_mode, config.volume_ul):
                issues.append(ValidationIssue(
                    "critical",
                    "tip volume rule",
                    f"Volume {config.volume_ul} uL is not valid for {tip_type} tips in {config.hardware_mode} mode for Set {set_order}.",
                    {"sequence": config.sequence, "tip_type": tip_type, "volume_ul": config.volume_ul},
                ))
    return issues


def detect_hardware_mode(sequence_name: str, selected_role: str = "") -> str:
    text = f"{sequence_name} {selected_role}".upper()
    return "head" if any(t in text for t in ("CO-RE", "CORE", "96", "HEAD")) else "channels"


def validate_step_sets(configs: List[SequenceStepConfig], variant_mode: bool, use_set_order: bool) -> str:
    if not use_set_order:
        return ""

    grouped = defaultdict(list)
    for config in configs:
        grouped[int(config.set_order)].append(config)

    for set_order, items in sorted(grouped.items()):
        roles = [item.role for item in items]
        for required in ROLE_OPTIONS:
            if variant_mode:
                if required not in roles:
                    return f"Set {set_order} must contain at least one {required} sequence."
            elif roles.count(required) != 1:
                return f"Set {set_order} must contain exactly one tip pick up, one aspirate, and one dispense sequence."
    return ""


def validate_step_configs(configs: List[SequenceStepConfig]) -> List[ValidationIssue]:
    issues = []
    roles = {c.role for c in configs}
    for required in ROLE_OPTIONS:
        if required not in roles:
            issues.append(ValidationIssue("critical", "configuration", f"Missing required step role: {required}"))

    for config in configs:
        if config.role not in ALL_ROLE_OPTIONS:
            issues.append(
                ValidationIssue(
                    "critical",
                    "configuration",
                    f"Invalid step role for {config.sequence}: {config.role}",
                )
            )

        if config.role in ("aspirate", "dispense") and not config.liquid_class:
            issues.append(ValidationIssue("critical", "configuration", f"Liquid class is required for {config.role}: {config.sequence}"))
        if config.volume_ul <= 0 or config.volume_ul > 1000:
            issues.append(ValidationIssue("critical", "configuration", f"Volume must be between 1 and 1000 uL: {config.sequence}"))
        if config.transfer_type not in TRANSFER_OPTIONS:
            issues.append(ValidationIssue("critical", "configuration", f"Invalid transfer type: {config.transfer_type}"))
        if config.replicate not in REPLICATE_OPTIONS:
            issues.append(ValidationIssue("critical", "configuration", f"Invalid replicate value: {config.replicate}"))
    return issues


def get_sequence_batch_size(config: SequenceStepConfig) -> int:
    if config.hardware_mode == "head":
        return 96
    return max(count_active_channels(config.channel_pattern), 1)


def get_sequence_iterations(config: SequenceStepConfig) -> int:
    if config.manual:
        return 1
    if config.sequence_count <= 0:
        return 1
    batch = get_sequence_batch_size(config)
    return max((config.sequence_count + batch - 1) // batch, 1)


def get_required_transfer_iterations(configs: List[SequenceStepConfig]) -> int:
    controls = [
        c for c in configs
        if c.control
        and not c.manual
        and c.role in ("tip pick up", "aspirate", "dispense", TIP_EJECT_ROLE)
    ]
    if controls:
        return max(get_sequence_iterations(c) for c in controls)
    dispenses = [c for c in configs if c.role == "dispense" and not c.manual]
    if dispenses:
        return max(get_sequence_iterations(c) for c in dispenses)
    aspirates = [c for c in configs if c.role == "aspirate" and not c.manual]
    if aspirates:
        return max(get_sequence_iterations(c) for c in aspirates)
    return 1


def build_execution_plan(
    configs: List[SequenceStepConfig],
    lay_metadata: Optional[Dict[str, Dict[str, Any]]] = None,
    variant_mode: bool = False,
) -> List[Dict[str, Any]]:
    """
    Build the ordered runtime plan.

    Normal behavior is unchanged when no sequence is assigned the "tip eject"
    role: DevPal automatically adds its normal Waste tip-eject step after each
    transfer.

    When one or more selected sequences are assigned "tip eject", those
    sequences replace the automatic default eject for their set. They retain
    their own channel pattern, hardware mode, sequence count, Manual flag,
    Control flag, and set assignment just like the other selected step roles.
    """
    lay_metadata = lay_metadata or {}
    grouped = defaultdict(list)

    for config in configs:
        grouped[int(config.set_order or 1)].append(config)

    plan: List[Dict[str, Any]] = []
    order = 1

    def action_for_role(role: str) -> str:
        role = str(role or "").strip().lower()
        if role == "tip pick up":
            return "tip_pick"
        if role == TIP_EJECT_ROLE:
            return "tip_eject"
        return role

    def append_config(
        config: SequenceStepConfig,
        set_order: int,
        replicate_index: int,
        transfer_iteration: int,
    ):
        nonlocal order

        batch = get_sequence_batch_size(config)
        iterations = get_sequence_iterations(config)

        autoincrement = (
            transfer_iteration > 1
            and transfer_iteration <= iterations
            and not config.manual
            and config.role in ("aspirate", "dispense", TIP_EJECT_ROLE)
        )

        meta = get_sequence_labware_metadata(
            config.sequence,
            lay_metadata,
            config.sequence_count,
        )

        plan.append({
            "order": order,
            "set_order": set_order,
            "action": action_for_role(config.role),
            "sequence": config.sequence,
            "hardware_mode": config.hardware_mode,
            "labware_name": meta.get("labware_name", ""),
            "labware_names": meta.get("labware_names", []),
            "labware_count": meta.get("labware_count", 0),
            "labware_format": meta.get("labware_format", ""),
            "replicate_index": replicate_index,
            "transfer_iteration": transfer_iteration,
            "manual": bool(config.manual),
            "control": bool(config.control),
            "autoincrement": autoincrement,
            "increment": batch if autoincrement else 0,
            "channel_pattern": config.channel_pattern,
            "volume_ul": (
                config.volume_ul
                if config.role in ("aspirate", "dispense")
                else ""
            ),
            "liquid_class": (
                config.liquid_class
                if config.role in ("aspirate", "dispense")
                else ""
            ),
            "transfer_type": config.transfer_type,
            "marker": config.marker,
            "sequence_count": config.sequence_count,
            "batch_size": batch,
            "unique_tip_eject": config.role == TIP_EJECT_ROLE,
        })
        order += 1

    def append_default_eject(
        tip: SequenceStepConfig,
        set_order: int,
        replicate_index: int,
        transfer_iteration: int,
    ):
        """Preserve the legacy/default eject exactly when no unique eject is selected."""
        nonlocal order

        plan.append({
            "order": order,
            "set_order": set_order,
            "action": "tip_eject",
            "sequence": "Waste",
            "hardware_mode": tip.hardware_mode,
            "labware_name": "Waste",
            "labware_names": ["Waste"],
            "labware_count": 1,
            "labware_format": "waste",
            "replicate_index": replicate_index,
            "transfer_iteration": transfer_iteration,
            "manual": False,
            "control": False,
            "autoincrement": False,
            "increment": 0,
            "channel_pattern": tip.channel_pattern,
            "volume_ul": "",
            "liquid_class": "",
            "transfer_type": tip.transfer_type,
            "marker": tip.marker,
            "sequence_count": tip.sequence_count,
            "batch_size": get_sequence_batch_size(tip),
            "unique_tip_eject": False,
        })
        order += 1

    for set_order in sorted(grouped):
        items = sorted(
            grouped[set_order],
            key=lambda config: config.selected_index,
        )

        tips = [c for c in items if c.role == "tip pick up"]
        aspirates = [c for c in items if c.role == "aspirate"]
        dispenses = [c for c in items if c.role == "dispense"]
        tip_ejects = [c for c in items if c.role == TIP_EJECT_ROLE]

        iterations = get_required_transfer_iterations(items)

        replicate_index = REPLICATE_MAP.get(
            items[0].replicate if items else "single",
            1,
        )

        for transfer_iteration in range(1, iterations + 1):
            if variant_mode:
                for index, dispense_config in enumerate(dispenses):
                    if not tips or not aspirates:
                        continue

                    tip_config = tips[index % len(tips)]
                    aspirate_config = aspirates[index % len(aspirates)]

                    append_config(
                        tip_config,
                        set_order,
                        replicate_index,
                        transfer_iteration,
                    )
                    append_config(
                        aspirate_config,
                        set_order,
                        replicate_index,
                        transfer_iteration,
                    )
                    append_config(
                        dispense_config,
                        set_order,
                        replicate_index,
                        transfer_iteration,
                    )

                    if tip_ejects:
                        eject_config = tip_ejects[index % len(tip_ejects)]
                        append_config(
                            eject_config,
                            set_order,
                            replicate_index,
                            transfer_iteration,
                        )
                    else:
                        append_default_eject(
                            tip_config,
                            set_order,
                            replicate_index,
                            transfer_iteration,
                        )

                continue

            max_len = max(
                len(tips),
                len(aspirates),
                len(dispenses),
                len(tip_ejects) if tip_ejects else 0,
                1,
            )

            for index in range(max_len):
                tip_config = (
                    tips[index]
                    if index < len(tips)
                    else tips[0] if tips else None
                )
                aspirate_config = (
                    aspirates[index]
                    if index < len(aspirates)
                    else aspirates[0] if aspirates else None
                )
                dispense_config = (
                    dispenses[index]
                    if index < len(dispenses)
                    else dispenses[0] if dispenses else None
                )

                for config in (
                    tip_config,
                    aspirate_config,
                    dispense_config,
                ):
                    if config is not None:
                        append_config(
                            config,
                            set_order,
                            replicate_index,
                            transfer_iteration,
                        )

                if tip_ejects:
                    eject_config = (
                        tip_ejects[index]
                        if index < len(tip_ejects)
                        else tip_ejects[0]
                    )
                    append_config(
                        eject_config,
                        set_order,
                        replicate_index,
                        transfer_iteration,
                    )
                elif tip_config is not None:
                    append_default_eject(
                        tip_config,
                        set_order,
                        replicate_index,
                        transfer_iteration,
                    )

    return plan

def _parse_marker_token(marker_token: Any) -> Tuple[str, Optional[int]]:
    text = str(marker_token or "").strip().upper()
    match = re.fullmatch(r"([A-Z][A-Z0-9 _\-]*?)(\d+)", text)
    if not match:
        return "", None
    prefix = re.sub(r"[\s_\-]+", "", match.group(1))
    return prefix, int(match.group(2))


def _add_layout_cell(
    plate_layout: Dict[str, List[Dict[str, Any]]],
    cell_text: str,
    dest_position: str,
    plate_row: str,
    plate_col: int,
    row_index: int,
    column_index: int,
    source_name: str,
    marker_filter: str = "",
):
    marker_filter = re.sub(r"[\s_\-]+", "", str(marker_filter or "").upper())
    token_pattern = re.compile(r"\b([A-Za-z][A-Za-z0-9 _\-]*?\d+)\b")

    for token in token_pattern.findall(cell_text):
        prefix, number = _parse_marker_token(token)
        if not prefix or number is None:
            continue
        if marker_filter and prefix != marker_filter:
            continue

        label = f"{prefix}{number}"
        plate_layout[label].append({
            "dest_value": dest_position,
            "cell_value": cell_text,
            "plate_row": plate_row,
            "plate_column": plate_col,
            "row_index": row_index,
            "column_index": column_index,
            "excel_position": f"{excel_col_name(column_index - 1)}{row_index}",
            "column_name": excel_col_name(column_index - 1),
            "layout_source": source_name,
        })


def parse_dataframe_plate_layout(df, plate_layout, source_name: str, marker_filter: str = ""):
    if df is None or df.empty:
        return

    for row_idx in range(1, df.shape[0]):
        row_raw = df.iat[row_idx, 0]
        if pd.isna(row_raw):
            continue
        plate_row = str(row_raw).strip().upper()
        if not re.fullmatch(r"[A-P]", plate_row):
            continue

        for col_idx in range(1, df.shape[1]):
            col_raw = df.iat[0, col_idx]
            if pd.isna(col_raw):
                continue
            col_match = re.search(r"\d+", str(col_raw))
            if not col_match:
                continue
            plate_col = int(col_match.group(0))

            cell = df.iat[row_idx, col_idx]
            if pd.isna(cell):
                continue
            cell_text = str(cell).strip().upper()
            if not cell_text:
                continue

            _add_layout_cell(
                plate_layout,
                cell_text,
                f"{plate_row}{plate_col}",
                plate_row,
                plate_col,
                row_idx + 1,
                col_idx + 1,
                source_name,
                marker_filter,
            )


def parse_plate_layout_file(layout_path: str, marker: str = "") -> Dict[str, List[Dict[str, Any]]]:
    """
    Marker is optional. If blank, numbered markers are discovered automatically.
    XLS/XLSX worksheet names and DOCX table identities are retained so multiple
    templates are not accidentally mistaken for replicate count.
    """
    if pd is None:
        raise RuntimeError("pandas is required for layout parsing.")

    plate_layout = defaultdict(list)
    ext = os.path.splitext(layout_path)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(layout_path, header=None, dtype=str)
        parse_dataframe_plate_layout(df, plate_layout, "CSV", marker)

    elif ext in (".xls", ".xlsx"):
        sheets = pd.read_excel(layout_path, sheet_name=None, header=None, dtype=str)
        for sheet_name, df in sheets.items():
            parse_dataframe_plate_layout(df, plate_layout, str(sheet_name), marker)

    elif ext == ".docx":
        if Document is None:
            raise RuntimeError("python-docx is required for DOCX layout parsing.")

        doc = Document(layout_path)
        for table_index, table in enumerate(doc.tables, start=1):
            if not table.rows:
                continue

            headers = table.rows[0].cells
            for row_idx in range(1, len(table.rows)):
                row = table.rows[row_idx]
                if not row.cells:
                    continue
                plate_row = str(row.cells[0].text).strip().upper()
                if not re.fullmatch(r"[A-P]", plate_row):
                    continue

                for col_idx in range(1, len(row.cells)):
                    if col_idx >= len(headers):
                        continue
                    col_match = re.search(r"\d+", str(headers[col_idx].text))
                    if not col_match:
                        continue
                    plate_col = int(col_match.group(0))
                    cell_text = str(row.cells[col_idx].text).strip().upper()
                    if not cell_text:
                        continue

                    _add_layout_cell(
                        plate_layout,
                        cell_text,
                        f"{plate_row}{plate_col}",
                        plate_row,
                        plate_col,
                        row_idx + 1,
                        col_idx + 1,
                        f"Table {table_index}",
                        marker,
                    )
    else:
        raise ValueError("Unsupported plate layout file type. Use CSV, XLS, XLSX, or DOCX.")

    return dict(plate_layout)


def get_layout_sources(
    layout_matches: Dict[str, List[Dict[str, Any]]],
) -> List[str]:
    sources: List[str] = []
    for entries in (layout_matches or {}).values():
        for entry in entries or []:
            source = str(entry.get("layout_source") or "default")
            if source not in sources:
                sources.append(source)
    return sources or ["default"]


def build_layout_templates(
    layout_matches: Dict[str, List[Dict[str, Any]]],
    configured_marker: str = "",
    configured_replicate: str = "",
) -> Tuple[str, List[Dict[str, Any]]]:
    if not layout_matches:
        raise ValueError("No plate-layout markers were found.")

    prefix = infer_marker_prefix(layout_matches, configured_marker)
    sources = get_layout_sources(layout_matches)
    templates: List[Dict[str, Any]] = []

    configured_count = REPLICATE_MAP.get(str(configured_replicate or "").lower())

    for source in sources:
        marker_numbers: List[int] = []
        wells_by_number: Dict[int, List[str]] = {}
        occurrences_by_number: Dict[int, int] = {}

        for marker_label, entries in (layout_matches or {}).items():
            marker_prefix, marker_number = _parse_marker_token(marker_label)
            if marker_prefix != prefix or marker_number is None:
                continue

            source_entries = [
                entry
                for entry in entries or []
                if str(entry.get("layout_source") or "default") == source
            ]
            if not source_entries:
                continue

            wells = [
                normalize_position_text(entry.get("dest_value", ""))
                for entry in source_entries
                if entry.get("dest_value") not in (None, "", "nan")
            ]

            if not wells:
                continue

            marker_numbers.append(marker_number)
            wells_by_number[marker_number] = wells
            occurrences_by_number[marker_number] = len(wells)

        marker_numbers = sorted(set(marker_numbers))
        if not marker_numbers:
            continue

        replicate_counts = sorted(set(occurrences_by_number.values()))
        if len(replicate_counts) != 1:
            raise ValueError(
                f"Plate layout source '{source}' has inconsistent replicate counts: "
                f"{occurrences_by_number}."
            )

        replicate_count = replicate_counts[0]
        if replicate_count not in (1, 2, 3):
            raise ValueError(
                f"Plate layout source '{source}' indicates {replicate_count} "
                "destination(s) per sample; only single, duplicate, and "
                "triplicate are supported."
            )

        if configured_count is not None and configured_count != replicate_count:
            raise ValueError(
                f"Replicate mismatch: UI is '{configured_replicate}' "
                f"({configured_count}) but plate layout source '{source}' "
                f"indicates {replicate_count}."
            )

        templates.append({
            "layout_source": source,
            "marker_prefix": prefix,
            "marker_numbers": marker_numbers,
            "wells_by_marker_number": wells_by_number,
            "occurrences_by_marker_number": occurrences_by_number,
            "samples_per_plate": len(marker_numbers),
            "replicate_count": replicate_count,
            "replicate": REPLICATE_NAME_BY_COUNT.get(
                replicate_count,
                f"{replicate_count}x",
            ),
        })

    if not templates:
        raise ValueError(f"No numbered markers were found for '{prefix}'.")

    replicate_counts = {template["replicate_count"] for template in templates}
    if len(replicate_counts) != 1:
        raise ValueError(
            "All plate-layout sheets/tables must use the same replicate count."
        )

    return prefix, templates


def detect_plate_layout_marker(layout_path: str) -> str:
    matches = parse_plate_layout_file(layout_path, "")
    return infer_marker_prefix(matches, "")

def infer_marker_prefix(
    layout_matches: Dict[str, List[Dict[str, Any]]],
    configured_marker: str = "",
) -> str:
    configured = re.sub(r"[\s_\-]+", "", str(configured_marker or "").strip().upper())
    prefix_numbers = defaultdict(set)

    for marker_label in layout_matches:
        prefix, number = _parse_marker_token(marker_label)
        if prefix and number is not None:
            prefix_numbers[prefix].add(number)

    if configured:
        if configured not in prefix_numbers:
            available = ", ".join(sorted(prefix_numbers)) or "none"
            raise ValueError(
                f"Marker '{configured_marker}' was not found in the plate layout. "
                f"Detected marker prefix(es): {available}."
            )
        return configured

    if not prefix_numbers:
        raise ValueError("No numbered sample markers were found, such as S1, S2, S3.")

    ranked = sorted(
        prefix_numbers,
        key=lambda p: (len(prefix_numbers[p]), max(prefix_numbers[p])),
        reverse=True,
    )

    if len(ranked) > 1:
        first_score = (len(prefix_numbers[ranked[0]]), max(prefix_numbers[ranked[0]]))
        second_score = (len(prefix_numbers[ranked[1]]), max(prefix_numbers[ranked[1]]))
        if first_score == second_score:
            raise ValueError(
                "Multiple equally likely marker prefixes were found. "
                f"Enter the intended marker in the Marker field: {', '.join(ranked)}."
            )

    return ranked[0]


def get_marker_numbers(
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker: str = "",
) -> List[int]:
    prefix = infer_marker_prefix(layout_matches, marker)
    numbers = []
    for key in layout_matches:
        p, n = _parse_marker_token(key)
        if p == prefix and n is not None:
            numbers.append(n)
    return sorted(set(numbers))


def get_layout_marker_positions(
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker_label: str,
    layout_source: str = "",
) -> List[str]:
    entries = layout_matches.get(str(marker_label).upper().strip(), [])
    if layout_source:
        entries = [
            entry for entry in entries
            if str(entry.get("layout_source") or "default") == str(layout_source)
        ]
    return [
        normalize_position_text(entry.get("dest_value", ""))
        for entry in entries
        if entry.get("dest_value") not in (None, "", "nan")
    ]

def analyze_plate_layout_pattern(
    layout_matches: Dict[str, List[Dict[str, Any]]],
    configured_marker: str = "",
    configured_replicate: str = "",
) -> PlateLayoutPattern:
    """
    Analyze the first physical plate template for summary/report compatibility.

    Full multi-sheet/sample-to-plate mapping is handled by
    expected_plate_layout_for_sample(). A single sheet is a reusable template;
    multiple sheets/tables are treated as sequential plate-specific templates.
    """
    prefix, templates = build_layout_templates(
        layout_matches,
        configured_marker,
        configured_replicate,
    )
    template = templates[0]

    return PlateLayoutPattern(
        marker_prefix=prefix,
        marker_numbers=list(template["marker_numbers"]),
        wells_by_marker_number=dict(template["wells_by_marker_number"]),
        occurrences_by_marker_number=dict(template["occurrences_by_marker_number"]),
        samples_per_plate=int(template["samples_per_plate"]),
        replicate_count=int(template["replicate_count"]),
    )

def resolve_repeating_layout_marker(
    source_number: int,
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker: str,
) -> Tuple[str, int]:
    expected = expected_plate_layout_for_sample(
        source_number,
        layout_matches,
        marker,
        "",
    )
    return expected["template_marker"], expected["marker_number"]

def expected_plate_layout_for_sample(
    sample_number: int,
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker: str = "",
    configured_replicate: str = "",
) -> Dict[str, Any]:
    prefix, templates = build_layout_templates(
        layout_matches,
        marker,
        configured_replicate,
    )

    sample_number = max(int(sample_number), 1)

    if len(templates) == 1:
        template = templates[0]
        count = int(template["samples_per_plate"])
        plate_cycle = ((sample_number - 1) // count) + 1
        local_index = (sample_number - 1) % count
    else:
        remaining = sample_number
        template = None
        plate_cycle = 1
        local_index = 0

        for index, candidate in enumerate(templates, start=1):
            count = int(candidate["samples_per_plate"])
            if remaining <= count:
                template = candidate
                plate_cycle = index
                local_index = remaining - 1
                break
            remaining -= count

        if template is None:
            # After explicitly supplied plate-specific layouts are exhausted,
            # reuse the final supplied plate pattern for subsequent plates.
            template = templates[-1]
            count = int(template["samples_per_plate"])
            extra_plate_offset = (remaining - 1) // count
            plate_cycle = len(templates) + extra_plate_offset
            local_index = (remaining - 1) % count

    marker_number = template["marker_numbers"][local_index]

    return {
        "sample_number": sample_number,
        "template_marker": f"{prefix}{marker_number}",
        "marker_number": marker_number,
        "plate_cycle": plate_cycle,
        "layout_source": template["layout_source"],
        "expected_wells": list(template["wells_by_marker_number"][marker_number]),
        "samples_per_plate": int(template["samples_per_plate"]),
        "replicate_count": int(template["replicate_count"]),
        "replicate": template["replicate"],
    }

def validate_layout_matches(
    layout_matches: Dict[str, List[Dict[str, Any]]],
    pairs: List[TransferPair],
    marker: str,
    configured_replicate: str = "",
) -> Tuple[List[Dict[str, Any]], List[ValidationIssue]]:
    issues = []
    findings = []

    try:
        pattern = analyze_plate_layout_pattern(layout_matches, marker, configured_replicate)
    except Exception as exc:
        add_issue(issues, "critical", "plate layout pattern", str(exc))
        return [{"type": "layout_summary", "status": "fail", "message": str(exc)}], issues

    seen_wells = defaultdict(list)
    for number, wells in pattern.wells_by_marker_number.items():
        label = f"{pattern.marker_prefix}{number}"
        for well in wells:
            seen_wells[well].append(label)

    duplicates = {well: labels for well, labels in seen_wells.items() if len(labels) > 1}
    if duplicates:
        add_issue(
            issues,
            "major",
            "plate layout",
            "One or more template wells are assigned to multiple markers.",
            {"duplicate_destination_wells": duplicates},
        )

    findings.append({
        "type": "layout_summary",
        "status": "pass",
        "marker": pattern.marker_prefix,
        "marker_count": len(pattern.marker_numbers),
        "layout_marker_numbers": pattern.marker_numbers,
        "samples_per_plate": pattern.samples_per_plate,
        "replicate": pattern.replicate_name,
        "replicate_count": pattern.replicate_count,
        "unique_layout_position_count": len(seen_wells),
        "repeating_layout_template": (
            f"{pattern.marker_prefix}1 through {pattern.marker_prefix}{pattern.samples_per_plate} "
            "define one physical plate pattern; that same pattern repeats on each later plate."
        ),
    })

    return findings, issues


def generate_deterministic_transfer_pairs(
    configs: List[SequenceStepConfig],
    layout_matches: Dict[str, List[Dict[str, Any]]],
) -> List[TransferPair]:
    if not configs:
        return []

    marker = configs[0].marker
    replicate = configs[0].replicate
    replicate_count = REPLICATE_MAP.get(replicate, 1)

    pattern = None
    if layout_matches:
        try:
            pattern = analyze_plate_layout_pattern(layout_matches, marker, replicate)
            marker = pattern.marker_prefix
            replicate_count = pattern.replicate_count
        except Exception:
            pass

    tips = [c for c in configs if c.role == "tip pick up"]
    asps = [c for c in configs if c.role == "aspirate"]
    dsps = [c for c in configs if c.role == "dispense"]

    if not dsps:
        return []

    max_dispense_positions = max(int(c.sequence_count or 0) for c in dsps)
    total_samples = max((max_dispense_positions + replicate_count - 1) // replicate_count, 1)

    pairs = []

    for sample_number in range(1, total_samples + 1):
        tip = tips[(sample_number - 1) % len(tips)] if tips else None
        asp = asps[(sample_number - 1) % len(asps)] if asps else None
        dsp = dsps[(sample_number - 1) % len(dsps)] if dsps else None

        if pattern:
            expected = expected_plate_layout_for_sample(
                sample_number, layout_matches, marker, replicate
            )
            template_marker = expected["template_marker"]
            plate_cycle = expected["plate_cycle"]
            expected_wells = expected["expected_wells"]
        else:
            template_marker = f"{marker}{sample_number}" if marker else ""
            plate_cycle = 1
            expected_wells = []

        for rep_index in range(1, replicate_count + 1):
            dest = expected_wells[rep_index - 1] if rep_index - 1 < len(expected_wells) else ""
            pairs.append(TransferPair(
                pair_id=f"sample-{sample_number}-rep-{rep_index}",
                tip_sequence=tip.sequence if tip else "",
                source_sequence=asp.sequence if asp else "",
                destination_sequence=dsp.sequence if dsp else "",
                source_position=str(sample_number),
                destination_position=dest,
                source_sample_number=sample_number,
                destination_sample_number=extract_position_number(dest),
                replicate_index=rep_index,
                volume_ul=asp.volume_ul if asp else (dsp.volume_ul if dsp else 0.0),
                liquid_class=asp.liquid_class if asp else (dsp.liquid_class if dsp else ""),
                transfer_type=asp.transfer_type if asp else (dsp.transfer_type if dsp else ""),
                marker_label=f"{marker}{sample_number}" if marker else "",
                template_marker=template_marker,
                plate_cycle=plate_cycle,
            ))

    return pairs


def get_latest_trc_file(
    run_start_time: Optional[float] = None,
    wait_seconds: int = 45,
    trace_dir: str = DEFAULT_TRACE_DIR,
    require_after_start: bool = True,
) -> Optional[str]:
    pattern = os.path.join(trace_dir, "STAR_OEM_noFan_*_Trace.trc")
    end_time = time.time() + wait_seconds

    while time.time() < end_time:
        files = sorted(glob.glob(pattern), key=os.path.getmtime, reverse=True)

        if files:
            if run_start_time is None:
                return files[0]
            fresh = [f for f in files if os.path.getmtime(f) >= run_start_time - 2]
            if fresh:
                return fresh[0]
            if not require_after_start:
                return files[0]

        time.sleep(1)

    return None


def parse_trc_file(trc_path: str) -> Tuple[List[str], List[str], List[Dict[str, Any]]]:
    aspirate_steps = []
    dispense_steps = []
    ordered_steps = []

    channel_prefix = r"(?:\d+(?:\.\d+)?\s*[µu]?l\s+)?channel"

    patterns = [
        ("tip_pick", re.compile(rf"{channel_prefix}\s+tip\s+pick\s+up\s+\(single step\)\s*-\s*complete[:;]\s*(.*)", re.I)),
        ("aspirate", re.compile(rf"{channel_prefix}\s+aspirate\s+\(single step\)\s*-\s*complete[:;]\s*(.*)", re.I)),
        ("dispense", re.compile(rf"{channel_prefix}\s+dispense\s+\(single step\)\s*-\s*complete[:;]\s*(.*)", re.I)),
        ("tip_eject", re.compile(rf"{channel_prefix}\s+tip\s+eject\s+\(single step\)\s*-\s*complete[:;]\s*(.*)", re.I)),
        ("tip_pick_96_head", re.compile(r"co-?re\s+96\s+head\s+tip\s+pick\s+up.*?complete[:;]\s*(.*)", re.I)),
        ("aspirate_96_head", re.compile(r"co-?re\s+96\s+head\s+aspirate.*?complete[:;]\s*(.*)", re.I)),
        ("dispense_96_head", re.compile(r"co-?re\s+96\s+head\s+dispense.*?complete[:;]\s*(.*)", re.I)),
        ("tip_eject_96_head", re.compile(r"co-?re\s+96\s+head\s+tip\s+eject.*?complete[:;]\s*(.*)", re.I)),
    ]

    with open(trc_path, "r", encoding="latin-1", errors="ignore") as handle:
        for line in handle:
            for step_type, pattern in patterns:
                match = pattern.search(line.strip())
                if not match:
                    continue

                payload = match.group(1).strip()
                if "aspirate" in step_type:
                    aspirate_steps.append(payload)
                if "dispense" in step_type:
                    dispense_steps.append(payload)

                ordered_steps.append({
                    "order": len(ordered_steps) + 1,
                    "type": step_type,
                    "raw": payload,
                })
                break

    return aspirate_steps, dispense_steps, ordered_steps


def normalize_channel_steps(step_lines: List[str]) -> List[Tuple[str, str, str, str]]:
    normalized = []

    pattern = re.compile(
        r">\s*channel\s+(\d+):\s*([^,]+),\s*([A-Za-z]*\d+)\s+([\d.]+)\s*[µu]L",
        re.I,
    )
    fallback = re.compile(
        r"channel\s+(\d+).*?([A-Za-z0-9_\- .]+?),\s*\b([A-P]\d{1,2}|\d+)\b.*?([\d.]+)\s*[µu]L",
        re.I,
    )

    for line in step_lines:
        matches = pattern.findall(line) or fallback.findall(line)
        for channel, labware, position, volume in matches:
            normalized.append((
                channel.strip(),
                labware.strip(),
                normalize_position_text(position),
                volume.strip(),
            ))

    return normalized


def cross_match_positions(
    aspirate_steps: List[str],
    dispense_steps: List[str],
    replicate_value: str,
    ordered_steps: Optional[List[Dict[str, Any]]] = None,
) -> Tuple[bool, Any]:
    """
    Cross-match source samples to destination dispenses without losing reused
    channels. If ordered trace events are available, repeated aspirates from the
    same source labware/position are treated as the same logical sample. This
    supports workflows that aspirate again before each duplicate/triplicate.
    """
    replicate_count = REPLICATE_MAP.get(str(replicate_value).lower(), 1)

    if ordered_steps:
        logical_groups: Dict[Tuple[str, str], Dict[str, Any]] = {}
        logical_order: List[Tuple[str, str]] = []
        active_source_by_channel: Dict[str, Tuple[str, str]] = {}

        for event in ordered_steps:
            event_type = str(event.get("type", "")).lower()
            raw = str(event.get("raw", ""))

            if "aspirate" in event_type:
                for asp in normalize_channel_steps([raw]):
                    channel, labware, position, volume = asp
                    source_key = (labware, position)
                    if source_key not in logical_groups:
                        logical_groups[source_key] = {
                            "aspirate": asp,
                            "dispenses": [],
                        }
                        logical_order.append(source_key)
                    active_source_by_channel[channel] = source_key

            elif "dispense" in event_type:
                for dsp in normalize_channel_steps([raw]):
                    channel = dsp[0]
                    source_key = active_source_by_channel.get(channel)
                    if source_key is None:
                        return False, (
                            f"Channel {channel} dispensed before a matching "
                            "aspirate was found in the trace."
                        )
                    logical_groups[source_key]["dispenses"].append(dsp[1:])

        if logical_order:
            result = {}
            for sample_order, source_key in enumerate(logical_order, start=1):
                group = logical_groups[source_key]
                assigned = group["dispenses"]
                if len(assigned) != replicate_count:
                    return False, (
                        f"Source sample {source_key[0]} {source_key[1]} had "
                        f"{len(assigned)} dispense(s); {replicate_count} expected "
                        f"for {replicate_value}."
                    )
                asp = group["aspirate"]
                result[asp + (f"sample_order={sample_order}",)] = assigned
            return True, result

    # Fallback for older/manual traces without usable chronological events.
    aspirates = normalize_channel_steps(aspirate_steps)
    dispenses = normalize_channel_steps(dispense_steps)

    if not aspirates:
        return False, "No channel-level aspirates could be parsed."
    if not dispenses:
        return False, "No channel-level dispenses could be parsed."

    asp_by_channel = defaultdict(list)
    dsp_by_channel = defaultdict(list)

    for global_index, asp in enumerate(aspirates):
        asp_by_channel[asp[0]].append((global_index, asp))
    for dsp in dispenses:
        dsp_by_channel[dsp[0]].append(dsp)

    unknown = sorted(set(dsp_by_channel) - set(asp_by_channel))
    if unknown:
        return False, f"Dispense channels with no matching aspirate: {', '.join(unknown)}."

    records = []

    for channel, asp_records in asp_by_channel.items():
        channel_dispenses = dsp_by_channel.get(channel, [])

        # Collapse repeated aspirates of the same source position on one channel
        # into one logical sample before applying replicate count.
        unique_sources: List[Tuple[int, Tuple[str, str, str, str]]] = []
        seen_source_keys = set()
        for global_index, asp in asp_records:
            source_key = (asp[1], asp[2])
            if source_key in seen_source_keys:
                continue
            seen_source_keys.add(source_key)
            unique_sources.append((global_index, asp))

        expected_count = len(unique_sources) * replicate_count
        if len(channel_dispenses) != expected_count:
            return False, (
                f"Channel {channel}: {len(unique_sources)} logical aspirated sample(s), "
                f"{len(channel_dispenses)} dispense(s); {expected_count} expected "
                f"for {replicate_value}."
            )

        for occurrence, (global_index, asp) in enumerate(unique_sources):
            start = occurrence * replicate_count
            assigned = [
                dsp[1:]
                for dsp in channel_dispenses[start:start + replicate_count]
            ]
            records.append((global_index, asp, assigned))

    records.sort(key=lambda item: item[0])

    result = {}
    for sample_order, (_, asp, assigned) in enumerate(records, start=1):
        result[asp + (f"sample_order={sample_order}",)] = assigned

    return True, result

def validate_expected_runtime_steps(
    execution_plan: List[Dict[str, Any]],
    ordered_steps: List[Dict[str, Any]],
    trace_text: str,
) -> Tuple[List[Dict[str, Any]], List[ValidationIssue]]:
    expected = [
        item["action"]
        for item in execution_plan
        if item.get("action") in {"tip_pick", "aspirate", "dispense", "tip_eject"}
    ]

    actual = []
    for step in ordered_steps:
        step_type = str(step.get("type", ""))
        if "aspirate" in step_type:
            actual.append("aspirate")
        elif "dispense" in step_type:
            actual.append("dispense")
        elif "tip_pick" in step_type:
            actual.append("tip_pick")
        elif "tip_eject" in step_type:
            actual.append("tip_eject")

    error_match = re.search(r"main\s*-\s*error;\s*(.*?)(?:\n|$)", trace_text, re.I)
    runtime_error = error_match.group(1).strip() if error_match else ""

    exact_count = len(expected) == len(actual)
    order_match = actual == expected
    status = "pass" if exact_count and order_match and not runtime_error else "fail"

    finding = {
        "type": "expected_runtime_step_check",
        "status": status,
        "planned_liquid_handling_step_count": len(expected),
        "trace_liquid_handling_step_count": len(actual),
        "planned_actions": expected,
        "trace_actions": actual,
        "action_order_match": order_match,
        "runtime_error_reason": runtime_error,
    }

    issues = []
    if status == "fail":
        issues.append(ValidationIssue(
            "critical",
            "runtime step execution",
            f"Runtime validation failed. Planned {len(expected)} actions, trace showed "
            f"{len(actual)}; action order match={order_match}"
            + (f"; reason: {runtime_error}" if runtime_error else ""),
            finding,
        ))

    return [finding], issues


def validate_trace_labware_against_layout_format(
    aspirate_to_dispense_map: Dict[Any, Any],
    layout_matches: Dict[str, List[Dict[str, Any]]],
) -> Tuple[List[Dict[str, Any]], List[ValidationIssue]]:
    """
    Validate destination plate format without mistaking a physical plate ID
    such as P1 or Plate2 for a plate geometry. Unknown/plate-unknown names are
    accepted when their actual well coordinates fit the uploaded layout.
    """
    expected_format = summarize_layout_format(layout_matches).get(
        "inferred_plate_format",
        "unknown",
    )
    findings: List[Dict[str, Any]] = []
    issues: List[ValidationIssue] = []

    bounds = {
        "24-well": (4, 6),
        "48-well": (6, 8),
        "96-well": (8, 12),
        "384-well": (16, 24),
    }
    expected_bounds = bounds.get(expected_format)

    destinations_by_labware: Dict[str, List[str]] = defaultdict(list)
    for dispenses in aspirate_to_dispense_map.values():
        for dsp in dispenses:
            if len(dsp) < 2:
                continue
            labware = str(dsp[0]).strip()
            position = normalize_position_text(dsp[1])
            if labware and position:
                destinations_by_labware[labware].append(position)

    for labware, positions in destinations_by_labware.items():
        fmt = infer_labware_format_from_name(labware)
        status = "pass"
        invalid_positions: List[str] = []

        if expected_bounds is not None:
            max_rows, max_cols = expected_bounds
            for position in positions:
                match = re.fullmatch(r"([A-P])(\d{1,2})", position)
                if not match:
                    invalid_positions.append(position)
                    continue
                row_number = ord(match.group(1)) - ord("A") + 1
                col_number = int(match.group(2))
                if row_number > max_rows or col_number > max_cols:
                    invalid_positions.append(position)

        explicit_conflict = (
            expected_format != "unknown"
            and fmt not in {
                "unknown",
                "plate/unknown format",
                "mixed",
                expected_format,
            }
        )

        if invalid_positions or explicit_conflict:
            status = "fail"
            reason = (
                f"Destination labware '{labware}' used well(s) that do not fit "
                f"the uploaded {expected_format} layout: {invalid_positions}."
                if invalid_positions
                else f"Destination labware '{labware}' appears to be {fmt}, "
                     f"but the uploaded plate layout appears to be {expected_format}."
            )
            issues.append(ValidationIssue(
                "critical",
                "URGENT_LABWARE_MISMATCH",
                reason,
                {
                    "labware": labware,
                    "labware_format": fmt,
                    "expected_layout_format": expected_format,
                    "destination_positions": positions,
                    "invalid_destination_positions": invalid_positions,
                },
            ))

        findings.append({
            "type": "labware_format_check",
            "status": status,
            "labware": labware,
            "destination": positions,
            "labware_format": fmt,
            "expected_layout_format": expected_format,
            "invalid_destination_positions": invalid_positions,
        })

    return findings, issues

def validate_trace_against_layout(
    aspirate_to_dispense_map: Dict[Any, Any],
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker: str,
    configured_replicate: str,
) -> Tuple[List[Dict[str, Any]], List[ValidationIssue]]:
    """
    Validate sequential samples across one or more physical destination plates.

    One uploaded layout sheet/table is a reusable plate template. Therefore
    S1..S6 with 12 logical samples means samples 1..6 use plate 1 and samples
    7..12 use plate 2 with the same S1..S6 well pattern.

    If multiple Excel sheets/DOCX tables are supplied, each is used as the
    plate-specific template for the corresponding physical plate.
    """
    findings: List[Dict[str, Any]] = []
    issues: List[ValidationIssue] = []

    prefix, templates = build_layout_templates(
        layout_matches,
        marker,
        configured_replicate,
    )

    plate_cycle_to_labware: Dict[int, str] = {}
    labware_to_plate_cycle: Dict[str, int] = {}
    physical_destinations: List[Tuple[str, str]] = []

    for sample_index, (aspirate_step, dispenses) in enumerate(
        aspirate_to_dispense_map.items(),
        start=1,
    ):
        expected = expected_plate_layout_for_sample(
            sample_index,
            layout_matches,
            prefix,
            configured_replicate,
        )

        marker_label = f"{prefix}{sample_index}"
        template_marker = expected["template_marker"]
        plate_cycle = int(expected["plate_cycle"])
        expected_positions = [
            normalize_position_text(well)
            for well in expected["expected_wells"]
        ]

        source_labware = str(aspirate_step[1]).strip() if len(aspirate_step) > 1 else ""
        source_position = normalize_position_text(aspirate_step[2]) if len(aspirate_step) > 2 else ""

        actual_records = []
        for dispense in dispenses:
            if len(dispense) < 2:
                continue
            labware = str(dispense[0]).strip()
            position = normalize_position_text(dispense[1])
            if labware and position:
                actual_records.append((labware, position))
                physical_destinations.append((labware, position))

        actual_positions = [position for _, position in actual_records]
        unique_labwares: List[str] = []
        for labware, _ in actual_records:
            if labware not in unique_labwares:
                unique_labwares.append(labware)

        missing = sorted(set(expected_positions) - set(actual_positions))
        unexpected = sorted(set(actual_positions) - set(expected_positions))
        reasons: List[str] = []

        expected_replicate_count = int(expected["replicate_count"])
        if len(actual_positions) != expected_replicate_count:
            reasons.append(
                f"expected {expected_replicate_count} dispense(s) for "
                f"{expected['replicate']}, trace showed {len(actual_positions)}"
            )

        if missing or unexpected:
            reasons.append(
                f"expected well(s) {expected_positions}, trace showed {actual_positions}"
            )

        if len(unique_labwares) != 1:
            reasons.append(
                "all replicates for one sample must remain on exactly one "
                f"physical destination plate; trace used {unique_labwares}"
            )
            actual_labware = ""
        else:
            actual_labware = unique_labwares[0]
            required_labware = plate_cycle_to_labware.get(plate_cycle)

            if required_labware is None:
                prior_cycle = labware_to_plate_cycle.get(actual_labware)
                if prior_cycle is not None and prior_cycle != plate_cycle:
                    reasons.append(
                        f"plate cycle {plate_cycle} reused destination labware "
                        f"'{actual_labware}' already assigned to plate cycle {prior_cycle}"
                    )
                else:
                    plate_cycle_to_labware[plate_cycle] = actual_labware
                    labware_to_plate_cycle[actual_labware] = plate_cycle
            elif actual_labware != required_labware:
                reasons.append(
                    f"plate cycle {plate_cycle} must remain on destination labware "
                    f"'{required_labware}', but this sample used '{actual_labware}'"
                )

        status = "fail" if reasons else "pass"
        finding = {
            "type": "trace_layout_match",
            "sample_number": sample_index,
            "marker": marker_label,
            "template_marker": template_marker,
            "plate_layout_marker_used": template_marker,
            "marker_number": expected["marker_number"],
            "plate_cycle": plate_cycle,
            "layout_source": expected.get("layout_source", ""),
            "samples_per_plate": expected["samples_per_plate"],
            "replicate": expected["replicate"],
            "replicate_count": expected_replicate_count,
            "source_labware": source_labware,
            "source_position": source_position,
            "expected_destinations": expected_positions,
            "actual_destinations": actual_positions,
            "actual_destination_labware": unique_labwares,
            "missing_destinations": missing,
            "unexpected_destinations": unexpected,
            "status": status,
            "reason": "; ".join(reasons) if reasons else "Plate-layout crossmatch passed.",
        }
        findings.append(finding)

        findings.append({
            "type": "actual_trace_transfer",
            "marker": marker_label,
            "transfers": [{
                "source_labware": source_labware,
                "source_position": source_position,
                "destination_labware": unique_labwares,
                "destination_wells": actual_positions,
                "plate_cycle": plate_cycle,
                "template_marker": template_marker,
                "layout_source": expected.get("layout_source", ""),
            }],
        })

        if reasons:
            add_issue(
                issues,
                "critical",
                "trace vs plate layout",
                (
                    f"{marker_label} (plate cycle {plate_cycle}, template "
                    f"{template_marker}) failed plate-layout crossmatch: "
                    f"{'; '.join(reasons)}."
                ),
                finding,
            )

    counts = defaultdict(int)
    for destination in physical_destinations:
        counts[destination] += 1

    duplicate_physical_wells = sorted(
        f"{labware}:{well}"
        for (labware, well), count in counts.items()
        if count > 1
    )

    if duplicate_physical_wells:
        add_issue(
            issues,
            "major",
            "trace duplicate destinations",
            "The same physical destination well was used more than once.",
            {"duplicate_destination_wells": duplicate_physical_wells},
        )

    total_samples = len(aspirate_to_dispense_map)
    expected_plate_count = 0
    if total_samples:
        expected_plate_count = max(
            int(expected_plate_layout_for_sample(
                sample_number,
                layout_matches,
                prefix,
                configured_replicate,
            )["plate_cycle"])
            for sample_number in range(1, total_samples + 1)
        )

    actual_plate_order = [
        plate_cycle_to_labware[cycle]
        for cycle in sorted(plate_cycle_to_labware)
    ]

    plate_count_status = (
        "pass" if len(actual_plate_order) == expected_plate_count else "fail"
    )

    findings.append({
        "type": "multi_plate_layout_summary",
        "marker_prefix": prefix,
        "layout_template_sources": [t["layout_source"] for t in templates],
        "single_template_reused": len(templates) == 1 and expected_plate_count > 1,
        "total_samples_processed": total_samples,
        "expected_destination_plate_count": expected_plate_count,
        "actual_destination_plate_count": len(actual_plate_order),
        "actual_destination_plate_order": actual_plate_order,
        "status": plate_count_status,
    })

    if plate_count_status == "fail":
        add_issue(
            issues,
            "critical",
            "multi-plate destination continuity",
            (
                f"Trace used {len(actual_plate_order)} destination plate(s), but "
                f"{expected_plate_count} were expected for {total_samples} sample(s)."
            ),
            {
                "expected_destination_plate_count": expected_plate_count,
                "actual_destination_plate_count": len(actual_plate_order),
                "actual_destination_plate_order": actual_plate_order,
            },
        )

    labware_findings, labware_issues = validate_trace_labware_against_layout_format(
        aspirate_to_dispense_map,
        layout_matches,
    )
    findings.extend(labware_findings)
    issues.extend(labware_issues)

    return findings, issues

def validate_tracking_source_order(
    aspirate_to_dispense_map: Dict[Any, Any],
    marker: str,
) -> Tuple[List[Dict[str, Any]], List[ValidationIssue]]:
    findings = []

    for expected_order, (aspirate_step, dispenses) in enumerate(
        aspirate_to_dispense_map.items(),
        start=1,
    ):
        findings.append({
            "type": "tracking_source_order",
            "marker": f"{marker}{expected_order}",
            "expected_order": expected_order,
            "source_labware": (
                str(aspirate_step[1]).strip()
                if len(aspirate_step) > 1
                else ""
            ),
            "source_position": (
                str(aspirate_step[2]).strip()
                if len(aspirate_step) > 2
                else ""
            ),
            "dispense_count": len(dispenses),
            "status": "pass",
        })

    return findings, []


def build_labware_validation_summary(
    layout_matches: Dict[str, List[Dict[str, Any]]],
    trace_findings: List[Dict[str, Any]],
    configs: List[SequenceStepConfig],
    lay_metadata: Dict[str, Dict[str, Any]],
) -> List[Dict[str, Any]]:
    summary = []
    expected_format = summarize_layout_format(layout_matches).get(
        "inferred_plate_format",
        "unknown",
    )

    for config in configs:
        if config.role not in ("aspirate", "dispense"):
            continue

        meta = get_sequence_labware_metadata(
            config.sequence,
            lay_metadata,
            config.sequence_count,
        )

        labware_names = meta.get("labware_names", []) or []
        formats = meta.get("labware_formats", []) or []
        sequence_format = meta.get("labware_format", "unknown")

        status = "pass"

        if config.role == "dispense":
            for fmt in (formats or [sequence_format]):
                if (
                    expected_format != "unknown"
                    and fmt not in {
                        "unknown",
                        "plate/unknown format",
                        "mixed",
                        expected_format,
                    }
                ):
                    status = "fail"
                    break

        summary.append({
            "type": "labware_validation_summary",
            "status": status,
            "step_role": config.role,
            "sequence": config.sequence,
            "trace_or_method_labware": labware_names,
            "detected_labware_format": sequence_format,
            "detected_labware_formats": formats,
            "labware_count": len(labware_names),
            "position_count": meta.get("position_count", 0),
            "expected_layout_format": (
                expected_format
                if config.role == "dispense"
                else "not applicable to source"
            ),
            "message": (
                f"Sequence '{config.sequence}' references "
                f"{len(labware_names)} labware object(s): "
                f"{', '.join(labware_names) if labware_names else 'unknown'}. "
                f"Detected format(s): "
                f"{', '.join(formats) if formats else sequence_format}."
            ),
        })

    return summary


def infer_labware_orientation_from_layout(
    layout_matches: Dict[str, List[Dict[str, Any]]],
) -> Dict[str, Any]:
    summary = summarize_layout_format(layout_matches)
    row_count = int(summary.get("row_count") or 0)
    column_count = int(summary.get("column_count") or 0)

    if column_count > row_count:
        orientation = "landscape"
    elif row_count > column_count:
        orientation = "portrait"
    else:
        orientation = "square/unknown"

    return {
        "type": "labware_orientation",
        "orientation": orientation,
        "row_count": row_count,
        "column_count": column_count,
        "inferred_plate_format": summary.get("inferred_plate_format", "unknown"),
    }


def analyze_orientation_and_tip_usage(
    configs: List[SequenceStepConfig],
    execution_plan: List[Dict[str, Any]],
    layout_matches: Dict[str, List[Dict[str, Any]]],
    aspirate_steps: List[str],
    dispense_steps: List[str],
    cross_match: Any,
) -> Tuple[List[Dict[str, Any]], List[ValidationIssue]]:
    findings = [
        infer_labware_orientation_from_layout(layout_matches)
        if layout_matches
        else {
            "type": "labware_orientation",
            "orientation": "unknown/no layout",
            "row_count": 0,
            "column_count": 0,
            "inferred_plate_format": "unknown",
        }
    ]

    for role, step_lines in (
        ("aspirate", aspirate_steps),
        ("dispense", dispense_steps),
    ):
        for step_index, step_line in enumerate(step_lines, start=1):
            normalized = normalize_channel_steps([step_line])
            channels = sorted(
                {item[0] for item in normalized},
                key=lambda value: int(value) if str(value).isdigit() else 999,
            )
            findings.append({
                "type": "tip_usage_by_trace_step",
                "step_number": step_index,
                "step_role": role,
                "tips_used_at_this_step": len(channels),
                "channels_used": channels,
                "labware": sorted({item[1] for item in normalized if item[1]}),
                "positions": [item[2] for item in normalized],
            })

    # Old code compared every pickup against a single global max-channel count,
    # which can generate false "unused tip" warnings across unrelated cycles.
    return findings, []


def resolve_pyhamilton_resource_class(labware_name: str):
    text = str(labware_name or "").upper()
    if "384" in text:
        preferred = ["Plate384", "Plate96"]
    elif "24" in text:
        preferred = ["Plate24", "Plate96"]
    elif "TIP" in text or "SLIM" in text:
        preferred = ["Tip96", "TipRack", "Plate96"]
    else:
        preferred = ["Plate96"]

    for class_name in preferred:
        for module in PYHAMILTON_MODULES:
            if module is not None and hasattr(module, class_name):
                return getattr(module, class_name)
    return Plate96


def assign_sequence_labware_resource(layout_manager, labware_name: str):
    labware_name = str(labware_name or "").strip()
    if not labware_name:
        return None

    resource_class = resolve_pyhamilton_resource_class(labware_name)
    attempts = []

    if hasattr(layout_manager, "assign_resource"):
        attempts.append(lambda: layout_manager.assign_resource(labware_name))

    if ResourceType is not None and resource_class is not None:
        named = ResourceType(resource_class, labware_name)
        blank = ResourceType(resource_class, "")

        if hasattr(layout_manager, "assign_resource"):
            attempts.append(lambda: layout_manager.assign_resource(named))

        if hasattr(layout_manager, "assign_unused_resource"):
            attempts.append(lambda: layout_manager.assign_unused_resource(named))
            attempts.append(lambda: layout_manager.assign_unused_resource(blank))

    last_error = None
    for attempt in attempts:
        try:
            resource = attempt()
            if resource is not None:
                return resource
        except Exception as exc:
            last_error = exc

    raise RuntimeError(
        f"Could not assign labware resource for '{labware_name}'. "
        f"Last error: {last_error}"
    )


def run_pyhamilton_simulation(
    lay_path: str,
    configs: List[SequenceStepConfig],
    execution_plan: List[Dict[str, Any]],
) -> Tuple[Optional[str], List[ValidationIssue]]:
    issues = []

    if phi is None or LayoutManager is None:
        return None, [
            ValidationIssue(
                "critical",
                "runtime",
                f"pyhamilton is not available: {PYHAMILTON_IMPORT_ERROR}",
            )
        ]

    run_start_time = time.time()
    resource_cache = {}

    try:
        # Do not instantiate LayoutManager before Hamilton starts.  On large
        # VENUS .lay files that constructor can be slow and was delaying the
        # Run Control window.  Channel-only methods do not need it at all; for
        # CO-RE 96/head steps it is created lazily after Run Control is open.
        layout_manager = None

        with AutoRunHamiltonInterface(simulate=True) as ham_int:
            normal_logging(ham_int, os.getcwd())
            initialize(ham_int)

            # Run Control is already open at this point.  Resolve LAY sequence
            # positions now (not before launch) so channel commands can use
            # explicit positions and are immune to an exhausted VENUS sequence
            # cursor.  If metadata parsing fails, the sequence-name fallback is
            # retained and VENUS will report the underlying layout issue.
            try:
                runtime_lay_metadata = parse_lay_sequence_metadata(lay_path)
            except Exception:
                runtime_lay_metadata = {}

            for item in execution_plan:
                action = str(item.get("action") or "").strip()
                seq = str(item.get("sequence") or "").strip()
                mode = str(item.get("hardware_mode") or "channels").strip().lower()
                channel = normalize_assigned_channel_pattern(
                    item.get("channel_pattern") or "111111111111",
                    mode,
                )
                volume = float(item.get("volume_ul") or 0)
                liquid_class = str(item.get("liquid_class") or "").strip()

                if action == "tip_pick":
                    if mode == "head":
                        ph_tip_pick_up_96_seq(ham_int, tip96_seq=seq)
                    else:
                        explicit_positions = _sequence_positions_for_plan_item(
                            item, runtime_lay_metadata
                        )
                        if explicit_positions:
                            ph_tip_pick_up_positions(
                                ham_int,
                                labware_positions=explicit_positions,
                                channel=channel,
                            )
                        else:
                            ph_tip_pick_up_seq(ham_int, tipseq=seq, channel=channel)

                elif action == "aspirate":
                    if mode == "head":
                        if layout_manager is None:
                            layout_manager = LayoutManager(lay_path)
                        labware_name = str(item.get("labware_name") or seq).strip()
                        cache_key = labware_name.lower()
                        if cache_key not in resource_cache:
                            resource_cache[cache_key] = assign_sequence_labware_resource(
                                layout_manager,
                                labware_name,
                            )
                        ph_aspirate_96_seq(
                            ham_int,
                            plate96=resource_cache[cache_key],
                            head_asp_seq=seq,
                            vols=volume,
                            liq_class=liquid_class,
                        )
                    else:
                        explicit_positions = _sequence_positions_for_plan_item(
                            item, runtime_lay_metadata
                        )
                        if explicit_positions:
                            ph_aspirate_positions(
                                ham_int,
                                labware_positions=explicit_positions,
                                vols=volume,
                                channel=channel,
                                liq_class=liquid_class,
                            )
                        else:
                            ph_aspirate_seq(
                                ham_int,
                                asp_seq=seq,
                                vols=volume,
                                channel=channel,
                                liq_class=liquid_class,
                            )

                elif action == "dispense":
                    if mode == "head":
                        if layout_manager is None:
                            layout_manager = LayoutManager(lay_path)
                        labware_name = str(item.get("labware_name") or seq).strip()
                        cache_key = labware_name.lower()
                        if cache_key not in resource_cache:
                            resource_cache[cache_key] = assign_sequence_labware_resource(
                                layout_manager,
                                labware_name,
                            )
                        ph_dispense_96_seq2(
                            ham_int,
                            plate96=resource_cache[cache_key],
                            head_disp_seq=seq,
                            vols=volume,
                            liq_class=liquid_class,
                        )
                    else:
                        explicit_positions = _sequence_positions_for_plan_item(
                            item, runtime_lay_metadata
                        )
                        if explicit_positions:
                            ph_dispense_positions(
                                ham_int,
                                labware_positions=explicit_positions,
                                vols=volume,
                                channel=channel,
                                liq_class=liquid_class,
                            )
                        else:
                            ph_dispense_seq(
                                ham_int,
                                disp_seq=seq,
                                vols=volume,
                                channel=channel,
                                liq_class=liquid_class,
                            )

                elif action == "tip_eject":
                    if mode == "head":
                        ph_tip_eject_96(ham_int, sequence=seq)
                    else:
                        ph_tip_eject_seq2(
                            ham_int,
                            waste_seq=seq or "Waste",
                            channel=channel,
                        )

    except Exception as exc:
        issues.append(ValidationIssue(
            "critical",
            "runtime",
            f"PyHamilton simulation failed: {type(exc).__name__}: {exc}",
        ))

    # Poll immediately. get_latest_trc_file already waits/retries, so a fixed
    # five-second sleep only adds latency and makes the application feel hung.
    trace_path = get_latest_trc_file(
        run_start_time=run_start_time,
        wait_seconds=45,
        trace_dir=DEFAULT_TRACE_DIR,
        require_after_start=True,
    )

    if not trace_path:
        issues.append(ValidationIssue(
            "major",
            "trace",
            "No new STAR_OEM_noFan trace file was found for this simulation.",
            {"trace_dir": DEFAULT_TRACE_DIR},
        ))

    return trace_path, issues


def generate_pyhamilton_review_script(
    configs: List[SequenceStepConfig],
    execution_plan: List[Dict[str, Any]],
    output_dir: str,
) -> str:
    os.makedirs(output_dir, exist_ok=True)

    script_path = os.path.join(
        output_dir,
        f"devpal_lite_review_required_pyhamilton_script_{datetime.now().strftime('%H%M%S')}.py",
    )

    payload = {
        "review_required": True,
        "generated_at": datetime.now().isoformat(),
        "selected_sequences": [asdict(config) for config in configs],
        "execution_plan": execution_plan,
    }

    lines = [
        "from pyhamilton import *",
        "",
        "REVIEW_REQUIRED = True",
        f"DEV_PAL_LITE_PLAN = {json.dumps(payload, indent=4)}",
        "",
        "def run_review_required_plan(ham_int):",
        "    if REVIEW_REQUIRED:",
        "        raise RuntimeError('Review required before running this generated PyHamilton script.')",
        "    initialize(ham_int)",
    ]

    for item in execution_plan:
        action = str(item.get("action") or "")
        seq = str(item.get("sequence") or "")
        mode = str(item.get("hardware_mode") or "channels").lower()
        channel = normalize_assigned_channel_pattern(
            item.get("channel_pattern") or "111111111111",
            mode,
        )
        volume = float(item.get("volume_ul") or 0)
        liquid_class = str(item.get("liquid_class") or "")

        lines.append("")
        lines.append(
            f"    # Order {item.get('order')} | Set {item.get('set_order')} "
            f"| {action} | {seq}"
        )

        if (
            item.get("autoincrement")
            and action in {"aspirate", "dispense"}
            and not item.get("manual")
        ):
            lines.append(
                f"    inc_sequence(ham_int, sequence={seq!r}, "
                f"increment={int(item.get('increment') or 0)!r})"
            )

        if action == "tip_pick":
            if mode == "head":
                lines.append(f"    tip_pick_up_96_seq(ham_int, tip96_seq={seq!r})")
            else:
                lines.append(
                    f"    tip_pick_up_seq(ham_int, tipseq={seq!r}, channel={channel!r})"
                )
        elif action == "aspirate":
            if mode == "head":
                lines.append(
                    "    raise RuntimeError('96-head review script requires "
                    "manual resource binding before execution.')"
                )
            else:
                lines.append(
                    f"    aspirate_seq(ham_int, asp_seq={seq!r}, vols={volume!r}, "
                    f"channel={channel!r}, liq_class={liquid_class!r})"
                )
        elif action == "dispense":
            if mode == "head":
                lines.append(
                    "    raise RuntimeError('96-head review script requires "
                    "manual resource binding before execution.')"
                )
            else:
                lines.append(
                    f"    dispense_seq(ham_int, disp_seq={seq!r}, vols={volume!r}, "
                    f"channel={channel!r}, liq_class={liquid_class!r})"
                )
        elif action == "tip_eject":
            if mode == "head":
                lines.append(f"    tip_eject_96(ham_int, sequence={seq!r})")
            else:
                lines.append(
                    f"    tip_eject_seq2(ham_int, waste_seq={seq!r}, channel={channel!r})"
                )

    Path(script_path).write_text("\n".join(lines) + "\n", encoding="utf-8")
    return script_path


def humanize_key(key: str) -> str:
    return str(key or "").replace("_", " ").replace("-", " ").strip().title()


def safe_text(value: Any) -> str:
    if value is None:
        return "None"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value)
    text = text.replace("{", "").replace("}", "")
    text = text.replace("[", "").replace("]", "")
    text = text.replace("'", "").replace('"', "")
    text = text.replace("_", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text if text else "None"


def join_values(values: Iterable[Any]) -> str:
    cleaned = [
        safe_text(value)
        for value in values
        if safe_text(value) not in {"", "None"}
    ]
    return ", ".join(cleaned) if cleaned else "None"


def normal_datetime() -> str:
    return datetime.now().strftime("%B %d, %Y at %I:%M %p")


def is_error_record(value: Any) -> bool:
    text = str(value).lower()
    return any(
        word in text
        for word in (
            "urgent",
            "critical",
            "error",
            "fail",
            "mismatch",
            "wrong labware",
            "missing",
            "unexpected",
            "not valid",
        )
    )


def _report_role_priority(role: Any) -> int:
    """Canonical Hamilton execution order for the What Was Run summary."""
    return {
        "tip pick up": 0,
        "aspirate": 1,
        "dispense": 2,
        TIP_EJECT_ROLE: 3,
    }.get(str(role or "").strip().lower(), 99)


def _format_trace_layout_issue(issue: Dict[str, Any]) -> str:
    """Render trace-vs-layout failures in the concise Version-1 report style."""
    evidence = issue.get("evidence") or {}
    marker = safe_text(evidence.get("marker") or evidence.get("template_marker") or "Sample")
    expected = join_values(evidence.get("expected_destinations", []) or [])
    actual = join_values(evidence.get("actual_destinations", []) or [])

    if expected != "None" or actual != "None":
        return (
            f"{marker}: sample was dispensed to the wrong destination well(s). "
            f"Expected {expected}; trace showed {actual}."
        )
    return safe_text(issue.get("message", "Trace-versus-plate-layout validation failed."))


def build_major_critical_groups(major_critical: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Group major/critical findings by validation category.  Trace-vs-layout
    failures intentionally use the concise, human-readable Version-1 wording
    instead of exposing the internal cross-match diagnostic string.
    """
    grouped: Dict[str, List[Dict[str, Any]]] = {}
    category_order: List[str] = []

    for issue in major_critical:
        raw_category = str(issue.get("category") or "validation").strip()
        key = raw_category.lower()
        if key not in grouped:
            grouped[key] = []
            category_order.append(key)
        grouped[key].append(issue)

    # Keep the original report's most actionable category first.
    if "trace vs plate layout" in category_order:
        category_order.remove("trace vs plate layout")
        category_order.insert(0, "trace vs plate layout")

    result: List[Dict[str, Any]] = []
    for number, key in enumerate(category_order, start=1):
        items = grouped[key]
        severities = {str(item.get("severity") or "").lower() for item in items}
        severity_label = "Critical" if "critical" in severities else "Major"

        if key == "trace vs plate layout":
            title = f"{number}. {severity_label} issue with Trace Vs Plate Layout"
            lines = [_format_trace_layout_issue(item) for item in items]
        else:
            category_label = humanize_key(key) or "Validation"
            title = f"{number}. {severity_label} issue with {category_label}"
            lines = [safe_text(item.get("message", "")) for item in items]

        result.append({
            "title": title,
            "lines": [line for line in lines if line and line != "None"],
            "severity": severity_label.lower(),
            "category": key,
        })

    return result


def build_report_sections(output: DevPalLiteOutput, report_path: str) -> Dict[str, Any]:
    major_critical = [
        issue
        for issue in output.issues
        if str(issue.get("severity", "")).lower() in {"critical", "major"}
    ]

    # Report configured steps in the same logical order Hamilton executes them,
    # regardless of the order in which sequence checkboxes appeared/selected.
    # selected_index is still used as the stable tie-breaker when a role repeats.
    what_was_run = []
    for config in sorted(
        output.selected_sequences,
        key=lambda c: (
            int(c.get("set_order", 1) or 1),
            _report_role_priority(c.get("role")),
            int(c.get("selected_index", 0) or 0),
            str(c.get("sequence") or "").lower(),
        ),
    ):
        role = str(config.get("role") or "").strip()
        line = (
            f"Set {config.get('set_order', 1)}: {role} - "
            f"{config.get('sequence')}; hardware {config.get('hardware_mode')}; "
            f"channel pattern {config.get('channel_pattern')}; "
            f"sequence count {config.get('sequence_count')}"
        )
        if role in {"aspirate", "dispense"}:
            line += (
                f"; {config.get('volume_ul')} uL; "
                f"LC {config.get('liquid_class')}"
            )
        what_was_run.append(line)

    marker_summary = []
    for marker_label, entries in sorted(
        output.marker_matches.items(),
        key=lambda item: extract_position_number(item[0]) or 0,
    ):
        wells = [
            safe_text(entry.get("dest_value"))
            for entry in entries
            if safe_text(entry.get("dest_value")) not in {"", "None"}
        ]
        if wells:
            marker_summary.append(f"{marker_label}: {', '.join(wells)}")

    layout_summary = []
    for finding in output.layout_findings:
        if finding.get("type") == "layout_summary":
            layout_summary.append(
                f"Marker {finding.get('marker')}: "
                f"{finding.get('samples_per_plate')} sample(s)/plate; "
                f"{finding.get('replicate')} "
                f"({finding.get('replicate_count')} destination(s)/sample); "
                f"status {finding.get('status')}."
            )

    # Trace / Runtime Summary: keep the clear diagnostic style from the
    # original report rather than reducing the trace to generic status lines.
    trace_summary = []
    trace_record = next(
        (f for f in output.trace_findings if f.get("type") == "trace_summary"),
        None,
    )
    if trace_record:
        asp_count = int(trace_record.get("aspirate_step_count") or 0)
        dsp_count = int(trace_record.get("dispense_step_count") or 0)
        cross_status = str(trace_record.get("cross_match_status") or "").lower()
        cross_text = (
            "Plate layout crossmatch passed."
            if cross_status == "pass"
            else "Plate layout crossmatch failed."
        )
        trace_summary.append(
            f"Trace parsed: {asp_count} aspirate step(s), {dsp_count} dispense step(s). "
            f"{cross_text}"
        )

    orientation = next(
        (f for f in output.trace_findings if f.get("type") == "labware_orientation"),
        None,
    )
    if orientation:
        trace_summary.append(
            f"Labware orientation: {orientation.get('orientation')}; "
            f"{orientation.get('row_count')} row(s), "
            f"{orientation.get('column_count')} column(s), detected as "
            f"{orientation.get('inferred_plate_format')}."
        )

    for finding in output.trace_findings:
        if finding.get("type") != "tip_usage_by_trace_step":
            continue
        role = str(finding.get("step_role") or "").strip().capitalize()
        step_number = finding.get("step_number")
        tips = int(finding.get("tips_used_at_this_step") or 0)
        channels = join_values(finding.get("channels_used", []))
        labware = join_values(finding.get("labware", []))
        positions = join_values(finding.get("positions", []))
        trace_summary.append(
            f"{role} step {step_number}: used {tips} tip(s) at one time; "
            f"channels {channels}; labware {labware}; positions {positions}."
        )

    # Report each actual pickup cycle, not merely the one selected configuration.
    trace_tip_max = max(
        [
            int(f.get("tips_used_at_this_step") or 0)
            for f in output.trace_findings
            if f.get("type") == "tip_usage_by_trace_step"
        ]
        or [0]
    )
    for item in output.execution_plan:
        if str(item.get("action") or "") != "tip_pick":
            continue
        seq = str(item.get("sequence") or "")
        mode = str(item.get("hardware_mode") or "channels").lower()
        if mode == "head":
            picked = 96
        else:
            picked = count_active_channels(str(item.get("channel_pattern") or ""))
        used = min(trace_tip_max or picked, picked)
        unused = max(picked - used, 0)
        trace_summary.append(
            f"Tip pickup usage: {seq} picked up {picked} tip(s); "
            f"trace used up to {used} tip(s) at one time; unused picked tips: {unused}."
        )

    for finding in output.trace_findings:
        if finding.get("type") != "trace_layout_match" or finding.get("status") != "fail":
            continue
        marker_label = finding.get("marker") or finding.get("template_marker") or "Sample"
        expected = join_values(finding.get("expected_destinations", []))
        actual = join_values(finding.get("actual_destinations", []))
        trace_summary.append(
            f"Mismatch - {marker_label}: expected plate layout wells {expected}; "
            f"trace dispensed to {actual}."
        )

    # Keep other high-level runtime validation results only when they add
    # information not already expressed above.
    for finding in output.trace_findings:
        if finding.get("type") == "expected_runtime_step_check" and finding.get("status") == "fail":
            trace_summary.append(
                f"Runtime action check failed: planned "
                f"{finding.get('planned_liquid_handling_step_count')} step(s); trace showed "
                f"{finding.get('trace_liquid_handling_step_count')}; "
                f"action order match {finding.get('action_order_match')}."
            )

    return {
        "major_critical": major_critical,
        "major_critical_groups": build_major_critical_groups(major_critical),
        "what_was_run": what_was_run,
        "marker_summary": marker_summary,
        "layout_summary": layout_summary,
        "trace_summary": trace_summary,
        "summary": {
            "Generated": normal_datetime(),
            "LAY File": output.lay_file,
            "Plate Layout File": output.layout_file,
            "Report Path": report_path,
            "Number of Critical Issues": sum(
                1
                for issue in output.issues
                if str(issue.get("severity", "")).lower() == "critical"
            ),
        },
    }


def write_docx_report(output: DevPalLiteOutput, output_dir: str) -> str:
    if Document is None:
        return write_text_report(output, output_dir)

    path = os.path.join(
        output_dir,
        f"DevPal_Lite_Validation_Report_{datetime.now().strftime('%H%M%S')}.docx",
    )
    sections = build_report_sections(output, path)

    doc = Document()
    doc.add_heading("DevPal Sequence Analyzer Report", 0)

    doc.add_heading("Run Summary", level=1)
    for label, value in sections["summary"].items():
        p = doc.add_paragraph()
        p.add_run(f"{label}: ").bold = True
        p.add_run(safe_text(value))

    doc.add_heading("Major / Critical Issues", level=1)
    if sections["major_critical_groups"]:
        for group in sections["major_critical_groups"]:
            heading_p = doc.add_paragraph()
            heading_run = heading_p.add_run(group["title"])
            heading_run.bold = True
            if RGBColor is not None:
                heading_run.font.color.rgb = RGBColor(255, 0, 0)

            for line in group["lines"]:
                p = doc.add_paragraph()
                run = p.add_run(f"- {safe_text(line)}")
                if RGBColor is not None:
                    run.font.color.rgb = RGBColor(255, 0, 0)
    else:
        doc.add_paragraph("No major or critical issues found.")

    for title, key in (
        ("What Was Run", "what_was_run"),
        ("Plate Layout Marker Review", "marker_summary"),
        ("Layout Summary", "layout_summary"),
        ("Trace / Runtime Summary", "trace_summary"),
    ):
        doc.add_heading(title, level=1)
        lines = sections[key]
        if not lines:
            doc.add_paragraph("None found.")
        else:
            for line in lines:
                p = doc.add_paragraph()
                run = p.add_run(f"• {safe_text(line)}")
                # In Trace / Runtime Summary, failures/mismatches are red;
                # informational/pass lines remain the document's default color.
                if title == "Trace / Runtime Summary" and is_error_record(line):
                    if RGBColor is not None:
                        run.font.color.rgb = RGBColor(255, 0, 0)

    doc.save(path)
    return path


def write_text_report(output: DevPalLiteOutput, output_dir: str) -> str:
    path = os.path.join(
        output_dir,
        f"DevPal_Lite_Validation_Report_{datetime.now().strftime('%H%M%S')}.txt",
    )
    sections = build_report_sections(output, path)

    with open(path, "w", encoding="utf-8") as handle:
        handle.write("DevPal Lite Validation Report\n")
        handle.write("=" * 80 + "\n\n")

        for label, value in sections["summary"].items():
            handle.write(f"{label}: {safe_text(value)}\n")

        major_lines = []
        for group in sections["major_critical_groups"]:
            major_lines.append(group["title"])
            major_lines.extend(f"- {line}" for line in group["lines"])

        blocks = [
            ("Major / Critical Issues", major_lines),
            ("What Was Run", sections["what_was_run"]),
            ("Plate Layout Marker Review", sections["marker_summary"]),
            ("Layout Summary", sections["layout_summary"]),
            ("Trace / Runtime Summary", sections["trace_summary"]),
        ]

        for title, lines in blocks:
            handle.write(f"\n{title}\n")
            handle.write("-" * 80 + "\n")
            if lines:
                for line in lines:
                    handle.write(f"- {safe_text(line)}\n")
            else:
                handle.write("None found.\n")

    return path


def write_pdf_report(output: DevPalLiteOutput, output_dir: str) -> str:
    if SimpleDocTemplate is None:
        return write_text_report(output, output_dir)

    path = os.path.join(
        output_dir,
        f"DevPal_Lite_Validation_Report_{datetime.now().strftime('%H%M%S')}.pdf",
    )
    sections = build_report_sections(output, path)

    styles = getSampleStyleSheet()
    story = [
        Paragraph("DevPal Lite Validation Report", styles["Title"]),
        Spacer(1, 12),
    ]

    summary_rows = [
        [safe_text(key), safe_text(value)]
        for key, value in sections["summary"].items()
    ]
    table = Table(summary_rows, colWidths=[150, 360])
    table.setStyle(
        TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("BACKGROUND", (0, 0), (0, -1), colors.lightgrey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ])
    )
    story.append(table)

    # Fixed from DP: use actual section keys. The original referenced
    # nonexistent "selected_sequences" and "transfer_summary".
    major_pdf_lines = []
    for group in sections["major_critical_groups"]:
        major_pdf_lines.append(group["title"])
        major_pdf_lines.extend(f"- {line}" for line in group["lines"])

    report_blocks = [
        ("Major / Critical Issues", major_pdf_lines),
        ("What Was Run", sections["what_was_run"]),
        ("Plate Layout Marker Review", sections["marker_summary"]),
        ("Layout Summary", sections["layout_summary"]),
        ("Trace / Runtime Summary", sections["trace_summary"]),
    ]

    for title, lines in report_blocks:
        story.append(Spacer(1, 10))
        story.append(Paragraph(safe_text(title), styles["Heading1"]))

        if not lines:
            story.append(Paragraph("None found.", styles["BodyText"]))
        else:
            for line in lines:
                clean = safe_text(line)
                if is_error_record(clean):
                    clean = f'<font color="red">{clean}</font>'
                story.append(Paragraph(f"• {clean}", styles["BodyText"]))

    SimpleDocTemplate(path, pagesize=letter).build(story)
    return path


def create_output_model(
    lay_file: str,
    layout_file: str,
    trace_file: str,
    configs: List[SequenceStepConfig],
    pairs: List[TransferPair],
    execution_plan: List[Dict[str, Any]],
    marker_matches: Dict[str, List[Dict[str, Any]]],
    layout_findings: List[Dict[str, Any]],
    trace_findings: List[Dict[str, Any]],
    issues: List[ValidationIssue],
    limitations: List[str],
    review_script_path: str,
    report_path: str = "",
) -> DevPalLiteOutput:
    validation_findings = [asdict(issue) for issue in issues]

    return DevPalLiteOutput(
        lay_file=lay_file,
        layout_file=layout_file,
        trace_file=trace_file,
        selected_sequences=[asdict(config) for config in configs],
        source_positions=sorted({
            pair.source_position
            for pair in pairs
            if pair.source_position
        }),
        destination_positions=sorted({
            pair.destination_position
            for pair in pairs
            if pair.destination_position
        }),
        transfer_pairs=[asdict(pair) for pair in pairs],
        execution_plan=execution_plan,
        marker_matches=marker_matches,
        layout_findings=layout_findings,
        trace_findings=trace_findings,
        validation_findings=validation_findings,
        issues=validation_findings,
        limitations=limitations,
        generated_review_script=review_script_path,
        report_path=report_path,
    )


class SearchableCombobox(ttk.Combobox):
    def __init__(self, master=None, completevalues=None, **kwargs):
        self.completevalues = sorted(completevalues or [], key=str.lower)
        super().__init__(master, values=self.completevalues, **kwargs)
        self.bind("<KeyRelease>", self._filter_values)

    def _filter_values(self, event):
        typed = self.get().strip().lower()
        if not typed:
            self["values"] = self.completevalues
            return
        self["values"] = [
            value for value in self.completevalues
            if typed in value.lower()
        ]


class DevPalLiteApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DevPal Lite Sequence Layout Runner")
        self.root.geometry("1180x780")
        self.root.minsize(1000, 650)
        self.root.option_add("*Font", APP_FONT)

        self.lay_path = ""
        self.layout_path = ""
        self.trace_path = ""

        self.sequence_names = []
        self.sequence_counts = {}
        self.liquid_classes = []

        self.sequence_vars = []
        self.sequence_role_vars = {}
        self.sequence_role_widgets = {}
        self.sequence_channel_vars = {}
        self.sequence_hardware_vars = {}
        self.sequence_count_vars = {}
        self.sequence_manual_vars = {}
        self.sequence_control_vars = {}
        self.sequence_set_vars = {}
        self.sequence_set_widgets = {}

        self.liquid_class_var = tk.StringVar()
        self.transfer_type_var = tk.StringVar(value="samples")
        self.replicate_var = tk.StringVar(value="single")
        self.volume_var = tk.StringVar(value="100")
        self.marker_var = tk.StringVar()
        self.report_format_var = tk.StringVar(value="docx")

        self.tracking_var = tk.IntVar(value=0)
        self.variant_var = tk.IntVar(value=0)
        self.unique_tip_eject_var = tk.IntVar(value=0)

        self.set_liquid_classes = []
        self.set_liquid_summary_var = tk.StringVar(
            value="No set liquid classes selected."
        )

        self.background_image_original = None
        self.background_image_tk = None

        # Cached parsed LAY metadata for the currently loaded method.  The old
        # Run path parsed the same .lay twice before Hamilton could even open.
        self._cached_lay_metadata_path = ""
        self._cached_lay_metadata_mtime = None
        self._cached_lay_metadata: Dict[str, Dict[str, Any]] = {}
        # Background plate-layout auto-detection token.  Marker/replicate
        # detection can involve pandas/openpyxl/python-docx, so never run it
        # synchronously on Tk's UI thread.
        self._layout_detection_generation = 0

        self._build_ui()

    def make_checkbutton(
        self,
        parent,
        text="",
        variable=None,
        command=None,
        borderless=False,
    ):
        """
        Create a persistent Tk checkbox.

        The previous UI used a white selection background together with white
        foreground/check-mark rendering on Windows. The variable actually
        toggled, but the selected indicator became visually blank immediately
        after the click, which made every checkbox look as though it had
        unchecked itself.

        Use an explicit integer on/off value and a contrasting purple selected
        indicator so the checked state remains visibly checked.
        """
        if variable is None:
            variable = tk.IntVar(master=self.root, value=0)

        checkbox = tk.Checkbutton(
            parent,
            text=text,
            variable=variable,
            onvalue=1,
            offvalue=0,
            bg=UI_BG,
            fg=UI_TEXT_FG,
            activebackground=UI_BG,
            activeforeground=UI_TEXT_FG,
            selectcolor=UI_TEXT_BG,
            indicatoron=True,
            highlightthickness=0,
            highlightbackground=UI_BG,
            highlightcolor=UI_BG,
            anchor="w",
            padx=2,
            pady=1,
            bd=0 if borderless else 2,
            relief="flat" if borderless else "raised",
            takefocus=True,
        )

        if command is not None:
            checkbox.configure(command=command)

        return checkbox

    def resize_background_image(self, event=None):
        if self.background_image_original is None or ImageTk is None:
            return

        try:
            width = max(self.background_canvas.winfo_width(), 1)
            height = max(self.background_canvas.winfo_height(), 1)
            image = self.background_image_original.copy()

            image_ratio = image.width / image.height
            canvas_ratio = width / height

            if image_ratio > canvas_ratio:
                new_height = height
                new_width = int(height * image_ratio)
            else:
                new_width = width
                new_height = int(width / image_ratio)

            image = image.resize((new_width, new_height))
            left = max((new_width - width) // 2, 0)
            top = max((new_height - height) // 2, 0)
            image = image.crop((left, top, left + width, top + height))

            self.background_image_tk = ImageTk.PhotoImage(image)
            self.background_canvas.delete("all")
            self.background_canvas.create_image(
                0,
                0,
                image=self.background_image_tk,
                anchor="nw",
            )
        except Exception:
            pass

    def _build_ui(self):
        self.root.configure(bg=UI_BG)

        icon_path = os.path.join(os.getcwd(), "thumbs_up.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                pass

        self.main_container = tk.Frame(self.root, bg=UI_BG)
        self.main_container.pack(fill="both", expand=True)
        self.main_container.grid_rowconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(0, weight=3)
        self.main_container.grid_columnconfigure(1, weight=2)

        self.left_ui_frame = tk.Frame(self.main_container, bg=UI_BG)
        self.left_ui_frame.grid(row=0, column=0, sticky="nsew")

        self.right_image_frame = tk.Frame(self.main_container, bg=UI_BG)
        self.right_image_frame.grid(row=0, column=1, sticky="nsew")

        self.background_canvas = tk.Canvas(
            self.right_image_frame,
            bg=UI_BG,
            highlightthickness=0,
            bd=0,
        )
        self.background_canvas.pack(fill="both", expand=True)

        background_path = os.path.join(os.getcwd(), "backgroundSA.png")
        if os.path.exists(background_path) and Image is not None:
            try:
                self.background_image_original = Image.open(background_path)
            except Exception:
                self.background_image_original = None

        self.background_canvas.bind(
            "<Configure>",
            self.resize_background_image,
        )

        def purple_label(parent, text, **kwargs):
            return tk.Label(
                parent,
                text=text,
                bg=UI_TEXT_BG,
                fg=UI_TEXT_FG,
                bd=2,
                relief="raised",
                padx=5,
                pady=2,
                **kwargs,
            )

        def purple_button(parent, text, command, **kwargs):
            return tk.Button(
                parent,
                text=text,
                command=command,
                bg=UI_BUTTON_BG,
                fg=UI_TEXT_FG,
                activebackground=UI_BUTTON_ACTIVE,
                activeforeground=UI_TEXT_FG,
                bd=3,
                relief="raised",
                padx=8,
                pady=3,
                **kwargs,
            )

        top = tk.Frame(self.left_ui_frame, bg=UI_BG)
        top.pack(fill="x", padx=8, pady=(8, 4))

        lay_row = tk.Frame(top, bg=UI_BG)
        lay_row.pack(fill="x", pady=2)

        purple_button(
            lay_row,
            "Upload .lay File",
            self.upload_lay,
            width=18,
        ).pack(side="left", padx=(0, 6))

        self.lay_label = tk.Label(
            lay_row,
            text="No .lay selected",
            bg=UI_BG,
            fg=UI_TEXT_FG,
            anchor="w",
        )
        self.lay_label.pack(side="left", fill="x", expand=True)

        layout_row = tk.Frame(top, bg=UI_BG)
        layout_row.pack(fill="x", pady=2)

        purple_button(
            layout_row,
            "Upload Plate Layout",
            self.upload_layout,
            width=18,
        ).pack(side="left", padx=(0, 6))

        self.layout_label = tk.Label(
            layout_row,
            text="No layout selected",
            bg=UI_BG,
            fg=UI_TEXT_FG,
            anchor="w",
        )
        self.layout_label.pack(side="left", fill="x", expand=True)

        form = tk.LabelFrame(
            self.left_ui_frame,
            text="Required Transfer Information",
            bg=UI_BG,
            fg=UI_TEXT_FG,
            bd=3,
            relief="ridge",
        )
        form.pack(fill="x", padx=8, pady=4)

        form_inner = tk.Frame(form, bg=UI_BG)
        form_inner.pack(fill="x", padx=6, pady=6)

        purple_label(form_inner, "Transfer Type").grid(
            row=0, column=0, padx=4, pady=3
        )
        ttk.Combobox(
            form_inner,
            textvariable=self.transfer_type_var,
            values=TRANSFER_OPTIONS,
            state="readonly",
            width=12,
        ).grid(row=0, column=1, padx=4, pady=3)

        purple_label(form_inner, "Replicate").grid(
            row=0, column=2, padx=4, pady=3
        )
        ttk.Combobox(
            form_inner,
            textvariable=self.replicate_var,
            values=REPLICATE_OPTIONS,
            state="readonly",
            width=10,
        ).grid(row=0, column=3, padx=4, pady=3)

        purple_label(form_inner, "Volume uL").grid(
            row=0, column=4, padx=4, pady=3
        )
        tk.Entry(
            form_inner,
            textvariable=self.volume_var,
            width=9,
            bg="white",
            fg="black",
        ).grid(row=0, column=5, padx=4, pady=3)

        purple_label(form_inner, "Marker (optional)").grid(
            row=0, column=6, padx=4, pady=3
        )
        tk.Entry(
            form_inner,
            textvariable=self.marker_var,
            width=9,
            bg="white",
            fg="black",
        ).grid(row=0, column=7, padx=4, pady=3)

        lc_row = tk.Frame(form, bg=UI_BG)
        lc_row.pack(fill="x", padx=6, pady=(0, 6))

        purple_label(lc_row, "Liquid Class").pack(side="left", padx=(0, 6))

        self.liquid_combo = SearchableCombobox(
            lc_row,
            textvariable=self.liquid_class_var,
            completevalues=[],
            width=50,
        )
        self.liquid_combo.pack(side="left", padx=(0, 6))

        purple_button(
            lc_row,
            "Add LC",
            self.add_liquid_class_for_next_set,
            width=8,
        ).pack(side="left", padx=(0, 4))

        purple_button(
            lc_row,
            "Clear",
            self.clear_set_liquid_classes,
            width=7,
        ).pack(side="left")

        self.set_liquid_summary_label = tk.Label(
            self.left_ui_frame,
            textvariable=self.set_liquid_summary_var,
            bg=UI_BG,
            fg=UI_TEXT_FG,
            anchor="w",
        )
        self.set_liquid_summary_label.pack(fill="x", padx=8, pady=2)

        options_row = tk.Frame(self.left_ui_frame, bg=UI_BG)
        options_row.pack(fill="x", padx=8, pady=4)

        self.make_checkbutton(
            options_row,
            text="Tracking Enabled",
            variable=self.tracking_var,
        ).pack(side="left", padx=(0, 8))

        self.make_checkbutton(
            options_row,
            text="Variant",
            variable=self.variant_var,
            command=self.update_set_column_visibility,
        ).pack(side="left", padx=(0, 8))

        self.make_checkbutton(
            options_row,
            text="Tip Eject",
            variable=self.unique_tip_eject_var,
            command=self.update_unique_tip_eject_role_options,
        ).pack(side="left", padx=(0, 12))

        purple_label(options_row, "Report Format").pack(
            side="left", padx=5, pady=5
        )
        ttk.Combobox(
            options_row,
            textvariable=self.report_format_var,
            values=["docx", "pdf", "txt"],
            state="readonly",
            width=8,
        ).pack(side="left", padx=(0, 6), pady=5)

        list_frame = tk.LabelFrame(
            self.left_ui_frame,
            text="Extracted Method/Deck Sequences",
            bg=UI_BG,
            fg=UI_TEXT_FG,
            bd=3,
            relief="ridge",
        )
        list_frame.pack(fill="both", expand=True, padx=8, pady=4)

        self.sequence_canvas = tk.Canvas(
            list_frame,
            bg=UI_BG,
            highlightthickness=0,
            bd=0,
        )

        y_scrollbar = ttk.Scrollbar(
            list_frame,
            orient="vertical",
            command=self.sequence_canvas.yview,
        )
        x_scrollbar = ttk.Scrollbar(
            list_frame,
            orient="horizontal",
            command=self.sequence_canvas.xview,
        )

        self.sequence_frame = tk.Frame(
            self.sequence_canvas,
            bg=UI_BG,
        )
        self.sequence_canvas.create_window(
            (0, 0),
            window=self.sequence_frame,
            anchor="nw",
        )

        self.sequence_frame.bind(
            "<Configure>",
            lambda event: self.sequence_canvas.configure(
                scrollregion=self.sequence_canvas.bbox("all")
            ),
        )

        self.sequence_canvas.configure(
            yscrollcommand=y_scrollbar.set,
            xscrollcommand=x_scrollbar.set,
        )

        self.sequence_canvas.grid(row=0, column=0, sticky="nsew")
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar.grid(row=1, column=0, sticky="ew")
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        bottom = tk.Frame(self.left_ui_frame, bg=UI_BG)
        bottom.pack(fill="x", padx=8, pady=(4, 8))

        purple_button(
            bottom,
            "Run Sequence Analyzer",
            self.generate_plan_run_and_validate,
            width=24,
        ).pack(side="left", padx=(0, 6))

        purple_button(
            bottom,
            "Select Existing .trc",
            self.upload_trace,
            width=17,
        ).pack(side="left", padx=(0, 6))

        self.status = tk.Label(
            bottom,
            text="Upload .lay file to begin.",
            bg=UI_BG,
            fg=UI_TEXT_FG,
            anchor="w",
        )
        self.status.pack(side="left", fill="x", expand=True)

    def load_liquid_classes(self):
        self.liquid_classes = load_combined_liquid_classes(self.lay_path)
        self.liquid_combo.completevalues = self.liquid_classes
        self.liquid_combo["values"] = self.liquid_classes

        if self.liquid_classes:
            self.liquid_class_var.set(self.liquid_classes[0])
        else:
            self.liquid_class_var.set("")
            messagebox.showwarning(
                "Liquid Class Warning",
                "No liquid classes were found in the uploaded method files.",
            )

    def add_liquid_class_for_next_set(self):
        liquid_class = self.liquid_class_var.get().strip()

        if not liquid_class:
            messagebox.showerror(
                "Missing Liquid Class",
                "Select a liquid class first.",
            )
            return

        if liquid_class not in self.liquid_classes:
            messagebox.showerror(
                "Invalid Liquid Class",
                "Selected liquid class is not from the uploaded method.",
            )
            return

        self.set_liquid_classes.append(liquid_class)
        self.set_liquid_summary_var.set(
            "; ".join(
                f"Set {index} = {value}"
                for index, value in enumerate(self.set_liquid_classes, start=1)
            )
        )

    def clear_set_liquid_classes(self):
        self.set_liquid_classes = []
        self.set_liquid_summary_var.set(
            "No set liquid classes selected."
        )

    def _start_layout_autodetect(self) -> None:
        """Auto-detect marker and replicate without blocking the Tk UI.

        Marker and replicate are properties of the uploaded plate-layout file,
        not the Hamilton .lay file.  This method is called after either upload
        so the UI fills them as soon as both inputs are available, regardless
        of which file the user selected first.
        """
        layout_path = str(self.layout_path or "").strip()
        if not layout_path or not os.path.isfile(layout_path):
            return

        self._layout_detection_generation += 1
        generation = self._layout_detection_generation

        def worker():
            try:
                marker_matches = parse_plate_layout_file(layout_path, "")
                detected_marker = infer_marker_prefix(marker_matches, "")
                _prefix, templates = build_layout_templates(
                    marker_matches,
                    detected_marker,
                    "",
                )
                detected_replicate = ""
                if templates:
                    candidate = str(templates[0].get("replicate", "")).strip().lower()
                    if candidate in REPLICATE_OPTIONS:
                        detected_replicate = candidate
                source_count = len(templates)
                template_text = (
                    "one reusable plate template"
                    if source_count == 1
                    else f"{source_count} plate-specific layout templates"
                )
                result = (detected_marker, detected_replicate, template_text, "")
            except Exception as exc:
                result = ("", "", "", str(exc))

            def apply_result():
                # Ignore an older worker if the user selected another layout.
                if generation != self._layout_detection_generation:
                    return
                detected_marker, detected_replicate, template_text, error = result
                if error:
                    self.status.config(
                        text=f"Plate layout loaded; auto-detection warning: {error}"
                    )
                    return
                if detected_marker:
                    self.marker_var.set(detected_marker)
                if detected_replicate:
                    self.replicate_var.set(detected_replicate)
                self.status.config(
                    text=(
                        f"Plate layout loaded. Marker {detected_marker} auto-detected; "
                        f"replicate {detected_replicate or self.replicate_var.get()} auto-detected; "
                        f"{template_text}."
                    )
                )

            try:
                self.root.after(0, apply_result)
            except Exception:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def upload_lay(self):
        path = filedialog.askopenfilename(
            title="Select .lay File",
            filetypes=[
                ("LAY files", "*.lay"),
                ("All files", "*.*"),
            ],
        )

        if not path:
            return

        self.lay_path = path
        self.lay_label.config(text=path)

        try:
            self.sequence_names = sorted(
                parse_lay_file(path),
                key=str.lower,
            )
            self.sequence_counts = parse_lay_sequence_counts(path)

            # Keep LAY upload lightweight, matching the original Version 1
            # behavior.  Do NOT parse the full LAY metadata here: that extra
            # synchronous parse is what made the UI appear frozen immediately
            # after selecting the .lay file.  Only invalidate the lazy runtime
            # cache; it will be populated later if/when validation needs it.
            self._cached_lay_metadata_path = ""
            self._cached_lay_metadata_mtime = None
            self._cached_lay_metadata = {}

            if not self.sequence_names:
                raise ValueError(
                    "No valid sequences were found in the selected .lay file."
                )

            self.load_liquid_classes()
            self.render_sequences()

            self.status.config(
                text=f"Loaded {len(self.sequence_names)} sequence(s)."
            )

            # If the plate layout was selected first, immediately refresh the
            # Marker and Replicate UI fields now that the .lay is loaded.
            # Detection runs in the background, so this does not freeze upload.
            if self.layout_path:
                self._start_layout_autodetect()
        except Exception as exc:
            messagebox.showerror(
                "LAY Parse Error",
                str(exc),
            )

    def upload_layout(self):
        path = filedialog.askopenfilename(
            title="Select Plate Layout File",
            filetypes=[
                ("Plate layout files", "*.csv *.xls *.xlsx *.docx"),
                ("All files", "*.*"),
            ],
        )

        if not path:
            return

        self.layout_path = path
        self.layout_label.config(text=path)

        # Auto-detect Marker + Replicate immediately, but do the spreadsheet/
        # DOCX parsing off the Tk UI thread so selecting a layout never makes
        # the application look frozen.
        self.status.config(text="Plate layout loaded; detecting marker and replicate...")
        self._start_layout_autodetect()

    def upload_trace(self):
        path = filedialog.askopenfilename(
            title="Select Trace File",
            filetypes=[
                ("Trace files", "*.trc"),
                ("All files", "*.*"),
            ],
        )

        if path:
            self.trace_path = path
            self.status.config(text=f"Selected trace: {path}")

    def render_sequences(self):
        for widget in self.sequence_frame.winfo_children():
            widget.destroy()

        self.sequence_vars = []
        self.sequence_role_vars = {}
        self.sequence_role_widgets = {}
        self.sequence_channel_vars = {}
        self.sequence_hardware_vars = {}
        self.sequence_count_vars = {}
        self.sequence_manual_vars = {}
        self.sequence_control_vars = {}
        self.sequence_set_vars = {}
        self.sequence_set_widgets = {}

        headers = [
            ("Use", 0, 5),
            ("Control", 1, 7),
            ("Set", 2, 4),
            ("Sequence", 3, 48),
            ("Step Role", 4, 12),
            ("Manual", 5, 7),
            ("Channel Pattern", 6, 16),
            ("Hardware", 7, 10),
            ("Seq Count", 8, 9),
        ]

        for text, col, width in headers:
            tk.Label(
                self.sequence_frame,
                text=text,
                width=width,
                bg=UI_TEXT_BG,
                fg=UI_TEXT_FG,
                font=("Arial", 9, "bold"),
                relief="raised",
            ).grid(row=0, column=col, padx=3, pady=4)

        for row_index, sequence_name in enumerate(
            self.sequence_names,
            start=1,
        ):
            sequence_count = int(
                self.sequence_counts.get(sequence_name, 0) or 0
            )
            default_hardware = detect_hardware_mode(sequence_name)

            use_var = tk.IntVar(value=0)
            control_var = tk.IntVar(value=0)
            set_var = tk.StringVar(value="")
            role_var = tk.StringVar(value="aspirate")
            manual_var = tk.IntVar(value=0)
            channel_var = tk.StringVar(
                value=(
                    "1" * 96
                    if default_hardware == "head"
                    else "111111111111"
                )
            )
            hardware_var = tk.StringVar(value=default_hardware)
            count_var = tk.StringVar(value=str(sequence_count))

            self.sequence_vars.append((sequence_name, use_var))
            self.sequence_control_vars[sequence_name] = control_var
            self.sequence_set_vars[sequence_name] = set_var
            self.sequence_role_vars[sequence_name] = role_var
            self.sequence_manual_vars[sequence_name] = manual_var
            self.sequence_channel_vars[sequence_name] = channel_var
            self.sequence_hardware_vars[sequence_name] = hardware_var
            self.sequence_count_vars[sequence_name] = count_var

            self.make_checkbutton(
                self.sequence_frame,
                variable=use_var,
                command=self.update_set_column_visibility,
                borderless=True,
            ).grid(row=row_index, column=0, padx=3)

            self.make_checkbutton(
                self.sequence_frame,
                variable=control_var,
                borderless=True,
            ).grid(row=row_index, column=1, padx=3)

            set_entry = tk.Entry(
                self.sequence_frame,
                textvariable=set_var,
                width=4,
                bg="white",
                fg="black",
            )
            set_entry.grid(row=row_index, column=2, padx=3)
            self.sequence_set_widgets[sequence_name] = set_entry

            tk.Label(
                self.sequence_frame,
                text=sequence_name,
                width=48,
                anchor="w",
                bg=UI_BG,
                fg=UI_TEXT_FG,
            ).grid(row=row_index, column=3, padx=3, sticky="w")

            role_combo = ttk.Combobox(
                self.sequence_frame,
                textvariable=role_var,
                values=self.get_step_role_options(),
                width=12,
                state="readonly",
            )
            role_combo.grid(row=row_index, column=4, padx=3)
            self.sequence_role_widgets[sequence_name] = role_combo

            self.make_checkbutton(
                self.sequence_frame,
                variable=manual_var,
                borderless=True,
            ).grid(row=row_index, column=5, padx=3)

            tk.Entry(
                self.sequence_frame,
                textvariable=channel_var,
                width=16,
                bg="white",
                fg="black",
            ).grid(row=row_index, column=6, padx=3)

            # Important: this remains user-selectable. Unlike the old DP,
            # changing another field does not silently force it back.
            ttk.Combobox(
                self.sequence_frame,
                textvariable=hardware_var,
                values=["channels", "head"],
                width=10,
                state="readonly",
            ).grid(row=row_index, column=7, padx=3)

            tk.Label(
                self.sequence_frame,
                textvariable=count_var,
                width=9,
                bg=UI_BG,
                fg=UI_TEXT_FG,
            ).grid(row=row_index, column=8, padx=3)

        self.update_set_column_visibility()

    def get_step_role_options(self) -> List[str]:
        """
        Tip Eject only appears as a selectable Step Role when Unique Tip Eject
        is enabled. Otherwise the UI is identical to the previous behavior.
        """
        if bool(self.unique_tip_eject_var.get()):
            return list(ALL_ROLE_OPTIONS)
        return list(ROLE_OPTIONS)

    def update_unique_tip_eject_role_options(self):
        """
        Refresh every Step Role combobox when Unique Tip Eject is toggled.

        If the option is turned back off, any sequence that was assigned
        "tip eject" is reset to "aspirate" so a hidden/invalid role cannot
        remain selected. The normal automatic Waste eject will then be used.
        """
        allowed_roles = self.get_step_role_options()
        unique_enabled = bool(self.unique_tip_eject_var.get())

        for sequence_name, combo in self.sequence_role_widgets.items():
            combo["values"] = allowed_roles

            role_var = self.sequence_role_vars.get(sequence_name)
            if (
                role_var is not None
                and not unique_enabled
                and role_var.get().strip().lower() == TIP_EJECT_ROLE
            ):
                role_var.set("aspirate")

    def update_set_column_visibility(self):
        selected_count = sum(
            1
            for _, use_var in self.sequence_vars
            if use_var.get()
        )
        show_sets = (
            selected_count > 3
            or bool(self.variant_var.get())
        )

        for sequence_name, widget in self.sequence_set_widgets.items():
            if show_sets:
                widget.grid()
            else:
                widget.grid_remove()
                self.sequence_set_vars[sequence_name].set("")

    def get_cached_lay_metadata(self) -> Dict[str, Dict[str, Any]]:
        """Parse the selected LAY once and reuse it until the file changes."""
        path = os.path.abspath(self.lay_path) if self.lay_path else ""
        try:
            mtime = os.path.getmtime(path) if path else None
        except OSError:
            mtime = None

        if (
            path
            and path == self._cached_lay_metadata_path
            and mtime == self._cached_lay_metadata_mtime
            and self._cached_lay_metadata
        ):
            return self._cached_lay_metadata

        metadata = parse_lay_sequence_metadata(path) if path else {}
        self._cached_lay_metadata_path = path
        self._cached_lay_metadata_mtime = mtime
        self._cached_lay_metadata = metadata
        return metadata

    def collect_configs(self) -> Optional[List[SequenceStepConfig]]:
        selected_sequences = [
            sequence_name
            for sequence_name, use_var in self.sequence_vars
            if use_var.get()
        ]

        if not selected_sequences:
            return None

        transfer_type = self.transfer_type_var.get().strip() or "samples"
        replicate = self.replicate_var.get().strip().lower()
        # Do not parse the plate-layout file here.  Marker discovery is a
        # validation/report concern and can happen after Hamilton has launched.
        # Keeping this path lightweight prevents the Run button from appearing
        # frozen before Hamilton Run Control opens.
        marker = self.marker_var.get().strip().upper()

        variant_mode = bool(self.variant_var.get())

        use_set_order = (
            len(selected_sequences) > 3
            or variant_mode
        )

        selected_set_numbers = []

        if use_set_order:
            for sequence_name in selected_sequences:
                raw_set = self.sequence_set_vars[
                    sequence_name
                ].get().strip()

                try:
                    set_number = int(raw_set)
                except Exception:
                    messagebox.showerror(
                        "Invalid Set Number",
                        f"Set number is required and must be numeric for "
                        f"{sequence_name}.",
                    )
                    return None

                if set_number < 1:
                    messagebox.showerror(
                        "Invalid Set Number",
                        "Set numbers must be 1 or higher.",
                    )
                    return None

                selected_set_numbers.append(set_number)

        set_count = (
            max(selected_set_numbers)
            if selected_set_numbers
            else 1
        )

        volume_parts = [
            part.strip()
            for part in self.volume_var.get().split(",")
            if part.strip()
        ]

        if not volume_parts:
            messagebox.showerror(
                "Missing Volume",
                "Enter at least one volume.",
            )
            return None

        try:
            set_volumes = [float(part) for part in volume_parts]
        except Exception:
            messagebox.showerror(
                "Invalid Volume",
                "All volume values must be numeric.",
            )
            return None

        if any(
            value <= 0 or value > 1000
            for value in set_volumes
        ):
            messagebox.showerror(
                "Invalid Volume",
                "Volume must be greater than 0 and no more than 1000 uL.",
            )
            return None

        if len(set_volumes) == 1:
            set_volumes *= set_count
        elif len(set_volumes) != set_count:
            messagebox.showerror(
                "Set Volume Error",
                f"Enter one volume or exactly {set_count} comma-separated "
                "volumes.",
            )
            return None

        if (
            set_count > 1
            and len(self.set_liquid_classes) < set_count
        ):
            messagebox.showerror(
                "Missing Set Liquid Classes",
                f"Assign a liquid class for all {set_count} sets.",
            )
            return None

        configs = []

        for selected_index, sequence_name in enumerate(selected_sequences):
            role = self.sequence_role_vars[
                sequence_name
            ].get().strip()

            if (
                role == TIP_EJECT_ROLE
                and not bool(self.unique_tip_eject_var.get())
            ):
                messagebox.showerror(
                    "Unique Tip Eject Disabled",
                    "A sequence is assigned the Tip Eject role, but "
                    "Unique Tip Eject is not enabled.",
                )
                return None
            hardware_mode = self.sequence_hardware_vars[
                sequence_name
            ].get().strip().lower()
            channel_pattern = self.sequence_channel_vars[
                sequence_name
            ].get().strip()
            manual = bool(
                self.sequence_manual_vars[
                    sequence_name
                ].get()
            )
            control = bool(
                self.sequence_control_vars[
                    sequence_name
                ].get()
            )

            try:
                sequence_count = int(
                    self.sequence_count_vars[
                        sequence_name
                    ].get()
                )
            except Exception:
                sequence_count = int(
                    self.sequence_counts.get(
                        sequence_name,
                        0,
                    )
                )

            set_order = (
                int(
                    self.sequence_set_vars[
                        sequence_name
                    ].get()
                )
                if use_set_order
                else 1
            )

            volume_ul = set_volumes[set_order - 1]

            if role in ("aspirate", "dispense"):
                liquid_class = (
                    self.set_liquid_classes[
                        set_order - 1
                    ]
                    if self.set_liquid_classes
                    else self.liquid_class_var.get().strip()
                )

                if not liquid_class:
                    messagebox.showerror(
                        "Missing Liquid Class",
                        f"Liquid class is required for {sequence_name}.",
                    )
                    return None
            else:
                liquid_class = ""

            if hardware_mode == "head":
                channel_pattern = "1" * 96
            elif not re.fullmatch(
                r"(?:[01]{8}|[01]{12}|[01]{14}|[01]{16})",
                channel_pattern,
            ):
                messagebox.showerror(
                    "Invalid Channel Pattern",
                    f"{sequence_name}: use 8, 12, 14, or 16 digits of 0/1.",
                )
                return None

            configs.append(
                SequenceStepConfig(
                    sequence=sequence_name,
                    role=role,
                    liquid_class=liquid_class,
                    transfer_type=transfer_type,
                    replicate=replicate,
                    volume_ul=volume_ul,
                    marker=marker,
                    channel_pattern=channel_pattern,
                    hardware_mode=hardware_mode,
                    selected_index=selected_index,
                    sequence_count=sequence_count,
                    manual=manual,
                    set_order=set_order,
                    control=control,
                )
            )

        if bool(self.unique_tip_eject_var.get()):
            grouped_configs = defaultdict(list)
            for config in configs:
                grouped_configs[int(config.set_order or 1)].append(config)

            missing_eject_sets = [
                set_order
                for set_order, items in sorted(grouped_configs.items())
                if not any(item.role == TIP_EJECT_ROLE for item in items)
            ]
            if missing_eject_sets:
                messagebox.showerror(
                    "Tip Eject Sequence Required",
                    "Tip Eject is checked, so each active set must include a "
                    "sequence assigned the Tip Eject Step Role. Missing set(s): "
                    + ", ".join(str(value) for value in missing_eject_sets),
                )
                return None

        validation_error = validate_step_sets(
            configs,
            variant_mode,
            use_set_order,
        )

        if validation_error:
            messagebox.showerror(
                "Invalid Step Sets",
                validation_error,
            )
            return None

        # Keep Run-button collection lightweight. Full LAY metadata parsing can
        # be expensive on large VENUS layouts, so do not perform it here on the
        # Tkinter UI thread. Tip/labware validation is completed after Hamilton
        # has launched (or before launch only for 96-head methods that truly need
        # LAY resource metadata to execute).
        return configs

    def generate_plan_run_and_validate(self):
        if not self.lay_path:
            messagebox.showerror(
                "Missing LAY",
                "Upload a .lay file first.",
            )
            return

        if not self.layout_path:
            messagebox.showerror(
                "Missing Plate Layout",
                "Upload a plate layout file first.",
            )
            return

        configs = self.collect_configs()

        if not configs:
            messagebox.showerror(
                "Missing Sequences",
                "Select sequences first.",
            )
            return

        issues = validate_step_configs(configs)

        # Do not parse the full .lay before opening Hamilton for ordinary
        # channel methods. That parse was the main source of the apparent
        # "freeze" between clicking Run and seeing Run Control. A 96-head step
        # needs labware-resource metadata for execution, so only that case pays
        # the parsing cost up front.
        needs_head_metadata = any(
            str(config.hardware_mode or "channels").lower() == "head"
            for config in configs
        )
        lay_metadata = self.get_cached_lay_metadata() if needs_head_metadata else {}

        critical = [
            issue
            for issue in issues
            if issue.severity == "critical"
        ]

        if critical:
            messagebox.showerror(
                "Configuration Error",
                "\n".join(
                    issue.message
                    for issue in critical[:10]
                ),
            )
            return

        today_folder = datetime.now().strftime("%d%b%Y")
        output_dir = os.path.join(
            DEFAULT_RESULTS_DIR,
            today_folder,
        )
        os.makedirs(output_dir, exist_ok=True)

        marker = configs[0].marker
        marker_matches = {}
        layout_findings = []
        trace_findings = []
        limitations = []
        pattern = None

        # Build only what Hamilton needs before launch.  Plate-layout parsing,
        # deterministic pair generation, and layout validation are intentionally
        # deferred until after the simulation has started/completed.  This keeps
        # the Run button responsive and gets Hamilton Run Control on screen first.
        execution_plan = build_execution_plan(
            configs=configs,
            lay_metadata=lay_metadata,
            variant_mode=bool(self.variant_var.get()),
        )

        # Generate the review script after Hamilton has run.  It is a report
        # artifact, not a launch prerequisite, so file I/O here only delays the
        # Run Control window.
        review_script = ""

        self.status.config(text="Launching Hamilton Run Control...")
        # Process pending paint/events now so the UI visibly responds before the
        # blocking PyHamilton simulation call begins.
        self.root.update()

        manually_selected_trace = self.trace_path

        runtime_trace, runtime_issues = (
            run_pyhamilton_simulation(
                self.lay_path,
                configs,
                execution_plan,
            )
        )
        issues.extend(runtime_issues)

        review_script = generate_pyhamilton_review_script(
            configs,
            execution_plan,
            output_dir,
        )

        if runtime_trace:
            self.trace_path = runtime_trace
        elif manually_selected_trace:
            self.trace_path = manually_selected_trace
            issues.append(
                ValidationIssue(
                    "major",
                    "trace",
                    "No new simulation trace was found; using the manually "
                    "selected trace as fallback.",
                    {
                        "uploaded_trace": manually_selected_trace
                    },
                )
            )
        else:
            self.trace_path = ""

        # Full LAY metadata and tip-volume validation happen only after Run
        # Control has launched/completed for channel methods. This preserves the
        # validation while keeping the launch path fast.
        if not lay_metadata:
            lay_metadata = self.get_cached_lay_metadata()
        issues.extend(validate_tip_volume_rules(configs, lay_metadata))

        # Now perform the plate-layout analysis.  It is required for the
        # validation report, but not for opening/running Hamilton itself.
        try:
            marker_matches = parse_plate_layout_file(
                self.layout_path,
                marker,
            )

            pattern = analyze_plate_layout_pattern(
                marker_matches,
                configured_marker=marker,
                configured_replicate=configs[0].replicate,
            )

            marker = pattern.marker_prefix
            if marker and not self.marker_var.get().strip():
                self.marker_var.set(marker)

            for config in configs:
                if config.transfer_type == "samples":
                    config.marker = marker

        except Exception as exc:
            issues.append(
                ValidationIssue(
                    "critical",
                    "layout",
                    f"Could not analyze plate layout: {exc}",
                )
            )

        pairs = generate_deterministic_transfer_pairs(
            configs,
            marker_matches,
        )

        if marker_matches and pattern is not None:
            layout_findings, layout_issues = validate_layout_matches(
                marker_matches,
                pairs,
                marker,
                configs[0].replicate,
            )
            issues.extend(layout_issues)
        else:
            limitations.append(
                "Plate-layout pattern validation was not completed."
            )

        labware_summary = build_labware_validation_summary(
            marker_matches,
            trace_findings,
            configs,
            lay_metadata,
        )
        trace_findings.extend(labware_summary)

        for item in labware_summary:
            if item.get("status") == "fail":
                issues.append(
                    ValidationIssue(
                        "critical",
                        "URGENT_LABWARE_MISMATCH",
                        item.get("message", "Wrong labware detected."),
                        item,
                    )
                )

        if self.trace_path:
            try:
                aspirate_steps, dispense_steps, ordered_steps = (
                    parse_trc_file(
                        self.trace_path
                    )
                )

                trace_text = read_text_file_best_effort(
                    self.trace_path
                )

                runtime_findings, runtime_step_issues = (
                    validate_expected_runtime_steps(
                        execution_plan,
                        ordered_steps,
                        trace_text,
                    )
                )

                trace_findings.extend(runtime_findings)
                issues.extend(runtime_step_issues)

                ok, cross_match = cross_match_positions(
                    aspirate_steps,
                    dispense_steps,
                    configs[0].replicate,
                    ordered_steps=ordered_steps,
                )

                trace_findings.append({
                    "type": "trace_summary",
                    "trace_file": self.trace_path,
                    "ordered_steps": ordered_steps,
                    "aspirate_step_count": len(aspirate_steps),
                    "dispense_step_count": len(dispense_steps),
                    "cross_match_status": "pass" if ok else "fail",
                    "cross_match": str(cross_match),
                })

                orientation_findings, orientation_issues = (
                    analyze_orientation_and_tip_usage(
                        configs,
                        execution_plan,
                        marker_matches,
                        aspirate_steps,
                        dispense_steps,
                        cross_match,
                    )
                )
                trace_findings.extend(
                    orientation_findings
                )
                issues.extend(orientation_issues)

                if ok and isinstance(cross_match, dict):
                    if marker_matches and pattern is not None:
                        t_findings, t_issues = (
                            validate_trace_against_layout(
                                cross_match,
                                marker_matches,
                                marker,
                                configs[0].replicate,
                            )
                        )
                        trace_findings.extend(t_findings)
                        issues.extend(t_issues)

                    if self.tracking_var.get():
                        tracking_findings, tracking_issues = (
                            validate_tracking_source_order(
                                cross_match,
                                marker,
                            )
                        )
                        trace_findings.extend(
                            tracking_findings
                        )
                        issues.extend(
                            tracking_issues
                        )
                else:
                    # If VENUS already stopped during execution, a missing
                    # aspirate/dispense cross-match is a consequence of that
                    # runtime failure, not a second independent critical defect.
                    runtime_failed = any(
                        f.get("type") == "expected_runtime_step_check"
                        and f.get("status") == "fail"
                        for f in runtime_findings
                    )
                    if runtime_failed and not aspirate_steps and not dispense_steps:
                        limitations.append(
                            "Trace plate-layout cross-match was not evaluated because "
                            "the runtime stopped before any aspirate/dispense step completed."
                        )
                    else:
                        issues.append(
                            ValidationIssue(
                                "critical",
                                "trace",
                                f"Trace cross-match failed: {cross_match}",
                            )
                        )

            except Exception as exc:
                issues.append(
                    ValidationIssue(
                        "critical",
                        "trace",
                        f"Trace analysis failed: "
                        f"{type(exc).__name__}: {exc}",
                    )
                )
        else:
            limitations.append(
                "No trace file was selected or generated; "
                "trace validation was not completed."
            )

        output = create_output_model(
            lay_file=self.lay_path,
            layout_file=self.layout_path,
            trace_file=self.trace_path,
            configs=configs,
            pairs=pairs,
            execution_plan=execution_plan,
            marker_matches=marker_matches,
            layout_findings=layout_findings,
            trace_findings=trace_findings,
            issues=issues,
            limitations=limitations,
            review_script_path=review_script,
        )

        json_path = os.path.join(
            output_dir,
            f"DevPal_Lite_Output_Model_{datetime.now().strftime('%H%M%S')}.json",
        )

        Path(json_path).write_text(
            json.dumps(
                asdict(output),
                indent=2,
                default=str,
            ),
            encoding="utf-8",
        )

        report_format = self.report_format_var.get()

        if report_format == "pdf":
            report_path = write_pdf_report(
                output,
                output_dir,
            )
        elif report_format == "txt":
            report_path = write_text_report(
                output,
                output_dir,
            )
        else:
            report_path = write_docx_report(
                output,
                output_dir,
            )

        output.report_path = report_path

        Path(json_path).write_text(
            json.dumps(
                asdict(output),
                indent=2,
                default=str,
            ),
            encoding="utf-8",
        )

        issue_count = sum(
            1
            for issue in issues
            if issue.severity in (
                "critical",
                "major",
            )
        )

        messagebox.showinfo(
            "DevPal Lite Complete",
            (
                "Validation complete.\n\n"
                f"Major/Critical issues: {issue_count}\n"
                f"Report: {report_path}"
            ),
        )

        self.status.config(
            text=f"Complete. Report saved: {report_path}"
        )

    def run(self):
        self.root.mainloop()


def main():
    app = DevPalLiteApp()
    app.run()


if __name__ == "__main__":
    main()
