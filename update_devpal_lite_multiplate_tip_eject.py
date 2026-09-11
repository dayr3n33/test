from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path


def replace_top_level_function(source: str, name: str, replacement: str) -> str:
    pattern = re.compile(rf"(?ms)^def {re.escape(name)}\s*\(.*?(?=^def |^class |\Z)")
    match = pattern.search(source)
    if not match:
        raise RuntimeError(f"Could not find top-level function: {name}")
    return source[:match.start()] + replacement.strip() + "\n\n" + source[match.end():]


def insert_after_top_level_function(source: str, name: str, insertion: str) -> str:
    pattern = re.compile(rf"(?ms)^def {re.escape(name)}\s*\(.*?(?=^def |^class |\Z)")
    match = pattern.search(source)
    if not match:
        raise RuntimeError(f"Could not find top-level function: {name}")
    return source[:match.end()] + "\n\n" + insertion.strip() + "\n\n" + source[match.end():]


def replace_once(source: str, old: str, new: str, description: str) -> str:
    if old not in source:
        raise RuntimeError(f"Could not find expected code for: {description}")
    return source.replace(old, new, 1)


NEW_VALIDATE_STEP_SETS = r'''
def validate_step_sets(
    configs: List[SequenceStepConfig],
    variant_mode: bool,
    use_set_order: bool,
    tip_eject_enabled: bool = False,
) -> str:
    if not configs:
        return "No sequence configurations were selected."

    grouped = defaultdict(list)
    for config in configs:
        grouped[int(config.set_order)].append(config)

    allowed_roles = set(ROLE_OPTIONS)
    if tip_eject_enabled:
        allowed_roles.add("tip eject")

    for config in configs:
        if config.role not in allowed_roles:
            return f"Invalid step role for {config.sequence}: {config.role}"

    required_roles = list(ROLE_OPTIONS)
    if tip_eject_enabled:
        required_roles.append("tip eject")

    if variant_mode:
        for set_order, items in sorted(grouped.items()):
            roles = [item.role for item in items]
            for required_role in required_roles:
                if required_role not in roles:
                    return f"Set {set_order} must contain at least one {required_role} sequence."
        return ""

    for set_order, items in sorted(grouped.items()):
        roles = [item.role for item in items]
        for required_role in required_roles:
            if roles.count(required_role) != 1:
                if tip_eject_enabled:
                    return (
                        f"Set {set_order} must contain exactly one tip pick up, one aspirate, "
                        f"one dispense, and one tip eject sequence."
                    )
                return (
                    f"Set {set_order} must contain exactly one tip pick up, "
                    f"one aspirate, and one dispense sequence."
                )

    return ""
'''

NEW_VALIDATE_STEP_CONFIGS = r'''
def validate_step_configs(
    configs: List[SequenceStepConfig],
    tip_eject_enabled: bool = False,
) -> List[ValidationIssue]:
    issues = []
    roles = {config.role for config in configs}

    required_roles = list(ROLE_OPTIONS)
    if tip_eject_enabled:
        required_roles.append("tip eject")

    allowed_roles = set(required_roles)

    for required_role in required_roles:
        if required_role not in roles:
            issues.append(ValidationIssue("critical", "configuration", f"Missing required step role: {required_role}"))

    for config in configs:
        if config.role not in allowed_roles:
            issues.append(ValidationIssue("critical", "configuration", f"Invalid step role for {config.sequence}: {config.role}"))

        if config.role in ("aspirate", "dispense") and not config.liquid_class:
            issues.append(ValidationIssue("critical", "configuration", f"Liquid class is required for {config.role}: {config.sequence}"))

        if config.volume_ul <= 0 or config.volume_ul > 1000:
            issues.append(ValidationIssue("critical", "configuration", f"Volume must be between 1 and 1000 uL: {config.sequence}"))

        if config.transfer_type not in TRANSFER_OPTIONS:
            issues.append(ValidationIssue("critical", "configuration", f"Invalid transfer type: {config.transfer_type}"))

        if config.replicate not in REPLICATE_OPTIONS:
            issues.append(ValidationIssue("critical", "configuration", f"Invalid replicate value: {config.replicate}"))

        if config.transfer_type == "samples" and not config.marker:
            issues.append(ValidationIssue("critical", "configuration", f"Marker is required for samples: {config.sequence}"))

    return issues
'''

NEW_GET_REQUIRED_TRANSFER_ITERATIONS = r'''
def get_required_transfer_iterations(configs: List[SequenceStepConfig]) -> int:
    control_configs = [
        config for config in configs
        if bool(getattr(config, "control", False))
        and not bool(getattr(config, "manual", False))
        and config.role in ("tip pick up", "aspirate", "dispense", "tip eject")
    ]

    if control_configs:
        return max(get_sequence_iterations(config) for config in control_configs)

    for role in ("dispense", "aspirate", "tip eject"):
        role_configs = [
            config for config in configs
            if config.role == role and not bool(getattr(config, "manual", False))
        ]
        if role_configs:
            return max(get_sequence_iterations(config) for config in role_configs)

    return 1
'''

NEW_BUILD_EXECUTION_PLAN = r'''
def build_execution_plan(
    configs: List[SequenceStepConfig],
    lay_metadata: Optional[Dict[str, Dict[str, Any]]] = None,
    variant_mode: bool = False,
) -> List[Dict[str, Any]]:
    lay_metadata = lay_metadata or {}
    plan = []
    order = 1

    grouped = defaultdict(list)
    for config in configs:
        grouped[int(getattr(config, "set_order", 1) or 1)].append(config)

    def should_loop_by_replicate(set_configs: List[SequenceStepConfig]) -> bool:
        aspirates = [c for c in set_configs if c.role == "aspirate"]
        dispenses = [c for c in set_configs if c.role == "dispense"]
        if not aspirates or not dispenses:
            return False
        return (
            all(bool(getattr(c, "manual", False)) for c in aspirates)
            and all(bool(getattr(c, "manual", False)) for c in dispenses)
        )

    def append_plan_step(config, set_order, rep_index, transfer_iteration):
        nonlocal order
        batch_size = get_sequence_batch_size(config)
        action = {
            "tip pick up": "tip_pick",
            "aspirate": "aspirate",
            "dispense": "dispense",
            "tip eject": "tip_eject",
        }.get(config.role, config.role)

        config_iterations = get_sequence_iterations(config)
        should_increment = (
            transfer_iteration > 1
            and transfer_iteration <= config_iterations
            and not bool(getattr(config, "manual", False))
            and config.role in ("aspirate", "dispense", "tip eject")
        )

        seq_meta = get_sequence_labware_metadata(
            config.sequence,
            lay_metadata,
            getattr(config, "sequence_count", 0),
        )

        plan.append({
            "order": order,
            "set_order": set_order,
            "action": action,
            "sequence": config.sequence,
            "hardware_mode": config.hardware_mode,
            "labware_name": seq_meta.get("labware_name", ""),
            "labware_format": seq_meta.get("labware_format", ""),
            "labware_positions": seq_meta.get("positions", []),
            "replicate_index": rep_index,
            "replicate_report_only": not should_loop_by_replicate(grouped[set_order]),
            "transfer_iteration": transfer_iteration,
            "manual": bool(getattr(config, "manual", False)),
            "control": bool(getattr(config, "control", False)),
            "autoincrement": should_increment,
            "increment": batch_size if should_increment else 0,
            "channel_pattern": config.channel_pattern,
            "volume_ul": config.volume_ul if config.role in ("aspirate", "dispense") else "",
            "liquid_class": config.liquid_class if config.role in ("aspirate", "dispense") else "",
            "transfer_type": config.transfer_type,
            "marker": config.marker,
            "sequence_count": config.sequence_count,
            "batch_size": batch_size,
        })
        order += 1

    def append_default_tip_eject(tip, set_order, rep_index, transfer_iteration):
        nonlocal order
        plan.append({
            "order": order,
            "set_order": set_order,
            "action": "tip_eject",
            "sequence": "Waste",
            "hardware_mode": tip.hardware_mode,
            "labware_name": "Waste",
            "labware_format": "waste",
            "labware_positions": [],
            "replicate_index": rep_index,
            "replicate_report_only": not should_loop_by_replicate(grouped[set_order]),
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
        })
        order += 1

    for set_order in sorted(grouped):
        set_configs = sorted(grouped[set_order], key=lambda c: c.selected_index)
        replicate = set_configs[0].replicate if set_configs else "single"
        configured_replicate_count = REPLICATE_MAP.get(str(replicate).lower(), 1)
        loop_by_replicate = should_loop_by_replicate(set_configs)
        runtime_replicate_count = configured_replicate_count if loop_by_replicate else 1

        tip_steps = [c for c in set_configs if c.role == "tip pick up"]
        asp_steps = [c for c in set_configs if c.role == "aspirate"]
        dsp_steps = [c for c in set_configs if c.role == "dispense"]
        eject_steps = [c for c in set_configs if c.role == "tip eject"]
        use_explicit_tip_eject = bool(eject_steps)

        transfer_iterations = get_required_transfer_iterations(set_configs)

        for rep_index in range(1, runtime_replicate_count + 1):
            report_replicate_index = rep_index if loop_by_replicate else configured_replicate_count

            for transfer_iteration in range(1, transfer_iterations + 1):
                if variant_mode:
                    if not tip_steps or not asp_steps or not dsp_steps:
                        continue

                    for dispense_index, dispense_config in enumerate(dsp_steps):
                        tip_config = tip_steps[dispense_index % len(tip_steps)]
                        aspirate_config = asp_steps[dispense_index % len(asp_steps)]

                        append_plan_step(tip_config, set_order, report_replicate_index, transfer_iteration)
                        append_plan_step(aspirate_config, set_order, report_replicate_index, transfer_iteration)
                        append_plan_step(dispense_config, set_order, report_replicate_index, transfer_iteration)

                        if use_explicit_tip_eject:
                            eject_config = eject_steps[dispense_index % len(eject_steps)]
                            append_plan_step(eject_config, set_order, report_replicate_index, transfer_iteration)
                        else:
                            append_default_tip_eject(tip_config, set_order, report_replicate_index, transfer_iteration)
                    continue

                max_len = max(
                    len(tip_steps),
                    len(asp_steps),
                    len(dsp_steps),
                    len(eject_steps) if use_explicit_tip_eject else 0,
                    1,
                )

                for idx in range(max_len):
                    tip_config = tip_steps[idx] if idx < len(tip_steps) else tip_steps[0] if tip_steps else None
                    aspirate_config = asp_steps[idx] if idx < len(asp_steps) else asp_steps[0] if asp_steps else None
                    dispense_config = dsp_steps[idx] if idx < len(dsp_steps) else dsp_steps[0] if dsp_steps else None

                    for config in (tip_config, aspirate_config, dispense_config):
                        if config is not None:
                            append_plan_step(config, set_order, report_replicate_index, transfer_iteration)

                    if use_explicit_tip_eject:
                        eject_config = eject_steps[idx] if idx < len(eject_steps) else eject_steps[0]
                        append_plan_step(eject_config, set_order, report_replicate_index, transfer_iteration)
                    elif tip_config is not None:
                        append_default_tip_eject(tip_config, set_order, report_replicate_index, transfer_iteration)

    return plan
'''

NEW_GENERATE_PAIRS = r'''
def generate_deterministic_transfer_pairs(
    configs: List[SequenceStepConfig],
    layout_matches: Dict[str, List[Dict[str, Any]]],
) -> List[TransferPair]:
    replicate = configs[0].replicate if configs else "single"
    replicate_count = REPLICATE_MAP.get(str(replicate).lower(), 1)

    tip_steps = [c for c in configs if c.role == "tip pick up"]
    asp_steps = [c for c in configs if c.role == "aspirate"]
    dsp_steps = [c for c in configs if c.role == "dispense"]

    if not dsp_steps:
        return []

    marker = configs[0].marker.upper().strip() if configs and configs[0].marker else ""
    pairs = []

    max_dispense_positions = max(int(dsp.sequence_count or 0) for dsp in dsp_steps)
    if max_dispense_positions <= 0:
        max_dispense_positions = 1

    source_count = max(
        (max_dispense_positions + max(replicate_count, 1) - 1) // max(replicate_count, 1),
        1,
    )

    for source_num in range(1, source_count + 1):
        asp = asp_steps[(source_num - 1) % len(asp_steps)] if asp_steps else None
        dsp = dsp_steps[(source_num - 1) % len(dsp_steps)] if dsp_steps else None
        tip = tip_steps[(source_num - 1) % len(tip_steps)] if tip_steps else None

        layout_marker, plate_index, _ = resolve_layout_marker_for_sample(
            source_num,
            layout_matches,
            marker,
        )
        displayed_marker = f"{marker}{source_num}" if marker else ""
        positions = get_layout_marker_positions(
            layout_matches,
            layout_marker,
            plate_index=plate_index,
        )

        for rep_index in range(1, replicate_count + 1):
            dest_position = ""
            if positions:
                dest_position = positions[min(rep_index - 1, len(positions) - 1)]

            pairs.append(TransferPair(
                pair_id=f"P{plate_index}-S{source_num}-R{rep_index}",
                tip_sequence=tip.sequence if tip else "",
                source_sequence=asp.sequence if asp else "",
                destination_sequence=dsp.sequence if dsp else "",
                source_position=str(source_num),
                destination_position=dest_position,
                source_sample_number=source_num,
                destination_sample_number=extract_position_number(dest_position),
                replicate_index=rep_index,
                volume_ul=asp.volume_ul if asp else dsp.volume_ul if dsp else 0.0,
                liquid_class=asp.liquid_class if asp else dsp.liquid_class if dsp else "",
                transfer_type=asp.transfer_type if asp else dsp.transfer_type if dsp else "",
                marker_label=displayed_marker,
            ))

    return pairs
'''

NEW_PARSE_DATAFRAME = r'''
def parse_dataframe_plate_layout(
    df,
    marker_pattern,
    plate_layout,
    plate_index: int = 1,
    plate_name: str = "",
):
    if df is None or df.empty:
        return

    header_row_idx = 0
    row_label_col_idx = 0

    for row_idx in range(1, df.shape[0]):
        plate_row_raw = df.iat[row_idx, row_label_col_idx]
        if pd.isna(plate_row_raw):
            continue

        plate_row = str(plate_row_raw).strip().upper()
        if not re.match(r"^[A-P]$", plate_row):
            continue

        for col_idx in range(1, df.shape[1]):
            plate_col_raw = df.iat[header_row_idx, col_idx]
            if pd.isna(plate_col_raw):
                continue

            plate_col_match = re.search(r"\d+", str(plate_col_raw).strip())
            if not plate_col_match:
                continue

            plate_col = int(plate_col_match.group(0))
            cell_value = df.iat[row_idx, col_idx]
            if pd.isna(cell_value):
                continue

            cell_text = str(cell_value).strip().upper()
            if not cell_text:
                continue

            matches = marker_pattern.findall(cell_text)
            if not matches:
                continue

            dest_position = f"{plate_row}{plate_col}"
            for found_marker in matches:
                plate_layout[found_marker.upper()].append({
                    "dest_value": dest_position,
                    "cell_value": cell_text,
                    "plate_row": plate_row,
                    "plate_column": plate_col,
                    "row_index": row_idx + 1,
                    "column_index": col_idx + 1,
                    "excel_position": f"{excel_col_name(col_idx)}{row_idx + 1}",
                    "column_name": excel_col_name(col_idx),
                    "plate_index": int(plate_index),
                    "plate_name": str(plate_name or f"Plate {plate_index}"),
                })
'''

NEW_PARSE_LAYOUT = r'''
def parse_plate_layout_file(layout_path: str, marker: str) -> Dict[str, List[Dict[str, Any]]]:
    if pd is None:
        raise RuntimeError("pandas is required for layout parsing.")

    marker = marker.strip().upper()
    if not marker:
        return {}

    marker_pattern = re.compile(rf"\b{re.escape(marker)}\d+\b", re.IGNORECASE)
    plate_layout = defaultdict(list)
    ext = os.path.splitext(layout_path)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(layout_path, header=None, dtype=str)
        parse_dataframe_plate_layout(df, marker_pattern, plate_layout, 1, "Plate 1")

    elif ext in (".xls", ".xlsx"):
        sheets = pd.read_excel(layout_path, sheet_name=None, header=None, dtype=str)
        for plate_index, (sheet_name, df) in enumerate(sheets.items(), start=1):
            parse_dataframe_plate_layout(
                df,
                marker_pattern,
                plate_layout,
                plate_index=plate_index,
                plate_name=str(sheet_name),
            )

    elif ext == ".docx":
        if Document is None:
            raise RuntimeError("python-docx is required for DOCX layout parsing.")

        doc = Document(layout_path)
        for plate_index, table in enumerate(doc.tables, start=1):
            if not table.rows:
                continue

            header_cells = table.rows[0].cells
            for row_idx in range(1, len(table.rows)):
                row = table.rows[row_idx]
                plate_row = str(row.cells[0].text).strip().upper()
                if not re.match(r"^[A-P]$", plate_row):
                    continue

                for col_idx in range(1, len(row.cells)):
                    if col_idx >= len(header_cells):
                        continue

                    plate_col_match = re.search(r"\d+", str(header_cells[col_idx].text).strip())
                    if not plate_col_match:
                        continue

                    plate_col = int(plate_col_match.group(0))
                    cell_text = str(row.cells[col_idx].text).strip().upper()
                    matches = marker_pattern.findall(cell_text)
                    if not matches:
                        continue

                    dest_position = f"{plate_row}{plate_col}"
                    for found_marker in matches:
                        plate_layout[found_marker.upper()].append({
                            "dest_value": dest_position,
                            "cell_value": cell_text,
                            "plate_row": plate_row,
                            "plate_column": plate_col,
                            "row_index": row_idx + 1,
                            "column_index": col_idx + 1,
                            "excel_position": f"{excel_col_name(col_idx)}{row_idx + 1}",
                            "column_name": excel_col_name(col_idx),
                            "plate_index": int(plate_index),
                            "plate_name": f"Table {plate_index}",
                        })
    else:
        raise ValueError("Unsupported plate layout file type. Use CSV, XLS, XLSX, or DOCX.")

    return dict(plate_layout)
'''

NEW_LAYOUT_HELPERS = r'''
def get_layout_plate_indices(layout_matches: Dict[str, List[Dict[str, Any]]]) -> List[int]:
    plate_indices = set()
    for entries in layout_matches.values():
        for entry in entries:
            try:
                plate_indices.add(int(entry.get("plate_index", 1) or 1))
            except Exception:
                plate_indices.add(1)
    return sorted(plate_indices) if plate_indices else [1]


def get_plate_marker_numbers(
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker: str,
    plate_index: int,
) -> List[int]:
    marker = str(marker or "").upper().strip()
    numbers = set()

    for marker_key, entries in layout_matches.items():
        match = re.match(rf"^{re.escape(marker)}(\d+)$", str(marker_key).upper().strip())
        if not match:
            continue

        for entry in entries:
            try:
                entry_plate = int(entry.get("plate_index", 1) or 1)
            except Exception:
                entry_plate = 1
            if entry_plate == int(plate_index):
                numbers.add(int(match.group(1)))
                break

    return sorted(numbers)


def resolve_layout_marker_for_sample(
    source_number: int,
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker: str,
) -> Tuple[str, int, int]:
    marker = str(marker or "").upper().strip()
    source_number = max(int(source_number or 1), 1)
    plate_indices = get_layout_plate_indices(layout_matches)

    if not layout_matches:
        return f"{marker}{source_number}", 1, source_number

    if len(plate_indices) == 1:
        template_plate = plate_indices[0]
        marker_numbers = get_plate_marker_numbers(layout_matches, marker, template_plate)
        if not marker_numbers:
            return f"{marker}{source_number}", 1, source_number

        samples_per_plate = len(marker_numbers)
        plate_index = ((source_number - 1) // samples_per_plate) + 1
        local_index = (source_number - 1) % samples_per_plate
        marker_number = marker_numbers[local_index]
        return f"{marker}{marker_number}", plate_index, local_index + 1

    remaining = source_number
    for plate_index in plate_indices:
        marker_numbers = get_plate_marker_numbers(layout_matches, marker, plate_index)
        if not marker_numbers:
            continue

        count = len(marker_numbers)
        if remaining <= count:
            marker_number = source_number if source_number in marker_numbers else marker_numbers[remaining - 1]
            return f"{marker}{marker_number}", plate_index, remaining
        remaining -= count

    last_plate = plate_indices[-1]
    last_numbers = get_plate_marker_numbers(layout_matches, marker, last_plate)
    if not last_numbers:
        return f"{marker}{source_number}", last_plate, source_number

    count = len(last_numbers)
    extra_plate_offset = (remaining - 1) // count
    local_index = (remaining - 1) % count
    logical_plate_index = last_plate + extra_plate_offset + 1
    return f"{marker}{last_numbers[local_index]}", logical_plate_index, local_index + 1
'''

NEW_GET_LAYOUT_MARKER_POSITIONS = r'''
def get_layout_marker_positions(
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker_label: str,
    plate_index: Optional[int] = None,
) -> List[str]:
    entries = layout_matches.get(str(marker_label).upper().strip(), [])

    if plate_index is not None and entries:
        available_plates = sorted({int(entry.get("plate_index", 1) or 1) for entry in entries})
        all_layout_plates = get_layout_plate_indices(layout_matches)

        if int(plate_index) in available_plates:
            entries = [entry for entry in entries if int(entry.get("plate_index", 1) or 1) == int(plate_index)]
        elif len(all_layout_plates) == 1:
            template_plate = all_layout_plates[0]
            entries = [entry for entry in entries if int(entry.get("plate_index", 1) or 1) == template_plate]
        elif available_plates:
            fallback_plate = available_plates[-1]
            entries = [entry for entry in entries if int(entry.get("plate_index", 1) or 1) == fallback_plate]

    return [
        normalize_position_text(entry.get("dest_value", ""))
        for entry in entries
        if entry.get("dest_value") not in (None, "", "nan")
    ]
'''

NEW_VALIDATE_LAYOUT = r'''
def validate_layout_matches(
    layout_matches: Dict[str, List[Dict[str, Any]]],
    pairs: List[TransferPair],
    marker: str,
) -> Tuple[List[Dict[str, Any]], List[ValidationIssue]]:
    findings = []
    issues = []
    marker = str(marker or "").upper().strip()
    all_layout_positions = []
    seen_positions = defaultdict(list)

    for marker_label, entries in layout_matches.items():
        for entry in entries:
            pos = normalize_position_text(entry.get("dest_value", ""))
            if not pos:
                continue
            plate_index = int(entry.get("plate_index", 1) or 1)
            all_layout_positions.append((plate_index, pos))
            seen_positions[(plate_index, pos)].append(marker_label)

    duplicate_layout_positions = {
        key: labels for key, labels in seen_positions.items() if len(labels) > 1
    }

    if duplicate_layout_positions:
        add_issue(
            issues,
            "major",
            "plate layout",
            "Duplicate destination wells were found within the same plate layout.",
            {
                "duplicate_destination_wells": [
                    f"Plate {plate_index}:{position}"
                    for plate_index, position in sorted(duplicate_layout_positions)
                ],
                "markers_using_same_wells": {
                    f"Plate {plate_index}:{position}": labels
                    for (plate_index, position), labels in duplicate_layout_positions.items()
                },
            },
        )

    marker_numbers = get_marker_numbers(layout_matches, marker)
    plate_indices = get_layout_plate_indices(layout_matches)

    findings.append({
        "type": "layout_summary",
        "marker": marker,
        "marker_count": len(layout_matches),
        "layout_position_count": len(all_layout_positions),
        "unique_layout_position_count": len(set(all_layout_positions)),
        "layout_marker_numbers": marker_numbers,
        "layout_plate_count": len(plate_indices),
        "layout_plate_indices": plate_indices,
        "repeating_layout_template": (
            "Single plate-layout sheet/table will be reused for later plates."
            if len(plate_indices) == 1
            else "Each sheet/table is treated as a plate-specific layout."
        ),
        "duplicate_layout_positions": [
            f"Plate {plate_index}:{position}"
            for plate_index, position in sorted(duplicate_layout_positions)
        ],
    })

    return findings, issues
'''

NEW_CROSS_MATCH = r'''
def cross_match_positions(
    aspirate_steps: List[str],
    dispense_steps: List[str],
    replicate_value: str,
    ordered_steps: Optional[List[Dict[str, Any]]] = None,
) -> Tuple[bool, Any]:
    replicate_count = REPLICATE_MAP.get(str(replicate_value).lower(), 1)
    aspirate_to_dispense_map = {}
    current_aspirate_by_channel = {}
    occurrence = 0

    if ordered_steps:
        for ordered_step in ordered_steps:
            step_type = str(ordered_step.get("type", "")).lower()
            raw = str(ordered_step.get("raw", ""))

            if "aspirate" in step_type:
                for aspirate in normalize_channel_steps([raw]):
                    occurrence += 1
                    key = (aspirate[0], aspirate[1], aspirate[2], aspirate[3], occurrence)
                    current_aspirate_by_channel[aspirate[0]] = key
                    aspirate_to_dispense_map.setdefault(key, [])

            elif "dispense" in step_type:
                for dispense in normalize_channel_steps([raw]):
                    channel = dispense[0]
                    aspirate_key = current_aspirate_by_channel.get(channel)
                    if aspirate_key is None:
                        return False, f"Channel {channel} in dispense step does not match any preceding aspirate step."
                    aspirate_to_dispense_map.setdefault(aspirate_key, []).append(dispense[1:])

    if not aspirate_to_dispense_map:
        normalized_aspirate_steps = normalize_channel_steps(aspirate_steps)
        normalized_dispense_steps = normalize_channel_steps(dispense_steps)
        aspirate_queues = defaultdict(list)

        for aspirate in normalized_aspirate_steps:
            occurrence += 1
            key = (aspirate[0], aspirate[1], aspirate[2], aspirate[3], occurrence)
            aspirate_queues[aspirate[0]].append(key)
            aspirate_to_dispense_map.setdefault(key, [])

        channel_cursor = defaultdict(int)
        for dispense in normalized_dispense_steps:
            channel = dispense[0]
            queue = aspirate_queues.get(channel, [])
            if not queue:
                return False, f"Channel {channel} in dispense step does not match any aspirate step."

            cursor = min(channel_cursor[channel], len(queue) - 1)
            aspirate_key = queue[cursor]
            aspirate_to_dispense_map[aspirate_key].append(dispense[1:])

            if (
                replicate_count > 0
                and len(aspirate_to_dispense_map[aspirate_key]) >= replicate_count
                and channel_cursor[channel] < len(queue) - 1
            ):
                channel_cursor[channel] += 1

    for aspirate_step, dispenses in aspirate_to_dispense_map.items():
        if replicate_count > 0 and len(dispenses) % replicate_count != 0:
            return False, (
                f"Aspirate step {aspirate_step[:4]} has {len(dispenses)} dispenses, "
                f"which is not divisible by replicate count {replicate_count}."
            )

    return True, dict(aspirate_to_dispense_map)
'''

NEW_VALIDATE_TRACE_LAYOUT = r'''
def validate_trace_against_layout(
    aspirate_to_dispense_map: Dict[Any, Any],
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker: str,
) -> Tuple[List[Dict[str, Any]], List[ValidationIssue]]:
    findings = []
    issues = []
    marker = marker.upper().strip()

    seen_trace_destinations = []
    trace_transfer_by_marker = defaultdict(list)
    mismatch_by_marker = {}
    replicate_mismatch_by_marker = {}
    plate_order_mismatch_by_marker = {}
    destination_labware_to_plate = {}
    next_destination_plate = 1

    for sample_index, (aspirate_step, dispenses) in enumerate(aspirate_to_dispense_map.items(), start=1):
        marker_label = f"{marker}{sample_index}"
        layout_marker, expected_plate_index, local_sample_number = resolve_layout_marker_for_sample(
            sample_index,
            layout_matches,
            marker,
        )

        source_labware = str(aspirate_step[1]).strip()
        source_position = normalize_position_text(aspirate_step[2])
        expected_positions = get_layout_marker_positions(
            layout_matches,
            layout_marker,
            plate_index=expected_plate_index,
        )
        actual_positions = [normalize_position_text(dispense[1]) for dispense in dispenses]

        actual_labware_in_order = []
        for dispense in dispenses:
            if not dispense:
                continue
            labware_name = str(dispense[0]).strip()
            if labware_name and labware_name not in actual_labware_in_order:
                actual_labware_in_order.append(labware_name)

        actual_plate_indices = []
        for labware_name in actual_labware_in_order:
            if labware_name not in destination_labware_to_plate:
                destination_labware_to_plate[labware_name] = next_destination_plate
                next_destination_plate += 1
            actual_plate_indices.append(destination_labware_to_plate[labware_name])

        expected_set = set(expected_positions)
        actual_set = set(actual_positions)
        missing_for_this_sample = sorted(expected_set - actual_set)
        wrong_for_this_sample = sorted(actual_set - expected_set)

        for labware_name in actual_labware_in_order or [""]:
            for position in actual_positions:
                seen_trace_destinations.append((labware_name, position))

        plate_order_ok = True
        if actual_plate_indices:
            plate_order_ok = all(index == expected_plate_index for index in actual_plate_indices)

        status = "pass"
        if (
            missing_for_this_sample
            or wrong_for_this_sample
            or len(actual_positions) != len(expected_positions)
            or not plate_order_ok
        ):
            status = "fail"

        trace_transfer_by_marker[marker_label].append({
            "source_labware": source_labware,
            "source_position": source_position,
            "destination_labware": actual_labware_in_order,
            "destination_wells": actual_positions,
            "expected_plate_index": expected_plate_index,
            "local_sample_number": local_sample_number,
        })

        findings.append({
            "type": "trace_layout_match",
            "marker": marker_label,
            "plate_layout_marker_used": layout_marker,
            "expected_plate_index": expected_plate_index,
            "local_sample_number_on_plate": local_sample_number,
            "source_labware": source_labware,
            "source_position": source_position,
            "expected_destinations": expected_positions,
            "actual_destinations": actual_positions,
            "actual_destination_labware": actual_labware_in_order,
            "actual_destination_plate_indices": actual_plate_indices,
            "missing_destinations": missing_for_this_sample,
            "unexpected_destinations": wrong_for_this_sample,
            "plate_order_match": plate_order_ok,
            "status": status,
        })

        if missing_for_this_sample or wrong_for_this_sample:
            mismatch_by_marker[marker_label] = {
                "expected_plate_index": expected_plate_index,
                "plate_layout_marker_used": layout_marker,
                "expected_positions": expected_positions,
                "actual_positions": actual_positions,
                "missing_for_this_sample": missing_for_this_sample,
                "wrong_for_this_sample": wrong_for_this_sample,
            }

        if len(actual_positions) != len(expected_positions):
            replicate_mismatch_by_marker[marker_label] = {
                "expected_plate_index": expected_plate_index,
                "plate_layout_marker_used": layout_marker,
                "expected_replicate_count_from_plate_layout": len(expected_positions),
                "actual_dispense_count_from_trace": len(actual_positions),
                "expected_destination_wells": expected_positions,
                "trace_destination_wells_for_this_sample": actual_positions,
            }

        if not plate_order_ok:
            plate_order_mismatch_by_marker[marker_label] = {
                "expected_plate_index": expected_plate_index,
                "actual_destination_plate_indices": actual_plate_indices,
                "actual_destination_labware": actual_labware_in_order,
                "plate_layout_marker_used": layout_marker,
            }

    for marker_label, transfers in trace_transfer_by_marker.items():
        findings.append({"type": "actual_trace_transfer", "marker": marker_label, "transfers": transfers})

    for marker_label, data in sorted(mismatch_by_marker.items(), key=lambda item: extract_position_number(item[0]) or 0):
        add_issue(
            issues,
            "critical",
            "trace vs plate layout",
            (
                f"{marker_label}: sample was dispensed to the wrong destination well(s). "
                f"Expected {join_values(data['expected_positions'])}; "
                f"trace showed {join_values(data['actual_positions'])}."
            ),
            {
                "expected_plate_index": data["expected_plate_index"],
                "plate_layout_marker_used": data["plate_layout_marker_used"],
                "missing_correct_wells_for_this_sample": data["missing_for_this_sample"],
                "wrong_wells_used_for_this_sample": data["wrong_for_this_sample"],
            },
        )

    for marker_label, evidence in sorted(replicate_mismatch_by_marker.items(), key=lambda item: extract_position_number(item[0]) or 0):
        add_issue(
            issues,
            "major",
            "replicate mismatch",
            f"{marker_label}: plate layout replicate count and trace dispense count do not match.",
            evidence,
        )

    for marker_label, evidence in sorted(plate_order_mismatch_by_marker.items(), key=lambda item: extract_position_number(item[0]) or 0):
        add_issue(
            issues,
            "critical",
            "trace vs plate layout",
            f"{marker_label}: destination plate/labware order does not match the expected sequential plate order.",
            evidence,
        )

    duplicate_trace_dests = sorted({
        f"{labware}:{position}"
        for labware, position in set(seen_trace_destinations)
        if seen_trace_destinations.count((labware, position)) > 1
    })

    if duplicate_trace_dests:
        add_issue(
            issues,
            "major",
            "trace duplicate destinations",
            "Duplicate destination wells were found on the same destination labware in the trace.",
            {"duplicate_destination_wells": duplicate_trace_dests},
        )

    labware_findings, labware_issues = validate_trace_labware_against_layout_format(
        aspirate_to_dispense_map,
        layout_matches,
    )
    findings.extend(labware_findings)
    issues.extend(labware_issues)

    return findings, issues
'''


def apply_patch(source: str) -> str:
    if 'OPTIONAL_ROLE_OPTIONS = ["tip eject"]' not in source:
        source = replace_once(
            source,
            'ROLE_OPTIONS = ["tip pick up", "aspirate", "dispense"]',
            'ROLE_OPTIONS = ["tip pick up", "aspirate", "dispense"]\nOPTIONAL_ROLE_OPTIONS = ["tip eject"]',
            "role constants",
        )

    for name, replacement in [
        ("validate_step_sets", NEW_VALIDATE_STEP_SETS),
        ("validate_step_configs", NEW_VALIDATE_STEP_CONFIGS),
        ("get_required_transfer_iterations", NEW_GET_REQUIRED_TRANSFER_ITERATIONS),
        ("build_execution_plan", NEW_BUILD_EXECUTION_PLAN),
        ("generate_deterministic_transfer_pairs", NEW_GENERATE_PAIRS),
        ("parse_dataframe_plate_layout", NEW_PARSE_DATAFRAME),
        ("parse_plate_layout_file", NEW_PARSE_LAYOUT),
        ("get_layout_marker_positions", NEW_GET_LAYOUT_MARKER_POSITIONS),
        ("validate_layout_matches", NEW_VALIDATE_LAYOUT),
        ("cross_match_positions", NEW_CROSS_MATCH),
        ("validate_trace_against_layout", NEW_VALIDATE_TRACE_LAYOUT),
    ]:
        source = replace_top_level_function(source, name, replacement)

    if "def resolve_layout_marker_for_sample(" not in source:
        source = insert_after_top_level_function(source, "get_marker_numbers", NEW_LAYOUT_HELPERS)

    source = replace_once(
        source,
        '        self.variant_var = tk.IntVar(value=0)\n',
        '        self.variant_var = tk.IntVar(value=0)\n        self.tip_eject_var = tk.IntVar(value=0)\n',
        "tip eject checkbox variable",
    )

    source = replace_once(
        source,
        '        self.sequence_set_widgets = {}\n\n        self.liquid_class_var',
        '        self.sequence_set_widgets = {}\n        self.sequence_role_widgets = {}\n\n        self.liquid_class_var',
        "role widget state",
    )

    old_variant_ui = '''        self.make_checkbutton(options_row, text="Variant", variable=self.variant_var,
                              command=self.update_set_column_visibility).pack(side="left", padx=(0, 12))

        report_box = tk.LabelFrame(options_row, text="Report Options", bg=UI_BG, fg=UI_TEXT_FG, bd=3, relief="ridge")'''
    new_variant_ui = '''        self.make_checkbutton(options_row, text="Variant", variable=self.variant_var,
                              command=self.update_set_column_visibility).pack(side="left", padx=(0, 8))

        self.make_checkbutton(
            options_row,
            text="Tip Eject",
            variable=self.tip_eject_var,
            command=self.update_tip_eject_role_options,
        ).pack(side="left", padx=(0, 12))

        report_box = tk.LabelFrame(options_row, text="Report Options", bg=UI_BG, fg=UI_TEXT_FG, bd=3, relief="ridge")'''
    source = replace_once(source, old_variant_ui, new_variant_ui, "Tip Eject checkbox UI")

    source = replace_once(
        source,
        '        self.sequence_set_widgets = {}\n        self.set_header_widget = None\n',
        '        self.sequence_set_widgets = {}\n        self.sequence_role_widgets = {}\n        self.set_header_widget = None\n',
        "render role widget reset",
    )

    source = replace_once(
        source,
        '                values=ROLE_OPTIONS,\n                width=12,\n',
        '                values=self.get_available_role_options(),\n                width=12,\n',
        "dynamic role options",
    )

    source = replace_once(
        source,
        '            role_combo.grid(row=row_index, column=4, padx=3, pady=1, sticky="w")\n',
        '            role_combo.grid(row=row_index, column=4, padx=3, pady=1, sticky="w")\n            self.sequence_role_widgets[sequence_name] = role_combo\n',
        "remember role comboboxes",
    )

    class_methods = '''    def get_available_role_options(self):
        options = list(ROLE_OPTIONS)
        if bool(self.tip_eject_var.get()):
            options.extend(OPTIONAL_ROLE_OPTIONS)
        return options

    def update_tip_eject_role_options(self):
        options = self.get_available_role_options()

        for sequence_name, combo in getattr(self, "sequence_role_widgets", {}).items():
            try:
                combo.configure(values=options)
            except Exception:
                continue

            if (
                not bool(self.tip_eject_var.get())
                and self.sequence_role_vars.get(sequence_name) is not None
                and self.sequence_role_vars[sequence_name].get().strip().lower() == "tip eject"
            ):
                self.sequence_role_vars[sequence_name].set("aspirate")

'''
    if "    def get_available_role_options(self):" not in source:
        marker = "    def update_set_column_visibility(self):"
        if marker not in source:
            raise RuntimeError("Could not find update_set_column_visibility method.")
        source = source.replace(marker, class_methods + marker, 1)

    source = replace_once(
        source,
        '        variant_mode = bool(self.variant_var.get())\n',
        '        variant_mode = bool(self.variant_var.get())\n        tip_eject_enabled = bool(self.tip_eject_var.get())\n',
        "collect tip eject state",
    )

    source = replace_once(
        source,
        '        validation_error = validate_step_sets(configs, variant_mode, use_set_order)\n',
        '        validation_error = validate_step_sets(\n            configs,\n            variant_mode,\n            use_set_order,\n            tip_eject_enabled=tip_eject_enabled,\n        )\n',
        "step-set validation call",
    )

    source = replace_once(
        source,
        '        issues = validate_step_configs(configs)\n',
        '        issues = validate_step_configs(\n            configs,\n            tip_eject_enabled=bool(self.tip_eject_var.get()),\n        )\n',
        "step config validation call",
    )

    old_cross_call = '''                ok, cross_match = cross_match_positions(
                    aspirate_steps,
                    dispense_steps,
                    configs[0].replicate,
                )'''
    new_cross_call = '''                ok, cross_match = cross_match_positions(
                    aspirate_steps,
                    dispense_steps,
                    configs[0].replicate,
                    ordered_steps=ordered_steps,
                )'''
    source = replace_once(source, old_cross_call, new_cross_call, "ordered trace cross-match call")

    source = source.replace(
        'item.get("action") in {"aspirate", "dispense"}',
        'item.get("action") in {"aspirate", "dispense", "tip_eject"}',
    )
    source = source.replace(
        '                        and action in {"aspirate", "dispense"}\n                        and not bool(item.get("manual"))',
        '                        and action in {"aspirate", "dispense", "tip_eject"}\n                        and not bool(item.get("manual"))',
        1,
    )

    return source


def main():
    candidate = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("devpal_lite_runner.py")
    if not candidate.exists() and candidate.name == "devpal_lite_runner.py":
        alt = Path("t3")
        if alt.exists():
            candidate = alt

    if not candidate.exists():
        raise SystemExit(
            "Could not find devpal_lite_runner.py. Run this updater from the project directory "
            "or pass the file path as the first argument."
        )

    original = candidate.read_text(encoding="utf-8")
    updated = apply_patch(original)

    backup = candidate.with_name(candidate.name + ".before_multiplate_tip_eject.bak")
    shutil.copy2(candidate, backup)

    output = candidate.with_name(candidate.stem + "_updated.py")
    output.write_text(updated, encoding="utf-8")

    print(f"Updated file created: {output}")
    print(f"Original preserved: {candidate}")
    print(f"Backup created: {backup}")
    print("The original is not overwritten.")


if __name__ == "__main__":
    main()
