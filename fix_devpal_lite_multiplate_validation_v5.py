from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path


MARKER_DETECTOR = r'''
def detect_plate_layout_marker(layout_path: str) -> str:
    if not layout_path:
        return ""

    ext = os.path.splitext(layout_path)[1].lower()
    candidates = defaultdict(set)
    occurrences = defaultdict(int)

    def add_cell(value):
        if value is None:
            return

        text = str(value).strip().upper()
        if not text:
            return

        for match in re.finditer(
            r"(?<![A-Z0-9])([A-Z][A-Z0-9 _\-]*?)[ _\-]*(\d+)(?![A-Z0-9])",
            text,
        ):
            prefix = re.sub(r"[ _\-]+$", "", match.group(1).strip())
            prefix = re.sub(r"[ _\-]+", " ", prefix).strip()

            if not prefix:
                continue

            try:
                number = int(match.group(2))
            except Exception:
                continue

            candidates[prefix].add(number)
            occurrences[prefix] += 1

    if ext == ".csv":
        if pd is None:
            raise RuntimeError("pandas is required for CSV plate layout parsing.")

        df = pd.read_csv(layout_path, header=None, dtype=str)
        for value in df.to_numpy().flatten():
            if pd.isna(value):
                continue
            add_cell(value)

    elif ext in (".xls", ".xlsx"):
        if pd is None:
            raise RuntimeError("pandas is required for Excel plate layout parsing.")

        sheets = pd.read_excel(layout_path, sheet_name=None, header=None, dtype=str)
        for df in sheets.values():
            for value in df.to_numpy().flatten():
                if pd.isna(value):
                    continue
                add_cell(value)

    elif ext == ".docx":
        if Document is None:
            raise RuntimeError("python-docx is required for DOCX plate layout parsing.")

        doc = Document(layout_path)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    add_cell(cell.text)

    else:
        return ""

    if not candidates:
        return ""

    def score(prefix):
        nums = sorted(candidates[prefix])
        unique_count = len(nums)
        contiguous = 0
        expected = nums[0] if nums else 1

        for number in nums:
            if number == expected:
                contiguous += 1
                expected += 1
            elif number > expected:
                break

        semantic_bonus = 0
        if "SAMPLE" in prefix:
            semantic_bonus += 1000
        if "SPECIMEN" in prefix:
            semantic_bonus += 500

        return (
            semantic_bonus + unique_count * 100 + contiguous * 10 + occurrences[prefix],
            unique_count,
            occurrences[prefix],
        )

    best = max(candidates, key=score)

    if len(candidates[best]) < 2 and "SAMPLE" not in best and "SPECIMEN" not in best:
        return ""

    return best
'''


TRACE_LABWARE_VALIDATOR = r'''
def validate_trace_labware_against_layout_format(
    aspirate_to_dispense_map: Dict[Any, Any],
    layout_matches: Dict[str, List[Dict[str, Any]]],
) -> Tuple[List[Dict[str, Any]], List[ValidationIssue]]:
    findings = []
    issues = []

    layout_positions = []

    for entries in (layout_matches or {}).values():
        for entry in entries or []:
            position = normalize_position_text(entry.get("dest_value", ""))
            if position:
                layout_positions.append(position)

    def position_rc(position):
        match = re.fullmatch(r"([A-P])(\d{1,2})", normalize_position_text(position))
        if not match:
            return None

        row_number = ord(match.group(1)) - ord("A") + 1
        col_number = int(match.group(2))
        return row_number, col_number

    parsed_layout_positions = [
        rc for rc in (position_rc(p) for p in layout_positions)
        if rc is not None
    ]

    if not parsed_layout_positions:
        return findings, issues

    max_layout_row = max(row for row, _ in parsed_layout_positions)
    max_layout_col = max(col for _, col in parsed_layout_positions)

    formats = [
        ("24-well", 4, 6),
        ("48-well", 6, 8),
        ("96-well", 8, 12),
        ("384-well", 16, 24),
    ]

    expected_format = "unknown"
    expected_rows = None
    expected_cols = None

    for format_name, rows, cols in formats:
        if max_layout_row <= rows and max_layout_col <= cols:
            expected_format = format_name
            expected_rows = rows
            expected_cols = cols
            break

    destinations_by_labware = defaultdict(list)

    for _aspirate, dispenses in (aspirate_to_dispense_map or {}).items():
        for dispense in dispenses or []:
            if len(dispense) < 2:
                continue

            labware = str(dispense[0]).strip()
            position = normalize_position_text(dispense[1])

            if labware and position:
                destinations_by_labware[labware].append(position)

    for labware, positions in destinations_by_labware.items():
        invalid_positions = []

        for position in positions:
            rc = position_rc(position)

            if rc is None:
                invalid_positions.append(position)
                continue

            row_number, col_number = rc

            if (
                expected_rows is not None
                and expected_cols is not None
                and (
                    row_number > expected_rows
                    or col_number > expected_cols
                )
            ):
                invalid_positions.append(position)

        findings.append({
            "type": "trace_labware_format",
            "labware": labware,
            "labware_format": expected_format if not invalid_positions else "incompatible",
            "expected_layout_format": expected_format,
            "destination_positions": positions,
            "invalid_positions": invalid_positions,
            "status": "pass" if not invalid_positions else "fail",
        })

        if invalid_positions:
            add_issue(
                issues,
                "critical",
                "labware mismatch",
                (
                    f"URGENT ERROR: Dispense used destination position(s) on {labware} "
                    f"that do not fit the uploaded {expected_format} plate layout."
                ),
                {
                    "labware": labware,
                    "expected_layout_format": expected_format,
                    "invalid_destination_positions": invalid_positions,
                    "all_destination_positions": positions,
                },
            )

    return findings, issues
'''


TRACE_VALIDATOR = r'''
def validate_trace_against_layout(
    aspirate_to_dispense_map: Dict[Any, Any],
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker: str,
) -> Tuple[List[Dict[str, Any]], List[ValidationIssue]]:
    findings = []
    issues = []
    marker = str(marker or "").strip().upper()

    destination_labware_to_plate = {}
    next_destination_plate = 1
    seen_physical_destinations = []
    mismatch_by_marker = {}
    replicate_mismatch_by_marker = {}
    plate_order_mismatch_by_marker = {}

    for sample_index, (aspirate_step, dispenses) in enumerate(
        (aspirate_to_dispense_map or {}).items(),
        start=1,
    ):
        display_marker = f"{marker}{sample_index}"

        layout_marker, expected_plate_index, local_sample_number = resolve_layout_marker_for_sample(
            sample_index,
            layout_matches,
            marker,
        )

        expected_positions = get_layout_marker_positions(
            layout_matches,
            layout_marker,
            plate_index=expected_plate_index,
        )

        actual_records = []

        for dispense in dispenses or []:
            if len(dispense) < 2:
                continue

            labware_name = str(dispense[0]).strip()
            position = normalize_position_text(dispense[1])

            if not labware_name or not position:
                continue

            if labware_name not in destination_labware_to_plate:
                destination_labware_to_plate[labware_name] = next_destination_plate
                next_destination_plate += 1

            actual_plate_index = destination_labware_to_plate[labware_name]

            actual_records.append({
                "labware": labware_name,
                "position": position,
                "plate_index": actual_plate_index,
            })

            seen_physical_destinations.append((labware_name, position))

        expected_plate_records = [
            record
            for record in actual_records
            if record["plate_index"] == expected_plate_index
        ]

        actual_positions = [record["position"] for record in expected_plate_records]

        actual_labware = []
        for record in actual_records:
            if record["labware"] not in actual_labware:
                actual_labware.append(record["labware"])

        actual_plate_indices = []
        for record in actual_records:
            if record["plate_index"] not in actual_plate_indices:
                actual_plate_indices.append(record["plate_index"])

        expected_set = set(expected_positions)
        actual_set = set(actual_positions)

        missing = sorted(expected_set - actual_set)
        unexpected = sorted(actual_set - expected_set)
        plate_order_ok = bool(expected_plate_records)

        status = "pass"
        if (
            missing
            or unexpected
            or len(actual_positions) != len(expected_positions)
            or not plate_order_ok
        ):
            status = "fail"

        findings.append({
            "type": "trace_layout_match",
            "marker": display_marker,
            "plate_layout_marker_used": layout_marker,
            "expected_plate_index": expected_plate_index,
            "local_sample_number_on_plate": local_sample_number,
            "expected_destinations": expected_positions,
            "actual_destinations": actual_positions,
            "actual_destination_labware": actual_labware,
            "actual_destination_plate_indices": actual_plate_indices,
            "missing_destinations": missing,
            "unexpected_destinations": unexpected,
            "plate_order_match": plate_order_ok,
            "status": status,
        })

        if missing or unexpected:
            mismatch_by_marker[display_marker] = {
                "expected_plate_index": expected_plate_index,
                "plate_layout_marker_used": layout_marker,
                "expected_positions": expected_positions,
                "actual_positions": actual_positions,
                "missing": missing,
                "unexpected": unexpected,
            }

        if len(actual_positions) != len(expected_positions):
            replicate_mismatch_by_marker[display_marker] = {
                "expected_plate_index": expected_plate_index,
                "plate_layout_marker_used": layout_marker,
                "expected_replicate_count_from_plate_layout": len(expected_positions),
                "actual_dispense_count_from_trace": len(actual_positions),
                "expected_destination_wells": expected_positions,
                "trace_destination_wells_for_this_sample": actual_positions,
            }

        if not plate_order_ok:
            plate_order_mismatch_by_marker[display_marker] = {
                "expected_plate_index": expected_plate_index,
                "actual_destination_plate_indices": actual_plate_indices,
                "actual_destination_labware": actual_labware,
                "plate_layout_marker_used": layout_marker,
            }

    for marker_label, data in mismatch_by_marker.items():
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
                "missing_correct_wells_for_this_sample": data["missing"],
                "wrong_wells_used_for_this_sample": data["unexpected"],
            },
        )

    for marker_label, evidence in replicate_mismatch_by_marker.items():
        add_issue(
            issues,
            "major",
            "replicate mismatch",
            f"{marker_label}: plate layout replicate count and trace dispense count do not match.",
            evidence,
        )

    for marker_label, evidence in plate_order_mismatch_by_marker.items():
        add_issue(
            issues,
            "critical",
            "trace vs plate layout",
            (
                f"{marker_label}: destination plate/labware order does not match "
                f"the expected sequential plate order."
            ),
            evidence,
        )

    counts = defaultdict(int)
    for physical_destination in seen_physical_destinations:
        counts[physical_destination] += 1

    duplicate_trace_dests = sorted(
        f"{labware}:{position}"
        for (labware, position), count in counts.items()
        if count > 1
    )

    if duplicate_trace_dests:
        add_issue(
            issues,
            "major",
            "trace duplicate destinations",
            (
                "Duplicate destination wells were found on the same physical "
                "destination labware in the trace."
            ),
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


def find_function_span(source: str, name: str):
    start_match = re.search(rf"(?m)^def {re.escape(name)}\s*\(", source)
    if not start_match:
        return None

    start = start_match.start()
    next_match = re.search(r"(?m)^def \w+\s*\(|^class \w+", source[start_match.end():])

    if next_match:
        end = start_match.end() + next_match.start()
    else:
        end = len(source)

    return start, end


def replace_function(source: str, name: str, replacement: str) -> str:
    span = find_function_span(source, name)
    if not span:
        raise RuntimeError(f"Could not find function: {name}")

    start, end = span
    return source[:start] + replacement.strip() + "\n\n" + source[end:]


def ensure_upload_autofill(source: str) -> str:
    if "Auto-detected plate layout marker:" in source:
        return source

    pattern = re.compile(r"(?m)^(?P<indent>\s*)self\.layout_path\s*=\s*path\s*$")
    match = pattern.search(source)
    if not match:
        return source

    indent = match.group("indent")
    addition = (
        match.group(0)
        + "\n"
        + indent + "try:\n"
        + indent + "    detected_marker = detect_plate_layout_marker(path)\n"
        + indent + "    if detected_marker:\n"
        + indent + "        self.marker_var.set(detected_marker)\n"
        + indent + "        try:\n"
        + indent + "            self.status.config(\n"
        + indent + '                text=f"Auto-detected plate layout marker: {detected_marker}"\n'
        + indent + "            )\n"
        + indent + "        except Exception:\n"
        + indent + "            pass\n"
        + indent + "except Exception:\n"
        + indent + "    pass"
    )

    return source[:match.start()] + addition + source[match.end():]


def ensure_collect_configs_autodetect(source: str) -> str:
    if "Marker could not be auto-detected from the uploaded plate layout." in source:
        return source

    pattern = re.compile(
        r'(?m)^(?P<indent>\s*)marker\s*=\s*self\.marker_var\.get\(\)\.strip\(\)(?:\.upper\(\))?\s*$'
    )
    match = pattern.search(source)
    if not match:
        return source

    indent = match.group("indent")
    replacement = (
        indent + "marker = self.marker_var.get().strip().upper()\n"
        + indent + "if not marker and self.layout_path:\n"
        + indent + "    try:\n"
        + indent + "        marker = detect_plate_layout_marker(self.layout_path).strip().upper()\n"
        + indent + "    except Exception:\n"
        + indent + '        marker = ""\n'
        + indent + "    if marker:\n"
        + indent + "        self.marker_var.set(marker)\n"
        + indent + "if not marker:\n"
        + indent + "    messagebox.showerror(\n"
        + indent + '        "Plate Layout Marker",\n'
        + indent + '        "Marker could not be auto-detected from the uploaded plate layout.",\n'
        + indent + "    )\n"
        + indent + "    return None"
    )

    return source[:match.start()] + replacement + source[match.end():]


def apply_fix(source: str) -> str:
    if "def detect_plate_layout_marker(" in source:
        source = replace_function(source, "detect_plate_layout_marker", MARKER_DETECTOR)
    else:
        match = re.search(r"(?m)^def parse_plate_layout_file\s*\(", source)
        if not match:
            raise RuntimeError("Could not find parse_plate_layout_file.")
        source = source[:match.start()] + MARKER_DETECTOR.strip() + "\n\n" + source[match.start():]

    source = replace_function(
        source,
        "validate_trace_labware_against_layout_format",
        TRACE_LABWARE_VALIDATOR,
    )
    source = replace_function(
        source,
        "validate_trace_against_layout",
        TRACE_VALIDATOR,
    )

    source = ensure_upload_autofill(source)
    source = ensure_collect_configs_autodetect(source)

    required = [
        "def detect_plate_layout_marker(",
        "def validate_trace_labware_against_layout_format(",
        "def validate_trace_against_layout(",
        "def resolve_layout_marker_for_sample(",
        "def get_layout_marker_positions(",
    ]
    missing = [item for item in required if item not in source]
    if missing:
        raise RuntimeError("Current DevPal file is missing required functions: " + ", ".join(missing))

    return source


def main():
    if len(sys.argv) < 2:
        raise SystemExit(
            "Pass the exact DevPal Lite Python file you are currently running.\n"
            "Example:\n"
            "python fix_devpal_lite_multiplate_validation_v5.py "
            "devpal_lite_runner_updated_multiplate_fix_v3_plate_helper_v4.py"
        )

    source_path = Path(sys.argv[1])
    if not source_path.exists():
        raise SystemExit(f"File not found: {source_path}")

    original = source_path.read_text(encoding="utf-8")
    updated = apply_fix(original)

    compile(updated, str(source_path), "exec")

    backup_path = source_path.with_name(source_path.name + ".before_multiplate_validation_v5.bak")
    shutil.copy2(source_path, backup_path)

    output_path = source_path.with_name(source_path.stem + "_multiplate_validation_v5.py")
    output_path.write_text(updated, encoding="utf-8")

    print(f"Created: {output_path}")
    print(f"Backup:  {backup_path}")
    print()
    print("V5 fixes applied:")
    print("- S1/S2/... short marker auto-detection")
    print("- UI marker auto-fill where layout_path is assigned")
    print("- marker auto-detect fallback when Run is clicked")
    print("- P1:A1 and Plate2:A1 treated as different physical destinations")
    print("- Plate2/P1 names no longer cause unknown-format critical errors")
    print("- destination geometry checked against uploaded 24/48/96/384 layout")
    print("- full generated DevPal source compiled successfully")


if __name__ == "__main__":
    main()
