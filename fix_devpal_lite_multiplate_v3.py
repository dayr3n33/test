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


def insert_before_top_level_function(source: str, name: str, insertion: str) -> str:
    pattern = re.compile(rf"(?m)^def {re.escape(name)}\s*\(")
    match = pattern.search(source)
    if not match:
        raise RuntimeError(f"Could not find top-level function: {name}")
    return source[:match.start()] + insertion.strip() + "\n\n" + source[match.start():]


MARKER_HELPERS = r'''
def normalize_marker_prefix(value: str) -> str:
    text = str(value or "").strip().upper()
    text = re.sub(r"[\s_\-]+$", "", text)
    return text


def extract_marker_candidates_from_text(text: str) -> List[Tuple[str, int]]:
    text = str(text or "").upper()
    candidates = []

    # Accept both long markers (SAMPLE1) and short markers (S1).
    # Row labels such as A/B/C by themselves are not candidates because they
    # contain no following number in the same cell.
    pattern = re.compile(
        r"(?<![A-Z0-9])([A-Z][A-Z0-9 _\-]*?)[\s_\-]*(\d+)(?![A-Z0-9])"
    )

    for prefix, number in pattern.findall(text):
        prefix = normalize_marker_prefix(prefix)
        if not prefix:
            continue

        try:
            marker_number = int(number)
        except Exception:
            continue

        candidates.append((prefix, marker_number))

    return candidates


def detect_plate_layout_marker(layout_path: str) -> str:
    if not layout_path:
        return ""

    ext = os.path.splitext(layout_path)[1].lower()
    marker_numbers = defaultdict(set)
    marker_occurrences = defaultdict(int)

    def add_text(value: Any):
        if value is None:
            return

        for prefix, number in extract_marker_candidates_from_text(value):
            marker_numbers[prefix].add(number)
            marker_occurrences[prefix] += 1

    if ext == ".csv":
        if pd is None:
            raise RuntimeError("pandas is required for CSV layout parsing.")

        df = pd.read_csv(layout_path, header=None, dtype=str)

        # Only inspect data cells, not the row-label column or column headers.
        # This prevents ordinary plate coordinates from competing with S1/S2/etc.
        for row_idx in range(1, df.shape[0]):
            for col_idx in range(1, df.shape[1]):
                value = df.iat[row_idx, col_idx]
                if pd.isna(value):
                    continue
                add_text(value)

    elif ext in (".xls", ".xlsx"):
        if pd is None:
            raise RuntimeError("pandas is required for Excel layout parsing.")

        sheets = pd.read_excel(layout_path, sheet_name=None, header=None, dtype=str)

        for df in sheets.values():
            for row_idx in range(1, df.shape[0]):
                for col_idx in range(1, df.shape[1]):
                    value = df.iat[row_idx, col_idx]
                    if pd.isna(value):
                        continue
                    add_text(value)

    elif ext == ".docx":
        if Document is None:
            raise RuntimeError("python-docx is required for DOCX layout parsing.")

        doc = Document(layout_path)
        for table in doc.tables:
            for row_idx in range(1, len(table.rows)):
                row = table.rows[row_idx]
                for col_idx in range(1, len(row.cells)):
                    add_text(row.cells[col_idx].text)

    else:
        raise ValueError("Unsupported plate layout file type. Use CSV, XLS, XLSX, or DOCX.")

    if not marker_numbers:
        return ""

    def score(prefix: str):
        numbers = sorted(marker_numbers[prefix])
        occurrences = marker_occurrences[prefix]
        unique_count = len(numbers)
        contiguous_from_one = 0
        expected = 1

        for number in numbers:
            if number == expected:
                contiguous_from_one += 1
                expected += 1
            elif number > expected:
                break

        upper = prefix.upper()
        semantic_bonus = 0
        if "SAMPLE" in upper:
            semantic_bonus += 100000
        if "SPECIMEN" in upper:
            semantic_bonus += 50000
        if upper == "S":
            semantic_bonus += 25000

        return (
            semantic_bonus + contiguous_from_one * 1000 + unique_count * 100 + occurrences,
            contiguous_from_one,
            unique_count,
            occurrences,
            prefix,
        )

    return max(marker_numbers, key=score)


def find_marker_labels_in_text(text: str, marker: str) -> List[str]:
    marker = normalize_marker_prefix(marker)
    if not marker:
        return []

    pattern = re.compile(
        rf"(?<![A-Z0-9]){re.escape(marker)}[\s_\-]*(\d+)(?![A-Z0-9])",
        re.IGNORECASE,
    )

    labels = []
    for number in pattern.findall(str(text or "")):
        try:
            labels.append(f"{marker}{int(number)}")
        except Exception:
            continue

    return labels
'''


NEW_PARSE_LAYOUT = r'''
def parse_plate_layout_file(layout_path: str, marker: str = "") -> Dict[str, List[Dict[str, Any]]]:
    if pd is None:
        raise RuntimeError("pandas is required for layout parsing.")

    marker = normalize_marker_prefix(marker)
    detected_marker = detect_plate_layout_marker(layout_path)

    # Auto-detection is authoritative when the box is blank. If a value was
    # supplied, keep it only when it actually exists in the file; otherwise use
    # the detected marker so stale UI text cannot break layout parsing.
    if not marker:
        marker = detected_marker
    elif detected_marker and marker != detected_marker:
        marker = detected_marker

    if not marker:
        raise ValueError(
            "DevPal Lite could not automatically identify a numbered plate-layout marker."
        )

    plate_layout = defaultdict(list)
    ext = os.path.splitext(layout_path)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(layout_path, header=None, dtype=str)
        parse_dataframe_plate_layout(
            df,
            marker,
            plate_layout,
            plate_index=1,
            plate_name="Plate 1",
        )

    elif ext in (".xls", ".xlsx"):
        sheets = pd.read_excel(layout_path, sheet_name=None, header=None, dtype=str)

        for plate_index, (sheet_name, df) in enumerate(sheets.items(), start=1):
            parse_dataframe_plate_layout(
                df,
                marker,
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

                    plate_col_text = str(header_cells[col_idx].text).strip()
                    plate_col_match = re.search(r"\d+", plate_col_text)
                    if not plate_col_match:
                        continue

                    plate_col = int(plate_col_match.group(0))
                    cell_text = str(row.cells[col_idx].text).strip().upper()
                    matches = find_marker_labels_in_text(cell_text, marker)

                    if not matches:
                        continue

                    dest_position = f"{plate_row}{plate_col}"
                    for found_marker in matches:
                        plate_layout[found_marker].append({
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


NEW_VALIDATE_TRACE_LAYOUT = r'''
def validate_trace_against_layout(
    aspirate_to_dispense_map: Dict[Any, Any],
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker: str,
) -> Tuple[List[Dict[str, Any]], List[ValidationIssue]]:
    findings = []
    issues = []
    marker = normalize_marker_prefix(marker)

    seen_trace_destinations = []
    trace_transfer_by_marker = defaultdict(list)
    mismatch_by_marker = {}
    replicate_mismatch_by_marker = {}
    plate_order_mismatch_by_marker = {}

    destination_labware_to_plate = {}
    next_destination_plate = 1

    for sample_index, (aspirate_step, dispenses) in enumerate(
        aspirate_to_dispense_map.items(),
        start=1,
    ):
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

        actual_positions = [
            normalize_position_text(dispense[1])
            for dispense in dispenses
        ]

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

        # Duplicate tracking MUST include destination labware. A1 on P1 and A1
        # on Plate2 are different physical destinations and are not duplicates.
        if actual_labware_in_order:
            for dispense in dispenses:
                if not dispense:
                    continue
                seen_trace_destinations.append(
                    (str(dispense[0]).strip(), normalize_position_text(dispense[1]))
                )
        else:
            for position in actual_positions:
                seen_trace_destinations.append((f"Plate{expected_plate_index}", position))

        trace_transfer_by_marker[marker_label].append({
            "source_labware": source_labware,
            "source_position": source_position,
            "destination_labware": actual_labware_in_order,
            "destination_wells": actual_positions,
            "expected_plate_index": expected_plate_index,
            "local_sample_number": local_sample_number,
            "plate_layout_marker_used": layout_marker,
        })

        plate_order_ok = True
        if actual_plate_indices:
            plate_order_ok = all(
                plate_index == expected_plate_index
                for plate_index in actual_plate_indices
            )

        status = "pass"
        if (
            missing_for_this_sample
            or wrong_for_this_sample
            or len(actual_positions) != len(expected_positions)
            or not plate_order_ok
        ):
            status = "fail"

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
        findings.append({
            "type": "actual_trace_transfer",
            "marker": marker_label,
            "transfers": transfers,
        })

    for marker_label, data in sorted(
        mismatch_by_marker.items(),
        key=lambda item: extract_position_number(item[0]) or 0,
    ):
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

    for marker_label, evidence in sorted(
        replicate_mismatch_by_marker.items(),
        key=lambda item: extract_position_number(item[0]) or 0,
    ):
        add_issue(
            issues,
            "major",
            "replicate mismatch",
            f"{marker_label}: plate layout replicate count and trace dispense count do not match.",
            evidence,
        )

    for marker_label, evidence in sorted(
        plate_order_mismatch_by_marker.items(),
        key=lambda item: extract_position_number(item[0]) or 0,
    ):
        add_issue(
            issues,
            "critical",
            "trace vs plate layout",
            (
                f"{marker_label}: destination plate/labware order does not match the "
                f"expected sequential plate order."
            ),
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


NEW_VALIDATE_LABWARE = r'''
def validate_trace_labware_against_layout_format(
    aspirate_to_dispense_map: Dict[Any, Any],
    layout_matches: Dict[str, List[Dict[str, Any]]],
) -> Tuple[List[Dict[str, Any]], List[ValidationIssue]]:
    findings = []
    issues = []

    layout_summary = summarize_layout_format(layout_matches)
    expected_format = str(layout_summary.get("inferred_plate_format", "unknown") or "unknown")

    expected_dims = {
        "24-well": (4, 6),
        "48-well": (6, 8),
        "96-well": (8, 12),
        "384-well": (16, 24),
    }

    dest_by_labware = defaultdict(list)

    for _, dispenses in aspirate_to_dispense_map.items():
        for dispense in dispenses:
            if not dispense:
                continue
            labware = str(dispense[0]).strip()
            position = normalize_position_text(dispense[1])
            if labware and position:
                dest_by_labware[labware].append(position)

    for labware, positions in dest_by_labware.items():
        inferred = infer_plate_format_from_labware_name(labware)
        inferred_lower = str(inferred or "").lower()

        format_known = inferred in expected_dims
        format_ambiguous = (
            not format_known
            or inferred_lower in {"unknown", "plate/unknown format"}
            or "unknown" in inferred_lower
        )

        positions_within_expected = True
        bad_positions = []

        if expected_format in expected_dims:
            max_rows, max_cols = expected_dims[expected_format]

            for position in positions:
                match = re.match(r"^([A-P])(\d{1,2})$", position)
                if not match:
                    continue

                row_num = ord(match.group(1)) - ord("A") + 1
                col_num = int(match.group(2))

                if row_num > max_rows or col_num > max_cols:
                    positions_within_expected = False
                    bad_positions.append(position)

        status = "pass"
        message = ""

        if expected_format in expected_dims and not positions_within_expected:
            status = "fail"
            message = (
                f"URGENT ERROR: Dispense used position(s) outside the uploaded "
                f"{expected_format} plate layout on {labware}."
            )
            issues.append(ValidationIssue(
                "critical",
                "URGENT_LABWARE_MISMATCH",
                message,
                {
                    "labware": labware,
                    "labware_format": inferred,
                    "expected_layout_format": expected_format,
                    "destination": sorted(set(positions)),
                    "outside_expected_format": sorted(set(bad_positions)),
                },
            ))

        elif (
            expected_format in expected_dims
            and format_known
            and inferred != expected_format
        ):
            status = "fail"
            message = (
                f"URGENT ERROR: Dispense used labware {labware}, detected as {inferred}, "
                f"but the uploaded plate layout appears to be {expected_format}."
            )
            issues.append(ValidationIssue(
                "critical",
                "URGENT_LABWARE_MISMATCH",
                message,
                {
                    "labware": labware,
                    "labware_format": inferred,
                    "expected_layout_format": expected_format,
                    "destination": sorted(set(positions)),
                },
            ))

        elif expected_format in expected_dims and format_ambiguous:
            # A trace name such as P1 or Plate2 identifies the physical plate,
            # not necessarily its geometry. Do not call that a mismatch when all
            # actual positions are valid for the uploaded layout format.
            message = (
                f"Trace destination labware {labware} does not encode a specific plate "
                f"format; positions are valid for the uploaded {expected_format} layout."
            )

        else:
            message = (
                f"Destination labware {labware} is compatible with the uploaded "
                f"plate layout based on available format and position information."
            )

        findings.append({
            "type": "trace_labware_vs_layout_format",
            "status": status,
            "labware": labware,
            "labware_format": inferred,
            "expected_layout_format": expected_format,
            "destination": sorted(set(positions)),
            "positions_within_expected_format": positions_within_expected,
            "message": message,
        })

    return findings, issues
'''


def apply_patch(source: str) -> str:
    # Replace marker helpers if V2 already inserted them, otherwise insert them.
    helper_names = [
        "normalize_marker_prefix",
        "extract_marker_candidates_from_text",
        "detect_plate_layout_marker",
        "find_marker_labels_in_text",
    ]

    if all(f"def {name}(" in source for name in helper_names):
        for name in helper_names:
            # Extract just this function from the combined helper block.
            pattern = re.compile(rf"(?ms)^def {re.escape(name)}\s*\(.*?(?=^def |^class |\Z)")
            m = pattern.search(MARKER_HELPERS)
            if not m:
                raise RuntimeError(f"Internal updater error for helper {name}")
            source = replace_top_level_function(source, name, m.group(0))
    else:
        source = insert_before_top_level_function(source, "parse_dataframe_plate_layout", MARKER_HELPERS)

    source = replace_top_level_function(source, "parse_plate_layout_file", NEW_PARSE_LAYOUT)
    source = replace_top_level_function(source, "validate_trace_against_layout", NEW_VALIDATE_TRACE_LAYOUT)
    source = replace_top_level_function(
        source,
        "validate_trace_labware_against_layout_format",
        NEW_VALIDATE_LABWARE,
    )

    # Ensure upload_layout populates the field from auto-detection even if V2
    # was not successfully applied.
    if "Auto-detected marker:" not in source:
        old = '''        if path:\n            self.layout_path = path\n            self.layout_label.config(text=path)\n'''
        new = '''        if path:\n            self.layout_path = path\n            self.layout_label.config(text=path)\n\n            try:\n                detected_marker = detect_plate_layout_marker(path)\n                if detected_marker:\n                    self.marker_var.set(detected_marker)\n                    self.status.config(\n                        text=f"Plate layout loaded. Auto-detected marker: {detected_marker}"\n                    )\n            except Exception:\n                pass\n'''
        if old in source:
            source = source.replace(old, new, 1)

    # If V2 upload code exists, make detection authoritative so a stale marker
    # from a previous file is replaced when a new layout is uploaded.
    source = source.replace(
        '            if not self.marker_var.get().strip():\n                try:\n                    detected_marker = detect_plate_layout_marker(path)',
        '            try:\n                detected_marker = detect_plate_layout_marker(path)'
    )
    source = source.replace(
        '                    if detected_marker:\n                        self.marker_var.set(detected_marker)\n                        self.status.config(text=f"Plate layout loaded. Auto-detected marker: {detected_marker}")\n                except Exception:\n                    pass',
        '                if detected_marker:\n                    self.marker_var.set(detected_marker)\n                    self.status.config(text=f"Plate layout loaded. Auto-detected marker: {detected_marker}")\n            except Exception:\n                pass'
    )

    return source


def main():
    candidate = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("devpal_lite_runner_updated_multiplate_fix_v2.py")

    if not candidate.exists():
        alternatives = [
            Path("devpal_lite_runner_updated.py"),
            Path("devpal_lite_runner.py"),
            Path("t3_updated_multiplate_fix_v2.py"),
            Path("t3_updated.py"),
            Path("t3"),
        ]
        candidate = next((path for path in alternatives if path.exists()), candidate)

    if not candidate.exists():
        raise SystemExit(
            "Could not find the DevPal Lite source file. Run this updater from the "
            "project directory or pass the source file path as the first argument."
        )

    original = candidate.read_text(encoding="utf-8")
    updated = apply_patch(original)

    backup = candidate.with_name(candidate.name + ".before_multiplate_fix_v3.bak")
    shutil.copy2(candidate, backup)

    output = candidate.with_name(candidate.stem + "_multiplate_fix_v3.py")
    output.write_text(updated, encoding="utf-8")

    compile(updated, str(output), "exec")

    print(f"Corrected file created: {output}")
    print(f"Original preserved: {candidate}")
    print(f"Backup created: {backup}")
    print()
    print("V3 fixes:")
    print("- short marker auto-detection such as S1/S2/S3")
    print("- marker detection scans layout data cells instead of row/column headers")
    print("- P1:A1 and Plate2:A1 are treated as different physical destinations")
    print("- Plate2 / P1 names are treated as plate identifiers when geometry is unknown")
    print("- unknown trace plate-name geometry is not a critical mismatch when positions fit the uploaded format")


if __name__ == "__main__":
    main()
