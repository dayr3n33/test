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


NEW_MARKER_HELPERS = r'''
def normalize_marker_prefix(value: str) -> str:
    text = str(value or "").strip().upper()
    text = re.sub(r"[\s_\-]+$", "", text)
    return text


def extract_marker_candidates_from_text(text: str) -> List[Tuple[str, int]]:
    text = str(text or "").upper()
    candidates = []

    pattern = re.compile(
        r"(?<![A-Z0-9])([A-Z][A-Z0-9 _\-]*?[A-Z])[\s_\-]*(\d+)(?![A-Z0-9])"
    )

    for prefix, number in pattern.findall(text):
        prefix = normalize_marker_prefix(prefix)
        if not prefix or len(prefix) == 1:
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
        for value in df.to_numpy().flatten():
            if pd.isna(value):
                continue
            add_text(value)

    elif ext in (".xls", ".xlsx"):
        if pd is None:
            raise RuntimeError("pandas is required for Excel layout parsing.")
        sheets = pd.read_excel(layout_path, sheet_name=None, header=None, dtype=str)
        for df in sheets.values():
            for value in df.to_numpy().flatten():
                if pd.isna(value):
                    continue
                add_text(value)

    elif ext == ".docx":
        if Document is None:
            raise RuntimeError("python-docx is required for DOCX layout parsing.")
        doc = Document(layout_path)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    add_text(cell.text)

    else:
        raise ValueError("Unsupported plate layout file type. Use CSV, XLS, XLSX, or DOCX.")

    if not marker_numbers:
        return ""

    def score(prefix: str):
        upper = prefix.upper()
        sample_bonus = 100000 if "SAMPLE" in upper else 0
        specimen_bonus = 50000 if "SPECIMEN" in upper else 0
        unique_count = len(marker_numbers[prefix])
        occurrences = marker_occurrences[prefix]
        contiguous_from_one = 0
        expected = 1

        for number in sorted(marker_numbers[prefix]):
            if number == expected:
                contiguous_from_one += 1
                expected += 1
            elif number > expected:
                break

        return (
            sample_bonus + specimen_bonus + contiguous_from_one * 1000 + unique_count * 100 + occurrences,
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

    marker = normalize_marker_prefix(marker_pattern if isinstance(marker_pattern, str) else "")
    if not marker:
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
                    "plate_name": str(plate_name or f"Plate {plate_index}"),
                })
'''


NEW_PARSE_LAYOUT = r'''
def parse_plate_layout_file(layout_path: str, marker: str = "") -> Dict[str, List[Dict[str, Any]]]:
    if pd is None:
        raise RuntimeError("pandas is required for layout parsing.")

    marker = normalize_marker_prefix(marker)
    if not marker:
        marker = detect_plate_layout_marker(layout_path)

    if not marker:
        raise ValueError("DevPal Lite could not automatically identify a numbered plate-layout marker.")

    plate_layout = defaultdict(list)
    ext = os.path.splitext(layout_path)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(layout_path, header=None, dtype=str)
        parse_dataframe_plate_layout(df, marker, plate_layout, plate_index=1, plate_name="Plate 1")

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

                    plate_col_match = re.search(r"\d+", str(header_cells[col_idx].text).strip())
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


NEW_CROSS_MATCH = r'''
def cross_match_positions(
    aspirate_steps: List[str],
    dispense_steps: List[str],
    replicate_value: str,
    ordered_steps: Optional[List[Dict[str, Any]]] = None,
) -> Tuple[bool, Any]:
    replicate_count = REPLICATE_MAP.get(str(replicate_value).lower(), 1)

    logical_groups = {}
    logical_group_order = []
    current_group_by_channel = {}

    def get_or_create_group(aspirate):
        source_identity = (
            str(aspirate[1]).strip(),
            normalize_position_text(aspirate[2]),
        )

        if source_identity not in logical_groups:
            sample_number = len(logical_group_order) + 1
            group_key = (
                aspirate[0],
                aspirate[1],
                aspirate[2],
                aspirate[3],
                sample_number,
            )
            logical_groups[source_identity] = {
                "key": group_key,
                "dispenses": [],
            }
            logical_group_order.append(source_identity)

        return logical_groups[source_identity]

    if ordered_steps:
        for ordered_step in ordered_steps:
            step_type = str(ordered_step.get("type", "")).lower()
            raw = str(ordered_step.get("raw", ""))

            if "aspirate" in step_type:
                for aspirate in normalize_channel_steps([raw]):
                    group = get_or_create_group(aspirate)
                    current_group_by_channel[aspirate[0]] = group

            elif "dispense" in step_type:
                for dispense in normalize_channel_steps([raw]):
                    channel = dispense[0]
                    group = current_group_by_channel.get(channel)
                    if group is None:
                        return False, (
                            f"Channel {channel} in dispense step does not match any preceding aspirate step."
                        )
                    group["dispenses"].append(dispense[1:])
    else:
        normalized_aspirates = normalize_channel_steps(aspirate_steps)
        normalized_dispenses = normalize_channel_steps(dispense_steps)
        channel_groups = defaultdict(list)

        for aspirate in normalized_aspirates:
            group = get_or_create_group(aspirate)
            if group not in channel_groups[aspirate[0]]:
                channel_groups[aspirate[0]].append(group)

        channel_cursor = defaultdict(int)

        for dispense in normalized_dispenses:
            channel = dispense[0]
            groups = channel_groups.get(channel, [])
            if not groups:
                return False, f"Channel {channel} in dispense step does not match any aspirate step."

            cursor = min(channel_cursor[channel], len(groups) - 1)
            group = groups[cursor]
            group["dispenses"].append(dispense[1:])

            if (
                replicate_count > 0
                and len(group["dispenses"]) >= replicate_count
                and channel_cursor[channel] < len(groups) - 1
            ):
                channel_cursor[channel] += 1

    result = {}

    for source_identity in logical_group_order:
        group = logical_groups[source_identity]
        dispenses = group["dispenses"]

        if not dispenses:
            return False, (
                f"Source sample {source_identity[0]} {source_identity[1]} had no matching dispense step."
            )

        if replicate_count > 0 and len(dispenses) % replicate_count != 0:
            return False, (
                f"Source sample {source_identity[0]} {source_identity[1]} has {len(dispenses)} dispenses, "
                f"which is not divisible by replicate count {replicate_count}."
            )

        result[group["key"]] = dispenses

    return True, result
'''


def apply_patch(source: str) -> str:
    if "def detect_plate_layout_marker(" not in source:
        source = insert_before_top_level_function(source, "parse_dataframe_plate_layout", NEW_MARKER_HELPERS)

    source = replace_top_level_function(source, "parse_dataframe_plate_layout", NEW_PARSE_DATAFRAME)
    source = replace_top_level_function(source, "parse_plate_layout_file", NEW_PARSE_LAYOUT)
    source = replace_top_level_function(source, "cross_match_positions", NEW_CROSS_MATCH)

    upload_old = '        if path:\n            self.layout_path = path\n            self.layout_label.config(text=path)\n'
    upload_new = '''        if path:\n            self.layout_path = path\n            self.layout_label.config(text=path)\n\n            if not self.marker_var.get().strip():\n                try:\n                    detected_marker = detect_plate_layout_marker(path)\n                    if detected_marker:\n                        self.marker_var.set(detected_marker)\n                        self.status.config(text=f"Plate layout loaded. Auto-detected marker: {detected_marker}")\n                except Exception:\n                    pass\n'''

    if "Auto-detected marker:" not in source:
        if upload_old not in source:
            raise RuntimeError("Could not find upload_layout method block.")
        source = source.replace(upload_old, upload_new, 1)

    marker_old = '        marker = self.marker_var.get().strip().upper()\n'
    marker_new = '''        marker = normalize_marker_prefix(self.marker_var.get())\n\n        if not marker and self.layout_path:\n            try:\n                marker = detect_plate_layout_marker(self.layout_path)\n            except Exception as exc:\n                messagebox.showerror(\n                    "Plate Layout Marker Error",\n                    f"Could not automatically identify the plate-layout marker: {exc}",\n                )\n                return None\n\n            if marker:\n                self.marker_var.set(marker)\n\n        if not marker:\n            messagebox.showerror(\n                "Missing Plate Layout Marker",\n                "DevPal Lite could not automatically identify a numbered marker in the plate layout.",\n            )\n            return None\n'''

    if "Could not automatically identify a numbered marker in the plate layout." not in source:
        if marker_old not in source:
            raise RuntimeError("Could not find collect_configs marker assignment.")
        source = source.replace(marker_old, marker_new, 1)

    old_cross = '''                ok, cross_match = cross_match_positions(\n                    aspirate_steps,\n                    dispense_steps,\n                    configs[0].replicate,\n                )'''
    new_cross = '''                ok, cross_match = cross_match_positions(\n                    aspirate_steps,\n                    dispense_steps,\n                    configs[0].replicate,\n                    ordered_steps=ordered_steps,\n                )'''

    if old_cross in source:
        source = source.replace(old_cross, new_cross, 1)

    return source


def main():
    candidate = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("devpal_lite_runner_updated.py")

    if not candidate.exists():
        alternatives = [Path("devpal_lite_runner.py"), Path("t3_updated.py"), Path("t3")]
        candidate = next((path for path in alternatives if path.exists()), candidate)

    if not candidate.exists():
        raise SystemExit(
            "Could not find the DevPal Lite source file. Run this updater from the project directory or pass the source path."
        )

    original = candidate.read_text(encoding="utf-8")
    updated = apply_patch(original)

    backup = candidate.with_name(candidate.name + ".before_multiplate_fix_v2.bak")
    shutil.copy2(candidate, backup)

    output = candidate.with_name(candidate.stem + "_multiplate_fix_v2.py")
    output.write_text(updated, encoding="utf-8")
    compile(updated, str(output), "exec")

    print(f"Corrected file created: {output}")
    print(f"Original preserved: {candidate}")
    print(f"Backup created: {backup}")
    print("Fixes marker auto-detection and logical-sample grouping for repeated aspirates / multi-plate runs.")


if __name__ == "__main__":
    main()
