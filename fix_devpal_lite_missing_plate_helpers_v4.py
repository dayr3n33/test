
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path


HELPERS = r'''
def get_layout_plate_indices(
    layout_matches: Dict[str, List[Dict[str, Any]]]
) -> List[int]:
    plate_indices = set()

    for entries in (layout_matches or {}).values():
        for entry in entries or []:
            try:
                plate_indices.add(int(entry.get("plate_index", 1)))
            except Exception:
                plate_indices.add(1)

    return sorted(plate_indices) or [1]


def get_plate_marker_numbers(
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker: str,
    plate_index: int,
) -> List[int]:
    marker = str(marker or "").strip().upper()
    numbers = set()

    for marker_label, entries in (layout_matches or {}).items():
        match = re.fullmatch(
            rf"{re.escape(marker)}(\d+)",
            str(marker_label or "").strip().upper(),
        )
        if not match:
            continue

        number = int(match.group(1))

        for entry in entries or []:
            try:
                entry_plate = int(entry.get("plate_index", 1))
            except Exception:
                entry_plate = 1

            if entry_plate == int(plate_index):
                numbers.add(number)
                break

    return sorted(numbers)


def resolve_layout_marker_for_sample(
    source_number: int,
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker: str,
) -> Tuple[str, int, int]:
    # Map sequential source/sample number to:
    # (layout marker label, logical destination plate index, local sample number).
    # One uploaded sheet/table is treated as a reusable plate template.
    marker = str(marker or "").strip().upper()
    source_number = int(source_number)

    if source_number < 1:
        raise ValueError("source_number must be 1 or greater.")

    plate_indices = get_layout_plate_indices(layout_matches)

    if len(plate_indices) == 1:
        template_plate = plate_indices[0]
        numbers = get_plate_marker_numbers(
            layout_matches,
            marker,
            template_plate,
        )

        if not numbers:
            raise ValueError(
                f"No numbered markers beginning with {marker!r} were found "
                f"in the uploaded plate layout."
            )

        samples_per_plate = len(numbers)
        logical_plate_index = ((source_number - 1) // samples_per_plate) + 1
        local_offset = (source_number - 1) % samples_per_plate
        local_marker_number = numbers[local_offset]

        return (
            f"{marker}{local_marker_number}",
            logical_plate_index,
            local_marker_number,
        )

    remaining = source_number

    for logical_plate_index, physical_plate_index in enumerate(
        plate_indices,
        start=1,
    ):
        numbers = get_plate_marker_numbers(
            layout_matches,
            marker,
            physical_plate_index,
        )

        if not numbers:
            continue

        if source_number in numbers:
            return (
                f"{marker}{source_number}",
                logical_plate_index,
                source_number,
            )

        if remaining <= len(numbers):
            local_marker_number = numbers[remaining - 1]
            return (
                f"{marker}{local_marker_number}",
                logical_plate_index,
                local_marker_number,
            )

        remaining -= len(numbers)

    final_physical_plate = plate_indices[-1]
    final_numbers = get_plate_marker_numbers(
        layout_matches,
        marker,
        final_physical_plate,
    )

    if not final_numbers:
        raise ValueError(
            f"No numbered markers beginning with {marker!r} were found "
            f"in the final uploaded plate layout."
        )

    samples_per_plate = len(final_numbers)
    extra_source_index = remaining - 1
    extra_plate_offset = extra_source_index // samples_per_plate
    local_offset = extra_source_index % samples_per_plate

    logical_plate_index = len(plate_indices) + extra_plate_offset
    local_marker_number = final_numbers[local_offset]

    return (
        f"{marker}{local_marker_number}",
        logical_plate_index,
        local_marker_number,
    )


def get_layout_marker_positions(
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker_label: str,
    plate_index: Optional[int] = None,
) -> List[str]:
    entries = list((layout_matches or {}).get(marker_label, []) or [])

    if not entries:
        return []

    plate_indices = get_layout_plate_indices(layout_matches)

    if plate_index is not None:
        if len(plate_indices) == 1:
            physical_plate_index = plate_indices[0]
        elif int(plate_index) <= len(plate_indices):
            physical_plate_index = plate_indices[int(plate_index) - 1]
        else:
            physical_plate_index = plate_indices[-1]

        filtered = []

        for entry in entries:
            try:
                entry_plate = int(entry.get("plate_index", 1))
            except Exception:
                entry_plate = 1

            if entry_plate == physical_plate_index:
                filtered.append(entry)

        entries = filtered

    positions = []
    seen = set()

    for entry in entries:
        position = normalize_position_text(entry.get("dest_value", ""))

        if position and position not in seen:
            seen.add(position)
            positions.append(position)

    return positions
'''


def find_function_span(source: str, name: str):
    pattern = re.compile(rf"(?m)^def {re.escape(name)}\s*\(")
    match = pattern.search(source)
    if not match:
        return None

    start = match.start()
    next_match = re.search(r"(?m)^def \w+\s*\(|^class \w+", source[match.end():])

    if next_match:
        end = match.end() + next_match.start()
    else:
        end = len(source)

    return start, end


def remove_function_if_present(source: str, name: str) -> str:
    span = find_function_span(source, name)
    if not span:
        return source

    start, end = span
    return source[:start] + source[end:]


def insert_before_function(source: str, before_name: str, block: str) -> str:
    match = re.search(rf"(?m)^def {re.escape(before_name)}\s*\(", source)

    if not match:
        raise RuntimeError(
            f"Could not find def {before_name}(...) to insert helpers before."
        )

    return source[:match.start()] + block.strip() + "\n\n" + source[match.start():]


def apply_fix(source: str) -> str:
    for name in (
        "get_layout_plate_indices",
        "get_plate_marker_numbers",
        "resolve_layout_marker_for_sample",
        "get_layout_marker_positions",
    ):
        source = remove_function_if_present(source, name)

    source = insert_before_function(
        source,
        "validate_trace_against_layout",
        HELPERS,
    )

    required = (
        "def get_layout_plate_indices(",
        "def get_plate_marker_numbers(",
        "def resolve_layout_marker_for_sample(",
        "def get_layout_marker_positions(",
    )

    missing = [item for item in required if item not in source]
    if missing:
        raise RuntimeError(
            "Fix did not insert all required helpers: " + ", ".join(missing)
        )

    return source


def main():
    if len(sys.argv) > 1:
        source_path = Path(sys.argv[1])
    else:
        candidates = [
            Path("devpal_lite_runner_updated_multiplate_fix_v3.py"),
            Path("devpal_lite_runner_updated_multiplate_fix_v2_multiplate_fix_v3.py"),
            Path("devpal_lite_runner_updated_multiplate_fix_v2.py"),
            Path("devpal_lite_runner_updated.py"),
            Path("devpal_lite_runner.py"),
            Path("t3"),
        ]
        source_path = next((p for p in candidates if p.exists()), None)

    if source_path is None or not source_path.exists():
        raise SystemExit(
            "DevPal Lite source file not found. Pass the Python file you are currently running."
        )

    original = source_path.read_text(encoding="utf-8")
    updated = apply_fix(original)

    compile(updated, str(source_path), "exec")

    backup_path = source_path.with_name(
        source_path.name + ".before_plate_helper_v4.bak"
    )
    shutil.copy2(source_path, backup_path)

    output_path = source_path.with_name(
        source_path.stem + "_plate_helper_v4.py"
    )
    output_path.write_text(updated, encoding="utf-8")

    print(f"Created: {output_path}")
    print(f"Backup:  {backup_path}")
    print("")
    print("Verified in generated DevPal source:")
    print("- get_layout_plate_indices")
    print("- get_plate_marker_numbers")
    print("- resolve_layout_marker_for_sample")
    print("- get_layout_marker_positions")
    print("- full generated source compiles successfully")


if __name__ == "__main__":
    main()
