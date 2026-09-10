
from __future__ import annotations

import ast
import re
import shutil
import sys
from pathlib import Path


TARGET_DEFAULT = "devpal_lite_runner.py"
OUTPUT_DEFAULT = "devpal_lite_runner_fixed.py"


def replace_top_level_function(source: str, function_name: str, replacement: str) -> str:
    pattern = re.compile(
        rf"(?ms)^def {re.escape(function_name)}\(.*?(?=^def |^class |\Z)"
    )
    match = pattern.search(source)
    if not match:
        raise RuntimeError(f"Could not find top-level function: {function_name}")
    return source[:match.start()] + replacement.rstrip() + "\n\n" + source[match.end():]


def require_once(source: str, old: str, description: str) -> None:
    count = source.count(old)
    if count != 1:
        raise RuntimeError(
            f"Expected exactly one {description}, found {count}. "
            "This patch is intentionally strict so it cannot silently alter the wrong version."
        )


def main():
    target = Path(sys.argv[1] if len(sys.argv) > 1 else TARGET_DEFAULT)
    output = Path(sys.argv[2] if len(sys.argv) > 2 else OUTPUT_DEFAULT)

    if not target.exists():
        raise FileNotFoundError(
            f"{target} was not found. Run this from the PyCharm project directory "
            "containing the ORIGINAL working devpal_lite_runner.py."
        )

    source = target.read_text(encoding="utf-8")

    # Refuse to patch the bad rewritten branch. This must start from the original DP.
    if 'TIP_EJECT_ROLE = "tip eject"' in source or "class PlateLayoutPattern" in source:
        raise RuntimeError(
            "This input already contains the rewritten update. "
            "Use the original working DP-based devpal_lite_runner.py instead."
        )

    required_markers = [
        'ROLE_OPTIONS = ["tip pick up", "aspirate", "dispense"]',
        "def build_execution_plan(",
        "def cross_match_positions(",
        "def validate_trace_against_layout(",
        '("What Was Run", sections["what_was_run"])',
        "require_after_start=False",
        'text="Variant"',
    ]
    missing = [marker for marker in required_markers if marker not in source]
    if missing:
        raise RuntimeError(
            "Input does not look like the original working DP version. "
            f"Missing baseline markers: {missing}"
        )

    # 1) Optional role constants. ROLE_OPTIONS itself remains unchanged.
    old = 'ROLE_OPTIONS = ["tip pick up", "aspirate", "dispense"]\n'
    require_once(source, old, "ROLE_OPTIONS definition")
    source = source.replace(
        old,
        old
        + 'TIP_EJECT_ROLE = "tip eject"\n'
        + 'ROLE_OPTIONS_WITH_TIP_EJECT = ROLE_OPTIONS + [TIP_EJECT_ROLE]\n',
        1,
    )

    # 2) Optional sequence-aware 96-head eject.
    source = replace_top_level_function(
        source,
        "ph_tip_eject_96",
        r'''def ph_tip_eject_96(ham_int, tip_eject_sequence: str = ""):
    tip_eject_sequence = str(tip_eject_sequence or "").strip()

    func = _get_optional_pyhamilton_function(
        "tip_eject_96",
        "tip_eject_96_seq",
        "tip_eject96",
    )

    if func:
        if tip_eject_sequence:
            return _call_best_effort(
                func,
                ham_int,
                sequence=tip_eject_sequence,
                waste_seq=tip_eject_sequence,
                tip_seq=tip_eject_sequence,
                tip96_seq=tip_eject_sequence,
            )
        return _call_best_effort(func, ham_int)

    if tip_eject_sequence:
        return _send_command_best_effort(
            ham_int,
            "tip_eject_96",
            sequence=tip_eject_sequence,
            waste_seq=tip_eject_sequence,
            tip_seq=tip_eject_sequence,
            tip96_seq=tip_eject_sequence,
        )

    return _send_command_best_effort(ham_int, "tip_eject_96")''',
    )

    # 3) Keep the same three required roles, but allow optional tip eject.
    source = replace_top_level_function(
        source,
        "validate_step_configs",
        r'''def validate_step_configs(configs: List[SequenceStepConfig]) -> List[ValidationIssue]:
    issues = []
    roles = {config.role for config in configs}

    for required_role in ROLE_OPTIONS:
        if required_role not in roles:
            issues.append(ValidationIssue(
                "critical",
                "configuration",
                f"Missing required step role: {required_role}",
            ))

    for config in configs:
        if config.role not in ROLE_OPTIONS_WITH_TIP_EJECT:
            issues.append(ValidationIssue(
                "critical",
                "configuration",
                f"Invalid step role for {config.sequence}: {config.role}",
            ))

        if config.role in ("aspirate", "dispense") and not config.liquid_class:
            issues.append(ValidationIssue(
                "critical",
                "configuration",
                f"Liquid class is required for {config.role}: {config.sequence}",
            ))

        if config.volume_ul <= 0 or config.volume_ul > 1000:
            issues.append(ValidationIssue(
                "critical",
                "configuration",
                f"Volume must be between 1 and 1000 uL: {config.sequence}",
            ))

        if config.transfer_type not in TRANSFER_OPTIONS:
            issues.append(ValidationIssue(
                "critical",
                "configuration",
                f"Invalid transfer type: {config.transfer_type}",
            ))

        if config.replicate not in REPLICATE_OPTIONS:
            issues.append(ValidationIssue(
                "critical",
                "configuration",
                f"Invalid replicate value: {config.replicate}",
            ))

        if config.transfer_type == "samples" and not config.marker:
            issues.append(ValidationIssue(
                "critical",
                "configuration",
                f"Marker is required for samples: {config.sequence}",
            ))

    return issues''',
    )

    # 4) Let a control-designated unique eject drive iterations if needed.
    old = 'and config.role in ("tip pick up", "aspirate", "dispense")'
    require_once(source, old, "control-role filter")
    source = source.replace(
        old,
        'and config.role in ("tip pick up", "aspirate", "dispense", TIP_EJECT_ROLE)',
        1,
    )

    # 5) Original execution-plan logic with only optional eject substitution added.
    source = replace_top_level_function(
        source,
        "build_execution_plan",
        r'''def build_execution_plan(
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

    def append_plan_step(
        config: SequenceStepConfig,
        set_order: int,
        rep_index: int,
        transfer_iteration: int,
    ):
        nonlocal order

        batch_size = get_sequence_batch_size(config)
        if config.role == "tip pick up":
            action = "tip_pick"
        elif config.role == TIP_EJECT_ROLE:
            action = "tip_eject"
        else:
            action = config.role

        config_iterations = get_sequence_iterations(config)

        should_increment = (
            transfer_iteration > 1
            and transfer_iteration <= config_iterations
            and not bool(getattr(config, "manual", False))
            and config.role in ("aspirate", "dispense")
        )

        seq_meta = get_sequence_labware_metadata(
            config.sequence,
            lay_metadata,
            getattr(config, "sequence_count", 0),
        )

        labware_names = seq_meta.get(
            "labware_names",
            [seq_meta.get("labware_name", "")]
            if seq_meta.get("labware_name") else [],
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
            "labware_names": labware_names,
            "labware_count": len(labware_names),
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
            "unique_tip_eject": config.role == TIP_EJECT_ROLE,
        })

        order += 1

    def append_default_tip_eject(
        tip: SequenceStepConfig,
        set_order: int,
        rep_index: int,
        transfer_iteration: int,
    ):
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
            "labware_names": ["Waste"],
            "labware_count": 1,
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
            "unique_tip_eject": False,
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
        eject_steps = [c for c in set_configs if c.role == TIP_EJECT_ROLE]

        transfer_iterations = get_required_transfer_iterations(set_configs)

        for rep_index in range(1, runtime_replicate_count + 1):
            report_replicate_index = (
                rep_index if loop_by_replicate else configured_replicate_count
            )

            for transfer_iteration in range(1, transfer_iterations + 1):

                if variant_mode:
                    if not tip_steps or not asp_steps or not dsp_steps:
                        continue

                    for dispense_index, dispense_config in enumerate(dsp_steps):
                        tip_config = tip_steps[dispense_index % len(tip_steps)]
                        aspirate_config = asp_steps[dispense_index % len(asp_steps)]

                        append_plan_step(
                            tip_config,
                            set_order,
                            report_replicate_index,
                            transfer_iteration,
                        )

                        append_plan_step(
                            aspirate_config,
                            set_order,
                            report_replicate_index,
                            transfer_iteration,
                        )

                        append_plan_step(
                            dispense_config,
                            set_order,
                            report_replicate_index,
                            transfer_iteration,
                        )

                        if eject_steps:
                            append_plan_step(
                                eject_steps[dispense_index % len(eject_steps)],
                                set_order,
                                report_replicate_index,
                                transfer_iteration,
                            )
                        else:
                            append_default_tip_eject(
                                tip_config,
                                set_order,
                                report_replicate_index,
                                transfer_iteration,
                            )

                    continue

                max_len = max(len(tip_steps), len(asp_steps), len(dsp_steps), 1)

                for idx in range(max_len):
                    tip_config = (
                        tip_steps[idx]
                        if idx < len(tip_steps)
                        else tip_steps[0] if tip_steps else None
                    )
                    aspirate_config = (
                        asp_steps[idx]
                        if idx < len(asp_steps)
                        else asp_steps[0] if asp_steps else None
                    )
                    dispense_config = (
                        dsp_steps[idx]
                        if idx < len(dsp_steps)
                        else dsp_steps[0] if dsp_steps else None
                    )

                    for config in (tip_config, aspirate_config, dispense_config):
                        if config is not None:
                            append_plan_step(
                                config,
                                set_order,
                                report_replicate_index,
                                transfer_iteration,
                            )

                    if tip_config is not None:
                        if eject_steps:
                            eject_config = (
                                eject_steps[idx]
                                if idx < len(eject_steps)
                                else eject_steps[0]
                            )
                            append_plan_step(
                                eject_config,
                                set_order,
                                report_replicate_index,
                                transfer_iteration,
                            )
                        else:
                            append_default_tip_eject(
                                tip_config,
                                set_order,
                                report_replicate_index,
                                transfer_iteration,
                            )

    return plan''',
    )

    # 6) Preserve original LAY metadata behavior and simply add all referenced labware.
    old = '''        metadata[seq_name] = {
            "sequence_id": seq_id,
            "sequence_name": seq_name,
            "selected_position_count": selected_position_count,
            "position_count": selected_position_count,
            "positions": [normalize_position_text(position) for position in positions],
            "labware_name": labware_name,
            "labware_file": labware_file,
            "template": template,
            "site_id": site_id,
            "labware_format": labware_format,
            "labware_format_basis": "sequence_item_objid_to_labware_id_file_template",
        }
'''
    require_once(source, old, "LAY metadata dictionary")
    source = source.replace(
        old,
        '''        ordered_labware_names = []
        for objid in clean_objids:
            resolved_name = labware_by_id.get(objid.lower(), {}).get(
                "labware_name",
                objid,
            )
            if resolved_name and resolved_name not in ordered_labware_names:
                ordered_labware_names.append(resolved_name)

        metadata[seq_name] = {
            "sequence_id": seq_id,
            "sequence_name": seq_name,
            "selected_position_count": selected_position_count,
            "position_count": selected_position_count,
            "positions": [normalize_position_text(position) for position in positions],
            "labware_name": labware_name,
            "labware_file": labware_file,
            "template": template,
            "site_id": site_id,
            "labware_format": labware_format,
            "labware_format_basis": "sequence_item_objid_to_labware_id_file_template",
            "labware_names": ordered_labware_names,
            "labware_count": len(ordered_labware_names),
        }
''',
        1,
    )

    # 7) Keep original crossmatch for ordinary runs; only use new branch when channels repeat.
    source = replace_top_level_function(
        source,
        "cross_match_positions",
        r'''def cross_match_positions(
    aspirate_steps: List[str],
    dispense_steps: List[str],
    replicate_value: str,
) -> Tuple[bool, Any]:
    replicate_count = REPLICATE_MAP.get(replicate_value, 1)

    normalized_aspirates = normalize_channel_steps(aspirate_steps)
    normalized_dispenses = normalize_channel_steps(dispense_steps)

    aspirates_by_channel = defaultdict(list)
    for aspirate_index, asp in enumerate(normalized_aspirates):
        aspirates_by_channel[asp[0]].append((aspirate_index, asp))

    channel_reuse_exists = any(
        len(items) > 1
        for items in aspirates_by_channel.values()
    )

    if not channel_reuse_exists:
        # Original DP behavior.
        aspirate_channel_map = {}
        for asp in normalized_aspirates:
            aspirate_channel_map[asp[0]] = asp

        aspirate_to_dispense_map = defaultdict(list)
        for dsp in normalized_dispenses:
            channel = dsp[0]
            if channel not in aspirate_channel_map:
                return False, (
                    f"Dispense channel {channel} did not match any aspirate channel."
                )
            asp = aspirate_channel_map[channel]
            aspirate_to_dispense_map[asp].append(dsp[1:])

        for asp, dispenses in aspirate_to_dispense_map.items():
            if len(dispenses) % replicate_count != 0:
                return False, (
                    f"Aspirate {asp} has {len(dispenses)} dispense events, "
                    f"which does not match replicate setting '{replicate_value}'."
                )

        return True, dict(aspirate_to_dispense_map)

    # New branch only for later batches that reuse physical channels.
    dispenses_by_channel = defaultdict(list)
    for dsp in normalized_dispenses:
        dispenses_by_channel[dsp[0]].append(dsp)

    unknown_channels = sorted(
        set(dispenses_by_channel) - set(aspirates_by_channel)
    )
    if unknown_channels:
        return False, (
            "Dispense channel(s) did not match any aspirate channel: "
            + ", ".join(unknown_channels)
        )

    chronological_records = []

    for channel, aspirate_occurrences in aspirates_by_channel.items():
        channel_dispenses = dispenses_by_channel.get(channel, [])
        expected_dispenses = len(aspirate_occurrences) * replicate_count

        if len(channel_dispenses) != expected_dispenses:
            return False, (
                f"Channel {channel} had {len(aspirate_occurrences)} aspirate "
                f"occurrence(s) and {len(channel_dispenses)} dispense event(s); "
                f"{expected_dispenses} are required for replicate setting "
                f"'{replicate_value}'."
            )

        for occurrence_index, (global_aspirate_index, asp) in enumerate(
            aspirate_occurrences
        ):
            start = occurrence_index * replicate_count
            stop = start + replicate_count
            assigned_dispenses = [
                dsp[1:]
                for dsp in channel_dispenses[start:stop]
            ]
            chronological_records.append(
                (global_aspirate_index, asp, assigned_dispenses)
            )

    chronological_records.sort(key=lambda item: item[0])

    aspirate_to_dispense_map = {}
    for sample_order, (_, asp, assigned_dispenses) in enumerate(
        chronological_records,
        start=1,
    ):
        unique_aspirate_key = asp + (f"sample_order={sample_order}",)
        aspirate_to_dispense_map[unique_aspirate_key] = assigned_dispenses

    return True, aspirate_to_dispense_map''',
    )

    # 8) Keep original single-plate issue logic; add plate identity only when samples exceed template.
    source = replace_top_level_function(
        source,
        "validate_trace_against_layout",
        r'''def validate_trace_against_layout(
    aspirate_to_dispense_map: Dict[Any, Any],
    layout_matches: Dict[str, List[Dict[str, Any]]],
    marker: str,
) -> Tuple[List[Dict[str, Any]], List[ValidationIssue]]:
    findings = []
    issues = []
    marker = marker.upper().strip()

    marker_numbers = get_marker_numbers(layout_matches, marker)
    samples_per_plate = max(marker_numbers) if marker_numbers else 0
    total_samples = len(aspirate_to_dispense_map)
    multi_plate_mode = (
        samples_per_plate > 0
        and total_samples > samples_per_plate
    )

    seen_trace_destinations = []
    trace_transfer_by_marker = defaultdict(list)
    mismatch_by_marker = {}
    replicate_mismatch_by_marker = {}

    plate_cycle_to_labware = {}
    labware_to_plate_cycle = {}
    plate_continuity_failures = []

    for sample_index, (aspirate_step, dispenses) in enumerate(
        aspirate_to_dispense_map.items(),
        start=1,
    ):
        marker_label = f"{marker}{sample_index}"

        layout_marker, repeated_layout_number = resolve_repeating_layout_marker(
            sample_index,
            layout_matches,
            marker,
        )

        source_labware = str(aspirate_step[1]).strip()
        source_position = normalize_position_text(aspirate_step[2])

        expected_positions = get_layout_marker_positions(layout_matches, layout_marker)
        actual_positions = [
            normalize_position_text(dispense[1])
            for dispense in dispenses
        ]
        actual_labware = sorted({
            str(dispense[0]).strip()
            for dispense in dispenses
            if dispense
        })

        expected_set = set(expected_positions)
        actual_set = set(actual_positions)

        missing_for_this_sample = sorted(expected_set - actual_set)
        wrong_for_this_sample = sorted(actual_set - expected_set)

        if multi_plate_mode:
            for dispense in dispenses:
                seen_trace_destinations.append((
                    str(dispense[0]).strip(),
                    normalize_position_text(dispense[1]),
                ))

            expected_plate_cycle = (
                (sample_index - 1) // samples_per_plate
            ) + 1

            if len(actual_labware) != 1:
                plate_continuity_failures.append({
                    "sample": marker_label,
                    "plate_cycle": expected_plate_cycle,
                    "message": (
                        f"{marker_label}: expected all destination wells for "
                        f"plate cycle {expected_plate_cycle} to be on one physical "
                        f"labware, but trace used {actual_labware}."
                    ),
                    "actual_labware": actual_labware,
                })
            else:
                sample_labware = actual_labware[0]
                assigned_labware = plate_cycle_to_labware.get(
                    expected_plate_cycle
                )

                if assigned_labware is None:
                    previous_cycle = labware_to_plate_cycle.get(sample_labware)

                    if (
                        previous_cycle is not None
                        and previous_cycle != expected_plate_cycle
                    ):
                        plate_continuity_failures.append({
                            "sample": marker_label,
                            "plate_cycle": expected_plate_cycle,
                            "message": (
                                f"{marker_label}: plate cycle "
                                f"{expected_plate_cycle} reused labware "
                                f"'{sample_labware}' that was already used for "
                                f"plate cycle {previous_cycle}."
                            ),
                            "actual_labware": actual_labware,
                        })
                    else:
                        plate_cycle_to_labware[
                            expected_plate_cycle
                        ] = sample_labware
                        labware_to_plate_cycle[
                            sample_labware
                        ] = expected_plate_cycle

                elif sample_labware != assigned_labware:
                    plate_continuity_failures.append({
                        "sample": marker_label,
                        "plate_cycle": expected_plate_cycle,
                        "message": (
                            f"{marker_label}: plate cycle "
                            f"{expected_plate_cycle} should continue on "
                            f"'{assigned_labware}', but trace used "
                            f"'{sample_labware}'."
                        ),
                        "actual_labware": actual_labware,
                    })
        else:
            # Exactly the old behavior for a normal single-plate run.
            seen_trace_destinations.extend(actual_positions)

        trace_transfer_by_marker[marker_label].append({
            "source_labware": source_labware,
            "source_position": source_position,
            "destination_labware": actual_labware,
            "destination_wells": actual_positions,
        })

        status = "pass"
        if (
            missing_for_this_sample
            or wrong_for_this_sample
            or len(actual_positions) != len(expected_positions)
        ):
            status = "fail"

        findings.append({
            "type": "trace_layout_match",
            "marker": marker_label,
            "plate_layout_marker_used": layout_marker,
            "source_labware": source_labware,
            "source_position": source_position,
            "expected_destinations": expected_positions,
            "actual_destinations": actual_positions,
            "actual_destination_labware": actual_labware,
            "missing_destinations": missing_for_this_sample,
            "unexpected_destinations": wrong_for_this_sample,
            "status": status,
        })

        if missing_for_this_sample or wrong_for_this_sample:
            mismatch_by_marker[marker_label] = {
                "expected_positions": expected_positions,
                "actual_positions": actual_positions,
                "missing_for_this_sample": missing_for_this_sample,
                "wrong_for_this_sample": wrong_for_this_sample,
            }

        if len(actual_positions) != len(expected_positions):
            replicate_mismatch_by_marker[marker_label] = {
                "plate_layout_marker_used": layout_marker,
                "expected_replicate_count_from_plate_layout": len(expected_positions),
                "actual_dispense_count_from_trace": len(actual_positions),
                "expected_destination_wells": expected_positions,
                "trace_destination_wells_for_this_sample": actual_positions,
            }

    for marker_label, transfers in trace_transfer_by_marker.items():
        findings.append({
            "type": "actual_trace_transfer",
            "marker": marker_label,
            "transfers": transfers,
        })

    # Original issue creation is preserved.
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
                "missing_correct_wells_for_this_sample": data[
                    "missing_for_this_sample"
                ],
                "wrong_wells_used_for_this_sample": data[
                    "wrong_for_this_sample"
                ],
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

    if multi_plate_mode:
        duplicate_trace_dests = sorted({
            f"{labware}:{position}"
            for labware, position in set(seen_trace_destinations)
            if seen_trace_destinations.count((labware, position)) > 1
        })
    else:
        duplicate_trace_dests = sorted([
            pos
            for pos in set(seen_trace_destinations)
            if seen_trace_destinations.count(pos) > 1
        ])

    if duplicate_trace_dests:
        add_issue(
            issues,
            "major",
            "trace duplicate destinations",
            "Duplicate destination wells were found in the trace.",
            {"duplicate_destination_wells": duplicate_trace_dests},
        )

    # Only brand-new issue type: a later plate did not continue the same
    # S1..SN pattern on its own physical labware.
    for failure in plate_continuity_failures:
        add_issue(
            issues,
            "critical",
            "trace vs plate layout",
            failure["message"],
            {
                "sample": failure["sample"],
                "plate_cycle": failure["plate_cycle"],
                "actual_labware": failure["actual_labware"],
                "samples_per_plate": samples_per_plate,
            },
        )

    labware_findings, labware_issues = validate_trace_labware_against_layout_format(
        aspirate_to_dispense_map,
        layout_matches,
    )
    findings.extend(labware_findings)
    issues.extend(labware_issues)

    return findings, issues''',
    )

    # 9) Runtime: unique sequence only affects unique eject. Preserve old trace timing.
    old = '''                elif action == "tip_eject":
                    if mode == "head":
                        ph_tip_eject_96(ham_int)
                    else:
                        ph_tip_eject_seq2(
                            ham_int,
                            waste_seq=seq or "Waste",
                            channel=channel,
                        )
'''
    require_once(source, old, "runtime tip-eject block")
    source = source.replace(
        old,
        '''                elif action == "tip_eject":
                    if mode == "head":
                        ph_tip_eject_96(
                            ham_int,
                            tip_eject_sequence=(
                                seq
                                if item.get("unique_tip_eject")
                                else ""
                            ),
                        )
                    else:
                        ph_tip_eject_seq2(
                            ham_int,
                            waste_seq=seq or "Waste",
                            channel=channel,
                        )
''',
        1,
    )

    # 10) UI state storage.
    old = '''        self.sequence_role_vars = {}
        self.sequence_channel_vars = {}
'''
    if source.count(old) != 2:
        raise RuntimeError(
            "Expected the original sequence role-var initialization twice."
        )
    source = source.replace(
        old,
        '''        self.sequence_role_vars = {}
        self.sequence_role_widgets = {}
        self.sequence_channel_vars = {}
''',
        2,
    )

    old = '''        self.tracking_var = tk.IntVar(value=0)
        self.variant_var = tk.IntVar(value=0)
'''
    require_once(source, old, "Tracking/Variant variable block")
    source = source.replace(
        old,
        '''        self.tracking_var = tk.IntVar(value=0)
        self.variant_var = tk.IntVar(value=0)
        self.unique_tip_eject_var = tk.IntVar(value=0)
''',
        1,
    )

    # Keep original checkbox helper untouched; it is the one that worked.
    old = '''        self.make_checkbutton(options_row, text="Variant", variable=self.variant_var,
                              command=self.update_set_column_visibility).pack(side="left", padx=(0, 12))

        report_box = tk.LabelFrame(options_row, text="Report Options", bg=UI_BG, fg=UI_TEXT_FG, bd=3, relief="ridge")
'''
    require_once(source, old, "Variant checkbox block")
    source = source.replace(
        old,
        '''        self.make_checkbutton(options_row, text="Variant", variable=self.variant_var,
                              command=self.update_set_column_visibility).pack(side="left", padx=(0, 8))

        self.make_checkbutton(
            options_row,
            text="Unique Tip Eject",
            variable=self.unique_tip_eject_var,
            command=self.update_unique_tip_eject_role_options,
        ).pack(side="left", padx=(0, 12))

        report_box = tk.LabelFrame(options_row, text="Report Options", bg=UI_BG, fg=UI_TEXT_FG, bd=3, relief="ridge")
''',
        1,
    )

    old = '''            role_combo = ttk.Combobox(
                sequence_frame,
                textvariable=role_var,
                values=ROLE_OPTIONS,
                width=12,
                state="readonly",
            )
            role_combo.grid(row=row_index, column=4, padx=3, pady=1, sticky="w")
'''
    require_once(source, old, "Step Role combo")
    source = source.replace(
        old,
        '''            role_combo = ttk.Combobox(
                sequence_frame,
                textvariable=role_var,
                values=self.get_step_role_options(),
                width=12,
                state="readonly",
            )
            role_combo.grid(row=row_index, column=4, padx=3, pady=1, sticky="w")
            self.sequence_role_widgets[sequence_name] = role_combo
''',
        1,
    )

    insertion_marker = "    def update_set_column_visibility(self):\n"
    require_once(source, insertion_marker, "update_set_column_visibility method")
    helper_methods = '''    def get_step_role_options(self) -> List[str]:
        if bool(self.unique_tip_eject_var.get()):
            return ROLE_OPTIONS_WITH_TIP_EJECT
        return ROLE_OPTIONS

    def update_unique_tip_eject_role_options(self):
        role_options = self.get_step_role_options()

        for sequence_name, role_combo in self.sequence_role_widgets.items():
            role_combo["values"] = role_options

            role_var = self.sequence_role_vars.get(sequence_name)
            if (
                role_var is not None
                and not self.unique_tip_eject_var.get()
                and role_var.get().strip() == TIP_EJECT_ROLE
            ):
                role_var.set("aspirate")

'''
    source = source.replace(
        insertion_marker,
        helper_methods + insertion_marker,
        1,
    )

    # 11) Generated review script: channel mode already uses seq as waste_seq.
    # Only 96-head needs the optional unique sequence argument.
    old = '                lines.append("    tip_eject_96(ham_int)")\n'
    require_once(source, old, "generated 96-head eject line")
    source = source.replace(
        old,
        '''                if item.get("unique_tip_eject"):
                    lines.append(
                        f"    tip_eject_96(ham_int, tip_eject_sequence={seq!r})"
                    )
                else:
                    lines.append("    tip_eject_96(ham_int)")
''',
        1,
    )

    # Safety guards: report and original non-blocking trace selection must remain.
    if '("What Was Run", sections["what_was_run"])' not in source:
        raise RuntimeError("Original report structure was unexpectedly altered.")

    if "require_after_start=False" not in source:
        raise RuntimeError(
            "Original trace-selection timing was unexpectedly altered."
        )

    ast.parse(source)

    if output.resolve() == target.resolve():
        backup = target.with_suffix(target.suffix + ".before_minimal_update.bak")
        shutil.copy2(target, backup)
        print(f"Backup created: {backup}")

    output.write_text(source, encoding="utf-8")

    print(f"Created: {output}")
    print("Syntax validation: PASS")
    print("Preserved: original report, original single-plate validation path, original trace timing, original checkbox code.")
    print("Added: multi-plate continuation + reused-channel branch + Unique Tip Eject.")


if __name__ == "__main__":
    main()
