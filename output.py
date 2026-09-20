"""
output.py
=========
Exports the schedule to a formatted .xlsx and prints console summaries.

Sheets:
  1. Schedule        — week-by-week ambulance grid (Sun -> Sat)
  2. Campus Response — week-by-week campus responder grid (Sun -> Sat)
  3. Hour Summary    — per-person totals (ambulance + campus) vs requirements
  4. Warnings        — unfilled shifts, ALS without EVDT, night/weekend crews
                       without a driver, under-hours volunteers
  5. Solver          — objective stages, attained values, bounds and status
"""

from __future__ import annotations

import csv
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from models import (
    CAMPUS_BLOCK_TIMES,
    HourCaps,
    Schedule,
    SHIFT_TIMES,
    Volunteer,
    crew_cap,
    is_big_weekend,
)

# ── Styling ──────────────────────────────────────────────────────────────────

C_HEADER_BG = "1F3864"
C_HEADER_FG = "FFFFFF"
C_DATE_BG = "2E5FA3"
C_DATE_FG = "FFFFFF"
C_EVDT_BG = "E2EFDA"
C_AUTH_BG = "FFF2CC"
C_EMT_BG = "FFFFFF"
C_BERT_BG = "D9E1F2"
C_WARN_BG = "FFE0E0"
C_ALT_ROW = "F8F8F8"
C_UNDER = "CC0000"

CERT_BG = {"EVDT": C_EVDT_BG, "Auth": C_AUTH_BG, "EMT": C_EMT_BG, "BERT": C_BERT_BG}

_thin = Side(style="thin", color="CCCCCC")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def _fill(hex_color):
    return PatternFill("solid", start_color=hex_color, fgColor=hex_color)


def _font(bold=False, color="000000", size=10):
    return Font(name="Arial", bold=bold, color=color, size=size)


def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def _header_row(ws, row, values, widths=None):
    for col, val in enumerate(values, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.font = _font(bold=True, color=C_HEADER_FG)
        c.fill = _fill(C_HEADER_BG)
        c.alignment = _align(h="center")
        c.border = BORDER
    if widths:
        for col, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w


def _week_start_sunday(d: date) -> date:
    return d - timedelta(days=(d.weekday() + 1) % 7)


DAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]


def _weekly_grid(ws, keys_to_people, sections, slot_rows, cell_for):
    """
    Shared week-by-week grid builder.
      sections:  [(section_key, time_label)] — one band of rows per section
      slot_rows: [slot_label, ...] — rows inside each band
      cell_for:  (people, slot_index) -> (text, fill_hex)
    """
    widths = [10, 12] + [18] * 7
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    if not keys_to_people:
        return

    all_dates = sorted({d for (d, _) in keys_to_people})
    current, end = _week_start_sunday(all_dates[0]), all_dates[-1]

    row = 1
    while current <= end:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
        c = ws.cell(row=row, column=1, value=f"Week of {current.isoformat()} (Sun–Sat)")
        c.font = _font(bold=True, color=C_DATE_FG)
        c.fill = _fill(C_DATE_BG)
        c.alignment = _align(h="left")
        c.border = BORDER
        row += 1

        headers = ["", "Slot"] + [
            f"{DAY_NAMES[i]} {(current + timedelta(days=i)).strftime('%m/%d')}" for i in range(7)
        ]
        _header_row(ws, row, headers)
        row += 1

        for section_key, time_label in sections:
            band_start = row
            for slot_idx, slot_label in enumerate(slot_rows):
                c = ws.cell(row=row, column=2, value=slot_label)
                c.border = BORDER
                c.alignment = _align(h="center")
                c.font = _font(bold=True)
                for i in range(7):
                    d = current + timedelta(days=i)
                    people = keys_to_people.get((d, section_key))
                    text, fill_hex = ("", C_EMT_BG)
                    if people is not None:
                        text, fill_hex = cell_for(people, slot_idx)
                    cell = ws.cell(row=row, column=3 + i, value=text)
                    cell.fill = _fill(fill_hex)
                    cell.border = BORDER
                    cell.alignment = _align(h="left", wrap=True)
                    cell.font = _font(size=9, bold=bool(text))
                row += 1

            ws.merge_cells(start_row=band_start, start_column=1, end_row=row - 1, end_column=1)
            c = ws.cell(row=band_start, column=1, value=f"{section_key}\n{time_label}")
            c.alignment = _align(h="center", wrap=True)
            c.font = _font(bold=True)
            for r in range(band_start, row):
                ws.cell(row=r, column=1).border = BORDER
                ws.cell(row=r, column=1).fill = _fill(C_ALT_ROW)
            row += 1  # spacer between bands

        row += 1  # spacer between weeks
        current += timedelta(days=7)


# ── Sheet 1: Ambulance schedule ──────────────────────────────────────────────

def _build_schedule_sheet(ws, assignments):
    ws.title = "Schedule"
    ws.freeze_panes = "C3"

    sections = [(s, f"{SHIFT_TIMES[s][0]}-{SHIFT_TIMES[s][1]}") for s in ("DAY", "AM", "PM", "NIGHT")]

    def cell_for(people, slot_idx):
        # Rows: EVDTs, Auths, then EMT-only in order.
        evdts = [v for v in people if v.is_evdt]
        auths = [v for v in people if v.certification == "Auth"]
        emts = [v for v in people if not v.is_driver]
        if slot_idx == 0:
            return ("\n".join(v.full_name for v in evdts), C_EVDT_BG if evdts else C_EMT_BG)
        if slot_idx == 1:
            return ("\n".join(v.full_name for v in auths), C_AUTH_BG if auths else C_EMT_BG)
        if slot_idx == 2:
            return (emts[0].full_name if emts else "", C_EMT_BG)
        # Last row stacks any remaining EMTs so no one falls off the grid.
        return ("\n".join(v.full_name for v in emts[1:]), C_EMT_BG)

    _weekly_grid(ws, assignments, sections, ["EVDT", "Auth", "EMT1", "EMT2"], cell_for)


# ── Sheet 2: Campus response ─────────────────────────────────────────────────

def _build_campus_sheet(ws, campus_assignments, responders_per_block: int):
    ws.title = "Campus Response"
    ws.freeze_panes = "C3"

    sections = [(b, CAMPUS_BLOCK_TIMES[b]) for b in ("A", "B", "C", "D")]
    slot_rows = [f"Responder{i + 1}" for i in range(responders_per_block)]

    def cell_for(people, slot_idx):
        if slot_idx < len(people):
            p = people[slot_idx]
            return (p.full_name, CERT_BG.get(p.certification, C_BERT_BG))
        return ("", C_EMT_BG)

    _weekly_grid(ws, campus_assignments, sections, slot_rows, cell_for)



def _fmt_keys(keys):
    return "; ".join(f"{d.month}/{d.day} {s}" for d, s in sorted(keys))


def _build_summary_sheet(ws, people, caps):
    ws.title = "Hour Summary"
    ws.freeze_panes = "A2"
    _header_row(ws, 1, ["Name", "Email", "Role", "Certification", "Ambulance Hours",
                       "Ambulance Shifts", "Campus Hours", "Campus Blocks",
                       "Ambulance Shortfall", "Campus Shortfall"],
                [28, 32, 10, 14, 18, 44, 14, 44, 20, 20])
    for row, p in enumerate(sorted(people, key=lambda p: p.full_name), 2):
        emt = isinstance(p, Volunteer)
        ambulance = p.assigned_hours if emt else 0
        values = [p.full_name, p.email, "EMT" if emt else "ERT/BERT", p.certification,
                  ambulance if emt else "—", _fmt_keys(p.assigned) if emt else "—",
                  p.campus_assigned_hours, _fmt_keys(p.campus_assigned),
                  caps.ambulance - ambulance if emt else "—",
                  caps.campus_for(p) - p.campus_assigned_hours]
        for col, value in enumerate(values, 1):
            cell = ws.cell(row, col, value)
            cell.font = _font()
            cell.fill = _fill(C_ALT_ROW if row % 2 == 0 else C_EMT_BG)
            cell.alignment = _align(wrap=col in (6, 8))
            cell.border = BORDER


def collect_warnings(schedule, providers, people, caps):
    """Report actual unmet needs, never a fictitious missing BLS truck driver."""
    issues = []
    for key, crew in schedule.ambulance.items():
        d, kind = key
        if not crew:
            issues.append(("NO AMBULANCE EMT", d, kind, "Supervisor is supplied separately"))
        if providers[key] == "ALS" and not any(p.is_evdt for p in crew):
            issues.append(("ALS — NO EVDT", d, kind,
                           "No assigned EVDT to drive while the ALS provider treats during transport"))
        if providers[key] == "BLS" and is_big_weekend(*key) and crew and not any(p.is_driver for p in crew):
            issues.append(("NO UTILITY DRIVER", d, kind,
                           "No assigned Utility-qualified volunteer for split crew; supervisor can drive ambulance"))
    for p in sorted(people, key=lambda p: p.full_name):
        if isinstance(p, Volunteer) and p.assigned_hours < caps.ambulance:
            issues.append(("AMBULANCE UNDER HOURS", None, "",
                           f"{p.full_name}: {p.assigned_hours}/{caps.ambulance}h"))
        if p.campus_assigned_hours < caps.campus_for(p):
            issues.append(("CAMPUS UNDER HOURS", None, "",
                           f"{p.full_name}: {p.campus_assigned_hours}/{caps.campus_for(p)}h"))
    for stage in schedule.stages:
        if stage.status != "OPTIMAL":
            issues.append(("SOLVER LIMIT", None, "", f"{stage.name}: {stage.status}; optimality not proven"))
    return issues


def _build_warnings_sheet(ws, issues):
    ws.title = "Warnings"
    _header_row(ws, 1, ["Type", "Date", "Day", "Shift", "Details"], [28, 13, 12, 10, 85])
    if not issues:
        ws.cell(2, 1, "No coverage, qualification or hour shortfalls reported.")
    for row, (kind, d, shift, detail) in enumerate(issues, 2):
        for col, value in enumerate([kind, d.isoformat() if d else "",
                                     d.strftime("%A") if d else "", shift, detail], 1):
            cell = ws.cell(row, col, value)
            cell.font = _font(bold=col == 1)
            cell.fill = _fill(C_WARN_BG)
            cell.alignment = _align(wrap=True)


def export_schedule_xlsx(schedule: Schedule, people, providers, caps: HourCaps,
                         output_path, campus_capacity=2):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    _build_schedule_sheet(wb.active, schedule.ambulance)
    _build_campus_sheet(wb.create_sheet(), schedule.campus, campus_capacity)
    _build_summary_sheet(wb.create_sheet(), people, caps)
    _build_warnings_sheet(wb.create_sheet(), collect_warnings(schedule, providers, people, caps))
    ws = wb.create_sheet("Solver")
    _header_row(ws, 1, ["Priority", "Objective", "Status", "Achieved", "Upper Bound", "Seconds"],
                [10, 52, 18, 14, 16, 12])
    for row, stage in enumerate(schedule.stages, 2):
        for col, value in enumerate([row - 1, stage.name, stage.status, stage.value,
                                     stage.bound, round(stage.seconds, 3)], 1):
            ws.cell(row, col, value)
    row = len(schedule.stages) + 3
    ws.cell(row, 1, "Each stage preserves earlier attained values. FEASIBLE is not proof of an optimum.")
    wb.save(output_path)
    return str(output_path)


def print_summary(schedule, people, providers, caps):
    print("\nSchedule summary")
    print(f"  Ambulance shifts with an EMT: {sum(bool(p) for p in schedule.ambulance.values())}/{len(schedule.ambulance)}")
    print(f"  Ambulance volunteer seats: {sum(map(len, schedule.ambulance.values()))}")
    print(f"  Campus blocks with a responder: {sum(bool(p) for p in schedule.campus.values())}/{len(schedule.campus)} (no required minimum)")
    for stage in schedule.stages:
        print(f"  {stage.name}: {stage.value}, {stage.status}, bound {stage.bound}, {stage.seconds:.2f}s")
    for kind, d, shift, details in collect_warnings(schedule, providers, people, caps):
        where = f"{d} {shift}: " if d else ""
        print(f"  {kind}: {where}{details}")


MASTER_SCHEDULE_HEADER = [
    "Block", "ShiftID", "Date", "Shift", "Vehicle", "Seat", "Requires", "Assigned/Name",
]


def _ambulance_seats(key, people, provider):
    """Vehicle-specific volunteer seats; supervisors are supplied separately.

    ALS reserves an R1/EVDT seat. On weekends a volunteer may drive U1; Auth
    and EVDT both qualify. BLS never fabricates an R1 driver shortage, and no
    crew-only EMT is labelled a driver to make the export fit.
    """
    remaining = sorted(people, key=lambda p: p.full_name)
    seats = []
    cap = crew_cap(*key)
    if provider == "ALS":
        driver = next((p for p in remaining if p.is_evdt), None)
        seats.append(("R1", "Driver", "EVDT", driver))
        if driver is not None:
            remaining.remove(driver)
    if is_big_weekend(*key):
        driver = next((p for p in remaining if p.is_driver), None)
        if driver is not None or len(seats) + len(remaining) < cap:
            seats.append(("U1", "Driver", "AUTH", driver))
            if driver is not None:
                remaining.remove(driver)
    for p in remaining:
        seats.append(("R1", f"C{len(seats) + 1}", "CREW", p))
    while len(seats) < cap:
        seats.append(("R1", f"C{len(seats) + 1}", "CREW", None))
    if len(seats) > cap:
        raise ValueError(f"Export would exceed capacity for {key}")
    return seats


def export_master_schedule_csv(schedule, providers, block_start, output_path,
                               block, daynum_start, campus_capacity=2):
    rows = []

    def row(key, vehicle, seat, requires, person):
        d, kind = key
        daynum = daynum_start + (d - block_start).days
        suffix = requires if seat == "Driver" else seat
        shift_id = f"{block}-{daynum:04d}-{kind}-{vehicle}-{suffix}"
        return [block, shift_id, f"{d.month}/{d.day}/{d:%y}", kind, vehicle,
                seat, requires, person.full_name if person else ""]

    for key, people in sorted(schedule.ambulance.items()):
        for vehicle, seat, requires, person in _ambulance_seats(key, people, providers[key]):
            rows.append(row(key, vehicle, seat, requires, person))
    for key, people in sorted(schedule.campus.items()):
        ordered = sorted(people, key=lambda p: p.full_name)
        if len(ordered) > campus_capacity:
            raise ValueError(f"Export would drop campus responders for {key}")
        for i in range(campus_capacity):
            rows.append(row(key, "CR", f"S{i + 1}", "CREW", ordered[i] if i < len(ordered) else None))
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(MASTER_SCHEDULE_HEADER)
        writer.writerows(rows)
    return str(output_path)
