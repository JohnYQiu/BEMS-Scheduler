"""
output.py
=========
Handles all output for the Brown EMS scheduler:
  - Formatted .xlsx file with three sheets:
      1. Schedule   — day-by-day, colour-coded by certification & ALS status
      2. Summary    — volunteer hour totals with under-18h flags
      3. Warnings   — ALS shifts missing EVDT, unfilled slots, no auth driver
  - Brief terminal summary
"""

from datetime import date
from scheduler import Shift
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SHIFT_TIMES = {
    "AM":    ("0700", "1300"),
    "PM":    ("1300", "1900"),
    "NIGHT": ("1900", "0700+1"),
    "DAY":   ("0700", "1900"),
}

C_HEADER_BG = "1F3864"
C_HEADER_FG = "FFFFFF"
C_DATE_BG   = "2E5FA3"
C_DATE_FG   = "FFFFFF"
C_ALS_BG    = "FCE4D6"
C_EVDT_BG   = "E2EFDA"
C_AUTH_BG   = "FFF2CC"
C_EMT_BG    = "FFFFFF"
C_UNFILLED  = "F4CCCC"
C_WARN_BG   = "FFE0E0"
C_ALT_ROW   = "F8F8F8"
C_UNDER18   = "CC0000"

CERT_BG = {"EVDT": C_EVDT_BG, "Auth": C_AUTH_BG, "EMT": C_EMT_BG}

_thin   = Side(style="thin", color="CCCCCC")
BORDER  = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def _fill(hex_color):
    return PatternFill("solid", start_color=hex_color, fgColor=hex_color)

def _font(bold=False, color="000000", size=10, italic=False):
    return Font(name="Arial", bold=bold, color=color, size=size, italic=italic)

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _header_row(ws, row, values, widths=None):
    for col, val in enumerate(values, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = _font(bold=True, color=C_HEADER_FG)
        c.fill      = _fill(C_HEADER_BG)
        c.alignment = _align(h="center")
        c.border    = BORDER
    if widths:
        for col, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w


def _build_schedule_sheet(ws, all_shifts):
    ws.title = "Schedule"
    ws.freeze_panes = "A2"
    headers = ["Date", "Day", "Shift", "Time", "ALS", "Volunteer", "Certification", "Total Hours"]
    widths  = [13, 12, 8, 14, 6, 28, 14, 12]
    _header_row(ws, 1, headers, widths)

    row = 2
    current_date = None

    for key in sorted(all_shifts):
        shift = all_shifts[key]
        t_start, t_end = SHIFT_TIMES[shift.shift_type]
        time_str = f"{t_start}-{t_end}"
        als_str  = "YES" if shift.has_als else ""

        if shift.date != current_date:
            current_date = shift.date
            date_label = shift.date.strftime("%A, %B %d %Y")
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            c = ws.cell(row=row, column=1, value=date_label)
            c.font      = _font(bold=True, color=C_DATE_FG)
            c.fill      = _fill(C_DATE_BG)
            c.alignment = _align(h="left")
            c.border    = BORDER
            row += 1

        if not shift.volunteers:
            vals = ["", "", shift.shift_type, time_str, als_str, "— UNFILLED —", "", ""]
            for col, val in enumerate(vals, 1):
                c = ws.cell(row=row, column=col, value=val)
                c.fill      = _fill(C_UNFILLED)
                c.font      = _font(bold=True, color="CC0000", italic=True)
                c.alignment = _align(h="center" if col in (3,4,5) else "left")
                c.border    = BORDER
            row += 1
            continue

        for i, v in enumerate(shift.volunteers):
            cert_bg = CERT_BG.get(v.certification, C_EMT_BG)
            vals = [
                shift.date.isoformat() if i == 0 else "",
                shift.date.strftime("%A") if i == 0 else "",
                shift.shift_type if i == 0 else "",
                time_str if i == 0 else "",
                als_str if i == 0 else "",
                v.full_name,
                v.certification,
                v.scheduled_hours,
            ]
            for col, val in enumerate(vals, 1):
                c = ws.cell(row=row, column=col, value=val)
                c.fill      = _fill(cert_bg)
                c.font      = _font(bold=(col == 6))
                c.alignment = _align(h="center" if col in (3,4,5,7,8) else "left")
                c.border    = BORDER
            row += 1

    row += 1
    legend = [("EVDT", C_EVDT_BG), ("Auth Driver", C_AUTH_BG),
              ("EMT", C_EMT_BG), ("Unfilled", C_UNFILLED), ("ALS Shift", C_ALS_BG)]
    ws.cell(row=row, column=1, value="Legend:").font = _font(bold=True)
    for i, (label, color) in enumerate(legend, 2):
        c = ws.cell(row=row, column=i, value=label)
        c.fill      = _fill(color)
        c.font      = _font(size=9)
        c.alignment = _align(h="center")
        c.border    = BORDER


def _build_summary_sheet(ws, volunteers):
    ws.title = "Hour Summary"
    ws.freeze_panes = "A2"
    headers = ["Name", "Email", "Certification", "Scheduled Hours", "Status"]
    widths  = [28, 32, 14, 16, 16]
    _header_row(ws, 1, headers, widths)

    sorted_vols = sorted(volunteers, key=lambda v: -v.scheduled_hours)
    for i, v in enumerate(sorted_vols, 2):
        under = v.scheduled_hours < 18
        bg    = C_ALT_ROW if i % 2 == 0 else C_EMT_BG
        vals  = [v.full_name, v.email, v.certification,
                 v.scheduled_hours, "⚠ Under 18h" if under else "OK"]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=i, column=col, value=val)
            c.fill      = _fill(bg)
            c.font      = _font(
                bold=(col == 5 and under),
                color=(C_UNDER18 if (col == 5 and under) else "000000")
            )
            c.alignment = _align(h="center" if col in (3,4,5) else "left")
            c.border    = BORDER


def _build_warnings_sheet(ws, all_shifts):
    ws.title = "Warnings"
    headers = ["Type", "Date", "Day", "Shift", "Details"]
    widths  = [25, 13, 12, 8, 45]
    _header_row(ws, 1, headers, widths)

    issues = []
    for key in sorted(all_shifts):
        shift = all_shifts[key]
        if not shift.volunteers:
            issues.append(("UNFILLED SHIFT", shift.date, shift.shift_type, "No volunteers assigned"))
        if shift.has_als and not shift.has_evdt:
            names = ", ".join(v.full_name for v in shift.volunteers) or "—"
            issues.append(("ALS - NO EVDT", shift.date, shift.shift_type, f"Assigned: {names}"))
        if not shift.has_auth:
            issues.append(("NO AUTH DRIVER", shift.date, shift.shift_type, "No EVDT or Auth on shift"))

    if not issues:
        c = ws.cell(row=2, column=1, value="No warnings — all shifts adequately staffed.")
        c.font = _font(bold=True, color="1E7E34")
        return

    for i, (type_, d, shift_type, detail) in enumerate(issues, 2):
        bg   = "FFF0F0" if i % 2 == 0 else C_WARN_BG
        vals = [type_, d.isoformat(), d.strftime("%A"), shift_type, detail]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=i, column=col, value=val)
            c.fill      = _fill(bg)
            c.font      = _font(bold=(col == 1), color=("CC0000" if col == 1 else "000000"))
            c.alignment = _align(h="center" if col in (2,3,4) else "left")
            c.border    = BORDER


def _build_strike_list_sheet(ws, violations):
    ws.title = "Strike List"
    ws.freeze_panes = "A2"
    headers = ["Name", "Email", "Certification", "Missing Requirements"]
    widths  = [28, 32, 14, 40]
    _header_row(ws, 1, headers, widths)

    if not violations:
        c = ws.cell(row=2, column=1, value="✓ All volunteers met the minimum availability requirements.")
        c.font = _font(bold=True, color="1E7E34")
        return

    for i, item in enumerate(violations, 2):
        v = item["volunteer"]
        missing = ", ".join(item.get("missing", []))
        bg = C_ALT_ROW if i % 2 == 0 else C_EMT_BG
        vals = [v.full_name, v.email, v.certification, missing]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=i, column=col, value=val)
            c.fill      = _fill(bg)
            c.font      = _font(bold=(col == 1))
            c.alignment = _align(h="center" if col in (3,) else "left", wrap=(col == 4))
            c.border    = BORDER


def export_schedule_xlsx(all_shifts, volunteers, output_path, violations=None):
    if not output_path.endswith(".xlsx"):
        output_path = output_path.rsplit(".", 1)[0] + ".xlsx"

    wb = Workbook()
    _build_schedule_sheet(wb.active, all_shifts)
    _build_summary_sheet(wb.create_sheet(), volunteers)
    _build_warnings_sheet(wb.create_sheet(), all_shifts)
    _build_strike_list_sheet(wb.create_sheet(), violations or [])
    wb.save(output_path)
    print(f"  Schedule exported -> {output_path}")
    return output_path


def print_summary(all_shifts, volunteers):
    total       = len(all_shifts)
    unfilled    = sum(1 for s in all_shifts.values() if not s.volunteers)
    als_no_evdt = sum(1 for s in all_shifts.values() if s.has_als and not s.has_evdt)
    under18     = sum(1 for v in volunteers if v.scheduled_hours < 18)

    print("\n" + "=" * 55)
    print("SCHEDULE SUMMARY")
    print("=" * 55)
    print(f"  Total shifts:          {total}")
    print(f"  Unfilled shifts:       {unfilled}" + (" ⚠" if unfilled else " ✓"))
    print(f"  ALS shifts w/o EVDT:   {als_no_evdt}" + (" ⚠" if als_no_evdt else " ✓"))
    print(f"  Volunteers under 18h:  {under18}" + (" ⚠" if under18 else " ✓"))
    print()


def print_warnings(all_shifts):
    issues = []
    for key in sorted(all_shifts):
        shift = all_shifts[key]
        if not shift.volunteers:
            issues.append(f"UNFILLED:       {shift.label}")
        if shift.has_als and not shift.has_evdt:
            issues.append(f"ALS / NO EVDT:  {shift.label}")
        if not shift.has_auth:
            issues.append(f"NO AUTH DRIVER: {shift.label}")

    print("=" * 55)
    print("WARNINGS")
    print("=" * 55)
    if not issues:
        print("  No warnings.")
    else:
        for issue in issues:
            print(f"  ⚠  {issue}")
    print()