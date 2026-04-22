"""Creates the IO_List_Input_Template.xlsx workbook with all sheets pre-populated."""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

NAVY   = "FFC0C0C0"
ORANGE = "FFFFC000"
LGRAY  = "FFD9D9D9"
DGRAY  = "FF595959"
WHITE  = "FFFFFFFF"
LBLUE  = "FFD6E4F0"
LGREEN = "FFE2EFDA"
LYELLOW= "FFFFF2CC"
LPINK  = "FFFCE4D6"

def hdr_font(color="FF000000", sz=10, bold=True):
    return Font(name="Arial", size=sz, bold=bold, color=color)

def body_font(sz=10, bold=False, color="FF000000"):
    return Font(name="Arial", size=sz, bold=bold, color=color)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def thin_border():
    s = Side(border_style="thin", color="FF000000")
    return Border(left=s, right=s, top=s, bottom=s)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left_middle():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def apply_header(cell, value, bg=NAVY, fg="FF000000", bold=True, sz=10):
    cell.value = value
    cell.font = hdr_font(fg, sz, bold)
    cell.fill = fill(bg)
    cell.border = thin_border()
    cell.alignment = center()

def apply_body(cell, value, bg=None, bold=False, align="left"):
    cell.value = value
    cell.font = body_font(bold=bold)
    if bg:
        cell.fill = fill(bg)
    cell.border = thin_border()
    cell.alignment = left_middle() if align == "left" else center()


# ─────────────────────────────────────────────
# Sheet 1: PROJECT_INFO
# ─────────────────────────────────────────────
def build_project_info(wb):
    ws = wb.create_sheet("PROJECT_INFO")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 40

    apply_header(ws["A1"], "PROJECT INFORMATION", NAVY, "FF000000", sz=12)
    ws.merge_cells("A1:B1")
    ws.row_dimensions[1].height = 30

    fields = [
        ("Project Name", "My Project"),
        ("Project Number", "PRJ-001"),
        ("Controller Tag Prefix", "PLC-1"),
        ("Controller Model", "5069-L310ER"),
        ("IO Address Prefix", "DI_"),
        ("Engineer Name", "Eng. Anas"),
        ("Client Name", ""),
        ("Site Name", ""),
        ("Revision", "A"),
        ("Notes", ""),
    ]
    for i, (label, val) in enumerate(fields, start=2):
        apply_header(ws.cell(i, 1), label, LGRAY, "FF000000")
        apply_body(ws.cell(i, 2), val)
        ws.row_dimensions[i].height = 20

    note_row = len(fields) + 3
    ws.cell(note_row, 1).value = "⚠ Fill in all yellow cells before running the generator."
    ws.cell(note_row, 1).font = Font(name="Arial", size=9, italic=True, color="FF7F0000")
    ws.merge_cells(f"A{note_row}:B{note_row}")


# ─────────────────────────────────────────────
# Sheet 2: MODULE_DB
# ─────────────────────────────────────────────
def build_module_db(wb):
    ws = wb.create_sheet("MODULE_DB")
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 12

    headers = ["Model No.", "Description", "Type (DI/DO/AI/AO)", "Channel Qty"]
    for col, h in enumerate(headers, 1):
        apply_header(ws.cell(1, col), h)
    ws.row_dimensions[1].height = 25

    modules = [
        ("1734-IB8",  "POINT I/O 8-Ch Digital Input 24VDC",          "DI", 8),
        ("1734-IB16", "POINT I/O 16-Ch Digital Input 24VDC",          "DI", 16),
        ("1734-OB8",  "POINT I/O 8-Ch Digital Output 24VDC",          "DO", 8),
        ("1734-OB16", "POINT I/O 16-Ch Digital Output 24VDC",         "DO", 16),
        ("1734-IE8C", "POINT I/O 8-Ch Analog Input 4-20mA",           "AI", 8),
        ("1734-IE4C", "POINT I/O 4-Ch Analog Input 4-20mA",           "AI", 4),
        ("1734-OE4C", "POINT I/O 4-Ch Analog Output 4-20mA",          "AO", 4),
        ("1734-OE2C", "POINT I/O 2-Ch Analog Output 4-20mA",          "AO", 2),
        ("1756-IB32", "ControlLogix 32-Ch Digital Input 24VDC",        "DI", 32),
        ("1756-IB16", "ControlLogix 16-Ch Digital Input 24VDC",        "DI", 16),
        ("1756-OB32", "ControlLogix 32-Ch Digital Output 24VDC",       "DO", 32),
        ("1756-OB16", "ControlLogix 16-Ch Digital Output 24VDC",       "DO", 16),
        ("1756-IF16", "ControlLogix 16-Ch Analog Input",               "AI", 16),
        ("1756-IF8",  "ControlLogix 8-Ch Analog Input",                "AI", 8),
        ("1756-OF8",  "ControlLogix 8-Ch Analog Output",               "AO", 8),
        ("1756-OF4",  "ControlLogix 4-Ch Analog Output",               "AO", 4),
        ("5069-IB16", "Compact 5000 16-Ch Digital Input 24VDC",        "DI", 16),
        ("5069-IB8",  "Compact 5000 8-Ch Digital Input 24VDC",         "DI", 8),
        ("5069-OB16", "Compact 5000 16-Ch Digital Output 24VDC",       "DO", 16),
        ("5069-OB8",  "Compact 5000 8-Ch Digital Output 24VDC",        "DO", 8),
        ("5069-IF8",  "Compact 5000 8-Ch Analog Input",                "AI", 8),
        ("5069-IF4",  "Compact 5000 4-Ch Analog Input",                "AI", 4),
        ("5069-OF4",  "Compact 5000 4-Ch Analog Output",               "AO", 4),
        ("5069-OF2",  "Compact 5000 2-Ch Analog Output",               "AO", 2),
    ]
    type_fill = {"DI": LBLUE, "DO": LGREEN, "AI": LYELLOW, "AO": LPINK}
    for row, (model, desc, typ, qty) in enumerate(modules, 2):
        bg = type_fill.get(typ, WHITE)
        apply_body(ws.cell(row, 1), model)
        apply_body(ws.cell(row, 2), desc)
        apply_body(ws.cell(row, 3), typ, bg=bg, align="center")
        apply_body(ws.cell(row, 4), qty, align="center")
        ws.row_dimensions[row].height = 16

    legend_row = len(modules) + 3
    ws.cell(legend_row, 1).value = "Color legend:"
    ws.cell(legend_row, 1).font = Font(name="Arial", size=9, bold=True)
    for i, (t, bg) in enumerate(type_fill.items(), 1):
        c = ws.cell(legend_row, i + 1)
        c.value = t
        c.fill = fill(bg)
        c.font = Font(name="Arial", size=9, bold=True)
        c.border = thin_border()
        c.alignment = center()


# ─────────────────────────────────────────────
# Sheet 3: HARDWARE_CONFIG
# ─────────────────────────────────────────────
def build_hardware_config(wb):
    ws = wb.create_sheet("HARDWARE_CONFIG")
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 12

    headers = ["Rack Name", "Module Model No.", "Modules Qty"]
    for col, h in enumerate(headers, 1):
        apply_header(ws.cell(1, col), h)
    ws.row_dimensions[1].height = 25

    samples = [
        ("RACK 01", "1734-IB8",  2),
        ("RACK 01", "1734-OB8",  1),
        ("RACK 01", "1734-IE8C", 1),
        ("RACK 01", "1734-OE4C", 1),
        ("RACK 02", "1734-IB8",  2),
        ("RACK 02", "1734-OB8",  1),
    ]
    for row, (rack, model, qty) in enumerate(samples, 2):
        apply_body(ws.cell(row, 1), rack)
        apply_body(ws.cell(row, 2), model)
        apply_body(ws.cell(row, 3), qty, align="center")
        ws.row_dimensions[row].height = 16

    note_row = len(samples) + 3
    ws.cell(note_row, 1).value = (
        "⚠ Rack Name must be consistent (e.g. 'RACK 01'). "
        "Model No. must match MODULE_DB exactly. "
        "Signals are arranged DI → DO → AI → AO per rack."
    )
    ws.cell(note_row, 1).font = Font(name="Arial", size=9, italic=True, color="FF7F0000")
    ws.merge_cells(f"A{note_row}:C{note_row}")


# ─────────────────────────────────────────────
# Sheet 4: SIGNAL_MATRIX
# ─────────────────────────────────────────────
def build_signal_matrix(wb):
    ws = wb.create_sheet("SIGNAL_MATRIX")

    eq_types   = ["VSD", "FSD", "HVALVE", "MVALVE", "AI", "DI"]
    sig_rows   = [
        ("HW", "DI", "DI Signal1"),
        ("HW", "DI", "DI Signal2"),
        ("HW", "DI", "DI Signal3"),
        ("HW", "DI", "DI Signal4"),
        ("HW", "DO", "DO Signal1"),
        ("HW", "DO", "DO Signal2"),
        ("HW", "AI", "AI Signal1"),
        ("HW", "AI", "AI Signal2"),
        ("HW", "AO", "AO Signal1"),
        ("HW", "AO", "AO Signal2"),
    ]
    matrix = {
        "DI Signal1": {"VSD": "",            "FSD": "Running",       "HVALVE": "Opened",        "MVALVE": "Opened",       "AI": "",    "DI": "Input"},
        "DI Signal2": {"VSD": "",            "FSD": "Fault",         "HVALVE": "Closed",        "MVALVE": "Closed",       "AI": "",    "DI": ""},
        "DI Signal3": {"VSD": "",            "FSD": "Remote",        "HVALVE": "",              "MVALVE": "Remote",       "AI": "",    "DI": ""},
        "DI Signal4": {"VSD": "",            "FSD": "",              "HVALVE": "",              "MVALVE": "",             "AI": "",    "DI": ""},
        "DO Signal1": {"VSD": "",            "FSD": "Start Command", "HVALVE": "",              "MVALVE": "Open Command", "AI": "",    "DI": ""},
        "DO Signal2": {"VSD": "",            "FSD": "",              "HVALVE": "",              "MVALVE": "Close Command","AI": "",    "DI": ""},
        "AI Signal1": {"VSD": "",            "FSD": "",              "HVALVE": "",              "MVALVE": "",             "AI": "Raw", "DI": ""},
        "AI Signal2": {"VSD": "",            "FSD": "",              "HVALVE": "",              "MVALVE": "",             "AI": "",    "DI": ""},
        "AO Signal1": {"VSD": "",            "FSD": "",              "HVALVE": "",              "MVALVE": "",             "AI": "",    "DI": ""},
        "AO Signal2": {"VSD": "",            "FSD": "",              "HVALVE": "",              "MVALVE": "",             "AI": "",    "DI": ""},
    }

    type_fill = {"DI": LBLUE, "DO": LGREEN, "AI": LYELLOW, "AO": LPINK}

    apply_header(ws.cell(1, 1), "Category")
    apply_header(ws.cell(1, 2), "Signal Type")
    apply_header(ws.cell(1, 3), "Signal Row Name")
    for col, et in enumerate(eq_types, 4):
        apply_header(ws.cell(1, col), et)
    ws.row_dimensions[1].height = 25

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 18
    for col in range(4, 4 + len(eq_types)):
        ws.column_dimensions[get_column_letter(col)].width = 18

    for row, (cat, sig_type, sig_name) in enumerate(sig_rows, 2):
        bg = type_fill.get(sig_type, WHITE)
        apply_body(ws.cell(row, 1), cat, align="center")
        apply_body(ws.cell(row, 2), sig_type, bg=bg, align="center")
        apply_body(ws.cell(row, 3), sig_name)
        for col, et in enumerate(eq_types, 4):
            val = matrix.get(sig_name, {}).get(et, "")
            apply_body(ws.cell(row, col), val, align="center")
        ws.row_dimensions[row].height = 16

    note_row = len(sig_rows) + 3
    ws.cell(note_row, 1).value = (
        "⚠ Signal Row Name must start with DI/DO/AI/AO. "
        "Cell value = signal description for that equipment type. Leave blank if not applicable. "
        "Add columns for new equipment types. Add rows for new signal rows."
    )
    ws.cell(note_row, 1).font = Font(name="Arial", size=9, italic=True, color="FF7F0000")
    ws.merge_cells(f"A{note_row}:{get_column_letter(3+len(eq_types))}{note_row}")


# ─────────────────────────────────────────────
# Sheet 5: EQUIPMENT_LIST
# ─────────────────────────────────────────────
def build_equipment_list(wb):
    ws = wb.create_sheet("EQUIPMENT_LIST")
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 20

    headers = ["Tag", "Description", "Type", "System (optional)"]
    for col, h in enumerate(headers, 1):
        apply_header(ws.cell(1, col), h)
    ws.row_dimensions[1].height = 25

    samples = [
        ("VFD-001", "Scum Sludge Pump 1",              "VSD",   "MCC-01"),
        ("VFD-002", "Scum Sludge Pump 2",              "VSD",   "MCC-01"),
        ("MOT-001", "Grit Blower 1",                   "FSD",   "MCC-01"),
        ("MOT-002", "Grit Blower 2",                   "FSD",   "MCC-01"),
        ("XV-001",  "Inlet Control Valve",             "HVALVE","PCV-01"),
        ("MV-001",  "Sludge Discharge Valve",          "MVALVE","PCV-01"),
        ("LIT-001", "Headworks Overflow Level",        "AI",    "INS-01"),
        ("DI-001",  "High Level Switch Headworks",     "DI",    "INS-01"),
    ]
    for row, (tag, desc, typ, sys) in enumerate(samples, 2):
        apply_body(ws.cell(row, 1), tag)
        apply_body(ws.cell(row, 2), desc)
        apply_body(ws.cell(row, 3), typ, align="center")
        apply_body(ws.cell(row, 4), sys)
        ws.row_dimensions[row].height = 16

    note_row = len(samples) + 3
    ws.cell(note_row, 1).value = (
        "⚠ Type must match exactly one of the equipment type columns in SIGNAL_MATRIX sheet. "
        "Tag and Description are required."
    )
    ws.cell(note_row, 1).font = Font(name="Arial", size=9, italic=True, color="FF7F0000")
    ws.merge_cells(f"A{note_row}:D{note_row}")


# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────
def main(out_path=None):
    if out_path is None:
        out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "IO_List_Input_Template.xlsx")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    build_project_info(wb)
    build_module_db(wb)
    build_hardware_config(wb)
    build_signal_matrix(wb)
    build_equipment_list(wb)

    wb.save(out_path)
    print(f"Template saved: {out_path}")

if __name__ == "__main__":
    main()
