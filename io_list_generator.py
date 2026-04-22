"""
IO List Generator — Allen-Bradley / Rockwell Automation
Reads a structured Excel workbook and generates a fully formatted IO List.
Supports normal generation and manufacturing revision mode.
"""
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────────────────────
# STYLING CONSTANTS
# ──────────────────────────────────────────────────────────────
C_NAVY    = "FFC0C0C0"   # light gray (used everywhere navy was)
C_ORANGE  = "FFFFC000"
C_LGRAY   = "FFD9D9D9"
C_DGRAY   = "FF595959"
C_WHITE   = "FFFFFFFF"
C_BLACK   = "FF000000"
C_LBLUE   = "FFD6E4F0"
C_LGREEN  = "FFE2EFDA"
C_LYELLOW = "FFFFF2CC"
C_LPINK   = "FFFCE4D6"
C_RED_HDR = "FFC00000"
C_UNASSIGNED = "FFFFD7D7"

WIRE_COLORS = ["Black","Red","White","Green","Brown","Blue","Orange","Yellow","Gray","Purple",
               "Pink","Violet","Tan","Cyan","Maroon","Olive"]

IO_COLUMNS = [
    "Controller","Tag No.","PlantPax","Equipment Description","Signal Description",
    "Rack","Slot","CH","Type","Wire Color","Marshalling\nTerminal Block",
    "Term1","Term2","Term3","Power","Feed\nF: Field\nS: Self","Intrinsically Safe",
    "System","Sub System","Drawing/Panel","ON State\n(1)","OFF State\n(0)","Alarm",
    "Scale Min","Scale Max","Eng. Unit","LL","L","H","HH","Field Cable\nID","Notes"
]

COL_WIDTHS = {
    1:12.14, 2:20.29, 3:9.29,  4:59.43, 5:23.57, 6:3.43,  7:13.0,  8:13.0,
    9:11.14, 10:14.0, 11:11.86,12:6.71, 13:6.0,  14:6.0,  15:7.14, 16:11.71,
    17:9.0,  18:8.14, 19:7.57, 20:15.14,21:9.0,  22:5.86, 23:6.43, 24:4.57,
    25:4.86, 26:8.0,  27:3.57, 28:2.14, 29:2.29, 30:3.57, 31:6.43, 32:6.0
}

# ──────────────────────────────────────────────────────────────
# EXCEL HELPERS
# ──────────────────────────────────────────────────────────────
def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(color=C_BLACK, sz=10, bold=False, name="Arial"):
    return Font(name=name, size=sz, bold=bold, color=color)

def _border(style="thin", color=C_BLACK):
    s = Side(border_style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _align(h="center", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _style(cell, value=None, bg=None, fg=C_BLACK, bold=False, sz=10,
           h="center", v="center", wrap=True, border=True):
    if value is not None:
        cell.value = value
    if bg:
        cell.fill = _fill(bg)
    cell.font = _font(fg, sz, bold)
    if border:
        cell.border = _border()
    cell.alignment = _align(h, v, wrap)


# ──────────────────────────────────────────────────────────────
# EXCEL READER
# ──────────────────────────────────────────────────────────────
def read_workbook(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    def sheet_rows(name, skip=1):
        if name not in wb.sheetnames:
            raise ValueError(f"Sheet '{name}' not found in workbook.")
        ws = wb[name]
        rows = []
        for row in ws.iter_rows(min_row=skip + 1, values_only=True):
            if any(v is not None and str(v).strip() for v in row):
                rows.append([str(v).strip() if v is not None else "" for v in row])
        return rows

    # PROJECT_INFO
    proj = {}
    ws_proj = wb["PROJECT_INFO"]
    for row in ws_proj.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1] is not None:
            proj[str(row[0]).strip()] = str(row[1]).strip() if row[1] else ""

    # MODULE_DB
    module_db = {}
    for row in sheet_rows("MODULE_DB"):
        if len(row) >= 4 and row[0]:
            module_db[row[0]] = {
                "desc": row[1],
                "type": row[2].upper(),
                "qty":  int(row[3]) if row[3].isdigit() else 0
            }

    # HARDWARE_CONFIG
    hw_config = []
    for row in sheet_rows("HARDWARE_CONFIG"):
        if len(row) >= 3 and row[0] and row[1]:
            try:
                qty = int(row[2])
            except:
                qty = 1
            hw_config.append({"rack": row[0], "model": row[1], "qty": qty})

    # SIGNAL_MATRIX — row format: Category | SigType | SigRowName | EqType1 | EqType2 ...
    ws_mat = wb["SIGNAL_MATRIX"]
    mat_header = [str(c).strip() if c else "" for c in next(ws_mat.iter_rows(min_row=1, max_row=1, values_only=True))]
    eq_types = mat_header[3:]
    matrix = {}
    sig_order = []
    for row in ws_mat.iter_rows(min_row=2, values_only=True):
        if not row[2]:
            continue
        sig_name = str(row[2]).strip()
        if sig_name not in matrix:
            matrix[sig_name] = {}
            sig_order.append(sig_name)
        for i, et in enumerate(eq_types):
            val = str(row[3 + i]).strip() if len(row) > 3 + i and row[3 + i] else ""
            matrix[sig_name][et] = val

    # EQUIPMENT_LIST
    equipment = []
    for row in sheet_rows("EQUIPMENT_LIST"):
        if len(row) >= 3 and row[0] and row[2]:
            equipment.append({
                "tag":    row[0],
                "desc":   row[1],
                "type":   row[2].upper(),
                "system": row[3] if len(row) > 3 else ""
            })

    return proj, module_db, hw_config, eq_types, matrix, sig_order, equipment


# ──────────────────────────────────────────────────────────────
# MANUFACTURING IO LIST PARSER
# ──────────────────────────────────────────────────────────────
def parse_manufacturing_io_list(path):
    """Parse a previously generated IO List Excel file into a flat channel list."""
    wb = openpyxl.load_workbook(path, data_only=True)
    if "IO List" not in wb.sheetnames:
        raise ValueError("Attached file does not contain an 'IO List' sheet.")
    ws = wb["IO List"]

    TYPE_SUFFIX_REV = {"DIM": "DI", "DOM": "DO", "AIM": "AI", "AOM": "AO"}

    channels = []
    current_rack_name = ""
    current_model = ""
    current_mod_desc = ""
    current_rack_num = 0
    current_slot_num = 0

    for row in ws.iter_rows(min_row=3, values_only=True):
        if not any(v for v in row):
            continue

        def cell(i, row=row):
            v = row[i] if i < len(row) else None
            return str(v).strip() if v is not None else ""

        type_val = cell(8)  # column 9

        if type_val == "Processor":
            continue

        if type_val in TYPE_SUFFIX_REV:
            current_rack_name = cell(0)
            col4_val = cell(3)
            if " — " in col4_val:
                parts = col4_val.split(" — ", 1)
                current_model = parts[0].strip()
                current_mod_desc = parts[1].strip()
            else:
                current_model = col4_val
                current_mod_desc = ""
            try:
                current_rack_num = int(cell(5))
            except:
                current_rack_num = 0
            try:
                current_slot_num = int(cell(6))
            except:
                current_slot_num = 0
            continue

        if type_val in ("DI", "DO", "AI", "AO"):
            tag = cell(1)
            sig_desc = cell(4)
            is_spare = (not tag or sig_desc == "SPARE")
            try:
                ch = int(cell(7))
            except:
                ch = 0

            channels.append({
                "rack_name":  current_rack_name,
                "rack_num":   current_rack_num,
                "slot_num":   current_slot_num,
                "ch":         ch,
                "sig_type":   type_val,
                "model":      current_model,
                "mod_desc":   current_mod_desc,
                "tag":        tag,
                "sig_desc":   sig_desc,
                "eq_desc":    cell(3),
                "system":     cell(17),
                "wire_color": cell(9),
                "term_block": cell(10),
                "t1":         cell(11),
                "t2":         cell(12),
                "power":      cell(14),
                "io_addr":    cell(0),
                "is_spare":   is_spare,
            })

    if not channels:
        raise ValueError("No IO channels found in the attached manufacturing IO list.")
    return channels


# ──────────────────────────────────────────────────────────────
# CORE ENGINE
# ──────────────────────────────────────────────────────────────
def get_sig_type(sig_row_name):
    n = sig_row_name.strip().upper()
    for t in ("DI", "DO", "AI", "AO"):
        if n.startswith(t):
            return t
    return None


def build_io_list(proj, module_db, hw_config, eq_types, matrix, sig_order, equipment, log_fn=None):
    def log(msg):
        if log_fn:
            log_fn(msg)

    prefix   = proj.get("Controller Tag Prefix", "PLC-1")

    # Build signal demand per type
    demand = {"DI": [], "DO": [], "AI": [], "AO": []}
    unknown_types = set()
    for eq in equipment:
        eq_type = eq["type"]
        if eq_type not in eq_types:
            unknown_types.add(eq_type)
            continue
        for sig_name in sig_order:
            sig_desc = matrix.get(sig_name, {}).get(eq_type, "")
            if not sig_desc.strip():
                continue
            st = get_sig_type(sig_name)
            if not st:
                continue
            demand[st].append({
                "tag":      eq["tag"],
                "eq_desc":  eq["desc"],
                "sig_desc": sig_desc,
                "system":   eq["system"],
                "eq_type":  eq_type,
                "sig_row":  sig_name,
                "_assigned": False,
            })

    if unknown_types:
        log(f"⚠  Unknown equipment types (not in signal matrix): {', '.join(unknown_types)}")

    # Build rack structure
    racks = {}
    rack_order = []
    for row in hw_config:
        rack_name = row["rack"]
        if rack_name not in racks:
            racks[rack_name] = []
            rack_order.append(rack_name)
        info = module_db.get(row["model"])
        if not info:
            log(f"⚠  Module '{row['model']}' not found in MODULE_DB — skipped.")
            continue
        for _ in range(row["qty"]):
            racks[rack_name].append({
                "model": row["model"],
                "desc":  info["desc"],
                "type":  info["type"],
                "qty":   info["qty"],
            })

    io_rows = []
    global_slot = 1
    TYPE_ORDER  = ["DI", "DO", "AI", "AO"]
    POWER_MAP   = {"DI": "24VDC", "DO": "NO", "AI": "", "AO": ""}

    for rack_name in rack_order:
        rack_num  = rack_order.index(rack_name) + 1
        modules   = racks[rack_name]
        rack_slot_start = global_slot
        rack_modules_by_type = {t: [m for m in modules if m["type"] == t] for t in TYPE_ORDER}

        for sig_type in TYPE_ORDER:
            type_modules = rack_modules_by_type[sig_type]
            if not type_modules:
                continue

            pending = [d for d in demand[sig_type] if not d["_assigned"]]
            pending_idx = 0

            for mod_idx, mod in enumerate(type_modules):
                slot_num = rack_slot_start + sum(
                    len(rack_modules_by_type[t])
                    for t in TYPE_ORDER[:TYPE_ORDER.index(sig_type)]
                ) + mod_idx

                io_rows.append({
                    "_row_type": "mod_header",
                    "rack_num":  rack_num,
                    "rack_name": rack_name,
                    "slot_num":  slot_num,
                    "model":     mod["model"],
                    "desc":      mod["desc"],
                    "sig_type":  sig_type,
                    "controller": prefix,
                })

                for ch in range(mod["qty"]):
                    addr = f"{sig_type}_{str(rack_num).zfill(2)}.{str(slot_num).zfill(2)}"
                    term = f"{sig_type}-{str(slot_num).zfill(3)}/{ch // 8 + 1}"

                    if pending_idx < len(pending):
                        d = pending[pending_idx]
                        io_rows.append({
                            "_row_type":  "signal",
                            "controller": addr,
                            "tag":        d["tag"],
                            "eq_desc":    d["eq_desc"],
                            "sig_desc":   d["sig_desc"],
                            "system":     d["system"],
                            "rack":       rack_num,
                            "rack_name":  rack_name,
                            "slot":       slot_num,
                            "ch":         ch,
                            "sig_type":   sig_type,
                            "wire_color": WIRE_COLORS[ch % len(WIRE_COLORS)],
                            "term_block": term,
                            "t1": f"{ch}A",
                            "t2": f"{ch}B",
                            "power": POWER_MAP.get(sig_type, ""),
                        })
                        d["_assigned"] = True
                        pending_idx += 1
                    else:
                        io_rows.append({
                            "_row_type":  "spare",
                            "controller": addr,
                            "sig_type":   sig_type,
                            "rack":       rack_num,
                            "rack_name":  rack_name,
                            "slot":       slot_num,
                            "ch":         ch,
                            "wire_color": WIRE_COLORS[ch % len(WIRE_COLORS)],
                            "term_block": term,
                            "t1": f"{ch}A",
                            "t2": f"{ch}B",
                            "power": POWER_MAP.get(sig_type, ""),
                        })

        global_slot += len(modules)

    # Collect any signals that couldn't fit in any rack
    warnings = []
    all_unassigned = []
    for sig_type in TYPE_ORDER:
        for d in demand[sig_type]:
            if not d["_assigned"]:
                all_unassigned.append(d)
                w = (f"Unassigned — {d['tag']} / {d['sig_desc']} ({sig_type}): "
                     f"no available channels across all racks.")
                warnings.append(w)
                log(f"⚠  {w}")

    # Build final_rows with rack headers injected before first module of each rack
    final_rows = []
    seen_racks = set()
    for r in io_rows:
        if r["_row_type"] in ("mod_header", "signal", "spare"):
            actual_rack = r.get("rack_name", f"RACK {str(r.get('rack', 0)).zfill(2)}")
            if actual_rack not in seen_racks:
                final_rows.append({
                    "_row_type":  "rack_header",
                    "rack_name":  actual_rack,
                    "controller": r.get("controller", prefix),
                })
                seen_racks.add(actual_rack)
        final_rows.append(r)

    for d in all_unassigned:
        final_rows.append({"_row_type": "unassigned", **d})

    counts = {t: sum(1 for r in final_rows if r["_row_type"] == "signal" and r.get("sig_type") == t)
              for t in TYPE_ORDER}
    total = sum(counts.values())
    log(f"✓ Generation complete — DI:{counts['DI']} DO:{counts['DO']} AI:{counts['AI']} AO:{counts['AO']} — Total:{total}")

    return final_rows, counts, warnings


def build_io_list_revision(proj, module_db, mfg_channels, eq_types, matrix, sig_order, equipment, log_fn=None):
    """
    Revision mode: update an existing manufacturing IO list with a new equipment list.
    Physical channel positions are preserved. Removed equipment becomes SPARE.
    New equipment is assigned to the first available SPARE channels of matching type.
    Raises ValueError if there are insufficient spare channels.
    """
    def log(msg):
        if log_fn:
            log_fn(msg)

    prefix = proj.get("Controller Tag Prefix", "PLC-1")
    TYPE_ORDER = ["DI", "DO", "AI", "AO"]

    # Build new demand from updated equipment list
    new_demand = []
    unknown_types = set()
    for eq in equipment:
        eq_type = eq["type"]
        if eq_type not in eq_types:
            unknown_types.add(eq_type)
            continue
        for sig_name in sig_order:
            sig_desc = matrix.get(sig_name, {}).get(eq_type, "")
            if not sig_desc.strip():
                continue
            st = get_sig_type(sig_name)
            if not st:
                continue
            new_demand.append({
                "tag":      eq["tag"],
                "eq_desc":  eq["desc"],
                "sig_desc": sig_desc,
                "system":   eq["system"],
                "sig_type": st,
                "_assigned": False,
            })

    if unknown_types:
        log(f"⚠  Unknown equipment types: {', '.join(unknown_types)}")

    # Index new demand by (tag, sig_desc) for O(1) lookup
    demand_lookup = {}
    for d in new_demand:
        key = (d["tag"], d["sig_desc"])
        if key not in demand_lookup:
            demand_lookup[key] = d

    # First pass: decide keep vs spare for each channel in the mfg IO list
    revised = []
    for ch in mfg_channels:
        key = (ch["tag"], ch["sig_desc"])
        if (not ch["is_spare"]
                and key in demand_lookup
                and not demand_lookup[key]["_assigned"]):
            revised.append({**ch, "status": "signal"})
            demand_lookup[key]["_assigned"] = True
        else:
            revised.append({
                **ch,
                "status":   "spare",
                "tag":      "",
                "sig_desc": "SPARE",
                "eq_desc":  "",
                "system":   "",
                "is_spare": True,
            })

    # Gather unassigned new signals
    unassigned_new = [d for d in new_demand if not d["_assigned"]]
    log(f"Signals kept from manufacturing list: {sum(1 for r in revised if r['status'] == 'signal')}")
    log(f"New signals to assign: {len(unassigned_new)}")

    # Count spare channels by type
    spare_indices_by_type = {t: [] for t in TYPE_ORDER}
    for i, ch in enumerate(revised):
        if ch["status"] == "spare" and ch["sig_type"] in spare_indices_by_type:
            spare_indices_by_type[ch["sig_type"]].append(i)

    # Check capacity before making any assignments
    needed_by_type = {}
    for d in unassigned_new:
        needed_by_type[d["sig_type"]] = needed_by_type.get(d["sig_type"], 0) + 1

    shortage = {}
    for sig_type, needed in needed_by_type.items():
        available = len(spare_indices_by_type.get(sig_type, []))
        if needed > available:
            shortage[sig_type] = {"needed": needed, "available": available, "short": needed - available}

    if shortage:
        DEFAULT_CH = {"DI": 8, "DO": 8, "AI": 8, "AO": 4}
        lines = []
        for sig_type, info in shortage.items():
            ch_per_mod = DEFAULT_CH.get(sig_type, 8)
            mods_needed = (info["short"] + ch_per_mod - 1) // ch_per_mod
            lines.append(
                f"  • {sig_type}: need {info['short']} more channel(s) "
                f"→ add at least {mods_needed} more {sig_type} module(s) to HARDWARE_CONFIG"
            )
        raise ValueError(
            "Cannot complete revision — insufficient spare channels:\n"
            + "\n".join(lines)
            + "\n\nTo resolve:\n"
              "  1. Open your input workbook\n"
              "  2. In HARDWARE_CONFIG, add the required modules\n"
              "  3. Run revision mode again with the updated workbook"
        )

    # Assign new signals to spare channels in order
    for d in unassigned_new:
        spare_list = spare_indices_by_type[d["sig_type"]]
        idx = spare_list.pop(0)
        ch = revised[idx]
        revised[idx] = {
            **ch,
            "status":   "signal",
            "tag":      d["tag"],
            "eq_desc":  d["eq_desc"],
            "sig_desc": d["sig_desc"],
            "system":   d["system"],
            "is_spare": False,
        }
        d["_assigned"] = True
        log(f"  Assigned {d['tag']} / {d['sig_desc']} → {ch['sig_type']} slot {ch['slot_num']} ch {ch['ch']}")

    # Build final_rows preserving physical order from the manufacturing IO list
    final_rows = []
    seen_racks = set()
    prev_slot_key = None

    for ch in revised:
        rack_key  = ch["rack_name"]
        slot_key  = (ch["rack_num"], ch["slot_num"])

        if rack_key not in seen_racks:
            final_rows.append({
                "_row_type":  "rack_header",
                "rack_name":  rack_key,
                "controller": prefix,
            })
            seen_racks.add(rack_key)

        if slot_key != prev_slot_key:
            mod_info = module_db.get(ch["model"], {})
            final_rows.append({
                "_row_type":  "mod_header",
                "rack_num":   ch["rack_num"],
                "rack_name":  ch["rack_name"],
                "slot_num":   ch["slot_num"],
                "model":      ch["model"],
                "desc":       mod_info.get("desc", ch.get("mod_desc", "")),
                "sig_type":   ch["sig_type"],
                "controller": prefix,
            })
            prev_slot_key = slot_key

        if ch["status"] == "signal":
            final_rows.append({
                "_row_type":  "signal",
                "controller": ch["io_addr"],
                "tag":        ch["tag"],
                "eq_desc":    ch["eq_desc"],
                "sig_desc":   ch["sig_desc"],
                "system":     ch["system"],
                "rack":       ch["rack_num"],
                "rack_name":  ch["rack_name"],
                "slot":       ch["slot_num"],
                "ch":         ch["ch"],
                "sig_type":   ch["sig_type"],
                "wire_color": ch["wire_color"],
                "term_block": ch["term_block"],
                "t1":         ch["t1"],
                "t2":         ch["t2"],
                "power":      ch["power"],
            })
        else:
            final_rows.append({
                "_row_type":  "spare",
                "controller": ch["io_addr"],
                "sig_type":   ch["sig_type"],
                "rack":       ch["rack_num"],
                "rack_name":  ch["rack_name"],
                "slot":       ch["slot_num"],
                "ch":         ch["ch"],
                "wire_color": ch["wire_color"],
                "term_block": ch["term_block"],
                "t1":         ch["t1"],
                "t2":         ch["t2"],
                "power":      ch["power"],
            })

    counts = {t: sum(1 for r in final_rows if r["_row_type"] == "signal" and r.get("sig_type") == t)
              for t in TYPE_ORDER}
    total = sum(counts.values())
    log(f"✓ Revision complete — DI:{counts['DI']} DO:{counts['DO']} AI:{counts['AI']} AO:{counts['AO']} — Total:{total}")

    return final_rows, counts, []


# ──────────────────────────────────────────────────────────────
# EXCEL WRITER
# ──────────────────────────────────────────────────────────────
def write_output(proj, module_db, final_rows, counts, warnings, out_path, log_fn=None):
    def log(msg):
        if log_fn:
            log_fn(msg)

    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = "IO List"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    prefix = proj.get("Controller Tag Prefix", "PLC-1")
    ctrl   = proj.get("Controller Model", "")

    for col_idx, width in COL_WIDTHS.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # ── Header row 1 ──
    for col, header in enumerate(IO_COLUMNS, 1):
        c = ws.cell(1, col)
        _style(c, header, bg=C_NAVY, fg=C_BLACK, bold=True, sz=10)
    ws.row_dimensions[1].height = 38

    # ── Sub-header row 2 ──
    for col in range(1, 33):
        c = ws.cell(2, col)
        _style(c, None, bg=C_NAVY, fg=C_BLACK, bold=True, sz=10)
    ws.cell(2, 24).value = "Min"
    ws.cell(2, 25).value = "Max"
    ws.cell(2, 27).value = "LL"
    ws.cell(2, 28).value = "L"
    ws.cell(2, 29).value = "H"
    ws.cell(2, 30).value = "HH"
    ws.row_dimensions[2].height = 15

    ws.merge_cells("X1:Y1")
    ws.merge_cells("AA1:AD1")

    # ── Controller row ──
    current_row = 3
    _style(ws.cell(current_row, 1), prefix, bg=C_ORANGE, bold=True)
    _style(ws.cell(current_row, 4), ctrl,   bg=C_ORANGE, bold=True)
    _style(ws.cell(current_row, 6), 0,      bg=C_ORANGE, bold=True)
    _style(ws.cell(current_row, 7), 0,      bg=C_ORANGE, bold=True)
    _style(ws.cell(current_row, 9), "Processor", bg=C_ORANGE, bold=True)
    for col in range(1, 33):
        c = ws.cell(current_row, col)
        if not c.value:
            c.fill = _fill(C_ORANGE)
        c.border = _border()
    ws.row_dimensions[current_row].height = 18
    current_row += 1

    SIG_BG = {"DI": C_LBLUE, "DO": C_LGREEN, "AI": C_LYELLOW, "AO": C_LPINK}

    for r in final_rows:
        rt = r["_row_type"]

        if rt == "rack_header":
            pass

        elif rt == "mod_header":
            rack_name = r["rack_name"]
            sig_type  = r["sig_type"]
            _style(ws.cell(current_row, 1), rack_name,
                   bg=C_NAVY, fg=C_BLACK, bold=True)
            _style(ws.cell(current_row, 4), f"{r['model']} — {r['desc']}",
                   bg=C_NAVY, fg=C_BLACK, bold=True, h="left")
            _style(ws.cell(current_row, 6), r["rack_num"],
                   bg=C_NAVY, fg=C_BLACK, bold=True)
            _style(ws.cell(current_row, 7), r["slot_num"],
                   bg=C_NAVY, fg=C_BLACK, bold=True)
            _style(ws.cell(current_row, 9),
                   {"DI": "DIM", "DO": "DOM", "AI": "AIM", "AO": "AOM"}.get(sig_type, sig_type),
                   bg=C_NAVY, fg=C_BLACK, bold=True)
            for col in [2,3,5,8,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32]:
                c = ws.cell(current_row, col)
                c.fill   = _fill(C_NAVY)
                c.border = _border()
            ws.row_dimensions[current_row].height = 16
            current_row += 1

        elif rt == "signal":
            sig_type = r["sig_type"]
            bg       = SIG_BG.get(sig_type, C_WHITE)
            row_data = [
                r["controller"], r["tag"], "", r["eq_desc"], r["sig_desc"],
                r["rack"], r["slot"], r["ch"], sig_type, r["wire_color"],
                r["term_block"], r["t1"], r["t2"], "",
                r["power"], "S", "",
                r["system"], "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            ]
            for col, val in enumerate(row_data, 1):
                c = ws.cell(current_row, col)
                _style(c, val if val != "" else None,
                       bg=bg if col == 1 else None,
                       bold=(col == 1),
                       h="center" if col not in (4, 5, 18) else "left",
                       sz=10)
            ws.row_dimensions[current_row].height = 16
            current_row += 1

        elif rt == "spare":
            sig_type = r["sig_type"]
            bg       = SIG_BG.get(sig_type, C_WHITE)
            row_data = [
                r.get("controller", ""), "", "", "", "SPARE",
                r.get("rack", ""), r.get("slot", ""), r.get("ch", ""), sig_type,
                r.get("wire_color", ""), r.get("term_block", ""),
                r.get("t1", ""), r.get("t2", ""), "",
                r.get("power", ""), "", "",
                "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            ]
            for col, val in enumerate(row_data, 1):
                c = ws.cell(current_row, col)
                _style(c, val if val != "" else None, bg=bg, sz=10,
                       h="center" if col != 5 else "left")
            ws.row_dimensions[current_row].height = 16
            current_row += 1

        elif rt == "unassigned":
            for col in range(1, 33):
                c = ws.cell(current_row, col)
                c.fill   = _fill(C_UNASSIGNED)
                c.border = _border()
                c.font   = _font(C_RED_HDR, bold=True, sz=10)
            ws.cell(current_row, 1).value = "UNASSIGNED"
            ws.cell(current_row, 2).value = r["tag"]
            ws.cell(current_row, 4).value = r["eq_desc"]
            ws.cell(current_row, 5).value = r["sig_desc"]
            ws.cell(current_row, 9).value = get_sig_type(r["sig_row"])
            ws.row_dimensions[current_row].height = 16
            current_row += 1

    # ── SUMMARY sheet ──
    ws2 = wb.create_sheet("Generation Summary")
    ws2.column_dimensions["A"].width = 35
    ws2.column_dimensions["B"].width = 25
    ws2.column_dimensions["C"].width = 50

    def s2h(cell, val, bg=C_NAVY, fg=C_BLACK):
        cell.value = val
        cell.font  = _font(fg, 11, True)
        cell.fill  = _fill(bg)
        cell.border = _border()
        cell.alignment = _align()

    def s2b(cell, val, bg=None, bold=False, fg=C_BLACK):
        cell.value = val
        cell.font  = _font(fg, 10, bold)
        if bg: cell.fill = _fill(bg)
        cell.border = _border()
        cell.alignment = _align("left")

    now = datetime.now()
    gen_info = [
        ("Project Name",    proj.get("Project Name",          "")),
        ("Project Number",  proj.get("Project Number",        "")),
        ("Controller",      proj.get("Controller Tag Prefix", "") + "  " + proj.get("Controller Model", "")),
        ("Engineer",        proj.get("Engineer Name",         "")),
        ("Generation Date", now.strftime("%Y-%m-%d")),
        ("Generation Time", now.strftime("%H:%M:%S")),
    ]

    r2 = 1
    s2h(ws2.cell(r2, 1), "Generation Report", C_NAVY)
    ws2.merge_cells(f"A{r2}:C{r2}")
    ws2.row_dimensions[r2].height = 25
    r2 += 1

    for label, val in gen_info:
        s2h(ws2.cell(r2, 1), label, C_LGRAY, C_BLACK)
        s2b(ws2.cell(r2, 2), val)
        ws2.merge_cells(f"B{r2}:C{r2}")
        r2 += 1

    r2 += 1
    s2h(ws2.cell(r2, 1), "IO Signal Counts", C_NAVY)
    ws2.merge_cells(f"A{r2}:C{r2}")
    ws2.row_dimensions[r2].height = 20
    r2 += 1

    type_fill = {"DI": C_LBLUE, "DO": C_LGREEN, "AI": C_LYELLOW, "AO": C_LPINK}
    for sig_type, count in counts.items():
        s2h(ws2.cell(r2, 1), sig_type, type_fill.get(sig_type, C_WHITE), C_BLACK)
        s2b(ws2.cell(r2, 2), count)
        r2 += 1
    s2h(ws2.cell(r2, 1), "TOTAL", C_NAVY)
    s2b(ws2.cell(r2, 2), sum(counts.values()), bold=True)
    r2 += 2

    if warnings:
        s2h(ws2.cell(r2, 1), f"Warnings ({len(warnings)})", C_RED_HDR, C_WHITE)
        ws2.merge_cells(f"A{r2}:C{r2}")
        ws2.row_dimensions[r2].height = 20
        r2 += 1
        for w in warnings:
            c = ws2.cell(r2, 1)
            c.value = w
            c.font  = _font(C_RED_HDR, 9)
            c.fill  = _fill(C_UNASSIGNED)
            c.border = _border()
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws2.merge_cells(f"A{r2}:C{r2}")
            ws2.row_dimensions[r2].height = 30
            r2 += 1
    else:
        s2h(ws2.cell(r2, 1), "No warnings — all signals assigned successfully ✓", C_LGREEN, "FF1E4D2B")
        ws2.merge_cells(f"A{r2}:C{r2}")

    wb.save(out_path)
    log(f"✓ File saved: {out_path}")


# ──────────────────────────────────────────────────────────────
# GUI
# ──────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("IO List Generator — Allen-Bradley / Rockwell")
        self.resizable(True, True)
        self.minsize(680, 640)
        self.input_path = None
        self.rev_file_path = None
        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = 780, 680
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        FONT_H  = ("Arial", 11, "bold")
        FONT_B  = ("Arial", 10)
        FONT_SM = ("Arial", 9)
        BG_MAIN = "#f5f5f5"
        BG_NAV  = "#c0c0c0"
        FG_NAV  = "#000000"

        self.configure(bg=BG_MAIN)

        # Title bar
        title_frame = tk.Frame(self, bg=BG_NAV, height=52)
        title_frame.pack(fill="x")
        title_frame.pack_propagate(False)
        tk.Label(title_frame, text="  IO List Generator", font=("Arial", 14, "bold"),
                 bg=BG_NAV, fg=FG_NAV, anchor="w").pack(side="left", padx=10, pady=10)
        tk.Label(title_frame, text="Allen-Bradley / Rockwell Automation",
                 font=("Arial", 9), bg=BG_NAV, fg="#444444").pack(side="left", padx=2)

        # Main content
        content = tk.Frame(self, bg=BG_MAIN, padx=18, pady=14)
        content.pack(fill="both", expand=True)

        # ── Step 1: Input file ──
        self._section(content, "1  Select Input Workbook", FONT_H)
        file_row = tk.Frame(content, bg=BG_MAIN)
        file_row.pack(fill="x", pady=(0, 10))
        self.lbl_file = tk.Label(file_row, text="No file selected", font=FONT_SM,
                                 bg="#e8eef4", fg="#444", relief="flat",
                                 anchor="w", padx=8, pady=5)
        self.lbl_file.pack(side="left", fill="x", expand=True)
        tk.Button(file_row, text="Browse…", font=FONT_B, bg=BG_NAV, fg=FG_NAV,
                  relief="flat", padx=14, pady=4, cursor="hand2",
                  command=self._browse_input).pack(side="left", padx=(6, 0))
        tk.Button(file_row, text="Create template", font=FONT_SM, bg="#f0c040", fg="#1a1a1a",
                  relief="flat", padx=10, pady=4, cursor="hand2",
                  command=self._create_template).pack(side="left", padx=(6, 0))

        # ── Step 2: Revision mode ──
        self._section(content, "2  Revision Mode (optional)", FONT_H)
        rev_check_row = tk.Frame(content, bg=BG_MAIN)
        rev_check_row.pack(fill="x", pady=(0, 4))
        self.revision_var = tk.BooleanVar(value=False)
        tk.Checkbutton(rev_check_row,
                       text="Enable revision mode — attach a manufacturing IO list to update assignments",
                       variable=self.revision_var,
                       command=self._toggle_revision,
                       font=FONT_B, bg=BG_MAIN, fg="#1a1a1a",
                       activebackground=BG_MAIN).pack(side="left")

        rev_attach_row = tk.Frame(content, bg=BG_MAIN)
        rev_attach_row.pack(fill="x", pady=(0, 10))
        self.lbl_rev_file = tk.Label(rev_attach_row, text="No manufacturing IO list attached",
                                     font=FONT_SM, bg="#e8eef4", fg="#888", relief="flat",
                                     anchor="w", padx=8, pady=5)
        self.lbl_rev_file.pack(side="left", fill="x", expand=True)
        self.btn_rev_attach = tk.Button(rev_attach_row, text="Attach…", font=FONT_B,
                                        bg="#888888", fg="white", relief="flat",
                                        padx=14, pady=4, cursor="hand2",
                                        command=self._browse_rev_file, state="disabled")
        self.btn_rev_attach.pack(side="left", padx=(6, 0))

        # ── Step 3: Output folder ──
        self._section(content, "3  Output Folder", FONT_H)
        out_row = tk.Frame(content, bg=BG_MAIN)
        out_row.pack(fill="x", pady=(0, 10))
        self.out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "IO List")
        self.lbl_out = tk.Label(out_row, text=self.out_dir, font=FONT_SM,
                                bg="#e8eef4", fg="#444", relief="flat",
                                anchor="w", padx=8, pady=5)
        self.lbl_out.pack(side="left", fill="x", expand=True)
        tk.Button(out_row, text="Change…", font=FONT_B, bg=BG_NAV, fg=FG_NAV,
                  relief="flat", padx=14, pady=4, cursor="hand2",
                  command=self._browse_out).pack(side="left", padx=(6, 0))

        # ── Generate button ──
        gen_frame = tk.Frame(content, bg=BG_MAIN)
        gen_frame.pack(fill="x", pady=(6, 10))
        self.btn_gen = tk.Button(gen_frame, text="▶  Generate IO List", font=("Arial", 11, "bold"),
                                 bg="#1E7C3A", fg="white", relief="flat",
                                 padx=22, pady=8, cursor="hand2",
                                 command=self._run_generate)
        self.btn_gen.pack(side="left")
        self.btn_open = tk.Button(gen_frame, text="Open output folder", font=FONT_B,
                                  bg=BG_NAV, fg=FG_NAV, relief="flat",
                                  padx=14, pady=8, cursor="hand2",
                                  command=self._open_out_folder, state="disabled")
        self.btn_open.pack(side="left", padx=(10, 0))

        # ── Progress bar ──
        self.progress = ttk.Progressbar(content, mode="indeterminate", length=300)
        self.progress.pack(fill="x", pady=(0, 6))

        # ── Log box ──
        self._section(content, "Log", FONT_H)
        self.log_box = scrolledtext.ScrolledText(content, height=11, font=("Courier New", 9),
                                                 bg="#1a1a2e", fg="#d0e8ff",
                                                 relief="flat", padx=8, pady=6,
                                                 insertbackground="white")
        self.log_box.pack(fill="both", expand=True)
        self.log_box.configure(state="disabled")

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(self, textvariable=self.status_var, font=FONT_SM,
                 bg=BG_NAV, fg=FG_NAV, anchor="w", padx=10).pack(
                 side="bottom", fill="x")

    def _section(self, parent, text, font):
        tk.Frame(parent, bg="#d0d0d0", height=1).pack(fill="x", pady=(4, 2))
        tk.Label(parent, text=text, font=font, bg="#f5f5f5", fg="#333333",
                 anchor="w").pack(fill="x")

    def _toggle_revision(self):
        if self.revision_var.get():
            self.btn_rev_attach.configure(state="normal", bg="#595959")
            self.lbl_rev_file.configure(fg="#444")
        else:
            self.btn_rev_attach.configure(state="disabled", bg="#888888")
            self.lbl_rev_file.configure(text="No manufacturing IO list attached", fg="#888")
            self.rev_file_path = None

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select Input Workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if path:
            self.input_path = path
            self.lbl_file.configure(text=f"  {os.path.basename(path)}  ({path})")

    def _browse_rev_file(self):
        path = filedialog.askopenfilename(
            title="Select Manufacturing IO List",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if path:
            self.rev_file_path = path
            self.lbl_rev_file.configure(text=f"  {os.path.basename(path)}  ({path})")

    def _browse_out(self):
        d = filedialog.askdirectory(title="Select Output Folder", initialdir=self.out_dir)
        if d:
            self.out_dir = d
            self.lbl_out.configure(text=d)

    def _open_out_folder(self):
        import subprocess, platform
        try:
            if platform.system() == "Windows":
                os.startfile(self.out_dir)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", self.out_dir])
            else:
                subprocess.Popen(["xdg-open", self.out_dir])
        except:
            pass

    def _log(self, msg):
        self.log_box.configure(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{ts}]  {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _set_status(self, msg):
        self.status_var.set(msg)

    def _create_template(self):
        from create_template import main as make_tpl
        save_path = filedialog.asksaveasfilename(
            title="Save Input Template As",
            defaultextension=".xlsx",
            initialfile="IO_List_Input_Template.xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not save_path:
            return
        try:
            make_tpl(save_path)
            messagebox.showinfo("Template created",
                                f"Template saved to:\n{save_path}\n\n"
                                "Fill in all sheets then run Generate.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _run_generate(self):
        if not self.input_path:
            messagebox.showwarning("No input file", "Please select the input workbook first.")
            return

        revision_mode = self.revision_var.get()
        if revision_mode and not self.rev_file_path:
            messagebox.showwarning("No manufacturing IO list",
                                   "Revision mode is enabled.\n"
                                   "Please attach the manufacturing IO list before generating.")
            return

        os.makedirs(self.out_dir, exist_ok=True)
        self.btn_gen.configure(state="disabled")
        self.progress.start(12)
        self._set_status("Generating…")

        def task():
            try:
                self._log(f"Reading workbook: {os.path.basename(self.input_path)}")
                proj, module_db, hw_config, eq_types, matrix, sig_order, equipment = \
                    read_workbook(self.input_path)

                self._log(f"Project: {proj.get('Project Name', '?')}  |  "
                          f"Equipment items: {len(equipment)}")

                if revision_mode:
                    self._log(f"Revision mode — parsing: {os.path.basename(self.rev_file_path)}")
                    mfg_channels = parse_manufacturing_io_list(self.rev_file_path)
                    self._log(f"Manufacturing IO list: {len(mfg_channels)} channels found")
                    final_rows, counts, warnings = build_io_list_revision(
                        proj, module_db, mfg_channels, eq_types, matrix, sig_order, equipment,
                        log_fn=self._log
                    )
                else:
                    self._log(f"Racks configured: {len(set(r['rack'] for r in hw_config))}  |  "
                              f"Module rows: {len(hw_config)}  |  Signal rows: {len(sig_order)}")
                    final_rows, counts, warnings = build_io_list(
                        proj, module_db, hw_config, eq_types, matrix, sig_order, equipment,
                        log_fn=self._log
                    )

                now = datetime.now()
                proj_name = proj.get("Project Name", "Project").replace(" ", "_")
                suffix = "_Revision" if revision_mode else ""
                fname = f"{proj_name}_{now.strftime('%Y%m%d_%H%M%S')}{suffix}_IO_List.xlsx"
                out_path = os.path.join(self.out_dir, fname)

                write_output(proj, module_db, final_rows, counts, warnings, out_path,
                             log_fn=self._log)
                self.after(0, lambda: self._on_done(out_path, warnings))

            except Exception as e:
                import traceback
                self._log(f"✗ ERROR: {e}")
                self._log(traceback.format_exc())
                self.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=task, daemon=True).start()

    def _on_done(self, out_path, warnings):
        self.progress.stop()
        self.btn_gen.configure(state="normal")
        self.btn_open.configure(state="normal")
        self._set_status(f"Done — {os.path.basename(out_path)}")
        if warnings:
            messagebox.showwarning("Completed with warnings",
                                   f"IO List generated with {len(warnings)} warning(s).\n"
                                   f"Check the 'Generation Summary' sheet for details.\n\n"
                                   f"File: {os.path.basename(out_path)}")
        else:
            messagebox.showinfo("Success",
                                f"IO List generated successfully!\n\nFile: {os.path.basename(out_path)}")

    def _on_error(self, msg):
        self.progress.stop()
        self.btn_gen.configure(state="normal")
        self._set_status("Error — see log")
        messagebox.showerror("Generation failed", f"An error occurred:\n\n{msg}\n\nSee log for details.")


if __name__ == "__main__":
    app = App()
    app.mainloop()
