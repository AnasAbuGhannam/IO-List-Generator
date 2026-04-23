"""
Engineering Automation Suite — Allen-Bradley / Rockwell Automation
Generates IO lists, EPLAN exports, L5X hardware configs, tag objects, and mirroring code.
"""

APP_VERSION = "1.0.0"
APP_DATE    = "2026-04-23"
APP_AUTHOR  = "Anas Abu Ghannam"
APP_COMPANY = "SAM Engineering Co."
APP_DESC    = (
    "Engineering Automation Suite automates the generation of\n"
    "PLC engineering deliverables for Allen-Bradley / Rockwell\n"
    "Automation systems, including IO lists, EPLAN exports,\n"
    "L5X hardware configurations, Studio 5000 tag objects,\n"
    "and PLC mirroring code.\n\n"
    "Designed to minimize engineering time and project costs\n"
    "while maximizing quality and delivery speed."
)
import csv
import os
import re
import random
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
C_NAVY       = "FFC0C0C0"
C_ORANGE     = "FFFFC000"
C_LGRAY      = "FFD9D9D9"
C_DGRAY      = "FF595959"
C_WHITE      = "FFFFFFFF"
C_BLACK      = "FF000000"
C_LBLUE      = "FFD6E4F0"
C_LGREEN     = "FFE2EFDA"
C_LYELLOW    = "FFFFF2CC"
C_LPINK      = "FFFCE4D6"
C_RED_HDR    = "FFC00000"
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
    1:16.0,  2:20.29, 3:9.29,  4:59.43, 5:23.57, 6:3.43,  7:13.0,  8:13.0,
    9:11.14, 10:14.0, 11:11.86,12:6.71, 13:6.0,  14:6.0,  15:7.14, 16:11.71,
    17:9.0,  18:8.14, 19:7.57, 20:15.14,21:9.0,  22:5.86, 23:6.43, 24:4.57,
    25:4.86, 26:8.0,  27:3.57, 28:2.14, 29:2.29, 30:3.57, 31:6.43, 32:6.0
}

TYPE_ORDER = ["DI", "DO", "AI", "AO"]
POWER_MAP  = {"DI": "24VDC", "DO": "NO", "AI": "", "AO": ""}

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

    # SIGNAL_MATRIX
    ws_mat = wb["SIGNAL_MATRIX"]
    mat_header = [str(c).strip() if c else "" for c in
                  next(ws_mat.iter_rows(min_row=1, max_row=1, values_only=True))]
    eq_types  = mat_header[3:]
    matrix    = {}
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

    # PLC_ATTRIBUTES (optional sheet added by user)
    plc_attributes = {}
    if "PLC_ATTRIBUTES" in wb.sheetnames:
        ws_plc = wb["PLC_ATTRIBUTES"]
        sig_col = attr_col = None
        for row in ws_plc.iter_rows(min_row=1, max_row=10, values_only=True):
            for i, v in enumerate(row):
                if v is not None and str(v).strip() == "Signal Description":
                    sig_col = i
                if v is not None and str(v).strip() == "PLC Attribute":
                    attr_col = i
            if sig_col is not None and attr_col is not None:
                break
        if sig_col is not None and attr_col is not None:
            for row in ws_plc.iter_rows(min_row=2, values_only=True):
                s = row[sig_col] if sig_col < len(row) else None
                a = row[attr_col] if attr_col < len(row) else None
                if s and a:
                    plc_attributes[str(s).strip()] = str(a).strip()

    return proj, module_db, hw_config, eq_types, matrix, sig_order, equipment, plc_attributes


def _quick_read_racks(path):
    """Lightweight read of rack names only from HARDWARE_CONFIG sheet."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        if "HARDWARE_CONFIG" not in wb.sheetnames:
            return []
        ws = wb["HARDWARE_CONFIG"]
        racks, seen = [], set()
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if row[0]:
                r = str(row[0]).strip()
                if r and r not in seen:
                    racks.append(r)
                    seen.add(r)
        wb.close()
        return racks
    except:
        return []


# ──────────────────────────────────────────────────────────────
# MODULE ASSIGNMENT HELPER  (used by IO list, EPLAN, L5X)
# ──────────────────────────────────────────────────────────────
def get_rack_module_assignments(hw_config, module_db):
    """
    Returns (assignments_by_rack, rack_order, slot_lookup) where:
      assignments_by_rack = {rack_name: [{model, type, name, slot, qty}, ...]}
      slot_lookup = {(rack_name, slot_num): {name, type, mod_idx_in_type}}
    Modules sorted DI→DO→AI→AO within each rack; global type counters across all racks.
    """
    racks = {}
    rack_order = []
    for row in hw_config:
        rack_name = row["rack"]
        if rack_name not in racks:
            racks[rack_name] = []
            rack_order.append(rack_name)
        info = module_db.get(row["model"])
        if not info:
            continue
        for _ in range(row["qty"]):
            racks[rack_name].append({"model": row["model"], "type": info["type"], "qty": info["qty"]})

    global_type_counter = {t: 0 for t in TYPE_ORDER}
    assignments_by_rack = {}
    slot_lookup = {}
    global_slot = 1

    for rack_name in rack_order:
        modules = racks[rack_name]
        rack_slot_start = global_slot
        rack_by_type = {t: [m for m in modules if m["type"] == t] for t in TYPE_ORDER}
        rack_assignments = []

        for sig_type in TYPE_ORDER:
            for mod_idx, mod in enumerate(rack_by_type[sig_type]):
                global_type_counter[sig_type] += 1
                mod_name = f"{sig_type}_{str(global_type_counter[sig_type]).zfill(2)}"
                slot_num = rack_slot_start + sum(
                    len(rack_by_type[t]) for t in TYPE_ORDER[:TYPE_ORDER.index(sig_type)]
                ) + mod_idx
                entry = {
                    "model":   mod["model"],
                    "type":    sig_type,
                    "name":    mod_name,
                    "slot":    slot_num,
                    "qty":     mod["qty"],
                    "mod_idx": global_type_counter[sig_type],
                }
                rack_assignments.append(entry)
                slot_lookup[(rack_name, slot_num)] = entry

        assignments_by_rack[rack_name] = rack_assignments
        global_slot += len(modules)

    return assignments_by_rack, rack_order, slot_lookup


# ──────────────────────────────────────────────────────────────
# MANUFACTURING IO LIST PARSER
# ──────────────────────────────────────────────────────────────
def parse_manufacturing_io_list(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    if "IO List" not in wb.sheetnames:
        raise ValueError("Attached file does not contain an 'IO List' sheet.")
    ws = wb["IO List"]

    TYPE_SUFFIX_REV = {"DIM": "DI", "DOM": "DO", "AIM": "AI", "AOM": "AO"}
    channels = []
    current_rack_name = current_model = current_mod_desc = ""
    current_rack_num = current_slot_num = 0

    for row in ws.iter_rows(min_row=3, values_only=True):
        if not any(v for v in row):
            continue

        def cell(i, row=row):
            v = row[i] if i < len(row) else None
            return str(v).strip() if v is not None else ""

        type_val = cell(8)
        if type_val == "Processor":
            continue
        if type_val in TYPE_SUFFIX_REV:
            current_rack_name = cell(0)
            col4 = cell(3)
            if " — " in col4:
                parts = col4.split(" — ", 1)
                current_model, current_mod_desc = parts[0].strip(), parts[1].strip()
            else:
                current_model, current_mod_desc = col4, ""
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
            tag      = cell(1)
            sig_desc = cell(4)
            channels.append({
                "rack_name":  current_rack_name,
                "rack_num":   current_rack_num,
                "slot_num":   current_slot_num,
                "ch":         int(cell(7)) if cell(7).isdigit() else 0,
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
                "is_spare":   (not tag or sig_desc == "SPARE"),
            })

    if not channels:
        raise ValueError("No IO channels found in the attached manufacturing IO list.")
    return channels


# ──────────────────────────────────────────────────────────────
# CORE ENGINE
# ──────────────────────────────────────────────────────────────
def get_sig_type(sig_row_name):
    n = sig_row_name.strip().upper()
    for t in TYPE_ORDER:
        if n.startswith(t):
            return t
    return None


def build_io_list(proj, module_db, hw_config, eq_types, matrix, sig_order, equipment, log_fn=None):
    def log(msg):
        if log_fn: log_fn(msg)

    prefix = proj.get("Controller Tag Prefix", "PLC-1")
    demand = {t: [] for t in TYPE_ORDER}
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
                "tag": eq["tag"], "eq_desc": eq["desc"], "sig_desc": sig_desc,
                "system": eq["system"], "eq_type": eq_type, "sig_row": sig_name,
                "_assigned": False,
            })

    if unknown_types:
        log(f"⚠  Unknown equipment types: {', '.join(unknown_types)}")

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
            racks[rack_name].append({"model": row["model"], "desc": info["desc"],
                                     "type": info["type"], "qty": info["qty"]})

    io_rows = []
    global_slot = 1

    for rack_name in rack_order:
        rack_num = rack_order.index(rack_name) + 1
        modules  = racks[rack_name]
        rack_slot_start = global_slot
        rack_by_type = {t: [m for m in modules if m["type"] == t] for t in TYPE_ORDER}

        for sig_type in TYPE_ORDER:
            type_modules = rack_by_type[sig_type]
            if not type_modules:
                continue
            pending = [d for d in demand[sig_type] if not d["_assigned"]]
            pending_idx = 0

            for mod_idx, mod in enumerate(type_modules):
                slot_num = rack_slot_start + sum(
                    len(rack_by_type[t]) for t in TYPE_ORDER[:TYPE_ORDER.index(sig_type)]
                ) + mod_idx

                # R@@.S##.%%**  @@=rack seq, ##=slot, %%=type, **=module order in type per rack
                addr = f"R{str(rack_num).zfill(2)}.S{str(slot_num).zfill(2)}.{sig_type}{str(mod_idx + 1).zfill(2)}"

                io_rows.append({
                    "_row_type": "mod_header",
                    "rack_num": rack_num, "rack_name": rack_name,
                    "slot_num": slot_num, "model": mod["model"],
                    "desc": mod["desc"], "sig_type": sig_type, "controller": prefix,
                })

                for ch in range(mod["qty"]):
                    term = f"{sig_type}-{str(slot_num).zfill(3)}/{ch // 8 + 1}"
                    if pending_idx < len(pending):
                        d = pending[pending_idx]
                        io_rows.append({
                            "_row_type": "signal",
                            "controller": addr, "tag": d["tag"],
                            "eq_desc": d["eq_desc"], "sig_desc": d["sig_desc"],
                            "system": d["system"], "rack": rack_num, "rack_name": rack_name,
                            "slot": slot_num, "ch": ch, "sig_type": sig_type,
                            "wire_color": WIRE_COLORS[ch % len(WIRE_COLORS)],
                            "term_block": term, "t1": f"{ch}A", "t2": f"{ch}B",
                            "power": POWER_MAP.get(sig_type, ""),
                        })
                        d["_assigned"] = True
                        pending_idx += 1
                    else:
                        io_rows.append({
                            "_row_type": "spare",
                            "controller": addr, "sig_type": sig_type,
                            "rack": rack_num, "rack_name": rack_name, "slot": slot_num, "ch": ch,
                            "wire_color": WIRE_COLORS[ch % len(WIRE_COLORS)],
                            "term_block": term, "t1": f"{ch}A", "t2": f"{ch}B",
                            "power": POWER_MAP.get(sig_type, ""),
                        })

        global_slot += len(modules)

    warnings = []
    final_rows = []
    seen_racks = set()

    for r in io_rows:
        if r["_row_type"] in ("mod_header", "signal", "spare"):
            rack_key = r.get("rack_name", f"RACK {str(r.get('rack', 0)).zfill(2)}")
            if rack_key not in seen_racks:
                final_rows.append({"_row_type": "rack_header", "rack_name": rack_key,
                                   "controller": r.get("controller", prefix)})
                seen_racks.add(rack_key)
        final_rows.append(r)

    for t in TYPE_ORDER:
        for d in demand[t]:
            if not d["_assigned"]:
                w = f"Unassigned — {d['tag']} / {d['sig_desc']} ({t}): no channels in any rack."
                warnings.append(w)
                log(f"⚠  {w}")
                final_rows.append({"_row_type": "unassigned", **d})

    counts = {t: sum(1 for r in final_rows if r["_row_type"] == "signal" and r.get("sig_type") == t)
              for t in TYPE_ORDER}
    log(f"✓ Generation complete — DI:{counts['DI']} DO:{counts['DO']} AI:{counts['AI']} AO:{counts['AO']} — Total:{sum(counts.values())}")
    return final_rows, counts, warnings


def build_io_list_revision(proj, module_db, mfg_channels, eq_types, matrix, sig_order, equipment, log_fn=None):
    def log(msg):
        if log_fn: log_fn(msg)

    prefix = proj.get("Controller Tag Prefix", "PLC-1")
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
                "tag": eq["tag"], "eq_desc": eq["desc"], "sig_desc": sig_desc,
                "system": eq["system"], "sig_type": st, "_assigned": False,
            })

    if unknown_types:
        log(f"⚠  Unknown equipment types: {', '.join(unknown_types)}")

    demand_lookup = {}
    for d in new_demand:
        k = (d["tag"], d["sig_desc"])
        if k not in demand_lookup:
            demand_lookup[k] = d

    revised = []
    for ch in mfg_channels:
        k = (ch["tag"], ch["sig_desc"])
        if not ch["is_spare"] and k in demand_lookup and not demand_lookup[k]["_assigned"]:
            revised.append({**ch, "status": "signal"})
            demand_lookup[k]["_assigned"] = True
        else:
            revised.append({**ch, "status": "spare", "tag": "", "sig_desc": "SPARE",
                            "eq_desc": "", "system": "", "is_spare": True})

    unassigned_new = [d for d in new_demand if not d["_assigned"]]
    log(f"Kept from manufacturing: {sum(1 for r in revised if r['status'] == 'signal')}, new to assign: {len(unassigned_new)}")

    spare_by_type = {t: [] for t in TYPE_ORDER}
    for i, ch in enumerate(revised):
        if ch["status"] == "spare" and ch["sig_type"] in spare_by_type:
            spare_by_type[ch["sig_type"]].append(i)

    needed_by_type = {}
    for d in unassigned_new:
        needed_by_type[d["sig_type"]] = needed_by_type.get(d["sig_type"], 0) + 1

    shortage = {}
    for sig_type, needed in needed_by_type.items():
        avail = len(spare_by_type.get(sig_type, []))
        if needed > avail:
            shortage[sig_type] = {"needed": needed, "available": avail, "short": needed - avail}

    if shortage:
        DEFAULT_CH = {"DI": 8, "DO": 8, "AI": 8, "AO": 4}
        lines = []
        for sig_type, info in shortage.items():
            mods = (info["short"] + DEFAULT_CH.get(sig_type, 8) - 1) // DEFAULT_CH.get(sig_type, 8)
            lines.append(f"  • {sig_type}: need {info['short']} more channel(s) → add {mods} more {sig_type} module(s) to HARDWARE_CONFIG")
        raise ValueError(
            "Cannot complete revision — insufficient spare channels:\n" + "\n".join(lines) +
            "\n\nTo resolve:\n  1. Open your input workbook\n"
            "  2. In HARDWARE_CONFIG, add the required modules\n"
            "  3. Run revision mode again with the updated workbook"
        )

    for d in unassigned_new:
        idx = spare_by_type[d["sig_type"]].pop(0)
        ch = revised[idx]
        revised[idx] = {**ch, "status": "signal", "tag": d["tag"], "eq_desc": d["eq_desc"],
                        "sig_desc": d["sig_desc"], "system": d["system"], "is_spare": False}
        d["_assigned"] = True
        log(f"  Assigned {d['tag']} / {d['sig_desc']} → {ch['sig_type']} slot {ch['slot_num']} ch {ch['ch']}")

    # Compute address lookups for new R@@.S##.%%** format
    revised_rack_order = list(dict.fromkeys(ch["rack_name"] for ch in revised))
    rack_seq_map = {name: idx + 1 for idx, name in enumerate(revised_rack_order)}

    slots_by_rack_type = {}
    for ch in revised:
        key = (ch["rack_name"], ch["sig_type"])
        if key not in slots_by_rack_type:
            slots_by_rack_type[key] = []
        if ch["slot_num"] not in slots_by_rack_type[key]:
            slots_by_rack_type[key].append(ch["slot_num"])
    slot_rank_map = {}
    for (rack, stype), slots in slots_by_rack_type.items():
        for rank, slot in enumerate(sorted(slots), 1):
            slot_rank_map[(rack, stype, slot)] = rank

    final_rows = []
    seen_racks = set()
    prev_slot_key = None

    for ch in revised:
        rack_key = ch["rack_name"]
        slot_key = (ch["rack_num"], ch["slot_num"])
        if rack_key not in seen_racks:
            final_rows.append({"_row_type": "rack_header", "rack_name": rack_key, "controller": prefix})
            seen_racks.add(rack_key)
        if slot_key != prev_slot_key:
            mod_info = module_db.get(ch["model"], {})
            final_rows.append({
                "_row_type": "mod_header", "rack_num": ch["rack_num"], "rack_name": ch["rack_name"],
                "slot_num": ch["slot_num"], "model": ch["model"],
                "desc": mod_info.get("desc", ch.get("mod_desc", "")),
                "sig_type": ch["sig_type"], "controller": prefix,
            })
            prev_slot_key = slot_key

        rack_seq  = rack_seq_map.get(ch["rack_name"], 1)
        type_rank = slot_rank_map.get((ch["rack_name"], ch["sig_type"], ch["slot_num"]), 1)
        addr = f"R{rack_seq:02d}.S{ch['slot_num']:02d}.{ch['sig_type']}{type_rank:02d}"

        if ch["status"] == "signal":
            final_rows.append({
                "_row_type": "signal", "controller": addr,
                "tag": ch["tag"], "eq_desc": ch["eq_desc"], "sig_desc": ch["sig_desc"],
                "system": ch["system"], "rack": ch["rack_num"], "rack_name": ch["rack_name"],
                "slot": ch["slot_num"], "ch": ch["ch"], "sig_type": ch["sig_type"],
                "wire_color": ch["wire_color"], "term_block": ch["term_block"],
                "t1": ch["t1"], "t2": ch["t2"], "power": ch["power"],
            })
        else:
            final_rows.append({
                "_row_type": "spare", "controller": addr, "sig_type": ch["sig_type"],
                "rack": ch["rack_num"], "rack_name": ch["rack_name"], "slot": ch["slot_num"],
                "ch": ch["ch"], "wire_color": ch["wire_color"], "term_block": ch["term_block"],
                "t1": ch["t1"], "t2": ch["t2"], "power": ch["power"],
            })

    counts = {t: sum(1 for r in final_rows if r["_row_type"] == "signal" and r.get("sig_type") == t)
              for t in TYPE_ORDER}
    log(f"✓ Revision complete — DI:{counts['DI']} DO:{counts['DO']} AI:{counts['AI']} AO:{counts['AO']} — Total:{sum(counts.values())}")
    return final_rows, counts, []


# ──────────────────────────────────────────────────────────────
# IO LIST EXCEL WRITER
# ──────────────────────────────────────────────────────────────
def write_output(proj, module_db, final_rows, counts, warnings, out_path, log_fn=None):
    def log(msg):
        if log_fn: log_fn(msg)

    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = "IO List"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    prefix = proj.get("Controller Tag Prefix", "PLC-1")
    ctrl   = proj.get("Controller Model", "")

    for col_idx, width in COL_WIDTHS.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    for col, header in enumerate(IO_COLUMNS, 1):
        _style(ws.cell(1, col), header, bg=C_NAVY, fg=C_BLACK, bold=True, sz=10)
    ws.row_dimensions[1].height = 38

    for col in range(1, 33):
        _style(ws.cell(2, col), None, bg=C_NAVY, fg=C_BLACK, bold=True, sz=10)
    for col, val in [(24,"Min"),(25,"Max"),(27,"LL"),(28,"L"),(29,"H"),(30,"HH")]:
        ws.cell(2, col).value = val
    ws.row_dimensions[2].height = 15
    ws.merge_cells("X1:Y1")
    ws.merge_cells("AA1:AD1")

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
            sig_type = r["sig_type"]
            _style(ws.cell(current_row, 1), r["rack_name"], bg=C_NAVY, fg=C_BLACK, bold=True)
            _style(ws.cell(current_row, 4), f"{r['model']} — {r['desc']}",
                   bg=C_NAVY, fg=C_BLACK, bold=True, h="left")
            _style(ws.cell(current_row, 6), r["rack_num"], bg=C_NAVY, fg=C_BLACK, bold=True)
            _style(ws.cell(current_row, 7), r["slot_num"], bg=C_NAVY, fg=C_BLACK, bold=True)
            _style(ws.cell(current_row, 9),
                   {"DI":"DIM","DO":"DOM","AI":"AIM","AO":"AOM"}.get(sig_type, sig_type),
                   bg=C_NAVY, fg=C_BLACK, bold=True)
            for col in [2,3,5,8,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32]:
                c = ws.cell(current_row, col)
                c.fill = _fill(C_NAVY)
                c.border = _border()
            ws.row_dimensions[current_row].height = 16
            current_row += 1

        elif rt == "signal":
            sig_type = r["sig_type"]
            bg = SIG_BG.get(sig_type, C_WHITE)
            row_data = [
                r["controller"], r["tag"], "", r["eq_desc"], r["sig_desc"],
                r["rack"], r["slot"], r["ch"], sig_type, r["wire_color"],
                r["term_block"], r["t1"], r["t2"], "",
                r["power"], "S", "", r["system"],
                "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            ]
            for col, val in enumerate(row_data, 1):
                c = ws.cell(current_row, col)
                _style(c, val if val != "" else None,
                       bg=bg if col == 1 else None, bold=(col == 1),
                       h="center" if col not in (4, 5, 18) else "left", sz=10)
            ws.row_dimensions[current_row].height = 16
            current_row += 1

        elif rt == "spare":
            sig_type = r["sig_type"]
            bg = SIG_BG.get(sig_type, C_WHITE)
            row_data = [
                r.get("controller",""), "", "", "", "SPARE",
                r.get("rack",""), r.get("slot",""), r.get("ch",""), sig_type,
                r.get("wire_color",""), r.get("term_block",""),
                r.get("t1",""), r.get("t2",""), "", r.get("power",""),
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
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
                c.fill = _fill(C_UNASSIGNED)
                c.border = _border()
                c.font = _font(C_RED_HDR, bold=True, sz=10)
            ws.cell(current_row, 1).value = "UNASSIGNED"
            ws.cell(current_row, 2).value = r["tag"]
            ws.cell(current_row, 4).value = r["eq_desc"]
            ws.cell(current_row, 5).value = r["sig_desc"]
            ws.cell(current_row, 9).value = get_sig_type(r["sig_row"])
            ws.row_dimensions[current_row].height = 16
            current_row += 1

    # Summary sheet
    ws2 = wb.create_sheet("Generation Summary")
    ws2.column_dimensions["A"].width = 35
    ws2.column_dimensions["B"].width = 25
    ws2.column_dimensions["C"].width = 50

    def s2h(cell, val, bg=C_NAVY, fg=C_BLACK):
        cell.value = val
        cell.font = _font(fg, 11, True)
        cell.fill = _fill(bg)
        cell.border = _border()
        cell.alignment = _align()

    def s2b(cell, val, bg=None, bold=False, fg=C_BLACK):
        cell.value = val
        cell.font = _font(fg, 10, bold)
        if bg: cell.fill = _fill(bg)
        cell.border = _border()
        cell.alignment = _align("left")

    now = datetime.now()
    gen_info = [
        ("Project Name",    proj.get("Project Name",          "")),
        ("Project Number",  proj.get("Project Number",        "")),
        ("Controller",      proj.get("Controller Tag Prefix","") + "  " + proj.get("Controller Model","")),
        ("Engineer",        proj.get("Engineer Name",         "")),
        ("Generation Date", now.strftime("%Y-%m-%d")),
        ("Generation Time", now.strftime("%H:%M:%S")),
    ]

    r2 = 1
    s2h(ws2.cell(r2,1), "Generation Report", C_NAVY)
    ws2.merge_cells(f"A{r2}:C{r2}")
    ws2.row_dimensions[r2].height = 25
    r2 += 1

    for label, val in gen_info:
        s2h(ws2.cell(r2,1), label, C_LGRAY, C_BLACK)
        s2b(ws2.cell(r2,2), val)
        ws2.merge_cells(f"B{r2}:C{r2}")
        r2 += 1

    r2 += 1
    s2h(ws2.cell(r2,1), "IO Signal Counts", C_NAVY)
    ws2.merge_cells(f"A{r2}:C{r2}")
    ws2.row_dimensions[r2].height = 20
    r2 += 1

    type_fill = {"DI":C_LBLUE,"DO":C_LGREEN,"AI":C_LYELLOW,"AO":C_LPINK}
    for sig_type, count in counts.items():
        s2h(ws2.cell(r2,1), sig_type, type_fill.get(sig_type, C_WHITE), C_BLACK)
        s2b(ws2.cell(r2,2), count)
        r2 += 1
    s2h(ws2.cell(r2,1), "TOTAL", C_NAVY)
    s2b(ws2.cell(r2,2), sum(counts.values()), bold=True)
    r2 += 2

    if warnings:
        s2h(ws2.cell(r2,1), f"Warnings ({len(warnings)})", C_RED_HDR, C_WHITE)
        ws2.merge_cells(f"A{r2}:C{r2}")
        ws2.row_dimensions[r2].height = 20
        r2 += 1
        for w in warnings:
            c = ws2.cell(r2,1)
            c.value = w
            c.font = _font(C_RED_HDR, 9)
            c.fill = _fill(C_UNASSIGNED)
            c.border = _border()
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws2.merge_cells(f"A{r2}:C{r2}")
            ws2.row_dimensions[r2].height = 30
            r2 += 1
    else:
        s2h(ws2.cell(r2,1), "No warnings — all signals assigned successfully ✓", C_LGREEN, "FF1E4D2B")
        ws2.merge_cells(f"A{r2}:C{r2}")

    wb.save(out_path)
    log(f"✓ File saved: {out_path}")


# ──────────────────────────────────────────────────────────────
# EPLAN EXPORT
# ──────────────────────────────────────────────────────────────
def generate_eplan_excel(final_rows, hw_config, module_db, out_path, log_fn=None):
    def log(msg):
        if log_fn: log_fn(msg)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EPLAN"

    # Row 1
    ws["B1"] = "EPLAN Software & Service"

    # Row 2 — headers
    headers = {
        "B2": "Function definition",
        "C2": "Function group",
        "D2": "Function category",
        "E2": "DT: Identifier",
        "F2": "DT: Counter",
        "G2": "DT: Subcounter",
        "H2": "Plug designation",
        "I2": "Function text en_US",
        "J2": "Symbolic address",
    }
    for cell_ref, val in headers.items():
        ws[cell_ref] = val

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 14
    ws.column_dimensions["I"].width = 50
    ws.column_dimensions["J"].width = 20

    _, _, slot_lookup = get_rack_module_assignments(hw_config, module_db)

    serial = random.randint(100000, 899999)
    current_row = 3

    for r in final_rows:
        rt = r["_row_type"]
        if rt not in ("signal", "spare"):
            continue

        sig_type  = r["sig_type"]
        rack_num  = r["rack"]
        slot_num  = r["slot"]
        ch        = r["ch"]
        rack_name = r.get("rack_name", "")

        mod_entry  = slot_lookup.get((rack_name, slot_num), {})
        mod_idx    = mod_entry.get("mod_idx", 0)
        subcounter = f"{sig_type}{str(mod_idx).zfill(2)}" if mod_idx else ""

        dt_counter = f"{str(rack_num).zfill(2)}.S{str(slot_num).zfill(2)}"

        plug = f"IN {ch}" if sig_type in ("DI", "AI") else f"OUT {ch}"

        if rt == "signal":
            func_text = f"{r['eq_desc']} – {r['sig_desc']}"
            tag_val   = r["tag"]
        else:  # spare
            func_text = "SPARE"
            tag_val   = "SPARE"

        ws.cell(current_row, 1).value  = f"17/{serial:06d}"
        ws.cell(current_row, 2).value  = "2"
        ws.cell(current_row, 3).value  = "1"
        ws.cell(current_row, 4).value  = "300"
        ws.cell(current_row, 5).value  = "R"
        ws.cell(current_row, 6).value  = dt_counter
        ws.cell(current_row, 7).value  = subcounter
        ws.cell(current_row, 8).value  = plug
        ws.cell(current_row, 9).value  = func_text
        ws.cell(current_row, 10).value = tag_val

        serial += 1
        current_row += 1

    wb.save(out_path)
    log(f"✓ EPLAN file saved: {out_path} ({current_row - 3} records)")


# ──────────────────────────────────────────────────────────────
# L5X HARDWARE CONFIG GENERATOR
# ──────────────────────────────────────────────────────────────
def generate_l5x_files(hw_config, module_db, template_dir, output_dir, sw_revision, log_fn=None):
    def log(msg):
        if log_fn: log_fn(msg)

    required = set(row["model"] for row in hw_config)
    missing  = [m for m in required if not os.path.exists(os.path.join(template_dir, f"{m}.L5X"))]
    if missing:
        raise ValueError(
            f"Missing L5X templates for the following modules:\n  " +
            "\n  ".join(missing) +
            f"\n\nPlease add the corresponding .L5X files to:\n  {template_dir}"
        )

    assignments_by_rack, rack_order, _ = get_rack_module_assignments(hw_config, module_db)
    os.makedirs(output_dir, exist_ok=True)

    for rack_name in rack_order:
        rack_mods = assignments_by_rack[rack_name]
        module_blocks = []
        first_mod_name = None

        for entry in rack_mods:
            model    = entry["model"]
            mod_name = entry["name"]
            slot_num = entry["slot"]

            if first_mod_name is None:
                first_mod_name = mod_name

            tpl_path = os.path.join(template_dir, f"{model}.L5X")
            with open(tpl_path, "r", encoding="utf-8-sig") as f:
                content = f.read()

            content = re.sub(r'SoftwareRevision="[^"]*"',
                             f'SoftwareRevision="{sw_revision}"', content)
            content = re.sub(r'TargetName="[^"]*"',
                             f'TargetName="{mod_name}"', content)
            content = re.sub(r'(<Module\s+Use="Target"\s+)Name="[^"]*"',
                             f'\\1Name="{mod_name}"', content)
            content = re.sub(r'ParentModule="[^"]*"',
                             f'ParentModule="{rack_name}"', content)
            content = re.sub(r'(<Port\s[^>]*)Address="[^"]*"',
                             f'\\1Address="{slot_num}"', content)

            m = re.search(r'<Module\s+Use="Target".*?</Module>', content, re.DOTALL)
            if m:
                module_blocks.append(m.group(0))
            else:
                log(f"⚠  Could not extract Module block from template: {model}.L5X")

        if not module_blocks:
            log(f"⚠  No modules for rack {rack_name} — skipped.")
            continue

        now_str = datetime.now().strftime("%a %b %d %H:%M:%S %Y")
        combined = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<RSLogix5000Content SchemaRevision="1.0" SoftwareRevision="{sw_revision}" '
            f'TargetName="{first_mod_name}" TargetType="Module" ContainsContext="true" '
            f'ExportDate="{now_str}" ExportOptions="References NoRawData L5KData DecoratedData '
            f'Context Dependencies ForceProtectedEncoding AllProjDocTrans">\n'
            '<Controller Use="Context" Name="C">\n'
            '<Modules Use="Context">\n'
        )
        for block in module_blocks:
            combined += block + "\n"
        combined += '</Modules>\n</Controller>\n</RSLogix5000Content>'

        safe_name = rack_name.replace(" ", "_")
        out_path  = os.path.join(output_dir, f"{safe_name}.L5X")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(combined)
        log(f"✓ L5X saved: {os.path.basename(out_path)} ({len(module_blocks)} modules)")


# ──────────────────────────────────────────────────────────────
# TAG OBJECTS CSV GENERATOR
# ──────────────────────────────────────────────────────────────
def generate_tag_objects(equipment, out_path, log_fn=None):
    def log(msg):
        if log_fn: log_fn(msg)

    BOOL_TYPES = {"DI", "DO"}

    # Validate and normalise tag names before writing anything
    validated = []
    for eq in equipment:
        raw = eq["tag"].strip()
        if not raw or raw.upper() == "SPARE":
            continue

        # Replace hyphens with underscores (allowed auto-fix)
        tag_name = raw.replace("-", "_")

        # Check for leading digit
        if tag_name and tag_name[0].isdigit():
            raise ValueError(
                f"Invalid tag name: \"{raw}\"\n\n"
                f"Tag starts with a number, which is not allowed in Studio 5000.\n"
                f"Equipment description: {eq.get('desc','')}\n\n"
                f"Please fix this tag in the EQUIPMENT_LIST sheet and regenerate."
            )

        # Check for remaining invalid characters (anything not a-z, A-Z, 0-9, _)
        bad_chars = sorted({c for c in tag_name if not (c.isalnum() or c == "_")})
        if bad_chars:
            raise ValueError(
                f"Invalid tag name: \"{raw}\"\n\n"
                f"Contains unsupported character(s): {' '.join(repr(c) for c in bad_chars)}\n"
                f"Equipment description: {eq.get('desc','')}\n\n"
                f"Only letters, digits, and underscores are allowed.\n"
                f"Please fix this tag in the EQUIPMENT_LIST sheet and regenerate."
            )

        validated.append((tag_name, eq))

    now_str = datetime.now().strftime("%a %b %d %H:%M:%S %Y")

    rows_written = 0
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        f.write('remark,"CSV-Import-Export"\n')
        f.write(f'remark,"Date = {now_str}"\n')
        f.write('remark,"Version = RSLogix 5000 v33.04"\n')
        f.write('remark,"Owner = "\n')
        f.write('remark,"Company = "\n')
        f.write('0.3\n')
        f.write('TYPE,SCOPE,NAME,DESCRIPTION,DATATYPE,SPECIFIER,ATTRIBUTES\n')

        writer = csv.writer(f, quoting=csv.QUOTE_ALL)
        for tag_name, eq in validated:
            eq_type   = eq["type"]
            data_type = "BOOL" if eq_type in BOOL_TYPES else eq_type
            desc      = eq.get("desc", "").strip()

            if data_type == "BOOL":
                attrs = "(RADIX := Decimal, Constant := false, ExternalAccess := Read/Write)"
            else:
                attrs = "(Constant := false, ExternalAccess := Read/Write)"

            writer.writerow(["TAG", "", tag_name, desc, data_type, "", attrs])
            rows_written += 1

    log(f"✓ Tag objects saved: {out_path} ({rows_written} tags)")


# ──────────────────────────────────────────────────────────────
# MIRRORING CODE GENERATOR
# ──────────────────────────────────────────────────────────────
def generate_mirroring_files(final_rows, plc_attributes, local_rack, code_type, output_dir, log_fn=None):
    def log(msg):
        if log_fn: log_fn(msg)

    os.makedirs(output_dir, exist_ok=True)

    TYPE_LABELS = {
        "DI": "Digital Inputs",
        "DO": "Digital Outputs",
        "AI": "Analog Inputs",
        "AO": "Analog Outputs",
    }

    # Group signals by rack
    signals_by_rack = {}
    for r in final_rows:
        if r["_row_type"] == "signal":
            rack_name = r.get("rack_name", "")
            signals_by_rack.setdefault(rack_name, []).append(r)

    if not signals_by_rack:
        raise ValueError("No signals found in IO list. Generate the IO list first.")

    if not plc_attributes:
        raise ValueError("PLC_ATTRIBUTES sheet not found or empty in the input workbook.\n"
                         "Please add the PLC_ATTRIBUTES sheet with Signal Description and PLC Attribute columns.")

    file_count = 0
    for rack_name, signals in signals_by_rack.items():
        adapter = "LOCAL" if rack_name == local_rack else rack_name.replace(" ", "")

        # Collect statements grouped by signal type
        stmts_by_type = {t: [] for t in TYPE_ORDER}

        for sig in signals:
            sig_desc = sig["sig_desc"]
            plc_attr = plc_attributes.get(sig_desc, "")
            if not plc_attr:
                continue

            inverted = plc_attr.startswith("!")
            if inverted:
                plc_attr = plc_attr[1:]

            tag_name = sig["tag"]           # taken as-is from input template
            tag_ref  = f"{tag_name}.{plc_attr}"
            sig_type = sig["sig_type"]
            slot     = sig["slot"]
            ch       = sig["ch"]

            if sig_type == "DI":
                phys = f"{adapter}:{slot}:I.Data.{ch}"
            elif sig_type == "DO":
                phys = f"{adapter}:{slot}:O.Data.{ch}"
            elif sig_type == "AI":
                phys = f"{adapter}:{slot}:I.Ch{ch}Data"
            else:  # AO
                phys = f"{adapter}:{slot}:O.Ch{ch}Data"

            is_input = sig_type in ("DI", "AI")

            if code_type == "ST":
                if is_input:
                    stmt = f"    {tag_ref} := {'NOT ' if inverted else ''}{phys};"
                else:
                    stmt = f"    {phys} := {'NOT ' if inverted else ''}{tag_ref};"
            else:  # Ladder Logic
                if sig_type == "DI":
                    contact = "XIO" if inverted else "XIC"
                    stmt = f"{contact}({phys})OTE({tag_ref});"
                elif sig_type == "DO":
                    contact = "XIO" if inverted else "XIC"
                    stmt = f"{contact}({tag_ref})OTE({phys});"
                elif sig_type == "AI":
                    src = f"NOT {phys}" if inverted else phys
                    stmt = f"COP({src},{tag_ref},1);"
                else:  # AO
                    src = f"NOT {tag_ref}" if inverted else tag_ref
                    stmt = f"COP({src},{phys},1);"

            stmts_by_type[sig_type].append(stmt)

        # Build output with regions (ST) or comments (LL) per type
        output_lines = []
        for sig_type in TYPE_ORDER:
            type_stmts = stmts_by_type[sig_type]
            if not type_stmts:
                continue
            label = TYPE_LABELS[sig_type]
            if code_type == "ST":
                output_lines.append(f"{{ region {label} }}")
                output_lines.extend(type_stmts)
                output_lines.append("{ end_region }")
                output_lines.append("")
            else:
                output_lines.append(f"(* === {label} === *)")
                output_lines.extend(type_stmts)
                output_lines.append("")

        # Remove trailing blank line
        while output_lines and output_lines[-1] == "":
            output_lines.pop()

        if not output_lines:
            log(f"⚠  No mapped signals for rack {rack_name} — file skipped.")
            continue

        safe_name = rack_name.replace(" ", "_")
        ext       = "st" if code_type == "ST" else "txt"
        out_path  = os.path.join(output_dir, f"{safe_name}_Mirror.{ext}")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("\n".join(output_lines))
        log(f"✓ Mirror saved: {os.path.basename(out_path)} ({sum(len(v) for v in stmts_by_type.values())} statements)")
        file_count += 1

    if file_count == 0:
        raise ValueError("No mirroring files generated. Check that PLC_ATTRIBUTES contains "
                         "Signal Description values that match signals in your equipment list.")


# ──────────────────────────────────────────────────────────────
# GUI
# ──────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Engineering Automation Suite — Allen-Bradley / Rockwell")
        self.resizable(True, True)
        self.minsize(820, 700)
        self.input_path    = None
        self.rev_file_path = None
        self.eplan_out_path = None
        self.tags_out_path  = None
        self.rack_local_var = tk.StringVar(value="NONE")
        self._racks_radio_frame = None
        self._last_final_rows = None
        self._last_hw_config  = None
        self._last_module_db  = None
        self._last_equipment  = None
        self._last_plc_attrs  = None
        # Open-location buttons (enabled after generation)
        self.btn_open_eplan  = None
        self.btn_open_l5x    = None
        self.btn_open_tags   = None
        self.btn_open_mirror = None
        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = 920, 780
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── Build UI ────────────────────────────────────────────
    def _build_ui(self):
        FH  = ("Arial", 11, "bold")
        FB  = ("Arial", 10)
        FSM = ("Arial", 9)
        BG  = "#f5f5f5"
        BN  = "#c0c0c0"
        FN  = "#000000"
        self.configure(bg=BG)

        # Title bar
        tf = tk.Frame(self, bg=BN, height=52)
        tf.pack(fill="x")
        tf.pack_propagate(False)
        tk.Label(tf, text="  Engineering Automation Suite", font=("Arial",14,"bold"),
                 bg=BN, fg=FN, anchor="w").pack(side="left", padx=10, pady=10)
        tk.Label(tf, text="Allen-Bradley / Rockwell Automation",
                 font=("Arial",9), bg=BN, fg="#444").pack(side="left", padx=2)
        tk.Button(tf, text="ℹ  About", font=("Arial",9), bg=BN, fg="#222",
                  relief="flat", padx=10, pady=4, cursor="hand2",
                  command=self._show_about).pack(side="right", padx=10, pady=10)

        # Common: input file
        cf = tk.Frame(self, bg=BG, padx=16, pady=8)
        cf.pack(fill="x")
        self._sec(cf, "Input Workbook", FH)
        fr = tk.Frame(cf, bg=BG)
        fr.pack(fill="x", pady=(0,4))
        self.lbl_file = tk.Label(fr, text="No file selected", font=FSM,
                                 bg="#e8eef4", fg="#444", relief="flat", anchor="w", padx=8, pady=5)
        self.lbl_file.pack(side="left", fill="x", expand=True)
        tk.Button(fr, text="Browse…", font=FB, bg=BN, fg=FN, relief="flat",
                  padx=14, pady=4, cursor="hand2", command=self._browse_input).pack(side="left", padx=(6,0))
        tk.Button(fr, text="Create template", font=FSM, bg="#f0c040", fg="#1a1a1a",
                  relief="flat", padx=10, pady=4, cursor="hand2",
                  command=self._create_template).pack(side="left", padx=(6,0))

        # Main notebook
        style = ttk.Style()
        style.configure("TNotebook", background=BG)
        style.configure("TNotebook.Tab", font=FB, padding=[12,4])

        mnb = ttk.Notebook(self)
        mnb.pack(fill="both", expand=True, padx=10, pady=2)

        dt = tk.Frame(mnb, bg=BG)
        mnb.add(dt, text="  Design Team  ")
        self._build_design_tab(dt, FH, FB, FSM, BG, BN, FN)

        at = tk.Frame(mnb, bg=BG)
        mnb.add(at, text="  Automation Team  ")
        self._build_auto_tab(at, FH, FB, FSM, BG, BN, FN)

        # Progress + log
        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=10, pady=(2,0))

        lf = tk.Frame(self, bg=BG, padx=10, pady=4)
        lf.pack(fill="both", expand=False)
        tk.Label(lf, text="Log", font=FH, bg=BG, fg="#333", anchor="w").pack(fill="x")
        self.log_box = scrolledtext.ScrolledText(lf, height=7, font=("Courier New",9),
                                                  bg="#1a1a2e", fg="#d0e8ff", relief="flat",
                                                  padx=8, pady=6, insertbackground="white")
        self.log_box.pack(fill="both", expand=True)
        self.log_box.configure(state="disabled")

        self.status_var = tk.StringVar(value="Ready")
        tk.Label(self, textvariable=self.status_var, font=FSM,
                 bg=BN, fg=FN, anchor="w", padx=10).pack(side="bottom", fill="x")

    def _sec(self, parent, text, font):
        tk.Frame(parent, bg="#c8c8c8", height=1).pack(fill="x", pady=(4,2))
        tk.Label(parent, text=text, font=font, bg="#f5f5f5", fg="#222", anchor="w").pack(fill="x")

    def _show_about(self):
        win = tk.Toplevel(self)
        win.title("About")
        win.resizable(False, False)
        win.configure(bg="#f5f5f5")
        win.grab_set()

        # Header bar
        hf = tk.Frame(win, bg="#c0c0c0", height=60)
        hf.pack(fill="x")
        hf.pack_propagate(False)
        tk.Label(hf, text="Engineering Automation Suite", font=("Arial",14,"bold"),
                 bg="#c0c0c0", fg="#000").pack(side="left", padx=16, pady=12)

        # Details
        df = tk.Frame(win, bg="#f5f5f5", padx=24, pady=16)
        df.pack(fill="both")

        fields = [
            ("Version",  APP_VERSION),
            ("Released", APP_DATE),
            ("Author",   APP_AUTHOR),
            ("Company",  APP_COMPANY),
            ("Platform", "Allen-Bradley / Rockwell Automation"),
        ]
        for label, value in fields:
            row = tk.Frame(df, bg="#f5f5f5")
            row.pack(fill="x", pady=2)
            tk.Label(row, text=f"{label}:", font=("Arial",9,"bold"), bg="#f5f5f5",
                     fg="#555", width=10, anchor="w").pack(side="left")
            tk.Label(row, text=value, font=("Arial",9), bg="#f5f5f5",
                     fg="#111", anchor="w").pack(side="left")

        tk.Frame(df, bg="#c8c8c8", height=1).pack(fill="x", pady=(12,8))
        tk.Label(df, text=APP_DESC, font=("Arial",9), bg="#f5f5f5", fg="#444",
                 justify="left", anchor="w").pack(fill="x")

        tk.Frame(df, bg="#c8c8c8", height=1).pack(fill="x", pady=(12,0))

        tk.Button(win, text="Close", font=("Arial",10), bg="#c0c0c0", fg="#000",
                  relief="flat", padx=20, pady=6, cursor="hand2",
                  command=win.destroy).pack(pady=12)

        # Center over parent
        win.update_idletasks()
        pw = self.winfo_x() + self.winfo_width()  // 2
        ph = self.winfo_y() + self.winfo_height() // 2
        ww, wh = win.winfo_width(), win.winfo_height()
        win.geometry(f"+{pw - ww//2}+{ph - wh//2}")

    # ── Design Team tab ─────────────────────────────────────
    def _build_design_tab(self, parent, FH, FB, FSM, BG, BN, FN):
        dnb = ttk.Notebook(parent)
        dnb.pack(fill="both", expand=True, padx=6, pady=6)

        # IO List sub-tab
        ilf = tk.Frame(dnb, bg=BG, padx=14, pady=10)
        dnb.add(ilf, text="  IO List  ")
        self._build_iolist_sub(ilf, FH, FB, FSM, BG, BN, FN)

        # EPLAN Export sub-tab
        ef = tk.Frame(dnb, bg=BG, padx=14, pady=10)
        dnb.add(ef, text="  EPLAN Export  ")
        self._build_eplan_sub(ef, FH, FB, FSM, BG, BN, FN)

    # ── Automation Team tab ──────────────────────────────────
    def _build_auto_tab(self, parent, FH, FB, FSM, BG, BN, FN):
        anb = ttk.Notebook(parent)
        anb.pack(fill="both", expand=True, padx=6, pady=6)

        # L5X sub-tab
        lf = tk.Frame(anb, bg=BG, padx=14, pady=10)
        anb.add(lf, text="  L5X Hardware  ")
        self._build_l5x_sub(lf, FH, FB, FSM, BG, BN, FN)

        # Tags sub-tab
        tf = tk.Frame(anb, bg=BG, padx=14, pady=10)
        anb.add(tf, text="  Tag Objects  ")
        self._build_tags_sub(tf, FH, FB, FSM, BG, BN, FN)

        # Mirroring sub-tab
        mf = tk.Frame(anb, bg=BG, padx=14, pady=10)
        anb.add(mf, text="  Mirroring  ")
        self._build_mirror_sub(mf, FH, FB, FSM, BG, BN, FN)

    def _build_iolist_sub(self, p, FH, FB, FSM, BG, BN, FN):
        self._sec(p, "Revision Mode (optional)", FH)
        rv = tk.Frame(p, bg=BG)
        rv.pack(fill="x", pady=(0,4))
        self.revision_var = tk.BooleanVar(value=False)
        tk.Checkbutton(rv, text="Enable revision mode — attach a manufacturing IO list",
                       variable=self.revision_var, command=self._toggle_revision,
                       font=FB, bg=BG, activebackground=BG).pack(side="left")
        ra = tk.Frame(p, bg=BG)
        ra.pack(fill="x", pady=(0,10))
        self.lbl_rev_file = tk.Label(ra, text="No manufacturing IO list attached",
                                     font=FSM, bg="#e8eef4", fg="#888", relief="flat",
                                     anchor="w", padx=8, pady=5)
        self.lbl_rev_file.pack(side="left", fill="x", expand=True)
        self.btn_rev_attach = tk.Button(ra, text="Attach…", font=FB, bg="#888", fg="white",
                                        relief="flat", padx=14, pady=4, cursor="hand2",
                                        command=self._browse_rev_file, state="disabled")
        self.btn_rev_attach.pack(side="left", padx=(6,0))

        self._sec(p, "Output Folder", FH)
        orf = tk.Frame(p, bg=BG)
        orf.pack(fill="x", pady=(0,10))
        self.out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "IO List")
        self.lbl_out = tk.Label(orf, text=self.out_dir, font=FSM,
                                bg="#e8eef4", fg="#444", relief="flat", anchor="w", padx=8, pady=5)
        self.lbl_out.pack(side="left", fill="x", expand=True)
        tk.Button(orf, text="Change…", font=FB, bg=BN, fg=FN, relief="flat",
                  padx=14, pady=4, cursor="hand2", command=self._browse_out).pack(side="left", padx=(6,0))

        bf = tk.Frame(p, bg=BG)
        bf.pack(fill="x", pady=(6,10))
        self.btn_gen = tk.Button(bf, text="▶  Generate IO List", font=("Arial",11,"bold"),
                                 bg="#1E7C3A", fg="white", relief="flat",
                                 padx=22, pady=8, cursor="hand2", command=self._run_generate)
        self.btn_gen.pack(side="left")
        self.btn_open = tk.Button(bf, text="Open output folder", font=FB, bg=BN, fg=FN,
                                  relief="flat", padx=14, pady=8, cursor="hand2",
                                  command=self._open_out_folder, state="disabled")
        self.btn_open.pack(side="left", padx=(10,0))

    def _build_eplan_sub(self, p, FH, FB, FSM, BG, BN, FN):
        self._sec(p, "EPLAN Import-Ready Excel", FH)
        tk.Label(p, text="Generates one row per IO signal (and SPARE channels), ready for EPLAN import.",
                 font=FSM, bg=BG, fg="#555").pack(anchor="w", pady=(0,8))

        er = tk.Frame(p, bg=BG)
        er.pack(fill="x", pady=(0,8))
        tk.Label(er, text="Output file:", font=FB, bg=BG).pack(side="left")
        self.lbl_eplan = tk.Label(er, text="Not set", font=FSM,
                                  bg="#e8eef4", fg="#444", relief="flat", anchor="w", padx=8, pady=5)
        self.lbl_eplan.pack(side="left", fill="x", expand=True, padx=(6,0))
        tk.Button(er, text="Browse…", font=FB, bg=BN, fg=FN, relief="flat",
                  padx=14, pady=4, cursor="hand2",
                  command=self._browse_eplan_out).pack(side="left", padx=(6,0))

        bf = tk.Frame(p, bg=BG)
        bf.pack(fill="x", pady=(8,0))
        tk.Button(bf, text="▶  Generate EPLAN File", font=("Arial",11,"bold"),
                  bg="#1E7C3A", fg="white", relief="flat", padx=22, pady=8, cursor="hand2",
                  command=self._run_eplan).pack(side="left")
        self.btn_open_eplan = tk.Button(bf, text="Open file location", font=FB, bg=BN, fg=FN,
                                        relief="flat", padx=14, pady=8, cursor="hand2",
                                        command=lambda: self._open_location(self.eplan_out_path),
                                        state="disabled")
        self.btn_open_eplan.pack(side="left", padx=(10,0))

    def _build_l5x_sub(self, p, FH, FB, FSM, BG, BN, FN):
        self._sec(p, "Studio 5000 Hardware Configuration", FH)
        tk.Label(p, text="Generates one L5X file per rack for import into Studio 5000 Logix Designer.",
                 font=FSM, bg=BG, fg="#555").pack(anchor="w", pady=(0,8))

        sr = tk.Frame(p, bg=BG)
        sr.pack(fill="x", pady=(0,8))
        tk.Label(sr, text="Software Revision:", font=FB, bg=BG).pack(side="left")
        self.l5x_sw_rev = tk.StringVar(value="33.04")
        tk.Entry(sr, textvariable=self.l5x_sw_rev, font=FB, width=10,
                 relief="flat", bg="#e8eef4").pack(side="left", padx=(8,0), ipady=3)

        tpl_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ModuleL5K")
        tk.Label(p, text=f"Templates folder:  {tpl_dir}", font=FSM, bg=BG, fg="#555").pack(anchor="w", pady=(0,8))

        self._sec(p, "Output Folder", FH)
        of = tk.Frame(p, bg=BG)
        of.pack(fill="x", pady=(0,10))
        self.l5x_out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "L5X Output")
        self.lbl_l5x_out = tk.Label(of, text=self.l5x_out_dir, font=FSM,
                                     bg="#e8eef4", fg="#444", relief="flat", anchor="w", padx=8, pady=5)
        self.lbl_l5x_out.pack(side="left", fill="x", expand=True)
        tk.Button(of, text="Change…", font=FB, bg=BN, fg=FN, relief="flat",
                  padx=14, pady=4, cursor="hand2", command=self._browse_l5x_out).pack(side="left", padx=(6,0))

        bf = tk.Frame(p, bg=BG)
        bf.pack(fill="x", pady=(8,0))
        tk.Button(bf, text="▶  Generate L5X Files", font=("Arial",11,"bold"),
                  bg="#1E7C3A", fg="white", relief="flat", padx=22, pady=8, cursor="hand2",
                  command=self._run_l5x).pack(side="left")
        self.btn_open_l5x = tk.Button(bf, text="Open output folder", font=FB, bg=BN, fg=FN,
                                      relief="flat", padx=14, pady=8, cursor="hand2",
                                      command=lambda: self._open_location(self.l5x_out_dir),
                                      state="disabled")
        self.btn_open_l5x.pack(side="left", padx=(10,0))

    def _build_tags_sub(self, p, FH, FB, FSM, BG, BN, FN):
        self._sec(p, "Studio 5000 Tag Objects", FH)
        tk.Label(p, text="Generates a CSV file with equipment tag definitions for Studio 5000 import.",
                 font=FSM, bg=BG, fg="#555").pack(anchor="w", pady=(0,8))

        self._sec(p, "Output File", FH)
        of = tk.Frame(p, bg=BG)
        of.pack(fill="x", pady=(0,10))
        self.lbl_tags = tk.Label(of, text="Not set", font=FSM,
                                  bg="#e8eef4", fg="#444", relief="flat", anchor="w", padx=8, pady=5)
        self.lbl_tags.pack(side="left", fill="x", expand=True)
        tk.Button(of, text="Browse…", font=FB, bg=BN, fg=FN, relief="flat",
                  padx=14, pady=4, cursor="hand2", command=self._browse_tags_out).pack(side="left", padx=(6,0))

        bf = tk.Frame(p, bg=BG)
        bf.pack(fill="x", pady=(8,0))
        tk.Button(bf, text="▶  Generate Tag Objects", font=("Arial",11,"bold"),
                  bg="#1E7C3A", fg="white", relief="flat", padx=22, pady=8, cursor="hand2",
                  command=self._run_tags).pack(side="left")
        self.btn_open_tags = tk.Button(bf, text="Open file location", font=FB, bg=BN, fg=FN,
                                       relief="flat", padx=14, pady=8, cursor="hand2",
                                       command=lambda: self._open_location(self.tags_out_path),
                                       state="disabled")
        self.btn_open_tags.pack(side="left", padx=(10,0))

    def _build_mirror_sub(self, p, FH, FB, FSM, BG, BN, FN):
        self._sec(p, "PLC Mirroring Code", FH)
        tk.Label(p, text="Generates mirroring statements linking physical IO addresses to PLC tag objects.",
                 font=FSM, bg=BG, fg="#555").pack(anchor="w", pady=(0,8))

        # Code type
        cr = tk.Frame(p, bg=BG)
        cr.pack(fill="x", pady=(0,8))
        tk.Label(cr, text="Code type:", font=FB, bg=BG).pack(side="left")
        self.mirror_code_type = tk.StringVar(value="ST")
        tk.Radiobutton(cr, text="Structured Text", variable=self.mirror_code_type,
                       value="ST", font=FB, bg=BG, activebackground=BG).pack(side="left", padx=(8,0))
        tk.Radiobutton(cr, text="Ladder Logic", variable=self.mirror_code_type,
                       value="LL", font=FB, bg=BG, activebackground=BG).pack(side="left", padx=(8,0))

        # LOCAL rack (radio — only one)
        self._sec(p, "LOCAL Rack", FH)
        tk.Label(p, text="Select the rack that is in the same chassis as the controller (LOCAL):",
                 font=FSM, bg=BG, fg="#555").pack(anchor="w")
        rack_outer = tk.Frame(p, bg=BG)
        rack_outer.pack(fill="x", pady=(4,8))
        self._racks_radio_frame = tk.Frame(rack_outer, bg="#e8eef4", relief="flat", padx=8, pady=6)
        self._racks_radio_frame.pack(side="left")
        tk.Label(self._racks_radio_frame, text="Browse an input file to see racks",
                 font=FSM, bg="#e8eef4", fg="#888").pack()

        # Output folder
        self._sec(p, "Output Folder", FH)
        of = tk.Frame(p, bg=BG)
        of.pack(fill="x", pady=(0,10))
        self.mirror_out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Mirroring")
        self.lbl_mirror_out = tk.Label(of, text=self.mirror_out_dir, font=FSM,
                                        bg="#e8eef4", fg="#444", relief="flat", anchor="w", padx=8, pady=5)
        self.lbl_mirror_out.pack(side="left", fill="x", expand=True)
        tk.Button(of, text="Change…", font=FB, bg=BN, fg=FN, relief="flat",
                  padx=14, pady=4, cursor="hand2", command=self._browse_mirror_out).pack(side="left", padx=(6,0))

        bf = tk.Frame(p, bg=BG)
        bf.pack(fill="x", pady=(8,0))
        tk.Button(bf, text="▶  Generate Mirroring Files", font=("Arial",11,"bold"),
                  bg="#1E7C3A", fg="white", relief="flat", padx=22, pady=8, cursor="hand2",
                  command=self._run_mirroring).pack(side="left")
        self.btn_open_mirror = tk.Button(bf, text="Open output folder", font=FB, bg=BN, fg=FN,
                                         relief="flat", padx=14, pady=8, cursor="hand2",
                                         command=lambda: self._open_location(self.mirror_out_dir),
                                         state="disabled")
        self.btn_open_mirror.pack(side="left", padx=(10,0))

    # ── Helpers ─────────────────────────────────────────────
    def _toggle_revision(self):
        if self.revision_var.get():
            self.btn_rev_attach.configure(state="normal", bg="#595959")
            self.lbl_rev_file.configure(fg="#444")
        else:
            self.btn_rev_attach.configure(state="disabled", bg="#888")
            self.lbl_rev_file.configure(text="No manufacturing IO list attached", fg="#888")
            self.rev_file_path = None

    def _log(self, msg):
        self.log_box.configure(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{ts}]  {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _set_status(self, msg):
        self.status_var.set(msg)

    def _update_rack_checkboxes(self):
        if not self.input_path or not self._racks_radio_frame:
            return
        racks = _quick_read_racks(self.input_path)
        for w in self._racks_radio_frame.winfo_children():
            w.destroy()
        self.rack_local_var.set("NONE")
        if not racks:
            tk.Label(self._racks_radio_frame, text="No racks found",
                     font=("Arial",9), bg="#e8eef4", fg="#888").pack()
            return
        # "None" option
        tk.Radiobutton(self._racks_radio_frame, text="None", variable=self.rack_local_var,
                       value="NONE", font=("Arial",10), bg="#e8eef4",
                       activebackground="#e8eef4").pack(anchor="w")
        for rack in racks:
            tk.Radiobutton(self._racks_radio_frame, text=rack, variable=self.rack_local_var,
                           value=rack, font=("Arial",10), bg="#e8eef4",
                           activebackground="#e8eef4").pack(anchor="w")

    # ── Browse handlers ──────────────────────────────────────
    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select Input Workbook",
            filetypes=[("Excel files","*.xlsx *.xlsm"),("All files","*.*")])
        if path:
            self.input_path = path
            self.lbl_file.configure(text=f"  {os.path.basename(path)}  ({path})")
            self._update_rack_checkboxes()

    def _browse_rev_file(self):
        path = filedialog.askopenfilename(
            title="Select Manufacturing IO List",
            filetypes=[("Excel files","*.xlsx *.xlsm"),("All files","*.*")])
        if path:
            self.rev_file_path = path
            self.lbl_rev_file.configure(text=f"  {os.path.basename(path)}  ({path})")

    def _browse_out(self):
        d = filedialog.askdirectory(title="Select Output Folder", initialdir=self.out_dir)
        if d:
            self.out_dir = d
            self.lbl_out.configure(text=d)

    def _browse_eplan_out(self):
        path = filedialog.asksaveasfilename(
            title="Save EPLAN File As", defaultextension=".xlsx",
            initialfile="EPLAN_Export.xlsx",
            filetypes=[("Excel files","*.xlsx")])
        if path:
            self.eplan_out_path = path
            self.lbl_eplan.configure(text=f"  {os.path.basename(path)}  ({path})")

    def _browse_l5x_out(self):
        d = filedialog.askdirectory(title="Select L5X Output Folder", initialdir=self.l5x_out_dir)
        if d:
            self.l5x_out_dir = d
            self.lbl_l5x_out.configure(text=d)

    def _browse_tags_out(self):
        path = filedialog.asksaveasfilename(
            title="Save Tag Objects File As", defaultextension=".csv",
            initialfile="Tag_Objects.csv",
            filetypes=[("CSV files","*.csv"),("All files","*.*")])
        if path:
            self.tags_out_path = path
            self.lbl_tags.configure(text=f"  {os.path.basename(path)}  ({path})")

    def _browse_mirror_out(self):
        d = filedialog.askdirectory(title="Select Mirroring Output Folder", initialdir=self.mirror_out_dir)
        if d:
            self.mirror_out_dir = d
            self.lbl_mirror_out.configure(text=d)

    def _open_location(self, path):
        import subprocess, platform
        if not path:
            return
        try:
            if platform.system() == "Windows":
                if os.path.isdir(path):
                    os.startfile(path)
                else:
                    # Open parent folder and highlight the file
                    subprocess.Popen(f'explorer /select,"{path}"')
            elif platform.system() == "Darwin":
                target = path if os.path.isdir(path) else os.path.dirname(path)
                subprocess.Popen(["open", target])
            else:
                target = path if os.path.isdir(path) else os.path.dirname(path)
                subprocess.Popen(["xdg-open", target])
        except:
            pass

    def _open_out_folder(self):
        self._open_location(self.out_dir)

    def _create_template(self):
        from create_template import main as make_tpl
        save_path = filedialog.asksaveasfilename(
            title="Save Input Template As", defaultextension=".xlsx",
            initialfile="IO_List_Input_Template.xlsx",
            filetypes=[("Excel files","*.xlsx")])
        if not save_path:
            return
        try:
            make_tpl(save_path)
            messagebox.showinfo("Template created",
                                f"Template saved to:\n{save_path}\n\nFill in all sheets then run Generate.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _require_input(self):
        if not self.input_path:
            messagebox.showwarning("No input file", "Please select the input workbook first.")
            return False
        return True

    def _start_task(self):
        self.progress.start(12)

    def _stop_task(self, status="Done"):
        self.progress.stop()
        self._set_status(status)

    # ── IO List generation ───────────────────────────────────
    def _run_generate(self):
        if not self._require_input():
            return
        revision_mode = self.revision_var.get()
        if revision_mode and not self.rev_file_path:
            messagebox.showwarning("No manufacturing IO list",
                                   "Revision mode is enabled.\nPlease attach the manufacturing IO list.")
            return
        os.makedirs(self.out_dir, exist_ok=True)
        self.btn_gen.configure(state="disabled")
        self._start_task()
        self._set_status("Generating IO list…")

        def task():
            try:
                self._log(f"Reading: {os.path.basename(self.input_path)}")
                proj, module_db, hw_config, eq_types, matrix, sig_order, equipment, plc_attrs = \
                    read_workbook(self.input_path)
                self._log(f"Project: {proj.get('Project Name','?')} | Equipment: {len(equipment)}")

                if revision_mode:
                    self._log(f"Revision — parsing: {os.path.basename(self.rev_file_path)}")
                    mfg = parse_manufacturing_io_list(self.rev_file_path)
                    self._log(f"Manufacturing IO list: {len(mfg)} channels")
                    final_rows, counts, warnings = build_io_list_revision(
                        proj, module_db, mfg, eq_types, matrix, sig_order, equipment, self._log)
                else:
                    self._log(f"Racks: {len(set(r['rack'] for r in hw_config))} | Modules: {len(hw_config)} | Signals: {len(sig_order)}")
                    final_rows, counts, warnings = build_io_list(
                        proj, module_db, hw_config, eq_types, matrix, sig_order, equipment, self._log)

                self._last_final_rows = final_rows
                self._last_hw_config  = hw_config
                self._last_module_db  = module_db
                self._last_equipment  = equipment
                self._last_plc_attrs  = plc_attrs

                now  = datetime.now()
                name = proj.get("Project Name","Project").replace(" ","_")
                sfx  = "_Revision" if revision_mode else ""
                fname = f"{name}_{now.strftime('%Y%m%d_%H%M%S')}{sfx}_IO_List.xlsx"
                out_path = os.path.join(self.out_dir, fname)
                write_output(proj, module_db, final_rows, counts, warnings, out_path, self._log)
                self.after(0, lambda: self._on_done_io(out_path, warnings))
            except Exception as e:
                import traceback
                self._log(f"✗ ERROR: {e}")
                self._log(traceback.format_exc())
                self.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=task, daemon=True).start()

    def _on_done_io(self, out_path, warnings):
        self._stop_task(f"Done — {os.path.basename(out_path)}")
        self.btn_gen.configure(state="normal")
        self.btn_open.configure(state="normal")
        if warnings:
            messagebox.showwarning("Completed with warnings",
                                   f"IO List generated with {len(warnings)} warning(s).\n"
                                   f"Check 'Generation Summary' sheet.\n\nFile: {os.path.basename(out_path)}")
        else:
            messagebox.showinfo("Success", f"IO List generated!\n\nFile: {os.path.basename(out_path)}")

    # ── EPLAN generation ─────────────────────────────────────
    def _run_eplan(self):
        if not self._require_input():
            return
        if not self.eplan_out_path:
            messagebox.showwarning("No output file", "Please select an output file path first.")
            return
        self._start_task()
        self._set_status("Generating EPLAN file…")

        def task():
            try:
                self._log(f"Reading: {os.path.basename(self.input_path)}")
                proj, module_db, hw_config, eq_types, matrix, sig_order, equipment, plc_attrs = \
                    read_workbook(self.input_path)
                final_rows, _, _ = build_io_list(
                    proj, module_db, hw_config, eq_types, matrix, sig_order, equipment, self._log)
                generate_eplan_excel(final_rows, hw_config, module_db, self.eplan_out_path, self._log)
                self.after(0, lambda: self._on_done_simple("EPLAN", self.eplan_out_path, self.btn_open_eplan))
            except Exception as e:
                import traceback
                self._log(f"✗ ERROR: {e}")
                self._log(traceback.format_exc())
                self.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=task, daemon=True).start()

    # ── L5X generation ───────────────────────────────────────
    def _run_l5x(self):
        if not self._require_input():
            return
        tpl_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ModuleL5K")
        self._start_task()
        self._set_status("Generating L5X files…")

        def task():
            try:
                self._log(f"Reading: {os.path.basename(self.input_path)}")
                _, module_db, hw_config, _, _, _, _, _ = read_workbook(self.input_path)
                generate_l5x_files(hw_config, module_db, tpl_dir,
                                   self.l5x_out_dir, self.l5x_sw_rev.get(), self._log)
                self.after(0, lambda: self._on_done_folder("L5X", self.l5x_out_dir, self.btn_open_l5x))
            except Exception as e:
                import traceback
                self._log(f"✗ ERROR: {e}")
                self._log(traceback.format_exc())
                self.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=task, daemon=True).start()

    # ── Tag objects generation ───────────────────────────────
    def _run_tags(self):
        if not self._require_input():
            return
        if not self.tags_out_path:
            messagebox.showwarning("No output file", "Please select an output file path first.")
            return
        self._start_task()
        self._set_status("Generating tag objects…")

        def task():
            try:
                self._log(f"Reading: {os.path.basename(self.input_path)}")
                _, _, _, _, _, _, equipment, _ = read_workbook(self.input_path)
                generate_tag_objects(equipment, self.tags_out_path, self._log)
                self.after(0, lambda: self._on_done_simple("Tag Objects", self.tags_out_path, self.btn_open_tags))
            except Exception as e:
                import traceback
                self._log(f"✗ ERROR: {e}")
                self._log(traceback.format_exc())
                self.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=task, daemon=True).start()

    # ── Mirroring generation ─────────────────────────────────
    def _run_mirroring(self):
        if not self._require_input():
            return
        local_rack = self.rack_local_var.get()
        if local_rack == "NONE":
            local_rack = ""
        code_type  = self.mirror_code_type.get()
        self._start_task()
        self._set_status("Generating mirroring files…")

        def task():
            try:
                self._log(f"Reading: {os.path.basename(self.input_path)}")
                proj, module_db, hw_config, eq_types, matrix, sig_order, equipment, plc_attrs = \
                    read_workbook(self.input_path)
                final_rows, _, _ = build_io_list(
                    proj, module_db, hw_config, eq_types, matrix, sig_order, equipment, self._log)
                generate_mirroring_files(final_rows, plc_attrs, local_rack,
                                         code_type, self.mirror_out_dir, self._log)
                self.after(0, lambda: self._on_done_folder("Mirroring", self.mirror_out_dir, self.btn_open_mirror))
            except Exception as e:
                import traceback
                self._log(f"✗ ERROR: {e}")
                self._log(traceback.format_exc())
                self.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=task, daemon=True).start()

    # ── Done/Error callbacks ─────────────────────────────────
    def _on_done_simple(self, label, path, open_btn=None):
        self._stop_task(f"{label} done — {os.path.basename(path)}")
        if open_btn:
            open_btn.configure(state="normal")
        messagebox.showinfo("Success", f"{label} file generated!\n\nFile: {path}")

    def _on_done_folder(self, label, folder, open_btn=None):
        self._stop_task(f"{label} done — {folder}")
        if open_btn:
            open_btn.configure(state="normal")
        messagebox.showinfo("Success", f"{label} files generated!\n\nFolder: {folder}")

    def _on_error(self, msg):
        self._stop_task("Error — see log")
        messagebox.showerror("Generation failed", f"An error occurred:\n\n{msg}\n\nSee log for details.")


if __name__ == "__main__":
    app = App()
    app.mainloop()
