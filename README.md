# IO List Generator
### Allen-Bradley / Rockwell Automation — Python Tool

---

## Requirements

- Python 3.8 or higher (https://www.python.org/downloads/)
- `openpyxl` library

### Install openpyxl (one time only)
Open Command Prompt and run:
```
pip install openpyxl
```

---

## How to Run

Double-click `io_list_generator.py` or run:
```
python io_list_generator.py
```

---

## Input Workbook Structure

The tool reads a single Excel workbook with **5 sheets**:

### Sheet 1 — `PROJECT_INFO`
| Column A (Field)       | Column B (Value)           |
|------------------------|----------------------------|
| Project Name           | My Project                 |
| Project Number         | PRJ-001                    |
| Controller Tag Prefix  | PLC-1                      |
| Controller Model       | 5069-L310ER                |
| IO Address Prefix      | DI_                        |
| Engineer Name          | Eng. Anas                  |
| Client Name            | (optional)                 |
| Site Name              | (optional)                 |
| Revision               | A                          |

---

### Sheet 2 — `MODULE_DB`
Define all Allen-Bradley I/O modules available.

| Model No.   | Description                          | Type | Channel Qty |
|-------------|--------------------------------------|------|-------------|
| 1734-IB8    | POINT I/O 8-Ch Digital Input 24VDC   | DI   | 8           |
| 1734-OB8    | POINT I/O 8-Ch Digital Output 24VDC  | DO   | 8           |
| 1734-IE8C   | POINT I/O 8-Ch Analog Input 4-20mA   | AI   | 8           |
| 1734-OE4C   | POINT I/O 4-Ch Analog Output 4-20mA  | AO   | 4           |

**Type must be exactly:** `DI`, `DO`, `AI`, or `AO`

---

### Sheet 3 — `HARDWARE_CONFIG`
Define the rack layout for your project.

| Rack Name | Module Model No. | Modules Qty |
|-----------|-----------------|-------------|
| RACK 01   | 1734-IB8        | 2           |
| RACK 01   | 1734-OB8        | 1           |
| RACK 01   | 1734-IE8C       | 1           |
| RACK 01   | 1734-OE4C       | 1           |
| RACK 02   | 1734-IB8        | 2           |

**Rules:**
- Rack Name must be consistent across rows (e.g., `RACK 01`)
- Module Model No. must match exactly what is in `MODULE_DB`
- Signals are arranged **DI → DO → AI → AO** per rack

---

### Sheet 4 — `SIGNAL_MATRIX`
Defines which hardware signals each equipment type uses.

| Category | Signal Type | Signal Row Name | VSD | FSD           | HVALVE | MVALVE        | AI  | DI    |
|----------|-------------|-----------------|-----|---------------|--------|---------------|-----|-------|
| HW       | DI          | DI Signal1      |     | Running       | Opened | Opened        |     | Input |
| HW       | DI          | DI Signal2      |     | Fault         | Closed | Closed        |     |       |
| HW       | DI          | DI Signal3      |     | Remote        |        | Remote        |     |       |
| HW       | DO          | DO Signal1      |     | Start Command |        | Open Command  |     |       |
| HW       | DO          | DO Signal2      |     |               |        | Close Command |     |       |
| HW       | AI          | AI Signal1      |     |               |        |               | Raw |       |

**Rules:**
- Signal Row Name must **start with** `DI`, `DO`, `AI`, or `AO`
- Cell value = the signal description for that equipment type
- Leave blank if the signal does not apply to that equipment type
- Add new columns for new equipment types
- Add new rows for additional signals

---

### Sheet 5 — `EQUIPMENT_LIST`
List of all equipment to be included in the IO list.

| Tag     | Description               | Type   | System (optional) |
|---------|---------------------------|--------|-------------------|
| VFD-001 | Scum Sludge Pump 1        | VSD    | MCC-01            |
| MOT-001 | Grit Blower 1             | FSD    | MCC-01            |
| XV-001  | Inlet Control Valve       | HVALVE | PCV-01            |
| MV-001  | Sludge Discharge Valve    | MVALVE | PCV-01            |
| LIT-001 | Headworks Overflow Level  | AI     | INS-01            |

**Rules:**
- `Tag` and `Description` are required
- `Type` must match exactly one of the equipment type columns in `SIGNAL_MATRIX`

---

## Output

The generated file is saved in an `IO List` subfolder next to the tool.

**File name format:**
```
ProjectName_YYYYMMDD_HHMMSS_IO_List.xlsx
```

**Output sheets:**
1. `IO List` — fully formatted IO list matching your template
2. `Generation Summary` — signal counts, warnings, unassigned signals

**IO list structure per rack:**
```
[Controller row]
  RACK 01 — [DI module header]
    DI signals...
  RACK 01 — [DO module header]
    DO signals...
  RACK 01 — [AI module header]
    AI signals...
  RACK 01 — [AO module header]
    AO signals...
  RACK 02 — ...
```

---

## Tips

- Use the **"Create template"** button in the GUI to generate a fresh input workbook
- If a rack does not have enough channels, unassigned signals are highlighted in red and listed in the summary sheet
- The tool does **not** overflow signals to the next rack — it flags and stops to let you adjust the hardware config
- You can run the tool multiple times; each run creates a new timestamped output file

---

*Generated by IO List Generator v1.0 — Anas Abu Ghanname*
