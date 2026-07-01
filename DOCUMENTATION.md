# Unimex FTZ / Air Separation Processor — Project Documentation (Executable Version)

**Prepared by:** Andres Torres (andres.torres@unimexgroup.com)
**Date:** June 4, 2026
**Version:** Ocean (UnimexFTZ) v1.6.0 · Air (UnimexAir) v1.1.0

---

## 1. Summary — What It Does

The Unimex Separation Processor is a desktop tool that automates the
preparation of **Foreign Trade Zone (FTZ) customs declarations** from raw
shipment data.

Every inbound shipment arrives with two Excel files:

1. **The master manifest** — a large file listing every individual parcel
   (line item) in the shipment, with its bag, tracking number, HS code,
   weight, quantity, and declared value.
2. **The separation list** — a smaller file listing only the bags that need
   to be declared for the FTZ entry.

Manually cross-referencing these two files, filtering thousands of line items,
and totaling them by HS code is slow and error-prone. This tool does it
automatically. For each shipment it:

- **Pairs** each master manifest with its matching separation list, identifying
  which files belong together by reading the **shipment ID inside the files**
  (filenames are not trusted).
- **Filters** the master down to only the line items whose **Bag ID** appears
  in the separation list.
- **Aggregates** those line items into **one row per HS Code**, summing the
  Quantity, Weight, and Value.
- **Applies customs business rules** (round-up minimums, fixed declaration
  columns).
- **Outputs** a clean, formatted Excel file per shipment, ready to submit.

The tool ships as **two standalone Windows programs** that share the same
logic but handle two transport modes:

| Program | Mode | Shipment ID format | Bag ID prefix |
|---|---|---|---|
| `UnimexFTZ.exe` | Ocean | Carrier SCAC + digits (e.g. `COSU1234567890`) | `ZXWR…` |
| `UnimexAir.exe` | Air | `999-` + 8 digits (e.g. `999-92338816`) | `CBZS…` |

A browser-based version (no installation required) is also available via the
`docs/` web app.

---

## 2. How It Was Implemented

### 2.1 Technology stack

- **Language:** Python 3.10+
- **Libraries:** `pandas` (data manipulation and aggregation), `openpyxl`
  (reading/writing and styling Excel workbooks)
- **Packaging:** `PyInstaller` — bundles Python and all dependencies into a
  single `.exe` so end users need nothing installed.
- **Web version:** The same processing logic is mirrored in the browser via a
  static site under `docs/` (`index.html`, `app.js`, `style.css`).

### 2.2 Design philosophy

The core challenge is that **input files are inconsistent**: filenames vary,
column headers differ between clients, shipment IDs appear in different forms,
and whitespace/encoding noise is common. The tool is built to be tolerant of
all of this rather than demanding perfectly formatted input.

### 2.3 Key implementation details

**File classification by content, not filename.**
Each Excel file dropped into the input folder is opened and inspected. A file
is classified as a *master* if it contains the expected manifest columns, or as
a *separation list* if its cells contain bag IDs matching the expected pattern
(`ZXWR…` for ocean, `CBZS…` for air). Files that match neither are reported as
unrecognized.

**Flexible column matching (two-pass resolver).**
Master manifests don't always use identical headers. The `resolve_columns()`
function maps each required field (MWB, Bag ID, Tracking #, HS Code, Weight,
Quantity, Value) using:
1. **Exact alias match** — a curated list of known header spellings, compared
   case-insensitively and with all whitespace stripped (so `HS CODE`,
   `hscode`, and `hs code` are equivalent).
2. **Keyword fallback** — if no alias matches, any header containing the right
   keyword tokens (e.g. anything with "weight") is accepted, preferring the
   shortest matching header to avoid grabbing a catch-all column.

**Robust shipment-ID matching.**
Shipment IDs are extracted with regular expressions tuned to each mode. The
ocean pattern accepts carrier prefixes of 4–7 letters (e.g. `COSU`, `EGLV`,
`ZIMUSHH`) followed by digits. When a master records a bare-digit MWB but the
separation list carries the prefixed form, the tool **reconciles them by
suffix** — re-keying the shorter ID under the longer one so they pair up,
provided exactly one candidate matches.

**ID normalization.**
The `clean_id()` helper normalizes every ID before comparison: Unicode NFKC
normalization plus stripping of all whitespace, including non-breaking spaces
(`\xa0`), which are a frequent source of "invisible" mismatches.

**Aggregation and customs rules.**
Matched line items are grouped by HS code and summed. Business rules are then
applied to the **aggregated totals** (not per line item):
- Any final **Weight < 1 kg** or **Value < $1** is rounded **up to 1**.
- Fixed columns are appended: **Zone = "P"**, **Charges = 3**,
  **Country of Origin = "CN"**.
- Output is sorted ascending by HS code and split into ≤998-row sheets if
  necessary (rarely triggered after aggregation, but part of the spec).

**Formatted output.**
The output workbook is styled with `openpyxl`: a colored bold header row,
number formats per column (currency for Value, thousands separators, etc.),
auto-sized columns, and frozen header panes.

**End-user ergonomics.**
- Input/output/log folders are auto-created next to the executable.
- Paths are anchored to the `.exe`'s real location, so double-clicking works
  even when Windows sets the working directory to `System32`.
- Every run writes a timestamped log file (`logs/run_YYYYMMDD_HHMMSS.log`) via
  a tee that mirrors console output to disk.
- The console window pauses on "Press ENTER to close" so it doesn't flash shut.
- All processing is wrapped in error handling that prints a full traceback to
  the log for troubleshooting.

**Ocean vs. Air differences.**
The air processor (`air_processor.py`) is a deliberate fork of the ocean
processor (`ftz_processor.py`), kept separate to avoid destabilizing the
approved ocean build. The differences:
- Shipment IDs are `999-########` (no carrier letters) instead of SCAC+digits.
- Bag IDs use the `CBZS` prefix instead of `ZXWR`.
- For air separation lists the shipment ID lives **only in the filename**, not
  inside the file contents, so it is parsed from the filename.
- Air master exports use a `sheet1` worksheet (matched case-insensitively),
  alongside the legacy Chinese-export sheets (`表1`, `0`).

### 2.4 Build process

`build.bat` automates the whole build:
1. Installs dependencies from `requirements.txt`.
2. Syncs `ftz_processor.py` into `docs/` for the web version.
3. Runs PyInstaller twice (`--onefile --console --clean`) to produce
   `dist/UnimexFTZ.exe` (ocean) and `dist/UnimexAir.exe` (air).

---

## 3. Instruction Manual — How to Use It

> These steps apply to both `UnimexFTZ.exe` (ocean) and `UnimexAir.exe` (air).
> Use the program that matches your shipment's transport mode.

### Step 1 — Place the program

Copy the correct `.exe` into any folder you like (for example, a folder on your
Desktop). No installation is required.

### Step 2 — First run (creates the folders)

Double-click the `.exe` **once**. It creates three folders next to it:

```
input\    ← you drop shipment files here
output\   ← finished FTZ files appear here
logs\     ← run logs, for troubleshooting
```

The window will tell you the input folder is empty and to add files; press
ENTER to close it.

### Step 3 — Add your shipment files

Drop **both** files for each shipment into the `input\` folder:

- **The master manifest** — the large file with all the line items.
- **The separation list** — the small file with bag IDs and tracking numbers.

You may add files for **multiple shipments at once**. The tool figures out which
master goes with which separation list by reading the shipment ID, so you do
**not** need to rename anything.

> **Air shipments only:** make sure the separation list's **filename** contains
> the shipment ID in the form `999-########` (e.g. `999-92338816`). For air, the
> ID is read from the filename, so a separation list without it cannot be paired.

### Step 4 — Run the tool

Double-click the `.exe` **again**. It will:

1. Scan every Excel file in `input\`.
2. Classify each as a master or a separation list.
3. Pair them up by shipment ID.
4. Write one `{ShipmentID}_FTZ.xlsx` file per shipment into `output\`.

### Step 5 — Read the summary

When finished, the window shows a per-shipment status and a final count, then
waits for you to press ENTER. The output files are in the `output\` folder.

### Understanding the status messages

| Tag | Meaning |
|---|---|
| `[OK   ]` | Shipment processed successfully. |
| `[SKIP ]` | Couldn't process — usually a master with no separation list, or vice versa. |
| `[WARN ]` | Heads-up — e.g. a bag ID in the separation list was not found in the master. |
| `[?????]` | A file in `input\` was recognized as neither a master nor a separation list. |
| `[INFO ]` | Informational — e.g. files were paired by suffix, or a column matched via keyword fallback. |
| `[FAIL ]` | Processing of a shipment threw an error; details are in the log. |

### What the output looks like

Each output file contains one sheet (`FTZ`) with seven columns:

| HS Code | Quantity | Weight | Value | Zone | Charges | Country of Origin |
|---|---|---|---|---|---|---|

- **One row per unique HS code** — line items are summed within each code.
- **Sorted ascending** by HS code.
- **Zone** is always `P`, **Charges** is always `3`, **Country of Origin** is
  always `CN`.
- Any final **Value < $1** or **Weight < 1 kg** is rounded up to **1**.

### If something looks wrong

1. Open the `logs\` folder and find the most recent `run_...log` file.
2. Send that log file — along with the input files you used — to Andy.

### Using the web version (optional)

A browser-based version requires no installation. Open the hosted page, drag
your shipment files onto the drop zone, and download the processed results as a
ZIP. The logic mirrors the desktop tool.

### Advanced: custom input/output folders

From a command prompt you can point the program at different folders:

```
UnimexFTZ.exe "C:\path\to\input" "C:\path\to\output"
```

---

## 4. Changelog

### Air Processor (`UnimexAir.exe`)

#### v1.1.0 — 2026-07-01
**Fixed**
- Shipments whose bag IDs use a consolidator prefix other than `CBZS`
  (e.g. `ADAS…`) are now recognized. Bag IDs are matched as any 4-letter
  prefix + digits instead of a single hard-coded prefix. Previously the
  separation list was reported "not recognized" and the shipment was skipped.
- Shipments whose air-waybill prefix is not `999-` (e.g. `369-10313494`) are
  now paired correctly. Shipment IDs are matched as any 3-digit airline prefix
  + 8-digit serial. This mainly affected separation lists, whose ID is read
  from the filename.

**Changed**
- The harmless openpyxl "Workbook contains no default style" warning (emitted
  by some "By SKU" master exports) is suppressed so it no longer clutters the
  run log.

#### v1.0.0 — 2026-06-01
*Initial release of the AIR variant.*
- Standalone fork of the ocean processor for air shipments.
- Recognizes air shipment IDs in the `999-########` format (no carrier
  letters).
- Recognizes `CBZS`-prefixed bag IDs.
- Reads the separation list's shipment ID from the **filename** (air
  separation lists do not carry the ID in their contents).
- Supports `sheet1` master exports (case-insensitive), alongside legacy
  Chinese-export sheets.
- Dual-executable build: `build.bat` now produces both `UnimexFTZ.exe` (ocean)
  and `UnimexAir.exe` (air).

---

### Ocean Processor (`UnimexFTZ.exe`)

#### v1.6.0 — 2026-05-28
**Fixed**
- Shipments where the master manifest's MWB column has bare digits
  (e.g. `2323289462`) while the separation list has the prefixed form
  (e.g. `OOLU2323289462`) are now paired correctly. IDs are matched by suffix
  when one is a prefixed version of the other.

#### v1.5.0 — 2026-05-20
**Fixed**
- Shipment IDs with prefixes longer than 4 letters (e.g. `ZIMUSHH` for Zim
  Integrated Shipping) are now recognized. Carrier prefixes between 4 and 7
  letters are now supported.

#### v1.4.0 — 2026-05-19
**Fixed**
- Shipment IDs preceded by a letter (e.g. in filenames like
  `Estatus_de_GuíasEGLV...`) now parse correctly. The previous word-boundary
  regex didn't fire when a letter sat directly before the carrier prefix.
- Tracking numbers embedded in mixed-text cells (notes alongside a JMX number)
  are now detected for diagnostic logs.

#### v1.3.0 — 2026-05-15
**Added**
- Support for shipment ID formats from any shipping line, not just COSCO.
  Previously only `COSU…` identifiers were recognized; the tool now handles the
  standard 4-letter-prefix + digits format used by all major carriers (EGLV for
  Evergreen, MAEU for Maersk, HLCU for Hapag-Lloyd, MSCU for MSC, etc.).

#### v1.2.0 — 2026-05-14
**Added**
- New **Country of Origin** column appended to output (after Charges),
  populated with `CN` on every row.

#### v1.1.0 — 2026-05-14
**Fixed**
- Master manifest files were not recognized when the client used alternate
  column header names (e.g. `WEIGHT` vs `PARCEL WEIGHT`, `HSCODE` vs `HS CODE`,
  `TOTAL QTY` vs `PRODUCT QTY`). Column matching is now case-insensitive and
  whitespace-insensitive with an alias list.

#### v1.0.0 — 2026-05-13
*Initial release.*
- Pairs master manifest files with their FTZ separation lists based on the
  shipment ID found inside the files (filename-independent).
- Filters master line items by Bag ID against the separation list.
- Aggregates output to one row per HS Code with summed Quantity, Weight, and
  Value.
- Applies the round-up-to-1 rule on aggregated Weight and Value totals.
- Adds fixed columns: Zone = `P`, Charges = `3`.
- Outputs sorted ascending by HS code, split into 998-row sheets if exceeded.
- Auto-creates `input/`, `output/`, and `logs/` folders next to the `.exe`.
- Writes a timestamped log file for every run.

---

## Appendix — Project Files

| File | Purpose |
|---|---|
| `ftz_processor.py` | Ocean processor source. |
| `air_processor.py` | Air processor source. |
| `build.bat` | One-step build script (installs deps, builds both `.exe`s). |
| `requirements.txt` | Python dependencies (`pandas`, `openpyxl`, `pyinstaller`). |
| `UnimexFTZ.spec` / `UnimexAir.spec` | PyInstaller build specifications. |
| `docs/` | Browser-based version (static web app) plus README and changelog. |
| `input/` `output/` `logs/` | Working folders created at runtime (not committed). |

**Department this tool was intended for:** US Customs Brokerage

**Built with:** Python 3.10+, pandas, openpyxl, packaged as standalone Windows
executables via PyInstaller.

This is a finished product. Future bugs will handled accordingly. 

**Contact:** Andres Torres — andres.torres@unimexgroup.com
