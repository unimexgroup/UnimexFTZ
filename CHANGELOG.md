# Changelog — UnimexFTZ-Build

All notable changes to the FTZ Customs Declaration Processor.

---

## Air Processor (`UnimexAir.exe`)

### [v1.3.1] — 2026-07-16

**Fixed**
- The repo moved from `andrestorres-unimex/UnimexFTZ` to `unimexgroup/UnimexFTZ`. The self-updater's hard-coded owner is updated to match, so already-installed exes keep auto-updating instead of relying on GitHub's transfer redirect indefinitely.

### [v1.3.0] — 2026-07-14

**Added**
- **Automatic updates from GitHub Releases.** The exe now updates itself, so it no longer has to be re-copied to each machine by hand. To avoid wasting time when everything is working, it only checks for a new version on a run that had a problem — a skipped, failed, or unrecognized file (exactly when a newer version might handle it). If a newer build is found it downloads, installs, and restarts automatically, then reprocesses the input so the previously-skipped files go through. Clean runs never touch the network. If the machine is offline or GitHub is unreachable, it prints a short note and keeps running the current version.
- Startup banner now shows the version (e.g. `Unimex Air Processor (UnimexAir)  v1.3.0`).

*Note:* auto-update takes effect from this version forward — install v1.3.0 once (manually), and v1.3.1+ will arrive on their own. Install the exe in a **user-writable** folder (e.g. under Documents or `%LOCALAPPDATA%`), not Program Files, or it can't replace itself.

### [v1.2.0] — 2026-07-14

**Fixed**
- Shipments whose ID is a carrier booking reference (e.g. `ZIMUSHH32215153`) instead of an air waybill (`999-########`) are now paired correctly. The separation list's ID is read from its filename; previously the filename parser only recognized the `###-########` waybill format, so a separation list named `ZIMUSHH32215153.xlsx` was reported "no shipment ID" and the shipment was skipped. Booking references (4-letter carrier prefix + letters/digits) are now recognized as a fallback, with the air-waybill format still tried first.

### [v1.1.0] — 2026-07-01

**Fixed**
- Shipments whose bag IDs use a consolidator prefix other than `CBZS` (e.g. `ADAS…`) are now recognized. Bag IDs are matched as any 4-letter prefix + digits instead of a single hard-coded prefix. Previously the separation list was reported "not recognized" and the shipment was skipped.
- Shipments whose air-waybill prefix is not `999-` (e.g. `369-10313494`) are now paired correctly. Shipment IDs are matched as any 3-digit airline prefix + 8-digit serial. This mainly affected separation lists, whose ID is read from the filename.

**Changed**
- The harmless openpyxl "Workbook contains no default style" warning (emitted by some "By SKU" master exports) is suppressed so it no longer clutters the run log.

### [v1.0.0] — 2026-06-01

*Initial release of the AIR variant.*

- Standalone fork of the ocean processor for air shipments.
- Recognizes air shipment IDs in `999-########` format (no carrier letters).
- Recognizes `CBZS`-prefixed bag IDs.
- Reads the separation list's shipment ID from the **filename** (air separation lists do not carry the ID in their contents).
- Supports `sheet1` master exports (case-insensitive), alongside legacy Chinese-export sheets (`表1`, `0`).
- Dual-executable build: `build.bat` now produces both `UnimexFTZ.exe` (ocean) and `UnimexAir.exe` (air).

---

## Ocean Processor (`UnimexFTZ.exe`)

### [v1.8.1] — 2026-07-16

**Fixed**
- The repo moved from `andrestorres-unimex/UnimexFTZ` to `unimexgroup/UnimexFTZ`. The self-updater's hard-coded owner is updated to match, so already-installed exes keep auto-updating instead of relying on GitHub's transfer redirect indefinitely.

### [v1.8.0] — 2026-07-16

**Fixed**
- **"By SKU" master manifests (sheet named `sheet1`) are now recognized.** Master detection previously only looked at sheets named `表1` or `0`; a master like `NLD_PVG_COSU6504476123(By SKU).xlsx` fell through to the bag-ID scan, was misclassified as a separation list, and overwrote the real separation file — so the shipment was skipped with "separation list found but no master manifest". Sheet names are now matched case-insensitively and `sheet1` is accepted, same as the Air processor has done since Air v1.0.0.

**Changed**
- The harmless openpyxl "Workbook contains no default style" warning (emitted by "By SKU" exports) is suppressed, matching Air v1.1.0.

### [v1.7.0] — 2026-07-14

**Added**
- **Automatic updates from GitHub Releases.** Same mechanism as Air v1.3.0: the exe self-updates instead of being hand-copied to each machine. It only checks for a new version after a run that skipped, failed, or didn't recognize a file — clean runs never check. A newer build downloads, installs, and restarts automatically, then reprocesses the input. Offline/unreachable simply falls back to the current version with a short note.
- Startup banner now shows the version (`Unimex FTZ Processor  v1.7.0`).

*Note:* takes effect from this version forward — install v1.7.0 once (manually), then v1.7.1+ arrive automatically. Install in a **user-writable** folder (not Program Files) so the exe can replace itself.

### [v1.6.0] — 2026-05-28

**Fixed**
- Shipments where the master manifest's MWB column has bare digits (e.g. `2323289462`) while the separation list has the prefixed form (e.g. `OOLU2323289462`) are now paired correctly. IDs are matched by suffix when one is a prefixed version of the other.

### [v1.5.0] — 2026-05-20

**Fixed**
- Shipment IDs with prefixes longer than 4 letters (e.g. `ZIMUSHH` for Zim Integrated Shipping) are now recognized. Carrier prefixes between 4 and 7 letters are now supported.

### [v1.4.0] — 2026-05-19

**Fixed**
- Shipment IDs preceded by a letter (e.g. in filenames like `Estatus_de_GuíasEGLV...`) now parse correctly. The previous word-boundary regex didn't fire when a letter sat directly before the carrier prefix.
- Tracking numbers embedded in mixed-text cells (notes alongside a JMX number) are now detected for diagnostic logs.

### [v1.3.0] — 2026-05-15

**Added**
- Support for shipment ID formats from any shipping line, not just COSCO. Previously only `COSU…` identifiers were recognized; the tool now handles the standard 4-letter-prefix + digits format used by all major carriers (EGLV, MAEU, HLCU, MSCU, etc.).

### [v1.2.0] — 2026-05-14

**Added**
- New **Country of Origin** column appended to output (after Charges), populated with `CN` on every row.

### [v1.1.0] — 2026-05-14

**Fixed**
- Master manifest files were not recognized when the client used alternate column header names (e.g. `WEIGHT` vs `PARCEL WEIGHT`, `HSCODE` vs `HS CODE`, `TOTAL QTY` vs `PRODUCT QTY`). Column matching is now case-insensitive and whitespace-insensitive with an alias list.

### [v1.0.0] — 2026-05-13

*Initial release.*

- Pairs master manifests with FTZ separation lists based on the shipment ID found inside the files (filename-independent).
- Filters master line items by Bag ID against the separation list.
- Aggregates output to one row per HS Code with summed Quantity, Weight, and Value.
- Applies the round-up-to-1 rule on aggregated Weight and Value totals.
- Adds fixed columns: Zone = `P`, Charges = `3`.
- Outputs sorted ascending by HS code, split into 998-row sheets if exceeded.
- Auto-creates `input/`, `output/`, and `logs/` folders next to the `.exe`.
- Writes a timestamped log file for every run.
