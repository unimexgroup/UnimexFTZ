# Changelog — UnimexFTZ-Build

All notable changes to the FTZ Customs Declaration Processor.

---

## Air Processor (`UnimexAir.exe`)

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
