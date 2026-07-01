# UnimexFTZ-Build — FTZ Customs Declaration Processor

Standalone Windows tools that automate the preparation of **Foreign Trade Zone (FTZ) customs declarations** from raw shipment manifests and separation lists. Two executables, one for each transport mode:

| Program | Mode | Shipment ID format | Bag ID prefix |
|---|---|---|---|
| `UnimexFTZ.exe` | Ocean | SCAC + digits (e.g. `COSU1234567890`) | `ZXWR…` |
| `UnimexAir.exe` | Air | `999-` + 8 digits (e.g. `999-92338816`) | `CBZS…` |

A browser-based version (no installation required) is also available via the `docs/` web app.

**Department:** US Customs Brokerage  
**Current versions:** Ocean v1.6.0 · Air v1.1.0

## What It Does

For each shipment it:

1. **Pairs** each master manifest with its matching separation list by reading the shipment ID inside the files (filenames are not trusted).
2. **Filters** the master down to line items whose Bag ID appears in the separation list.
3. **Aggregates** those lines to one row per HS Code, summing Quantity, Weight, and Value.
4. **Applies customs rules** — round-up minimums (Weight < 1 kg → 1, Value < $1 → 1), fixed columns (Zone = P, Charges = 3, Country of Origin = CN).
5. **Outputs** a formatted Excel file per shipment, ready to submit.

## Quick Start (End Users)

1. Copy `UnimexFTZ.exe` (or `UnimexAir.exe`) to any folder.
2. Double-click once — creates `input\`, `output\`, `logs\` folders next to it.
3. Drop both files for each shipment into `input\` (master manifest + separation list).
4. Double-click again — processed files appear in `output\`.

> **Air only:** the separation list filename must contain the shipment ID in `999-########` form.

If something looks wrong, send the relevant `logs\run_...log` file to Andy.

## Building the Executables

Requires Python 3.10+ and the packages in `requirements.txt`.

```bat
build.bat
```

This installs dependencies, syncs the ocean processor into `docs/`, and runs PyInstaller twice to produce `dist\UnimexFTZ.exe` and `dist\UnimexAir.exe`.

## Tech Stack

- **Language:** Python 3.10+
- **Libraries:** pandas (data manipulation), openpyxl (Excel read/write/styling)
- **Packaging:** PyInstaller (single-file `.exe`, no Python required on end-user machines)
- **Web version:** static site in `docs/` mirroring the same processing logic

## Project Files

| File | Purpose |
|---|---|
| `ftz_processor.py` | Ocean processor source |
| `air_processor.py` | Air processor source (fork of ocean) |
| `build.bat` | One-step build script |
| `requirements.txt` | Python dependencies |
| `UnimexFTZ.spec` / `UnimexAir.spec` | PyInstaller specs |
| `docs/` | Browser-based version + README + changelog |
| `DOCUMENTATION.md` | Full technical documentation and user manual |
| `CHANGELOG.md` | Version history |

## Full Documentation

See [DOCUMENTATION.md](DOCUMENTATION.md) for the complete technical reference, implementation details, and user manual.
