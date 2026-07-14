# UnimexFTZ-Build — FTZ Customs Declaration Processor

Standalone Windows tools that automate the preparation of **Foreign Trade Zone (FTZ) customs declarations** from raw shipment manifests and separation lists. Two executables, one for each transport mode:

| Program | Mode | Shipment ID format | Bag ID prefix |
|---|---|---|---|
| `UnimexFTZ.exe` | Ocean | SCAC + digits (e.g. `COSU1234567890`) | `ZXWR…` |
| `UnimexAir.exe` | Air | `999-` + 8 digits (e.g. `999-92338816`) or a carrier booking ref (e.g. `ZIMUSHH32215153`) | 4-letter prefix + digits (e.g. `CBZS…`, `ADAS…`) |

A browser-based version (no installation required) is also available via the `docs/` web app.

**Department:** US Customs Brokerage  
**Current versions:** Ocean v1.7.0 · Air v1.3.0

## Automatic Updates

From Ocean v1.7.0 / Air v1.3.0 onward, the exes **update themselves** from GitHub Releases — no more hand-copying a new build to every machine. To avoid wasting time when the tool is working fine, an update check only runs after a run that had a problem (a skipped, failed, or unrecognized file). When a newer build exists it downloads, installs, and restarts automatically, then reprocesses the input so the previously-skipped files go through. Offline or GitHub unreachable → it prints a short note and keeps running the current version.

Install v1.7.0 / v1.3.0 **once**, manually, into a **user-writable folder** (e.g. a subfolder of Documents or `%LOCALAPPDATA%` — *not* Program Files, or the exe can't replace itself). Every later version arrives on its own. The first launch may show a one-time Windows "unknown publisher" SmartScreen prompt (the exes are unsigned); choose *More info → Run anyway*.

Publishing a new version is automated: push a `ocean-vX.Y.Z` or `air-vX.Y.Z` tag and GitHub Actions builds the exe and creates the Release (see [DOCUMENTATION.md](DOCUMENTATION.md) and `.github/workflows/release.yml`).

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

> **Air only:** the separation list filename must contain the shipment ID — either `999-########` form or a carrier booking reference like `ZIMUSHH32215153`.

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
| `_version.py` | Single source of truth for both tools' versions |
| `updater.py` | Self-updater (checks GitHub Releases, swaps the exe in place) |
| `build.bat` | One-step build script |
| `requirements.txt` | Python dependencies |
| `UnimexFTZ.spec` / `UnimexAir.spec` | PyInstaller specs |
| `.github/workflows/release.yml` | CI: builds the exe and publishes a Release on a version tag |
| `docs/` | Browser-based version + README + changelog |
| `DOCUMENTATION.md` | Full technical documentation and user manual |
| `CHANGELOG.md` | Version history |

## Full Documentation

See [DOCUMENTATION.md](DOCUMENTATION.md) for the complete technical reference, implementation details, and user manual.
