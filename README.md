# Davidoff Physics – Trajectory Log QA

A click-button Windows app that analyzes TrueBeam **Trajectory Log `.bin`** files and
extracts **all the recorded data** + standard delivery-QA metrics, with a "time in room"
session view and CSV/Excel export.

## Engine
`trjlog_bin.py` — a **pure-Python** parser of the Varian *TrueBeam Trajectory Log File
Specification* (log v3.0/4.0). No pylinac/scipy/scikit-image → freezes into a small (~26 MB),
robust exe that runs on locked-down hospital PCs. **Validated against pylinac** on real LA1
logs: header, metadata (incl. Plan UID), gantry, MU exact; MLC RMS within **<0.3 %**.

Extracts: header (version, sampling, axes, MLC model, truncation, snapshots, subbeams),
metadata (PatientID, PlanName, **Plan UID**, MU, Energy, BeamName), every axis (expected +
actual) at 50 Hz, MLC leaves, beam holds, subbeam timing — plus MLC RMS/errors, axis errors,
beam-on, holds, duration, and per-session time-at-machine.

## Files
| File | Purpose |
|------|---------|
| `trj_gui.py` | the GUI (tkinter) |
| `trjlog_bin.py` | the validated binary parser (engine) |
| `requirements.txt` | numpy, openpyxl, pyinstaller |
| `build_exe.bat` | local Windows build → `dist\TrajectoryLogQA.exe` |
| `installer.iss` | Inno Setup → `Output\TrajectoryLogQA_Setup.exe` |
| `.github/workflows/build.yml` | GitHub Actions: builds exe + installer in the cloud |
| `קרא_אותי.txt` | Hebrew user guide (bundled into the installer) |

## How to get the click-button installer

**macOS cannot build a Windows `.exe`.** Two ways to produce it:

### Option A — GitHub Actions (recommended, no Windows/Python needed)
1. Push this folder to a GitHub repo.
2. The workflow runs automatically on `windows-latest` (Python 3.11 x64).
3. Download the **`TrajectoryLogQA-Windows`** artifact → it contains
   `TrajectoryLogQA_Setup.exe`. Double-click to install; a desktop icon is created.

### Option B — build on any Windows PC with Python 3.9+ x64
1. `build_exe.bat`  → `dist\TrajectoryLogQA.exe` (already runnable by double-click).
2. Open `installer.iss` in [Inno Setup](https://jrsoftware.org/isdl.php), press **F9** →
   `Output\TrajectoryLogQA_Setup.exe`.

## Usage
Run the app → **בחר תיקייה** (folder of `.bin` logs) → **הרץ ניתוח**. Results show in the
table and save to `trajectory_qa_results.csv`. Yellow rows = imaging; red = truncated.

**Graphics & analysis (matplotlib):**
- **התפלגות** — distribution histograms of MLC RMS / errors / holds (outlier tail in red).
- **מגמה** — MLC RMS over time (machine-drift detection).
- **גרף לוג** (or double-click a row) — per-snapshot traces: MLC error, gantry actual-vs-expected,
  MU, with beam-holds shaded → shows *where* in the delivery a deviation occurred.
- **outliers / מטופלים** — ranks suspect treatments by **clinical thresholds + statistical
  (mean+3σ) outliers**, with the triggering reason, plus a per-patient rollup.
- **זמן בחדר** — per-session time-at-machine.

> The outlier flags are **mechanical** delivery accuracy (MLC/gantry/jaw/MU) — candidates for dose
> inaccuracy. Dosimetric confirmation needs measurement (ArcCHECK); the log cannot see TPS
> dose-model or machine output/calibration errors. Thresholds are editable consts (`CLINICAL`,
> `Z_THRESHOLD`) at the top of `trj_gui.py`.

## Note on gamma/fluence
This app reports everything **recorded in the log** + mechanical/MLC QA. Fluence-map **gamma**
needs heavy dose reconstruction (pylinac/scipy) and is intentionally **not** bundled here to
keep the exe small and robust. It can be added as a separate optional module if required.
