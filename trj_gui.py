# -*- coding: utf-8 -*-
"""
trj_gui.py  -  Davidoff Physics Trajectory Log
==============================================
תוכנה גרפית לניתוח קבצי Trajectory Log (.bin) של TrueBeam.
- בחירת תיקייה, הרצה עם פס התקדמות (סריקה רקורסיבית של *.bin).
- טבלת סיכום לכל לוג: מטופל, תוכנית, אנרגיה, MLC RMS, שגיאות צירים, beam holds, beam-on, snapshots.
- חלון "זמן בחדר": קיבוץ לפי מטופל+תאריך -> משך ליד המכונה, beam-on, תקורה.
- ייצוא ל-CSV ול-Excel.

מנוע: trjlog_bin (פרסר בינארי טהור, מאומת מול pylinac). תלות: numpy בלבד.
"""
from __future__ import print_function

import csv
import os
import re
import sys
import threading
from datetime import datetime, timedelta

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except Exception:
    sys.stderr.write("tkinter is not available.\n")
    raise

try:
    import trjlog_bin as TB
    _ENGINE_OK = True
except Exception as _e:
    _ENGINE_OK = False
    _ENGINE_ERR = str(_e)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    _OPENPYXL_OK = True
except Exception:
    _OPENPYXL_OK = False

try:
    import matplotlib
    matplotlib.use("TkAgg")
    import matplotlib.pyplot as plt  # noqa: F401  (kept for Figure styling)
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg,
                                                   NavigationToolbar2Tk)
    _MPL_OK = True
except Exception:
    _MPL_OK = False

# Flag thresholds for "suspect delivery accuracy" (editable).
# flu_* are the DOSIMETRIC-impact metrics (actual-vs-expected fluence) and lead the
# danger ranking; the rest are the mechanical metrics.
CLINICAL = {
    "flu_err_pct": 1.0,           # % integral fluence error (dosimetric, primary)
    "flu_max_local_pct": 5.0,     # % peak local fluence error
    "mlc_rms_max_mm": 0.5,        # mm
    "mlc_err_max_mm": 1.0,        # mm
    "max_gantry_error_deg": 0.5,  # deg
    "max_jaw_error_cm": 0.1,      # cm (= 1 mm)
    "mu_diff_pct": 0.5,           # %
}
Z_THRESHOLD = 3.0                 # statistical outlier: z > 3 vs the cohort
FLAG_METRICS = list(CLINICAL.keys())
# QA (ArcCHECK) action limit — a measured gamma below this is a real dosimetric fail.
QA_ACTION_LIMIT = 90.0
QA_MATCH_DAYS = 4                 # match a log to a QA measurement within ±N days

try:
    from pdfminer.converter import PDFPageAggregator
    from pdfminer.layout import LAParams, LTChar, LTTextContainer
    from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
    from pdfminer.pdfpage import PDFPage
    _PDFMINER_OK = True
except Exception:
    _PDFMINER_OK = False


def _qa_layout(pdf_path):
    """Reconstruct visual rows of a QA PDF by y-grouping (same as MPC/GammaQA)."""
    rsrcmgr = PDFResourceManager()
    device = PDFPageAggregator(rsrcmgr, laparams=LAParams())
    interp = PDFPageInterpreter(rsrcmgr, device)
    out = []
    with open(pdf_path, "rb") as fp:
        for page in PDFPage.get_pages(fp):
            interp.process_page(page)
            chars = []

            def walk(o):
                for c in o:
                    if isinstance(c, LTChar):
                        chars.append(c)
                    elif isinstance(c, LTTextContainer):
                        walk(c)
            walk(device.get_result())
            rows = {}
            for c in chars:
                rows.setdefault(int(round(c.y0 / 3.0)), []).append(c)
            for k in sorted(rows, reverse=True):
                out.append("".join(ch.get_text() for ch in sorted(rows[k], key=lambda c: c.x0)))
    return "\n".join(out)


def _qa_parse(text):
    """Extract (patient_id, date, percent_passed, device) from a QA report."""
    import datetime as _dt
    pid = re.search(r"Patient\s*ID\s*:?\s*(\S+)", text)
    pct = re.search(r"%\s*Passed\s*:?\s*([0-9.]+)", text)
    # Prefer "Plan Date" (M/D/YYYY); the header "Date:" is D/M/YYYY and ambiguous.
    dm = re.search(r"Plan\s*Date\s*:?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})", text)
    dev = "ArcCHECK" if "ArcCHECK" in text[:800] else ("MapCHECK" if "MapCHECK" in text[:800] else "")
    date = None
    if dm:
        try:
            date = _dt.datetime.strptime(dm.group(1), "%m/%d/%Y").date()
        except ValueError:
            date = None
    return {"patient_id": pid.group(1) if pid else "",
            "date": date, "percent_passed": float(pct.group(1)) if pct else None,
            "device": dev}


def load_qa_folder(folder, progress_cb):
    """Parse all QA PDFs in a folder → list of measurement dicts."""
    pdfs = []
    for dp, _, fns in os.walk(folder):
        for fn in fns:
            if fn.lower().endswith(".pdf"):
                pdfs.append(os.path.join(dp, fn))
    pdfs.sort()
    meas = []
    for i, p in enumerate(pdfs, 1):
        progress_cb(i, len(pdfs), os.path.basename(p))
        try:
            r = _qa_parse(_qa_layout(p))
            if r["patient_id"] and r["percent_passed"] is not None:
                meas.append(r)
        except Exception:  # noqa: BLE001
            pass
    return meas


def match_qa(rows, measurements):
    """Attach measured_gamma_pct to each log row by patient_id + nearest date (±days)."""
    from collections import defaultdict
    by_pid = defaultdict(list)
    for m in measurements:
        if m["date"]:
            by_pid[str(m["patient_id"])].append(m)
    n = 0
    for r in rows:
        r["measured_gamma_pct"] = ""
        dt = r.get("_dt")
        pid = str(r.get("patient_id", ""))
        if dt is None or pid not in by_pid:
            continue
        ld = dt.date()
        best = min(by_pid[pid], key=lambda m: abs((m["date"] - ld).days))
        if abs((best["date"] - ld).days) <= QA_MATCH_DAYS:
            r["measured_gamma_pct"] = best["percent_passed"]
            n += 1
    return n

APP_TITLE = "Davidoff Physics Trajectory Log"
APP_SUBTITLE = "ניתוח קבצי Trajectory Log (.bin) - TrueBeam"

# full output columns (CSV/Excel)
COLUMNS = ["filename", "machine", "patient_id", "plan_name", "beam_name", "energy", "mlc_model",
           "num_snapshots", "sampling_interval_ms", "duration_s", "num_subbeams",
           "mu_planned", "mu_delivered", "mu_diff_pct",
           "flu_err_pct", "flu_max_local_pct",
           "mlc_rms_avg_mm", "mlc_rms_max_mm",
           "mlc_err_p95_mm", "mlc_err_max_mm", "mlc_bias_mm", "max_gantry_error_deg",
           "max_collimator_error_deg", "max_jaw_error_cm", "num_beam_holds",
           "beam_hold_fraction", "beam_on_s", "measured_gamma_pct", "is_truncated", "is_imaging",
           "plan_uid", "error"]

# columns shown in the on-screen table (key subset)
TREE_COLS = [
    ("filename", "קובץ", 200),
    ("machine", "מאיץ", 60),
    ("patient_id", "מטופל", 80),
    ("beam_name", "שדה", 90),
    ("energy", "אנרגיה", 70),
    ("duration_s", "משך (ש')", 70),
    ("mu_delivered", "MU", 65),
    ("flu_err_pct", "fluence Δ%", 75),
    ("mlc_rms_max_mm", "MLC max", 70),
    ("max_gantry_error_deg", "גנטרי err", 70),
    ("num_beam_holds", "holds", 50),
    ("measured_gamma_pct", "γ נמדד", 60),
    ("is_truncated", "חתוך", 50),
]


def scan_bin(root):
    out = []
    for dp, _, fns in os.walk(root):
        for fn in fns:
            if fn.lower().endswith(".bin") and not fn.startswith("._"):
                out.append(os.path.join(dp, fn))
    out.sort()
    return out


# Machine name is NOT stored in the .bin — derive it from the folder path.
# Recognized linac tokens (Varian naming): H193868, LA1, TrueBeam3868, SN1022.
_MACHINE_PATTERNS = [
    re.compile(r"H(\d{5,6})", re.I),          # H193868 -> 3868 (Varian host id)
    re.compile(r"\b(LA\d+)\b", re.I),         # LA1, LA3
    re.compile(r"TrueBeam\s*(\d+)", re.I),    # TrueBeam3868
    re.compile(r"\bSN\s*(\d+)\b", re.I),      # SN1022
]


def machine_from_path(path, root):
    """Best-effort machine label from the path. Tries known linac tokens; falls
    back to the first folder component under the chosen root."""
    rel = os.path.relpath(path, root)
    parts = rel.replace("\\", "/").split("/")
    hay = rel.replace("\\", "/")
    for pat in _MACHINE_PATTERNS:
        m = pat.search(hay)
        if m:
            tok = m.group(1)
            # H193868 -> last 4 digits are the common machine id (3868)
            if pat.pattern.startswith("H") and len(tok) >= 4:
                return tok[-4:]
            return tok.upper()
    # fallback: top folder under root (the machine's directory), else "?"
    return parts[0] if len(parts) > 1 else "?"


def analyze_folder(root, progress_cb):
    files = scan_bin(root)
    rows, fails = [], 0
    total = len(files)
    for i, path in enumerate(files, 1):
        progress_cb(i, total, os.path.basename(path))
        try:
            log = TB.load_log(path)
            row = TB.qa_metrics(log)
            row["error"] = ""
        except Exception as e:                       # corrupt/truncated: skip, continue
            row = {"filename": os.path.basename(path), "error": "%s: %s" % (type(e).__name__, e)}
            fails += 1
        # session-timing + per-log re-load helpers
        row["_dt"] = TB_parse_dt(os.path.basename(path))
        row["_path"] = path
        row["machine"] = machine_from_path(path, root)
        rows.append(row)
    return rows, total, fails


# ------------------------- outlier flagging -------------------------
def cohort_stats(rows):
    """mean/std per flag-metric over treatment (non-imaging, error-free) logs."""
    import statistics
    stats = {}
    for k in FLAG_METRICS:
        vals = [abs(r[k]) for r in rows if not r.get("is_imaging") and not r.get("error")
                and isinstance(r.get(k), (int, float))]
        if len(vals) >= 5:
            stats[k] = (statistics.mean(vals), statistics.pstdev(vals) or 1e-9)
    return stats


def flag_log(row, stats):
    """(score, reasons) for one treatment log: clinical thresholds + z>3 statistical."""
    if row.get("is_imaging") or row.get("error"):
        return 0.0, []
    reasons, score = [], 0.0
    for k in FLAG_METRICS:
        v = row.get(k)
        if not isinstance(v, (int, float)):
            continue
        av = abs(v)
        thr = CLINICAL.get(k)
        if thr is not None and av > thr:
            reasons.append("{0} {1:g}>{2:g}".format(k, av, thr)); score += 100.0
        if k in stats:
            mu, sd = stats[k]
            z = (av - mu) / sd
            if z > Z_THRESHOLD:
                reasons.append("{0} {1:.1f}σ".format(k, z)); score = max(score, z)
    if row.get("is_truncated"):
        reasons.append("truncated"); score += 100.0
    return round(score, 2), reasons


def TB_parse_dt(filename):
    import re
    m = re.search(r"(\d{14})(?=\.\w+$)", filename)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y%m%d%H%M%S")
    except ValueError:
        return None


def sessions_from_rows(rows):
    """Group logs into (patient, date) sessions -> time-at-machine analytics."""
    from collections import defaultdict
    sess = defaultdict(list)
    for r in rows:
        dt = r.get("_dt")
        if dt is None:
            continue
        pid = r.get("patient_id") or r.get("filename", "").split("_")[0]
        dur = r.get("duration_s") or 0.0
        sess[(pid, dt.date())].append((dt, float(dur), r))
    out = []
    for (pid, day), items in sess.items():
        items.sort()
        start = items[0][0]
        end = max(dt + timedelta(seconds=dur) for dt, dur, _ in items)
        span = (end - start).total_seconds()
        beamon = sum(d for _, d, _ in items)
        out.append({
            "patient_id": pid, "date": day.isoformat(), "n_logs": len(items),
            "first": start.strftime("%H:%M:%S"), "last": end.strftime("%H:%M:%S"),
            "time_at_machine_min": round(span / 60, 1),
            "beam_on_min": round(beamon / 60, 2),
            "overhead_min": round((span - beamon) / 60, 1),
        })
    out.sort(key=lambda s: (s["date"], s["first"]))
    return out


def write_csv(rows, out_path):
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow(r)


def write_xlsx(rows, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Trajectory QA"
    hf = PatternFill("solid", fgColor="305496")
    hfont = Font(bold=True, color="FFFFFF")
    trunc_fill = PatternFill("solid", fgColor="F8CBAD")
    img_fill = PatternFill("solid", fgColor="FFEB9C")
    ws.append(COLUMNS)
    for c in ws[1]:
        c.fill, c.font, c.alignment = hf, hfont, Alignment(horizontal="center")
    for r in rows:
        ws.append([r.get(k, "") for k in COLUMNS])
        fill = trunc_fill if r.get("is_truncated") else img_fill if r.get("is_imaging") else None
        if fill:
            for c in ws[ws.max_row]:
                c.fill = fill
    ws.freeze_panes = "A2"
    wb.save(out_path)


class App(object):
    def __init__(self, master):
        self.master = master
        master.title(APP_TITLE)
        master.geometry("1120x720")
        self.folder = tk.StringVar()
        self.f_type = tk.StringVar(value="הכל")
        self.f_machine = tk.StringVar(value="הכל")
        self.f_date_from = tk.StringVar()
        self.f_date_to = tk.StringVar()
        self.search = tk.StringVar()
        self.rows = []
        self.out_csv = None
        self.running = False
        self._sort_col = None
        self._sort_reverse = False
        self._build()

    def _build(self):
        head = ttk.Frame(self.master); head.pack(fill="x", padx=12, pady=(10, 2))
        ttk.Label(head, text=APP_TITLE, font=("Helvetica", 18, "bold")).pack(anchor="center")
        ttk.Label(head, text=APP_SUBTITLE, font=("Helvetica", 10)).pack(anchor="center")
        ttk.Separator(self.master).pack(fill="x", padx=12, pady=6)

        ctl = ttk.Frame(self.master); ctl.pack(fill="x", padx=12)
        ttk.Label(ctl, text=":תיקיית קבצי Trajectory Log (.bin)").pack(anchor="e")
        row = ttk.Frame(ctl); row.pack(fill="x")
        ttk.Button(row, text="בחר תיקייה...", command=self.choose).pack(side="right")
        ttk.Entry(row, textvariable=self.folder).pack(side="right", fill="x", expand=True, padx=6)

        opt = ttk.Frame(self.master); opt.pack(fill="x", padx=12, pady=6)
        self.run_btn = ttk.Button(opt, text="הרץ ניתוח", command=self.run)
        self.run_btn.pack(side="right")
        self.prog = ttk.Progressbar(self.master, mode="determinate")
        self.prog.pack(fill="x", padx=12, pady=2)
        self.status = ttk.Label(self.master, text="מוכן."); self.status.pack(anchor="e", padx=12)
        self.summary = ttk.Label(self.master, text="", font=("Helvetica", 11, "bold"))
        self.summary.pack(anchor="center", pady=4)
        ttk.Label(self.master, font=("Helvetica", 9), foreground="#666",
                  text="מקרא:  ◻ טיפול תקין   ▦ צהוב=הדמיה   ▦ אדום=חתוך   |   "
                       "MLC RMS/max=דיוק עלים (מ\"מ)   |   'מקרא / עזרה' להסבר מלא").pack(anchor="center")

        flt = ttk.Frame(self.master); flt.pack(fill="x", padx=12, pady=(2, 4))
        ttk.Button(flt, text="מקרא / עזרה", command=self.show_legend).pack(side="left", padx=(0, 8))
        ttk.Button(flt, text="ייצא ל-Excel", command=self.export_xlsx).pack(side="left")
        ttk.Button(flt, text="פתח CSV", command=self.open_csv).pack(side="left", padx=6)
        ttk.Button(flt, text="זמן בחדר", command=self.show_sessions).pack(side="left", padx=6)
        ttk.Button(flt, text="התפלגות", command=self.show_histograms).pack(side="left", padx=2)
        ttk.Button(flt, text="מגמה", command=self.show_trend).pack(side="left", padx=2)
        ttk.Button(flt, text="גרף לוג", command=self.show_log_graph).pack(side="left", padx=2)
        ttk.Button(flt, text="מפת fluence", command=self.show_fluence).pack(side="left", padx=2)
        ttk.Button(flt, text="outliers / מטופלים", command=self.show_outliers).pack(side="left", padx=2)
        ttk.Button(flt, text="טען QA (ArcCHECK)", command=self.load_qa).pack(side="left", padx=2)
        ttk.Label(flt, text=":חיפוש").pack(side="right")
        e = ttk.Entry(flt, textvariable=self.search, width=16); e.pack(side="right", padx=6)
        e.bind("<KeyRelease>", lambda ev: self.refresh_table())
        ttk.Label(flt, text=":סוג").pack(side="right", padx=(8, 0))
        cb = ttk.Combobox(flt, textvariable=self.f_type, width=10, state="readonly",
                          values=["הכל", "טיפול", "הדמיה", "חתוך"])
        cb.pack(side="right", padx=4); cb.bind("<<ComboboxSelected>>", lambda ev: self.refresh_table())

        # second filter row: machine + date range
        flt2 = ttk.Frame(self.master); flt2.pack(fill="x", padx=12, pady=(0, 2))
        ttk.Label(flt2, text=":מאיץ").pack(side="right")
        self.cb_machine = ttk.Combobox(flt2, textvariable=self.f_machine, width=10,
                                       state="readonly", values=["הכל"])
        self.cb_machine.pack(side="right", padx=4)
        self.cb_machine.bind("<<ComboboxSelected>>", lambda ev: self.refresh_table())
        ttk.Label(flt2, text=":תאריך מ (YYYY-MM-DD)").pack(side="right", padx=(10, 0))
        ef = ttk.Entry(flt2, textvariable=self.f_date_from, width=12); ef.pack(side="right", padx=4)
        ef.bind("<KeyRelease>", lambda ev: self.refresh_table())
        ttk.Label(flt2, text=":עד").pack(side="right")
        et = ttk.Entry(flt2, textvariable=self.f_date_to, width=12); et.pack(side="right", padx=4)
        et.bind("<KeyRelease>", lambda ev: self.refresh_table())
        ttk.Button(flt2, text="נקה סינון", command=self._clear_filters).pack(side="left")

        tf = ttk.Frame(self.master); tf.pack(fill="both", expand=True, padx=12, pady=8)
        cols = [c[0] for c in TREE_COLS]
        self.tree = ttk.Treeview(tf, columns=cols, show="headings")
        self._titles = {}
        for key, title, w in TREE_COLS:
            self._titles[key] = title
            self.tree.heading(key, text=title, command=lambda k=key: self._sort_by(k))
            self.tree.column(key, width=w, anchor="center")
        self.tree.column("filename", anchor="w")
        vs = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True); vs.pack(side="right", fill="y")
        self.tree.tag_configure("trunc", background="#f8d7da")
        self.tree.tag_configure("imaging", background="#fff3cd")
        self.tree.bind("<Double-1>", lambda e: self.show_log_graph())

        if not _ENGINE_OK:
            self.status.config(text="שגיאת מנוע: " + _ENGINE_ERR)
            self.run_btn.config(state="disabled")

    def choose(self):
        d = filedialog.askdirectory(title="בחר תיקייה עם קבצי .bin")
        if d:
            self.folder.set(d)

    def run(self):
        if self.running:
            return
        root = self.folder.get().strip()
        if not root or not os.path.isdir(root):
            messagebox.showwarning("שים לב", "יש לבחור תיקייה תקינה.")
            return
        self.running = True
        self.run_btn.config(state="disabled")
        for i in self.tree.get_children():
            self.tree.delete(i)
        t = threading.Thread(target=self._worker, args=(root,)); t.daemon = True; t.start()

    def _worker(self, root):
        def prog(d, t, m):
            self.master.after(0, self._on_prog, d, t, m)
        try:
            rows, total, fails = analyze_folder(root, prog)
            out = os.path.join(root, "trajectory_qa_results.csv")
            try:
                write_csv(rows, out)
            except Exception:
                out = None
            self.master.after(0, self._on_done, rows, total, fails, out)
        except Exception as e:
            self.master.after(0, self._on_err, str(e))

    def _on_prog(self, d, t, m):
        if t:
            self.prog["maximum"] = t; self.prog["value"] = d
        self.status.config(text="מעבד {0}/{1}  {2}".format(d, t, m[:38]))

    def _on_done(self, rows, total, fails, out):
        self.running = False; self.run_btn.config(state="normal")
        self.rows = rows; self.out_csv = out
        n_img = sum(1 for r in rows if r.get("is_imaging"))
        n_tr = sum(1 for r in rows if r.get("is_truncated"))
        machines = sorted(set(r.get("machine") for r in rows if r.get("machine")))
        self.cb_machine["values"] = ["הכל"] + machines
        self.summary.config(text="לוגים: {0}  |  מאיצים: {1}  |  טיפול: {2}  |  הדמיה: {3}  |  חתוכים: {4}  |  כשלים: {5}".format(
            total, len(machines), total - n_img - fails, n_img, n_tr, fails))
        self.status.config(text="הסתיים." + ("  נשמר: " + out if out else ""))
        self.refresh_table()
        messagebox.showinfo("הסתיים", "נותחו {0} לוגים מ-{1} מאיצים ({2} כשלים).".format(
            total, len(machines), fails))

    def _on_err(self, msg):
        self.running = False; self.run_btn.config(state="normal")
        self.status.config(text="שגיאה."); messagebox.showerror("שגיאה", msg)

    def _filtered(self):
        q = self.search.get().strip().lower()
        ft = self.f_type.get()
        machine = self.f_machine.get()
        d_from = self.f_date_from.get().strip()
        d_to = self.f_date_to.get().strip()
        out = []
        for r in self.rows:
            if ft == "טיפול" and (r.get("is_imaging") or r.get("error")):
                continue
            if ft == "הדמיה" and not r.get("is_imaging"):
                continue
            if ft == "חתוך" and not r.get("is_truncated"):
                continue
            if machine not in ("הכל", "") and r.get("machine") != machine:
                continue
            if d_from or d_to:
                dt = r.get("_dt")
                ds = dt.date().isoformat() if dt else ""
                if not ds:
                    continue                       # undated logs drop out when filtering by date
                if d_from and ds < d_from:
                    continue
                if d_to and ds > d_to:
                    continue
            if q:
                hay = (str(r.get("filename", "")) + " " + str(r.get("patient_id", "")) +
                       " " + str(r.get("plan_name", ""))).lower()
                if q not in hay:
                    continue
            out.append(r)
        return out

    def _clear_filters(self):
        self.f_type.set("הכל")
        self.f_machine.set("הכל")
        self.f_date_from.set("")
        self.f_date_to.set("")
        self.search.set("")
        self.refresh_table()

    def refresh_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for r in self._filtered():
            vals = []
            for key, _, _ in TREE_COLS:
                v = r.get(key, "")
                vals.append("{0:g}".format(v) if isinstance(v, float) else v)
            tag = "trunc" if r.get("is_truncated") else "imaging" if r.get("is_imaging") else ""
            self.tree.insert("", "end", values=vals, tags=(tag,))
        if self.rows:
            self.status.config(text="מוצגות {0} מתוך {1}".format(len(self._filtered()), len(self.rows)))

    def _sort_by(self, col):
        if self._sort_col == col:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_col, self._sort_reverse = col, False
        items = [(self.tree.set(i, col), i) for i in self.tree.get_children("")]

        def key(t):
            try:
                return (0, float(t[0]), "")
            except ValueError:
                return (1, 0.0, str(t[0]).lower())
        items.sort(key=key, reverse=self._sort_reverse)
        for idx, (_, i) in enumerate(items):
            self.tree.move(i, "", idx)
        arrow = " ▼" if self._sort_reverse else " ▲"
        for k, t in self._titles.items():
            self.tree.heading(k, text=t + (arrow if k == col else ""))

    def open_csv(self):
        if self.out_csv and os.path.exists(self.out_csv):
            self._open(self.out_csv)
        else:
            messagebox.showinfo("מידע", "עדיין לא בוצע ניתוח.")

    def export_xlsx(self):
        if not self.rows:
            messagebox.showinfo("מידע", "עדיין לא בוצע ניתוח."); return
        if not _OPENPYXL_OK:
            messagebox.showwarning("Excel לא זמין", "openpyxl לא מותקן. השתמש ב-CSV."); return
        folder = os.path.dirname(self.out_csv) if self.out_csv else self.folder.get()
        out = os.path.join(folder, "trajectory_qa_results.xlsx")
        try:
            write_xlsx(self._filtered(), out)
        except Exception as e:
            messagebox.showerror("שגיאה", str(e)); return
        if messagebox.askyesno("הושלם", "נשמר:\n{0}\n\nלפתוח?".format(out)):
            self._open(out)

    def show_sessions(self):
        if not self.rows:
            messagebox.showinfo("מידע", "עדיין לא בוצע ניתוח."); return
        sess = sessions_from_rows(self.rows)
        if not sess:
            messagebox.showinfo("מידע", "אין חותמות זמן בשמות הקבצים."); return
        win = tk.Toplevel(self.master); win.title("זמן בחדר - לפי session (מטופל × תאריך)")
        win.geometry("820x520")
        import statistics as st
        tam = [s["time_at_machine_min"] for s in sess]
        bo = [s["beam_on_min"] for s in sess]
        ttk.Label(win, font=("Helvetica", 11, "bold"),
                  text="{0} sessions | זמן ליד המכונה: חציון {1:.1f} ד' (ממוצע {2:.1f}) | beam-on חציון {3:.2f} ד'".format(
                      len(sess), st.median(tam), st.mean(tam), st.median(bo))).pack(pady=6)
        cols = ["date", "patient_id", "n_logs", "first", "last",
                "time_at_machine_min", "beam_on_min", "overhead_min"]
        titles = {"date": "תאריך", "patient_id": "מטופל", "n_logs": "לוגים", "first": "התחלה",
                  "last": "סיום", "time_at_machine_min": "ליד המכונה (ד')",
                  "beam_on_min": "beam-on (ד')", "overhead_min": "תקורה (ד')"}
        tf = ttk.Frame(win); tf.pack(fill="both", expand=True, padx=8, pady=8)
        tv = ttk.Treeview(tf, columns=cols, show="headings")
        for c in cols:
            tv.heading(c, text=titles[c]); tv.column(c, width=95, anchor="center")
        vs = ttk.Scrollbar(tf, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=vs.set)
        tv.pack(side="left", fill="both", expand=True); vs.pack(side="right", fill="y")
        for s in sess:
            tv.insert("", "end", values=[s[c] for c in cols])
        # export sessions
        def exp_sess():
            folder = os.path.dirname(self.out_csv) if self.out_csv else self.folder.get()
            out = os.path.join(folder, "sessions_time_in_room.csv")
            with open(out, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=cols); w.writeheader(); w.writerows(sess)
            messagebox.showinfo("הושלם", "נשמר:\n" + out)
        ttk.Button(win, text="ייצא sessions ל-CSV", command=exp_sess).pack(pady=4)

    # ---------- מקרא / עזרה ----------
    def show_legend(self):
        win = tk.Toplevel(self.master)
        win.title("מקרא והסבר - Trajectory Log QA")
        win.geometry("780x660")
        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True)
        txt = tk.Text(frame, wrap="word", font=("Helvetica", 11), padx=12, pady=12)
        sb = ttk.Scrollbar(frame, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)
        c = CLINICAL
        content = (
            "מקרא והסבר\n==========\n\n"
            "תיקייה ומאיצים\n--------------\n"
            "אפשר להצביע על תיקיית-אב שמכילה את כל המאיצים (סריקה רקורסיבית).\n"
            "שם המאיץ נגזר מנתיב התיקייה (H193868->3868, LA1, TrueBeam3597, SN1022).\n"
            "סינון: מאיץ (תפריט) | טווח תאריכים (מ/עד, YYYY-MM-DD) | חיפוש מטופל | סוג.\n\n"
            "עמודות הטבלה\n------------\n"
            "קובץ        - שם קובץ הלוג (.bin)\n"
            "מאיץ        - המאיץ (נגזר מהנתיב)\n"
            "מטופל       - מספר זהות המטופל\n"
            "שדה         - שם השדה / הקשת\n"
            "אנרגיה      - אנרגיית הקרן (6x, 10x, 10xFFF...)\n"
            "משך (ש')    - משך ההקלטה בשניות (snapshots x 0.02)\n"
            "MU          - מוניטור-יוניטס שנמסרו בפועל\n"
            "fluence Δ%  - הפרש פלואנס נמסר↔מתוכנן (%) = ההשפעה הדוזימטרית [המדד העיקרי]\n"
            "MLC max     - שגיאת מיקום העלה המקסימלית (מ\"מ) - מכני\n"
            "גנטרי err   - שגיאת זווית הגנטרי המקסימלית (מעלות)\n"
            "holds       - מספר עצירות קרן (gating / מעברי שדה / השהיות)\n"
            "γ נמדד      - gamma שנמדד ב-ArcCHECK (אחרי 'טען QA') = אמת המנה האמיתית\n"
            "חתוך        - האם הלוג נקטע (טיפול > 20 דק'). True = נתונים חלקיים\n\n"
            "מה באמת מסוכן?\n--------------\n"
            "fluence Δ% (דוזימטרי) מוביל את הדירוג - הוא אומר אם הסטייה המכנית\n"
            "באמת שינתה את הפלואנס/מנה. 'מפת fluence' מראה איפה על השדה.\n"
            "systematic (בלוח מטופלים) = הטיית עלים עקבית בין פרקציות -\n"
            "מצטברת לאורך הקורס, מסוכנת יותר מרעש אקראי.\n"
            "γ נמדד מתחת ל-90% = כשל דוזימטרי אמיתי במדידה.\n\n"
            "צבעי שורות\n----------\n"
            "לבן  - לוג טיפול תקין ומלא\n"
            "צהוב - לוג הדמיה (kV/CBCT/setup) - אין MLC/מנה\n"
            "אדום - לוג חתוך (truncated) - נתונים חלקיים\n\n"
            "גבולות לדגל \"חשוד לאי-דיוק\"\n---------------------------\n"
            "התוכנה מדגלת לוג בשתי דרכים:\n\n"
            "1) ספים קליניים קבועים (מבוסס MPPG 8.b / TG-142):\n"
            "   fluence Δ > {flu:g}%  (דוזימטרי, עיקרי)\n"
            "   MLC RMS   > {rms:g} מ\"מ\n"
            "   MLC max   > {emax:g} מ\"מ\n"
            "   גנטרי err > {gan:g} מעלות\n"
            "   סורגים    > {jaw:g} מ\"מ\n"
            "   פער MU    > {mu:g}%\n"
            "   לוג חתוך  = תמיד מדגל\n\n"
            "2) חריג סטטיסטי: ערך מעל ממוצע + 3-sigma של המכונה שלך עצמה\n"
            "   (חשוב במיוחד כי הערכים בפועל נמוכים בהרבה מהספים הקליניים)\n\n"
            "לוג מסומן מופיע בחלון \"outliers / מטופלים\" עם הסיבה והניקוד.\n\n"
            "חשוב\n----\n"
            "הדגלים מודדים דיוק *מכני* (MLC/גנטרי/סורגים/MU) - מועמדים לאי-דיוק\n"
            "במנה. אישור דוזימטרי מצריך מדידה (ArcCHECK). הלוג אינו רואה שגיאות\n"
            "מודל-מנה ב-TPS או סחף כיול/פלט של המאיץ.\n"
        ).format(flu=c["flu_err_pct"], rms=c["mlc_rms_max_mm"], emax=c["mlc_err_max_mm"],
                 gan=c["max_gantry_error_deg"], jaw=c["max_jaw_error_cm"] * 10,
                 mu=c["mu_diff_pct"])
        txt.insert("1.0", content)
        txt.tag_add("all", "1.0", "end")
        txt.tag_configure("all", justify="right", spacing1=1)
        txt.configure(state="disabled")
        ttk.Button(win, text="סגור", command=win.destroy).pack(pady=4)

    # ---------- גרפיקה (matplotlib) ----------
    def _need_mpl(self):
        if not _MPL_OK:
            messagebox.showwarning("גרפיקה לא זמינה", "matplotlib לא מותקן בגרסה זו.")
            return False
        if not self.rows:
            messagebox.showinfo("מידע", "עדיין לא בוצע ניתוח.")
            return False
        return True

    def _fig_window(self, title, figsize=(9, 6)):
        win = tk.Toplevel(self.master)
        win.title(title)
        fig = Figure(figsize=figsize, dpi=100)
        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(canvas, win).update()
        return win, fig, canvas

    def _tx_rows(self):
        """Filtered treatment logs (non-imaging, error-free)."""
        return [r for r in self._filtered()
                if not r.get("is_imaging") and not r.get("error")]

    def _selected_row(self):
        sel = self.tree.selection()
        if not sel:
            return None
        fn = self.tree.set(sel[0], "filename")
        for r in self.rows:
            if r.get("filename") == fn:
                return r
        return None

    def show_histograms(self):
        if not self._need_mpl():
            return
        import numpy as np
        rows = self._tx_rows()
        if not rows:
            messagebox.showinfo("מידע", "אין לוגי טיפול להצגה."); return
        specs = [("mlc_rms_max_mm", "MLC RMS max (mm)"),
                 ("mlc_err_max_mm", "MLC error max (mm)"),
                 ("max_gantry_error_deg", "Gantry error (deg)"),
                 ("num_beam_holds", "Beam holds")]
        win, fig, canvas = self._fig_window("התפלגות מדדים", (10, 7))
        for i, (key, label) in enumerate(specs, 1):
            ax = fig.add_subplot(2, 2, i)
            vals = [r[key] for r in rows if isinstance(r.get(key), (int, float))]
            if not vals:
                ax.set_title(label + " (אין נתונים)"); continue
            arr = np.asarray(vals, float)
            counts, bins, patches = ax.hist(arr, bins=40, color="#4e79a7")
            mu, sd = float(arr.mean()), float(arr.std())
            thr = CLINICAL.get(key)
            cut = max(mu + 3 * sd, thr if thr is not None else mu + 3 * sd)
            for p, b in zip(patches, bins[:-1]):
                if b >= cut:
                    p.set_facecolor("#e15759")
            ax.axvline(mu + 3 * sd, color="#e15759", ls="--", lw=1)
            if thr is not None:
                ax.axvline(thr, color="#f28e2b", ls=":", lw=1)
            n_out = int((arr > cut).sum())
            ax.set_title("{0}\nmean {1:.3g} · max {2:.3g} · חריגים {3}".format(
                label, mu, float(arr.max()), n_out), fontsize=9)
        fig.suptitle("התפלגות על {0} לוגי טיפול  (אדום = זנב חריג)".format(len(rows)))
        fig.tight_layout(); canvas.draw()

    def show_trend(self):
        if not self._need_mpl():
            return
        import numpy as np
        rows = [r for r in self._tx_rows() if r.get("_dt")]
        if not rows:
            messagebox.showinfo("מידע", "אין חותמות זמן בשמות הקבצים."); return
        rows.sort(key=lambda r: r["_dt"])
        dts = [r["_dt"] for r in rows]
        rms = [r.get("mlc_rms_max_mm") for r in rows]
        good = [v for v in rms if isinstance(v, (int, float))]
        mu = float(np.mean(good)); sd = float(np.std(good)); cut = mu + 3 * sd
        win, fig, canvas = self._fig_window("מגמה לאורך זמן", (10, 6))
        ax = fig.add_subplot(111)
        colors = ["#e15759" if isinstance(v, (int, float)) and v > cut else "#4e79a7" for v in rms]
        ax.scatter(dts, rms, c=colors, s=12)
        ax.axhline(cut, color="#e15759", ls="--", lw=1, label="mean+3σ ({0:.3g})".format(cut))
        ax.set_ylabel("MLC RMS max (mm)"); ax.set_xlabel("תאריך")
        ax.set_title("מגמת MLC RMS על {0} לוגי טיפול".format(len(rows)))
        ax.legend(); fig.autofmt_xdate(); fig.tight_layout(); canvas.draw()

    def show_log_graph(self):
        if not _MPL_OK:
            messagebox.showwarning("גרפיקה לא זמינה", "matplotlib לא מותקן."); return
        r = self._selected_row()
        if not r:
            messagebox.showinfo("בחר לוג", "בחר שורה בטבלה (לחיצה כפולה, או בחר ואז 'גרף לוג')."); return
        if r.get("error") or not r.get("_path"):
            messagebox.showwarning("שגיאה", "לוג זה לא נטען בהצלחה."); return
        import numpy as np
        try:
            log = TB.load_log(r["_path"])
        except Exception as e:
            messagebox.showerror("שגיאה", str(e)); return
        t = np.arange(log.num_snapshots) * log.sampling_interval_ms / 1000.0
        win, fig, canvas = self._fig_window("גרף לוג: " + r.get("filename", "")[:48], (10, 8))
        ax1 = fig.add_subplot(3, 1, 1)
        if not r.get("is_imaging") and "mlc" in log.axes:
            ts, mean_e, max_e = TB.mlc_error_series(log)
            ax1.plot(ts, max_e, color="#e15759", lw=0.8, label="max")
            ax1.plot(ts, mean_e, color="#4e79a7", lw=0.8, label="mean")
            ax1.set_ylabel("MLC error (mm)"); ax1.legend(loc="upper right", fontsize=8)
        else:
            ax1.set_title("לוג הדמיה - אין נתוני MLC")
        ax2 = fig.add_subplot(3, 1, 2, sharex=ax1)
        if "gantry" in log.axes:
            ge, ga = log.axes["gantry"]
            ax2.plot(t, ga, color="#59a14f", lw=0.8, label="actual")
            ax2.plot(t, ge, color="#999999", lw=0.8, ls="--", label="expected")
            ax2.set_ylabel("Gantry (deg)"); ax2.legend(loc="upper right", fontsize=8)
        ax3 = fig.add_subplot(3, 1, 3, sharex=ax1)
        if "mu" in log.axes:
            _, mu_a = log.axes["mu"]
            ax3.plot(t, mu_a, color="#4e79a7", lw=0.9)
            ax3.set_ylabel("MU"); ax3.set_xlabel("זמן (ש')")
        if "beam_hold" in log.axes:
            _, bh = log.axes["beam_hold"]
            hold = bh.astype(int) == 2
            for ax in (ax1, ax2, ax3):
                lo, hi = ax.get_ylim()
                ax.fill_between(t, lo, hi, where=hold, color="#ffd9a0", alpha=0.5, step="mid")
                ax.set_ylim(lo, hi)
        fig.suptitle("מטופל {0}  |  שדה {1}  |  {2}  |  holds מוצללים".format(
            r.get("patient_id", ""), r.get("beam_name", ""), r.get("energy", "")))
        fig.tight_layout(); canvas.draw()

    def show_fluence(self):
        if not _MPL_OK:
            messagebox.showwarning("גרפיקה לא זמינה", "matplotlib לא מותקן."); return
        r = self._selected_row()
        if not r:
            messagebox.showinfo("בחר לוג", "בחר שורה בטבלה (לוג טיפול)."); return
        if r.get("error") or not r.get("_path") or r.get("is_imaging"):
            messagebox.showwarning("לא זמין", "מפת fluence זמינה ללוגי טיפול בלבד."); return
        try:
            log = TB.load_log(r["_path"])
            g = TB.fluence_gamma(log, dose_pct=2, dist_mm=2, threshold=0.1)
        except Exception as e:
            messagebox.showerror("שגיאה", str(e)); return
        if not g:
            messagebox.showinfo("מידע", "אין fluence (לוג הדמיה?)."); return
        win, fig, canvas = self._fig_window("מפת fluence: " + r.get("filename", "")[:42], (11, 7))
        panels = [(g["actual"], "Actual (נמסר)", "viridis"),
                  (g["expected"], "Expected (מתוכנן)", "viridis"),
                  (g["diff_map"], "Difference act-exp", "RdBu_r"),
                  (g["gamma_map"], "Gamma 2%/2mm", "inferno")]
        for i, (arr, title, cmap) in enumerate(panels, 1):
            ax = fig.add_subplot(2, 2, i)
            im = ax.imshow(arr, aspect="auto", cmap=cmap)
            ax.set_title(title, fontsize=9)
            ax.set_xlabel("x (mm)", fontsize=7); ax.set_ylabel("leaf pair", fontsize=7)
            fig.colorbar(im, ax=ax, fraction=0.046)
        fig.suptitle("מטופל {0} | gamma pass {1}% | fluence Δ {2}%  (השוואת נמסר↔מתוכנן מהלוג)".format(
            r.get("patient_id", ""), g["pass_pct"], r.get("flu_err_pct", "")))
        fig.tight_layout(); canvas.draw()

    def load_qa(self):
        if not self.rows:
            messagebox.showinfo("מידע", "קודם הרץ ניתוח לוגים."); return
        if not _PDFMINER_OK:
            messagebox.showwarning("לא זמין", "pdfminer לא מותקן בגרסה זו."); return
        folder = filedialog.askdirectory(title="בחר תיקיית דוחות QA (ArcCHECK PDF)")
        if not folder:
            return
        win = tk.Toplevel(self.master); win.title("טוען QA..."); win.geometry("440x90")
        lbl = ttk.Label(win, text="טוען דוחות QA..."); lbl.pack(pady=8)
        pb = ttk.Progressbar(win, mode="determinate"); pb.pack(fill="x", padx=10)

        def prog(d, t, m):
            self.master.after(0, lambda: (pb.configure(maximum=t, value=d),
                                          lbl.configure(text="QA {0}/{1}".format(d, t))))

        def work():
            try:
                meas = load_qa_folder(folder, prog)
                n = match_qa(self.rows, meas)
                self.master.after(0, lambda: self._qa_done(win, len(meas), n))
            except Exception as e:
                self.master.after(0, lambda: (win.destroy(), messagebox.showerror("שגיאה", str(e))))
        th = threading.Thread(target=work); th.daemon = True; th.start()

    def _qa_done(self, win, nmeas, nmatch):
        win.destroy()
        self.refresh_table()
        messagebox.showinfo("QA נטען",
                            "נטענו {0} מדידות QA; הותאמו ל-{1} לוגים.\n"
                            "פתח 'outliers / מטופלים' לתצוגת חזוי-מול-נמדד.".format(nmeas, nmatch))

    def show_outliers(self):
        if not self.rows:
            messagebox.showinfo("מידע", "עדיין לא בוצע ניתוח."); return
        import statistics
        from collections import defaultdict
        rows = self._tx_rows()
        if not rows:
            messagebox.showinfo("מידע", "אין לוגי טיפול."); return
        stats = cohort_stats(self.rows)
        flagged = []
        for r in rows:
            score, reasons = flag_log(r, stats)
            if reasons:
                flagged.append((score, r, reasons))
        flagged.sort(key=lambda x: -x[0])
        flagged_ids = set(id(r) for _, r, _ in flagged)

        win = tk.Toplevel(self.master); win.title("Outliers / per-patient"); win.geometry("1000x660")
        ttk.Label(win, font=("Helvetica", 12, "bold"),
                  text="{0} לוגים מסומנים מתוך {1} לוגי טיפול".format(len(flagged), len(rows))).pack(pady=(6, 0))
        ttk.Label(win, foreground="#a02040", font=("Helvetica", 9),
                  text="דגלים מכניים בלבד (MLC/גנטרי/סורגים/MU) — מועמדים לאי-דיוק. "
                       "אישור דוזימטרי מצריך מדידה (ArcCHECK).").pack()
        nb = ttk.Notebook(win); nb.pack(fill="both", expand=True, padx=6, pady=6)

        # --- tab 1: flagged logs ---
        t1 = ttk.Frame(nb); nb.add(t1, text="לוגים מסומנים")
        cols = ["score", "date", "patient_id", "beam_name", "flu_err_pct", "measured_γ", "reasons"]
        tv = ttk.Treeview(t1, columns=cols, show="headings")
        for c, w in zip(cols, [55, 85, 80, 85, 80, 80, 360]):
            tv.heading(c, text=c); tv.column(c, width=w, anchor="center")
        tv.column("reasons", anchor="w")
        vs = ttk.Scrollbar(t1, orient="vertical", command=tv.yview); tv.configure(yscrollcommand=vs.set)
        tv.pack(side="left", fill="both", expand=True); vs.pack(side="right", fill="y")
        for score, r, reasons in flagged:
            d = r["_dt"].date().isoformat() if r.get("_dt") else ""
            tv.insert("", "end", values=[score, d, r.get("patient_id"), r.get("beam_name"),
                                         r.get("flu_err_pct"), r.get("measured_gamma_pct", ""),
                                         "; ".join(reasons)])

        # --- tab 2: per-patient rollup ---
        t2 = ttk.Frame(nb); nb.add(t2, text="לפי מטופל")
        bypat = defaultdict(list)
        for r in rows:
            bypat[r.get("patient_id")].append(r)
        pcols = ["patient_id", "n_fx", "flu_err_mean", "systematic", "n_flagged", "measured_γ_min"]
        ptv = ttk.Treeview(t2, columns=pcols, show="headings")
        for c, w in zip(pcols, [110, 55, 105, 90, 80, 110]):
            ptv.heading(c, text=c); ptv.column(c, width=w, anchor="center")
        vs2 = ttk.Scrollbar(t2, orient="vertical", command=ptv.yview); ptv.configure(yscrollcommand=vs2.set)
        ptv.pack(side="left", fill="both", expand=True); vs2.pack(side="right", fill="y")
        prows = []
        for pid, rs in bypat.items():
            flu = [r["flu_err_pct"] for r in rs if isinstance(r.get("flu_err_pct"), (int, float))]
            bias = [r["mlc_bias_mm"] for r in rs if isinstance(r.get("mlc_bias_mm"), (int, float))]
            mg = [r["measured_gamma_pct"] for r in rs if isinstance(r.get("measured_gamma_pct"), (int, float))]
            nfl = sum(1 for r in rs if id(r) in flagged_ids)
            # systematic ratio: |mean bias| / std(bias). High + consistent sign = systematic
            # (accumulates across the course = more dangerous than random jitter)
            if len(bias) >= 2:
                mb = statistics.mean(bias); sb = statistics.pstdev(bias) or 1e-9
                sysr = round(abs(mb) / sb, 2)
            else:
                sysr = ""
            prows.append((pid, len(rs), round(statistics.mean(flu), 4) if flu else "",
                          sysr, nfl, round(min(mg), 1) if mg else ""))
        prows.sort(key=lambda x: (-x[4], -(x[2] if isinstance(x[2], float) else 0)))
        for pr in prows:
            tags = ("sysflag",) if (isinstance(pr[3], float) and pr[3] > 1.5) else ()
            ptv.insert("", "end", values=pr, tags=tags)
        ptv.tag_configure("sysflag", background="#ffe0b0")
        ttk.Label(t2, font=("Helvetica", 8), foreground="#a05000",
                  text="systematic > 1.5 (כתום) = הטיית עלים עקבית בין פרקציות = מצטברת לאורך הקורס").pack(side="bottom")

        # --- tab 3: bar chart of top outliers ---
        if _MPL_OK and flagged:
            t3 = ttk.Frame(nb); nb.add(t3, text="גרף")
            fig = Figure(figsize=(9, 5), dpi=100); ax = fig.add_subplot(111)
            top = flagged[:20]
            labels = [(f[1].get("patient_id", "") + " " +
                       (f[1]["_dt"].date().isoformat() if f[1].get("_dt") else "")) for f in top]
            ax.barh(range(len(top)), [f[0] for f in top], color="#e15759")
            ax.set_yticks(range(len(top))); ax.set_yticklabels(labels, fontsize=7)
            ax.invert_yaxis(); ax.set_xlabel("score"); ax.set_title("עד 20 הלוגים החריגים ביותר")
            c = FigureCanvasTkAgg(fig, master=t3); c.get_tk_widget().pack(fill="both", expand=True)
            NavigationToolbar2Tk(c, t3).update(); fig.tight_layout(); c.draw()

        def exp_csv():
            folder = os.path.dirname(self.out_csv) if self.out_csv else self.folder.get()
            out = os.path.join(folder, "outliers.csv")
            with open(out, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["score", "date", "patient_id", "beam_name", "mlc_rms_max_mm",
                            "mu_diff_pct", "reasons"])
                for score, r, reasons in flagged:
                    d = r["_dt"].date().isoformat() if r.get("_dt") else ""
                    w.writerow([score, d, r.get("patient_id"), r.get("beam_name"),
                                r.get("mlc_rms_max_mm"), r.get("mu_diff_pct"), "; ".join(reasons)])
            messagebox.showinfo("הושלם", "נשמר:\n" + out)
        ttk.Button(win, text="ייצא outliers ל-CSV", command=exp_csv).pack(pady=4)

    def _open(self, path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore
            elif sys.platform == "darwin":
                os.system('open "%s"' % path)
            else:
                os.system('xdg-open "%s"' % path)
        except Exception:
            messagebox.showinfo("מידע", "נשמר ב:\n" + path)


def main():
    root = tk.Tk(); App(root); root.mainloop()


if __name__ == "__main__":
    main()
