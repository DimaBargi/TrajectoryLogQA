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

APP_TITLE = "Davidoff Physics Trajectory Log"
APP_SUBTITLE = "ניתוח קבצי Trajectory Log (.bin) - TrueBeam"

# full output columns (CSV/Excel)
COLUMNS = ["filename", "patient_id", "plan_name", "beam_name", "energy", "mlc_model",
           "num_snapshots", "sampling_interval_ms", "duration_s", "num_subbeams",
           "mu_planned", "mu_delivered", "mlc_rms_avg_mm", "mlc_rms_max_mm",
           "mlc_err_p95_mm", "mlc_err_max_mm", "max_gantry_error_deg",
           "max_collimator_error_deg", "max_jaw_error_cm", "num_beam_holds",
           "beam_hold_fraction", "beam_on_s", "is_truncated", "is_imaging",
           "plan_uid", "error"]

# columns shown in the on-screen table (key subset)
TREE_COLS = [
    ("filename", "קובץ", 230),
    ("patient_id", "מטופל", 80),
    ("beam_name", "שדה", 90),
    ("energy", "אנרגיה", 70),
    ("duration_s", "משך (ש')", 70),
    ("mu_delivered", "MU", 70),
    ("mlc_rms_avg_mm", "MLC RMS", 75),
    ("mlc_rms_max_mm", "MLC max", 75),
    ("max_gantry_error_deg", "גנטרי err", 75),
    ("num_beam_holds", "holds", 55),
    ("is_truncated", "חתוך", 55),
]


def scan_bin(root):
    out = []
    for dp, _, fns in os.walk(root):
        for fn in fns:
            if fn.lower().endswith(".bin") and not fn.startswith("._"):
                out.append(os.path.join(dp, fn))
    out.sort()
    return out


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
        # session-timing helpers
        row["_dt"] = TB_parse_dt(os.path.basename(path))
        rows.append(row)
    return rows, total, fails


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

        flt = ttk.Frame(self.master); flt.pack(fill="x", padx=12, pady=(2, 4))
        ttk.Button(flt, text="ייצא ל-Excel", command=self.export_xlsx).pack(side="left")
        ttk.Button(flt, text="פתח CSV", command=self.open_csv).pack(side="left", padx=6)
        ttk.Button(flt, text="זמן בחדר (sessions)", command=self.show_sessions).pack(side="left", padx=6)
        ttk.Label(flt, text=":חיפוש").pack(side="right")
        e = ttk.Entry(flt, textvariable=self.search, width=16); e.pack(side="right", padx=6)
        e.bind("<KeyRelease>", lambda ev: self.refresh_table())
        ttk.Label(flt, text=":סוג").pack(side="right", padx=(8, 0))
        cb = ttk.Combobox(flt, textvariable=self.f_type, width=10, state="readonly",
                          values=["הכל", "טיפול", "הדמיה", "חתוך"])
        cb.pack(side="right", padx=4); cb.bind("<<ComboboxSelected>>", lambda ev: self.refresh_table())

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
        self.summary.config(text="לוגים: {0}  |  טיפול: {1}  |  הדמיה: {2}  |  חתוכים: {3}  |  כשלים: {4}".format(
            total, total - n_img - fails, n_img, n_tr, fails))
        self.status.config(text="הסתיים." + ("  נשמר: " + out if out else ""))
        self.refresh_table()
        messagebox.showinfo("הסתיים", "נותחו {0} לוגים ({1} כשלים).".format(total, fails))

    def _on_err(self, msg):
        self.running = False; self.run_btn.config(state="normal")
        self.status.config(text="שגיאה."); messagebox.showerror("שגיאה", msg)

    def _filtered(self):
        q = self.search.get().strip().lower()
        ft = self.f_type.get()
        out = []
        for r in self.rows:
            if ft == "טיפול" and (r.get("is_imaging") or r.get("error")):
                continue
            if ft == "הדמיה" and not r.get("is_imaging"):
                continue
            if ft == "חתוך" and not r.get("is_truncated"):
                continue
            if q:
                hay = (str(r.get("filename", "")) + " " + str(r.get("patient_id", "")) +
                       " " + str(r.get("plan_name", ""))).lower()
                if q not in hay:
                    continue
            out.append(r)
        return out

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
