"""Athanase Revaluation — desktop GUI for re-running the engine on fresh prices.

An analyst picks a price file (Instrument, Market Cap, EV), the tool patches those
onto the baseline screener fundamentals, re-runs the ROIIC-DCF engine, and writes a
results spreadsheet (ER@7%/ER@12% + carry/re-rate decomposition per name).

Run from source:  python revalue_gui.py
Build executable: see build_revalue_exe.bat (PyInstaller)

Data files the tool needs (place next to the .exe, or Browse to them):
  * baseline screener  (e.g. LTM_baseline.xlsx) — the stable fundamentals
  * roiic_dcf.py       — the valuation engine (bundled by the build script)
  * (optional) a scored book to value on the researched moat
"""
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def _resource_dirs():
    """Where bundled/companion files might live (PyInstaller onefile + dev)."""
    dirs = [os.path.dirname(os.path.abspath(__file__))]
    if getattr(sys, "frozen", False):
        dirs = [getattr(sys, "_MEIPASS", ""), os.path.dirname(sys.executable)] + dirs
    return [d for d in dirs if d]


def _find(name):
    for d in _resource_dirs():
        p = os.path.join(d, name)
        if os.path.exists(p):
            return p
    return ""


# Make local modules importable and point the engine env at a bundled roiic_dcf.py
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
for d in _resource_dirs():
    sys.path.insert(0, d)
_engine = _find("roiic_dcf.py")
if _engine and not os.environ.get("AIP_VALUE_ENGINE"):
    os.environ["AIP_VALUE_ENGINE"] = _engine


class App:
    def __init__(self, root):
        root.title("Athanase Revaluation")
        root.geometry("980x620")
        self.root = root

        frm = ttk.Frame(root, padding=10)
        frm.pack(fill="x")

        self.prices = tk.StringVar()
        self.baseline = tk.StringVar(value=_find("LTM_baseline.xlsx"))
        self.moats = tk.StringVar()
        self.out = tk.StringVar()
        self.re = tk.StringVar(value="7")
        self.re2 = tk.StringVar(value="12")

        self._row(frm, 0, "Analyst prices (Instrument, Market Cap, EV):", self.prices,
                  self._pick_prices, required=True)
        self._row(frm, 1, "Baseline screener (fundamentals):", self.baseline, self._pick_baseline)
        self._row(frm, 2, "Researched moats book (optional):", self.moats, self._pick_moats)
        self._row(frm, 3, "Output spreadsheet:", self.out, self._pick_out)

        rates = ttk.Frame(frm); rates.grid(row=4, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(rates, text="Realistic discount % (re):").pack(side="left")
        ttk.Entry(rates, textvariable=self.re, width=6).pack(side="left", padx=(4, 16))
        ttk.Label(rates, text="Conservative discount % (re2):").pack(side="left")
        ttk.Entry(rates, textvariable=self.re2, width=6).pack(side="left", padx=4)

        btns = ttk.Frame(frm); btns.grid(row=5, column=0, columnspan=3, sticky="w", pady=8)
        self.run_btn = ttk.Button(btns, text="Revalue", command=self.run)
        self.run_btn.pack(side="left")
        self.open_btn = ttk.Button(btns, text="Open results", command=self._open, state="disabled")
        self.open_btn.pack(side="left", padx=6)
        frm.columnconfigure(1, weight=1)

        cols = ("ticker", "company", "er7", "er12", "carry", "rerate", "driver")
        heads = ("Ticker", "Company", "ER@re", "ER@re2", "Carry", "Re-rate", "Driver")
        self.tree = ttk.Treeview(root, columns=cols, show="headings")
        for c, h in zip(cols, heads):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=130 if c == "company" else 90, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=10, pady=4)

        self.status = tk.StringVar(value="Pick a price file and click Revalue.")
        ttk.Label(root, textvariable=self.status, anchor="w", relief="sunken").pack(fill="x", side="bottom")

    def _row(self, frm, r, label, var, cmd, required=False):
        ttk.Label(frm, text=label).grid(row=r, column=0, sticky="w", pady=2)
        ttk.Entry(frm, textvariable=var, width=70).grid(row=r, column=1, sticky="ew", padx=6)
        ttk.Button(frm, text="Browse...", command=cmd).grid(row=r, column=2)

    def _pick_prices(self):
        p = filedialog.askopenfilename(title="Analyst prices",
                                       filetypes=[("Spreadsheets", "*.xlsx *.csv"), ("All", "*.*")])
        if p:
            self.prices.set(p)
            if not self.out.get():
                base, _ = os.path.splitext(p)
                self.out.set(base + "_revalued.xlsx")

    def _pick_baseline(self):
        p = filedialog.askopenfilename(title="Baseline screener", filetypes=[("Excel", "*.xlsx")])
        if p: self.baseline.set(p)

    def _pick_moats(self):
        p = filedialog.askopenfilename(title="Scored book (researched moats)", filetypes=[("Excel", "*.xlsx")])
        if p: self.moats.set(p)

    def _pick_out(self):
        p = filedialog.asksaveasfilename(title="Save results as", defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")])
        if p: self.out.set(p)

    def _busy(self, b):
        self.run_btn.config(state="disabled" if b else "normal")

    def run(self):
        if not self.prices.get() or not os.path.exists(self.prices.get()):
            messagebox.showerror("Missing prices", "Pick a valid analyst price file."); return
        if not self.baseline.get() or not os.path.exists(self.baseline.get()):
            messagebox.showerror("Missing baseline",
                                 "Pick the baseline screener (fundamentals). It should sit next to "
                                 "this program as LTM_baseline.xlsx, or Browse to it."); return
        if not self.out.get():
            base, _ = os.path.splitext(self.prices.get()); self.out.set(base + "_revalued.xlsx")
        try:
            re = float(self.re.get()) / 100.0; re2 = float(self.re2.get()) / 100.0
        except ValueError:
            messagebox.showerror("Bad rate", "Discount rates must be numbers, e.g. 7 and 12."); return
        self._busy(True); self.status.set("Revaluing...")
        threading.Thread(target=self._worker, args=(re, re2), daemon=True).start()

    def _worker(self, re, re2):
        try:
            import revalue
            prices = revalue.load_prices(self.prices.get())
            if not prices:
                raise ValueError("No rows read. Needs columns 'Instrument' and 'Market Cap' (and 'EV').")
            moats = revalue.load_moats(self.moats.get()) if self.moats.get() else None
            rows = revalue.revalue(self.baseline.get(), prices, moats=moats, re=re, re2=re2,
                                   only_uploaded=True)
            ok, bad = revalue.write_results(self.out.get(), rows, re=re, re2=re2)
        except SystemExit as e:   # aip.engine() exits if the engine is missing
            self._error(f"Valuation engine not found. Keep roiic_dcf.py next to this program.\n\n{e}"); return
        except Exception as e:
            self._error(str(e)); return
        self._done(rows, ok, bad)

    def _fmt(self, v):
        return f"{v*100:.1f}%" if isinstance(v, (int, float)) else ""

    def _done(self, rows, ok, bad):
        def apply():
            self.tree.delete(*self.tree.get_children())
            for r in sorted([x for x in rows if "error" not in x], key=lambda x: -(x["er_7"] or -9)):
                self.tree.insert("", "end", values=(r["ticker"], (r["company"] or "")[:34],
                                 self._fmt(r["er_7"]), self._fmt(r["er_12"]), self._fmt(r["carry"]),
                                 self._fmt(r["rerate"]), r["driver"]))
            for r in [x for x in rows if "error" in x]:
                self.tree.insert("", "end", values=(r["ticker"], r["error"], "", "", "", "", ""))
            self.status.set(f"Revalued {ok} names ({bad} skipped). Saved to {self.out.get()}")
            self.open_btn.config(state="normal"); self._busy(False)
        self.root.after(0, apply)

    def _error(self, msg):
        def apply():
            messagebox.showerror("Revaluation failed", msg)
            self.status.set("Error"); self._busy(False)
        self.root.after(0, apply)

    def _open(self):
        path = self.out.get()
        if not os.path.exists(path):
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)              # noqa
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except Exception as e:
            messagebox.showinfo("Saved", f"Results saved to:\n{path}\n({e})")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
