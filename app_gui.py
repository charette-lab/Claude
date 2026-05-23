"""Simple GUI for querying the Athanase Azure SQL sample DB.

Run from source:  python app_gui.py
Build executable: see build_exe.bat (PyInstaller)
"""

import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

import pyodbc


CONNECTION_STRING = (
    "Driver={ODBC Driver 18 for SQL Server};"
    "Server=tcp:msqlserver-aip.database.windows.net,1433;"
    "Database=mySampleDatabase;"
    "Uid=readonly_user;"
    "Pwd=H,dc9axiq51z;"
    "Encrypt=yes;"
    "Connection Timeout=30;"
)


QUERIES = {
    "Server version": "SELECT @@VERSION",
    "List tables": (
        "SELECT TABLE_SCHEMA + '.' + TABLE_NAME AS TableName "
        "FROM INFORMATION_SCHEMA.TABLES "
        "WHERE TABLE_TYPE='BASE TABLE' "
        "ORDER BY TABLE_SCHEMA, TABLE_NAME"
    ),
    "10 cheapest by EV/EBITA": (
        "SELECT TOP 10 Company_Name, [EV/EBITA], Country_of_Headquarters, "
        "GICS_Industry_Group_Name "
        "FROM dbo.PublicFinalAnalyticsActive "
        "WHERE [EV/EBITA] IS NOT NULL AND [EV/EBITA] > 0 "
        "AND Company_Name IS NOT NULL "
        "ORDER BY [EV/EBITA] ASC"
    ),
    "10 most expensive by EV/EBITA": (
        "SELECT TOP 10 Company_Name, [EV/EBITA], Country_of_Headquarters, "
        "GICS_Industry_Group_Name "
        "FROM dbo.PublicFinalAnalyticsActive "
        "WHERE [EV/EBITA] IS NOT NULL "
        "AND Company_Name IS NOT NULL "
        "ORDER BY [EV/EBITA] DESC"
    ),
    "10 highest 7Y return": (
        "SELECT TOP 10 m.Company_Name, s.Return_7Y, s.EV_EBITA_7Y "
        "FROM dbo.ShareReturnsPlusEVEBITA s "
        "JOIN dbo.PublicCompaniesMeta m ON m.Instrument = s.Instrument "
        "WHERE s.Return_7Y IS NOT NULL "
        "ORDER BY s.Return_7Y DESC"
    ),
}


class App:
    def __init__(self, root: tk.Tk) -> None:
        root.title("Athanase SQL Explorer")
        root.geometry("960x600")

        top = ttk.Frame(root, padding=8)
        top.pack(fill="x")

        ttk.Label(top, text="Query:").pack(side="left")
        self.query_var = tk.StringVar(value=next(iter(QUERIES)))
        self.combo = ttk.Combobox(
            top, textvariable=self.query_var, values=list(QUERIES.keys()),
            state="readonly", width=40,
        )
        self.combo.pack(side="left", padx=6)

        self.run_btn = ttk.Button(top, text="Run", command=self.run)
        self.run_btn.pack(side="left", padx=4)

        self.custom_btn = ttk.Button(top, text="Custom SQL...", command=self.custom)
        self.custom_btn.pack(side="left", padx=4)

        columns = ("c0", "c1", "c2", "c3", "c4", "c5", "c6", "c7")
        self.tree = ttk.Treeview(root, columns=columns, show="headings")
        self.tree.pack(fill="both", expand=True, padx=8, pady=4)

        self.status = tk.StringVar(value="Idle")
        ttk.Label(root, textvariable=self.status, anchor="w", relief="sunken"
                  ).pack(fill="x", side="bottom")

    def set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.run_btn.config(state=state)
        self.custom_btn.config(state=state)
        self.combo.config(state="disabled" if busy else "readonly")

    def run(self) -> None:
        label = self.query_var.get()
        sql = QUERIES[label]
        self._execute(sql, label)

    def custom(self) -> None:
        dlg = tk.Toplevel()
        dlg.title("Custom SQL")
        dlg.geometry("700x300")
        txt = scrolledtext.ScrolledText(dlg, wrap="word")
        txt.pack(fill="both", expand=True, padx=8, pady=8)
        txt.insert("1.0", "SELECT TOP 10 * FROM dbo.PublicFinalAnalyticsActive")

        def submit() -> None:
            sql = txt.get("1.0", "end").strip()
            dlg.destroy()
            self._execute(sql, "Custom SQL")

        ttk.Button(dlg, text="Run", command=submit).pack(pady=4)

    def _execute(self, sql: str, label: str) -> None:
        self.set_busy(True)
        self.status.set(f"Running: {label} ...")
        threading.Thread(target=self._worker, args=(sql, label), daemon=True).start()

    def _worker(self, sql: str, label: str) -> None:
        try:
            with pyodbc.connect(CONNECTION_STRING) as conn:
                cur = conn.cursor()
                cur.execute(sql)
                cols = [d[0] for d in cur.description] if cur.description else []
                rows = cur.fetchall() if cols else []
        except Exception as exc:
            self._on_error(exc)
            return
        self._on_result(label, cols, rows)

    def _on_result(self, label, cols, rows) -> None:
        def apply() -> None:
            for col_id in self.tree["columns"]:
                self.tree.heading(col_id, text="")
                self.tree.column(col_id, width=0, stretch=False)
            self.tree.delete(*self.tree.get_children())

            shown = cols[:8]
            for i, name in enumerate(shown):
                col_id = f"c{i}"
                self.tree.heading(col_id, text=name)
                self.tree.column(col_id, width=140, stretch=True, anchor="w")
            for row in rows:
                values = [("" if v is None else str(v)) for v in row[:8]]
                self.tree.insert("", "end", values=values)
            self.status.set(f"{label}: {len(rows)} row(s)")
            self.set_busy(False)
        self.tree.after(0, apply)

    def _on_error(self, exc: Exception) -> None:
        def apply() -> None:
            messagebox.showerror("Query failed", str(exc))
            self.status.set("Error")
            self.set_busy(False)
        self.tree.after(0, apply)


def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
