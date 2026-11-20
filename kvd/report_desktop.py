"""Local P3 dashboard and English report review."""
import argparse
import sqlite3
from pathlib import Path
from tkinter import StringVar, Text, Tk, filedialog, font, messagebox, ttk

from openpyxl.utils.exceptions import InvalidFileException

from .journal import Journal
from .material_desktop import MaterialDashboard
from .report_data import FILTERS, SECTIONS, approved_snapshot, comparison, display_quantity, select
from .report_render import HEADERS, TITLES, content
from . import reports


class ReportingDesktop:
    def __init__(self, root, journal, workbook=""):
        self.root = root
        self.journal = journal
        self.snapshot = None
        self.visible = {}
        self.display_rows = []
        self.workbook = StringVar(value=workbook)
        self.feedback = StringVar(value="Select a committed synthetic workbook.")
        self.source = StringVar(value="No snapshot loaded")
        self.summary = StringVar()
        self.filters = {key: StringVar() for key in FILTERS}
        self.section = StringVar(value=SECTIONS["progress"])
        self.command = StringVar(value="MC10")
        self.anchor = StringVar()
        self.actor = StringVar()
        self.reviewer = StringVar()
        root.title("Construction Management | Engineering Dashboard")
        root.geometry("1320x850")
        root.minsize(1060, 720)
        self.build()
        self.refresh_history()

    def build(self):
        theme = ttk.Style()
        theme.theme_use("clam")
        fonts = {name.casefold() for name in font.families()}
        face = "Segoe UI" if "segoe ui" in fonts else "Helvetica"
        theme.configure(".", font=(face, 10), background="#edf3f8", foreground="#20364a")
        theme.configure("Header.TLabel", font=(face, 23, "bold"), foreground="#ffffff", background="#102d43")
        theme.configure("HeaderSub.TLabel", foreground="#bcdbe8", background="#102d43")
        theme.configure("Header.TFrame", background="#102d43")
        theme.configure("Accent.TButton", background="#006b65", foreground="#ffffff", padding=(10, 6))
        theme.map("Accent.TButton", background=[("active", "#07514d")])
        theme.configure("Treeview", rowheight=26, background="white", fieldbackground="white")
        theme.configure("Treeview.Heading", font=(face, 10, "bold"))
        theme.configure("TNotebook.Tab", padding=(13, 7))
        header = ttk.Frame(self.root, style="Header.TFrame", padding=(24, 13))
        header.pack(fill="x")
        ttk.Label(header, text="Engineering Management Dashboard", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header, textvariable=self.source, style="HeaderSub.TLabel").pack(anchor="w")
        main = ttk.Frame(self.root, padding=15)
        main.pack(fill="both", expand=True)
        ttk.Label(main, textvariable=self.feedback, wraplength=985).pack(side="bottom", anchor="w", pady=(8, 0))
        source_bar = ttk.Frame(main)
        source_bar.pack(fill="x", pady=(0, 8))
        ttk.Label(source_bar, text="Approved workbook").pack(side="left", padx=(0, 8))
        ttk.Entry(source_bar, textvariable=self.workbook).pack(side="left", fill="x", expand=True)
        ttk.Button(source_bar, text="Browse", command=self.browse).pack(side="left", padx=6)
        ttk.Button(source_bar, text="Load snapshot", command=lambda: self.call(self.load),
                   style="Accent.TButton").pack(side="left")
        self.tabs = ttk.Notebook(main)
        self.tabs.pack(fill="both", expand=True)
        dashboard = ttk.Frame(self.tabs, padding=11)
        self.reporting = ttk.Frame(self.tabs, padding=11)
        self.tabs.add(dashboard, text="Dashboard")
        self.tabs.add(self.reporting, text="Report review and issue")
        self.dashboard(dashboard)
        self.report_panel(self.reporting)
        materials = ttk.Frame(self.tabs, padding=11)
        self.tabs.add(materials, text="Material control")
        self.material_dashboard = MaterialDashboard(materials, self.journal, self.call)

    def call(self, callback):
        try:
            return callback()
        except (ValueError, OSError, sqlite3.Error, InvalidFileException) as error:
            self.feedback.set(str(error))
            messagebox.showerror("Review required", str(error), parent=self.root)

    def browse(self):
        path = filedialog.askopenfilename(parent=self.root, filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.workbook.set(path)

    @staticmethod
    def grid(parent, columns, labels, widths):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        grid = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        for col, label, width in zip(columns, labels, widths):
            grid.heading(col, text=label)
            grid.column(col, width=width)
        scroll = ttk.Scrollbar(frame, command=grid.yview)
        grid.configure(yscrollcommand=scroll.set)
        grid.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        return grid

    def dashboard(self, tab):
        bar = ttk.Frame(tab)
        bar.pack(fill="x")
        labels = ("From (YYYY-MM-DD)", "To (YYYY-MM-DD)", "Workfront ID", "Activity",
                  "Subcontractor", "Status")
        self.choices = {}
        for i, (key, label) in enumerate(zip(FILTERS, labels)):
            ttk.Label(bar, text=label).grid(row=0, column=i, sticky="w")
            if i < 2:
                box = ttk.Entry(bar, textvariable=self.filters[key], width=15)
            else:
                box = ttk.Combobox(bar, textvariable=self.filters[key], width=18, state="readonly")
                self.choices[key] = box
            box.grid(row=1, column=i, sticky="ew", padx=(0, 6))
            bar.columnconfigure(i, weight=1)
        ttk.Button(bar, text="Apply filters", style="Accent.TButton",
                   command=lambda: self.call(self.refresh_dashboard)).grid(row=1, column=6, padx=(4, 0))
        ttk.Button(bar, text="Reset", command=lambda: self.call(self.reset)).grid(row=1, column=7, padx=(5, 0))
        ttk.Label(tab, textvariable=self.summary, wraplength=990).pack(anchor="w", pady=(12, 6))
        section_bar = ttk.Frame(tab)
        section_bar.pack(fill="x", pady=(0, 6))
        ttk.Label(section_bar, text="Register view").pack(side="left", padx=(0, 8))
        box = ttk.Combobox(section_bar, textvariable=self.section, values=list(SECTIONS.values()),
                           width=28, state="readonly")
        box.pack(side="left")
        box.bind("<<ComboboxSelected>>", lambda _: self.show_section())
        self.detail = StringVar()
        ttk.Label(tab, textvariable=self.detail, wraplength=1000).pack(side="bottom", anchor="w", pady=(6, 0))
        self.data = self.grid(tab, ("a", "b", "c", "d", "e", "f", "g"), HEADERS, (150, 145, 140, 75, 100, 100, 100))
        self.data.bind("<<TreeviewSelect>>", self.show_detail)

    def load(self):
        self.snapshot = approved_snapshot(self.journal, self.workbook.get())
        self.source.set(f'Synthetic snapshot  |  {self.snapshot["file"]}  |  SHA-256 {self.snapshot["fingerprint"][:16]}')
        for key, box in self.choices.items():
            values = {row[key] for rows in self.snapshot["sections"].values() for row in rows if row[key]}
            box.configure(values=("", *sorted(values, key=str.casefold)))
        self.reset()
        self.feedback.set("Snapshot verified. Date filters exclude undated rows. Other filters match recorded fields.")

    def reset(self):
        for value in self.filters.values():
            value.set("")
        self.refresh_dashboard()

    def refresh_dashboard(self):
        if not self.snapshot:
            raise ValueError("Load an approved workbook snapshot first.")
        self.visible = select(self.snapshot, **{key: field.get() for key, field in self.filters.items()})
        self.summary.set("  |  ".join(
            f"{len(self.visible[key])} {label}" for key, label in
            (("progress", "site records"), ("technical", "technical issues"), ("quality", "QA/QC items"),
             ("toolbox", "toolbox / attendance"), ("readiness", "readiness records"),
             ("tasks", "junior tasks"), ("actions", "actions"))
        ))
        self.show_section()

    def show_section(self):
        self.data.delete(*self.data.get_children())
        self.detail.set("")
        section = next(key for key, label in SECTIONS.items() if label == self.section.get())
        records = self.visible.get(section, [])
        if section == "progress":
            self.display_rows = comparison(records)
            labels = HEADERS
            for group in self.display_rows:
                result = tuple(display_quantity(group["sums"][key], group["counts"][key], group["records"])
                               for key in ("planned", "reported", "verified"))
                dimensions = tuple(value or "Not recorded" for value in group["dimensions"])
                self.data.insert("", "end", values=dimensions + result)
            self.detail.set("Quantities stay grouped by workfront, activity, subcontractor and unit. "
                            "A fraction shows records with a supplied quantity.")
        else:
            self.display_rows = records
            labels = ("Record", "Description", "Date", "Workfront", "Status", "Source", "Details")
            for row in records:
                self.data.insert("", "end", values=(row["id"], row["title"], row["day"], row["workfront"],
                                 row["status"], f'{row["sheet"]}:{row["row"]}',
                                 " | ".join(v for _, v in row["details"])))
        for column, label in zip(self.data["columns"], labels):
            self.data.heading(column, text=label)

    def show_detail(self, _):
        selection = self.data.selection()
        if not selection:
            return
        row = self.display_rows[self.data.index(selection[0])]
        if self.section.get() == SECTIONS["progress"]:
            sources = row["sources"]
            suffix = f" | {len(sources) - 8} more rows" if len(sources) > 8 else ""
            self.detail.set("Site records: " + ", ".join(sources[:8]) + suffix)
        else:
            self.detail.set(" | ".join(f"{key}: {value}" for key, value in row["details"]))

    def report_panel(self, tab):
        bar = ttk.Frame(tab)
        bar.pack(fill="x", pady=(0, 10))
        for field, variable in (("Command", self.command), ("Anchor date", self.anchor),
                                ("Prepared by", self.actor), ("Reviewer", self.reviewer)):
            ttk.Label(bar, text=field).pack(side="left", padx=(0, 4))
            if field == "Command":
                widget = ttk.Combobox(bar, textvariable=variable, values=list(TITLES), state="readonly", width=7)
            else:
                widget = ttk.Entry(bar, textvariable=variable, width=17)
            widget.pack(side="left", padx=(0, 10))
        ttk.Button(bar, text="Create draft", style="Accent.TButton",
                   command=lambda: self.call(self.draft)).pack(side="right")
        ttk.Label(tab, text="MC10 End-of-Day Summary (daily)  |  MC02 Weekly Progress Report (Monday to Sunday)").pack(anchor="w")
        panel = ttk.Frame(tab)
        panel.pack(fill="both", expand=True, pady=(8, 0))
        history = ttk.Frame(panel)
        history.pack(side="left", fill="both")
        self.report_grid = self.grid(history, ("command", "date", "rev", "state"),
                                     ("Command", "Start date", "Revision", "Status"), (85, 110, 75, 95))
        self.report_grid.bind("<<TreeviewSelect>>", lambda _: self.call(self.show_report))
        self.preview = Text(panel, width=60, wrap="word", padx=12, pady=10, state="disabled")
        self.preview.pack(side="right", fill="both", expand=True, padx=(10, 0))
        footer = ttk.Frame(tab)
        footer.pack(fill="x", pady=(8, 0))
        ttk.Button(footer, text="Approve selected",
                   command=lambda: self.call(self.approve_report)).pack(side="left", padx=(0, 8))
        ttk.Button(footer, text="Issue and export Word + PDF", style="Accent.TButton",
                   command=lambda: self.call(self.export_report)).pack(side="left")

    def refresh_history(self, identifier=None):
        self.report_grid.delete(*self.report_grid.get_children())
        for row in reports.history(self.journal):
            self.report_grid.insert("", "end", iid=str(row["id"]),
                                    values=(row["command"], row["start"], row["revision"], row["state"].capitalize()))
        if identifier:
            self.report_grid.selection_set(str(identifier))
            self.show_report()

    def selected(self):
        chosen = self.report_grid.selection()
        if not chosen:
            raise ValueError("Select a report revision.")
        return int(chosen[0])

    def draft(self):
        if self.snapshot is None or str(Path(self.workbook.get()).resolve()) != self.snapshot["path"]:
            raise ValueError("Load the approved workbook before creating a report.")
        identifier = reports.create(self.journal, self.workbook.get(), self.command.get(),
                                    self.anchor.get().strip(), self.actor.get().strip(),
                                    expected_sha=self.snapshot["fingerprint"])
        self.refresh_history(identifier)
        self.feedback.set("Draft captured from the approved snapshot. Review it before approving.")

    def show_report(self):
        if not self.report_grid.selection():
            return
        state, approver, payload = reports.get(self.journal, self.selected())
        lines = [" | ".join(text) if kind == "quantity" else text for kind, text in content(payload, approver)]
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", f"{state.capitalize()}\n\n" + "\n\n".join(lines))
        self.preview.configure(state="disabled")

    def approve_report(self):
        identifier = self.selected()
        if messagebox.askyesno("Confirm approval", "Approve the selected report revision?",
                               parent=self.root):
            reports.approve(self.journal, identifier, self.reviewer.get().strip())
            self.refresh_history(identifier)
            self.feedback.set("Report approved. Word and PDF can now be issued together.")

    def export_report(self):
        identifier = self.selected()
        directory = filedialog.askdirectory(parent=self.root, mustexist=True)
        if directory:
            reports.issue(self.journal, identifier, self.actor.get().strip())
            path = reports.export(self.journal, identifier, directory)
            self.refresh_history(identifier)
            self.feedback.set(f"Issued Word and PDF to {path.name}.")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--journal", default="local/audit.sqlite3")
    parser.add_argument("--workbook", default="")
    args = parser.parse_args()
    Path(args.journal).parent.mkdir(parents=True, exist_ok=True)
    root = Tk()
    ReportingDesktop(root, Journal(args.journal), args.workbook)
    root.mainloop()


if __name__ == "__main__":
    main()
