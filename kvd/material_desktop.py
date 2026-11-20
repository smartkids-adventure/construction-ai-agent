"""Read-only material forecast views inside the Engineering Dashboard."""
from datetime import date
from tkinter import StringVar, ttk

from . import materials
from .material_forecast import forecast


COLOURS = {"RED": "#fce1df", "AMBER": "#fff1cf", "GREEN": "#e1f3e8", "UNKNOWN": "#e8edf4"}
VIEWS = ("14-day projection", "Deliveries", "Consumption variance", "Urgent procurement", "Requisition history")


def label(value):
    return "—" if value is None or value == "" else str(value)


class MaterialDashboard:
    def __init__(self, frame, journal, call):
        self.journal, self.call = journal, call
        self.rows = []
        self.day = StringVar(value=date.today().isoformat())
        self.workfront = StringVar()
        self.material = StringVar()
        self.status = StringVar()
        self.view = StringVar(value=VIEWS[0])
        self.summary = StringVar(value="Choose a forecast start date and refresh.")
        self.detail = StringVar()
        filters = ttk.Frame(frame)
        filters.pack(fill="x")
        for column, (title, value) in enumerate((("Start date (YYYY-MM-DD)", self.day), ("Workfront ID", self.workfront),
                                                 ("Material ID", self.material), ("Status", self.status))):
            ttk.Label(filters, text=title).grid(row=0, column=column, sticky="w")
            if value == self.status:
                control = ttk.Combobox(filters, textvariable=value, values=("", *COLOURS), state="readonly", width=14)
            else:
                control = ttk.Entry(filters, textvariable=value, width=21)
            control.grid(row=1, column=column, sticky="ew", padx=(0, 8))
            filters.columnconfigure(column, weight=1)
        ttk.Button(filters, text="Refresh forecast", style="Accent.TButton",
                   command=lambda: call(self.refresh)).grid(row=1, column=4, padx=5)
        ttk.Label(frame, textvariable=self.summary, wraplength=950).pack(anchor="w", pady=(10, 5))
        self.top = self.tree(frame, (
            ("Material", 127), ("Workfront", 145), ("Location", 140), ("Unit", 45),
            ("Physical", 68), ("Reserved", 68), ("Available", 68), ("Demand", 68),
            ("Confirmed", 75), ("Shortage", 72), ("Stockout", 91), ("Status", 80),
        ), height=7, expands=False)
        self.top.bind("<<TreeviewSelect>>", lambda _: call(self.refresh_detail))
        for status, colour in COLOURS.items():
            self.top.tag_configure(status, background=colour)
        bar = ttk.Frame(frame)
        bar.pack(fill="x", pady=(10, 5))
        ttk.Label(bar, text="Details").pack(side="left", padx=(0, 7))
        selector = ttk.Combobox(bar, textvariable=self.view, state="readonly", width=25, values=VIEWS)
        selector.pack(side="left")
        selector.bind("<<ComboboxSelected>>", lambda _: call(self.refresh_detail))
        ttk.Label(frame, textvariable=self.detail, wraplength=950).pack(side="bottom", anchor="w", pady=(5, 0))
        self.bottom = self.tree(frame, (), height=8)

    @staticmethod
    def tree(parent, headings, height, expands=True):
        pane = ttk.Frame(parent)
        pane.pack(fill="both", expand=expands)
        tree = ttk.Treeview(pane, show="headings", height=height, selectmode="browse")
        scrollbar = ttk.Scrollbar(pane, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(pane, orient="horizontal", command=tree.xview)
        pane.rowconfigure(0, weight=1)
        pane.columnconfigure(0, weight=1)
        tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        MaterialDashboard.headings(tree, headings)
        return tree

    @staticmethod
    def headings(tree, headings):
        keys = tuple(f"field{i}" for i in range(len(headings)))
        tree.configure(columns=keys)
        for key, (name, width) in zip(keys, headings):
            tree.heading(key, text=name)
            tree.column(key, width=width, minwidth=40)

    def refresh(self):
        self.rows = materials.read(self.journal, forecast, start=self.day.get().strip(),
                                   workfront=self.workfront.get().strip(), material_id=self.material.get().strip(),
                                   status=self.status.get())
        self.top.delete(*self.top.get_children())
        for index, row in enumerate(self.rows):
            values = (row["material"], row["workfront"], row["location"], row["unit"],
                      row["stock"]["physical"], row["stock"]["reserved"], row["stock"]["available"],
                      row["demand"], row["confirmed"], row["shortage"], row["stockout"], row["status"])
            self.top.insert("", "end", iid=str(index), values=tuple(map(label, values)), tags=(row["status"],))
        self.summary.set(f"{len(self.rows)} material and workfront rows | 14 calendar days | "
                         "Location stock is shared across its workfront rows.")
        if self.rows:
            self.top.selection_set("0")
        self.refresh_detail()

    def fill(self, headings, records):
        self.bottom.delete(*self.bottom.get_children())
        self.headings(self.bottom, headings)
        for values in records:
            self.bottom.insert("", "end", values=tuple(map(label, values)))

    def refresh_detail(self):
        selected = self.top.selection()
        row = self.rows[int(selected[0])] if selected else None
        self.detail.set("")
        view = self.view.get()
        if view == "Urgent procurement":
            records = sorted((line for line in self.rows if line["status"] in {"RED", "AMBER", "UNKNOWN"}),
                             key=lambda line: ({"RED": 0, "UNKNOWN": 1, "AMBER": 2}[line["status"]],
                                               line["requisition_by"] or "9999"))
            self.fill((("Material", 120), ("Workfront", 145), ("Status", 90), ("Shortage", 85),
                       ("Stockout", 100), ("Requisition by", 110), ("Action", 300)),
                      ((line["material"], line["workfront"], line["status"], line["shortage"],
                        line["stockout"], line["requisition_by"],
                        "Review draft requisition" if line["status"] == "RED"
                        else "Check safety stock or overdue PO" if line["status"] == "AMBER"
                        else "Resolve missing or unverified evidence") for line in records))
            return
        if view == "Requisition history":
            records = materials.read(self.journal, materials.list_requisitions)
            selected_workfront, selected_material = self.workfront.get().strip(), self.material.get().strip()
            self.fill((("ID", 75), ("Revision", 75), ("Status", 100), ("Workfront", 145), ("Material", 130),
                       ("Quantity", 90), ("Required delivery", 125), ("Reviewer", 190)),
                      ((rec["id"], rec["revision"], rec["state"], rec["workfront"], rec["material"],
                        rec["quantity"], rec["required"], rec["reviewer"]) for rec in records
                       if (not selected_workfront or rec["workfront"] == selected_workfront)
                       and (not selected_material or rec["material"] == selected_material)))
            return
        if not row:
            self.fill((), ())
            return
        self.detail.set(" | ".join(row["reasons"] or [f'Safety stock {row["safety_stock"]} {row["unit"]}',
                                                     f'Quarantined {row["stock"]["quarantine"]} {row["unit"]}',
                                                     f'Requisition by {label(row["requisition_by"])}']))
        if view == "Deliveries":
            self.fill((("PO", 110), ("Requisition", 90), ("Due", 100), ("Confirmed ETA", 120),
                       ("Accepted", 80), ("Rejected", 80), ("Quarantined", 95), ("Remaining", 85),
                       ("Actual arrivals", 145), ("Status", 110)),
                      ((item["po"], item["requisition"], item["due"], item["eta"] if item["confirmed"] else None,
                        item["accepted"], item["rejected"], item["quarantined"], item["remaining"],
                        ", ".join(item["actual_dates"]), "Overdue" if item["overdue"]
                        else "Arrived late" if item["last_arrival_late"] else "Confirmed" if item["confirmed"]
                        and item["remaining"] != "0" else "Complete" if item["remaining"] == "0" else "Unconfirmed")
                       for item in row["deliveries"]))
        elif view == "Consumption variance":
            values = row["variance"]
            self.fill((("Basis", 210), ("Planned / design", 155), ("Actual / issued", 155),
                       ("Variance", 130), ("Flag", 220)),
                      (("14-day consumption", values["planned"], values["actual"], values["variance"],
                        "Potential wastage" if values["potential_wastage"] else " — "),
                       ("Design vs issued to date", values["design"], values["issued"], None,
                        "Above design allowance threshold" if values["design_over_issue"] else " — ")))
        else:
            self.fill((("Date", 118), ("Available to workfront", 160), ("Planned demand", 145),
                       ("Confirmed incoming", 150), ("Workfront deficit", 140),
                       ("Location projected", 148), ("Unallocated net", 144)),
                      ((record["date"], record["available"], record["demand"], record["incoming"],
                        record["deficit"], record["projected"], record["unallocated"]) for record in row["daily"]))
