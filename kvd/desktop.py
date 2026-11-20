"""Tkinter workspace for local P2 intake and proposal review."""
import argparse
import sqlite3
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, font, messagebox, ttk

from openpyxl.utils.exceptions import InvalidFileException

from .excel_commit import commit, recover
from .intake import COMMANDS, FIELDS, parse
from .journal import Journal
from .proposals import approve, preview


BG = "#edf2f6"
NAVY = "#101e32"
INK = "#12243a"
MUTED = "#526679"
TEAL = "#087d71"


class Desktop:
    def __init__(self, root, journal, workbook=""):
        self.root = root
        self.journal = journal
        self.drafts = []
        self.decisions = []
        self.vars = {}
        self.selected = None
        self.token = None
        self.payload = None
        self.command = tk.StringVar(value="MC01")
        self.workbook = tk.StringVar(value=workbook)
        self.operator = tk.StringVar()
        self.reviewer = tk.StringVar()
        self.attested = tk.BooleanVar()
        self.status = tk.StringVar(value="Select an expanded synthetic workbook to begin.")
        self.summary = tk.StringVar(value="Build a preview to see the proposed cells.")
        self.root.title("Construction Management | Site Intake")
        self.root.geometry("1330x890")
        self.root.minsize(1040, 790)
        self.root.configure(bg=BG)
        families = {family.casefold() for family in font.families(root)}
        self.ui_font = "Segoe UI" if "segoe ui" in families else "Helvetica"
        self.mono_font = "Consolas" if "consolas" in families else "Courier"
        self.style()
        self.build()
        self.workbook.trace_add("write", self.invalidate)
        self.operator.trace_add("write", self.invalidate)
        self.command.trace_add("write", self.invalidate_paste)

    def style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(".", font=(self.ui_font, 10), background=BG, foreground=INK)
        style.configure("TButton", padding=(11, 7))
        style.configure("Accent.TButton", background=TEAL, foreground="white")
        style.map("Accent.TButton", background=[("disabled", "#718887"), ("active", "#09685f")])
        style.configure("TNotebook.Tab", padding=(14, 8))
        style.configure("Treeview", rowheight=26, fieldbackground="white", background="white")
        style.configure("Treeview.Heading", font=(self.ui_font, 10, "bold"))
        style.configure("Muted.TLabel", foreground=MUTED)

    def build(self):
        top = tk.Frame(self.root, bg=NAVY, padx=25, pady=15)
        top.pack(fill="x")
        tk.Label(top, text="Site intake & review", bg=NAVY, fg="white",
                 font=(self.ui_font, 24, "bold")).pack(anchor="w")
        tk.Label(top, text="Synthetic workspace  |  Approved updates create a new workbook revision",
                 bg=NAVY, fg="#b7ccde", font=(self.ui_font, 11)).pack(anchor="w")
        main = ttk.Frame(self.root, padding=(20, 12))
        main.pack(fill="both", expand=True)
        config = ttk.Frame(main)
        config.pack(fill="x", pady=(0, 10))
        config.columnconfigure(1, weight=1)
        ttk.Label(config, text="Workbook").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(config, textvariable=self.workbook).grid(row=0, column=1, sticky="ew")
        ttk.Button(config, text="Browse...", command=self.choose).grid(row=0, column=2, padx=8)
        ttk.Label(config, text="Operator").grid(row=0, column=3, padx=(8, 5))
        ttk.Entry(config, textvariable=self.operator, width=22).grid(row=0, column=4)
        ttk.Label(main, textvariable=self.status, wraplength=990, style="Muted.TLabel").pack(
            side="bottom", anchor="w", pady=(10, 0))
        self.tabs = ttk.Notebook(main)
        self.tabs.pack(fill="both", expand=True)
        self.intake_tab()
        self.review_tab()
        self.proposal_tab()

    def intake_tab(self):
        self.intake = ttk.Frame(self.tabs, padding=14)
        self.tabs.add(self.intake, text="1  WhatsApp intake")
        bar = ttk.Frame(self.intake)
        bar.pack(fill="x", pady=(0, 12))
        ttk.Label(bar, text="Default command").pack(side="left", padx=(0, 8))
        ttk.Combobox(bar, textvariable=self.command, values=list(COMMANDS), state="readonly",
                     width=9).pack(side="left")
        ttk.Label(bar, text="Start with Workfront: alias or Zon kerja: alias. Use one field per line.",
                  style="Muted.TLabel").pack(side="left", padx=12)
        ttk.Button(bar, text="Parse into drafts", style="Accent.TButton",
                   command=lambda: self.run(self.parse_paste)).pack(side="right")
        ttk.Label(self.intake, text="Pasted WhatsApp text").pack(anchor="w")
        pane = ttk.Frame(self.intake)
        pane.pack(fill="both", expand=True, pady=(5, 0))
        self.paste = tk.Text(pane, wrap="word", bg="white", fg=INK, relief="flat",
                             padx=12, pady=10, font=(self.mono_font, 11), undo=True)
        scroll = ttk.Scrollbar(pane, orient="vertical", command=self.paste.yview)
        self.paste.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.paste.pack(side="left", fill="both", expand=True)
        self.paste.bind("<<Modified>>", self.paste_changed)

    def paste_changed(self, _):
        if self.paste.edit_modified():
            self.paste.edit_modified(False)
            if self.paste.get("1.0", "end-1c") != getattr(self, "parsed_paste", None):
                self.invalidate_paste()

    def invalidate_paste(self, *_):
        self.invalidate()
        if self.drafts:
            self.drafts = []
            self.decisions = []
            self.selected = None
            self.list.delete(*self.list.get_children())
            for widget in self.form.winfo_children():
                widget.destroy()
            self.source.configure(state="normal")
            self.source.delete("1.0", "end")
            self.source.configure(state="disabled")
            self.status.set("Pasted text changed. Parse it again to review drafts.")

    def review_tab(self):
        self.review = ttk.Frame(self.tabs, padding=14)
        self.tabs.add(self.review, text="2  Review drafts")
        top = ttk.Frame(self.review)
        top.pack(fill="x", pady=(0, 8))
        ttk.Label(top, text="Select a draft to correct its fields or record a duplicate decision.").pack(side="left")
        ttk.Button(top, text="Build Excel preview", style="Accent.TButton",
                   command=lambda: self.run(self.build_preview)).pack(side="right")
        body = ttk.Frame(self.review)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)
        left = ttk.Frame(body)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.list = ttk.Treeview(left, columns=("command", "workfront"), show="headings",
                                 selectmode="browse", height=8)
        for column, title, width in (("command", "Command", 85), ("workfront", "Workfront", 200)):
            self.list.heading(column, text=title)
            self.list.column(column, width=width)
        self.list.pack(fill="x")
        self.list.bind("<<TreeviewSelect>>", self.pick)
        ttk.Label(left, text="Original source excerpt").pack(anchor="w", pady=(12, 4))
        self.source = tk.Text(left, width=35, wrap="word", state="disabled", bg="white", fg=INK,
                              padx=8, pady=8, relief="flat")
        self.source.pack(fill="both", expand=True)
        self.form = ttk.LabelFrame(body, text="Editable draft and review decision", padding=10)
        self.form.grid(row=0, column=1, sticky="nsew")
        self.form.columnconfigure(1, weight=1)
        self.form.columnconfigure(3, weight=1)

    def proposal_tab(self):
        self.proposal = ttk.Frame(self.tabs, padding=14)
        self.tabs.add(self.proposal, text="3  Excel proposal")
        ttk.Label(self.proposal, textvariable=self.summary, wraplength=1160).pack(anchor="w", pady=(0, 8))
        panel = ttk.Frame(self.proposal)
        panel.pack(fill="both", expand=True)
        cols = ("sheet", "cell", "before", "after")
        self.cells = ttk.Treeview(panel, columns=cols, show="headings")
        for column, label, width in zip(cols, ("Worksheet", "Cell", "Current value", "Proposed value"),
                                        (220, 65, 230, 460)):
            self.cells.heading(column, text=label)
            self.cells.column(column, width=width)
        scrollbar = ttk.Scrollbar(panel, orient="vertical", command=self.cells.yview)
        self.cells.configure(yscrollcommand=scrollbar.set)
        self.cells.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        foot = ttk.Frame(self.proposal)
        foot.pack(fill="x", pady=(10, 0))
        ttk.Label(foot, text="Reviewer").pack(side="left")
        ttk.Entry(foot, textvariable=self.reviewer, width=23).pack(side="left", padx=8)
        ttk.Checkbutton(foot, text="I reviewed the separate quantity evidence", variable=self.attested).pack(side="left")
        self.approve_button = ttk.Button(foot, text="Approve proposal", style="Accent.TButton",
                                         command=self.confirm_approval, state="disabled")
        self.approve_button.pack(side="right")
        self.commit_button = ttk.Button(foot, text="Save approved revision", state="disabled",
                                        command=lambda: self.run(self.save_revision))
        self.commit_button.pack(side="right", padx=8)
        ttk.Button(foot, text="Recover", command=lambda: self.run(self.recover_revisions)).pack(side="right")

    def choose(self):
        path = filedialog.askopenfilename(filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.workbook.set(path)

    def invalidate(self, *_):
        if self.token is not None:
            self.token = None
            self.payload = None
            self.summary.set("Draft changed. Rebuild the Excel preview before approval.")
            self.cells.delete(*self.cells.get_children())
            self.approve_button.configure(state="disabled")
            self.commit_button.configure(state="disabled")

    def run(self, operation):
        try:
            operation()
        except (ValueError, OSError, sqlite3.Error, InvalidFileException) as error:
            self.status.set(str(error))
            messagebox.showerror("Review required", str(error), parent=self.root)

    def parse_paste(self):
        text = self.paste.get("1.0", "end-1c")
        incoming = parse(text, self.command.get())
        self.parsed_paste = text
        self.invalidate()
        self.selected = None
        self.drafts = incoming
        self.decisions = [{} for _ in incoming]
        self.list.delete(*self.list.get_children())
        for index, draft in enumerate(incoming):
            self.list.insert("", "end", iid=str(index),
                             values=(draft.command, draft.fields.get("workfront", "")))
        self.list.selection_set("0")
        self.tabs.select(self.review)
        self.status.set(f"{len(incoming)} draft records ready for review.")

    def save_fields(self):
        if self.selected is None:
            return
        draft = self.drafts[self.selected]
        draft.fields = {key: variable.get().strip() for key, variable in self.vars.items()
                        if key in FIELDS[draft.command] and variable.get().strip()}
        self.decisions[self.selected] = {key: variable.get().strip() for key, variable in self.vars.items()
                                         if key not in FIELDS[draft.command]}
        self.list.item(str(self.selected), values=(draft.command, draft.fields.get("workfront", "")))

    def pick(self, _=None):
        choice = self.list.selection()
        if not choice or self.selected == int(choice[0]):
            return
        self.save_fields()
        self.selected = int(choice[0])
        draft = self.drafts[self.selected]
        for widget in self.form.winfo_children():
            widget.destroy()
        self.source.configure(state="normal")
        self.source.delete("1.0", "end")
        self.source.insert("1.0", draft.source)
        self.source.configure(state="disabled")
        self.vars = {}
        ttk.Label(self.form, text=f"{draft.command}  {COMMANDS[draft.command]}",
                  font=(self.ui_font, 12, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 6))
        for index, key in enumerate(FIELDS[draft.command]):
            row, pair = divmod(index, 2)
            self.field(key, draft.fields.get(key, ""), row + 1, pair)
        row = (len(FIELDS[draft.command]) + 1) // 2 + 2
        ttk.Separator(self.form).grid(row=row - 1, column=0, columnspan=4, sticky="ew", pady=7)
        options = self.decisions[self.selected]
        self.field("action", options.get("action", "add"), row, 0,
                   ("add", "keep distinct", "suppress", "correct existing"))
        self.field("reason", options.get("reason", ""), row, 1)
        if draft.command in {"MC01", "MC14"}:
            self.field("record_id", options.get("record_id", ""), row + 1, 0)
            self.field("verified_qty", options.get("verified_qty", ""), row + 1, 1)
            self.field("verification_evidence", options.get("verification_evidence", ""), row + 2, 0)

    def field(self, key, value, row, pair, options=None):
        column = pair * 2
        label = key.replace("_", " ").capitalize()
        ttk.Label(self.form, text=label).grid(row=row, column=column, sticky="w", padx=(4, 6), pady=4)
        variable = tk.StringVar(value=value)
        self.vars[key] = variable
        variable.trace_add("write", self.invalidate)
        if options:
            widget = ttk.Combobox(self.form, textvariable=variable, values=options, state="readonly", width=18)
        else:
            widget = ttk.Entry(self.form, textvariable=variable, width=18)
        widget.grid(row=row, column=column + 1, sticky="ew", padx=(0, 9), pady=4)

    def build_preview(self):
        self.save_fields()
        self.invalidate()
        self.token, self.payload = preview(self.journal, self.workbook.get(), self.drafts,
                                           self.decisions, self.operator.get().strip())
        for change in self.payload["changes"]:
            self.cells.insert("", "end", values=(change["sheet"], change["cell"],
                                                  change["before"], change["after"]))
        flags = sum(bool(record.get("duplicates")) for record in self.payload["records"])
        self.summary.set(f"{len(self.payload['changes'])} proposed cell changes  |  {flags} duplicate overrides"
                         f"  |  Workbook fingerprint {self.payload['workbook_sha'][:12]}")
        self.approve_button.configure(state="normal")
        self.tabs.select(self.proposal)
        self.status.set("Review the proposed cells and confirm the reviewer’s decision.")

    def confirm_approval(self):
        if self.token and messagebox.askyesno("Approve Excel proposal",
                                             "Record approval for this exact preview?",
                                             parent=self.root):
            self.run(self.approve_current)

    def approve_current(self):
        approve(self.journal, self.token, self.reviewer.get().strip(), self.attested.get())
        self.approve_button.configure(state="disabled")
        if self.payload["changes"]:
            self.commit_button.configure(state="normal")
        self.status.set(f"Proposal {self.token[:12]} approved. Save it as a new workbook revision.")

    def save_revision(self):
        output = filedialog.asksaveasfilename(
            parent=self.root, defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if output:
            self.write_revision(output)

    def write_revision(self, output):
        result = commit(self.journal, self.token, output, self.operator.get().strip())
        self.commit_button.configure(state="disabled")
        self.status.set(f"Saved {Path(result['output']).name} | Workbook fingerprint {result['output_sha'][:12]}")

    def recover_revisions(self):
        result = recover(self.journal, self.operator.get().strip())
        if not result:
            self.status.set("No interrupted proposal commits found.")
            return
        self.status.set("  |  ".join(f"{token[:12]}: {state}" for token, state in result))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--journal", default="local/audit.sqlite3")
    parser.add_argument("--workbook", default="")
    args = parser.parse_args()
    Path(args.journal).parent.mkdir(parents=True, exist_ok=True)
    root = tk.Tk()
    Desktop(root, Journal(args.journal), args.workbook)
    root.mainloop()


if __name__ == "__main__":
    main()
