"""
To-Do App - a lightweight, offline task tracker organized by project.

Pure Python standard library (tkinter + csv + json) so it runs on an
airgapped machine with no pip installs. XLSX export uses openpyxl if it
happens to be available; CSV export always works.

Everything happens inside the main window - no pop-up dialogs for adding
projects or editing tasks. The sidebar is a tree: projects at the top with
their tasks nested underneath. The only OS dialog used is the native file
picker for exports.

Runs on Windows, macOS, and Linux.
"""

import csv
import json
import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

# ---------------------------------------------------------------------------
# Optional dependency: openpyxl (for .xlsx export). Degrade gracefully.
# ---------------------------------------------------------------------------
try:
    import openpyxl  # noqa: F401
    from openpyxl.styles import Font, PatternFill

    HAVE_OPENPYXL = True
except ImportError:
    HAVE_OPENPYXL = False

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(APP_DIR, "todo-data.json")

# Task field order used everywhere (storage, export).
TASK_FIELDS = [
    ("title", "Title"),
    ("description", "Description"),
    ("added_date", "Added Date"),
    ("due_date", "Due Date"),
    ("completed_date", "Completed Date"),
    ("jira_made", "Jira Ticket Made?"),
    ("jira_ref", "Jira Link / Number"),
    ("blockers", "Blockers"),
    ("notes", "Notes"),
]


def now_stamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


# ---------------------------------------------------------------------------
# Data layer
# ---------------------------------------------------------------------------
class Store:
    """Holds all projects/tasks and persists them to a local JSON file."""

    def __init__(self, path):
        self.path = path
        # {project_name: [task_dict, ...]}
        self.projects = {}
        self.load()

    def load(self):
        if os.path.exists(self.path):
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    self.projects = json.load(f)
            except (json.JSONDecodeError, OSError):
                self.projects = {}

    def save(self):
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self.projects, f, indent=2)

    def add_project(self, name):
        if name in self.projects:
            return False
        self.projects[name] = []
        self.save()
        return True

    def delete_project(self, name):
        self.projects.pop(name, None)
        self.save()

    def add_task(self, project, task):
        self.projects[project].append(task)
        self.save()

    def update_task(self, project, index, task):
        self.projects[project][index] = task
        self.save()

    def delete_task(self, project, index):
        del self.projects[project][index]
        self.save()


# ---------------------------------------------------------------------------
# Main application window - everything lives here, no pop-up dialogs.
# ---------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self, store):
        super().__init__()
        self.store = store
        self.title("To-Do Tracker")
        self.geometry("1120x680")
        self.minsize(960, 560)

        # The project the form is currently working under.
        self.active_project = None
        # Index of the task loaded in the form (None = unsaved new task).
        self.editing_index = None
        # Map tree item id -> ("project", name) or ("task", name, index).
        self.node_meta = {}

        self._build_layout()
        self._refresh_tree()

    # -- layout -----------------------------------------------------------
    def _build_layout(self):
        # Status bar (bottom)
        self.status_var = tk.StringVar(value="Ready.")
        status_bar = ttk.Frame(self, relief="sunken")
        status_bar.pack(side="bottom", fill="x")
        ttk.Label(
            status_bar, textvariable=self.status_var, anchor="w", padding=(8, 3)
        ).pack(side="left", fill="x", expand=True)

        # --- Left: projects/tasks tree -----------------------------------
        left = ttk.Frame(self, padding=8)
        left.pack(side="left", fill="y")

        ttk.Label(left, text="Projects & Tasks", font=("", 11, "bold")).pack(
            anchor="w"
        )

        add_row = ttk.Frame(left)
        add_row.pack(fill="x", pady=4)
        self.project_entry = ttk.Entry(add_row, width=18)
        self.project_entry.pack(side="left", fill="x", expand=True)
        self.project_entry.bind("<Return>", lambda e: self._add_project())
        ttk.Button(add_row, text="Add", width=5, command=self._add_project).pack(
            side="left", padx=(4, 0)
        )

        tree_wrap = ttk.Frame(left)
        tree_wrap.pack(fill="both", expand=True, pady=4)
        self.tree = ttk.Treeview(tree_wrap, show="tree", height=22, selectmode="browse")
        self.tree.column("#0", width=250)
        self.tree.pack(side="left", fill="both", expand=True)
        yscroll = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        yscroll.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._on_tree_select())
        # Visual cue: completed tasks dimmed/struck.
        self.tree.tag_configure("done", foreground="#999")
        self.tree.tag_configure("project", font=("", 10, "bold"))

        btns = ttk.Frame(left)
        btns.pack(fill="x", pady=(4, 0))
        ttk.Button(btns, text="New Task", command=self._new_task).pack(
            side="left", expand=True, fill="x", padx=1
        )
        ttk.Button(btns, text="Del Task", command=self._delete_task).pack(
            side="left", expand=True, fill="x", padx=1
        )
        ttk.Button(left, text="Delete Project", command=self._delete_project).pack(
            fill="x", pady=(4, 0)
        )

        # --- Middle: export toolbar + spacer -----------------------------
        center = ttk.Frame(self, padding=8)
        center.pack(side="left", fill="both", expand=True)

        toolbar = ttk.Frame(center)
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="Export CSV", command=self._export_csv).pack(
            side="left"
        )
        self.xlsx_btn = ttk.Button(
            toolbar, text="Export XLSX", command=self._export_xlsx
        )
        self.xlsx_btn.pack(side="left", padx=4)
        self.dot = tk.Canvas(toolbar, width=14, height=14, highlightthickness=0)
        self.dot.pack(side="left", padx=(6, 2))
        color = "#2ecc71" if HAVE_OPENPYXL else "#e74c3c"
        self.dot.create_oval(2, 2, 12, 12, fill=color, outline="")
        ttk.Label(
            toolbar,
            text="openpyxl ready" if HAVE_OPENPYXL else "no openpyxl (XLSX off)",
            foreground="#555",
        ).pack(side="left")
        if not HAVE_OPENPYXL:
            self.xlsx_btn.state(["disabled"])

        # Context label - which project the form is editing under.
        self.context_var = tk.StringVar(value="No project selected.")
        ttk.Label(
            center, textvariable=self.context_var, font=("", 10, "italic"),
            foreground="#336", padding=(0, 10)
        ).pack(anchor="w")

        # --- Right: embedded task form -----------------------------------
        form = ttk.LabelFrame(self, text="Task Details", padding=10)
        form.pack(side="left", fill="y")
        self._build_form(form)

    def _build_form(self, form):
        pad = {"padx": 6, "pady": 3}
        r = 0

        ttk.Label(form, text="Title *").grid(row=r, column=0, sticky="ne", **pad)
        self.f_title = ttk.Entry(form, width=36)
        self.f_title.grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(form, text="Description").grid(row=r, column=0, sticky="ne", **pad)
        self.f_desc = tk.Text(form, width=36, height=3, wrap="word")
        self.f_desc.grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(form, text="Due Date").grid(row=r, column=0, sticky="ne", **pad)
        due_wrap = ttk.Frame(form)
        due_wrap.grid(row=r, column=1, sticky="w", **pad)
        self.f_due = ttk.Entry(due_wrap, width=20)
        self.f_due.pack(side="left")
        ttk.Label(due_wrap, text="e.g. 2026-10-12", foreground="#888").pack(
            side="left", padx=6
        )
        r += 1

        ttk.Label(form, text="Jira ticket made?").grid(
            row=r, column=0, sticky="ne", **pad
        )
        self.f_jira_made = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            form, text="Yes", variable=self.f_jira_made, command=self._toggle_jira
        ).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(form, text="Jira link / number").grid(
            row=r, column=0, sticky="ne", **pad
        )
        self.f_jira_ref = ttk.Entry(form, width=36)
        self.f_jira_ref.grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(form, text="Blockers").grid(row=r, column=0, sticky="ne", **pad)
        self.f_blockers = tk.Text(form, width=36, height=2, wrap="word")
        self.f_blockers.grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(form, text="Notes").grid(row=r, column=0, sticky="ne", **pad)
        self.f_notes = tk.Text(form, width=36, height=4, wrap="word")
        self.f_notes.grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(form, text="Completed?").grid(row=r, column=0, sticky="ne", **pad)
        self.f_completed = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            form, text="Mark complete (stamps date)", variable=self.f_completed
        ).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        self.f_added_lbl = ttk.Label(form, text="Added: -", foreground="#666")
        self.f_added_lbl.grid(row=r, column=1, sticky="w", padx=6)
        r += 1
        self.f_completed_lbl = ttk.Label(form, text="Completed: -", foreground="#666")
        self.f_completed_lbl.grid(row=r, column=1, sticky="w", padx=6)
        r += 1

        btnrow = ttk.Frame(form)
        btnrow.grid(row=r, column=0, columnspan=2, pady=(10, 0))
        self.save_btn = ttk.Button(btnrow, text="Save Task", command=self._save_task)
        self.save_btn.pack(side="left", padx=4)
        ttk.Button(btnrow, text="Clear / New", command=self._new_task).pack(
            side="left", padx=4
        )

        self._added_date = ""
        self._completed_date = ""
        self._toggle_jira()
        self._set_form_enabled(False)

    # -- form helpers -----------------------------------------------------
    def _toggle_jira(self):
        state = "normal" if self.f_jira_made.get() else "disabled"
        self.f_jira_ref.configure(state=state)

    def _set_form_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        for w in (self.f_title, self.f_due, self.f_desc, self.f_blockers, self.f_notes):
            w.configure(state=state)
        if enabled:
            self._toggle_jira()
        else:
            self.f_jira_ref.configure(state="disabled")
        self.save_btn.state(["!disabled"] if enabled else ["disabled"])

    def _set_text(self, widget, value):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", value)

    def _clear_form(self):
        self._set_form_enabled(True)
        self.f_title.delete(0, "end")
        self.f_due.delete(0, "end")
        self.f_jira_ref.configure(state="normal")
        self.f_jira_ref.delete(0, "end")
        self._set_text(self.f_desc, "")
        self._set_text(self.f_blockers, "")
        self._set_text(self.f_notes, "")
        self.f_jira_made.set(False)
        self.f_completed.set(False)
        self._added_date = ""
        self._completed_date = ""
        self.f_added_lbl.configure(text="Added: (on save)")
        self.f_completed_lbl.configure(text="Completed: -")
        self._toggle_jira()

    def _load_task_into_form(self, task):
        self._set_form_enabled(True)
        self.f_title.delete(0, "end")
        self.f_title.insert(0, task.get("title", ""))
        self.f_due.delete(0, "end")
        self.f_due.insert(0, task.get("due_date", ""))
        self._set_text(self.f_desc, task.get("description", ""))
        self._set_text(self.f_blockers, task.get("blockers", ""))
        self._set_text(self.f_notes, task.get("notes", ""))
        self.f_jira_made.set(task.get("jira_made", False))
        self._toggle_jira()
        self.f_jira_ref.configure(state="normal")
        self.f_jira_ref.delete(0, "end")
        self.f_jira_ref.insert(0, task.get("jira_ref", ""))
        self._toggle_jira()
        self._added_date = task.get("added_date", "")
        self._completed_date = task.get("completed_date", "")
        self.f_completed.set(bool(self._completed_date))
        self.f_added_lbl.configure(text=f"Added: {self._added_date or '-'}")
        self.f_completed_lbl.configure(text=f"Completed: {self._completed_date or '-'}")

    # -- tree -------------------------------------------------------------
    def _refresh_tree(self, select_item=None):
        self.tree.delete(*self.tree.get_children())
        self.node_meta = {}
        for project in sorted(self.store.projects):
            tasks = self.store.projects[project]
            pid = self.tree.insert(
                "", "end", text=f"\U0001F4C1 {project}  ({len(tasks)})",
                open=True, tags=("project",)
            )
            self.node_meta[pid] = ("project", project)
            for i, task in enumerate(tasks):
                done = bool(task.get("completed_date"))
                mark = "✓ " if done else "• "
                label = mark + (task.get("title") or "(untitled)")
                tid = self.tree.insert(
                    pid, "end", text=label, tags=("done",) if done else ()
                )
                self.node_meta[tid] = ("task", project, i)
        if select_item and select_item in self.node_meta:
            self.tree.selection_set(select_item)
            self.tree.see(select_item)

    def _find_task_node(self, project, index):
        for item, meta in self.node_meta.items():
            if meta[0] == "task" and meta[1] == project and meta[2] == index:
                return item
        return None

    def _find_project_node(self, project):
        for item, meta in self.node_meta.items():
            if meta[0] == "project" and meta[1] == project:
                return item
        return None

    def _on_tree_select(self):
        sel = self.tree.selection()
        if not sel:
            return
        meta = self.node_meta.get(sel[0])
        if not meta:
            return
        if meta[0] == "project":
            self.active_project = meta[1]
            self.editing_index = None
            self._clear_form()
            self.context_var.set(
                f"Project: {self.active_project}  -  click 'New Task' to add one."
            )
        else:  # task
            _, project, index = meta
            self.active_project = project
            self.editing_index = index
            self._load_task_into_form(self.store.projects[project][index])
            self.context_var.set(f"Editing task under project: {project}")

    # -- project actions --------------------------------------------------
    def _add_project(self):
        name = self.project_entry.get().strip()
        if not name:
            return
        if not self.store.add_project(name):
            self._status(f"Project '{name}' already exists.")
            return
        self.project_entry.delete(0, "end")
        self._refresh_tree()
        pid = self._find_project_node(name)
        if pid:
            self.tree.selection_set(pid)
        self._status(f"Added project '{name}'. Now click 'New Task' to add tasks.")

    def _delete_project(self):
        if not self.active_project:
            self._status("Select a project first.")
            return
        name = self.active_project
        if messagebox.askyesno(
            "Delete project", f"Delete project '{name}' and all its tasks?"
        ):
            self.store.delete_project(name)
            self.active_project = None
            self.editing_index = None
            self._set_form_enabled(False)
            self.context_var.set("No project selected.")
            self._refresh_tree()
            self._status(f"Deleted project '{name}'.")

    # -- task actions -----------------------------------------------------
    def _new_task(self):
        if not self.active_project:
            self._status("Select a project in the list first, then click New Task.")
            return
        self.editing_index = None
        self._clear_form()
        self.context_var.set(f"New task under project: {self.active_project}")
        self.f_title.focus_set()
        self._status(f"Entering new task under '{self.active_project}'. Fill in + Save.")

    def _save_task(self):
        if not self.active_project:
            self._status("Select a project first.")
            return
        project = self.active_project
        title = self.f_title.get().strip()
        if not title:
            self._status("Title is required.")
            self.f_title.focus_set()
            return

        completed = self.f_completed.get()
        completed_date = (self._completed_date or now_stamp()) if completed else ""

        task = {
            "title": title,
            "description": self.f_desc.get("1.0", "end").strip(),
            "added_date": self._added_date or now_stamp(),
            "due_date": self.f_due.get().strip(),
            "completed_date": completed_date,
            "jira_made": self.f_jira_made.get(),
            "jira_ref": self.f_jira_ref.get().strip()
            if self.f_jira_made.get()
            else "",
            "blockers": self.f_blockers.get("1.0", "end").strip(),
            "notes": self.f_notes.get("1.0", "end").strip(),
        }

        if self.editing_index is None:
            self.store.add_task(project, task)
            self.editing_index = len(self.store.projects[project]) - 1
            self._status(f"Added task '{title}' to project '{project}'.")
        else:
            self.store.update_task(project, self.editing_index, task)
            self._status(f"Saved task '{title}'.")

        self._refresh_tree(
            select_item=None
        )
        node = self._find_task_node(project, self.editing_index)
        if node:
            self.tree.selection_set(node)
            self.tree.see(node)

    def _delete_task(self):
        if self.active_project is None or self.editing_index is None:
            self._status("Select a task to delete.")
            return
        project = self.active_project
        idx = self.editing_index
        title = self.store.projects[project][idx].get("title", "")
        if messagebox.askyesno("Delete task", f"Delete task '{title}'?"):
            self.store.delete_task(project, idx)
            self.editing_index = None
            self._clear_form()
            self._refresh_tree()
            pid = self._find_project_node(project)
            if pid:
                self.tree.selection_set(pid)
            self._status(f"Deleted task '{title}'.")

    # -- export -----------------------------------------------------------
    def _rows_for_export(self):
        headers = ["Project"] + [label for _, label in TASK_FIELDS]
        rows = []
        for project in sorted(self.store.projects):
            for task in self.store.projects[project]:
                row = [project]
                for key, _ in TASK_FIELDS:
                    v = task.get(key, "")
                    if key == "jira_made":
                        v = "Yes" if v else "No"
                    row.append(v)
                rows.append(row)
        return headers, rows

    def _export_csv(self):
        headers, rows = self._rows_for_export()
        path = filedialog.asksaveasfilename(
            title="Export CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="todo-export.csv",
        )
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerow(headers)
            writer.writerows(rows)
        self._status(f"CSV exported to {path}")

    def _export_xlsx(self):
        if not HAVE_OPENPYXL:
            self._status("openpyxl not installed - XLSX export unavailable. Use CSV.")
            return
        headers, rows = self._rows_for_export()
        path = filedialog.asksaveasfilename(
            title="Export XLSX",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="todo-export.xlsx",
        )
        if not path:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Tasks"
        ws.append(headers)
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="4472C4")
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
        for row in rows:
            ws.append(row)
        for col_idx, header in enumerate(headers, start=1):
            letter = openpyxl.utils.get_column_letter(col_idx)
            ws.column_dimensions[letter].width = max(14, min(40, len(header) + 6))
        ws.freeze_panes = "A2"
        wb.save(path)
        self._status(f"XLSX exported to {path}")

    # -- misc -------------------------------------------------------------
    def _status(self, msg):
        self.status_var.set(msg)


def main():
    store = Store(DATA_FILE)
    app = App(store)
    app.mainloop()


if __name__ == "__main__":
    main()
