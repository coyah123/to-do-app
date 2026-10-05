"""
To-Do App - a lightweight, offline task tracker organized by project.

Pure Python standard library (tkinter + csv + json) so it runs on an
airgapped machine with no pip installs. XLSX export uses openpyxl if it
happens to be available; CSV export always works.

Run with pythonw / the .pyw launcher to avoid a console window.
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

# Task field order used everywhere (storage, UI, export).
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
# Task editor dialog
# ---------------------------------------------------------------------------
class TaskDialog(tk.Toplevel):
    def __init__(self, master, task=None):
        super().__init__(master)
        self.title("Edit Task" if task else "New Task")
        self.resizable(False, False)
        self.result = None
        self.transient(master)
        self.grab_set()

        task = task or {}
        self.vars = {}
        row = 0
        pad = {"padx": 8, "pady": 4}

        def label(text):
            nonlocal row
            ttk.Label(self, text=text).grid(row=row, column=0, sticky="ne", **pad)

        # Title
        label("Title *")
        self.title_entry = ttk.Entry(self, width=50)
        self.title_entry.insert(0, task.get("title", ""))
        self.title_entry.grid(row=row, column=1, sticky="w", **pad)
        row += 1

        # Description
        label("Description")
        self.desc_text = tk.Text(self, width=50, height=3)
        self.desc_text.insert("1.0", task.get("description", ""))
        self.desc_text.grid(row=row, column=1, sticky="w", **pad)
        row += 1

        # Due date
        label("Due Date")
        self.due_entry = ttk.Entry(self, width=50)
        self.due_entry.insert(0, task.get("due_date", ""))
        self.due_entry.grid(row=row, column=1, sticky="w", **pad)
        ttk.Label(self, text="(e.g. 2026-10-12)", foreground="#888").grid(
            row=row, column=2, sticky="w"
        )
        row += 1

        # Jira made?
        label("Jira Ticket Made?")
        self.jira_made_var = tk.BooleanVar(value=task.get("jira_made", False))
        jira_chk = ttk.Checkbutton(
            self,
            text="Yes",
            variable=self.jira_made_var,
            command=self._toggle_jira,
        )
        jira_chk.grid(row=row, column=1, sticky="w", **pad)
        row += 1

        # Jira ref
        label("Jira Link / Number")
        self.jira_ref_entry = ttk.Entry(self, width=50)
        self.jira_ref_entry.insert(0, task.get("jira_ref", ""))
        self.jira_ref_entry.grid(row=row, column=1, sticky="w", **pad)
        row += 1

        # Blockers
        label("Blockers")
        self.blockers_text = tk.Text(self, width=50, height=2)
        self.blockers_text.insert("1.0", task.get("blockers", ""))
        self.blockers_text.grid(row=row, column=1, sticky="w", **pad)
        row += 1

        # Notes
        label("Notes")
        self.notes_text = tk.Text(self, width=50, height=4)
        self.notes_text.insert("1.0", task.get("notes", ""))
        self.notes_text.grid(row=row, column=1, sticky="w", **pad)
        row += 1

        # Completed
        label("Completed?")
        self.completed_var = tk.BooleanVar(value=bool(task.get("completed_date")))
        self._existing_completed = task.get("completed_date", "")
        ttk.Checkbutton(
            self, text="Mark complete (stamps date)", variable=self.completed_var
        ).grid(row=row, column=1, sticky="w", **pad)
        row += 1

        # Buttons
        btns = ttk.Frame(self)
        btns.grid(row=row, column=0, columnspan=3, pady=10)
        ttk.Button(btns, text="Save", command=self._save).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="left", padx=6)

        self._existing_added = task.get("added_date", "")
        self._toggle_jira()
        self.title_entry.focus_set()

    def _toggle_jira(self):
        state = "normal" if self.jira_made_var.get() else "disabled"
        self.jira_ref_entry.configure(state=state)

    def _save(self):
        title = self.title_entry.get().strip()
        if not title:
            messagebox.showwarning("Missing title", "Title is required.", parent=self)
            return

        completed = self.completed_var.get()
        if completed:
            completed_date = self._existing_completed or now_stamp()
        else:
            completed_date = ""

        self.result = {
            "title": title,
            "description": self.desc_text.get("1.0", "end").strip(),
            "added_date": self._existing_added or now_stamp(),
            "due_date": self.due_entry.get().strip(),
            "completed_date": completed_date,
            "jira_made": self.jira_made_var.get(),
            "jira_ref": self.jira_ref_entry.get().strip()
            if self.jira_made_var.get()
            else "",
            "blockers": self.blockers_text.get("1.0", "end").strip(),
            "notes": self.notes_text.get("1.0", "end").strip(),
        }
        self.destroy()


# ---------------------------------------------------------------------------
# Main application window
# ---------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self, store):
        super().__init__()
        self.store = store
        self.title("To-Do Tracker")
        self.geometry("1000x600")
        self.minsize(820, 480)

        self._build_layout()
        self._refresh_projects()

    # -- layout -----------------------------------------------------------
    def _build_layout(self):
        # Left: projects panel
        left = ttk.Frame(self, padding=8)
        left.pack(side="left", fill="y")

        ttk.Label(left, text="Projects", font=("", 11, "bold")).pack(anchor="w")
        self.project_list = tk.Listbox(left, width=24, exportselection=False)
        self.project_list.pack(fill="y", expand=True, pady=4)
        self.project_list.bind("<<ListboxSelect>>", lambda e: self._refresh_tasks())

        pbtns = ttk.Frame(left)
        pbtns.pack(fill="x")
        ttk.Button(pbtns, text="Add", command=self._add_project).pack(
            side="left", expand=True, fill="x", padx=2
        )
        ttk.Button(pbtns, text="Delete", command=self._delete_project).pack(
            side="left", expand=True, fill="x", padx=2
        )

        # Right: tasks panel
        right = ttk.Frame(self, padding=8)
        right.pack(side="left", fill="both", expand=True)

        # Toolbar
        toolbar = ttk.Frame(right)
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="New Task", command=self._add_task).pack(side="left")
        ttk.Button(toolbar, text="Edit", command=self._edit_task).pack(
            side="left", padx=4
        )
        ttk.Button(toolbar, text="Delete", command=self._delete_task).pack(side="left")

        ttk.Separator(toolbar, orient="vertical").pack(
            side="left", fill="y", padx=8
        )
        ttk.Button(toolbar, text="Export CSV", command=self._export_csv).pack(
            side="left"
        )
        self.xlsx_btn = ttk.Button(
            toolbar, text="Export XLSX", command=self._export_xlsx
        )
        self.xlsx_btn.pack(side="left", padx=4)

        # openpyxl status indicator (green/red dot)
        status = ttk.Frame(toolbar)
        status.pack(side="left", padx=6)
        self.dot = tk.Canvas(status, width=14, height=14, highlightthickness=0)
        self.dot.pack(side="left")
        color = "#2ecc71" if HAVE_OPENPYXL else "#e74c3c"
        self.dot.create_oval(2, 2, 12, 12, fill=color, outline="")
        msg = (
            "openpyxl ready"
            if HAVE_OPENPYXL
            else "openpyxl not installed (XLSX off)"
        )
        self.status_lbl = ttk.Label(status, text=msg, foreground="#555")
        self.status_lbl.pack(side="left", padx=4)

        if not HAVE_OPENPYXL:
            self.xlsx_btn.state(["disabled"])

        # Tasks table
        cols = [key for key, _ in TASK_FIELDS]
        self.tree = ttk.Treeview(right, columns=cols, show="headings", height=18)
        for key, header in TASK_FIELDS:
            self.tree.heading(key, text=header)
            width = 160 if key in ("notes", "description", "blockers") else 110
            self.tree.column(key, width=width, anchor="w")
        self.tree.pack(fill="both", expand=True, pady=8)
        self.tree.bind("<Double-1>", lambda e: self._edit_task())

        yscroll = ttk.Scrollbar(right, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscroll=yscroll.set)
        yscroll.pack(fill="x")

    # -- project actions --------------------------------------------------
    def _current_project(self):
        sel = self.project_list.curselection()
        if not sel:
            return None
        return self.project_list.get(sel[0])

    def _refresh_projects(self):
        self.project_list.delete(0, "end")
        for name in sorted(self.store.projects):
            self.project_list.insert("end", name)
        if self.project_list.size():
            self.project_list.selection_set(0)
        self._refresh_tasks()

    def _add_project(self):
        name = SimplePrompt.ask(self, "New Project", "Project name:")
        if not name:
            return
        if not self.store.add_project(name.strip()):
            messagebox.showinfo("Exists", "A project with that name already exists.")
            return
        self._refresh_projects()
        # select the new one
        names = list(self.project_list.get(0, "end"))
        if name.strip() in names:
            self.project_list.selection_clear(0, "end")
            self.project_list.selection_set(names.index(name.strip()))
            self._refresh_tasks()

    def _delete_project(self):
        name = self._current_project()
        if not name:
            return
        if messagebox.askyesno(
            "Delete project",
            f"Delete project '{name}' and all its tasks?",
        ):
            self.store.delete_project(name)
            self._refresh_projects()

    # -- task actions -----------------------------------------------------
    def _refresh_tasks(self):
        self.tree.delete(*self.tree.get_children())
        project = self._current_project()
        if not project:
            return
        for i, task in enumerate(self.store.projects[project]):
            values = []
            for key, _ in TASK_FIELDS:
                v = task.get(key, "")
                if key == "jira_made":
                    v = "Yes" if v else "No"
                # keep table cells tidy
                v = str(v).replace("\n", " ")
                values.append(v)
            self.tree.insert("", "end", iid=str(i), values=values)

    def _selected_task_index(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return int(sel[0])

    def _add_task(self):
        project = self._current_project()
        if not project:
            messagebox.showinfo("No project", "Create or select a project first.")
            return
        dlg = TaskDialog(self)
        self.wait_window(dlg)
        if dlg.result:
            self.store.add_task(project, dlg.result)
            self._refresh_tasks()

    def _edit_task(self):
        project = self._current_project()
        idx = self._selected_task_index()
        if project is None or idx is None:
            return
        task = self.store.projects[project][idx]
        dlg = TaskDialog(self, task)
        self.wait_window(dlg)
        if dlg.result:
            self.store.update_task(project, idx, dlg.result)
            self._refresh_tasks()

    def _delete_task(self):
        project = self._current_project()
        idx = self._selected_task_index()
        if project is None or idx is None:
            return
        if messagebox.askyesno("Delete task", "Delete the selected task?"):
            self.store.delete_task(project, idx)
            self._refresh_tasks()

    # -- export -----------------------------------------------------------
    def _rows_for_export(self):
        """Flatten all projects/tasks into header + rows."""
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
        # csv.writer quotes fields containing commas/newlines automatically,
        # so commas in notes are safe.
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerow(headers)
            writer.writerows(rows)
        messagebox.showinfo("Exported", f"CSV saved to:\n{path}")

    def _export_xlsx(self):
        if not HAVE_OPENPYXL:
            messagebox.showwarning(
                "openpyxl missing",
                "XLSX export needs the 'openpyxl' package.\n\n"
                "Install with: pip install openpyxl\n"
                "(Use CSV export if you can't install packages.)",
            )
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
        # reasonable column widths
        for col_idx, header in enumerate(headers, start=1):
            letter = openpyxl.utils.get_column_letter(col_idx)
            ws.column_dimensions[letter].width = max(14, min(40, len(header) + 6))
        ws.freeze_panes = "A2"
        wb.save(path)
        messagebox.showinfo("Exported", f"XLSX saved to:\n{path}")


# ---------------------------------------------------------------------------
# Tiny modal text prompt (avoids simpledialog import quirks)
# ---------------------------------------------------------------------------
class SimplePrompt(tk.Toplevel):
    def __init__(self, master, title, prompt):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self.value = None

        ttk.Label(self, text=prompt).pack(padx=12, pady=(12, 4))
        self.entry = ttk.Entry(self, width=36)
        self.entry.pack(padx=12, pady=4)
        self.entry.focus_set()
        self.entry.bind("<Return>", lambda e: self._ok())

        btns = ttk.Frame(self)
        btns.pack(pady=10)
        ttk.Button(btns, text="OK", command=self._ok).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="left", padx=6)

    def _ok(self):
        self.value = self.entry.get()
        self.destroy()

    @classmethod
    def ask(cls, master, title, prompt):
        dlg = cls(master, title, prompt)
        master.wait_window(dlg)
        return dlg.value


def main():
    store = Store(DATA_FILE)
    app = App(store)
    app.mainloop()


if __name__ == "__main__":
    main()
