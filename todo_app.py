"""
To-Do App - a lightweight, offline task tracker organized by project.

Pure Python standard library (tkinter + csv + json) so it runs on an
airgapped machine with no pip installs. XLSX export uses openpyxl if it
happens to be available; CSV export always works.

Everything happens inside the main window - no pop-up dialogs for adding
projects or editing tasks. Three resizable panes:
  - left:   project/task tree
  - center: task table for the selected project (group by epic / sprint /
            status, sort by clicking a column header)
  - right:  task details form
The only OS dialog used is the native file picker for exports.

Runs on Windows, macOS, and Linux.
"""

import csv
import json
import os
import re
import tkinter as tk
import webbrowser
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
# Per-machine window state (last view, window positions, pin).
SETTINGS_FILE = os.path.join(APP_DIR, "todo-settings.json")

# Status name -> (color, tag). Order here is the workflow order used for
# sorting and grouping.
STATUSES = {
    "Not started": ("#7f8c8d", "st_not_started"),
    "In Progress": ("#8e44ad", "st_in_progress"),
    "Blocked": ("#c0392b", "st_blocked"),
    "Waiting for approval": ("#e67e22", "st_waiting"),
    "Done": ("#27ae60", "st_done"),
}
STATUS_ORDER = list(STATUSES)
DEFAULT_STATUS = "Not started"
DONE = "Done"

# Task field order used everywhere (storage, export).
# "added_date" is the created date (key kept for backwards compatibility).
TASK_FIELDS = [
    ("title", "Title"),
    ("status", "Status"),
    ("group", "Subgroup"),
    ("epic", "Epic"),
    ("sprint", "Sprint"),
    ("description", "Description"),
    ("added_date", "Created Date"),
    ("due_date", "Due Date"),
    ("completed_date", "Completed Date"),
    ("jira_made", "Jira Ticket Made?"),
    ("jira_ref", "Jira Link / Number"),
    ("blockers", "Blockers"),
    ("notes", "Notes"),
    ("links", "Links"),  # list of {"title", "description", "url"}
]

# Center table columns: (key, heading, width, stretch). The title lives in
# the tree column (#0) so group headers and task titles share it.
TABLE_COLS = [
    ("status", "Status", 140, False),
    ("epic", "Epic", 90, False),
    ("sprint", "Sprint", 80, False),
    ("due_date", "Due", 80, False),
    ("added_date", "Created", 110, False),
]

GROUP_OPTIONS = ["None", "Subgroup", "Epic", "Sprint", "Status"]

# Named per-project lists that tasks are assigned to. "group" is a plain
# sub-folder of a project (e.g. an app), shown nested in the sidebar trees.
GROUP_KINDS = ["group", "epic", "sprint"]
# What each kind is called on screen ("group" shows as "Subgroup").
KIND_LABEL = {"group": "Subgroup", "epic": "Epic", "sprint": "Sprint"}


def kind_of(label):
    """Map an on-screen label ("Subgroup", "Epic", "Status"...) to its task key."""
    return {v: k for k, v in KIND_LABEL.items()}.get(label, label.lower())

# Views, selectable from the View menu. Buddy is the default.
VIEWS = ["Buddy", "Manager", "Visual Planner"]
DEFAULT_VIEW = "Buddy"
MANAGER_MIN = (1000, 520)
VIEW_TITLES = {
    "Buddy": "To-Do Buddy",
    "Manager": "To-Do Tracker",
    "Visual Planner": "To-Do Visual Planner",
}
BUDDY_MIN = (230, 180)
BUDDY_SIZE = (300, 440)

# Buddy shows the most active work first.
BUDDY_ORDER = ["In Progress", "Blocked", "Waiting for approval", "Not started", "Done"]
BUDDY_BG = "#fff8c5"  # sticky-note yellow
ALL_PROJECTS = "All projects"
ALL_SPRINTS = "All sprints"
NO_SPRINT = "(no sprint)"

# Visual Planner (Trello-style board).
PLANNER_COLUMNS = ["Status", "Subgroup", "Epic", "Sprint", "Project"]
PLANNER_ORDER = ["Manual", "Due date", "Created", "Title"]
ANY = "All"
NONE_LABEL = "(none)"
BOARD_BG = "#e4e9f0"
LIST_BG = "#d5dce6"
CARD_BG = "#ffffff"
CARD_BORDER = "#c4ccd6"
ACCENT = "#2b7de9"  # selected card, drop target
LIST_WIDTH = 250


def now_stamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def new_task(title):
    """A task with every field present and defaults filled in."""
    task = {key: "" for key, _ in TASK_FIELDS}
    task.update(title=title, status=DEFAULT_STATUS, added_date=now_stamp(),
                jira_made=False, links=[])
    return task


def make_dot(master, color, size=16, radius=5.2):
    """A small filled circle image, used as a status marker in trees."""
    img = tk.PhotoImage(master=master, width=size, height=size)
    c = (size - 1) / 2
    for y in range(size):
        row = [color if (x - c) ** 2 + (y - c) ** 2 <= radius ** 2 else "" for x in range(size)]
        for x, px in enumerate(row):
            if px:
                img.put(px, (x, y))
    return img


def normalize_url(url):
    """Add https:// to bare addresses like 'example.com/page'."""
    if (re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", url)      # https://, ftp://
            or re.match(r"^(mailto|tel|file):", url, re.I)
            or re.match(r"^[a-zA-Z]:[\\/]", url)               # C:\path
            or url.startswith("\\\\")):                         # \\server\share
        return url
    return "https://" + url


def set_status(task, status):
    """Change status, stamping/clearing the completed date to match."""
    task["status"] = status
    if status == DONE:
        task["completed_date"] = task.get("completed_date") or now_stamp()
    else:
        task["completed_date"] = ""


def load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            settings = json.load(f)
        settings.pop("view", None)  # replaced by "default_view"
        return settings
    except (OSError, json.JSONDecodeError):
        return {}


def save_settings(settings):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Data layer
# ---------------------------------------------------------------------------
class Store:
    """Holds all projects/tasks and persists them to a local JSON file."""

    def __init__(self, path):
        self.path = path
        # {project_name: [task_dict, ...]}
        self.projects = {}
        # {"epic": {project_name: [name, ...]}, "sprint": {...}} - kept in
        # the order they were added (sprints are usually chronological).
        self.groups = {kind: {} for kind in GROUP_KINDS}
        self.load()

    def load(self):
        if os.path.exists(self.path):
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except (json.JSONDecodeError, OSError):
                data = {}
            if data.get("version") == 2:
                self.projects = data.get("projects", {})
                for kind in GROUP_KINDS:
                    self.groups[kind] = data.get(kind + "s", {})
            else:  # v1 file: the whole file was the projects dict
                self.projects = data
        self._migrate()

    def _migrate(self):
        """Fill in fields added after a task was first saved."""
        for project, tasks in self.projects.items():
            for task in tasks:
                if task.get("status") not in STATUSES:
                    task["status"] = DONE if task.get("completed_date") else DEFAULT_STATUS
                for kind in GROUP_KINDS:
                    task.setdefault(kind, "")
                task.setdefault("links", [])
            for task in tasks:
                if not isinstance(task.get("rank"), (int, float)):
                    task["rank"] = self._next_rank()
            # Any epic/sprint typed onto a task before lists existed.
            for kind in GROUP_KINDS:
                names = self.groups[kind].setdefault(project, [])
                for task in tasks:
                    if task[kind] and task[kind] not in names:
                        names.append(task[kind])

    def save(self):
        data = {"version": 2, "projects": self.projects}
        for kind in GROUP_KINDS:
            data[kind + "s"] = self.groups[kind]
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

    def add_project(self, name):
        if name in self.projects:
            return False
        self.projects[name] = []
        for kind in GROUP_KINDS:
            self.groups[kind][name] = []
        self.save()
        return True

    def delete_project(self, name):
        self.projects.pop(name, None)
        for kind in GROUP_KINDS:
            self.groups[kind].pop(name, None)
        self.save()

    def _next_rank(self):
        """Rank orders cards on the Visual Planner; new ones go last."""
        ranks = [t["rank"] for ts in self.projects.values() for t in ts
                 if isinstance(t.get("rank"), (int, float))]
        return max(ranks, default=-1) + 1

    def add_task(self, project, task):
        task.setdefault("rank", self._next_rank())
        self.projects[project].append(task)
        self.save()

    def move_task(self, src, index, dst, task):
        """Move a task to another project; returns its new index."""
        del self.projects[src][index]
        self.projects[dst].append(task)
        self.save()
        return len(self.projects[dst]) - 1

    def update_task(self, project, index, task):
        self.projects[project][index] = task
        self.save()

    def delete_task(self, project, index):
        del self.projects[project][index]
        self.save()

    # -- epics / sprints (kind is "epic" or "sprint") -----------------------
    def names(self, project, kind):
        return self.groups[kind].get(project, [])

    def usage(self, project, kind, name):
        return sum(1 for t in self.projects.get(project, []) if t.get(kind) == name)

    def add_name(self, project, kind, name):
        names = self.groups[kind].setdefault(project, [])
        if name in names:
            return False
        names.append(name)
        self.save()
        return True

    def rename_name(self, project, kind, old, new):
        names = self.groups[kind][project]
        if new in names:
            return False
        names[names.index(old)] = new
        for t in self.projects[project]:
            if t.get(kind) == old:
                t[kind] = new
        self.save()
        return True

    def delete_name(self, project, kind, name):
        self.groups[kind][project].remove(name)
        for t in self.projects[project]:
            if t.get(kind) == name:
                t[kind] = ""
        self.save()


# ---------------------------------------------------------------------------
# Main application window - everything lives here, no pop-up dialogs.
# ---------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self, store):
        super().__init__()
        self.store = store
        self.title("To-Do Tracker")
        self.settings = load_settings()
        self.view = None

        # The project the form is currently working under.
        self.active_project = None
        # Index of the task loaded in the form (None = unsaved new task).
        self.editing_index = None
        # Map tree item id -> ("project", name) or ("task", name, index).
        self.node_meta = {}
        # Map table row id -> task index within the active project.
        self.row_meta = {}
        # Table sort state.
        self.sort_col = "status"
        self.sort_rev = False
        # Guards selection handlers while we programmatically sync views.
        self._syncing = False

        self._build_menu()
        self.manager_frame = ttk.Frame(self)
        self.buddy_frame = tk.Frame(self, bg=BUDDY_BG)
        self.planner_frame = ttk.Frame(self)
        self.view_frames = {
            "Buddy": self.buddy_frame,
            "Manager": self.manager_frame,
            "Visual Planner": self.planner_frame,
        }
        self._build_layout()
        self._build_buddy()
        self._build_planner()
        self._refresh_views()

        self.bind_all("<Control-s>", lambda e: self._on_ctrl_s())
        self.bind_all("<Control-n>", lambda e: self._on_ctrl_n())
        for i, name in enumerate(VIEWS, start=1):
            self.bind_all(f"<Control-Key-{i}>", lambda e, n=name: self._show_view(n))
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._show_view(self.default_view_var.get())

    # -- views ------------------------------------------------------------
    def _build_menu(self):
        menubar = tk.Menu(self)
        self.view_var = tk.StringVar(value=DEFAULT_VIEW)
        view_menu = tk.Menu(menubar, tearoff=False)
        for i, name in enumerate(VIEWS, start=1):
            view_menu.add_radiobutton(
                label=name, variable=self.view_var, value=name,
                accelerator=f"Ctrl+{i}", command=lambda n=name: self._show_view(n),
            )

        # Which view the app opens in.
        view_menu.add_separator()
        view_menu.add_command(label="Default view", state="disabled")
        default = self.settings.get("default_view")
        self.default_view_var = tk.StringVar(
            value=default if default in VIEWS else DEFAULT_VIEW
        )
        for name in VIEWS:
            view_menu.add_radiobutton(
                label="    " + name, variable=self.default_view_var, value=name,
                command=self._save_default_view,
            )

        view_menu.add_separator()
        self.pin_var = tk.BooleanVar(value=self.settings.get("buddy_on_top", True))
        view_menu.add_checkbutton(
            label="Buddy always on top", variable=self.pin_var,
            command=self._apply_pin,
        )
        menubar.add_cascade(label="View", menu=view_menu)
        self.config(menu=menubar)

    def _save_default_view(self):
        self.settings["default_view"] = self.default_view_var.get()
        save_settings(self.settings)
        self._status(f"App will open in {self.default_view_var.get()} view.")

    def _default_geometry(self, view):
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        if view != "Buddy":
            w, h = min(1400, sw - 40), min(760, sh - 80)
            return f"{w}x{h}+{(sw - w) // 2}+{max(0, (sh - h) // 3)}"
        w, h = BUDDY_SIZE
        return f"{w}x{h}+{sw - w - 16}+16"  # top-right corner

    def _remember_geometry(self):
        if self.view:
            self.update_idletasks()  # make geometry() report what's on screen
            self.settings[f"geometry_{self.view}"] = self.geometry()

    def _show_view(self, view):
        if view == self.view:
            return
        self._remember_geometry()
        for frame in (*self.view_frames.values(), self.status_bar, self.form_frame):
            frame.pack_forget()
        self.view = view
        self.view_var.set(view)
        self.title(VIEW_TITLES[view])

        if view == "Buddy":
            self.minsize(*BUDDY_MIN)
            self.buddy_frame.pack(fill="both", expand=True)
            self._refresh_buddy()
        else:
            self.minsize(*MANAGER_MIN)
            self.status_bar.pack(side="bottom", fill="x")
            self.view_frames[view].pack(fill="both", expand=True)
            # One task form, shown inside whichever view needs it.
            slot = self.manager_form_slot if view == "Manager" else self.planner_form_slot
            self.form_frame.pack(in_=slot, fill="both", expand=True)
            self.form_frame.lift()
            self._refresh_views()
        # Pin first: toggling -topmost on Windows re-applies the old position.
        self._apply_pin()
        self.update_idletasks()
        self.geometry(
            self.settings.get(f"geometry_{view}") or self._default_geometry(view)
        )
        save_settings(self.settings)

    def _apply_pin(self):
        on_top = self.view == "Buddy" and self.pin_var.get()
        self.attributes("-topmost", on_top)
        self.settings["buddy_on_top"] = self.pin_var.get()
        self.pin_btn.configure(relief="sunken" if self.pin_var.get() else "flat")
        save_settings(self.settings)

    def _toggle_pin(self):
        self.pin_var.set(not self.pin_var.get())
        self._apply_pin()

    def _on_close(self):
        self._remember_geometry()
        save_settings(self.settings)
        self.destroy()

    def _on_ctrl_s(self):
        if self.view != "Buddy":
            self._save_task()

    def _on_ctrl_n(self):
        if self.view != "Buddy":
            self._new_task()
        else:
            self.b_entry.focus_set()

    # -- Visual Planner view ---------------------------------------------
    def _build_planner(self):
        f = self.planner_frame
        bar = ttk.Frame(f, padding=(8, 6))
        bar.pack(fill="x")

        def combo(label, key, values, default, width):
            ttk.Label(bar, text=label).pack(side="left")
            var = tk.StringVar(value=self.settings.get("planner_" + key, default))
            cb = ttk.Combobox(bar, textvariable=var, values=values,
                              state="readonly", width=width)
            cb.pack(side="left", padx=(4, 12))
            cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_planner())
            return var, cb

        self.p_columns, _ = combo("Columns", "columns", PLANNER_COLUMNS, "Status", 8)
        if self.p_columns.get() not in PLANNER_COLUMNS:
            self.p_columns.set("Status")
        self.p_project, self.p_project_cb = combo("Project", "project", [], ALL_PROJECTS, 16)
        self.p_epic, self.p_epic_cb = combo("Epic", "epic", [], ANY, 12)
        self.p_sprint, self.p_sprint_cb = combo("Sprint", "sprint", [], ANY, 12)
        self.p_order, _ = combo("Order", "order", PLANNER_ORDER, "Manual", 9)
        self.p_hide_done = tk.BooleanVar(value=self.settings.get("planner_hide_done", False))
        ttk.Checkbutton(bar, text="Hide Done", variable=self.p_hide_done,
                        command=self._refresh_planner).pack(side="left")
        ttk.Button(bar, text="Delete card", command=self._delete_task).pack(side="right")

        self.planner_form_slot = ttk.Frame(f, padding=(0, 0, 6, 6))
        self.planner_form_slot.pack(side="right", fill="y")

        wrap = tk.Frame(f, bg=BOARD_BG)
        wrap.pack(side="left", fill="both", expand=True, padx=6, pady=(0, 6))
        self.board = tk.Canvas(wrap, bg=BOARD_BG, highlightthickness=0,
                               xscrollincrement=20, yscrollincrement=20)
        xs = ttk.Scrollbar(wrap, orient="horizontal", command=self.board.xview)
        ys = ttk.Scrollbar(wrap, orient="vertical", command=self.board.yview)
        xs.pack(side="bottom", fill="x")
        ys.pack(side="right", fill="y")
        self.board.pack(side="left", fill="both", expand=True)
        self.board.configure(xscrollcommand=xs.set, yscrollcommand=ys.set)
        self.board_inner = tk.Frame(self.board, bg=BOARD_BG)
        self.board.create_window(0, 0, window=self.board_inner, anchor="nw")
        self.board_inner.bind(
            "<Configure>",
            lambda e: self.board.configure(scrollregion=self.board.bbox("all")),
        )
        self._bind_wheel(self.board)
        self.drop_line = tk.Frame(self.board_inner, bg=ACCENT, height=3)

        # [(column value, list frame, cards frame, [(task key, card frame)])]
        self.p_lists = []
        self.p_cards = {}  # (project, index) -> card frame
        self._drag = None

    def _bind_wheel(self, widget):
        widget.bind("<MouseWheel>", lambda e: self.board.yview_scroll(
            int(-e.delta / 120), "units"))
        widget.bind("<Shift-MouseWheel>", lambda e: self.board.xview_scroll(
            int(-e.delta / 120), "units"))
        for child in widget.winfo_children():
            self._bind_wheel(child)

    def _planner_scope(self):
        """Refresh filter choices; return the projects the board shows."""
        projects = sorted(self.store.projects, key=str.lower)
        self.p_project_cb.configure(values=[ALL_PROJECTS] + projects)
        if self.p_project.get() not in projects:
            self.p_project.set(ALL_PROJECTS)
        by_project = self.p_columns.get() == "Project"
        # Columns-by-project always shows every project.
        self.p_project_cb.configure(state="disabled" if by_project else "readonly")
        chosen = self.p_project.get()
        scope = projects if by_project or chosen == ALL_PROJECTS else [chosen]
        for kind, var, cb in (("epic", self.p_epic, self.p_epic_cb),
                              ("sprint", self.p_sprint, self.p_sprint_cb)):
            values = [ANY, NONE_LABEL] + self._names_in(scope, kind)
            cb.configure(values=values)
            if var.get() not in values:
                var.set(ANY)
        return scope

    def _names_in(self, scope, kind):
        """Epic/sprint names across projects, in each project's list order."""
        names = []
        for project in scope:
            for name in self.store.names(project, kind):
                if name not in names:
                    names.append(name)
        return names

    def _planner_columns(self, scope):
        mode = self.p_columns.get()
        if mode == "Status":
            return [(st, st) for st in STATUS_ORDER
                    if not (self.p_hide_done.get() and st == DONE)]
        if mode == "Project":
            return [(p, p) for p in scope]
        kind = kind_of(mode)
        return [("", f"No {mode.lower()}")] + [(n, n) for n in self._names_in(scope, kind)]

    def _card_value(self, project, task):
        mode = self.p_columns.get()
        return project if mode == "Project" else task.get(kind_of(mode), "")

    def _card_sort_key(self, task):
        order = self.p_order.get()
        if order == "Due date":
            due = task.get("due_date", "")
            return (due == "", due)
        if order == "Created":
            return (task.get("added_date", ""),)
        if order == "Title":
            return (task.get("title", "").lower(),)
        return (task.get("rank", 0),)

    def _refresh_planner(self):
        for key, var in (("columns", self.p_columns), ("project", self.p_project),
                         ("epic", self.p_epic), ("sprint", self.p_sprint),
                         ("order", self.p_order), ("hide_done", self.p_hide_done)):
            self.settings["planner_" + key] = var.get()
        save_settings(self.settings)

        scope = self._planner_scope()
        columns = self._planner_columns(scope)
        mode = self.p_columns.get()
        filters = {"epic": self.p_epic.get(), "sprint": self.p_sprint.get()}

        buckets = {value: [] for value, _ in columns}
        for project in scope:
            for i, t in enumerate(self.store.projects[project]):
                if self.p_hide_done.get() and t.get("status") == DONE:
                    continue
                if any(want != ANY and t.get(kind, "") != ("" if want == NONE_LABEL else want)
                       for kind, want in filters.items()):
                    continue
                value = self._card_value(project, t)
                if value in buckets:
                    buckets[value].append(((project, i), t))

        x0, y0 = self.board.xview()[0], self.board.yview()[0]
        for w in self.board_inner.winfo_children():
            if w is not self.drop_line:
                w.destroy()
        self.p_lists, self.p_cards = [], {}
        today = datetime.now().strftime("%Y-%m-%d")

        for col, (value, label) in enumerate(columns):
            items = sorted(buckets[value], key=lambda kt: self._card_sort_key(kt[1]))
            lst = tk.Frame(self.board_inner, bg=LIST_BG, highlightthickness=2,
                           highlightbackground=LIST_BG)
            lst.grid(row=0, column=col, sticky="n", padx=(8 if col == 0 else 4, 4), pady=8)
            tk.Frame(lst, width=LIST_WIDTH, height=0, bg=LIST_BG).pack()
            if mode == "Status":
                tk.Frame(lst, bg=STATUSES[value][0], height=4).pack(fill="x")
            head = tk.Frame(lst, bg=LIST_BG)
            head.pack(fill="x", padx=8, pady=(6, 2))
            tk.Label(head, text=label, bg=LIST_BG, font=("", 10, "bold"),
                     anchor="w").pack(side="left")
            tk.Label(head, text=str(len(items)), bg=LIST_BG, fg="#667",
                     font=("", 9)).pack(side="right")

            cards = tk.Frame(lst, bg=LIST_BG, height=6)
            cards.pack(fill="x", padx=6)
            entries = []
            for key, t in items:
                entries.append((key, self._make_card(cards, key, t, mode, len(scope) > 1, today)))

            add = tk.Label(lst, text="+ Add card", bg=LIST_BG, fg="#556",
                           anchor="w", cursor="hand2")
            add.pack(fill="x", padx=8, pady=(2, 8))
            add.bind("<Button-1>", lambda e, v=value: self._planner_add(v))
            self.p_lists.append((value, lst, cards, entries))

        self._bind_wheel(self.board_inner)
        self.board.update_idletasks()
        self.board.xview_moveto(x0)
        self.board.yview_moveto(y0)

    def _make_card(self, parent, key, task, mode, show_project, today):
        project = key[0]
        status = task.get("status", DEFAULT_STATUS)
        color = STATUSES[status][0]
        card = tk.Frame(parent, bg=CARD_BG, highlightthickness=1,
                        highlightbackground=CARD_BORDER, cursor="hand2")
        card.pack(fill="x", pady=3)
        tk.Frame(card, bg=color, width=5).pack(side="left", fill="y")
        body = tk.Frame(card, bg=CARD_BG)
        body.pack(side="left", fill="both", expand=True, padx=8, pady=6)
        tk.Label(body, text=task.get("title") or "(untitled)", bg=CARD_BG,
                 font=("", 10), wraplength=LIST_WIDTH - 40, justify="left",
                 anchor="w").pack(fill="x")

        # Status + due date on one line.
        due = task.get("due_date", "")
        if mode != "Status" or due:
            row = tk.Frame(body, bg=CARD_BG)
            row.pack(fill="x", pady=(3, 0))
            if mode != "Status":
                tk.Label(row, text="\u25cf " + status, bg=CARD_BG, fg=color,
                         font=("", 8, "bold")).pack(side="left")
            if due:
                overdue = due < today and status != DONE
                tk.Label(row, text=("! " if overdue else "") + "due " + due,
                         bg=CARD_BG, fg="#c0392b" if overdue else "#667",
                         font=("", 8, "bold" if overdue else "normal")).pack(side="right")

        # Everything the columns aren't already showing.
        meta = []
        if show_project and mode != "Project":
            meta.append(project)
        if task.get("group") and mode != "Subgroup":
            meta.append("\u25a3 " + task["group"])
        if task.get("epic") and mode != "Epic":
            meta.append("\u25c6 " + task["epic"])
        if task.get("sprint") and mode != "Sprint":
            meta.append("\u21bb " + task["sprint"])
        if task.get("jira_ref"):
            meta.append(task["jira_ref"])
        if task.get("links"):
            meta.append(f"\U0001F517 {len(task['links'])}")
        if meta:
            tk.Label(body, text="   ".join(meta), bg=CARD_BG, fg="#667",
                     font=("", 8), wraplength=LIST_WIDTH - 40, justify="left",
                     anchor="w").pack(fill="x", pady=(2, 0))

        self._bind_drag(card, key)
        self.p_cards[key] = card
        return card

    def _highlight_cards(self):
        current = (self.active_project, self.editing_index)
        for key, card in self.p_cards.items():
            on = key == current
            card.configure(highlightthickness=2 if on else 1,
                           highlightbackground=ACCENT if on else CARD_BORDER)

    def _planner_add(self, value):
        """'+ Add card': start a new task with this column's value preset."""
        mode = self.p_columns.get()
        if mode == "Project":
            project = value
        elif self.p_project.get() != ALL_PROJECTS:
            project = self.p_project.get()
        else:
            project = self.active_project
        if not project:
            self._status("Pick a project (top bar) to add cards to.")
            return
        self.active_project = project
        self.editing_index = None
        self._clear_form()
        if mode == "Status":
            self.f_status.set(value)
            self._on_status_change()
        elif kind_of(mode) in GROUP_KINDS and value:
            kind = kind_of(mode)
            if value not in self.store.names(project, kind):
                self.store.add_name(project, kind, value)
                self._refresh_choices()
            self.f_kinds[kind].set(value)
        self._highlight_cards()
        self.f_title.focus_set()
        self._status(f"New card in '{project}' - give it a title, then Save (Ctrl+S).")

    # drag and drop: press/motion/release are bound on every widget in a card
    def _bind_drag(self, widget, key):
        widget.bind("<ButtonPress-1>", lambda e: self._drag_start(e, key))
        widget.bind("<B1-Motion>", self._drag_move)
        widget.bind("<ButtonRelease-1>", self._drag_end)
        for child in widget.winfo_children():
            self._bind_drag(child, key)

    def _drag_start(self, event, key):
        self._drag = {"key": key, "x": event.x_root, "y": event.y_root, "ghost": None}

    def _drag_move(self, event):
        d = self._drag
        if not d:
            return
        if d["ghost"] is None:
            if abs(event.x_root - d["x"]) + abs(event.y_root - d["y"]) < 6:
                return  # still a click
            project, index = d["key"]
            task = self.store.projects[project][index]
            ghost = tk.Toplevel(self)
            ghost.overrideredirect(True)
            ghost.attributes("-topmost", True)
            try:
                ghost.attributes("-alpha", 0.85)
            except tk.TclError:
                pass
            tk.Label(ghost, text=task.get("title") or "(untitled)", bg=CARD_BG,
                     fg=STATUSES[task.get("status", DEFAULT_STATUS)][0],
                     font=("", 10, "bold"), padx=10, pady=6, relief="solid", bd=1,
                     wraplength=LIST_WIDTH - 40, justify="left").pack()
            d["ghost"] = ghost
            self.p_cards[d["key"]].configure(bg="#eef2f7")
        d["ghost"].geometry(f"+{event.x_root + 12}+{event.y_root + 8}")
        self._autoscroll(event.x_root, event.y_root)
        self._show_drop(self._drop_target(event.x_root, event.y_root, d["key"]))

    def _drag_end(self, event):
        d, self._drag = self._drag, None
        if not d:
            return
        if d["ghost"] is None:  # a click: edit the card
            self._select_task(*d["key"])
            return
        d["ghost"].destroy()
        target = self._drop_target(event.x_root, event.y_root, d["key"])
        self._show_drop(None)
        if target:
            self._move_card(d["key"], *target)
        else:
            self._refresh_planner()

    def _drop_target(self, x, y, key):
        """(column index, insert position, other cards) under the pointer."""
        for col, (value, lst, cards, entries) in enumerate(self.p_lists):
            left = lst.winfo_rootx()
            if left <= x < left + lst.winfo_width():
                others = [(k, c) for k, c in entries if k != key]
                pos = len(others)
                for n, (_, c) in enumerate(others):
                    if y < c.winfo_rooty() + c.winfo_height() / 2:
                        pos = n
                        break
                return col, pos, others
        return None

    def _show_drop(self, target):
        for _, lst, _, _ in self.p_lists:
            lst.configure(highlightbackground=LIST_BG)
        self.drop_line.place_forget()
        if not target:
            return
        col, pos, others = target
        _, lst, cards, _ = self.p_lists[col]
        lst.configure(highlightbackground=ACCENT)
        if self.p_order.get() != "Manual":
            return  # order is automatic; only the column matters
        if not others:
            y = 0
        elif pos < len(others):
            y = others[pos][1].winfo_y() - 3
        else:
            last = others[-1][1]
            y = last.winfo_y() + last.winfo_height() + 1
        self.drop_line.place(in_=cards, x=0, y=max(0, y), relwidth=1)
        self.drop_line.lift()

    def _autoscroll(self, x, y):
        b = self.board
        if x < b.winfo_rootx() + 30:
            b.xview_scroll(-1, "units")
        elif x > b.winfo_rootx() + b.winfo_width() - 30:
            b.xview_scroll(1, "units")
        if y < b.winfo_rooty() + 20:
            b.yview_scroll(-1, "units")
        elif y > b.winfo_rooty() + b.winfo_height() - 20:
            b.yview_scroll(1, "units")

    def _move_card(self, key, col, pos, others):
        value = self.p_lists[col][0]
        project, index = key
        task = dict(self.store.projects[project][index])
        mode = self.p_columns.get()

        if self.p_order.get() == "Manual" and others:
            ranks = [self.store.projects[p][i].get("rank", 0) for (p, i), _ in others]
            if pos == 0:
                task["rank"] = ranks[0] - 1
            elif pos == len(ranks):
                task["rank"] = ranks[-1] + 1
            else:
                task["rank"] = (ranks[pos - 1] + ranks[pos]) / 2

        if mode == "Status":
            set_status(task, value)
        elif kind_of(mode) in GROUP_KINDS:
            task[kind_of(mode)] = value

        dest = value if mode == "Project" else project
        self._relocate_task(project, index, task, dest)
        self._refresh_planner()
        self._highlight_cards()
        where = f"{mode.lower()} '{value or 'none'}'"
        self._status(f"Moved '{task.get('title', '')}' to {where}.")

    def _relocate_task(self, project, index, task, dest):
        """Save a task that was dragged somewhere - possibly into another
        project - keeping the form pointed at the right task."""
        # Groups/epics/sprints are per project: make sure dest knows this task's.
        for kind in GROUP_KINDS:
            names = self.store.groups[kind].setdefault(dest, [])
            if task.get(kind) and task[kind] not in names:
                names.append(task[kind])

        editing = (self.active_project, self.editing_index) == (project, index)
        if dest != project:
            new_index = self.store.move_task(project, index, dest, task)
            if editing:
                self.active_project, self.editing_index = dest, new_index
            elif (self.active_project == project and self.editing_index is not None
                  and self.editing_index > index):
                self.editing_index -= 1
        else:
            self.store.update_task(project, index, task)

        if editing:
            # Update only what the drag changed; keep other unsaved edits.
            self._refresh_choices()
            self.f_status.set(task["status"])
            self._completed_date = task.get("completed_date", "")
            self.f_completed_lbl.configure(text=f"Completed: {self._completed_date or '-'}")
            self._update_status_pill()
            for kind, combo in self.f_kinds.items():
                combo.set(task.get(kind, ""))

    # -- Buddy view -------------------------------------------------------
    def _build_buddy(self):
        f = self.buddy_frame
        style = ttk.Style(self)
        style.configure(
            "Buddy.Treeview", background=BUDDY_BG, fieldbackground=BUDDY_BG,
            borderwidth=0, rowheight=22,
        )
        style.layout("Buddy.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

        top = tk.Frame(f, bg=BUDDY_BG)
        top.pack(fill="x", padx=6, pady=(6, 2))
        self.b_sprint = ttk.Combobox(top, state="readonly", width=16)
        self.b_sprint.set(self.settings.get("buddy_sprint", ALL_SPRINTS))
        self.b_sprint.pack(side="left", fill="x", expand=True)
        self.b_sprint.bind("<<ComboboxSelected>>", lambda e: self._refresh_buddy())
        self.pin_btn = tk.Button(
            top, text="\U0001F4CC", bg=BUDDY_BG, activebackground=BUDDY_BG, bd=1,
            command=self._toggle_pin,
        )
        self.pin_btn.pack(side="left", padx=(4, 0))
        self.b_hide_done = tk.BooleanVar(value=self.settings.get("buddy_hide_done", True))
        tk.Checkbutton(
            top, text="Hide done", variable=self.b_hide_done, bg=BUDDY_BG,
            activebackground=BUDDY_BG, command=self._refresh_buddy,
        ).pack(side="left", padx=(4, 0))

        self.b_summary = tk.Label(
            f, text="", bg=BUDDY_BG, fg="#665", anchor="w", font=("", 8)
        )
        self.b_summary.pack(fill="x", padx=8)
        self.b_summary.bind(
            "<Configure>", lambda e: self.b_summary.configure(wraplength=e.width - 4)
        )

        add = tk.Frame(f, bg=BUDDY_BG)
        add.pack(side="bottom", fill="x", padx=6, pady=6)
        self.b_add_lbl = tk.Label(add, text="", bg=BUDDY_BG, fg="#665", font=("", 8))
        self.b_add_lbl.pack(side="left", padx=(0, 4))
        self.b_entry = ttk.Entry(add)
        self.b_entry.pack(side="left", fill="x", expand=True)
        self.b_entry.bind("<Return>", lambda e: self._buddy_add())
        self.b_add_btn = ttk.Button(add, text="+", width=3, command=self._buddy_add)
        self.b_add_btn.pack(side="left", padx=(4, 0))

        wrap = tk.Frame(f, bg=BUDDY_BG)
        wrap.pack(fill="both", expand=True, padx=6)
        self.b_tree = ttk.Treeview(
            wrap, columns=("due",), show="tree", selectmode="browse",
            style="Buddy.Treeview",
        )
        self.b_tree.column("#0", width=200, stretch=True)
        self.b_tree.column("due", width=62, stretch=False, anchor="e")
        self.b_tree.pack(side="left", fill="both", expand=True)
        ys = ttk.Scrollbar(wrap, orient="vertical", command=self.b_tree.yview)
        ys.pack(side="left", fill="y")
        self.b_tree.configure(yscrollcommand=ys.set)
        self.b_tree.tag_configure("project", font=("", 10, "bold"))
        self.b_tree.tag_configure("overdue", font=("", 9, "bold"))
        # Status shows as a colored dot; text keeps the normal color.
        self.b_dots = {name: make_dot(self, color) for name, (color, _) in STATUSES.items()}
        self.b_tree.bind("<Double-1>", self._buddy_dblclick)
        self.b_tree.bind("<<TreeviewSelect>>", lambda e: self._buddy_update_add())
        self.b_tree.bind("<<TreeviewOpen>>", lambda e: self._buddy_fold(True))
        self.b_tree.bind("<<TreeviewClose>>", lambda e: self._buddy_fold(False))
        # project/group row id -> (project, group); group is "" for a project row
        self.b_folder_rows = {}
        self.b_tree.tag_configure("group", font=("", 9, "bold"), foreground="#554")
        self.b_tree.bind("<Button-3>", self._buddy_menu)
        self.b_meta = {}  # row id -> (project, index)

        self.b_status_menu = tk.Menu(self, tearoff=False)

    def _buddy_sprints(self, projects):
        """Sprint filter choices: every project's sprints, in list order."""
        names = []
        for project in projects:
            for name in self.store.names(project, "sprint"):
                if name not in names:
                    names.append(name)
        values = [ALL_SPRINTS, NO_SPRINT] + names
        self.b_sprint.configure(values=values)
        if self.b_sprint.get() not in values:
            self.b_sprint.set(ALL_SPRINTS)
        return self.b_sprint.get()

    def _refresh_buddy(self):
        projects = sorted(self.store.projects, key=str.lower)
        sprint = self._buddy_sprints(projects)
        self.settings["buddy_sprint"] = sprint
        self.settings["buddy_hide_done"] = self.b_hide_done.get()
        # Collapsed folders: "project" or "project/group".
        collapsed = set(self.settings.get("buddy_collapsed", []))
        today = datetime.now().strftime("%Y-%m-%d")

        selected = self.b_tree.selection()
        keep = None
        if selected:
            keep = self.b_meta.get(selected[0]) or self.b_folder_rows.get(selected[0])

        self.b_tree.delete(*self.b_tree.get_children())
        self.b_meta, self.b_folder_rows = {}, {}
        counts = {st: 0 for st in STATUS_ORDER}
        reselect = None

        def folder(parent, project, group, items):
            open_n = sum(1 for _, t in items if t.get("status") != DONE)
            fold_key = f"{project}/{group}" if group else project
            row = self.b_tree.insert(
                parent, "end", text=f"{group or project}  ({open_n})",
                open=fold_key not in collapsed, tags=("group" if group else "project",),
            )
            self.b_folder_rows[row] = (project, group)
            return row

        for project in projects:
            items = []
            for i, t in enumerate(self.store.projects[project]):
                if self.b_hide_done.get() and t.get("status") == DONE:
                    continue
                if sprint == NO_SPRINT and t.get("sprint"):
                    continue
                if sprint not in (ALL_SPRINTS, NO_SPRINT) and t.get("sprint") != sprint:
                    continue
                items.append((i, t))
            items.sort(key=lambda it: (
                BUDDY_ORDER.index(it[1].get("status", DEFAULT_STATUS)),
                it[1].get("due_date") or "~",
            ))
            for _, t in items:
                counts[t.get("status", DEFAULT_STATUS)] += 1

            parents = {"": folder("", project, "", items)}
            for name in self.store.names(project, "group"):
                mine = [it for it in items if it[1].get("group") == name]
                parents[name] = folder(parents[""], project, name, mine)
            for row, key in self.b_folder_rows.items():
                if key == keep:
                    reselect = row

            for i, t in items:
                status = t.get("status", DEFAULT_STATUS)
                due = t.get("due_date", "")
                overdue = due and due < today and status != DONE
                if len(due) == 10:
                    due = due[5:]  # MM-DD keeps the column narrow
                if overdue:
                    due = "! " + due
                rid = self.b_tree.insert(
                    parents.get(t.get("group"), parents[""]), "end",
                    text=" " + (t.get("title") or "(untitled)"),
                    image=self.b_dots[status], values=(due,),
                    tags=("overdue",) if overdue else (),
                )
                self.b_meta[rid] = (project, i)
                if keep == (project, i):
                    reselect = rid

        if reselect:
            self.b_tree.selection_set(reselect)
        parts = [f"{n} {st.lower().replace(' for approval', '')}"
                 for st, n in counts.items() if n and st != DONE]
        self.b_summary.configure(text="  \u00b7  ".join(parts) or "Nothing open. Nice.")
        self._buddy_update_add()

    def _buddy_fold(self, opened):
        """Remember which projects/groups are collapsed."""
        key = self.b_folder_rows.get(self.b_tree.focus())
        if not key:
            return
        project, group = key
        fold_key = f"{project}/{group}" if group else project
        collapsed = set(self.settings.get("buddy_collapsed", []))
        (collapsed.discard if opened else collapsed.add)(fold_key)
        self.settings["buddy_collapsed"] = sorted(collapsed)
        save_settings(self.settings)

    def _buddy_target(self):
        """(project, group) quick-add goes to: the selected row's, else the
        last project used."""
        sel = self.b_tree.selection()
        if sel:
            if sel[0] in self.b_folder_rows:
                return self.b_folder_rows[sel[0]]
            if sel[0] in self.b_meta:
                project, i = self.b_meta[sel[0]]
                return project, self.store.projects[project][i].get("group", "")
        last = self.settings.get("buddy_add_project")
        return (last, "") if last in self.store.projects else (None, "")

    def _buddy_update_add(self):
        project, group = self._buddy_target()
        where = f"{project} / {group}" if group else project
        self.b_add_lbl.configure(text=f"{where} \u25b8" if project else "select a project \u25b8")
        for w in (self.b_entry, self.b_add_btn):
            w.state(["!disabled"] if project else ["disabled"])

    def _buddy_add(self):
        project, group = self._buddy_target()
        title = self.b_entry.get().strip()
        if not title or not project:
            return
        task = new_task(title)
        task["group"] = group
        sprint = self.b_sprint.get()
        if sprint not in (ALL_SPRINTS, NO_SPRINT) and sprint in self.store.names(project, "sprint"):
            task["sprint"] = sprint  # stays visible under the current filter
        self.store.add_task(project, task)
        self.settings["buddy_add_project"] = project
        self.b_entry.delete(0, "end")
        self._refresh_buddy()

    def _buddy_dblclick(self, event):
        rid = self.b_tree.identify_row(event.y)
        if rid in self.b_meta:
            self._buddy_open_task(*self.b_meta[rid])

    def _buddy_menu(self, event):
        rid = self.b_tree.identify_row(event.y)
        if rid not in self.b_meta:
            return
        self.b_tree.selection_set(rid)
        project, index = self.b_meta[rid]
        current = self.store.projects[project][index].get("status")
        m = self.b_status_menu
        m.delete(0, "end")
        for name in STATUS_ORDER:
            color = STATUSES[name][0]
            m.add_command(
                label=("\u2713 " if current == name else "    ") + name,
                foreground=color, activeforeground=color,
                command=lambda n=name: self._buddy_set_status(project, index, n),
            )
        m.add_separator()
        m.add_command(label="Open in Manager",
                      command=lambda: self._buddy_open_task(project, index))
        m.tk_popup(event.x_root, event.y_root)

    def _buddy_open_task(self, project, index):
        self._show_view("Manager")
        self.notebook.select(0)
        self._select_task(project, index)

    def _buddy_set_status(self, project, index, status):
        task = dict(self.store.projects[project][index])
        set_status(task, status)
        self.store.update_task(project, index, task)
        # Keep the Manager form in step if it has this task open.
        if project == self.active_project and index == self.editing_index:
            self.f_status.set(status)
            self._completed_date = task["completed_date"]
            self.f_completed_lbl.configure(
                text=f"Completed: {self._completed_date or '-'}")
            self._update_status_pill()
        self._refresh_buddy()

    # -- layout -----------------------------------------------------------
    def _build_layout(self):
        # Status bar (bottom)
        self.status_var = tk.StringVar(value="Ready.")
        # Shared by Manager and Visual Planner; packed by _show_view.
        status_bar = self.status_bar = ttk.Frame(self, relief="sunken")
        ttk.Label(
            status_bar, textvariable=self.status_var, anchor="w", padding=(8, 2)
        ).pack(side="left", fill="x", expand=True)
        for name, (color, _) in STATUSES.items():
            tk.Label(
                status_bar, text=" " + name + " ", bg=color, fg="white",
                font=("", 8, "bold"), padx=2,
            ).pack(side="left", padx=1, pady=2)

        panes = ttk.Panedwindow(self.manager_frame, orient="horizontal")
        panes.pack(fill="both", expand=True)

        left = ttk.Frame(panes, padding=(6, 6, 4, 6))
        center = ttk.Frame(panes, padding=(4, 6))
        self.manager_form_slot = ttk.Frame(panes)
        panes.add(left, weight=0)
        panes.add(center, weight=1)
        panes.add(self.manager_form_slot, weight=0)
        # The form is a child of the root so it can be shown inside either
        # the Manager or the Visual Planner (pack in_=slot).
        form = self.form_frame = ttk.LabelFrame(
            self, text="Task Details", padding=(8, 4)
        )

        self._build_sidebar(left)
        self._build_table(center)
        self._build_form(form)

    def _build_sidebar(self, left):
        ttk.Label(left, text="Projects & Tasks", font=("", 11, "bold")).pack(
            anchor="w"
        )

        # One add box: a new project when nothing is selected, otherwise a
        # new subgroup inside the selected project.
        mode_row = ttk.Frame(left)
        mode_row.pack(fill="x", pady=(4, 0))
        self.add_mode_lbl = ttk.Label(mode_row, text="New project", foreground="#555",
                                      font=("", 8))
        self.add_mode_lbl.pack(side="left")
        self.add_switch = ttk.Label(mode_row, text="new project instead",
                                    style="LinkAction.TLabel", cursor="hand2")
        self.add_switch.bind("<Button-1>", lambda e: self._clear_tree_selection())

        add_row = ttk.Frame(left)
        add_row.pack(fill="x", pady=(1, 4))
        self.project_entry = ttk.Entry(add_row, width=18)
        self.project_entry.pack(side="left", fill="x", expand=True)
        self.project_entry.bind("<Return>", lambda e: self._add_from_box())
        self.project_entry.bind("<Escape>", lambda e: self._clear_tree_selection())
        self.add_btn = ttk.Button(add_row, text="+ Project", width=11,
                                  command=self._add_from_box)
        self.add_btn.pack(side="left", padx=(4, 0))

        btns = ttk.Frame(left)
        btns.pack(side="bottom", fill="x", pady=(4, 0))
        ttk.Button(btns, text="New Task", command=self._new_task).grid(
            row=0, column=0, sticky="ew", padx=1
        )
        ttk.Button(btns, text="Del Task", command=self._delete_task).grid(
            row=0, column=1, sticky="ew", padx=1
        )
        ttk.Button(btns, text="Delete Project", command=self._delete_project).grid(
            row=1, column=0, columnspan=2, sticky="ew", padx=1, pady=(3, 0)
        )
        btns.columnconfigure((0, 1), weight=1)

        tree_wrap = ttk.Frame(left)
        tree_wrap.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_wrap, show="tree", selectmode="browse")
        self.tree.column("#0", width=220)
        self.tree.pack(side="left", fill="both", expand=True)
        yscroll = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        yscroll.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._on_tree_select())
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._update_add_mode(), add="+")
        # Drag tasks onto a group / project / other task to regroup them.
        self.tree.bind("<ButtonPress-1>", self._tree_press, add="+")
        self.tree.bind("<B1-Motion>", self._tree_motion, add="+")
        self.tree.bind("<ButtonRelease-1>", self._tree_release, add="+")
        self._tdrag = None
        self.tree.tag_configure("drop", background="#cfe0fb")
        self.tree.tag_configure("group", font=("", 9, "italic"), foreground="#446")
        self.tree.tag_configure("project", font=("", 10, "bold"))
        for color, tag in STATUSES.values():
            self.tree.tag_configure(tag, foreground=color)

    def _build_table(self, center):
        # Header: project title on the left, exports on the right.
        header = ttk.Frame(center)
        header.pack(fill="x")
        self.context_var = tk.StringVar(value="No project selected.")
        ttk.Label(
            header, textvariable=self.context_var, font=("", 11, "bold")
        ).pack(side="left")

        self.dot = tk.Canvas(header, width=14, height=14, highlightthickness=0)
        color = "#2ecc71" if HAVE_OPENPYXL else "#e74c3c"
        self.dot.create_oval(2, 2, 12, 12, fill=color, outline="")
        self.xlsx_btn = ttk.Button(header, text="Export XLSX", command=self._export_xlsx)
        self.xlsx_btn.pack(side="right")
        ttk.Button(header, text="Export CSV", command=self._export_csv).pack(
            side="right", padx=4
        )
        self.dot.pack(side="right", padx=(0, 2))
        if not HAVE_OPENPYXL:
            self.xlsx_btn.state(["disabled"])

        self.notebook = ttk.Notebook(center)
        self.notebook.pack(fill="both", expand=True, pady=(6, 0))
        tasks_tab = ttk.Frame(self.notebook, padding=4)
        groups_tab = ttk.Frame(self.notebook, padding=4)
        self.notebook.add(tasks_tab, text="Tasks")
        self.notebook.add(groups_tab, text="Subgroups, Epics & Sprints")
        self._build_groups_tab(groups_tab)
        center = tasks_tab

        # Organize controls.
        ctrl = ttk.Frame(center)
        ctrl.pack(fill="x", pady=(0, 4))
        ttk.Label(ctrl, text="Group by").pack(side="left")
        self.group_var = tk.StringVar(value="Epic")
        group_cb = ttk.Combobox(
            ctrl, textvariable=self.group_var, values=GROUP_OPTIONS,
            state="readonly", width=8,
        )
        group_cb.pack(side="left", padx=(4, 12))
        group_cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_table())
        self.hide_done = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            ctrl, text="Hide Done", variable=self.hide_done,
            command=self._refresh_table,
        ).pack(side="left")

        wrap = ttk.Frame(center)
        wrap.pack(fill="both", expand=True)
        self.table = ttk.Treeview(
            wrap, columns=[c[0] for c in TABLE_COLS], selectmode="browse",
        )
        self.table.column("#0", width=150, minwidth=120, stretch=True)
        self.table.heading("#0", text="Title", command=lambda: self._sort_by("title"))
        for key, heading, width, stretch in TABLE_COLS:
            self.table.heading(key, text=heading, command=lambda k=key: self._sort_by(k))
            self.table.column(key, width=width, minwidth=50, stretch=stretch)
        self.table.pack(side="left", fill="both", expand=True)
        ys = ttk.Scrollbar(wrap, orient="vertical", command=self.table.yview)
        ys.pack(side="left", fill="y")
        self.table.configure(yscrollcommand=ys.set)
        self.table.bind("<<TreeviewSelect>>", lambda e: self._on_table_select())
        self.table.tag_configure("group", font=("", 10, "bold"), background="#eef1f6")
        for color, tag in STATUSES.values():
            self.table.tag_configure(tag, foreground=color)

    def _build_groups_tab(self, tab):
        top = ttk.Frame(tab)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, text="Project").pack(side="left")
        self.g_project = ttk.Combobox(top, state="readonly", width=24)
        self.g_project.pack(side="left", padx=6)
        self.g_project.bind("<<ComboboxSelected>>", lambda e: self._on_group_project())
        ttk.Label(
            top, text="Select an item to rename or delete it.", foreground="#888"
        ).pack(side="left", padx=8)

        cols = ttk.Frame(tab)
        cols.pack(fill="both", expand=True)
        cols.columnconfigure(tuple(range(len(GROUP_KINDS))), weight=1, uniform="g")
        cols.rowconfigure(0, weight=1)
        # kind -> (listing treeview, name entry)
        self.g_widgets = {}
        for c, kind in enumerate(GROUP_KINDS):
            box = ttk.LabelFrame(cols, text=KIND_LABEL[kind] + "s", padding=6)
            box.grid(row=0, column=c, sticky="nsew", padx=(0 if c == 0 else 4, 0))

            entry_row = ttk.Frame(box)
            entry_row.pack(side="bottom", fill="x", pady=(6, 0))
            entry = ttk.Entry(entry_row)
            entry.pack(side="left", fill="x", expand=True)
            entry.bind("<Return>", lambda e, k=kind: self._group_add(k))
            for text, cmd in (("Add", self._group_add), ("Rename", self._group_rename),
                              ("Delete", self._group_delete)):
                ttk.Button(
                    entry_row, text=text, width=7, command=lambda k=kind, f=cmd: f(k)
                ).pack(side="left", padx=(4, 0))

            lst = ttk.Treeview(box, columns=("tasks", "open"), selectmode="browse")
            lst.heading("#0", text="Name")
            lst.heading("tasks", text="Tasks")
            lst.heading("open", text="Open")
            lst.column("#0", width=160, stretch=True)
            lst.column("tasks", width=55, stretch=False, anchor="center")
            lst.column("open", width=55, stretch=False, anchor="center")
            lst.pack(fill="both", expand=True)
            lst.bind("<<TreeviewSelect>>", lambda e, k=kind: self._on_group_select(k))
            self.g_widgets[kind] = (lst, entry)

    def _build_form(self, form):
        pad = {"padx": 4, "pady": 3}
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)
        r = 0

        ttk.Label(form, text="Title *").grid(row=r, column=0, sticky="e", **pad)
        self.f_title = ttk.Entry(form, width=30)
        self.f_title.grid(row=r, column=1, columnspan=3, sticky="ew", **pad)
        r += 1

        ttk.Label(form, text="Status").grid(row=r, column=0, sticky="e", **pad)
        st_wrap = ttk.Frame(form)
        st_wrap.grid(row=r, column=1, columnspan=3, sticky="ew", **pad)
        self.f_status = tk.StringVar(value=DEFAULT_STATUS)
        self.status_cb = ttk.Combobox(
            st_wrap, textvariable=self.f_status, values=STATUS_ORDER,
            state="readonly", width=20,
        )
        self.status_cb.pack(side="left")
        self.status_cb.bind("<<ComboboxSelected>>", lambda e: self._on_status_change())
        self.status_pill = tk.Label(
            st_wrap, text="", fg="white", font=("", 9, "bold"), padx=8
        )
        self.status_pill.pack(side="left", padx=6)
        r += 1

        ttk.Label(form, text="Epic").grid(row=r, column=0, sticky="e", **pad)
        self.f_epic = ttk.Combobox(form, width=12)
        self.f_epic.grid(row=r, column=1, sticky="ew", **pad)
        ttk.Label(form, text="Sprint").grid(row=r, column=2, sticky="e", **pad)
        self.f_sprint = ttk.Combobox(form, width=10)
        self.f_sprint.grid(row=r, column=3, sticky="ew", **pad)
        r += 1

        ttk.Label(form, text="Subgroup").grid(row=r, column=0, sticky="e", **pad)
        self.f_group = ttk.Combobox(form, width=12)
        self.f_group.grid(row=r, column=1, sticky="ew", **pad)
        ttk.Label(form, text="Due").grid(row=r, column=2, sticky="e", **pad)
        self.f_due = ttk.Entry(form, width=10)
        self.f_due.grid(row=r, column=3, sticky="ew", **pad)
        self.f_due.bind("<FocusIn>", lambda e: self._status("Due date format: YYYY-MM-DD"))
        r += 1
        # kind -> dropdown, for code that treats group/epic/sprint alike
        self.f_kinds = {"group": self.f_group, "epic": self.f_epic, "sprint": self.f_sprint}

        ttk.Label(form, text="Jira").grid(row=r, column=0, sticky="e", **pad)
        self.f_jira_made = tk.BooleanVar(value=False)
        self.jira_chk = ttk.Checkbutton(
            form, text="Ticket made", variable=self.f_jira_made,
            command=self._toggle_jira,
        )
        self.jira_chk.grid(row=r, column=1, sticky="w", **pad)
        self.f_jira_ref = ttk.Entry(form, width=14)
        self.f_jira_ref.grid(row=r, column=2, columnspan=2, sticky="ew", **pad)
        r += 1

        # Text areas share the remaining height.
        self.f_desc = self._text_row(form, r, "Description", 4, weight=2)
        r += 1
        self.f_blockers = self._text_row(form, r, "Blockers", 2, weight=1)
        r += 1
        self.f_notes = self._text_row(form, r, "Notes", 4, weight=2)
        r += 1

        ttk.Label(form, text="Links").grid(row=r, column=0, sticky="ne", **pad)
        self._build_links(form, r)
        r += 1

        dates = ttk.Frame(form)
        dates.grid(row=r, column=1, columnspan=3, sticky="w", padx=4)
        self.f_added_lbl = ttk.Label(dates, text="Created: -", foreground="#666",
                                     font=("", 8))
        self.f_added_lbl.pack(side="left")
        self.f_completed_lbl = ttk.Label(dates, text="Completed: -", foreground="#666",
                                         font=("", 8))
        self.f_completed_lbl.pack(side="left", padx=(12, 0))
        r += 1

        btnrow = ttk.Frame(form)
        btnrow.grid(row=r, column=0, columnspan=4, pady=(6, 2))
        self.save_btn = ttk.Button(btnrow, text="Save  (Ctrl+S)", command=self._save_task)
        self.save_btn.pack(side="left", padx=4)
        ttk.Button(btnrow, text="New  (Ctrl+N)", command=self._new_task).pack(
            side="left", padx=4
        )

        self._added_date = ""
        self._completed_date = ""
        self._update_status_pill()
        self._toggle_jira()
        self._set_form_enabled(False)

    def _build_links(self, form, r):
        style = ttk.Style(self)
        style.configure("Link.TLabel", foreground="#1a5fb4", font=("", 9, "underline"))
        style.configure("LinkAction.TLabel", foreground="#888", font=("", 8))

        box = ttk.Frame(form)
        box.grid(row=r, column=1, columnspan=3, sticky="ew", padx=4, pady=3)
        box.columnconfigure(1, weight=1)
        self.f_links_list = ttk.Frame(box)
        self.f_links_list.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 4))

        self.f_link_title = ttk.Entry(box)
        self.f_link_desc = ttk.Entry(box)
        self.f_link_url = ttk.Entry(box)
        for i, (label, entry) in enumerate((("Link title", self.f_link_title),
                                           ("Link description", self.f_link_desc),
                                           ("Link", self.f_link_url)), start=1):
            ttk.Label(box, text=label, foreground="#555", font=("", 8)).grid(
                row=i, column=0, sticky="w", padx=(0, 6))
            entry.grid(row=i, column=1, columnspan=2 if i < 3 else 1, sticky="ew", pady=1)
            entry.bind("<Return>", lambda e: self._add_link())
        self.f_link_btn = ttk.Button(box, text="Add link", width=9, command=self._add_link)
        self.f_link_btn.grid(row=3, column=2, padx=(4, 0))

        self._links = []
        self._link_edit_index = None

    def _render_links(self):
        for w in self.f_links_list.winfo_children():
            w.destroy()
        if not self._links:
            ttk.Label(self.f_links_list, text="No links yet.", foreground="#888",
                      font=("", 8)).pack(anchor="w")
        for n, link in enumerate(self._links):
            row = ttk.Frame(self.f_links_list)
            row.pack(fill="x")
            a = ttk.Label(row, text=link.get("title") or link["url"],
                          style="Link.TLabel", cursor="hand2")
            a.pack(side="left")
            a.bind("<Button-1>", lambda e, u=link["url"]: self._open_link(u))
            a.bind("<Enter>", lambda e, u=link["url"]: self._status(u))
            for text, cmd in (("\u2715", self._remove_link), ("edit", self._edit_link)):
                act = ttk.Label(row, text=text, style="LinkAction.TLabel", cursor="hand2")
                act.pack(side="right", padx=(6, 0))
                act.bind("<Button-1>", lambda e, n=n, f=cmd: f(n))
            if link.get("description"):
                ttk.Label(self.f_links_list, text=link["description"], foreground="#666",
                          font=("", 8), wraplength=280, justify="left").pack(
                    anchor="w", padx=(10, 0), pady=(0, 2))

    def _open_link(self, url):
        try:
            webbrowser.open(url)
            self._status(f"Opened {url}")
        except Exception as exc:  # noqa: BLE001 - surface any launcher failure
            self._status(f"Couldn't open {url}: {exc}")

    def _reset_link_entries(self):
        for entry in (self.f_link_title, self.f_link_desc, self.f_link_url):
            entry.delete(0, "end")
        self._link_edit_index = None
        self.f_link_btn.configure(text="Add link")

    def _add_link(self):
        url = self.f_link_url.get().strip()
        if not url:
            self._status("Enter the link (URL) first.")
            self.f_link_url.focus_set()
            return
        url = normalize_url(url)
        link = {
            "title": self.f_link_title.get().strip() or url,
            "description": self.f_link_desc.get().strip(),
            "url": url,
        }
        if self._link_edit_index is not None:
            self._links[self._link_edit_index] = link
        else:
            self._links.append(link)
        self._reset_link_entries()
        self._render_links()
        self._persist_links(f"Link '{link['title']}' saved.")

    def _edit_link(self, n):
        link = self._links[n]
        self._reset_link_entries()
        self.f_link_title.insert(0, link.get("title", ""))
        self.f_link_desc.insert(0, link.get("description", ""))
        self.f_link_url.insert(0, link["url"])
        self._link_edit_index = n
        self.f_link_btn.configure(text="Update")
        self.f_link_title.focus_set()

    def _remove_link(self, n):
        link = self._links.pop(n)
        self._reset_link_entries()
        self._render_links()
        self._persist_links(f"Removed link '{link.get('title', '')}'.")

    def _persist_links(self, msg):
        """Links save immediately on an existing task; new tasks keep them until Save."""
        if self.active_project is not None and self.editing_index is not None:
            task = self.store.projects[self.active_project][self.editing_index]
            task["links"] = [dict(link) for link in self._links]
            self.store.save()
            if self.view == "Visual Planner":
                self._refresh_planner()
            self._status(msg)
        else:
            self._status("Link added - save the task to keep it.")

    def _text_row(self, form, r, label, height, weight):
        ttk.Label(form, text=label).grid(row=r, column=0, sticky="ne", padx=4, pady=3)
        w = tk.Text(form, width=30, height=height, wrap="word")
        w.grid(row=r, column=1, columnspan=3, sticky="nsew", padx=4, pady=3)
        form.rowconfigure(r, weight=weight)
        return w

    # -- form helpers -----------------------------------------------------
    def _toggle_jira(self):
        state = "normal" if self.f_jira_made.get() else "disabled"
        self.f_jira_ref.configure(state=state)

    def _update_status_pill(self):
        status = self.f_status.get()
        color = STATUSES.get(status, STATUSES[DEFAULT_STATUS])[0]
        self.status_pill.configure(text=status, bg=color)

    def _on_status_change(self):
        self._update_status_pill()
        # Preview the completed stamp; it's persisted on save.
        if self.f_status.get() == DONE:
            stamp = self._completed_date or "(on save)"
        else:
            stamp = "-"
        self.f_completed_lbl.configure(text=f"Completed: {stamp}")

    def _set_form_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        for w in (self.f_title, self.f_due, self.f_desc, self.f_blockers, self.f_notes):
            w.configure(state=state)
        for w in (self.status_cb, self.f_group, self.f_epic, self.f_sprint):
            w.configure(state="readonly" if enabled else "disabled")
        self.jira_chk.state(["!disabled"] if enabled else ["disabled"])
        for w in (self.f_link_title, self.f_link_desc, self.f_link_url, self.f_link_btn):
            w.state(["!disabled"] if enabled else ["disabled"])
        if enabled:
            self._toggle_jira()
        else:
            self.f_jira_ref.configure(state="disabled")
        self.save_btn.state(["!disabled"] if enabled else ["disabled"])

    def _set_text(self, widget, value):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", value)

    def _set_entry(self, widget, value):
        widget.delete(0, "end")
        widget.insert(0, value)

    def _refresh_choices(self):
        """Epic/sprint dropdowns offer the project's lists (blank = none)."""
        p = self.active_project
        for kind, combo in self.f_kinds.items():
            combo.configure(values=[""] + self.store.names(p, kind))

    def _clear_form(self):
        self._load_task_into_form({})
        self.f_added_lbl.configure(text="Created: (on save)")

    def _load_task_into_form(self, task):
        self._set_form_enabled(True)
        self._refresh_choices()
        self._set_entry(self.f_title, task.get("title", ""))
        self._set_entry(self.f_due, task.get("due_date", ""))
        for kind, combo in self.f_kinds.items():
            combo.set(task.get(kind, ""))
        self._set_text(self.f_desc, task.get("description", ""))
        self._set_text(self.f_blockers, task.get("blockers", ""))
        self._set_text(self.f_notes, task.get("notes", ""))
        self.f_jira_made.set(task.get("jira_made", False))
        self.f_jira_ref.configure(state="normal")
        self._set_entry(self.f_jira_ref, task.get("jira_ref", ""))
        self._toggle_jira()
        self.f_status.set(task.get("status", DEFAULT_STATUS))
        self._added_date = task.get("added_date", "")
        self._completed_date = task.get("completed_date", "")
        self.f_added_lbl.configure(text=f"Created: {self._added_date or '-'}")
        self.f_completed_lbl.configure(text=f"Completed: {self._completed_date or '-'}")
        self._update_status_pill()
        self._links = [dict(link) for link in task.get("links", [])]
        self._reset_link_entries()
        self._render_links()

    # -- views ------------------------------------------------------------
    def _refresh_views(self):
        self._refresh_tree()
        self._refresh_table()
        self._refresh_groups()
        if self.view == "Visual Planner":
            self._refresh_planner()
        self._sync_selection()

    def _refresh_tree(self):
        # Keep collapsed projects/groups collapsed across rebuilds.
        closed = {meta for item, meta in self.node_meta.items()
                  if meta[0] != "task" and not self.tree.item(item, "open")}
        self.tree.delete(*self.tree.get_children())
        self.node_meta = {}
        for project in sorted(self.store.projects, key=str.lower):
            tasks = self.store.projects[project]
            open_count = sum(1 for t in tasks if t.get("status") != DONE)
            pid = self.tree.insert(
                "", "end", text=f"\U0001F4C1 {project}  ({open_count}/{len(tasks)})",
                open=("project", project) not in closed, tags=("project",)
            )
            self.node_meta[pid] = ("project", project)
            # Groups first, then tasks not in a group directly under the project.
            group_nodes = {}
            for name in self.store.names(project, "group"):
                mine = [t for t in tasks if t.get("group") == name]
                open_n = sum(1 for t in mine if t.get("status") != DONE)
                gid = self.tree.insert(
                    pid, "end", text=f"\U0001F4C2 {name}  ({open_n}/{len(mine)})",
                    open=("group", project, name) not in closed, tags=("group",),
                )
                self.node_meta[gid] = ("group", project, name)
                group_nodes[name] = gid
            for i, task in enumerate(tasks):
                status = task.get("status", DEFAULT_STATUS)
                mark = "✓ " if status == DONE else "● "
                tid = self.tree.insert(
                    group_nodes.get(task.get("group"), pid), "end",
                    text=mark + (task.get("title") or "(untitled)"),
                    tags=(STATUSES[status][1],),
                )
                self.node_meta[tid] = ("task", project, i)

    def _tree_press(self, event):
        row = self.tree.identify_row(event.y)
        if not row:  # empty space: back to "new project"
            self._clear_tree_selection()
            return
        meta = self.node_meta.get(row)
        self._tdrag = None
        if meta and meta[0] == "task":
            self._tdrag = {"key": meta[1:], "x": event.x, "y": event.y,
                           "active": False, "target": None}

    def _tree_mark(self, item, on):
        if item and self.tree.exists(item):
            tags = [t for t in self.tree.item(item, "tags") if t != "drop"]
            self.tree.item(item, tags=tags + (["drop"] if on else []))

    def _tree_motion(self, event):
        d = self._tdrag
        if not d:
            return None
        if not d["active"]:
            if abs(event.x - d["x"]) + abs(event.y - d["y"]) < 6:
                return None
            d["active"] = True
            self.tree.configure(cursor="fleur")
        if event.y < 12:
            self.tree.yview_scroll(-1, "units")
        elif event.y > self.tree.winfo_height() - 12:
            self.tree.yview_scroll(1, "units")
        target = self.tree.identify_row(event.y)
        if target != d["target"]:
            self._tree_mark(d["target"], False)
            self._tree_mark(target, True)
            d["target"] = target
        return "break"

    def _tree_release(self, event):
        d, self._tdrag = self._tdrag, None
        if not d or not d["active"]:
            return None
        self.tree.configure(cursor="")
        self._tree_mark(d["target"], False)
        meta = self.node_meta.get(d["target"])
        if not meta:
            return "break"
        if meta[0] == "project":
            dest, group = meta[1], ""          # out of any group
        elif meta[0] == "group":
            dest, group = meta[1], meta[2]
        else:                                  # onto a task: join its group
            dest = meta[1]
            group = self.store.projects[dest][meta[2]].get("group", "")
        project, index = d["key"]
        task = dict(self.store.projects[project][index])
        if (dest, group) == (project, task.get("group", "")):
            return "break"
        task["group"] = group
        self._relocate_task(project, index, task, dest)
        self._refresh_views()
        where = f"subgroup '{group}'" if group else "no subgroup"
        self._status(f"Moved '{task.get('title', '')}' to {dest} / {where}.")
        return "break"

    def _sort_key(self, task):
        col = self.sort_col
        if col == "status":
            primary = STATUS_ORDER.index(task.get("status", DEFAULT_STATUS))
            return (primary, task.get("due_date") or "~")
        v = (task.get(col) or "").lower()
        return (v == "", v)  # blanks last

    def _sort_by(self, col):
        if self.sort_col == col:
            self.sort_rev = not self.sort_rev
        else:
            self.sort_col, self.sort_rev = col, False
        for key, heading in [("title", "Title")] + [c[:2] for c in TABLE_COLS]:
            arrow = (" ▼" if self.sort_rev else " ▲") if key == col else ""
            self.table.heading("#0" if key == "title" else key, text=heading + arrow)
        self._refresh_table()
        self._sync_selection()

    def _refresh_table(self):
        self.table.delete(*self.table.get_children())
        self.row_meta = {}
        project = self.active_project
        if not project or project not in self.store.projects:
            self.context_var.set("No project selected.")
            return

        tasks = self.store.projects[project]
        open_count = sum(1 for t in tasks if t.get("status") != DONE)
        self.context_var.set(f"{project}  -  {open_count} open / {len(tasks)} total")

        items = [(i, t) for i, t in enumerate(tasks)
                 if not (self.hide_done.get() and t.get("status") == DONE)]
        items.sort(key=lambda it: self._sort_key(it[1]), reverse=self.sort_rev)

        group = self.group_var.get()
        if group == "None":
            for i, t in items:
                self._insert_row("", i, t)
            return

        key = kind_of(group)
        groups = {}
        for i, t in items:
            groups.setdefault(t.get(key) or "", []).append((i, t))
        if key == "status":
            names = [s for s in STATUS_ORDER if s in groups]
        else:
            # The project's list order, then anything not in the list.
            names = [g for g in self.store.names(project, key) if g in groups]
            names += sorted((g for g in groups if g and g not in names), key=str.lower)
            if "" in groups:
                names.append("")
        for name in names:
            label = name or f"(no {group.lower()})"
            gid = self.table.insert(
                "", "end", text=f"{label}  ({len(groups[name])})",
                open=True, tags=("group",),
            )
            for i, t in groups[name]:
                self._insert_row(gid, i, t)

    def _insert_row(self, parent, index, task):
        status = task.get("status", DEFAULT_STATUS)
        values = []
        for key, *_ in TABLE_COLS:
            v = task.get(key, "")
            if key == "status":
                v = "● " + v
            values.append(v)
        rid = self.table.insert(
            parent, "end", text=task.get("title") or "(untitled)",
            values=values, tags=(STATUSES[status][1],),
        )
        self.row_meta[rid] = index

    def _sync_selection(self):
        """Highlight the current project/task in both tree and table."""
        self._syncing = True
        try:
            node = None
            if self.active_project is not None:
                if self.editing_index is not None:
                    node = self._find_task_node(self.active_project, self.editing_index)
                if node is None:
                    # Starting a task from a group row: keep that group selected.
                    sel = self.tree.selection()
                    cur = self.node_meta.get(sel[0]) if sel else None
                    if cur and cur[0] == "group" and cur[1] == self.active_project:
                        node = sel[0]
                if node is None:
                    node = self._find_project_node(self.active_project)
            if node:
                self.tree.selection_set(node)
                self.tree.see(node)
            else:
                self.tree.selection_remove(self.tree.selection())

            row = None
            for rid, idx in self.row_meta.items():
                if idx == self.editing_index:
                    row = rid
                    break
            if row:
                self.table.selection_set(row)
                self.table.see(row)
            else:
                self.table.selection_remove(self.table.selection())
            self._highlight_cards()
            self._update_add_mode()
        finally:
            # Selection events are delivered after this returns.
            self.after_idle(self._end_sync)

    def _end_sync(self):
        self._syncing = False

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

    # -- epics & sprints tab ----------------------------------------------
    def _refresh_groups(self):
        p = self.active_project
        self.g_project.configure(values=sorted(self.store.projects, key=str.lower))
        self.g_project.set(p or "")
        tasks = self.store.projects.get(p, [])
        for kind, (lst, _) in self.g_widgets.items():
            lst.delete(*lst.get_children())
            for name in self.store.names(p, kind):
                mine = [t for t in tasks if t.get(kind) == name]
                open_n = sum(1 for t in mine if t.get("status") != DONE)
                lst.insert("", "end", iid=name, text=name, values=(len(mine), open_n))

    def _on_group_project(self):
        project = self.g_project.get()
        if project and project != self.active_project:
            self._select_project(project)

    def _on_group_select(self, kind):
        lst, entry = self.g_widgets[kind]
        sel = lst.selection()
        if sel:
            entry.delete(0, "end")
            entry.insert(0, sel[0])

    def _group_changed(self, kind, old, new):
        """Refresh views after a group/epic/sprint was renamed or deleted."""
        combo = self.f_kinds[kind]
        current = combo.get()
        self._refresh_choices()
        if old is not None and current == old:
            combo.set(new)
        self._refresh_tree()
        self._refresh_table()
        self._refresh_groups()
        self._sync_selection()

    def _group_add(self, kind):
        lst, entry = self.g_widgets[kind]
        name = entry.get().strip()
        if not self.active_project:
            self._status("Select a project first.")
            return
        if not name:
            return
        if not self.store.add_name(self.active_project, kind, name):
            self._status(f"{KIND_LABEL[kind]} '{name}' already exists.")
            return
        entry.delete(0, "end")
        self._group_changed(kind, None, None)
        lst.selection_set(name)
        self._status(f"Added {KIND_LABEL[kind].lower()} '{name}' to '{self.active_project}'.")

    def _group_rename(self, kind):
        lst, entry = self.g_widgets[kind]
        sel = lst.selection()
        new = entry.get().strip()
        if not sel:
            self._status(f"Select the {KIND_LABEL[kind].lower()} to rename, edit its name, then Rename.")
            return
        old = sel[0]
        if not new or new == old:
            return
        if not self.store.rename_name(self.active_project, kind, old, new):
            self._status(f"{KIND_LABEL[kind]} '{new}' already exists.")
            return
        self._group_changed(kind, old, new)
        lst.selection_set(new)
        self._status(f"Renamed {KIND_LABEL[kind].lower()} '{old}' to '{new}'.")

    def _group_delete(self, kind):
        lst, entry = self.g_widgets[kind]
        sel = lst.selection()
        if not sel:
            self._status(f"Select the {KIND_LABEL[kind].lower()} to delete.")
            return
        name = sel[0]
        used = self.store.usage(self.active_project, kind, name)
        msg = f"Delete {KIND_LABEL[kind].lower()} '{name}'?"
        if used:
            msg += f"\n\n{used} task(s) will be left with no {KIND_LABEL[kind].lower()}."
        if not messagebox.askyesno(f"Delete {KIND_LABEL[kind].lower()}", msg):
            return
        self.store.delete_name(self.active_project, kind, name)
        entry.delete(0, "end")
        self._group_changed(kind, name, "")
        self._status(f"Deleted {KIND_LABEL[kind].lower()} '{name}'.")

    def _select_project(self, project):
        changed = project != self.active_project
        self.active_project = project
        self.editing_index = None
        self._clear_form()
        if changed:
            self._refresh_table()
            self._refresh_groups()
        self._sync_selection()

    def _select_task(self, project, index):
        changed = project != self.active_project
        self.active_project = project
        self.editing_index = index
        self._load_task_into_form(self.store.projects[project][index])
        if changed:
            self._refresh_table()
            self._refresh_groups()
        self._sync_selection()

    def _on_tree_select(self):
        if self._syncing:
            return
        sel = self.tree.selection()
        meta = self.node_meta.get(sel[0]) if sel else None
        if not meta:
            return
        if meta[0] == "project":
            self._select_project(meta[1])
            self._status(f"Project '{meta[1]}' - click 'New Task' to add one.")
        elif meta[0] == "group":
            self._select_project(meta[1])
            self.f_group.set(meta[2])  # a new task starts in this group
            self._status(f"Subgroup '{meta[2]}' - 'New Task' adds to it; drag tasks here.")
        else:
            self._select_task(meta[1], meta[2])

    def _on_table_select(self):
        if self._syncing:
            return
        sel = self.table.selection()
        if sel and sel[0] in self.row_meta:
            self._select_task(self.active_project, self.row_meta[sel[0]])

    # -- project actions --------------------------------------------------
    def _add_project(self):
        name = self.project_entry.get().strip()
        if not name:
            return
        if not self.store.add_project(name):
            self._status(f"Project '{name}' already exists.")
            return
        self.project_entry.delete(0, "end")
        self.active_project = name
        self.editing_index = None
        self._clear_form()
        self._refresh_views()
        self._status(f"Added project '{name}'. Now click 'New Task' to add tasks.")

    def _add_target(self):
        """Project the add box creates a subgroup in (None = new project)."""
        sel = self.tree.selection()
        meta = self.node_meta.get(sel[0]) if sel else None
        return meta[1] if meta else None

    def _update_add_mode(self):
        project = self._add_target()
        if project:
            self.add_mode_lbl.configure(text=f"New subgroup in '{project}'")
            self.add_btn.configure(text="+ Subgroup")
            self.add_switch.pack(side="right")
        else:
            self.add_mode_lbl.configure(text="New project")
            self.add_btn.configure(text="+ Project")
            self.add_switch.pack_forget()

    def _clear_tree_selection(self):
        self._syncing = True  # deselect only; keep the form as it is
        self.tree.selection_remove(self.tree.selection())
        self.after_idle(self._end_sync)
        self._update_add_mode()
        self.project_entry.focus_set()

    def _add_from_box(self):
        if self._add_target():
            self._add_group()
        else:
            self._add_project()

    def _add_group(self):
        name = self.project_entry.get().strip()
        project = self._add_target()
        if not project:
            return
        if not name:
            self._status(f"Type a subgroup name, then '+ Subgroup' to add it to '{project}'.")
            self.project_entry.focus_set()
            return
        if not self.store.add_name(project, "group", name):
            self._status(f"Subgroup '{name}' already exists in '{project}'.")
            return
        self.active_project = project
        self.project_entry.delete(0, "end")
        # Same as clicking the new group: form ready for a task in it.
        self.editing_index = None
        self._clear_form()
        self.f_group.set(name)
        self._refresh_views()
        for item, meta in self.node_meta.items():
            if meta == ("group", project, name):
                self._syncing = True  # selection only; don't reload the form
                self.tree.selection_set(item)
                self.tree.see(item)
                self.after_idle(self._end_sync)
        self._update_add_mode()
        self._status(f"Added subgroup '{name}' to '{project}'. Drag tasks onto it.")

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
            self._refresh_views()
            self._status(f"Deleted project '{name}'.")

    # -- task actions -----------------------------------------------------
    def _new_task(self):
        if not self.active_project:
            self._status("Select a project in the list first, then click New Task.")
            return
        # Start in the group that's selected in the sidebar (or the selected
        # task's group), so new tasks land next to related ones.
        group = ""
        if self.view == "Manager":
            sel = self.tree.selection()
            meta = self.node_meta.get(sel[0]) if sel else None
            if meta and meta[1] == self.active_project:
                if meta[0] == "group":
                    group = meta[2]
                elif meta[0] == "task":
                    group = self.store.projects[meta[1]][meta[2]].get("group", "")
        self.editing_index = None
        self._clear_form()
        self.f_group.set(group)
        self._sync_selection()
        self.f_title.focus_set()
        where = f"{self.active_project} / {group}" if group else self.active_project
        self._status(f"Entering new task under '{where}'. Fill in + Save.")

    def _save_task(self):
        if not self.active_project or self.save_btn.instate(["disabled"]):
            self._status("Select a project first.")
            return
        project = self.active_project
        title = self.f_title.get().strip()
        if not title:
            self._status("Title is required.")
            self.f_title.focus_set()
            return

        status = self.f_status.get()
        completed_date = (self._completed_date or now_stamp()) if status == DONE else ""

        task = {
            "title": title,
            "status": status,
            "group": self.f_group.get().strip(),
            "epic": self.f_epic.get().strip(),
            "sprint": self.f_sprint.get().strip(),
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
            "links": [dict(link) for link in self._links],
        }

        if self.editing_index is None:
            self.store.add_task(project, task)
            self.editing_index = len(self.store.projects[project]) - 1
            self._status(f"Added task '{title}' to project '{project}'.")
        else:
            # Keep fields the form doesn't show (e.g. planner rank).
            task = {**self.store.projects[project][self.editing_index], **task}
            self.store.update_task(project, self.editing_index, task)
            self._status(f"Saved task '{title}'.")

        self._load_task_into_form(task)
        self._refresh_views()

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
            self._refresh_views()
            self._status(f"Deleted task '{title}'.")

    # -- export -----------------------------------------------------------
    def _rows_for_export(self):
        headers = ["Project"] + [label for _, label in TASK_FIELDS]
        rows = []
        for project in sorted(self.store.projects, key=str.lower):
            for task in self.store.projects[project]:
                row = [project]
                for key, _ in TASK_FIELDS:
                    v = task.get(key, "")
                    if key == "links":
                        v = "\n".join(f"{link.get('title', '')}: {link['url']}"
                                      for link in v or [])
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
        status_col = headers.index("Status") + 1
        for row in rows:
            ws.append(row)
            cell = ws.cell(row=ws.max_row, column=status_col)
            color = STATUSES.get(cell.value, (None,))[0]
            if color:
                cell.fill = PatternFill("solid", fgColor=color.lstrip("#").upper())
                cell.font = Font(bold=True, color="FFFFFF")
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
