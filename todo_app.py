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

import base64
import csv
import html
import json
import os
import re
import ssl
import sys
import threading
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
import uuid
import webbrowser
import xml.etree.ElementTree as ET
from datetime import datetime
from email.utils import parsedate_to_datetime
from html.parser import HTMLParser
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
# Everything personal lives in profile/ (gitignored): tasks, settings and
# the custom field list. Older versions kept files next to the app; they're
# moved in on first start (see init_profile).
PROFILE_DIR = os.path.join(APP_DIR, "profile")
DATA_FILE = os.path.join(PROFILE_DIR, "data.json")
# Per-machine window state (last view, window positions, pin).
SETTINGS_FILE = os.path.join(PROFILE_DIR, "settings.json")
# Optional Jira connection (address, auth type, token if remembered).
JIRA_FILE = os.path.join(PROFILE_DIR, "jira.json")
LEGACY_FILES = {
    DATA_FILE: os.path.join(APP_DIR, "todo-data.json"),
    SETTINGS_FILE: os.path.join(APP_DIR, "todo-settings.json"),
}


def init_profile():
    """Create profile/ and move files from older versions into it."""
    os.makedirs(PROFILE_DIR, exist_ok=True)
    for new, old in LEGACY_FILES.items():
        if os.path.exists(old) and not os.path.exists(new):
            os.replace(old, new)

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
    ("priority", "Priority"),
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
    ("shared_by", "Shared By"),  # set on tasks imported from a coworker
]

# Shared-task files (File > Share selected / Import shared tasks).
SHARE_FORMAT = "todo-tracker-share"
SHARE_VERSION = 1

# Center table columns: (key, heading, width, stretch). The title lives in
# the tree column (#0) so group headers and task titles share it.
TABLE_COLS = [
    ("status", "Status", 140, False),
    ("priority", "Priority", 85, False),
    ("epic", "Epic", 90, False),
    ("sprint", "Sprint", 80, False),
    ("due_date", "Due", 80, False),
    ("added_date", "Created", 110, False),
]

# "Project" shows every project's tasks; the others show the selected project.
GROUP_OPTIONS = ["None", "Project", "Subgroup", "Epic", "Sprint", "Status", "Priority"]

# Priority levels (Jira's), highest first: name -> (color, arrow glyph).
PRIORITIES = {
    "Highest": ("#cd1317", "\u21c8"),   # double up arrow
    "High": ("#e9494a", "\u2191"),
    "Medium": ("#e97f33", "="),
    "Low": ("#2d8738", "\u2193"),
    "Lowest": ("#57a55a", "\u21ca"),    # double down arrow
}
PRIORITY_ORDER = list(PRIORITIES)
# Sort table / sort dropdown labels -> task keys.
SORT_OPTIONS = {"Status": "status", "Priority": "priority", "Due date": "due_date",
                "Created": "added_date", "Title": "title", "Epic": "epic",
                "Sprint": "sprint"}
BUDDY_SORTS = ["Activity", "Priority", "Due date", "Title", "Created"]


def priority_rank(task):
    """0 = Highest ... 4 = Lowest; tasks without a priority sort last."""
    p = task.get("priority", "")
    return PRIORITY_ORDER.index(p) if p in PRIORITIES else len(PRIORITY_ORDER)


def priority_label(p):
    return f"{PRIORITIES[p][1]} {p}" if p in PRIORITIES else ""


JIRA_PRIORITY = {
    "highest": "Highest", "blocker": "Highest", "p1": "Highest",
    "high": "High", "critical": "High", "p2": "High",
    "medium": "Medium", "major": "Medium", "normal": "Medium", "p3": "Medium",
    "low": "Low", "minor": "Low", "p4": "Low",
    "lowest": "Lowest", "trivial": "Lowest", "p5": "Lowest",
}


def jira_priority(name):
    """Map a Jira priority name (Cloud or Server) onto our five levels."""
    n = (name or "").strip().lower()
    if n in JIRA_PRIORITY:
        return JIRA_PRIORITY[n]
    m = re.match(r"^(p[1-5])\b", n)  # e.g. "P2 - High"
    return JIRA_PRIORITY.get(m.group(1), "") if m else ""

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
# New tasks from Jira imports land here; drag them to your own projects.
IMPORTED_PROJECT = "Imported"
# Two arrows chasing each other = synced from Jira. Tk draws the emoji as a
# tiny box on Windows, so use the plain symbol there (and a safe one on Linux).
JIRA_ICON = ("\U0001F5D8" if sys.platform == "win32" else
             "\U0001F504" if sys.platform == "darwin" else "\u27F3")


def jira_mark(task):
    """Suffix shown after a task's title when it has a Jira ticket attached."""
    return f"  {JIRA_ICON}" if task.get("jira") else ""
AUTOSAVE_DELAY = 500  # ms after the last keystroke before saving
ALL_SPRINTS = "All sprints"
NO_SPRINT = "(no sprint)"

# Visual Planner (Trello-style board).
PLANNER_COLUMNS = ["Status", "Priority", "Subgroup", "Epic", "Sprint", "Project"]
PLANNER_ORDER = ["Manual", "Priority", "Due date", "Created", "Title"]
ANY = "All"
NONE_LABEL = "(none)"
BOARD_BG = "#e4e9f0"
LIST_BG = "#d5dce6"
CARD_BG = "#ffffff"
CARD_BORDER = "#c4ccd6"
ACCENT = "#2b7de9"  # selected card, drop target
LIST_WIDTH = 250


# ---------------------------------------------------------------------------
# Cross-platform input: macOS uses Cmd for shortcuts, Button-2 for right
# click and small wheel deltas; Linux (X11) sends wheel as Button-4/5.
# ---------------------------------------------------------------------------
IS_MAC = sys.platform == "darwin"
MOD = "Command" if IS_MAC else "Control"   # shortcut modifier
MOD_LABEL = "Cmd" if IS_MAC else "Ctrl"


def wheel_steps(event):
    """Scroll steps (+down / -up) from a wheel event on any platform."""
    if getattr(event, "num", None) == 4:
        return -1
    if getattr(event, "num", None) == 5:
        return 1
    delta = getattr(event, "delta", 0) or 0
    if IS_MAC:  # small values, one per notch/trackpad tick
        return -delta
    return int(-delta / 120) or (-1 if delta > 0 else 1)


def bind_wheel(widget, scroll, shift_scroll=None):
    """Bind vertical (and optional Shift = horizontal) wheel scrolling."""
    widget.bind("<MouseWheel>", lambda e: scroll(wheel_steps(e)))
    widget.bind("<Button-4>", lambda e: scroll(wheel_steps(e)))
    widget.bind("<Button-5>", lambda e: scroll(wheel_steps(e)))
    if shift_scroll:
        widget.bind("<Shift-MouseWheel>", lambda e: shift_scroll(wheel_steps(e)))
        widget.bind("<Shift-Button-4>", lambda e: shift_scroll(wheel_steps(e)))
        widget.bind("<Shift-Button-5>", lambda e: shift_scroll(wheel_steps(e)))


def bind_right_click(widget, callback):
    """Right-click: Button-3 on Windows/Linux; Button-2 or Ctrl-click on macOS."""
    if IS_MAC:
        widget.bind("<Button-2>", callback)
        widget.bind("<Control-Button-1>", callback)
    else:
        widget.bind("<Button-3>", callback)


def now_stamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def new_task(title):
    """A task with every field present and defaults filled in."""
    task = {key: "" for key, _ in TASK_FIELDS}
    task.update(title=title, status=DEFAULT_STATUS, added_date=now_stamp(),
                jira_made=False, links=[], fields={}, priority="", id=uuid.uuid4().hex)
    return task


def normalize_url(url):
    """Add https:// to bare addresses like 'example.com/page'."""
    if (re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", url)      # https://, ftp://
            or re.match(r"^(mailto|tel|file):", url, re.I)
            or re.match(r"^[a-zA-Z]:[\\/]", url)               # C:\path
            or url.startswith("\\\\")):                         # \\server\share
        return url
    return "https://" + url


# ---------------------------------------------------------------------------
# Jira XML import (Jira's "Export > XML" / RSS format, one or many issues)
# ---------------------------------------------------------------------------
class _HTMLText(HTMLParser):
    """Jira descriptions/comments are HTML; flatten them to readable text."""

    BLOCKS = {"p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6",
              "pre", "blockquote", "table", "ul", "ol"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.links = []   # [(href, link text)]
        self._href = None
        self._href_text = []

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            href = dict(attrs).get("href") or ""
            if href and not href.startswith("#"):
                self._href, self._href_text = href, []
        if tag == "li":
            self.parts.append("\n- ")
        elif tag in self.BLOCKS:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag == "a" and self._href:
            self.links.append((self._href, "".join(self._href_text).strip()))
            self._href = None
        if tag in self.BLOCKS and tag != "li":
            self.parts.append("\n")

    def handle_data(self, data):
        self.parts.append(data)
        if self._href:
            self._href_text.append(data)


def html_to_text(markup):
    if not markup:
        return ""
    parser = _HTMLText()
    parser.feed(markup)
    text = html.unescape("".join(parser.parts)).replace("\xa0", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def html_links(markup):
    """[(href, text)] for every <a href> in some HTML."""
    if not markup:
        return []
    parser = _HTMLText()
    parser.feed(markup)
    return parser.links


def jira_date(value, with_time=False):
    """'Fri, 13 Oct 2023 09:30:00 -0400' -> '2023-10-13' (or with HH:MM)."""
    if not value or not value.strip():
        return ""
    try:
        dt = parsedate_to_datetime(value.strip())
    except (TypeError, ValueError):
        return value.strip()
    return dt.strftime("%Y-%m-%d %H:%M" if with_time else "%Y-%m-%d")


def jira_status(name, category):
    """Map a Jira status (+ its category) onto this app's statuses."""
    n = (name or "").lower()
    if "block" in n or "impediment" in n or "on hold" in n:
        return "Blocked"
    if any(w in n for w in ("approv", "review", "waiting", "pending", "verif", "test")):
        return "Waiting for approval"
    if category == "done" or n in ("done", "closed", "resolved", "complete", "completed"):
        return DONE
    if category == "indeterminate" or "progress" in n:
        return "In Progress"
    return DEFAULT_STATUS


def parse_jira_issues(text):
    """Parse a Jira XML export into issue snapshots (plain dicts).

    A snapshot is a read-only copy of the ticket that gets attached to a task
    (task["jira"]) and shown in the form's Jira tab - it never overwrites
    the task's own fields."""
    root = ET.fromstring(text)
    items = root.findall(".//item")
    if not items:
        raise ValueError("No Jira issues (<item>) found in that XML.")

    issues = []
    for item in items:
        def get(tag):
            el = item.find(tag)
            return (el.text or "").strip() if el is not None and el.text else ""

        def all_text(path):
            return [(e.text or "").strip() for e in item.findall(path) if (e.text or "").strip()]

        key = get("key")
        cat_el = item.find("statusCategory")
        if cat_el is None:
            cat_el = item.find("statuscategory")

        custom = {}
        for cf in item.findall("customfields/customfield"):
            name = (cf.findtext("customfieldname") or "").strip()
            values = [(v.text or "").strip() for v in cf.findall("customfieldvalues/customfieldvalue")
                      if (v.text or "").strip()]
            if name and values:
                custom[name] = values
        sprint = ""
        if custom.get("Sprint"):
            sprint = custom["Sprint"][-1]  # the latest sprint it was in
            m = re.search(r"name=([^,\]]+)", sprint)  # Jira Server's long form
            if m:
                sprint = m.group(1).strip()
        epic = (custom.get("Epic Link") or custom.get("Parent Link") or [""])[0]
        if not epic and item.find("parent") is not None and get("type").lower() not in ("sub-task", "subtask"):
            epic = get("parent")

        links = []
        for lt in item.findall("issuelinks/issuelinktype"):
            for direction in lt:
                desc = direction.get("description", "")
                for k in direction.findall("issuelink/issuekey"):
                    links.append(f"{desc} {(k.text or '').strip()}".strip())

        fields = {}
        for label, value in (
            ("Labels", ", ".join(all_text("labels/label"))),
            ("Components", ", ".join(all_text("component"))),
            ("Fix versions", ", ".join(all_text("fixVersion"))),
            ("Affects versions", ", ".join(all_text("version"))),
            ("Environment", html_to_text(get("environment"))),
            ("Parent", get("parent")),
            ("Subtasks", ", ".join(all_text("subtasks/subtask"))),
            ("Linked issues", "\n".join(links)),
        ):
            if value:
                fields[label] = value
        for name, values in custom.items():
            if name not in ("Sprint", "Epic Link", "Parent Link", "Rank"):
                fields[name] = ", ".join(values)

        # Every link in the XML: the ticket, linked issues, parent, epic,
        # subtasks, attachments and any <a href> in the description/comments.
        url = get("link")
        base = url.split("/browse/")[0] if "/browse/" in url else ""
        found = []

        def add_link(title, href, description=""):
            if href and href not in {l["url"] for l in found}:
                found.append({"title": title or href, "description": description, "url": href})

        add_link(key or "Jira ticket", url, "Jira ticket" + (f": {get('summary')}" if get("summary") else ""))
        if base:
            browse = base + "/browse/"
            if epic and re.match(r"^[A-Z][A-Z0-9_]*-\d+$", epic):
                add_link(f"Epic {epic}", browse + epic, "Jira epic")
            if get("parent"):
                add_link(f"Parent {get('parent')}", browse + get("parent"), "Jira parent issue")
            for lt in item.findall("issuelinks/issuelinktype"):
                for direction in lt:
                    desc = direction.get("description", "")
                    for k in direction.findall("issuelink/issuekey"):
                        other = (k.text or "").strip()
                        add_link(f"{desc} {other}".strip(), browse + other, "Linked Jira issue")
            for sub in all_text("subtasks/subtask"):
                add_link(f"Subtask {sub}", browse + sub, "Jira subtask")
            for att in item.findall("attachments/attachment"):
                name, att_id = att.get("name", ""), att.get("id", "")
                if att_id and name:
                    add_link(name, f"{base}/secure/attachment/{att_id}/{name}", "Jira attachment")
        for markup in [get("description")] + [c.text or "" for c in item.findall("comments/comment")]:
            for href, text in html_links(markup):
                if href.startswith("/") and base:
                    href = base + href
                add_link(text or href, href, "Link from the Jira ticket")

        project_el = item.find("project")
        project = (project_el.text or "").strip() if project_el is not None else ""
        issues.append({
            "key": key,
            "url": get("link"),
            "project": project or (key.split("-")[0] if "-" in key else "Jira"),
            "summary": get("summary") or re.sub(r"^\[[^\]]+\]\s*", "", get("title")) or key,
            "status": get("status"),
            "status_category": cat_el.get("key", "") if cat_el is not None else "",
            "type": get("type"),
            "priority": get("priority"),
            "priority_level": jira_priority(get("priority")),
            "resolution": get("resolution"),
            "assignee": get("assignee"),
            "reporter": get("reporter"),
            "created": jira_date(get("created"), True),
            "updated": jira_date(get("updated"), True),
            "due": jira_date(get("due")),
            "resolved": jira_date(get("resolved"), True),
            "sprint": sprint,
            "epic": epic,
            "blocked_by": [l for l in links if "blocked by" in l.lower()],
            "description": html_to_text(get("description")),
            "fields": fields,
            "comments": [{"author": c.get("author", ""),
                          "created": jira_date(c.get("created", ""), True),
                          "body": html_to_text(c.text or "")}
                         for c in item.findall("comments/comment")],
            "links": found,
            "attached": now_stamp(),
        })
    return issues


# Task field <- ticket value that a Jira ticket fills in on your side.
JIRA_SYNCED = (("epic", "epic"), ("sprint", "sprint"),
               ("description", "description"), ("due_date", "due"),
               ("jira_ref", "key"), ("priority", "priority_level"))


def jira_updates(task, issue, old_issue):
    """What a (re)attached ticket fills in on the task.

    Epic, sprint, description and due date: taken from the ticket only if
    your field is empty, or still holds what the previous copy of the ticket
    put there - anything you wrote yourself is kept. Resolved in Jira: the
    task becomes Done with Jira's resolved date as its completed date (only
    for a resolution not seen before, so reopening it yourself sticks). Links: every link in the
    ticket that isn't on the task yet (ones you removed after an earlier
    attach aren't re-added)."""
    old = old_issue or {}
    updates = {}
    for field, key in JIRA_SYNCED:
        value = (issue.get(key) or "").strip()
        current = (task.get(field) or "").strip()
        if value and value != current and (not current or current == (old.get(key) or "").strip()):
            updates[field] = value
    if issue.get("key") and not task.get("jira_made"):
        updates["jira_made"] = True  # "Ticket made" - it clearly exists
    resolved = issue.get("resolved", "")
    done_date = task.get("completed_date", "")
    if resolved and task.get("status") != DONE and resolved != old.get("resolved"):
        # Resolved in Jira -> Done here. Only for a resolution we haven't seen
        # yet, so reopening the task yourself sticks across refreshes.
        updates["status"] = DONE
        updates["completed_date"] = resolved
    elif (resolved and task.get("status") == DONE and resolved != done_date
            and (not done_date or done_date == old.get("resolved"))):
        updates["completed_date"] = resolved
    have = {link.get("url") for link in task.get("links") or []}
    offered_before = {link.get("url") for link in old.get("links") or []}
    new_links = [dict(link) for link in issue.get("links", [])
                 if link["url"] not in have and link["url"] not in offered_before]
    if new_links:
        updates["links"] = [dict(link) for link in task.get("links") or []] + new_links
    return updates


def jira_task(issue):
    """A new task from a Jira issue: title, epic, sprint, description, due
    date, links and Jira's created date; the full ticket rides along in
    task["jira"] (read-only, shown in the form's Jira tab)."""
    task = new_task(issue["summary"] or issue["key"] or "(untitled)")
    task.update(id=f"jira:{issue['key']}" if issue["key"] else task["id"], jira=issue)
    if issue.get("created"):
        task["added_date"] = issue["created"]
    task.update(jira_updates(task, issue, None))
    return task


def make_marker(master, status_color, priority):
    """Buddy row icon: status dot, then the priority arrow in its color."""
    w, h = 32, 16
    img = tk.PhotoImage(master=master, width=w, height=h)
    c, r = 7.5, 5.2
    for y in range(h):
        for x in range(16):
            if (x - c) ** 2 + (y - c) ** 2 <= r * r:
                img.put(status_color, (x, y))
    if priority in PRIORITIES:
        color = PRIORITIES[priority][0]

        def tri(top, up, height=4, cx=24):
            for i in range(height):
                half = i if up else height - 1 - i
                for x in range(cx - half - 1, cx + half + 2):
                    img.put(color, (x, top + i))

        if priority == "Highest":
            tri(2, True); tri(8, True)
        elif priority == "High":
            tri(5, True, 5)
        elif priority == "Medium":
            for y in (5, 6, 9, 10):
                for x in range(19, 30):
                    img.put(color, (x, y))
        elif priority == "Low":
            tri(5, False, 5)
        else:  # Lowest
            tri(2, False); tri(8, False)
    return img


# ---------------------------------------------------------------------------
# Optional Jira connection (read-only). Fetches the same XML you'd get from
# Jira's "Export > XML", so it reuses parse_jira_issues. Supports Jira Server
# / Data Center personal access tokens (Bearer) and Jira Cloud (email + API
# token). Settings live in profile/jira.json (gitignored).
# ---------------------------------------------------------------------------
JIRA_KEY_RE = re.compile(r"\b([A-Z][A-Z0-9_]+-\d+)\b")


def _dpapi(data, protect):
    """Windows: encrypt/decrypt bytes with the current user's login (DPAPI)."""
    import ctypes
    from ctypes import wintypes

    class Blob(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    buf = ctypes.create_string_buffer(data, len(data))
    blob_in = Blob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = Blob()
    crypt = ctypes.windll.crypt32
    fn = crypt.CryptProtectData if protect else crypt.CryptUnprotectData
    if not fn(ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)):
        raise OSError("Windows couldn't " + ("protect" if protect else "unlock") + " the token")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


def protect_token(token):
    if sys.platform == "win32":
        return "dpapi:" + base64.b64encode(_dpapi(token.encode("utf-8"), True)).decode("ascii")
    return "plain:" + token  # file is chmod 600 (owner-only) on macOS/Linux


def unprotect_token(stored):
    if not stored:
        return ""
    kind, _, value = stored.partition(":")
    if kind == "dpapi":
        return _dpapi(base64.b64decode(value), False).decode("utf-8")
    return value if kind == "plain" else ""


def load_jira_config():
    try:
        with open(JIRA_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        return cfg if isinstance(cfg, dict) else {}
    except (OSError, ValueError):
        return {}


def save_jira_config(cfg):
    os.makedirs(os.path.dirname(JIRA_FILE), exist_ok=True)
    with open(JIRA_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)
    if sys.platform != "win32":
        try:
            os.chmod(JIRA_FILE, 0o600)
        except OSError:
            pass


def jira_base_url(text):
    """Clean up whatever was typed: add https://, drop /browse/... and slashes."""
    url = normalize_url(text.strip()) if text.strip() else ""
    url = re.split(r"/(?:browse|secure|projects|issues|rest|si|sr)/", url, maxsplit=1)[0]
    return url.rstrip("/")


class JiraError(Exception):
    pass


class JiraClient:
    """Minimal read-only Jira client (standard library only)."""

    def __init__(self, base_url, auth, token, email="", ca_file="", timeout=20):
        self.base = jira_base_url(base_url)
        self.auth, self.token, self.email = auth, token, email
        self.ca_file, self.timeout = ca_file, timeout

    def _request(self, path, accept="application/xml"):
        if not self.base:
            raise JiraError("No Jira address set (Settings > Jira connection).")
        if not self.token:
            raise JiraError("No Jira token - enter it in Settings > Jira connection.")
        req = urllib.request.Request(self.base + path, headers={
            "Accept": accept, "User-Agent": "todo-tracker",
            "Authorization": (f"Bearer {self.token}" if self.auth == "pat" else
                              "Basic " + base64.b64encode(
                                  f"{self.email}:{self.token}".encode()).decode()),
        })
        ctx = ssl.create_default_context(cafile=self.ca_file or None)
        try:
            with urllib.request.urlopen(req, timeout=self.timeout, context=ctx) as resp:
                body = resp.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError as exc:
            hints = {
                401: "Jira rejected the token (401). Check it's correct and not expired.",
                403: "Jira refused access (403). Your token may lack permission, or "
                     "XML export may be disabled on this Jira.",
                404: "Not found (404). Check the ticket key and the Jira address.",
            }
            raise JiraError(hints.get(exc.code, f"Jira returned HTTP {exc.code}.")) from None
        except ssl.SSLError as exc:
            raise JiraError("SSL certificate problem talking to Jira. If your company "
                            "uses its own certificates, set the CA file in Settings > "
                            f"Jira connection. ({exc.reason or exc})") from None
        except urllib.error.URLError as exc:
            reason = getattr(exc, "reason", exc)
            if isinstance(reason, ssl.SSLError):
                raise JiraError("SSL certificate problem talking to Jira. If your company "
                                "uses its own certificates, set the CA file in Settings > "
                                "Jira connection.") from None
            raise JiraError(f"Couldn't reach Jira at {self.base} ({reason}).") from None
        except (TimeoutError, OSError) as exc:
            raise JiraError(f"Couldn't reach Jira at {self.base} ({exc}).") from None
        head = body.lstrip()[:200].lower()
        if accept.endswith("xml") and (head.startswith("<!doctype html") or "<html" in head):
            raise JiraError("Jira sent back a web page instead of data - usually a "
                            "login/SSO page, meaning the token wasn't accepted.")
        return body

    def myself(self):
        """Who the token belongs to - used by 'Test connection'."""
        data = json.loads(self._request("/rest/api/2/myself", accept="application/json"))
        return data.get("displayName") or data.get("name") or data.get("emailAddress") or "?"

    def issue_xml(self, key):
        key = urllib.parse.quote(key)
        return self._request(f"/si/jira.issueviews:issue-xml/{key}/{key}.xml")

    def search_xml(self, jql, limit=200):
        q = urllib.parse.urlencode({"jqlQuery": jql, "tempMax": limit})
        return self._request(f"/sr/jira.issueviews:searchrequest-xml/temp/SearchRequest.xml?{q}")


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

    def __init__(self, path, fields_path=None):
        self.path = path
        # Custom field definitions: [{"name", "type": "text"|"long", "source"}].
        # Created on the fly when an import (e.g. Jira) brings a field we
        # don't have yet; values live on each task under task["fields"].
        self.fields_path = fields_path or os.path.join(os.path.dirname(path), "fields.json")
        self.fields = []
        # {project_name: [task_dict, ...]}
        self.projects = {}
        # {"epic": {project_name: [name, ...]}, "sprint": {...}} - kept in
        # the order they were added (sprints are usually chronological).
        self.groups = {kind: {} for kind in GROUP_KINDS}
        # Epic details: {project: {epic name: {"sprint", "status",
        # "description", "jira_key", "jira_url"}}}
        self.epic_info = {}
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
                self.epic_info = data.get("epic_info", {})
            else:  # v1 file: the whole file was the projects dict
                self.projects = data
        try:
            with open(self.fields_path, "r", encoding="utf-8") as f:
                self.fields = [d for d in json.load(f) if isinstance(d, dict) and d.get("name")]
        except (OSError, ValueError):
            self.fields = []
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
                task.setdefault("priority", "")
                if not isinstance(task.get("fields"), dict):
                    task["fields"] = {}
                if not task.get("id"):  # stable identity for sharing
                    task["id"] = uuid.uuid4().hex
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
        data["epic_info"] = self.epic_info
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

    def save_fields(self):
        with open(self.fields_path, "w", encoding="utf-8") as f:
            json.dump(self.fields, f, indent=2)

    def field_type(self, name):
        for d in self.fields:
            if d["name"] == name:
                return d.get("type", "text")
        return "text"

    def ensure_field(self, name, ftype="text", source="user"):
        """Add a custom field definition if it's new. Returns True if added."""
        if any(d["name"] == name for d in self.fields):
            return False
        self.fields.append({"name": name, "type": ftype, "source": source})
        self.save_fields()
        return True

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
        self.epic_info.pop(name, None)
        self.save()

    def _next_rank(self):
        """Rank orders cards on the Visual Planner; new ones go last."""
        ranks = [t["rank"] for ts in self.projects.values() for t in ts
                 if isinstance(t.get("rank"), (int, float))]
        return max(ranks, default=-1) + 1

    def add_task(self, project, task):
        task.setdefault("rank", self._next_rank())
        if not task.get("id"):
            task["id"] = uuid.uuid4().hex
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
        info = self.epic_info.get(project, {})
        if kind == "epic" and old in info:
            info[new] = info.pop(old)
        self.save()
        return True

    def delete_name(self, project, kind, name):
        self.groups[kind][project].remove(name)
        if kind == "epic":
            self.epic_info.get(project, {}).pop(name, None)
        for t in self.projects[project]:
            if t.get(kind) == name:
                t[kind] = ""
        self.save()


    # -- epic details -------------------------------------------------------
    def epic_meta(self, project, name):
        return self.epic_info.get(project, {}).get(name, {})

    def set_epic_meta(self, project, name, **fields):
        self.epic_info.setdefault(project, {}).setdefault(name, {}).update(fields)
        self.save()

    def move_epic(self, src, name, dst):
        """Attach an epic to another project, taking its tasks with it.
        Returns the ids of the tasks that moved."""
        if src == dst:
            return []
        if name in self.groups["epic"].get(src, []):
            self.groups["epic"][src].remove(name)
        dst_names = self.groups["epic"].setdefault(dst, [])
        if name not in dst_names:
            dst_names.append(name)
        meta = self.epic_info.get(src, {}).pop(name, {})
        merged = {**meta, **self.epic_info.get(dst, {}).get(name, {})}
        if merged:
            self.epic_info.setdefault(dst, {})[name] = merged
        moving = [t for t in self.projects.get(src, []) if t.get("epic") == name]
        self.projects[src] = [t for t in self.projects.get(src, []) if t.get("epic") != name]
        for t in moving:
            t["group"] = ""  # subgroups belong to the old project
            sprints = self.groups["sprint"].setdefault(dst, [])
            if t.get("sprint") and t["sprint"] not in sprints:
                sprints.append(t["sprint"])
        if merged.get("sprint") and merged["sprint"] not in self.groups["sprint"].setdefault(dst, []):
            self.groups["sprint"][dst].append(merged["sprint"])
        self.projects.setdefault(dst, []).extend(moving)
        self.save()
        return [t.get("id") for t in moving]

    def note_jira_epics(self):
        """Give epics that came from Jira their key + link automatically."""
        changed = False
        for project, tasks in self.projects.items():
            for t in tasks:
                j = t.get("jira") or {}
                epic, url = j.get("epic", ""), j.get("url", "")
                if (epic and t.get("epic") == epic and "/browse/" in url
                        and JIRA_KEY_RE.fullmatch(epic)):
                    meta = self.epic_info.setdefault(project, {}).setdefault(epic, {})
                    if not meta.get("jira_key"):
                        meta["jira_key"] = epic
                        meta["jira_url"] = url.split("/browse/")[0] + "/browse/" + epic
                        changed = True
        if changed:
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

        self.bind_all(f"<{MOD}-s>", lambda e: self._on_ctrl_s())
        self.bind_all(f"<{MOD}-n>", lambda e: self._on_ctrl_n())
        for i, name in enumerate(VIEWS, start=1):
            self.bind_all(f"<{MOD}-Key-{i}>", lambda e, n=name: self._show_view(n))
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
                accelerator=f"{MOD_LABEL}+{i}", command=lambda n=name: self._show_view(n),
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
        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Share selected\u2026", command=self._share)
        file_menu.add_command(label="Import shared tasks\u2026", command=self._import_shared)
        file_menu.add_command(label="Import Jira XML\u2026", command=self._import_jira)
        file_menu.add_separator()
        file_menu.add_command(label="Import from Jira search\u2026", command=self._jira_search_import)
        file_menu.add_command(label="Refresh all Jira tickets", command=self._jira_refresh_all)
        file_menu.add_command(label=f"Gather Jira imports into '{IMPORTED_PROJECT}'",
                              command=self._gather_imports)
        menubar.add_cascade(label="File", menu=file_menu)
        menubar.add_cascade(label="View", menu=view_menu)
        settings_menu = tk.Menu(menubar, tearoff=False)
        settings_menu.add_command(label="Your name\u2026", command=lambda: self._ask_name())
        settings_menu.add_command(label="Jira connection\u2026", command=self._jira_settings)
        menubar.add_cascade(label="Settings", menu=settings_menu)
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
        self._flush_autosave()
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
        bind_wheel(widget, lambda n: self.board.yview_scroll(n, "units"),
                   lambda n: self.board.xview_scroll(n, "units"))
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
        if mode == "Priority":
            return [(p, priority_label(p)) for p in PRIORITY_ORDER] + [("", "No priority")]
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
        if order == "Priority":
            due = task.get("due_date", "")
            return (priority_rank(task), due == "", due)
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
        tk.Label(body, text=(task.get("title") or "(untitled)") + jira_mark(task), bg=CARD_BG,
                 font=("", 10), wraplength=LIST_WIDTH - 40, justify="left",
                 anchor="w").pack(fill="x")

        # Status, priority + due date on one line.
        due = task.get("due_date", "")
        if mode != "Status" or due or task.get("priority") in PRIORITIES:
            row = tk.Frame(body, bg=CARD_BG)
            row.pack(fill="x", pady=(3, 0))
            if mode != "Status":
                tk.Label(row, text="\u25cf " + status, bg=CARD_BG, fg=color,
                         font=("", 8, "bold")).pack(side="left")
            prio = task.get("priority", "")
            if prio in PRIORITIES and mode != "Priority":
                tk.Label(row, text=priority_label(prio), bg=CARD_BG, fg=PRIORITIES[prio][0],
                         font=("", 8, "bold")).pack(side="left", padx=(6, 0))
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
        if task.get("shared_by"):
            meta.append(f"from {task['shared_by']}")
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
        elif mode == "Priority":
            self.f_priority.set(value)
            self._update_priority_icon()
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
        elif mode == "Priority":
            task["priority"] = value
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
            self.f_priority.set(task.get("priority", ""))
            self._update_priority_icon()

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

        row2 = tk.Frame(f, bg=BUDDY_BG)
        row2.pack(fill="x", padx=6, pady=(0, 2))
        tk.Label(row2, text="Sort", bg=BUDDY_BG, fg="#665", font=("", 8)).pack(side="left")
        self.b_sort = ttk.Combobox(row2, state="readonly", width=10, values=BUDDY_SORTS)
        saved_sort = self.settings.get("buddy_sort", "Activity")
        self.b_sort.set(saved_sort if saved_sort in BUDDY_SORTS else "Activity")
        self.b_sort.pack(side="left", padx=(4, 0))
        self.b_sort.bind("<<ComboboxSelected>>", lambda e: self._refresh_buddy())
        self.b_hide_done = tk.BooleanVar(value=self.settings.get("buddy_hide_done", True))
        tk.Checkbutton(
            row2, text="Hide done", variable=self.b_hide_done, bg=BUDDY_BG,
            activebackground=BUDDY_BG, command=self._refresh_buddy,
        ).pack(side="right")

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
        self.b_tree.tag_configure("link", foreground="#1a5fb4", font=("", 8, "underline"))
        self.b_link_rows = {}  # link row id -> url
        self.b_tree.bind("<ButtonRelease-1>", self._buddy_link_click, add="+")
        # Status shows as a colored dot; text keeps the normal color.
        # Row icon = status dot + priority arrow (built on demand, cached).
        self.b_markers = {}
        self.b_tree.bind("<Double-1>", self._buddy_dblclick)
        self.b_tree.bind("<<TreeviewSelect>>", lambda e: self._buddy_update_add())
        self.b_tree.bind("<<TreeviewOpen>>", lambda e: self._buddy_fold(True))
        self.b_tree.bind("<<TreeviewClose>>", lambda e: self._buddy_fold(False))
        # project/group row id -> (project, group); group is "" for a project row
        self.b_folder_rows = {}
        self.b_tree.tag_configure("group", font=("", 9, "bold"), foreground="#554")
        bind_right_click(self.b_tree, self._buddy_menu)
        self.b_meta = {}  # row id -> (project, index)

        self.b_status_menu = tk.Menu(self, tearoff=False)

    def _buddy_marker(self, status, priority):
        key = (status, priority if priority in PRIORITIES else "")
        if key not in self.b_markers:
            self.b_markers[key] = make_marker(self, STATUSES[status][0], key[1])
        return self.b_markers[key]

    def _buddy_sort_key(self, task):
        sort = self.b_sort.get()
        status = BUDDY_ORDER.index(task.get("status", DEFAULT_STATUS))
        due = task.get("due_date") or "~"
        if sort == "Priority":
            return (priority_rank(task), status, due)
        if sort == "Due date":
            return (due, priority_rank(task))
        if sort == "Title":
            return ((task.get("title") or "").lower(),)
        if sort == "Created":
            return (task.get("added_date", ""),)
        return (status, priority_rank(task), due)  # Activity

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
        self.settings["buddy_sort"] = self.b_sort.get()
        # Collapsed folders: "project" or "project/group".
        collapsed = set(self.settings.get("buddy_collapsed", []))
        today = datetime.now().strftime("%Y-%m-%d")

        selected = self.b_tree.selection()
        keep = None
        if selected:
            keep = self.b_meta.get(selected[0]) or self.b_folder_rows.get(selected[0])

        self.b_tree.delete(*self.b_tree.get_children())
        self.b_meta, self.b_folder_rows, self.b_link_rows = {}, {}, {}
        open_tasks = set(self.settings.get("buddy_open_tasks", []))
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
            items.sort(key=lambda it: self._buddy_sort_key(it[1]))
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
                links = t.get("links") or []
                rid = self.b_tree.insert(
                    parents.get(t.get("group"), parents[""]), "end",
                    text=" " + (t.get("title") or "(untitled)") + jira_mark(t)
                    + (f"  \U0001F517{len(links)}" if links else ""),
                    image=self._buddy_marker(status, t.get("priority", "")), values=(due,),
                    tags=("overdue",) if overdue else (),
                    open=t.get("id") in open_tasks,
                )
                self.b_meta[rid] = (project, i)
                # Links tuck under the task: click the arrow to show/hide.
                for link in links:
                    lid = self.b_tree.insert(rid, "end", text="\u2197 " + (link.get("title") or link["url"]),
                                             tags=("link",))
                    self.b_link_rows[lid] = link["url"]
                if keep == (project, i):
                    reselect = rid

        if reselect:
            self.b_tree.selection_set(reselect)
        parts = [f"{n} {st.lower().replace(' for approval', '')}"
                 for st, n in counts.items() if n and st != DONE]
        self.b_summary.configure(text="  \u00b7  ".join(parts) or "Nothing open. Nice.")
        self._buddy_update_add()

    def _buddy_fold(self, opened):
        """Remember which projects/groups are collapsed and which tasks
        have their links shown."""
        row = self.b_tree.focus()
        if row in self.b_meta:
            project, i = self.b_meta[row]
            task_id = self.store.projects[project][i].get("id")
            shown = set(self.settings.get("buddy_open_tasks", []))
            (shown.add if opened else shown.discard)(task_id)
            self.settings["buddy_open_tasks"] = sorted(shown)
            save_settings(self.settings)
            return
        key = self.b_folder_rows.get(row)
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
            row = sel[0]
            if row in self.b_link_rows:
                row = self.b_tree.parent(row)
            if row in self.b_meta:
                project, i = self.b_meta[row]
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

    def _buddy_link_click(self, event):
        """A click on a link row opens it."""
        row = self.b_tree.identify_row(event.y)
        if row in self.b_link_rows and self.b_tree.identify_element(event.x, event.y) != "Treeitem.indicator":
            self._open_link(self.b_link_rows[row])

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
        m.add_command(label="Share task\u2026", command=self._share)
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
        # extended: Ctrl/Shift-click several tasks, then drag them together
        self.tree = ttk.Treeview(tree_wrap, show="tree", selectmode="extended")
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
        ttk.Separator(header, orient="vertical").pack(side="right", fill="y", padx=8)
        ttk.Button(header, text="Import\u2026", command=self._import_shared).pack(side="right")
        ttk.Button(header, text="Share\u2026", command=self._share).pack(side="right", padx=4)
        if not HAVE_OPENPYXL:
            self.xlsx_btn.state(["disabled"])

        self.notebook = ttk.Notebook(center)
        self.notebook.pack(fill="both", expand=True, pady=(6, 0))
        tasks_tab = ttk.Frame(self.notebook, padding=4)
        groups_tab = ttk.Frame(self.notebook, padding=4)
        epics_tab = ttk.Frame(self.notebook, padding=4)
        self.notebook.add(tasks_tab, text="Tasks")
        self.notebook.add(epics_tab, text="Epics")
        self.notebook.add(groups_tab, text="Subgroups, Epics & Sprints")
        self._build_epics_tab(epics_tab)
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
        ttk.Label(ctrl, text="Sort by").pack(side="left")
        self.sort_var = tk.StringVar(value="Status")
        sort_cb = ttk.Combobox(ctrl, textvariable=self.sort_var, values=list(SORT_OPTIONS),
                               state="readonly", width=9)
        sort_cb.pack(side="left", padx=(4, 12))
        sort_cb.bind("<<ComboboxSelected>>",
                     lambda e: self._sort_by(SORT_OPTIONS[self.sort_var.get()], keep_dir=True))
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

    def _build_form(self, outer):
        self.form_tabs = ttk.Notebook(outer)
        self.form_tabs.pack(fill="both", expand=True)
        # Quick links: the task's Jira ticket(s), one click away from any tab.
        bg = ttk.Style(self).lookup("TFrame", "background") or "SystemButtonFace"
        self.quick_links = tk.Text(outer, height=1, wrap="word", relief="flat",
                                   borderwidth=0, highlightthickness=0, background=bg,
                                   font=("", 9), cursor="arrow", padx=2, pady=4)
        self.quick_links.pack(side="bottom", fill="x", before=self.form_tabs)
        self.quick_links.tag_configure("label", foreground="#555", font=("", 9, "bold"))
        self.quick_links.tag_configure("key", foreground="#1a5fb4", underline=True,
                                       font=("", 9, "bold"))
        self.quick_links.tag_configure("link", foreground="#1a5fb4", underline=True)
        self.quick_links.tag_configure("dim", foreground="#888")
        self.quick_links.configure(state="disabled")
        self._ql_urls = {}  # tag -> url
        self.quick_links.bind("<Button-1>", self._quick_link_click)
        self.quick_links.bind("<Motion>", self._quick_link_hover)
        form = ttk.Frame(self.form_tabs, padding=(0, 4))
        fields_tab = ttk.Frame(self.form_tabs, padding=(4, 4))
        self.form_tabs.add(form, text="Details")
        self.form_tabs.add(fields_tab, text="Fields")
        jira_tab = ttk.Frame(self.form_tabs, padding=(4, 4))
        self.form_tabs.add(jira_tab, text="Jira")
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

        ttk.Label(form, text="Priority").grid(row=r, column=0, sticky="e", **pad)
        pr_wrap = ttk.Frame(form)
        pr_wrap.grid(row=r, column=1, columnspan=3, sticky="ew", **pad)
        self.f_priority = ttk.Combobox(pr_wrap, values=[""] + PRIORITY_ORDER,
                                       state="readonly", width=20)
        self.f_priority.pack(side="left")
        self.f_priority.bind("<<ComboboxSelected>>", lambda e: self._update_priority_icon())
        self.f_priority_icon = tk.Label(pr_wrap, text="", font=("", 11, "bold"))
        self.f_priority_icon.pack(side="left", padx=6)
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
            command=lambda: (self._toggle_jira(), self._form_changed(delay=1)),
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
        self.f_from_lbl = ttk.Label(dates, text="", foreground="#8e44ad", font=("", 8, "bold"))
        self.f_from_lbl.pack(side="left", padx=(12, 0))
        r += 1

        btnrow = ttk.Frame(form)
        btnrow.grid(row=r, column=0, columnspan=4, pady=(6, 2))
        self.save_btn = ttk.Button(btnrow, text=f"Save  ({MOD_LABEL}+S)", command=self._save_task)
        self.save_btn.pack(side="left", padx=4)
        ttk.Button(btnrow, text=f"New  ({MOD_LABEL}+N)", command=self._new_task).pack(
            side="left", padx=4
        )
        self.f_saved_lbl = ttk.Label(btnrow, text="", foreground="#2d8738", font=("", 8))
        self.f_saved_lbl.pack(side="left", padx=(8, 0))

        # Autosave: every edit is saved shortly after you make it.
        self._autosave_job = None
        self._autosave_target = None
        self._loading_form = False
        for w in (self.f_title, self.f_due, self.f_jira_ref,
                  self.f_desc, self.f_blockers, self.f_notes):
            self._watch_typing(w)
        for combo in (self.status_cb, self.f_priority, self.f_group, self.f_epic, self.f_sprint):
            combo.bind("<<ComboboxSelected>>", lambda e: self._form_changed(delay=1), add="+")

        self._build_fields_tab(fields_tab)
        self._build_jira_tab(jira_tab)
        self._added_date = ""
        self._completed_date = ""
        self._update_status_pill()
        self._toggle_jira()
        self._set_form_enabled(False)

    # -- Jira tab: a read-only copy of the ticket, attached to the task ----
    def _build_jira_tab(self, tab):
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)
        head = ttk.Frame(tab)
        head.grid(row=0, column=0, sticky="ew")
        self.j_key = ttk.Label(head, text="", style="Link.TLabel", cursor="hand2",
                               font=("", 10, "bold", "underline"))
        self.j_key.pack(side="left")
        self.j_key.bind("<Button-1>", lambda e: self._jira and self._jira.get("url")
                        and self._open_link(self._jira["url"]))
        self.j_meta = ttk.Label(head, text="", foreground="#666", font=("", 8))
        self.j_meta.pack(side="left", padx=(8, 0))

        view_wrap = ttk.Frame(tab)
        view_wrap.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        self.j_view = tk.Text(view_wrap, wrap="word", width=40, height=12, font=("", 9),
                              relief="flat", padx=6, pady=4, background="#f6f8fa")
        ys = ttk.Scrollbar(view_wrap, orient="vertical", command=self.j_view.yview)
        self.j_view.configure(yscrollcommand=ys.set)
        ys.pack(side="right", fill="y")
        self.j_view.pack(side="left", fill="both", expand=True)
        self.j_view.tag_configure("h", font=("", 9, "bold"), foreground="#1a5fb4",
                                  spacing1=8, spacing3=2)
        self.j_view.tag_configure("k", font=("", 9, "bold"), foreground="#444")
        self.j_view.tag_configure("dim", foreground="#888")
        # Sprint / epic values: click to see everything in that sprint/epic.
        self.j_view.tag_configure("filter", foreground="#1a5fb4", underline=True)
        for kind in ("sprint", "epic"):
            tag = "filter_" + kind
            self.j_view.tag_bind(tag, "<Button-1>",
                                 lambda e, k=kind: self._jira and self._filter_by(k, self._jira.get(k)))
            self.j_view.tag_bind(tag, "<Enter>", lambda e: self.j_view.configure(cursor="hand2"))
            self.j_view.tag_bind(tag, "<Leave>", lambda e: self.j_view.configure(cursor=""))

        fetch = ttk.Frame(tab)
        fetch.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        ttk.Label(fetch, text="Ticket").pack(side="left")
        self.j_fetch_key = ttk.Entry(fetch, width=16)
        self.j_fetch_key.pack(side="left", padx=4, fill="x", expand=True)
        self.j_fetch_key.bind("<Return>", lambda e: self._jira_fetch())
        self.j_fetch_btn = ttk.Button(fetch, text="Fetch from Jira", command=self._jira_fetch)
        self.j_fetch_btn.pack(side="left")
        self.j_fetch_hint = ttk.Label(tab, text="", foreground="#888", font=("", 8))
        self.j_fetch_hint.grid(row=3, column=0, sticky="w")

        paste = ttk.LabelFrame(tab, text=f"Or paste Jira XML ({MOD_LABEL}+Enter to attach)", padding=4)
        paste.grid(row=4, column=0, sticky="ew", pady=(4, 0))
        paste.columnconfigure(0, weight=1)
        self.j_paste = tk.Text(paste, height=3, width=40, wrap="none", font="TkFixedFont")
        self.j_paste.grid(row=0, column=0, columnspan=4, sticky="ew")
        self.j_paste.bind(f"<{MOD}-Return>", lambda e: (self._jira_attach(), "break")[1])
        self.j_attach_btn = ttk.Button(paste, text="Attach", command=self._jira_attach)
        self.j_attach_btn.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.j_file_btn = ttk.Button(paste, text="Load .xml file\u2026", command=self._jira_attach_file)
        self.j_file_btn.grid(row=1, column=1, padx=(4, 0), pady=(4, 0))
        self.j_remove_btn = ttk.Button(paste, text="Remove", command=self._jira_remove)
        self.j_remove_btn.grid(row=1, column=3, padx=(4, 0), pady=(4, 0))
        self._jira = None
        self._jira_token = ""  # this session's token (if not remembered)

    def _render_jira(self):
        # Fetch bar: pre-fill this task's ticket key; hint if not connected.
        key = (self._jira or {}).get("key") or (
            self.f_jira_ref.get().strip() if self.f_jira_made.get() else "")
        m = JIRA_KEY_RE.search(key.upper()) if key else None
        self.j_fetch_key.delete(0, "end")
        if m:
            self.j_fetch_key.insert(0, m.group(1))
        cfg = load_jira_config()
        self.j_fetch_btn.configure(text="Refresh from Jira" if self._jira else "Fetch from Jira")
        self.j_fetch_hint.configure(
            text="" if cfg.get("base_url") else
            "Connect Jira under Settings \u2192 Jira connection to fetch tickets directly.")
        v, j = self.j_view, self._jira
        v.configure(state="normal")
        v.delete("1.0", "end")
        has = bool(j)
        self.j_remove_btn.state(["!disabled"] if has else ["disabled"])
        if not j:
            self.j_key.configure(text="")
            self.j_meta.configure(text="")
            v.insert("end", "No Jira ticket attached to this task.\n\n", "dim")
            v.insert("end", "In Jira use Export \u2192 XML, copy the XML and paste it below "
                            "(or load the .xml file). The ticket is kept here for "
                            "reference and fills in empty fields on your task (number, "
                            "links, description, dates, epic, sprint) - anything you "
                            "wrote yourself is kept.", "dim")
            v.configure(state="disabled")
            self.form_tabs.tab(2, text="Jira")
            self._render_quick_links()
            return
        self.j_key.configure(text=j.get("key") or "Jira ticket")
        self.j_meta.configure(text=f"attached {j.get('attached', '')}")
        v.insert("end", (j.get("summary") or "") + "\n", "h")
        for label, key in (("Status", "status"), ("Type", "type"), ("Priority", "priority"),
                           ("Resolution", "resolution"), ("Assignee", "assignee"),
                           ("Reporter", "reporter"), ("Sprint", "sprint"), ("Epic", "epic"),
                           ("Due", "due"), ("Created", "created"), ("Updated", "updated"),
                           ("Resolved", "resolved")):
            if j.get(key):
                v.insert("end", f"{label}: ", "k")
                if key in ("sprint", "epic"):
                    v.insert("end", j[key], ("filter", "filter_" + key))
                    v.insert("end", "  (click to filter)" + "\n", "dim")
                else:
                    v.insert("end", j[key] + "\n")
        for name, value in j.get("fields", {}).items():
            v.insert("end", f"{name}: ", "k")
            v.insert("end", value + "\n")
        v.insert("end", "Description\n", "h")
        v.insert("end", (j.get("description") or "(none)") + "\n",
                 () if j.get("description") else "dim")
        comments = j.get("comments", [])
        v.insert("end", f"Comments ({len(comments)})\n", "h")
        for c in comments:
            v.insert("end", f"{c.get('author', '')}  {c.get('created', '')}\n", "k")
            v.insert("end", (c.get("body") or "") + "\n\n")
        v.configure(state="disabled")
        self.form_tabs.tab(2, text=f"Jira {JIRA_ICON}")
        self._render_quick_links()

    def _quick_link_at(self, event):
        for tag in self.quick_links.tag_names(f"@{event.x},{event.y}"):
            if tag in self._ql_urls:
                return self._ql_urls[tag]
        return None

    def _quick_link_click(self, event):
        url = self._quick_link_at(event)
        if url:
            self._open_link(url)
        return "break"

    def _quick_link_hover(self, event):
        url = self._quick_link_at(event)
        self.quick_links.configure(cursor="hand2" if url else "arrow")

    def _render_quick_links(self):
        """Bottom bar: the Jira ticket and its related Jira links, clickable."""
        q = self.quick_links
        q.configure(state="normal")
        q.delete("1.0", "end")
        self._ql_urls = {}
        links = []  # (text, url, tag)
        j = self._jira or {}
        if j.get("url"):
            links.append((j.get("key") or "Jira ticket", j["url"], "key"))
            jira_urls = [l for l in j.get("links", []) if l["url"] != j["url"]
                         and "/browse/" in l["url"]]
            links += [(l["title"], l["url"], "link") for l in jira_urls]
        else:
            ref = self.f_jira_ref.get().strip() if self.f_jira_made.get() else ""
            if ref.startswith(("http://", "https://")):
                links.append((ref.rsplit("/", 1)[-1] or ref, ref, "key"))
            elif ref:
                q.insert("end", "Jira: ", "label")
                q.insert("end", ref + "  (attach the ticket's XML to make it a link)", "dim")
        if links:
            q.insert("end", "Jira: ", "label")
            for n, (text, url, tag) in enumerate(links):
                if n:
                    q.insert("end", "  \u00b7  ", "dim")
                link_tag = f"ql{n}"
                q.insert("end", text, (tag, link_tag))
                self._ql_urls[link_tag] = url
        # grow to fit (wrapped) content, up to 3 lines; hide when empty
        lines = int(q.count("1.0", "end", "displaylines")[0]) if q.get("1.0", "end").strip() else 0
        q.configure(height=max(1, min(3, lines)), state="disabled")
        if lines:
            q.pack(side="bottom", fill="x", before=self.form_tabs)
        else:
            q.pack_forget()

    def _jira_attach(self, text=None):
        """Attach (or refresh) the ticket from pasted XML. Task fields stay as they are."""
        text = text if text is not None else self.j_paste.get("1.0", "end").strip()
        if not text:
            self._status("Paste the Jira XML (Export > XML) into the box first.")
            self.j_paste.focus_set()
            return False
        try:
            issues = parse_jira_issues(text)
        except (ValueError, ET.ParseError) as exc:
            self._status(f"That isn't Jira XML I can read: {exc}")
            return False
        # Prefer the issue matching this task's Jira ref when several were pasted.
        ref = (self._jira or {}).get("key") or self.f_jira_ref.get().strip()
        issue = next((i for i in issues if i["key"] == ref), issues[0])
        old, self._jira = self._jira, issue
        self.j_paste.delete("1.0", "end")

        # Fill in epic, sprint, description, due/completed dates and links -
        # only where you haven't written your own (see jira_updates).
        project = self.active_project
        current = {
            "epic": self.f_epic.get().strip(), "sprint": self.f_sprint.get().strip(),
            "description": self.f_desc.get("1.0", "end").strip(),
            "due_date": self.f_due.get().strip(), "status": self.f_status.get(),
            "priority": self.f_priority.get(),
            "completed_date": self._completed_date, "links": self._links,
            "jira_made": self.f_jira_made.get(),
            "jira_ref": self.f_jira_ref.get().strip() if self.f_jira_made.get() else "",
        }
        updates = jira_updates(current, issue, old)
        for kind in ("epic", "sprint"):
            if issue.get(kind) and issue[kind] not in self.store.names(project, kind):
                self.store.add_name(project, kind, issue[kind])
        self._refresh_choices()
        for kind in ("epic", "sprint"):
            if kind in updates:
                self.f_kinds[kind].set(updates[kind])
        if "description" in updates:
            self._set_text(self.f_desc, updates["description"])
        if "status" in updates:
            self.f_status.set(updates["status"])
            self._update_status_pill()
        if "priority" in updates:
            self.f_priority.set(updates["priority"])
            self._update_priority_icon()
        if "jira_made" in updates:
            self.f_jira_made.set(True)
            self._toggle_jira()
        if "jira_ref" in updates:
            self.f_jira_ref.configure(state="normal")
            self._set_entry(self.f_jira_ref, updates["jira_ref"])
            self._toggle_jira()
        if "due_date" in updates:
            self._set_entry(self.f_due, updates["due_date"])
        if "completed_date" in updates:
            self._completed_date = updates["completed_date"]
            self.f_completed_lbl.configure(text=f"Completed: {self._completed_date}")
        if "links" in updates:
            self._links = updates["links"]
            self._render_links()
        if updates and self.editing_index is not None:
            self.store.projects[project][self.editing_index].update(
                {k: ([dict(x) for x in v] if k == "links" else v) for k, v in updates.items()})

        self._render_jira()
        labels = {"epic": "epic", "sprint": "sprint", "description": "description",
                  "due_date": "due date", "completed_date": "completed date",
                  "jira_ref": "Jira ticket number", "priority": "priority",
                  "status": "status Done (resolved in Jira)"}
        filled = [labels[k] for k in updates if k in labels]
        if "links" in updates:
            n = len(updates["links"]) - len(current["links"])
            filled.append(f"{n} link{'s' if n != 1 else ''}")
        self._persist_jira(f"Attached Jira {issue['key']}"
                           + (f" - filled in {', '.join(filled)}." if filled else
                              " - nothing new to fill in (your values kept)."))
        self.store.note_jira_epics()
        if updates:
            self._refresh_tree()
            self._refresh_table()
            self._refresh_groups()
            self._refresh_epics()
            self._sync_selection()
        return True

    # -- Jira connection -------------------------------------------------------
    def _jira_client(self):
        cfg = load_jira_config()
        token = self._jira_token
        if not token and cfg.get("token"):
            try:
                token = self._jira_token = unprotect_token(cfg["token"])
            except OSError:
                token = ""
        if not cfg.get("base_url") or not token:
            return None
        return JiraClient(cfg["base_url"], cfg.get("auth", "pat"), token,
                          cfg.get("email", ""), cfg.get("ca_file", ""))

    def _run_bg(self, work, done, busy_msg):
        """Run a network call off the UI thread; deliver the result on it."""
        self._status(busy_msg)
        self.configure(cursor="watch")
        box = {}

        def worker():
            try:
                box["result"] = work()
            except JiraError as exc:
                box["error"] = str(exc)
            except Exception as exc:  # noqa: BLE001 - report anything to the user
                box["error"] = f"Unexpected error: {exc}"

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

        def poll():
            if thread.is_alive():
                self.after(100, poll)
                return
            self.configure(cursor="")
            if "error" in box:
                self._status(box["error"])
                messagebox.showerror("Jira", box["error"])
            else:
                done(box["result"])

        self.after(100, poll)

    def _need_jira(self):
        client = self._jira_client()
        if client is None:
            self._status("Set up the Jira connection first (Settings > Jira connection).")
            self._jira_settings()
        return client

    def _jira_fetch(self):
        """Fetch (or refresh) a ticket by key/URL and attach it to this task."""
        if not self.active_project or self.save_btn.instate(["disabled"]):
            self._status("Open a task first.")
            return
        m = JIRA_KEY_RE.search(self.j_fetch_key.get().strip().upper())
        if not m:
            self._status("Type a ticket key like MDA-12 (or paste its URL).")
            self.j_fetch_key.focus_set()
            return
        client = self._need_jira()
        if not client:
            return
        key = m.group(1)
        target = (self.active_project, self.editing_index)

        def done(xml_text):
            if (self.active_project, self.editing_index) != target:
                self._status(f"Fetched {key}, but you'd moved to another task - not attached.")
                return
            self._jira_attach(xml_text)

        self._run_bg(lambda: client.issue_xml(key), done, f"Fetching {key} from Jira\u2026")

    def _jira_refresh_all(self):
        keys = sorted({(t.get("jira") or {}).get("key") for ts in self.store.projects.values()
                       for t in ts if (t.get("jira") or {}).get("key")})
        if not keys:
            self._status("No tasks have a Jira ticket attached yet.")
            return
        client = self._need_jira()
        if not client:
            return

        def work():
            parts = []
            for i in range(0, len(keys), 50):  # keep the query a sane length
                parts.append(client.search_xml(f"key in ({', '.join(keys[i:i + 50])})"))
            return parts

        def done(parts):
            for xml_text in parts:
                self._import_jira_text(xml_text, create=False)

        self._run_bg(work, done, f"Refreshing {len(keys)} Jira ticket(s)\u2026")

    def _gather_imports(self):
        """Move tasks created by Jira imports back into the Imported project."""
        ids = [t["id"] for p, ts in self.store.projects.items() if p != IMPORTED_PROJECT
               for t in ts if str(t.get("id", "")).startswith("jira:")]
        if not ids:
            self._status("No Jira-imported tasks outside 'Imported'.")
            return
        if not messagebox.askyesno(
                "Gather Jira imports",
                f"Move {len(ids)} task(s) created by Jira imports into '{IMPORTED_PROJECT}'?\n\n"
                "This includes any you've already moved to other projects. Tasks you "
                "attached a ticket to yourself stay where they are."):
            return
        if IMPORTED_PROJECT not in self.store.projects:
            self.store.add_project(IMPORTED_PROJECT)
        self._move_tasks(ids, IMPORTED_PROJECT, "")
        self._status(f"Moved {len(ids)} Jira-imported task(s) into '{IMPORTED_PROJECT}'.")

    def _move_tasks(self, ids, dest, group):
        """Move tasks (by id) into dest / group, keeping the form pointed right."""
        for task_id in ids:
            found = self._find_task(task_id)
            if not found:
                continue
            project, index = found
            task = dict(self.store.projects[project][index])
            if (project, task.get("group", "")) == (dest, group):
                continue
            task["group"] = group
            self._relocate_task(project, index, task, dest)
        self._refresh_views()

    def _jira_search_import(self):
        client = self._need_jira()
        if not client:
            return
        where = f"the '{IMPORTED_PROJECT}' project"

        def run(jql):
            if not jql:
                return
            self.settings["jira_last_jql"] = jql
            save_settings(self.settings)
            self._run_bg(lambda: client.search_xml(jql),
                         lambda xml_text: self._import_jira_text(xml_text),
                         "Searching Jira\u2026")

        self._prompt("Import from Jira search",
                     f"JQL query. Matching tickets become tasks in {where}; ones you "
                     "already have are just refreshed.",
                     self.settings.get("jira_last_jql",
                                       "assignee = currentUser() AND resolution = Unresolved"),
                     run)

    def _jira_settings(self):
        """In-app panel for the optional Jira connection."""
        self._close_prompt()
        cfg = load_jira_config()
        bg = "#ffffff"
        panel = self._prompt_panel = tk.Frame(self, bg=bg, highlightthickness=2,
                                              highlightbackground=ACCENT, padx=14, pady=12)
        tk.Label(panel, text="Jira connection (optional)", bg=bg,
                 font=("", 11, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
        tk.Label(panel, bg=bg, fg="#555", wraplength=320, justify="left",
                 text="Read-only: fetches tickets the same way as Export \u2192 XML. "
                      "Saved in profile/jira.json (not in git).").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(2, 8))

        auth = tk.StringVar(value=cfg.get("auth", "pat"))
        remember = tk.BooleanVar(value=bool(cfg.get("token")))
        rows = {}

        def row(r, label, widget):
            tk.Label(panel, text=label, bg=bg).grid(row=r, column=0, sticky="w", pady=2)
            widget.grid(row=r, column=1, columnspan=2, sticky="ew", pady=2)
            rows[label] = widget
            return widget

        base = row(2, "Jira address", ttk.Entry(panel, width=36))
        base.insert(0, cfg.get("base_url", ""))
        kind = tk.Frame(panel, bg=bg)
        for text, value in (("Server / Data Center (personal access token)", "pat"),
                            ("Cloud (email + API token)", "cloud")):
            tk.Radiobutton(kind, text=text, variable=auth, value=value, bg=bg,
                           activebackground=bg, anchor="w",
                           command=lambda: toggle_email()).pack(anchor="w")
        row(3, "Type", kind)
        email = row(4, "Email (Cloud)", ttk.Entry(panel, width=36))
        email.insert(0, cfg.get("email", ""))
        token = row(5, "Token", ttk.Entry(panel, width=36, show="\u2022"))
        token.insert(0, self._jira_token)
        if cfg.get("token") and not self._jira_token:
            tk.Label(panel, text="(saved - leave blank to keep)", bg=bg,
                     fg="#888", font=("", 8)).grid(row=6, column=1, sticky="w")
        tk.Checkbutton(panel, text="Remember token on this computer" +
                       (" (encrypted with your Windows login)" if sys.platform == "win32"
                        else " (file readable only by you)"),
                       variable=remember, bg=bg, activebackground=bg).grid(
            row=7, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ca = row(8, "CA file (optional)", ttk.Entry(panel, width=30))
        ca.insert(0, cfg.get("ca_file", ""))
        ttk.Button(panel, text="Browse\u2026", command=lambda: (
            lambda f: f and (ca.delete(0, "end"), ca.insert(0, f)))(
            filedialog.askopenfilename(title="Company CA certificate",
                                       filetypes=[("Certificates", "*.pem *.crt *.cer"),
                                                  ("All files", "*.*")]))).grid(
            row=9, column=2, sticky="e")
        result = tk.Label(panel, text="", bg=bg, fg="#555", wraplength=320, justify="left")
        result.grid(row=10, column=0, columnspan=3, sticky="w", pady=(6, 0))

        def toggle_email():
            email.configure(state="normal" if auth.get() == "cloud" else "disabled")
        toggle_email()

        def current_token():
            return token.get().strip() or self._jira_token or unprotect_token(cfg.get("token", ""))

        def gather():
            new = {"base_url": jira_base_url(base.get()), "auth": auth.get(),
                   "email": email.get().strip(), "ca_file": ca.get().strip()}
            tok = current_token()
            if remember.get() and tok:
                new["token"] = protect_token(tok)
            return new, tok

        def test():
            new, tok = gather()
            client = JiraClient(new["base_url"], new["auth"], tok, new["email"], new["ca_file"])
            result.configure(text="Testing\u2026", fg="#555")

            def ok(name):
                if result.winfo_exists():
                    result.configure(text=f"\u2713 Connected as {name}.", fg="#2d8738")
                self._status(f"Jira connection works - connected as {name}.")

            self._run_bg(client.myself, ok, "Testing the Jira connection\u2026")

        def save():
            new, tok = gather()
            if not new["base_url"]:
                result.configure(text="Enter your Jira address first.", fg="#c0392b")
                return
            save_jira_config(new)
            self._jira_token = tok
            self._close_prompt()
            self._render_jira()
            self._status("Jira connection saved" + ("" if new.get("token") else
                         " (token kept for this session only)") + ".")

        def disconnect():
            if os.path.exists(JIRA_FILE):
                os.remove(JIRA_FILE)
            self._jira_token = ""
            self._close_prompt()
            self._render_jira()
            self._status("Jira connection removed from this computer.")

        btns = tk.Frame(panel, bg=bg)
        btns.grid(row=11, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        ttk.Button(btns, text="Save", command=save).pack(side="right")
        ttk.Button(btns, text="Cancel", command=self._close_prompt).pack(side="right", padx=4)
        ttk.Button(btns, text="Test connection", command=test).pack(side="left")
        if cfg:
            ttk.Button(btns, text="Disconnect", command=disconnect).pack(side="left", padx=4)
        panel.columnconfigure(1, weight=1)
        panel.place(relx=0.5, rely=0.06, anchor="n")
        panel.lift()
        base.focus_set()

    def _jira_attach_file(self):
        path = filedialog.askopenfilename(
            title="Attach Jira XML", filetypes=[("Jira XML export", "*.xml"), ("All files", "*.*")])
        if path:
            try:
                with open(path, "r", encoding="utf-8-sig") as f:
                    self._jira_attach(f.read())
            except OSError as exc:
                self._status(f"Couldn't read {path}: {exc}")

    def _jira_remove(self):
        if self._jira and messagebox.askyesno(
                "Remove Jira ticket", f"Remove the attached Jira copy of {self._jira.get('key')}?"
                "\n\nYour task itself isn't changed."):
            self._jira = None
            self._render_jira()
            self._persist_jira("Removed the attached Jira ticket.")

    def _filter_by(self, kind, value):
        """Show everything in this sprint/epic: Visual Planner, filtered."""
        if not value:
            return
        project = self.active_project
        self._show_view("Visual Planner")
        if self.p_columns.get() == "Project":
            self.p_columns.set("Status")
        if project in self.store.projects:
            self.p_project.set(project)
        self.p_epic.set(value if kind == "epic" else ANY)
        self.p_sprint.set(value if kind == "sprint" else ANY)
        self._refresh_planner()
        self._status(f"Showing {kind} '{value}'" + (f" in '{project}'." if project else "."))

    def _persist_jira(self, msg):
        """Like links: saves straight away on an existing task."""
        if self.active_project is not None and self.editing_index is not None:
            task = self.store.projects[self.active_project][self.editing_index]
            if self._jira:
                task["jira"] = self._jira
            else:
                task.pop("jira", None)
            self.store.save()
            self._status(msg)
        else:
            self._form_changed(delay=1)  # saved with the task once it has a title
            self._status(msg)

    # -- custom fields tab --------------------------------------------------
    def _build_fields_tab(self, tab):
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)
        ttk.Label(
            tab, foreground="#666", font=("", 8), wraplength=320, justify="left",
            text="Extra fields for this task. Jira imports create these automatically; "
                 "pick one below or type a new name to add your own.",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

        bg = ttk.Style(self).lookup("TFrame", "background") or "SystemButtonFace"
        canvas = tk.Canvas(tab, highlightthickness=0, bg=bg)
        ys = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        canvas.grid(row=1, column=0, sticky="nsew")
        ys.grid(row=1, column=1, sticky="ns")
        canvas.configure(yscrollcommand=ys.set)
        self.f_fields_box = ttk.Frame(canvas)
        self.f_fields_box.columnconfigure(1, weight=1)
        win = canvas.create_window(0, 0, window=self.f_fields_box, anchor="nw")
        self.f_fields_box.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))
        self.f_fields_canvas = canvas

        add = ttk.Frame(tab)
        add.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        self.f_field_pick = ttk.Combobox(add)
        self.f_field_pick.pack(side="left", fill="x", expand=True)
        self.f_field_pick.bind("<Return>", lambda e: self._add_field())
        self.f_field_add_btn = ttk.Button(add, text="Add field", command=self._add_field)
        self.f_field_add_btn.pack(side="left", padx=(4, 0))
        self.fields_save_btn = ttk.Button(tab, text=f"Save  ({MOD_LABEL}+S)", command=self._save_task)
        self.fields_save_btn.grid(row=3, column=0, columnspan=2, pady=(8, 2))
        self._field_widgets = {}

    def _field_value(self, widget):
        if isinstance(widget, tk.Text):
            return widget.get("1.0", "end-1c")
        return widget.get()

    def _current_fields(self):
        return {name: self._field_value(w) for name, w in self._field_widgets.items()}

    def _collect_fields(self):
        """Fields to save: the non-empty ones."""
        return {n: v.strip() for n, v in self._current_fields().items() if v.strip()}

    def _render_fields(self, values):
        box = self.f_fields_box
        for w in box.winfo_children():
            w.destroy()
        self._field_widgets = {}
        defined = [d["name"] for d in self.store.fields]
        names = [n for n in defined if n in values] + [n for n in values if n not in defined]
        if not names:
            ttk.Label(box, text="No extra fields on this task.", foreground="#888",
                      font=("", 8)).grid(row=0, column=0, columnspan=3, sticky="w")
        for r, name in enumerate(names):
            ttk.Label(box, text=name, font=("", 8, "bold"), foreground="#444",
                      wraplength=110, justify="right").grid(
                row=r, column=0, sticky="ne", padx=(0, 6), pady=2)
            if self.store.field_type(name) == "long":
                w = tk.Text(box, height=4, width=26, wrap="word", font=("", 9))
                w.insert("1.0", values[name])
            else:
                w = ttk.Entry(box, width=26)
                w.insert(0, values[name])
            w.grid(row=r, column=1, sticky="ew", pady=2)
            x = ttk.Label(box, text="\u2715", style="LinkAction.TLabel", cursor="hand2")
            x.grid(row=r, column=2, sticky="n", padx=(4, 0), pady=2)
            x.bind("<Button-1>", lambda e, n=name: self._remove_field(n))
            self._field_widgets[name] = w
            self._watch_typing(w)
        for w in [box, self.f_fields_canvas, *box.winfo_children()]:
            if not isinstance(w, tk.Text):
                bind_wheel(w, lambda n: self.f_fields_canvas.yview_scroll(n, "units"))
        self.f_field_pick.configure(values=[n for n in defined if n not in values])
        self.form_tabs.tab(1, text=f"Fields ({len(names)})" if names else "Fields")
        self.f_fields_canvas.yview_moveto(0)

    def _add_field(self):
        name = self.f_field_pick.get().strip()
        if not name:
            self._status("Pick a field or type a new field name, then 'Add field'.")
            return
        if self.store.ensure_field(name):
            self._status(f"Created new field '{name}' - it's now available on every task.")
        values = self._current_fields()
        values.setdefault(name, "")
        self._render_fields(values)
        self._form_changed(delay=1)
        self.f_field_pick.set("")
        self._field_widgets[name].focus_set()

    def _remove_field(self, name):
        values = self._current_fields()
        values.pop(name, None)
        self._render_fields(values)
        self._form_changed(delay=1)
        self._status(f"Removed '{name}' from this task.")

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
            self._form_changed(delay=1)  # creates the task once it has a title
            self._status(msg if self.f_title.get().strip()
                         else "Link added - give the task a title to save it.")

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

    def _update_priority_icon(self):
        p = self.f_priority.get()
        color, glyph = PRIORITIES.get(p, ("#888", ""))
        self.f_priority_icon.configure(text=glyph, fg=color)

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
        for w in (self.status_cb, self.f_priority, self.f_group, self.f_epic, self.f_sprint):
            w.configure(state="readonly" if enabled else "disabled")
        self.jira_chk.state(["!disabled"] if enabled else ["disabled"])
        for w in (self.f_link_title, self.f_link_desc, self.f_link_url, self.f_link_btn,
                  self.f_field_pick, self.f_field_add_btn, self.fields_save_btn,
                  self.j_attach_btn, self.j_file_btn):
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
        self._flush_autosave()  # don't lose an edit made just before switching
        self._loading_form = True
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
        self.f_priority.set(task.get("priority", ""))
        self._update_priority_icon()
        self._added_date = task.get("added_date", "")
        self._completed_date = task.get("completed_date", "")
        self.f_added_lbl.configure(text=f"Created: {self._added_date or '-'}")
        self.f_completed_lbl.configure(text=f"Completed: {self._completed_date or '-'}")
        self._update_status_pill()
        shared_by = task.get("shared_by", "")
        self.f_from_lbl.configure(text=f"From: {shared_by}" if shared_by else "")
        self._render_fields(dict(task.get("fields") or {}))
        self._jira = task.get("jira") or None
        self.j_paste.delete("1.0", "end")
        self._render_jira()
        self._links = [dict(link) for link in task.get("links", [])]
        self._reset_link_entries()
        self._render_links()
        self._loading_form = False
        self.f_saved_lbl.configure(text="")

    # -- autosave -----------------------------------------------------------
    def _watch_typing(self, widget):
        for seq in ("<KeyRelease>", "<<Paste>>", "<<Cut>>"):
            widget.bind(seq, lambda e: self._form_changed(), add="+")

    def _form_changed(self, event=None, delay=AUTOSAVE_DELAY):
        """Something in the form changed: save it shortly (debounced)."""
        if self._loading_form or self.save_btn.instate(["disabled"]) or not self.active_project:
            return
        if self._autosave_target is None:
            # Remember *which* task this edit belongs to (by id, so it survives
            # index shifts), or that it's a new task in this project.
            if self.editing_index is not None:
                task = self.store.projects[self.active_project][self.editing_index]
                self._autosave_target = ("task", task.get("id"))
            else:
                self._autosave_target = ("new", self.active_project)
        if self._autosave_job:
            self.after_cancel(self._autosave_job)
        self._autosave_job = self.after(delay, self._autosave)

    def _flush_autosave(self):
        if self._autosave_job:
            self.after_cancel(self._autosave_job)
            self._autosave()

    def _find_task(self, task_id):
        for project, tasks in self.store.projects.items():
            for i, t in enumerate(tasks):
                if t.get("id") == task_id:
                    return project, i
        return None

    def _task_from_form(self):
        """The form's contents as a task dict (None if there's no title)."""
        title = self.f_title.get().strip()
        if not title:
            return None
        status = self.f_status.get()
        task = {
            "title": title,
            "status": status,
            "priority": self.f_priority.get(),
            "group": self.f_group.get().strip(),
            "epic": self.f_epic.get().strip(),
            "sprint": self.f_sprint.get().strip(),
            "description": self.f_desc.get("1.0", "end").strip(),
            "added_date": self._added_date or now_stamp(),
            "due_date": self.f_due.get().strip(),
            "completed_date": (self._completed_date or now_stamp()) if status == DONE else "",
            "jira_made": self.f_jira_made.get(),
            "jira_ref": self.f_jira_ref.get().strip() if self.f_jira_made.get() else "",
            "blockers": self.f_blockers.get("1.0", "end").strip(),
            "notes": self.f_notes.get("1.0", "end").strip(),
            "links": [dict(link) for link in self._links],
            "fields": self._collect_fields(),
        }
        if self._jira:
            task["jira"] = self._jira
        return task

    def _autosave(self):
        self._autosave_job = None
        target, self._autosave_target = self._autosave_target, None
        if not target:
            return
        kind, ref = target
        if kind == "task":
            found = self._find_task(ref)
            if not found:
                return  # deleted meanwhile
            project, index = found
        else:
            project, index = ref, None
            if project not in self.store.projects:
                return
        task = self._task_from_form()
        if task is None:
            if self.f_title.winfo_ismapped():
                self.f_saved_lbl.configure(text="Add a title to save", foreground="#b9770e")
            return

        same_form = (self.active_project, self.editing_index) == (project, index)
        if index is None:
            self.store.add_task(project, task)
            index = len(self.store.projects[project]) - 1
            if self.active_project == project and self.editing_index is None:
                self.editing_index = index  # the form now edits the new task
                same_form = True
            self._status(f"Added task '{task['title']}' to '{project}'.")
        else:
            stored = self.store.projects[project][index]
            merged = {**stored, **task}  # keep fields the form doesn't show
            if merged == stored:
                return
            self.store.update_task(project, index, merged)
            task = merged

        if same_form:
            # Reflect stamps without reloading the form (keeps your cursor).
            self._added_date = task["added_date"]
            self._completed_date = task["completed_date"]
            self.f_added_lbl.configure(text=f"Created: {self._added_date}")
            self.f_completed_lbl.configure(text=f"Completed: {self._completed_date or '-'}")
            self.f_saved_lbl.configure(text=f"\u2713 Saved {datetime.now():%H:%M:%S}",
                                       foreground="#2d8738")
        self._refresh_views()

    # -- views ------------------------------------------------------------
    def _refresh_views(self):
        self._refresh_tree()
        self._refresh_table()
        self._refresh_groups()
        self._refresh_epics()
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
                    text=mark + (task.get("title") or "(untitled)") + jira_mark(task),
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
        if event.state & 0x5:  # Shift/Ctrl: let the tree extend the selection
            return None
        if meta and meta[0] == "task":
            selected = [r for r in self.tree.selection()
                        if self.node_meta.get(r, ("",))[0] == "task"]
            group = selected if row in selected and len(selected) > 1 else [row]
            ids = [self.store.projects[self.node_meta[r][1]][self.node_meta[r][2]].get("id")
                   for r in group]
            self._tdrag = {"key": meta[1:], "ids": ids, "row": row, "x": event.x,
                           "y": event.y, "active": False, "target": None}
            if len(group) > 1:
                return "break"  # keep the multi-selection for dragging

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
        if d and not d["active"] and len(d.get("ids", [])) > 1:
            self.tree.selection_set(d["row"])  # plain click inside a multi-selection
            return None
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
        where = f"subgroup '{group}'" if group else "no subgroup"
        if len(d["ids"]) > 1:
            self._move_tasks(d["ids"], dest, group)
            self._status(f"Moved {len(d['ids'])} tasks to {dest} / {where}.")
            return "break"
        project, index = d["key"]
        task = dict(self.store.projects[project][index])
        if (dest, group) == (project, task.get("group", "")):
            return "break"
        task["group"] = group
        self._relocate_task(project, index, task, dest)
        self._refresh_views()
        self._status(f"Moved '{task.get('title', '')}' to {dest} / {where}.")
        return "break"

    def _sort_key(self, task):
        col = self.sort_col
        if col == "status":
            primary = STATUS_ORDER.index(task.get("status", DEFAULT_STATUS))
            return (primary, priority_rank(task), task.get("due_date") or "~")
        if col == "priority":
            return (priority_rank(task), STATUS_ORDER.index(task.get("status", DEFAULT_STATUS)),
                    task.get("due_date") or "~")
        v = (task.get(col) or "").lower()
        return (v == "", v)  # blanks last

    def _sort_by(self, col, keep_dir=False):
        if self.sort_col == col and not keep_dir:
            self.sort_rev = not self.sort_rev
        elif self.sort_col != col:
            self.sort_col, self.sort_rev = col, False
        # keep the Sort by dropdown in step with header clicks
        label = next((k for k, v in SORT_OPTIONS.items() if v == col), None)
        if label:
            self.sort_var.set(label)
        for key, heading in [("title", "Title")] + [c[:2] for c in TABLE_COLS]:
            arrow = (" ▼" if self.sort_rev else " ▲") if key == col else ""
            self.table.heading("#0" if key == "title" else key, text=heading + arrow)
        self._refresh_table()
        self._sync_selection()

    def _refresh_table(self):
        self.table.delete(*self.table.get_children())
        self.row_meta = {}  # row id -> (project, task index)
        group = self.group_var.get()
        all_projects = group == "Project"
        project = self.active_project
        if all_projects:
            scope = sorted(self.store.projects, key=str.lower)
        elif project in self.store.projects:
            scope = [project]
        else:
            self.context_var.set("No project selected.")
            return

        every = [(p, i, t) for p in scope for i, t in enumerate(self.store.projects[p])]
        open_count = sum(1 for _, _, t in every if t.get("status") != DONE)
        title = "All projects" if all_projects else project
        self.context_var.set(f"{title}  -  {open_count} open / {len(every)} total")

        items = [it for it in every
                 if not (self.hide_done.get() and it[2].get("status") == DONE)]
        items.sort(key=lambda it: self._sort_key(it[2]), reverse=self.sort_rev)

        if group == "None":
            for p, i, t in items:
                self._insert_row("", p, i, t)
            return

        key = "project" if all_projects else kind_of(group)
        groups = {}
        for p, i, t in items:
            value = p if all_projects else (t.get(key) or "")
            groups.setdefault(value, []).append((p, i, t))
        if all_projects:
            names = [p for p in scope if p in groups]
        elif key == "status":
            names = [s for s in STATUS_ORDER if s in groups]
        elif key == "priority":
            names = [p for p in PRIORITY_ORDER if p in groups] + ([""] if "" in groups else [])
        else:
            # The project's list order, then anything not in the list.
            names = [g for g in self.store.names(project, key) if g in groups]
            names += sorted((g for g in groups if g and g not in names), key=str.lower)
            if "" in groups:
                names.append("")
        for name in names:
            label = name or f"(no {group.lower()})"
            if all_projects:
                label = "\U0001F4C1 " + name
            gid = self.table.insert(
                "", "end", text=f"{label}  ({len(groups[name])})",
                open=True, tags=("group",),
            )
            for p, i, t in groups[name]:
                self._insert_row(gid, p, i, t)

    def _insert_row(self, parent, project, index, task):
        status = task.get("status", DEFAULT_STATUS)
        values = []
        for key, *_ in TABLE_COLS:
            v = task.get(key, "")
            if key == "status":
                v = "● " + v
            elif key == "priority":
                v = priority_label(v)
            values.append(v)
        rid = self.table.insert(
            parent, "end", text=(task.get("title") or "(untitled)") + jira_mark(task),
            values=values, tags=(STATUSES[status][1],),
        )
        self.row_meta[rid] = (project, index)

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
            if node and node in self.tree.selection():
                self.tree.see(node)  # already selected (maybe with others) - keep it
            elif node:
                self.tree.selection_set(node)
                self.tree.see(node)
            else:
                self.tree.selection_remove(self.tree.selection())

            row = None
            current = (self.active_project, self.editing_index)
            for rid, key in self.row_meta.items():
                if self.editing_index is not None and key == current:
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

    # -- Epics tab: epics with their own details --------------------------------
    def _build_epics_tab(self, tab):
        top = ttk.Frame(tab)
        top.pack(fill="x", pady=(0, 4))
        ttk.Label(top, text="Project").pack(side="left")
        self.e_filter = ttk.Combobox(top, state="readonly", width=18)
        self.e_filter.set(ALL_PROJECTS)
        self.e_filter.pack(side="left", padx=(4, 10))
        self.e_filter.bind("<<ComboboxSelected>>", lambda e: self._refresh_epics())
        self.e_hide_done = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Hide done", variable=self.e_hide_done,
                        command=self._refresh_epics).pack(side="left")
        self.e_new_name = ttk.Entry(top, width=18)
        ttk.Button(top, text="+ Epic", command=self._epic_new).pack(side="right")
        self.e_new_name.pack(side="right", padx=4)
        self.e_new_name.bind("<Return>", lambda e: self._epic_new())
        ttk.Label(top, text="New epic").pack(side="right")

        # List on top (full width); details + its tasks side by side below.
        panes = ttk.Panedwindow(tab, orient="vertical")
        panes.pack(fill="both", expand=True)
        left = ttk.Frame(panes)
        right = ttk.LabelFrame(panes, text="Epic", padding=8)
        panes.add(left, weight=1)
        panes.add(right, weight=1)

        self.e_list = ttk.Treeview(left, columns=("project", "sprint", "status", "tasks"),
                                   selectmode="browse", height=6)
        for col, text, width in (("#0", "Epic", 170), ("project", "Project", 110),
                                 ("sprint", "Sprint", 90), ("status", "Status", 100),
                                 ("tasks", "Open / all", 70)):
            self.e_list.heading(col, text=text)
            self.e_list.column(col, width=width, stretch=(col == "#0"))
        ys = ttk.Scrollbar(left, orient="vertical", command=self.e_list.yview)
        self.e_list.configure(yscrollcommand=ys.set)
        ys.pack(side="right", fill="y")
        self.e_list.pack(fill="both", expand=True)
        self.e_list.bind("<<TreeviewSelect>>", lambda e: self._on_epic_select())
        for color, tag in STATUSES.values():
            self.e_list.tag_configure(tag, foreground=color)
        self.e_rows = {}   # row id -> (project, epic name)
        self.e_sel = None  # the epic shown on the right

        right.columnconfigure(1, weight=1)
        right.columnconfigure(3, weight=1)
        pad = {"padx": 4, "pady": 3}
        ttk.Label(right, text="Name").grid(row=0, column=0, sticky="e", **pad)
        self.e_name = ttk.Entry(right)
        self.e_name.grid(row=0, column=1, columnspan=2, sticky="ew", **pad)
        self.e_name.bind("<Return>", lambda e: self._epic_rename())
        self.e_name.bind("<FocusOut>", lambda e: self._epic_rename())
        ttk.Label(right, text="Project").grid(row=1, column=0, sticky="e", **pad)
        self.e_project = ttk.Combobox(right, state="readonly")
        self.e_project.grid(row=1, column=1, columnspan=2, sticky="ew", **pad)
        self.e_project.bind("<<ComboboxSelected>>", lambda e: self._epic_move())
        ttk.Label(right, text="Sprint").grid(row=2, column=0, sticky="e", **pad)
        self.e_sprint = ttk.Combobox(right, state="readonly")
        self.e_sprint.grid(row=2, column=1, columnspan=2, sticky="ew", **pad)
        self.e_sprint.bind("<<ComboboxSelected>>",
                           lambda e: self._epic_set(sprint=self.e_sprint.get()))
        ttk.Label(right, text="Status").grid(row=3, column=0, sticky="e", **pad)
        self.e_status = ttk.Combobox(right, state="readonly",
                                     values=["(from its tasks)"] + STATUS_ORDER)
        self.e_status.grid(row=3, column=1, columnspan=2, sticky="ew", **pad)
        self.e_status.bind("<<ComboboxSelected>>", lambda e: self._epic_set(
            status="" if self.e_status.current() == 0 else self.e_status.get()))
        ttk.Label(right, text="Jira").grid(row=4, column=0, sticky="e", **pad)
        self.e_jira = ttk.Entry(right)
        self.e_jira.grid(row=4, column=1, sticky="ew", **pad)
        self.e_jira.bind("<KeyRelease>", lambda e: self._epic_set_later(
            jira_key=self.e_jira.get().strip()))
        self.e_jira_open = ttk.Label(right, text="Open \u2197", style="Link.TLabel",
                                     cursor="hand2")
        self.e_jira_open.grid(row=4, column=2, **pad)
        self.e_jira_open.bind("<Button-1>", lambda e: self._epic_open_jira())
        ttk.Label(right, text="Description").grid(row=5, column=0, sticky="ne", **pad)
        self.e_desc = tk.Text(right, height=5, width=30, wrap="word", font=("", 9))
        self.e_desc.grid(row=5, column=1, columnspan=2, sticky="nsew", **pad)
        self.e_desc.bind("<KeyRelease>", lambda e: self._epic_set_later(
            description=self.e_desc.get("1.0", "end").strip()))
        self.e_tasks_lbl = ttk.Label(right, text="Tasks", font=("", 9, "bold"))
        self.e_tasks_lbl.grid(row=0, column=3, sticky="w", padx=(16, 4))
        self.e_sprints_lbl = ttk.Label(right, text="", foreground="#666", font=("", 8),
                                       wraplength=230, justify="left")
        self.e_sprints_lbl.grid(row=1, column=3, sticky="w", padx=(16, 4))
        self.e_tasks = ttk.Treeview(right, show="tree", height=6, selectmode="browse")
        self.e_tasks.grid(row=2, column=3, rowspan=4, sticky="nsew", padx=(16, 4), pady=3)
        right.rowconfigure(5, weight=1)
        for color, tag in STATUSES.values():
            self.e_tasks.tag_configure(tag, foreground=color)
        self.e_tasks.bind("<<TreeviewSelect>>", lambda e: self._epic_open_task())
        self.e_task_rows = {}
        btns = ttk.Frame(right)
        btns.grid(row=6, column=0, columnspan=4, sticky="ew", pady=(6, 0))
        self.e_newtask_btn = ttk.Button(btns, text="+ New task in this epic",
                                        command=self._epic_new_task)
        self.e_newtask_btn.pack(side="left")
        self.e_delete_btn = ttk.Button(btns, text="Delete epic", command=self._epic_delete)
        self.e_delete_btn.pack(side="right")
        self._epic_job = None
        self._epic_pending = {}
        self._load_epic(None)

    def _epic_status(self, project, name):
        """The epic's own status, or one worked out from its tasks."""
        meta = self.store.epic_meta(project, name)
        if meta.get("status") in STATUSES:
            return meta["status"]
        tasks = [t for t in self.store.projects.get(project, []) if t.get("epic") == name]
        states = {t.get("status") for t in tasks}
        if tasks and states == {DONE}:
            return DONE
        if states & {"In Progress", "Blocked", "Waiting for approval", DONE}:
            return "In Progress"
        return DEFAULT_STATUS

    def _refresh_epics(self):
        projects = sorted(self.store.projects, key=str.lower)
        self.e_filter.configure(values=[ALL_PROJECTS] + projects)
        if self.e_filter.get() not in projects:
            self.e_filter.set(ALL_PROJECTS)
        chosen = self.e_filter.get()
        scope = projects if chosen == ALL_PROJECTS else [chosen]
        self.e_list.delete(*self.e_list.get_children())
        self.e_rows = {}
        reselect = None
        for project in scope:
            for name in self.store.names(project, "epic"):
                status = self._epic_status(project, name)
                if self.e_hide_done.get() and status == DONE:
                    continue
                meta = self.store.epic_meta(project, name)
                tasks = [t for t in self.store.projects[project] if t.get("epic") == name]
                open_n = sum(1 for t in tasks if t.get("status") != DONE)
                rid = self.e_list.insert(
                    "", "end", text=name + (f"  {JIRA_ICON}" if meta.get("jira_key") else ""),
                    values=(project, meta.get("sprint", ""), "\u25cf " + status,
                            f"{open_n} / {len(tasks)}"),
                    tags=(STATUSES[status][1],))
                self.e_rows[rid] = (project, name)
                if (project, name) == self.e_sel:
                    reselect = rid
        if reselect:
            self.e_list.selection_set(reselect)
        if self.e_sel and not reselect:
            self._load_epic(None)
        elif self.e_sel:
            self._load_epic_tasks()

    def _on_epic_select(self):
        sel = self.e_list.selection()
        key = self.e_rows.get(sel[0]) if sel else None
        if key and key != self.e_sel:
            self._flush_epic()
            self._load_epic(key)

    def _load_epic(self, key):
        """Show an epic's details on the right (None = nothing selected)."""
        self.e_sel = key
        widgets = (self.e_name, self.e_jira)
        for w in widgets:
            w.configure(state="normal")
            w.delete(0, "end")
        self.e_desc.configure(state="normal")
        self.e_desc.delete("1.0", "end")
        self.e_project.set("")
        self.e_sprint.set("")
        self.e_status.set("")
        enabled = key is not None
        if enabled:
            project, name = key
            meta = self.store.epic_meta(project, name)
            self.e_name.insert(0, name)
            self.e_project.configure(values=sorted(self.store.projects, key=str.lower))
            self.e_project.set(project)
            self.e_sprint.configure(values=[""] + self.store.names(project, "sprint"))
            self.e_sprint.set(meta.get("sprint", ""))
            self.e_status.set(meta.get("status") or "(from its tasks)")
            self.e_jira.insert(0, meta.get("jira_key", ""))
            self.e_desc.insert("1.0", meta.get("description", ""))
        for w in (self.e_name, self.e_jira):
            w.configure(state="normal" if enabled else "disabled")
        self.e_desc.configure(state="normal" if enabled else "disabled")
        for w in (self.e_project, self.e_sprint, self.e_status):
            w.configure(state="readonly" if enabled else "disabled")
        for w in (self.e_newtask_btn, self.e_delete_btn):
            w.state(["!disabled"] if enabled else ["disabled"])
        self._load_epic_tasks()

    def _load_epic_tasks(self):
        self.e_tasks.delete(*self.e_tasks.get_children())
        self.e_task_rows = {}
        if not self.e_sel:
            self.e_tasks_lbl.configure(text="Select an epic on the left.")
            self.e_sprints_lbl.configure(text="")
            return
        project, name = self.e_sel
        tasks = [(i, t) for i, t in enumerate(self.store.projects.get(project, []))
                 if t.get("epic") == name]
        self.e_tasks_lbl.configure(text=f"Tasks in this epic ({len(tasks)})")
        sprints = sorted({t.get("sprint") for _, t in tasks if t.get("sprint")}, key=str.lower)
        own = self.store.epic_meta(project, name).get("sprint", "")
        parts = [f"Epic sprint: {own}" if own else "Epic not attached to a sprint"]
        parts.append("Its tasks are in: " + (", ".join(sprints) if sprints else "no sprint"))
        self.e_sprints_lbl.configure(text="  \u00b7  ".join(parts))
        tasks.sort(key=lambda it: (STATUS_ORDER.index(it[1].get("status", DEFAULT_STATUS)),
                                   priority_rank(it[1])))
        for i, t in tasks:
            status = t.get("status", DEFAULT_STATUS)
            rid = self.e_tasks.insert("", "end", text=("\u2713 " if status == DONE else "\u25cf ")
                                      + (t.get("title") or "(untitled)") + jira_mark(t),
                                      tags=(STATUSES[status][1],))
            self.e_task_rows[rid] = (project, i)

    def _epic_set(self, **fields):
        if self.e_sel:
            self.store.set_epic_meta(*self.e_sel, **fields)
            self._refresh_epics()
            self._status(f"Saved epic '{self.e_sel[1]}'.")

    def _epic_set_later(self, **fields):
        """Typing in epic fields: save shortly after you pause."""
        self._epic_pending.update(fields)
        if self._epic_job:
            self.after_cancel(self._epic_job)
        self._epic_job = self.after(AUTOSAVE_DELAY, self._flush_epic)

    def _flush_epic(self):
        if self._epic_job:
            self.after_cancel(self._epic_job)
            self._epic_job = None
        if self._epic_pending and self.e_sel:
            self.store.set_epic_meta(*self.e_sel, **self._epic_pending)
            self._status(f"Saved epic '{self.e_sel[1]}'.")
            self._refresh_epics()
        self._epic_pending = {}

    def _epic_rename(self):
        if not self.e_sel:
            return
        project, old = self.e_sel
        new = self.e_name.get().strip()
        if not new or new == old:
            return
        self._flush_epic()
        if not self.store.rename_name(project, "epic", old, new):
            self._status(f"An epic called '{new}' already exists in '{project}'.")
            return
        self.e_sel = (project, new)
        self._refresh_choices()
        if self.active_project == project and self.f_epic.get() == old:
            self.f_epic.set(new)
        self._refresh_views()
        self._status(f"Renamed epic '{old}' to '{new}'.")

    def _epic_move(self):
        """Attach the epic to another project (its tasks come along)."""
        if not self.e_sel:
            return
        src, name = self.e_sel
        dst = self.e_project.get()
        if not dst or dst == src:
            return
        n = self.store.usage(src, "epic", name)
        if not messagebox.askyesno(
                "Attach epic to project",
                f"Attach epic '{name}' to '{dst}'?" +
                (f"\n\nIts {n} task(s) in '{src}' move to '{dst}' too." if n else "")):
            self.e_project.set(src)
            return
        self._flush_epic()
        editing_id = None
        if self.active_project is not None and self.editing_index is not None:
            editing_id = self.store.projects[self.active_project][self.editing_index].get("id")
        self._flush_autosave()
        self.store.move_epic(src, name, dst)
        if editing_id:  # keep the form pointed at the same task
            found = self._find_task(editing_id)
            if found:
                self.active_project, self.editing_index = found
        self.e_sel = (dst, name)
        self._refresh_choices()
        self._refresh_views()
        self._load_epic(self.e_sel)
        self._status(f"Epic '{name}' is now in '{dst}'" + (f" with its {n} task(s)." if n else "."))

    def _epic_new(self):
        name = self.e_new_name.get().strip()
        project = (self.e_filter.get() if self.e_filter.get() != ALL_PROJECTS
                   else (self.e_sel[0] if self.e_sel else self.active_project))
        if not project:
            self._status("Pick a project (top left) to add the epic to.")
            return
        if not name:
            self._status("Type a name for the new epic.")
            self.e_new_name.focus_set()
            return
        if not self.store.add_name(project, "epic", name):
            self._status(f"Epic '{name}' already exists in '{project}'.")
            return
        self.e_new_name.delete(0, "end")
        self.e_sel = (project, name)
        self._refresh_choices()
        self._refresh_views()
        self._load_epic(self.e_sel)
        self._status(f"Added epic '{name}' to '{project}'.")

    def _epic_delete(self):
        if not self.e_sel:
            return
        project, name = self.e_sel
        n = self.store.usage(project, "epic", name)
        if not messagebox.askyesno(
                "Delete epic", f"Delete epic '{name}'?" +
                (f"\n\nIts {n} task(s) stay, just without an epic." if n else "")):
            return
        self._epic_pending = {}
        self.store.delete_name(project, "epic", name)
        if self.active_project == project and self.f_epic.get() == name:
            self.f_epic.set("")
        self.e_sel = None
        self._refresh_choices()
        self._refresh_views()
        self._load_epic(None)
        self._status(f"Deleted epic '{name}'.")

    def _epic_new_task(self):
        if not self.e_sel:
            return
        project, name = self.e_sel
        sprint = self.store.epic_meta(project, name).get("sprint", "")
        self.notebook.select(0)
        self._select_project(project)
        self.f_epic.set(name)
        if sprint:
            self.f_sprint.set(sprint)
        self.f_title.focus_set()
        self._status(f"New task in epic '{name}' - type a title (it saves automatically).")

    def _epic_open_task(self):
        sel = self.e_tasks.selection()
        if sel and sel[0] in self.e_task_rows:
            project, index = self.e_task_rows[sel[0]]
            self._select_task(project, index)

    def _epic_open_jira(self):
        if not self.e_sel:
            return
        meta = self.store.epic_meta(*self.e_sel)
        key = self.e_jira.get().strip()
        url = meta.get("jira_url") if meta.get("jira_key") == key else ""
        if not url and key:
            base = load_jira_config().get("base_url", "")
            url = f"{base}/browse/{key}" if base else ""
        if url:
            self._open_link(url)
        else:
            self._status("No Jira link for this epic - set a Jira key (and connect Jira).")

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
        tasks = [r for r in sel if self.node_meta.get(r, ("",))[0] == "task"]
        if len(tasks) > 1:  # multi-select: leave the form alone, offer the drag
            self._status(f"{len(tasks)} tasks selected - drag them onto a project or "
                         "subgroup to move them together.")
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
            self._select_task(*self.row_meta[sel[0]])

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
        if self._autosave_job:
            self.after_cancel(self._autosave_job)
            self._autosave_job = None
        self._autosave_target = None
        if (self.j_paste.get("1.0", "end").strip() and self.active_project
                and not self.save_btn.instate(["disabled"])):
            if not self._jira_attach():
                return  # bad XML: leave it in the box so it can be fixed
        if not self.active_project or self.save_btn.instate(["disabled"]):
            self._status("Select a project first.")
            return
        project = self.active_project
        title = self.f_title.get().strip()
        if not title:
            self._status("Title is required.")
            self.f_title.focus_set()
            return

        task = self._task_from_form()

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

    # -- sharing with coworkers -------------------------------------------
    def _prompt(self, title, message, initial, on_ok):
        """Ask for a line of text in a panel drawn inside the main window.

        (tkinter's simpledialog opens a separate window that breaks on some
        macOS Tk builds, so all typing stays in the app.)"""
        self._close_prompt()
        panel = self._prompt_panel = tk.Frame(
            self, bg="#ffffff", highlightthickness=2, highlightbackground=ACCENT,
            padx=14, pady=12)
        tk.Label(panel, text=title, bg="#ffffff", fg="#222",
                 font=("", 11, "bold")).pack(anchor="w")
        tk.Label(panel, text=message, bg="#ffffff", fg="#555", wraplength=240,
                 justify="left").pack(anchor="w", pady=(2, 8))
        entry = ttk.Entry(panel, width=30)
        entry.insert(0, initial)
        entry.pack(fill="x")
        btns = tk.Frame(panel, bg="#ffffff")
        btns.pack(fill="x", pady=(10, 0))

        def ok(event=None):
            value = entry.get().strip()
            self._close_prompt()
            on_ok(value)
            return "break"

        ttk.Button(btns, text="OK", command=ok).pack(side="right")
        ttk.Button(btns, text="Cancel", command=self._close_prompt).pack(side="right", padx=(0, 6))
        entry.bind("<Return>", ok)
        entry.bind("<Escape>", lambda e: (self._close_prompt(), "break")[1])
        panel.place(relx=0.5, rely=0.12, anchor="n")
        panel.lift()
        entry.focus_set()
        entry.select_range(0, "end")

    def _close_prompt(self):
        panel = getattr(self, "_prompt_panel", None)
        if panel is not None:
            panel.destroy()
            self._prompt_panel = None

    def _ask_name(self, then=None):
        """Settings > Your name. `then` runs after a name is saved."""
        def save(name):
            if not name:
                self._status("Name not changed.")
                return
            self.settings["user_name"] = name
            save_settings(self.settings)
            self._status(f"Your name is set to '{name}'.")
            if then:
                then()

        self._prompt("Your name", "Coworkers see this on tasks you share with them.",
                     self.settings.get("user_name", ""), save)

    def _share_scope(self):
        """(project, [task indexes], label) for what's selected, or None."""
        def folder(project, group):
            idx = [i for i, t in enumerate(self.store.projects[project])
                   if not group or t.get("group") == group]
            return project, idx, group or project

        def one(project, index):
            return project, [index], self.store.projects[project][index].get("title", "task")

        if self.view == "Buddy":
            sel = self.b_tree.selection()
            if sel and sel[0] in self.b_meta:
                return one(*self.b_meta[sel[0]])
            if sel and sel[0] in self.b_folder_rows:
                return folder(*self.b_folder_rows[sel[0]])
        elif self.view == "Manager":
            sel = self.tree.selection()
            meta = self.node_meta.get(sel[0]) if sel else None
            if meta and meta[0] == "task":
                return one(meta[1], meta[2])
            if meta and meta[0] == "group":
                return folder(meta[1], meta[2])
            if meta and meta[0] == "project":
                return folder(meta[1], "")
        if self.active_project is not None and self.editing_index is not None:
            return one(self.active_project, self.editing_index)
        return None

    def _share(self):
        scope = self._share_scope()
        if not scope or not scope[1]:
            self._status("Select a task, subgroup or project to share.")
            return
        project, indexes, label = scope
        name = self.settings.get("user_name")
        if not name:
            # Ask in-app, then carry on sharing once it's set.
            self._ask_name(then=self._share)
            return
        # Rank is board order on this machine only; everything else travels.
        tasks = [{k: v for k, v in self.store.projects[project][i].items() if k != "rank"}
                 for i in indexes]
        data = {
            "format": SHARE_FORMAT,
            "version": SHARE_VERSION,
            "from": name,
            "exported": now_stamp(),
            "project": project,
            "lists": {kind: [n for n in self.store.names(project, kind)
                             if any(t.get(kind) == n for t in tasks)]
                      for kind in GROUP_KINDS},
            "tasks": tasks,
            "field_defs": [d for d in self.store.fields
                           if any(d["name"] in t.get("fields", {}) for t in tasks)],
        }
        default = re.sub(r'[\\/:*?"<>|]+', "-", f"{project} - {label}"
                         if label != project else project).strip() + ".json"
        path = filedialog.asksaveasfilename(
            title="Share tasks", defaultextension=".json", initialfile=default,
            filetypes=[("Shared tasks", "*.json")],
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        n = len(tasks)
        self._status(f"Shared {n} task{'s' if n != 1 else ''} to {path} - send that file to a coworker.")

    def _import_shared(self):
        path = filedialog.askopenfilename(
            title="Import shared tasks",
            filetypes=[("Shared tasks or Jira XML", "*.json *.xml"), ("All files", "*.*")],
        )
        if not path:
            return
        if path.lower().endswith(".xml"):
            self._import_jira(path)
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError) as exc:
            messagebox.showerror("Import", f"Couldn't read that file:\n{exc}")
            return
        if (not isinstance(data, dict) or data.get("format") != SHARE_FORMAT
                or not isinstance(data.get("tasks"), list)):
            messagebox.showerror("Import", "That file isn't a shared To-Do task export.")
            return
        self._import_data(data)

    def _import_jira(self, path=None):
        """Create tasks from a Jira XML file. Tickets you already have as
        tasks only get their attached Jira copy refreshed - your task's own
        fields are never overwritten."""
        path = path or filedialog.askopenfilename(
            title="Import Jira XML export",
            filetypes=[("Jira XML export", "*.xml"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                text = f.read()
        except OSError as exc:
            messagebox.showerror("Import Jira XML", f"Couldn't read that file:\n{exc}")
            return
        self._import_jira_text(text)

    def _import_jira_text(self, text, create=True):
        """Apply Jira XML (one or many issues): refresh tickets you already
        have; with create=True, add the rest as new tasks."""
        try:
            issues = parse_jira_issues(text)
        except (ValueError, ET.ParseError) as exc:
            messagebox.showerror("Import Jira XML", f"Couldn't read that Jira export:\n{exc}")
            return

        by_key = {}
        for project, tasks in self.store.projects.items():
            for i, t in enumerate(tasks):
                key = (t.get("jira") or {}).get("key") or t.get("jira_ref")
                if key:
                    by_key.setdefault(key, []).append((project, i))
        refreshed, by_project = [], {}
        for issue in issues:
            if issue["key"] in by_key:
                for project, i in by_key[issue["key"]]:  # every task linked to it
                    task = self.store.projects[project][i]
                    task.update(jira_updates(task, issue, task.get("jira")))
                    for kind in ("epic", "sprint"):
                        names = self.store.groups[kind].setdefault(project, [])
                        if issue.get(kind) and issue[kind] not in names:
                            names.append(issue[kind])
                    task["jira"] = issue
                refreshed.append(issue["key"])
            elif create:
                by_project.setdefault(IMPORTED_PROJECT, []).append(jira_task(issue))
        if refreshed:
            self.store.save()
        for project, tasks in by_project.items():
            lists = {kind: sorted({t[kind] for t in tasks if t.get(kind)}) for kind in GROUP_KINDS}
            self._import_data({"project": project, "from": "Jira", "lists": lists,
                               "tasks": tasks}, replace=False)
        if refreshed and not by_project:
            self._refresh_views()
            if self.editing_index is not None:  # it may be the one on screen
                self._load_task_into_form(self.store.projects[self.active_project][self.editing_index])
        created = sum(len(t) for t in by_project.values())
        parts = []
        if created:
            parts.append(f"created {created} task{'s' if created != 1 else ''} in "
                         + ", ".join(repr(p) for p in by_project))
        if refreshed:
            parts.append("refreshed the Jira copy on " + ", ".join(refreshed[:5])
                         + (f" (+{len(refreshed) - 5} more)" if len(refreshed) > 5 else ""))
        self.store.note_jira_epics()
        self._refresh_epics()
        if by_project:
            parts.append("drag them from 'Imported' into your projects")
        self._status("Jira import: " + ("; ".join(parts) or "nothing new") + ".")

    def _import_data(self, data, replace=None):
        """Merge shared/imported tasks into a project. replace=None asks the
        user what to do with tasks that are already there."""
        project = str(data.get("project") or "Imported").strip()
        sender = str(data.get("from") or "a coworker").strip()
        incoming = [t for t in data["tasks"] if isinstance(t, dict) and t.get("title")]
        if not incoming:
            self._status("That file has no tasks in it.")
            return

        if project not in self.store.projects:
            self.store.add_project(project)
        lists = data.get("lists") if isinstance(data.get("lists"), dict) else {}
        for kind in GROUP_KINDS:
            names = self.store.groups[kind].setdefault(project, [])
            for name in lists.get(kind, []):
                if isinstance(name, str) and name and name not in names:
                    names.append(name)

        tasks = self.store.projects[project]
        existing = {t.get("id"): i for i, t in enumerate(tasks) if t.get("id")}
        dupes = [t for t in incoming if t.get("id") in existing]
        if replace is None:
            replace = bool(dupes) and messagebox.askyesno(
                "Import",
                f"{len(dupes)} of these task(s) are already in '{project}'.\n\n"
                f"Replace them with this version?\n"
                f"(No keeps your copies and skips them.)",
            )

        # Create any custom fields this profile doesn't have yet.
        types = {d.get("name"): d.get("type", "text") for d in data.get("field_defs", [])
                 if isinstance(d, dict)}
        for raw in incoming:
            if not isinstance(raw.get("fields"), dict):
                raw["fields"] = {}
            raw["fields"] = {str(k): str(v) for k, v in raw["fields"].items() if str(k).strip()}
            for name, value in raw["fields"].items():
                long_value = "\n" in value or len(value) > 80
                self.store.ensure_field(name, types.get(name) or ("long" if long_value else "text"),
                                        source=sender)

        known = {key for key, _ in TASK_FIELDS} | {"id", "fields", "jira"}
        added = replaced = 0
        for raw in incoming:
            task = new_task(str(raw["title"]))
            task.update({k: v for k, v in raw.items() if k in known})
            if task.get("status") not in STATUSES:
                task["status"] = DEFAULT_STATUS
            if not isinstance(task.get("links"), list):
                task["links"] = []
            task["shared_by"] = sender
            for kind in GROUP_KINDS:
                names = self.store.groups[kind][project]
                if task.get(kind) and task[kind] not in names:
                    names.append(task[kind])
            if task.get("id") in existing:
                if not replace:
                    continue
                i = existing[task["id"]]
                task["rank"] = tasks[i].get("rank", 0)
                tasks[i] = task
                replaced += 1
            else:
                self.store.add_task(project, task)
                added += 1
        self.store.save()

        self.active_project, self.editing_index = project, None
        self._clear_form()
        self._refresh_views()
        if self.view == "Buddy":
            self._refresh_buddy()
        parts = [f"{added} new"] + ([f"{replaced} updated"] if replaced else [])
        self._status(f"Imported {' + '.join(parts)} task(s) from {sender} into '{project}'.")

    # -- export -----------------------------------------------------------
    def _rows_for_export(self):
        extra = [d["name"] for d in self.store.fields]
        headers = ["Project"] + [label for _, label in TASK_FIELDS] + extra
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
                row += [task.get("fields", {}).get(name, "") for name in extra]
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
    init_profile()
    store = Store(DATA_FILE)
    app = App(store)
    app.mainloop()


if __name__ == "__main__":
    main()
