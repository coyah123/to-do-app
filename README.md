# To-Do Tracker

A lightweight, offline task tracker organized by project. Pure Python
standard library — no pip installs required to run it, so it works on an
airgapped work machine. Runs on **Windows, macOS and Linux**.

- [Quick start](#quick-start)
- [Core ideas](#core-ideas)
- [Views](#views) — Buddy, Manager, Visual Planner
- [Working with tasks](#working-with-tasks) — Details, Fields and Jira tabs
- [Organizing: projects, subgroups, epics & sprints](#organizing-projects-subgroups-epics--sprints)
- [Sharing tasks with coworkers](#sharing-tasks-with-coworkers)
- [Jira tickets](#jira-tickets)
- [Exporting (CSV / XLSX for Confluence)](#exporting-csv--xlsx-for-confluence)
- [Menus & keyboard shortcuts](#menus--keyboard-shortcuts)
- [Your data: the `profile/` folder](#your-data-the-profile-folder)
- [Troubleshooting](#troubleshooting)

---

## Quick start

1. **Start the app**
   - **Windows:** double-click `Run To-Do.bat` (no console window).
   - **macOS / Linux:** `chmod +x run-todo.sh` once, then double-click or run
     `./run-todo.sh`.
   - **Any OS, from a terminal:** `python3 todo_app.py`
2. It opens as **Buddy**, a small sticky note in the top-right corner. Use
   **View → Manager** (Ctrl+2) to set things up.
3. In Manager, type a project name in the box at the top left and click
   **+ Project**. With the project selected, type more names to add
   **subgroups**, then **New Task** to add tasks.
4. Set your name once under **Settings → Your name…** (used when you share
   tasks with coworkers).

**Requirements:** Python 3 with tkinter.
- Windows / macOS: tkinter ships with the standard python.org installer.
- Linux: if you see `ModuleNotFoundError: No module named 'tkinter'`, install
  it via your package manager, e.g. `sudo apt install python3-tk`.
- Optional: `openpyxl` for XLSX export (`pip install openpyxl`). Everything
  else works without it.

---

## Core ideas

| Thing | What it is |
|---|---|
| **Project** | Top-level container. Every task belongs to exactly one project. |
| **Subgroup** | A folder *inside* a project for related tasks (e.g. an app). Optional. |
| **Epic / Sprint** | Per-project labels you can group, filter and plan by. Optional. |
| **Task** | The to-do item: title, status, dates, notes, links, custom fields, an attached Jira ticket… |
| **Status** | Not started (grey) · In Progress (purple) · Blocked (red) · Waiting for approval (orange) · Done (green) |

Subgroups, epics and sprints are each a named list per project, managed in
Manager's **Subgroups, Epics & Sprints** tab. Statuses are color-coded
everywhere.

---

## Views

Switch views from the **View** menu or with **Ctrl+1 / Ctrl+2 / Ctrl+3**.
Each view remembers its own window size and position.

**Default view:** **View → Default view** picks the view the app opens in
(Buddy unless you change it). The list there includes every view, so new
views show up automatically.

### Buddy (Ctrl+1) — the all-day sticky note

A small window that opens in the **top-right corner** and **stays on top**
of other windows (toggle with the 📌 button or **View → Buddy always on
top**).

- Every project is a **bold, collapsible** folder, with its **subgroups** as
  sub-folders and tasks underneath. Collapsed folders stay collapsed between
  sessions.
- A **colored dot** on the left of each task shows its status; text stays
  black. Most active work is listed first (In Progress, Blocked, Waiting,
  Not started), then by due date.
- Due dates show on the right as `MM-DD`; **overdue** tasks are bold and
  marked `!`.
- **Sprint filter** at the top (All sprints / no sprint / a specific
  sprint) and a **Hide done** checkbox.
- The summary line counts what's open (e.g. `2 in progress · 1 blocked`).
- **Quick-add:** select a project, subgroup or task, then type in the box
  at the bottom and press Enter. The label beside it shows where the task
  will go (e.g. `MDA / Web app ▸`). When a sprint filter is active, new
  tasks get that sprint so they stay visible.
- **Right-click a task:** set its status, **Open in Manager**, or **Share
  task…**. **Double-click** opens it in Manager.

### Manager (Ctrl+2) — the full editor

Three resizable panes:

1. **Left — Projects & Tasks tree.** Projects (📁, bold) contain subgroups
   (📂, italic) and tasks. The add box at the top creates projects or
   subgroups depending on what's selected (see
   [Organizing](#organizing-projects-subgroups-epics--sprints)). **Drag
   tasks** to reorganize them. Buttons: **New Task**, **Del Task**,
   **Delete Project**.
2. **Center — two tabs:**
   - **Tasks:** a table of the selected project's tasks (Title, Status,
     Epic, Sprint, Due, Created). **Group by** Subgroup / Epic / Sprint /
     Status / None, click a column header to sort (click again to reverse),
     and **Hide Done**. The header also holds **Share… / Import…** and
     **Export CSV / Export XLSX**.
   - **Subgroups, Epics & Sprints:** pick a project, then add, rename or
     delete its subgroups, epics and sprints. Each list shows how many tasks
     (and open tasks) use each name. Renaming updates every task using it;
     deleting leaves those tasks unassigned.
3. **Right — Task Details:** the task form (see
   [Working with tasks](#working-with-tasks)).

### Visual Planner (Ctrl+3) — a Trello-style board

- **Columns by:** Status, Subgroup, Epic, Sprint or Project. Status columns
  have a colored bar; every column shows a count.
- **Filters:** Project (or All projects), Epic, Sprint (each with a
  "(none)" option) and **Hide Done**.
- **Order:** Manual (you arrange cards), Due date, Created or Title.
- **Cards** show a status-colored stripe, the title, status, due date
  (overdue in red with `!`), and tags for project, subgroup, epic, sprint,
  Jira ref, link count (🔗) and who shared it — leaving out whatever the
  columns already show.
- **Drag a card** to another column to change that field: status (moving to
  Done stamps the completed date), subgroup, epic, sprint, or — with
  columns by Project — move the task to another project. In **Manual**
  order you can also drag cards up and down; a blue line shows where it
  will land and the order is saved.
- **Click a card** to edit it in Task Details on the right.
- **+ Add card** at the bottom of a column starts a new task with that
  column's value filled in. **Delete card** (top right) deletes the task
  open in the form.
- Mouse wheel scrolls; Shift + wheel scrolls sideways; the board
  auto-scrolls while you drag near an edge.

---

## Working with tasks

Click a task (in any view) to open it in **Task Details**. **Save**
(Ctrl+S) writes your changes; **New** (Ctrl+N) starts a fresh task in the
selected project (and subgroup, if one is selected). The form has three
tabs.

### Details tab

| Field | Notes |
|---|---|
| **Title** | Required. |
| **Status** | Dropdown with a colored badge. Choosing **Done** stamps the Completed date; moving off Done clears it. |
| **Epic / Sprint / Subgroup** | Picked from the project's lists (blank = none). |
| **Due** | Free text; use `YYYY-MM-DD` so sorting and overdue marks work. |
| **Jira** | "Ticket made" checkbox + a box for the ticket number/link. |
| **Description, Blockers, Notes** | Free text; these boxes grow with the window. |
| **Links** | Any number of links, each with a **title**, **description** and **URL**. The title is a hyperlink — click it to open in your browser (hover shows the URL in the status bar). Bare addresses like `example.com/page` get `https://` added. **edit** loads a link back for changes; **✕** removes it. On a saved task, link changes are stored immediately. |
| **Created / Completed** | Stamped automatically from the system clock. |
| **From** | Shown when the task came from a coworker (or Jira). |

### Fields tab — custom fields

For extra data that doesn't fit the standard fields — e.g. Priority,
Customer, Story Points.

- Pick an existing field in the box at the bottom, or **type a new name**,
  and click **Add field**. A new name is created once and is then available
  on every task.
- Edit values inline; long fields get a multi-line box. **✕** removes a
  field from this task. **Save** to keep changes.
- The tab label shows how many fields the task has, e.g. **Fields (3)**.
- Shared tasks from a coworker create any fields you don't have yet.
- Every custom field gets its own column in CSV/XLSX exports.

### Jira tab

Holds a read-only copy of a Jira ticket attached to the task. See
[Jira tickets](#jira-tickets).

---

## Organizing: projects, subgroups, epics & sprints

### Adding projects and subgroups

One add box sits at the top of Manager's tree. The line above it tells you
what it will create:

| Selected in the tree | Box says | Button |
|---|---|---|
| Nothing | **New project** | **+ Project** |
| A project, or a subgroup/task in it | **New subgroup in 'MDA'** | **+ Subgroup** |

To add a project while something is selected, click **new project
instead**, click empty space in the tree, or press **Esc** in the box. After
adding a project it stays selected, so you can type its subgroups straight
away. Subgroups can't exist outside a project.

### Moving tasks (drag and drop in Manager's tree)

- Drop on a **subgroup** → the task moves into it.
- Drop on the **project** row → it leaves its subgroup.
- Drop on **another task** → it joins that task's subgroup.
- Drop anywhere in a **different project** → the task moves to that
  project (its subgroup/epic/sprint names are added there if missing).

The target row is highlighted while you drag.

### Epics and sprints

Add, rename and delete them in **Subgroups, Epics & Sprints**, assign them
in the task form, and use them to group (Manager table), filter (Buddy,
Visual Planner) and plan (Visual Planner columns). Epics and sprints also
arrive automatically from attached [Jira tickets](#jira-tickets).

---

## Sharing tasks with coworkers

Everyone runs their own copy of the app with their own data. To hand tasks
to someone:

1. **Set your name** once: **Settings → Your name…** (stored in
   `profile/settings.json`). If you haven't, the app asks the first time you
   share.
2. **Select what to share:** a task, a subgroup (all its tasks) or a whole
   project — in Manager's tree, in Buddy, or the task open in the form.
3. **File → Share selected…** (or **Share…** in Manager's header, or
   **Share task…** in Buddy's right-click menu). Save the `.json` file it
   offers (e.g. `MDA - Web app.json`) and send it to your coworker.
4. **They import it:** **File → Import shared tasks…** (or **Import…**).

What happens on import:

- Tasks go into the project with the **same name**, created if they don't
  have it, with their subgroups, epics, sprints, links, custom fields and
  any attached Jira ticket.
- Each task shows **From: (your name)** in its form and "from …" on its
  Planner card. After that it's an ordinary task they can change however
  they like.
- Every task has a hidden ID. If you send an **updated version** of
  something they already have, they're asked whether to **replace** their
  copy or **keep** it.
- Board order (Visual Planner manual ranking) isn't shared — it's personal.

<details>
<summary>Share file format</summary>

```json
{
  "format": "todo-tracker-share",
  "version": 1,
  "from": "Connor",
  "exported": "2026-10-05 12:00",
  "project": "MDA",
  "lists": {"group": ["Web app"], "epic": ["Authentication"], "sprint": ["Sprint 1"]},
  "field_defs": [{"name": "Priority", "type": "text", "source": "user"}],
  "tasks": [{"id": "…", "title": "Build login", "status": "In Progress", "...": "..."}]
}
```
</details>

---

## Jira tickets

A Jira ticket is **attached to a task** and kept as a read-only copy. It
also fills in your task's **epic, sprint, description, dates and links** —
but **never over something you wrote yourself** (see the rules below).

### Attach a ticket to a task

1. In Jira, open the issue and use **Export → XML**. Copy the XML (or save
   the `.xml` file).
2. In the app, click your task and open the **Jira** tab in Task Details.
3. Paste the XML into the box and press **Ctrl+Enter** or **Attach** (or
   use **Load .xml file…**). **Save** also attaches anything left in the box.

The Jira tab then shows the ticket **read-only**:

- The **key** (e.g. `MDA-12`) — click it to open the ticket in your browser.
- Status, type, priority, resolution, assignee, reporter, sprint, epic, due,
  created, updated and resolved dates.
- Labels, components, fix/affects versions, parent, subtasks, linked issues
  and any Jira custom fields (e.g. Story Points).
- The Jira **description** (HTML converted to readable text) and all
  **comments** with author and date.

Paste a newer export at any time to **refresh** it; **Remove** detaches it.
The tab reads **Jira ✓** when a ticket is attached. If the pasted text
isn't valid Jira XML, it stays in the box and the status bar says why.

### What it fills in on your task

| Your field | From the ticket |
|---|---|
| **Epic / Sprint** | The ticket's epic and sprint (also added to the project's lists, so you can filter by them) |
| **Description** | The Jira description (HTML converted to readable text) |
| **Due** | Jira's due date |
| **Completed** | Jira's resolved date — only if your task's status is **Done** |
| **Created** | Jira's created date — only for tasks *created* from a Jira import |
| **Links** | Every link in the XML: the ticket itself, the epic, parent, linked issues (e.g. "is blocked by MDA-9"), subtasks, attachments, and any links inside the description or comments |

**Your edits always win:**

- Epic, sprint, description and due date are filled in only when **your
  field is empty**, or when it **still holds what the ticket put there
  last time** (so a refreshed ticket — e.g. a new due date — updates it).
  As soon as you type your own value, refreshing the ticket leaves it alone.
  No duplicate descriptions: the ticket's text is never appended to yours.
- Links are **added**, never replaced. Links already on the task are
  skipped, and a link you **removed** after an earlier attach isn't added
  back when you refresh.
- Status, title, blockers, notes and custom fields are **never touched**.

The status bar says what was filled in, e.g. *"Attached Jira MDA-12 —
filled in description, due date, 8 links."*

In the Jira tab, the **Sprint** and **Epic** values are links: click one to
open the **Visual Planner filtered** to that sprint/epic in the project.

### Create tasks from a Jira export

**File → Import Jira XML…** (the regular **Import…** picker accepts `.xml`
too) creates one task per issue — a single ticket or a whole search export:

- Tasks go into the **selected project**, or a project named after the
  Jira project if none is selected.
- New tasks get the **title**, plus everything in the table above
  (epic, sprint, description, due date, created date, links); the full
  ticket is attached in the Jira tab.
- Tickets you **already have** as tasks aren't duplicated. Their attached
  copy is refreshed and the same "your edits win" rules apply.

> **Epic names:** Jira's XML only includes the epic's *key* (e.g. `MDA-3`),
> not its name, so that's what the epic is called. Rename it in **Subgroups,
> Epics & Sprints** if you like.

---

## Exporting (CSV / XLSX for Confluence)

Buttons in Manager's header export **every task in every project**:

- **Export CSV** — always works (standard library). Every field is quoted,
  so commas in notes are safe.
- **Export XLSX** — needs `openpyxl`. Status cells are color-filled. The dot
  by the buttons shows availability:
  - 🟢 green: openpyxl is installed, XLSX export enabled.
  - 🔴 red: openpyxl not found, XLSX button disabled. Install with
    `pip install openpyxl` if your environment allows it; otherwise use CSV.

Columns: Project, Title, Status, Subgroup, Epic, Sprint, Description,
Created/Due/Completed dates, Jira fields, Blockers, Notes, Links (one
`title: URL` per line), Shared By, then **one column per custom field**.
Paste the result straight into a Confluence page.

---

## Menus & keyboard shortcuts

| Menu | Items |
|---|---|
| **File** | Share selected… · Import shared tasks… · Import Jira XML… |
| **View** | Buddy · Manager · Visual Planner · **Default view** (one option per view) · Buddy always on top |
| **Settings** | Your name… |

On **macOS** use **Cmd** wherever this guide says Ctrl (Cmd+1, Cmd+S, …);
right-click is a two-finger click or **Ctrl+click**.

| Shortcut | Action |
|---|---|
| Ctrl+1 / Ctrl+2 / Ctrl+3 | Buddy / Manager / Visual Planner |
| Ctrl+S | Save the task in the form (Manager, Planner) |
| Ctrl+N | New task (Manager, Planner) · focus quick-add (Buddy) |
| Ctrl+Enter | Attach the pasted Jira XML (Jira tab) |
| Esc | In the add box: switch back to "New project" |
| Enter | Add project/subgroup, quick-add task, add link, add field |

---

## Your data: the `profile/` folder

Everything personal lives in **`profile/`** next to the app. The whole
folder is **gitignored** — nothing in it is ever pushed to GitHub.

| File | What's in it |
|---|---|
| `profile/data.json` | Your projects, subgroups, epics, sprints and tasks (incl. links, custom field values and attached Jira tickets) |
| `profile/settings.json` | Your name, default view, window sizes/positions, Buddy & Planner filters, collapsed folders |
| `profile/fields.json` | Your custom field list |

- **Upgrading from an older version:** `todo-data.json` and
  `todo-settings.json` next to the app are moved into `profile/`
  automatically the first time you run this version. Older data files are
  upgraded in place (statuses, IDs and new fields are filled in).
- **Backup:** copy the `profile/` folder. **Restore / move to another
  machine:** put it back next to `todo_app.py`.
- **Start fresh:** close the app and delete or rename `profile/`.
- Exports (`*.csv`, `*.xlsx`) are gitignored too. Share files (`.json`) are
  saved wherever you choose.

---

## Troubleshooting

- **The app doesn't show my latest changes / features.** An older copy may
  still be running. Close every To-Do window and start it again. (Running
  two copies at once means the last one to save wins.)
- **My tasks disappeared after updating.** Check `profile/data.json`. If an
  old window was still open during the upgrade, it may have saved to the
  old `todo-data.json` next to the app — close everything, then move that
  file into `profile/` as `data.json` (back up first).
- **XLSX export is greyed out.** `openpyxl` isn't installed — use CSV, or
  `pip install openpyxl`.
- **macOS: errors like `::tk::unsupported::MacWindowStyle` or an extra
  Python window when typing your name.** Fixed — every text prompt (e.g.
  **Settings → Your name…**) now opens as a panel *inside* the app window.
  Update to the latest version.
- **"That isn't Jira XML I can read."** Make sure you used Jira's
  **Export → XML** (not Word/printable) and copied the whole document.
- **Buddy is off-screen** (e.g. after unplugging a monitor). Close the app,
  remove the `geometry_Buddy` line from `profile/settings.json`, and start
  it again — Buddy returns to the top-right corner.
