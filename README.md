# To-Do Tracker

A lightweight, offline task tracker organized by project. Pure Python
standard library — no pip installs required to run it, so it works on an
airgapped work machine.

## Running it

Runs on **Windows, macOS, and Linux** — it's pure Python standard library.

- **Windows:** double-click `Run To-Do.bat` (launches with no console window).
- **macOS / Linux:** `chmod +x run-todo.sh` once, then double-click or run
  `./run-todo.sh`.
- **Any OS, from a terminal:** `python3 todo_app.py`

Requires Python 3 with tkinter:
- Windows / macOS: tkinter ships with the standard python.org installer.
- Linux: if you get `ModuleNotFoundError: No module named 'tkinter'`, install
  it via your package manager, e.g. `sudo apt install python3-tk`.

## Views

Switch with the **View** menu (or Ctrl+1 / Ctrl+2). The app reopens in the
view chosen under **View → Default view** (Buddy unless you change it), and
each view remembers its own window size/position.

- **Buddy** (default) — a small sticky note that opens in the top-right
  corner and stays on top of other windows (toggle with the pin button).
  Shows every project as a bold, collapsible folder (with its subgroups as
  sub-folders) and its tasks
  underneath — a colored dot marks each task's status, most active first,
  with due dates (overdue ones bold and marked `!`). Filter by sprint at the
  top. Select a project (or one of its tasks) and type in the box at the
  bottom to quick-add a task there. Right-click a task to change its status;
  double-click to open it in Manager.
- **Manager** — the full editor described below.
- **Visual Planner** — a Trello-style board. Pick what the columns are
  (Status / Subgroup / Epic / Sprint / Project), filter by project, epic and sprint, and
  choose the card order (Manual / Due date / Created / Title).
  - **Drag** a card to another column to change that field — status, epic,
    sprint, or even move it to another project. In Manual order you can
    also drag cards up/down to reorder; the order is saved.
  - **Click** a card to edit it in the Task Details panel on the right.
  - **+ Add card** at the bottom of a column starts a new task with that
    column's value filled in.

## Features

- Three resizable panes: project/task tree, task table, task details form.
- Projects: add / delete (left panel).
- Projects and subgroups share one add box at the top of the left tree:
  - Nothing selected: it says **New project** — type a name, **+ Project**.
  - A project (or anything in it) selected: it says **New subgroup in
    'Project'** — type a name, **+ Subgroup**. Subgroups only exist inside a
    project. Click empty space in the tree, press Esc, or click "new project
    instead" to go back to adding projects.
- Subgroups are sub-folders inside a project for related tasks (e.g. an
  app), shown indented in italics. **Drag tasks** onto a subgroup to move
  them in, onto the project to take them out, or onto another task to join
  its subgroup. Dragging onto a different project moves the task there.
  Selecting a subgroup then **New Task** creates the task inside it.
- Tasks under each project with fields:
  - Title, Description
  - Status (dropdown, color-coded): Not started (grey), In Progress (purple),
    Blocked (red), Waiting for approval (orange), Done (green)
  - Subgroup, Epic and Sprint — picked from the project's lists
  - Created Date (auto-stamped from system time)
  - Due Date
  - Completed Date (auto-stamped when status is set to Done)
  - "Jira ticket made?" checkbox — when on, paste the ticket link/number
  - Blockers, Notes
  - Links — any number, each with a title, description and URL. The title
    is a hyperlink: click it to open the link in your browser. Bare
    addresses like `example.com/page` get `https://` added. Use "edit" / ✕
    to change or remove one.
- Task table: group by Subgroup / Epic / Sprint / Status, click a column header
  to sort, "Hide Done" filter. Click a task in the tree or table to edit it.
- Subgroups, Epics & Sprints tab: pick a project, then add / rename / delete its
  subgroups, epics and sprints. Renaming updates every task using it; deleting
  leaves those tasks unassigned. Shows task and open counts for each.
- Shortcuts: Ctrl+S save, Ctrl+N new task.

## Custom fields

Each task has a **Fields** tab next to **Details** for extra data that
doesn't fit the standard fields — e.g. Priority, Customer, Story Points.
Pick an existing field from the box at the bottom (or type a new name) and
click **Add field**; a new name becomes available on every task. Shared
tasks from a coworker create any fields you don't have yet. Custom fields
get their own columns in CSV/XLSX exports.

## Sharing tasks with coworkers

1. Set your name once: **Settings → Your name…** (saved in your local
   `profile/settings.json`).
2. Select a task, a subgroup or a whole project (in Manager's tree, Buddy,
   or the task open in the form) and choose **File → Share selected…** (or
   the **Share…** button in Manager / Buddy's right-click menu). This saves a
   `.json` file — send it to your coworker.
3. They choose **File → Import shared tasks…**. The tasks land in the project
   with the same name (created if needed), with their subgroups, epics,
   sprints and links, marked **From: (your name)**. From then on they're
   ordinary tasks they can edit however they like.

Every task carries a hidden ID, so if you send an updated version of
something they already imported, they're asked whether to replace their
copy or keep it.

## Jira tickets

A Jira ticket is **attached to a task for reference — it never overwrites
your task.** The only things from a ticket that enter your own data are its
**epic** and **sprint**, so you can filter by them.

**Attach a ticket to a task**

1. In Jira, open the issue and use **Export → XML**; copy the XML (or save
   the `.xml` file).
2. In the app, click your task and open the **Jira** tab in Task Details.
3. Paste the XML into the box and press **Ctrl+Enter** or **Attach** (or
   **Load .xml file…**). **Save** also attaches anything left in the box.

The Jira tab then shows the ticket read-only: the key (click it to open the
ticket in your browser), status, type, priority, assignee, reporter, dates,
labels, components, versions, linked issues, any Jira custom fields, the
Jira description and the comments. Paste a newer export any time to refresh
it; **Remove** detaches it. The tab reads **Jira ✓** when a ticket is
attached.

**What it does to your data — only epic and sprint**

- The ticket's epic and sprint are added to the project's epic/sprint lists.
- If your task has no epic/sprint yet, it takes the ticket's. If you've set
  your own, yours is kept. (If the value originally came from the ticket and
  the ticket later moves to a new sprint, refreshing follows it.)
- Your title, description, status, dates, notes, links and custom fields are
  never touched.
- In the Jira tab, the **Sprint** and **Epic** values are links: click one to
  open the Visual Planner filtered to that sprint/epic in the project.

**Create tasks from a Jira export**

**File → Import Jira XML…** (the regular **Import…** picker accepts `.xml`
too) creates a task per issue — one ticket or a whole search export — in the
selected project, or a project named after the Jira project if none is
selected. New tasks get only the title, epic and sprint; the full ticket is
attached in the Jira tab. Tickets you already have as tasks aren't
duplicated or overwritten — only their attached copy (and Jira-sourced
epic/sprint) is refreshed.

## Data & privacy

Everything personal lives in the **`profile/`** folder next to the app,
which is **gitignored** — nothing in it is pushed to GitHub:

| File | What's in it |
|---|---|
| `profile/data.json` | Your projects, subgroups, epics, sprints and tasks |
| `profile/settings.json` | Your name, default view, window positions, filters |
| `profile/fields.json` | Your custom field list (see "Custom fields") |

Older versions kept `todo-data.json` / `todo-settings.json` next to the
app; they're moved into `profile/` automatically the first time you run
this version. Exports (CSV/XLSX) are gitignored too.

## Exporting for Confluence

- **Export CSV** — always works (standard library). Every field is quoted,
  so commas in notes are safe.
- **Export XLSX** — needs the `openpyxl` package. The colored dot by the
  buttons shows status:
  - 🟢 green: openpyxl is installed, XLSX export enabled.
  - 🔴 red: openpyxl not found, XLSX button disabled. Install with
    `pip install openpyxl` if your environment allows it; otherwise use CSV.

Paste the resulting table straight into a Confluence page.
