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
  Shows every project as a bold, collapsible folder with its tasks
  underneath — a colored dot marks each task's status, most active first,
  with due dates (overdue ones bold and marked `!`). Filter by sprint at the
  top. Select a project (or one of its tasks) and type in the box at the
  bottom to quick-add a task there. Right-click a task to change its status;
  double-click to open it in Manager.
- **Manager** — the full editor described below.
- **Visual Planner** — a Trello-style board. Pick what the columns are
  (Status / Epic / Sprint / Project), filter by project, epic and sprint, and
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
- Tasks under each project with fields:
  - Title, Description
  - Status (dropdown, color-coded): Not started (grey), In Progress (purple),
    Blocked (red), Waiting for approval (orange), Done (green)
  - Epic and Sprint — picked from the project's lists
- Epics & Sprints tab: pick a project, then add / rename / delete its epics
  and sprints. Renaming updates every task using it; deleting leaves those
  tasks unassigned. Shows task and open counts for each.
  - Created Date (auto-stamped from system time)
  - Due Date
  - Completed Date (auto-stamped when status is set to Done)
  - "Jira ticket made?" checkbox — when on, paste the ticket link/number
  - Blockers, Notes
  - Links — any number, each with a title, description and URL. The title
    is a hyperlink: click it to open the link in your browser. Bare
    addresses like `example.com/page` get `https://` added. Use "edit" / ✕
    to change or remove one.
- Task table: group by Epic / Sprint / Status, click a column header to sort,
  "Hide Done" filter. Click a task in the tree or table to edit it.
- Shortcuts: Ctrl+S save, Ctrl+N new task.

## Data & privacy

- All data lives in **`todo-data.json`** next to the app.
- Window layout is saved per machine in `todo-settings.json`.
- That file (and any exports) are **gitignored**, so nothing is pushed to GitHub.

## Exporting for Confluence

- **Export CSV** — always works (standard library). Every field is quoted,
  so commas in notes are safe.
- **Export XLSX** — needs the `openpyxl` package. The colored dot by the
  buttons shows status:
  - 🟢 green: openpyxl is installed, XLSX export enabled.
  - 🔴 red: openpyxl not found, XLSX button disabled. Install with
    `pip install openpyxl` if your environment allows it; otherwise use CSV.

Paste the resulting table straight into a Confluence page.
