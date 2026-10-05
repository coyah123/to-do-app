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
| **Epic** | A body of work inside a project, with its own details (sprint, status, Jira key, description) — see the **Epics** tab. Optional. |
| **Sprint** | A per-project label you can group, filter and plan by. Optional. |
| **Task** | The to-do item: title, status, dates, notes, links, custom fields, an attached Jira ticket… |
| **Status** | Not started (grey) · In Progress (purple) · Blocked (red) · Waiting for approval (orange) · Done (green) |
| **Priority** | ⇈ Highest · ↑ High · = Medium · ↓ Low · ⇊ Lowest (Jira's levels), or none |

Subgroups, epics and sprints are each a named list per project, managed in
Manager's **Subgroups, Epics & Sprints** tab; epics also have their own
**Epics** tab. Statuses are color-coded everywhere. Tasks with a Jira ticket
attached show the **Jira sync icon** (two arrows chasing each other: 🗘 on
Windows, 🔄 on macOS, ⟳ on Linux) after their title in every view.

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
- A **colored dot** on the left of each task shows its status, followed by
  its **priority arrow** (red ⇈ Highest, ↑ High, orange = Medium, green ↓
  Low, ⇊ Lowest); text stays black.
- **Sort** dropdown: **Activity** (default — In Progress, Blocked, Waiting,
  Not started, then priority and due date), **Priority** (highest first),
  **Due date**, **Title** or **Created**.
- Due dates show on the right as `MM-DD`; **overdue** tasks are bold and
  marked `!`.
- **Sprint filter** at the top (All sprints / no sprint / a specific
  sprint) and a **Hide done** checkbox.
- **Links:** a task with links shows 🔗 and a count; click its arrow to
  **show/hide** them underneath, then click a link to open it in your
  browser. Buddy remembers which tasks you left expanded.
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
   (📂, italic) and tasks. **Ctrl/Shift-click** to select several tasks. The add box at the top creates projects or
   subgroups depending on what's selected (see
   [Organizing](#organizing-projects-subgroups-epics--sprints)). **Drag
   tasks** to reorganize them. Buttons: **New Task**, **Del Task**,
   **Delete Project**.
2. **Center — two tabs:**
   - **Tasks:** a table of the selected project's tasks (Title, Status,
     Priority, Epic, Sprint, Due, Created). **Group by** Project / Subgroup /
     Epic / Sprint / Status / Priority / None — **Project** shows *every*
     project's tasks grouped under each project (click any one to edit it);
     the other groupings show the selected project. **Sort by** Status, Priority, Due
     date, Created, Title, Epic or Sprint — or click a column header (click
     again to reverse). Sorting by Priority puts Highest first, then by
     status and due date; tasks without a priority go last. **Hide Done**. The header also holds **Share… / Import…** and
     **Export CSV / Export XLSX**.
   - **Epics:** every epic with its project, sprint, status and open/all
     task counts (filter by project, hide done ones). Click an epic to see
     and edit its details below — see [Epics](#epics).
   - **Subgroups, Epics & Sprints:** pick a project, then add, rename or
     delete its subgroups, epics and sprints. Each list shows how many tasks
     (and open tasks) use each name. Renaming updates every task using it;
     deleting leaves those tasks unassigned.
3. **Right — Task Details:** the task form (see
   [Working with tasks](#working-with-tasks)).

### Visual Planner (Ctrl+3) — a Trello-style board

- **Columns by:** Status, Priority, Subgroup, Epic, Sprint or Project. Status columns
  have a colored bar; every column shows a count.
- **Filters:** Project (or All projects), Epic, Sprint (each with a
  "(none)" option) and **Hide Done**.
- **Order:** Manual (you arrange cards), Priority (highest first), Due
  date, Created or Title.
- **Cards** show a status-colored stripe, the title, status, due date
  (overdue in red with `!`), priority, and tags for project, subgroup, epic, sprint,
  Jira ref, link count (🔗) and who shared it — leaving out whatever the
  columns already show.
- **Drag a card** to another column to change that field: status (moving to
  Done stamps the completed date), priority, subgroup, epic, sprint, or — with
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

Click a task (in any view) to open it in **Task Details**. **New** (Ctrl+N)
starts a fresh task in the selected project (and subgroup, if one is
selected). The form has three tabs.

**Everything saves automatically — no Save button needed.**

- Typing in any field (title, description, blockers, notes, due date, Jira
  number, custom fields) is saved about half a second after you pause.
- Dropdowns and checkboxes (status, priority, subgroup, epic, sprint, Jira
  "ticket made") save immediately. Links, custom fields added/removed and
  attached Jira tickets save as soon as you add them.
- Every view updates right away — the tree, table, Visual Planner cards and
  Buddy — and a small **✓ Saved** with the time appears next to the buttons.
- A **new task** is created as soon as you give it a title (until then the
  form shows "Add a title to save").
- Nothing is lost if you click another task, switch views or close the app
  straight after typing — the pending change is saved first.
- **Save** (Ctrl+S) is still there if you want to force a save.

### Details tab

| Field | Notes |
|---|---|
| **Title** | Required. |
| **Status** | Dropdown with a colored badge. Choosing **Done** stamps the Completed date; moving off Done clears it. |
| **Priority** | Highest / High / Medium / Low / Lowest (blank = none), shown with its colored arrow. |
| **Epic / Sprint / Subgroup** | Picked from the project's lists (blank = none). |
| **Due** | Free text; use `YYYY-MM-DD` so sorting and overdue marks work. |
| **Jira** | "Ticket made" checkbox + a box for the ticket number/link. Filled in automatically from an attached ticket; the ticket and its Jira links also appear as clickable **quick links** at the bottom of the form. |
| **Description, Blockers, Notes** | Free text; these boxes grow with the window. |
| **Links** | Any number of links, each with a **title**, **description** and **URL**. The title is a hyperlink — click it to open in your browser (hover shows the URL in the status bar). Bare addresses like `example.com/page` get `https://` added. **edit** loads a link back for changes; **✕** removes it. Link changes save immediately. |
| **Created / Completed** | Stamped automatically from the system clock. |
| **From** | Shown when the task came from a coworker (or Jira). |

### Fields tab — custom fields

For extra data that doesn't fit the standard fields — e.g. Priority,
Customer, Story Points.

- Pick an existing field in the box at the bottom, or **type a new name**,
  and click **Add field**. A new name is created once and is then available
  on every task.
- Edit values inline; long fields get a multi-line box. **✕** removes a
  field from this task. Changes save automatically.
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

Drag one task — or **select several with Ctrl/Shift-click and drag them
together**:

- Drop on a **subgroup** → the task moves into it.
- Drop on the **project** row → it leaves its subgroup.
- Drop on **another task** → it joins that task's subgroup.
- Drop anywhere in a **different project** → the task moves to that
  project (its subgroup/epic/sprint names are added there if missing).

The target row is highlighted while you drag.

### Epics

Manager's **Epics** tab (next to Tasks) treats each epic as something with
its own details:

- **The list** shows every epic: project, sprint, status (its own, or worked
  out from its tasks), and open / all tasks. Filter by **Project**, tick
  **Hide done**, or add one with **New epic** + **+ Epic**.
- **Click an epic** to edit it below (changes save automatically):

  | Field | |
  |---|---|
  | **Name** | Renaming updates every task in the epic |
  | **Project** | The project the epic belongs to. **Changing it attaches the epic to that project and moves its tasks there too** (you're asked first) — e.g. move a Jira epic out of *Imported* into your project |
  | **Sprint** | Attach the epic to one of the project's sprints |
  | **Status** | Set it yourself, or leave "(from its tasks)" |
  | **Jira** | The epic's Jira key; **Open ↗** opens it. Filled in automatically for epics that came from Jira |
  | **Description** | Notes about the epic |

- On the right: **the epic's tasks** (click one to open it) and **which
  sprints** the epic and its tasks are in.
- **+ New task in this epic** starts a task in the epic's project with the
  epic (and its sprint) already set. **Delete epic** removes it; its tasks
  stay, just without an epic.

### Sprints

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
| **Priority** | The ticket's priority. Jira Server names are converted: Blocker → Highest, Critical → High, Major → Medium, Minor → Low, Trivial → Lowest (also P1–P5) |
| **Jira (ticket made / number)** | "Ticket made" is ticked and the number is set to the ticket key, e.g. `MDA-12` |
| **Epic / Sprint** | The ticket's epic and sprint (also added to the project's lists, so you can filter by them) |
| **Description** | The Jira description (HTML converted to readable text) |
| **Due** | Jira's due date |
| **Completed** | Jira's resolved date — only if your task's status is **Done** |
| **Created** | Jira's created date — only for tasks *created* from a Jira import |
| **Links** | Every link in the XML: the ticket itself, the epic, parent, linked issues (e.g. "is blocked by MDA-9"), subtasks, attachments, and any links inside the description or comments |

**Your edits always win:**

- Jira number, priority, epic, sprint, description and due date are filled in only when **your
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

**Quick links bar:** whenever a task has a Jira ticket, a bar at the
**bottom of Task Details** (visible on every tab) shows the ticket key and
its related Jira links — epic, parent, linked issues, subtasks — e.g.
`Jira: MDA-12 · Epic MDA-3 · is blocked by MDA-9`. One click opens it in
your browser. If you typed a full Jira URL into the Jira field yourself
(without attaching XML), that shows there as a link too.

### Create tasks from a Jira export

**File → Import Jira XML…** (the regular **Import…** picker accepts `.xml`
too) creates one task per issue — a single ticket or a whole search export:

- New tasks always go into a project called **Imported**, so your own
  projects stay tidy. From there, drag them (several at once with
  Ctrl/Shift-click) into your projects and subgroups, or move a whole Jira
  epic with its tasks from the **Epics** tab.
- New tasks get the **title**, plus everything in the table above
  (epic, sprint, description, due date, created date, links); the full
  ticket is attached in the Jira tab.
- Tickets you **already have** as tasks aren't duplicated. Their attached
  copy is refreshed and the same "your edits win" rules apply.

### Optional: connect directly to Jira (personal access token)

Instead of exporting XML by hand, the app can fetch tickets itself. It's
**optional** and **read-only** (it never changes anything in Jira) — it
requests the same XML as **Export → XML**, so everything above applies.

**Set up — Settings → Jira connection…**

| Field | What to enter |
|---|---|
| **Jira address** | Your Jira's address, e.g. `https://jira.yourcompany.com` (custom domains are fine; pasting any ticket URL works too — the app keeps just the address) |
| **Type** | **Server / Data Center** for a *personal access token* (Profile → Personal Access Tokens in Jira), or **Cloud** (`*.atlassian.net`) for email + *API token* |
| **Email (Cloud)** | Only for Cloud |
| **Token** | Your personal access token / API token (shown as dots) |
| **Remember token** | Optional. Off = the token is kept in memory until you close the app. On = saved encrypted with your **Windows login** (Windows), or in a file only your user can read (macOS/Linux) |
| **CA file (optional)** | If your company uses its own HTTPS certificates and you get an SSL error, point this at the company CA certificate (`.pem`/`.crt`) |

Click **Test connection** ("✓ Connected as …"), then **Save**.
**Disconnect** removes the saved connection from this computer.

**Use it**

- **Jira tab → Ticket:** type a key like `MDA-12` (or paste the ticket URL)
  and click **Fetch from Jira** (or press Enter). Once attached, the button
  becomes **Refresh from Jira** to pull the latest version.
- **File → Import from Jira search…:** enter a JQL query (default
  `assignee = currentUser() AND resolution = Unresolved`). Matching tickets
  become tasks in **Imported**; ones you already have are refreshed wherever
  you've moved them.
- **File → Refresh all Jira tickets:** re-fetches every attached ticket in
  one go.
- **File → Gather Jira imports into 'Imported':** moves every task that a
  Jira import created back into *Imported* (handy if an import landed in
  the wrong place). Tasks you attached a ticket to yourself aren't moved.

All "your edits win" rules apply to fetched tickets. Network calls run in
the background, so the window never freezes; problems (wrong/expired token,
no permission, SSO login page, SSL, can't reach Jira) are explained in plain
words.

**Security notes:** the connection lives in `profile/jira.json`, which is
gitignored with the rest of `profile/`. Create the token with an expiry
date, use it only for this, and check your company's policy on personal
access tokens. Never paste a token into chats or tickets.

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

Columns: Project, Title, Status, Priority, Subgroup, Epic, Sprint, Description,
Created/Due/Completed dates, Jira fields, Blockers, Notes, Links (one
`title: URL` per line), Shared By, then **one column per custom field**.
Paste the result straight into a Confluence page.

---

## Menus & keyboard shortcuts

| Menu | Items |
|---|---|
| **File** | Share selected… · Import shared tasks… · Import Jira XML… · Import from Jira search… · Refresh all Jira tickets · Gather Jira imports into 'Imported' |
| **View** | Buddy · Manager · Visual Planner · **Default view** (one option per view) · Buddy always on top |
| **Settings** | Your name… · Jira connection… |

On **macOS** use **Cmd** wherever this guide says Ctrl (Cmd+1, Cmd+S, …);
right-click is a two-finger click or **Ctrl+click**.

| Shortcut | Action |
|---|---|
| Ctrl+1 / Ctrl+2 / Ctrl+3 | Buddy / Manager / Visual Planner |
| Ctrl+S | Save now (tasks also save automatically) |
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
| `profile/data.json` | Your projects, subgroups, epics (and their details), sprints and tasks (incl. links, custom field values and attached Jira tickets) |
| `profile/settings.json` | Your name, default view, window sizes/positions, Buddy & Planner filters, collapsed folders |
| `profile/fields.json` | Your custom field list |
| `profile/jira.json` | Optional Jira connection: address, type, and the token *only* if you chose "Remember" (encrypted on Windows) |

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
