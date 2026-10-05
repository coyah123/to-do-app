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

## Features

- Projects: add / delete (left panel).
- Tasks under each project with fields:
  - Title, Description
  - Added Date (auto-stamped from system time)
  - Due Date
  - Completed Date (auto-stamped when you mark complete)
  - "Jira ticket made?" checkbox — when on, paste the ticket link/number
  - Blockers
  - Notes
- Double-click a task to edit it.

## Data & privacy

- All data lives in **`todo-data.json`** next to the app.
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
