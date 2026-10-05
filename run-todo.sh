#!/usr/bin/env bash
# Launches the To-Do app on macOS / Linux with no terminal needed.
# Make executable once:  chmod +x run-todo.sh
# Then double-click (or run ./run-todo.sh).
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# Prefer pythonw if present (rare on mac/linux); otherwise python3.
if command -v pythonw >/dev/null 2>&1; then
    exec pythonw "$DIR/todo_app.py"
elif command -v python3 >/dev/null 2>&1; then
    exec python3 "$DIR/todo_app.py"
else
    exec python "$DIR/todo_app.py"
fi
