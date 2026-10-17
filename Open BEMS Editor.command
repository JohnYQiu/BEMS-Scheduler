#!/bin/bash
# Launch the local-only editor; no solver or external upload is started.
set -euo pipefail
pause() { printf '\nPress Return to close this window... '; read -r _answer || true; }
trap 'status=$?; if (( status != 0 )); then printf "\nEditor failed (exit %s). Check the message above.\n" "$status" >&2; fi; pause' EXIT
here=$(cd -- "$(dirname -- "$0")" && pwd -P) || exit 1
if [[ -f "$here/editor.py" && -x "$here/.venv/bin/python" ]]; then project="$here"
else project="$HOME/Desktop/BEMS Scheduler"; fi
[[ -f "$project/editor.py" && -x "$project/.venv/bin/python" ]] || { printf 'Installed project not found at %s. Run Install on Desktop.command first.\n' "$project" >&2; exit 1; }
printf 'Configuration to edit (default: %s/config.json): ' "$project"
read -r chosen || exit 1
config=${chosen:-"$project/config.json"}
if [[ "$config" != /* ]]; then config="$project/$config"; fi
[[ -f "$config" ]] || { printf 'Config file not found: %s\n' "$config" >&2; exit 1; }
cd -- "$project"
printf 'Starting the local editor. It opens your browser; leave this Terminal window open. Ctrl-C stops it.\n'
"$project/.venv/bin/python" "$project/editor.py" "$config"
