#!/bin/bash
# Explicit check or export; outputs are only written after confirmation.
set -euo pipefail
pause() { printf '\nPress Return to close this window... '; read -r _answer || true; }
trap 'status=$?; if (( status != 0 )); then printf "\nScheduler failed (exit %s). Check the message above.\n" "$status" >&2; fi; pause' EXIT
here=$(cd -- "$(dirname -- "$0")" && pwd -P) || exit 1
if [[ -f "$here/main.py" && -x "$here/.venv/bin/python" ]]; then project="$here"
else project="$HOME/Desktop/BEMS Scheduler"; fi
[[ -f "$project/main.py" && -x "$project/.venv/bin/python" ]] || { printf 'Installed project not found at %s. Run Install on Desktop.command first.\n' "$project" >&2; exit 1; }
printf 'Config file (default: %s/config.json; or config_f26b3.json): ' "$project"
read -r chosen || exit 1
config=${chosen:-"$project/config.json"}
if [[ "$config" != /* ]]; then config="$project/$config"; fi
[[ -f "$config" ]] || { printf 'Config file not found: %s\n' "$config" >&2; exit 1; }
printf '\n1) Check only: solve and validate without exporting [default]\n2) Full run: export files to the configured paths (may overwrite existing outputs)\nChoose 1 or 2: '
read -r mode || exit 1
cd -- "$project"
case "$mode" in
  ''|1) printf '\nRunning check only with %s\n\n' "$config"; "$project/.venv/bin/python" "$project/main.py" --check "$config" ;;
  2) printf '\nFull run may overwrite files at the output paths in %s.\nType EXPORT to continue: ' "$config"
     read -r confirm || exit 1
     [[ "$confirm" == EXPORT ]] || { printf 'Canceled; no schedule was run.\n'; exit 0; }
     "$project/.venv/bin/python" "$project/main.py" "$config" ;;
  *) printf 'Unrecognized choice. Nothing was run.\n' >&2; exit 1 ;;
esac
