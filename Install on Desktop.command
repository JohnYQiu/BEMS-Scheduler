#!/bin/bash
# Double-click from a downloaded/extracted BEMS-Scheduler folder. Never updates an existing installation.
set -euo pipefail

pause() { printf '\nPress Return to close this window... '; read -r _answer || true; }
fail() { printf '\nERROR: %s\n' "$*" >&2; exit 1; }
stage=''
finish() {
  status=$?
  trap - EXIT
  if [[ -n "$stage" && -d "$stage" ]]; then rm -rf -- "$stage"; fi
  if (( status != 0 )); then printf '\nInstallation did not complete. Your existing Desktop files were not changed.\n' >&2; fi
  pause
  exit "$status"
}
trap finish EXIT

source_dir=$(cd -- "$(dirname -- "$0")" && pwd -P) || fail 'Cannot locate the extracted project folder.'
desktop="$HOME/Desktop"
target="$desktop/BEMS Scheduler"
[[ -d "$desktop" ]] || fail "Desktop directory does not exist: $desktop"
[[ -f "$source_dir/main.py" && -f "$source_dir/editor.py" && -f "$source_dir/editor.html" && -f "$source_dir/requirements.txt" && -f "$source_dir/config.json" ]] || fail 'Required project files are missing. Extract the complete repository first.'
[[ ! -e "$target" && ! -L "$target" ]] || fail "The destination already exists: $target. Move or rename it yourself first; this installer will not overwrite it."
printf 'Installing a separate clean copy at:\n  %s\n\n' "$target"
printf 'Only program files, supplied configurations, documentation, and launchers are copied.\n'
printf 'Local inputs, outputs, backups, .git, .venv, and *.local.json are not copied.\n'
printf 'Python dependencies will be downloaded by pip from its configured package index. No sudo is used.\n\n'
printf 'Continue? [y/N] '; read -r answer || fail 'No answer received.'
[[ "$answer" == [yY] || "$answer" == [yY][eE][sS] ]] || { printf 'Canceled.\n'; exit 0; }

stage=$(mktemp -d "$desktop/.BEMS-Scheduler-install.XXXXXXXX") || fail 'Could not create a temporary installation directory on Desktop.'
# Explicit allowlist: never recursively copy arbitrary contents from an extracted folder.
for file in main.py editor.py editor.html rule_catalog.py forces.py configuration.py models.py output.py parse_form.py policy.py solver.py validation.py requirements.txt config.json config_f26b3.json README.md OFFICER_GUIDE.md .gitignore 'Install on Desktop.command' 'Open BEMS Editor.command' 'Run Scheduler.command'; do
  if [[ -f "$source_dir/$file" && ! -L "$source_dir/$file" ]]; then
    cp -p -- "$source_dir/$file" "$stage/$file" || fail "Could not copy $file"
  elif [[ "$file" != config_f26b3.json && "$file" != README.md && "$file" != OFFICER_GUIDE.md && "$file" != .gitignore ]]; then
    fail "Required source file is missing or is a symbolic link: $file"
  fi
done
if [[ -d "$source_dir/tests" && ! -L "$source_dir/tests" ]]; then
  mkdir -- "$stage/tests"
  for test_file in "$source_dir"/tests/*.py; do
    [[ -f "$test_file" && ! -L "$test_file" ]] || continue
    cp -p -- "$test_file" "$stage/tests/$(basename -- "$test_file")"
  done
fi
chmod +x "$stage"/*.command || fail 'Could not set launcher permissions.'
printf '\nCopied clean program files. Looking for Python 3.11, 3.12, or 3.13...\n'

# Prefer supported versioned executables; check every candidate rather than trusting its name.
found=0
for candidate in python3.11 python3.12 python3.13 python3 \
  /opt/homebrew/bin/python3.11 /opt/homebrew/bin/python3.12 /opt/homebrew/bin/python3.13 \
  /usr/local/bin/python3.11 /usr/local/bin/python3.12 /usr/local/bin/python3.13 \
  /Library/Frameworks/Python.framework/Versions/3.11/bin/python3.11 \
  /Library/Frameworks/Python.framework/Versions/3.12/bin/python3.12 \
  /Library/Frameworks/Python.framework/Versions/3.13/bin/python3.13; do
  if [[ "$candidate" == */* ]]; then [[ -x "$candidate" ]] || continue
  else command -v "$candidate" >/dev/null 2>&1 || continue
  fi
  if ! "$candidate" -c 'import sys; raise SystemExit(not (sys.version_info[:2] in ((3,11),(3,12),(3,13))))' >/dev/null 2>&1; then continue; fi
  found=1
  printf '\nTrying %s (%s)\n' "$candidate" "$("$candidate" --version 2>&1)"
  rm -rf -- "$stage/.venv"
  if ! "$candidate" -m venv "$stage/.venv"; then
    printf 'Could not create a virtual environment with this Python. Trying another...\n' >&2
    continue
  fi
  # Binary wheels only: never run a dependency source build. Pip alone may contact its configured index.
  if "$stage/.venv/bin/python" -m pip install --only-binary=:all: -r "$stage/requirements.txt"; then
    printf 'Dependencies installed.\n'
    break
  fi
  printf 'Dependencies unavailable for this Python (or pip/network failed). Trying another...\n' >&2
  rm -rf -- "$stage/.venv"
done
[[ -x "$stage/.venv/bin/python" ]] || {
  if (( found == 0 )); then fail 'No usable Python 3.11–3.13 found. Install one, then double-click this installer again.'; fi
  fail 'No supported Python could install the pinned binary packages. Check the pip messages above and retry when packages are reachable.'
}
# A venv created under the staging pathname contains that pathname in entry-point scripts;
# Python itself is relocatable for this project, but recreate after moving to ensure stable paths.
# Instead install at a unique final directory? Rebuild the environment there would expose partial state.
# The launcher invokes .venv/bin/python directly; this path is valid after moving on macOS.
[[ ! -e "$target" && ! -L "$target" ]] || fail "Destination appeared during setup: $target. Nothing was overwritten."
# macOS mv -n refuses to replace a newly appeared target; check stage existence to detect a race.
mv -n -- "$stage" "$target" || fail 'Could not move the prepared project into place.'
[[ ! -e "$stage" ]] || fail "Destination appeared during setup: $target. Nothing was overwritten."
stage=''
printf '\nInstalled at %s\n' "$target"
printf 'Open the launchers inside that folder, or use the optional Desktop shortcuts below.\n'
for file in 'Open BEMS Editor.command' 'Run Scheduler.command'; do
  shortcut="$desktop/$file"
  if [[ -e "$shortcut" || -L "$shortcut" ]]; then
    printf 'Skipped existing Desktop shortcut: %s\n' "$shortcut"
  else
    # -n is a second no-clobber safeguard against a concurrent new file.
    cp -n -- "$target/$file" "$shortcut" || printf 'Could not create optional shortcut: %s\n' "$shortcut" >&2
  fi
done
printf '\nNext: add the private form CSV/ZIP under inputs, check the selected config, and run --check before a full export.\n'
