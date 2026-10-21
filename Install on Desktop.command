#!/bin/bash
# Double-click from a downloaded/extracted BEMS-Scheduler folder. Updates preserve local data and keep a full rollback copy.
set -euo pipefail

pause() { printf '\nPress Return to close this window... '; read -r _answer || true; }
fail() { printf '\nERROR: %s\n' "$*" >&2; exit 1; }
stage=''
target=''
backup=''
updating=0
finish() {
  status=$?
  trap - EXIT
  if (( status != 0 )) && (( updating == 1 )) && [[ -n "$backup" && -d "$backup" && -n "$target" && ! -e "$target" ]]; then
    mv -n -- "$backup" "$target" || printf 'Rollback copy remains at %s\n' "$backup" >&2
  fi
  if [[ -n "$stage" && -d "$stage" ]]; then rm -rf -- "$stage"; fi
  if (( status != 0 )); then printf '\nInstallation did not complete. Any previous installation remains available or was restored.\n' >&2; fi
  pause
  exit "$status"
}
trap finish EXIT

source_dir=$(cd -- "$(dirname -- "$0")" && pwd -P) || fail 'Cannot locate the extracted project folder.'
desktop="$HOME/Desktop"
target="$desktop/BEMS Scheduler"
[[ -d "$desktop" ]] || fail "Desktop directory does not exist: $desktop"
[[ -f "$source_dir/main.py" && -f "$source_dir/editor.py" && -f "$source_dir/editor.html" && -f "$source_dir/requirements.txt" && -f "$source_dir/config.json" ]] || fail 'Required project files are missing. Extract the complete repository first.'
if [[ -e "$target" || -L "$target" ]]; then
  [[ -d "$target" && ! -L "$target" && -f "$target/main.py" && -f "$target/config.json" ]] || fail "The destination exists but does not look like a BEMS Scheduler installation: $target"
  updating=1
  stamp=$(date '+%Y%m%d-%H%M%S')
  backup="$desktop/BEMS Scheduler backup $stamp"
  [[ ! -e "$backup" && ! -L "$backup" ]] || fail "A rollback folder already exists: $backup"
  printf 'Updating the existing installation at:\n  %s\n\n' "$target"
  printf 'Your config files, local inputs, outputs, backups, and other non-program files will be copied into the update.\n'
  printf 'A complete rollback copy will be kept at:\n  %s\n\n' "$backup"
else
  printf 'Installing a separate clean copy at:\n  %s\n\n' "$target"
  printf 'Only program files, supplied configurations, documentation, and launchers are copied.\n'
  printf 'Local inputs, outputs, backups, .git, .venv, and *.local.json are not copied.\n'
fi
printf 'If needed, a private Python runtime will be downloaded and SHA-256 verified. Python dependencies also require internet access. No sudo or system-wide changes are used.\n\n'
if (( updating == 1 )); then printf 'Continue with update? [y/N] '; else printf 'Continue? [y/N] '; fi
read -r answer || fail 'No answer received.'
[[ "$answer" == [yY] || "$answer" == [yY][eE][sS] ]] || { printf 'Canceled.\n'; exit 0; }

stage=$(mktemp -d "$desktop/.BEMS-Scheduler-install.XXXXXXXX") || fail 'Could not create a temporary installation directory on Desktop.'
# Explicit allowlist: never recursively copy arbitrary contents from an extracted folder.
for file in main.py editor.py editor.html rule_catalog.py member_attributes.py forces.py configuration.py models.py output.py parse_form.py policy.py solver.py validation.py requirements.txt config.json config_f26b3.json README.md OFFICER_GUIDE.md .gitignore 'Install on Desktop.command' 'Open BEMS Editor.command' 'Run Scheduler.command'; do
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
if (( updating == 1 )); then
  # Carry user data forward, while replacing application files from the new source.
  for entry in "$target"/* "$target"/.[!.]* "$target"/..?*; do
    [[ -e "$entry" || -L "$entry" ]] || continue
    name=$(basename -- "$entry")
    case "$name" in
      .venv|.git|tests|__pycache__|.DS_Store|main.py|editor.py|editor.html|rule_catalog.py|member_attributes.py|forces.py|configuration.py|models.py|output.py|parse_form.py|policy.py|solver.py|validation.py|requirements.txt|README.md|OFFICER_GUIDE.md|.gitignore|"Install on Desktop.command"|"Open BEMS Editor.command"|"Run Scheduler.command")
        continue ;;
    esac
    cp -pR -- "$entry" "$stage/$name" || fail "Could not preserve existing local data: $name"
  done
fi
chmod +x "$stage"/*.command || fail 'Could not set launcher permissions.'
printf '\nCopied clean program files. Checking for Python 3.11, 3.12, or 3.13...\n'

setup_environment() {
  local python_exe="$1"
  rm -rf -- "$stage/.venv"
  if ! "$python_exe" -m venv "$stage/.venv"; then
    printf 'Could not create a virtual environment with %s.\n' "$python_exe" >&2
    rm -rf -- "$stage/.venv"
    return 1
  fi
  # Binary wheels only: never run a dependency source build. Pip alone contacts its configured index.
  if "$stage/.venv/bin/python" -m pip install --only-binary=:all: -r "$stage/requirements.txt"; then
    printf 'Dependencies installed.\n'
    return 0
  fi
  printf 'Pinned binary packages could not be installed with %s.\n' "$python_exe" >&2
  rm -rf -- "$stage/.venv"
  return 1
}

# If no suitable Python is already available, install a pinned, checksum-verified
# standalone runtime in this user's Application Support folder. This does not
# modify macOS, PATH, shell profiles, or require administrator privileges.
ensure_private_python() {
  local machine arch filename url checksum runtime_parent runtime_dir temp_dir
  machine=$(uname -m)
  case "$machine" in
    arm64|aarch64) arch='aarch64' ;;
    x86_64) arch='x86_64' ;;
    *) printf 'Automatic Python setup supports Apple Silicon (arm64) and Intel (x86_64), not %s.\n' "$machine" >&2; return 1 ;;
  esac
  filename="cpython-3.13.16+20261009-${arch}-apple-darwin-install_only_stripped.tar.gz"
  case "$arch" in
    aarch64) checksum='11785155d0f70e634178219c20925d8f20be00a3da4684fc4f5019dd15602ef7' ;;
    x86_64) checksum='bd085c51e960f00ee33b630e2b13ac90fe90cdc3da25a0f0bc0719b1294920bf' ;;
  esac
  url="https://github.com/astral-sh/python-build-standalone/releases/download/20261009/${filename/+/%2B}"
  runtime_parent="$HOME/Library/Application Support/BEMS Scheduler"
  runtime_dir="$runtime_parent/cpython-3.13.16-20261009-${arch}"
  mkdir -p -- "$runtime_parent" || { printf 'Could not create the per-user Python runtime folder: %s\n' "$runtime_parent" >&2; return 1; }

  if [[ -e "$runtime_dir" || -L "$runtime_dir" ]]; then
    if [[ ! -f "$runtime_dir/bin/python3.13" || ! -f "$runtime_dir/.bems-python-sha256" ]] || \
       [[ "$(cat "$runtime_dir/.bems-python-sha256" 2>/dev/null)" != "$checksum" ]] || \
       ! "$runtime_dir/bin/python3.13" -c 'import sys; raise SystemExit(sys.version_info[:3] != (3,13,16))' >/dev/null 2>&1; then
      printf 'A Python runtime path already exists but is not the expected verified runtime: %s\n' "$runtime_dir" >&2
      printf 'Nothing was replaced. Inspect or move that folder yourself, then rerun.\n' >&2
      return 1
    fi
    PRIVATE_PYTHON="$runtime_dir/bin/python3.13"
    printf 'Reusing verified private Python 3.13.16 at %s\n' "$PRIVATE_PYTHON"
    return 0
  fi

  command -v curl >/dev/null 2>&1 || { printf 'curl is required to download the private Python runtime.\n' >&2; return 1; }
  command -v shasum >/dev/null 2>&1 || { printf 'shasum is required to verify the private Python runtime.\n' >&2; return 1; }
  command -v tar >/dev/null 2>&1 || { printf 'tar is required to unpack the private Python runtime.\n' >&2; return 1; }
  temp_dir=$(mktemp -d "$runtime_parent/.python-runtime.XXXXXXXX") || { printf 'Could not create a temporary runtime folder.\n' >&2; return 1; }
  printf '\nNo existing Python completed setup. Downloading private Python 3.13.16 for %s...\n' "$machine"
  if ! curl --proto '=https' --tlsv1.2 -fL --retry 2 --connect-timeout 20 "$url" -o "$temp_dir/python.tar.gz"; then
    rm -rf -- "$temp_dir"
    printf 'Python download failed. Check your internet connection and try again.\n' >&2
    return 1
  fi
  if ! (cd "$temp_dir" && printf '%s  %s\n' "$checksum" python.tar.gz | shasum -a 256 -c -); then
    rm -rf -- "$temp_dir"
    printf 'Python archive checksum did not match. The archive was not installed.\n' >&2
    return 1
  fi
  if ! tar -xzf "$temp_dir/python.tar.gz" -C "$temp_dir" || \
     [[ ! -x "$temp_dir/python/bin/python3.13" ]] || \
     ! "$temp_dir/python/bin/python3.13" -c 'import sys; raise SystemExit(sys.version_info[:3] != (3,13,16))' >/dev/null 2>&1; then
    rm -rf -- "$temp_dir"
    printf 'The verified Python archive did not contain a usable Python 3.13.16 runtime.\n' >&2
    return 1
  fi
  printf '%s\n' "$checksum" > "$temp_dir/python/.bems-python-sha256"
  if ! mv -n -- "$temp_dir/python" "$runtime_dir" || [[ -e "$temp_dir/python" ]]; then
    rm -rf -- "$temp_dir"
    printf 'Could not install the private Python runtime without replacing an existing path.\n' >&2
    return 1
  fi
  rm -rf -- "$temp_dir"
  PRIVATE_PYTHON="$runtime_dir/bin/python3.13"
  printf 'Installed verified private Python at %s\n' "$PRIVATE_PYTHON"
  return 0
}

installed=0
for candidate in python3.11 python3.12 python3.13 python3 \
  "$HOME/.aside/runtime/bin/python3.11" "$HOME/.aside/runtime/bin/python3.12" "$HOME/.aside/runtime/bin/python3.13" \
  /opt/homebrew/bin/python3.11 /opt/homebrew/bin/python3.12 /opt/homebrew/bin/python3.13 \
  /usr/local/bin/python3.11 /usr/local/bin/python3.12 /usr/local/bin/python3.13 \
  /Library/Frameworks/Python.framework/Versions/3.11/bin/python3.11 \
  /Library/Frameworks/Python.framework/Versions/3.12/bin/python3.12 \
  /Library/Frameworks/Python.framework/Versions/3.13/bin/python3.13; do
  if [[ "$candidate" == */* ]]; then [[ -x "$candidate" ]] || continue
  else command -v "$candidate" >/dev/null 2>&1 || continue
  fi
  if ! "$candidate" -c 'import sys; raise SystemExit(not (sys.version_info[:2] in ((3,11),(3,12),(3,13))))' >/dev/null 2>&1; then continue; fi
  printf '\nTrying existing Python: %s (%s)\n' "$candidate" "$("$candidate" --version 2>&1)"
  if setup_environment "$candidate"; then installed=1; break; fi
done

if (( installed == 0 )); then
  ensure_private_python || fail 'Could not install a verified private Python runtime. Your existing Desktop files were not changed.'
  setup_environment "$PRIVATE_PYTHON" || fail 'Python is ready, but the pinned binary dependencies could not be installed. Check the pip messages above and retry when the package index is reachable.'
fi

# venv console-script shebangs retain the staging path, so launchers and documented
# commands invoke .venv/bin/python directly. The base Python stays at a stable path,
# including the private runtime under Application Support when bootstrap was needed.
if (( updating == 1 )); then
  mv -n -- "$target" "$backup" || fail 'Could not move the existing installation to its rollback folder.'
  [[ ! -e "$target" && -d "$backup" ]] || fail 'The existing installation changed during setup; no files were replaced.'
  mv -n -- "$stage" "$target" || fail 'Could not move the updated project into place.'
  [[ ! -e "$stage" ]] || fail 'Destination appeared during setup; no files were overwritten.'
else
  [[ ! -e "$target" && ! -L "$target" ]] || fail "Destination appeared during setup: $target. Nothing was overwritten."
  # macOS mv -n refuses to replace a newly appeared target; check stage existence to detect a race.
  mv -n -- "$stage" "$target" || fail 'Could not move the prepared project into place.'
  [[ ! -e "$stage" ]] || fail "Destination appeared during setup: $target. Nothing was overwritten."
fi
stage=''
if (( updating == 1 )); then
  printf '\nUpdated at %s\n' "$target"
  printf 'Rollback copy: %s\n' "$backup"
else
  printf '\nInstalled at %s\n' "$target"
fi
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
