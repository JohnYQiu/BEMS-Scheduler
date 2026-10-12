"""Local-only browser editor for a scheduler configuration (never runs the solver).

Usage: python editor.py [config.json] [--port 0] [--no-browser]
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import secrets
import tempfile
import webbrowser
from datetime import date
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit

from configuration import load_config
from forces import crew_options, load_volunteers, validate_forced_crews
from models import all_shift_keys, block_dates
from policy import (CUSTOM_PRIORITY_CATALOG, DEFAULT_OBJECTIVE_ORDER, HARD_REQUIREMENT_CATALOG,
                    OBJECTIVE_LABELS, validate_custom_priorities, validate_hard_requirements,
                    validate_objective_order)
from rule_catalog import FIELD_CATALOG, FIELD_BY_PATH, OPTIONAL_RULES, RULE_BY_PATH

DEFAULT_PRIORITIES = DEFAULT_OBJECTIVE_ORDER
ROOT = Path(__file__).resolve().parent
MAX_BODY = 256_000
MAX_DAYS = 366
RESERVED = {"block_start", "block_end", "shift_providers", "objective_order", "custom_priorities", "hard_requirements", "rules"}


class EditorError(ValueError):
    pass


def _digest(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def read_config(path: Path) -> tuple[dict, str]:
    raw = path.read_bytes()
    cfg = json.loads(raw)
    if not isinstance(cfg, dict):
        raise EditorError("Configuration must be a JSON object")
    return cfg, _digest(raw)


def scalar_paths(cfg: dict) -> dict[str, type]:
    """JSON-pointer-like paths for existing scalar settings only."""
    found = {}

    def visit(value, pieces):
        if isinstance(value, dict):
            for key, child in value.items():
                if isinstance(key, str):
                    visit(child, pieces + [key])
        elif type(value) in (str, int, float, bool):
            found["/" + "/".join(p.replace("~", "~0").replace("/", "~1") for p in pieces)] = type(value)

    for key, value in cfg.items():
        if key not in RESERVED:
            visit(value, [key])
    return {pointer: kind for pointer, kind in found.items() if pointer in FIELD_BY_PATH}


def _parts(pointer: str) -> list[str]:
    if not pointer.startswith("/"):
        raise EditorError("Invalid scalar setting path")
    return [p.replace("~1", "/").replace("~0", "~") for p in pointer[1:].split("/")]


def _dates_and_keys(cfg: dict) -> list[str]:
    try:
        start, end = date.fromisoformat(cfg["block_start"]), date.fromisoformat(cfg["block_end"])
        dates = block_dates(start, end)
    except (KeyError, TypeError, ValueError) as exc:
        raise EditorError(f"Invalid block dates: {exc}") from exc
    if len(dates) > MAX_DAYS:
        raise EditorError(f"Block cannot exceed {MAX_DAYS} days")
    return [f"{d.isoformat()}:{kind}" for d, kind in all_shift_keys(dates)]


def _priority_order(value, custom_priorities=()) -> list[str]:
    try:
        return list(validate_objective_order(value, custom_priorities))
    except ValueError as exc:
        raise EditorError(str(exc)) from exc


def build_candidate(original: dict, submitted: dict) -> dict:
    if not isinstance(submitted, dict):
        raise EditorError("Expected a JSON object")
    cfg = copy.deepcopy(original)
    for name in ("block_start", "block_end"):
        if not isinstance(submitted.get(name), str):
            raise EditorError(f"{name} must be a date")
        cfg[name] = submitted[name]
    expected = set(_dates_and_keys(cfg))
    providers = submitted.get("shift_providers")
    if not isinstance(providers, dict) or set(providers) != expected:
        missing = sorted(expected - set(providers)) if isinstance(providers, dict) else []
        extra = sorted(set(providers) - expected) if isinstance(providers, dict) else []
        raise EditorError(f"Provider grid must cover exactly the block shifts; missing: {missing[:4]}, extra: {extra[:4]}")
    if any(v not in ("ALS", "BLS") for v in providers.values()):
        raise EditorError("Each shift provider must be ALS or BLS")
    cfg["shift_providers"] = {key: providers[key] for key in sorted(expected)}

    scalar_types = scalar_paths(original)
    changes = submitted.get("scalars")
    if not isinstance(changes, dict) or set(changes) != set(scalar_types):
        raise EditorError("Scalar settings changed since the page loaded; reload before saving")
    for pointer, kind in scalar_types.items():
        value = changes[pointer]
        if type(value) is not kind and not (
            pointer == "/solver_time_limit_s" and type(value) in (int, float)
        ):
            raise EditorError(f"Wrong type for {pointer}: expected {kind.__name__}")
        if isinstance(value, str) and "\x00" in value:
            raise EditorError(f"Invalid null character in {pointer}")
        node = cfg
        parts = _parts(pointer)
        for part in parts[:-1]:
            node = node[part]
        node[parts[-1]] = value
    optional = submitted.get("optional_rules", {})
    if not isinstance(optional, dict) or not set(optional) <= set(RULE_BY_PATH):
        raise EditorError("Only implemented rules from the supported rule library may be added")
    if "rules" in cfg and not isinstance(cfg["rules"], dict):
        raise EditorError("rules must be a JSON object")
    rules = copy.deepcopy(cfg.get("rules", {}))
    for pointer, spec in RULE_BY_PATH.items():
        key = _parts(pointer)[-1]
        if pointer in optional:
            value = optional[pointer]
            if type(value) is not int or not spec["min"] <= value <= spec["max"]:
                raise EditorError(f"{spec['label']} must be an integer from {spec['min']} to {spec['max']}")
            rules[key] = value
        else:
            rules.pop(key, None)
    if rules or "rules" in cfg:
        cfg["rules"] = rules
    forced = submitted.get("forced_crews", cfg.get("forced_crews", []))
    if not isinstance(forced, list) or len(forced) > 200:
        raise EditorError("Forced crews must be a list of at most 200 shifts")
    cfg["forced_crews"] = forced
    try:
        validate_custom_priorities(original.get("custom_priorities", []))
    except ValueError as exc:
        raise EditorError("Existing custom priorities are incompatible with this editor; saving is blocked to prevent data loss: " + str(exc)) from exc
    try:
        custom_priorities = validate_custom_priorities(submitted.get("custom_priorities", []))
    except ValueError as exc:
        raise EditorError(str(exc)) from exc
    cfg["custom_priorities"] = list(custom_priorities)
    try:
        validate_hard_requirements(original.get("hard_requirements", []))
    except ValueError as exc:
        raise EditorError("Existing hard requirements are incompatible with this editor; saving is blocked to prevent data loss: " + str(exc)) from exc
    try:
        hard_requirements = validate_hard_requirements(submitted.get("hard_requirements", []))
    except ValueError as exc:
        raise EditorError(str(exc)) from exc
    cfg["hard_requirements"] = list(hard_requirements)
    cfg["objective_order"] = _priority_order(submitted.get("objective_order"), custom_priorities)
    return cfg


def validate_candidate(cfg: dict, config_path: Path) -> bytes:
    raw = (json.dumps(cfg, indent=2, ensure_ascii=False, allow_nan=False) + "\n").encode("utf-8")
    # Same directory ensures relative paths in configuration.load_config resolve exactly as in the real config.
    fd, name = tempfile.mkstemp(prefix=".editor-validate-", suffix=".json", dir=config_path.parent)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(raw)
        loaded = load_config(name)
        if loaded.forced_crews:
            validate_forced_crews(loaded, load_volunteers(loaded))
    finally:
        Path(name).unlink(missing_ok=True)
    return raw


def save_config(path: Path, expected_hash: str, submitted: dict) -> Path:
    if not isinstance(expected_hash, str):
        raise EditorError("Missing version; reload before saving")
    original, current_hash = read_config(path)
    if current_hash != expected_hash:
        raise EditorError("Config changed on disk since this page loaded. Reload to avoid overwriting those changes")
    cfg = build_candidate(original, submitted)
    data = validate_candidate(cfg, path)
    # No overwrite of existing backup. A backup is created before replacing the config.
    backup = path.with_name(path.name + ".bak-" + secrets.token_hex(6))
    fd, temp_name = tempfile.mkstemp(prefix=".editor-save-", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.chmod(temp_name, path.stat().st_mode & 0o777)
        backup_fd = os.open(backup, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with path.open("rb") as src, os.fdopen(backup_fd, "wb") as dst:
            os.fchmod(dst.fileno(), path.stat().st_mode & 0o777)
            while chunk := src.read(65536):
                dst.write(chunk)
            dst.flush()
            os.fsync(dst.fileno())
        # Detect an external write during validation/backup, rather than silently replacing it.
        if _digest(path.read_bytes()) != current_hash:
            raise EditorError("Config changed during save. No replacement made; reload before retrying")
        os.replace(temp_name, path)
        if hasattr(os, "O_DIRECTORY"):
            directory_fd = os.open(path.parent, os.O_RDONLY | os.O_DIRECTORY)
            try:
                os.fsync(directory_fd)
            finally:
                os.close(directory_fd)
        return backup
    finally:
        Path(temp_name).unlink(missing_ok=True)


def make_handler(config_path: Path, token: str):
    html = (ROOT / "editor.html").read_bytes()

    class Handler(BaseHTTPRequestHandler):
        def _local_request(self) -> bool:
            host = self.headers.get("Host", "")
            parsed_host = host.partition(":")
            if parsed_host[0] != "127.0.0.1" or (parsed_host[2] and parsed_host[2] != str(self.server.server_port)):
                return False
            origin = self.headers.get("Origin")
            if origin and origin != f"http://127.0.0.1:{self.server.server_port}":
                return False
            return True

        def _reply(self, code: int, data: bytes, content_type="application/json; charset=utf-8"):
            self.send_response(code)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("Content-Security-Policy", "default-src 'self'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; connect-src 'self'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'")
            self.send_header("Referrer-Policy", "no-referrer")
            self.end_headers()
            self.wfile.write(data)

        def _json(self, code: int, payload: dict):
            self._reply(code, json.dumps(payload, ensure_ascii=False).encode("utf-8"))

        def do_GET(self):
            if not self._local_request():
                return self._json(403, {"error": "Localhost only"})
            route = urlsplit(self.path).path
            if route == "/":
                return self._reply(200, html, "text/html; charset=utf-8")
            if route == "/api/crew-options":
                try:
                    current = load_config(config_path)
                    self._json(200, {"available": True, "shifts": crew_options(current),
                                     "source": str(current.form_csv)})
                except (OSError, ValueError, TypeError, KeyError) as exc:
                    self._json(200, {"available": False, "reason": str(exc), "shifts": []})
                return
            if route == "/api/config":
                try:
                    cfg, version = read_config(config_path)
                    _dates_and_keys(cfg)
                    self._json(200, {"config": cfg, "version": version, "token": token,
                                     "path": str(config_path), "default_priorities": DEFAULT_PRIORITIES,
                                     "priority_labels": OBJECTIVE_LABELS,
                                     "custom_priority_catalog": CUSTOM_PRIORITY_CATALOG,
                                     "hard_requirement_catalog": HARD_REQUIREMENT_CATALOG,
                                     "field_catalog": FIELD_CATALOG, "rule_catalog": OPTIONAL_RULES,
                                     "scalar_paths": list(scalar_paths(cfg))})
                except (OSError, ValueError, json.JSONDecodeError) as exc:
                    self._json(400, {"error": str(exc)})
                return
            self._json(404, {"error": "Not found"})

        def do_POST(self):
            if not self._local_request() or not secrets.compare_digest(self.headers.get("X-Editor-Token", ""), token):
                return self._json(403, {"error": "Local request and editor token required"})
            if urlsplit(self.path).path != "/api/save":
                return self._json(404, {"error": "Not found"})
            if self.headers.get("Content-Type", "").split(";", 1)[0] != "application/json":
                return self._json(415, {"error": "Expected JSON"})
            try:
                length = int(self.headers.get("Content-Length", "0"))
                if length < 1 or length > MAX_BODY:
                    raise EditorError("Request size out of range")
                payload = json.loads(self.rfile.read(length))
                if not isinstance(payload, dict):
                    raise EditorError("Expected a JSON object")
                backup = save_config(config_path, payload.get("version"), payload.get("config"))
                _, version = read_config(config_path)
                self._json(200, {"ok": True, "backup": str(backup), "version": version})
            except (OSError, ValueError, TypeError, KeyError) as exc:
                self._json(400, {"error": str(exc)})

    return Handler


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("config", nargs="?", default="config.json", help="chosen config JSON file (default: config.json beside editor.py)")
    parser.add_argument("--port", type=int, default=0, help="local port; 0 chooses a free random port")
    parser.add_argument("--no-browser", action="store_true", help="print URL without opening a browser")
    args = parser.parse_args(argv)
    if not 0 <= args.port <= 65535:
        parser.error("--port must be 0 through 65535")
    config_path = (Path(args.config) if args.config != "config.json" else ROOT / "config.json").resolve()
    try:
        cfg, _ = read_config(config_path)
        _dates_and_keys(cfg)
        load_config(config_path)
    except (OSError, ValueError, TypeError, KeyError) as exc:
        parser.error(f"Cannot load {config_path}: {exc}")
    token = secrets.token_urlsafe(32)
    with ThreadingHTTPServer(("127.0.0.1", args.port), make_handler(config_path, token)) as server:
        url = f"http://127.0.0.1:{server.server_port}/"
        print(f"Editing {config_path}\nOpen {url}\nCtrl-C to stop. Saving makes a timestamp-independent unique backup; solver is never run.", flush=True)
        if not args.no_browser:
            webbrowser.open(url)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass


if __name__ == "__main__":
    main()
