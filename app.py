#!/usr/bin/env python3
"""
Department Inventory Manager
----------------------------
A single-user web app for managing categorized inventory (units, electronics, etc.).

- Dynamic category / subcategory tree (any depth)
- Items with dynamic custom fields (serial number, origin, notes, ...) only where needed
- Per-category required-field templates (auto-applied to new items in that category)
- Auto-save on every change (no save button)
- Full change history with timestamps
- Data stored as Excel-readable CSV files on a folder you choose (e.g. a shared drive)

Run:
    python app.py
Then open http://127.0.0.1:8765 in your browser (it opens automatically).

The first time it runs it asks (in the browser) for a data folder, or you can set
the INVENTORY_DIR environment variable to point at your shared drive.
"""

import csv
import os
import re
import sys
import json
import uuid
import shutil
import threading
import subprocess
import webbrowser
from datetime import datetime
from pathlib import Path

def fatal(message):
    """
    Report a startup failure and quit. Falls back to a dialog box when there is
    no console to print to (Windows pythonw.exe sets sys.stderr to None).
    """
    if sys.stderr is not None:
        sys.stderr.write("\n[Inventory Manager] " + message + "\n\n")
    elif os.name == "nt":
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                None, message, "Inventory Manager", 0x10)
        except Exception:
            pass
    sys.exit(1)


try:
    from flask import Flask, request, jsonify, Response
except ModuleNotFoundError:
    fatal(
        "Flask is not installed.\n\nInstall it with:\n\n"
        "    python -m pip install flask\n\n"
        "(use python3 instead of python on macOS/Linux), "
        "then run this program again."
    )

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------

APP = Flask(__name__)
LOCK = threading.RLock()  # single-user, but guards against overlapping requests

# Bump this when you publish a new version. The updater compares it with the
# APP_VERSION in the copy of app.py on GitHub.
APP_VERSION = "1.8.0"

# Where updates come from. Normally auto-detected from this checkout's git
# remote; set INVENTORY_REPO ("owner/name") to override.
DEFAULT_REPO = "AmjadTrablsiD1/Inventory"
UPDATE_BRANCH = os.environ.get("INVENTORY_BRANCH", "main")
APP_DIR = Path(__file__).resolve().parent

# Where the CSVs live. Change this, set INVENTORY_DIR, or pick a folder in the UI.
DEFAULT_DIR = os.environ.get(
    "INVENTORY_DIR",
    str(Path.home() / "InventoryData"),
)

# Per-user setting, private to one machine.
CONFIG_FILE = Path.home() / ".inventory_manager_config.json"

# Setting stored next to app.py. When the app sits on a shared drive this file
# is shared too, so every computer running that copy finds the same data.
PORTABLE_CONFIG_NAME = "inventory_config.json"

INVENTORY_CSV = "inventory.csv"
CATEGORIES_CSV = "categories.csv"
HISTORY_CSV = "history.csv"
DELETED_CSV = "deleted_items.csv"   # every deleted item is archived here
ORDER_CSV = "order_list.csv"        # the working "things to order" list
COLUMNS_JSON = "column_settings.json"   # per-folder column order/labels/hidden

# Shown on order lists and in exports. Change to "$", "CHF", ... as needed.
CURRENCY = os.environ.get("INVENTORY_CURRENCY", "€")

# Human-readable mirror: one CSV per top-level category, subcategories as
# labelled sections inside it. inventory.csv stays the master copy.
HIER_DIR = "by_category"
HIER_MARKER = "# Inventory Manager - readable category view (auto-generated)"
UNCATEGORIZED = "_Uncategorized"

# Fixed base columns that always exist on every item row.
#   color      - "#rrggbb" chosen in the app or carried over from an import
#   status     - "" (in stock) or "used"
#   used_note  - free text: where/what it is used for
#   used_date  - when it was marked used
BASE_COLUMNS = ["id", "category_path", "name", "quantity",
                "color", "status", "used_note", "used_date",
                "date_added", "last_modified"]

# Columns that exist for the program's benefit rather than the user's.
# A "clean" export leaves these out.
INTERNAL_COLUMNS = ["id"]


# ----------------------------------------------------------------------------
# Data directory resolution
# ----------------------------------------------------------------------------

def load_config():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_config(cfg):
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    except Exception:
        pass


def env_data_dir():
    """The data folder pinned by INVENTORY_DIR, if that variable is set."""
    return (os.environ.get("INVENTORY_DIR") or "").strip()


def portable_config_file():
    return APP_DIR / PORTABLE_CONFIG_NAME


def load_portable_config():
    f = portable_config_file()
    if f.exists():
        try:
            return json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_portable_config(cfg):
    try:
        portable_config_file().write_text(
            json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")
        return True
    except Exception:
        return False


def portable_data_dir():
    """
    The data folder from the config beside app.py, if there is one.

    A path *inside* the app's own folder is stored relative, so a shared drive
    mounted as Z:\\ on one machine and /Volumes/share on another still resolves.
    """
    cfg = load_portable_config()
    rel = cfg.get("data_dir_relative")
    if rel:
        try:
            return (APP_DIR / rel).resolve()
        except Exception:
            return None
    d = cfg.get("data_dir")
    return Path(d).expanduser() if d else None


def get_data_dir():
    """
    Where the CSVs live, in order of precedence:
      1. INVENTORY_DIR                        - explicit, wins over everything
      2. this machine's setting, if it was    - "just this computer" choice
         saved with override_shared
      3. the config next to app.py            - shared drive: same for everyone
      4. this machine's setting
      5. ~/InventoryData
    """
    env = env_data_dir()
    if env:
        return Path(env).expanduser()

    cfg = load_config()
    if cfg.get("override_shared") and cfg.get("data_dir"):
        return Path(cfg["data_dir"])

    shared = portable_data_dir()
    if shared:
        return shared

    return Path(cfg.get("data_dir") or DEFAULT_DIR)


def data_dir_source():
    """Human-readable note about where the current setting comes from."""
    if env_data_dir():
        return {"scope": "env", "where": "INVENTORY_DIR environment variable"}
    cfg = load_config()
    if cfg.get("override_shared") and cfg.get("data_dir"):
        return {"scope": "user", "where": str(CONFIG_FILE)}
    if portable_data_dir():
        return {"scope": "shared", "where": str(portable_config_file())}
    if cfg.get("data_dir"):
        return {"scope": "user", "where": str(CONFIG_FILE)}
    return {"scope": "default", "where": "built-in default"}


def set_data_dir(path, scope="shared"):
    """
    Remember the data folder. scope="shared" writes next to app.py so every
    computer running this copy agrees; scope="user" keeps it to this machine.
    Returns the scope actually used (shared falls back to user if app.py sits
    somewhere unwritable).
    """
    path = Path(path)
    if scope == "shared":
        cfg = load_portable_config()
        rel = None
        try:
            # Relative wherever it is sensible - including a sibling folder on
            # the same share - so the setting survives the drive being mounted
            # as Z:\ on one machine and /Volumes/share on another.
            r = os.path.relpath(str(path.resolve()), str(APP_DIR.resolve()))
            if not os.path.isabs(r) and r.count("..") <= 3:
                rel = r
        except Exception:
            rel = None
        if rel:
            cfg.pop("data_dir", None)
            cfg["data_dir_relative"] = rel
        else:
            cfg.pop("data_dir_relative", None)
            cfg["data_dir"] = str(path)
        if save_portable_config(cfg):
            user = load_config()
            user.pop("override_shared", None)   # stop overriding the shared one
            user["data_dir"] = str(path)        # harmless fallback copy
            save_config(user)
            return "shared"

    user = load_config()
    user["data_dir"] = str(path)
    user["override_shared"] = True
    save_config(user)
    return "user"


def paths():
    d = get_data_dir()
    return (
        d,
        d / INVENTORY_CSV,
        d / CATEGORIES_CSV,
        d / HISTORY_CSV,
    )


def deleted_path():
    return get_data_dir() / DELETED_CSV


def order_path():
    return get_data_dir() / ORDER_CSV


# ----------------------------------------------------------------------------
# Column settings: which columns show, in what order, under what heading.
# Stored beside the data so they travel with it.
# ----------------------------------------------------------------------------

def read_col_settings():
    f = get_data_dir() / COLUMNS_JSON
    out = {"labels": {}, "hidden": [], "order": []}
    if f.exists():
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
            for k in out:
                if k in data and isinstance(data[k], type(out[k])):
                    out[k] = data[k]
        except Exception:
            pass
    return out


def write_col_settings(s):
    try:
        (get_data_dir() / COLUMNS_JSON).write_text(
            json.dumps(s, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def rename_custom_field(old, new):
    """Rename a custom column everywhere: the CSV header, every row, and any
    category that lists it as a required field."""
    rows, custom = read_inventory()
    if old in BASE_COLUMNS:
        return False, "That is a built-in column - you can rename its heading, but not the field itself."
    if old not in custom:
        return False, f"There is no column called '{old}'."
    new = sanitize_field_name(new)
    if not new:
        return False, "A name is required."
    if new == old:
        return True, ""
    if new in BASE_COLUMNS or new in custom:
        return False, f"A column called '{new}' already exists."

    custom[custom.index(old)] = new
    for r in rows:
        r[new] = r.pop(old, "")
    write_inventory(rows, custom)

    cat_map = read_categories()
    if any(old in v for v in cat_map.values()):
        for k, v in cat_map.items():
            cat_map[k] = [new if x == old else x for x in v]
        write_categories(cat_map)

    s = read_col_settings()
    s["hidden"] = [new if c == old else c for c in s["hidden"]]
    s["order"] = [new if c == old else c for c in s["order"]]
    if old in s["labels"]:
        s["labels"][new] = s["labels"].pop(old)
    write_col_settings(s)

    log_history("rename_column", old, f"-> {new}")
    return True, ""


def delete_custom_field(name):
    """Remove a custom column and its data from every item."""
    rows, custom = read_inventory()
    if name in BASE_COLUMNS:
        return False, "Built-in columns cannot be deleted - hide it instead."
    if name not in custom:
        return False, f"There is no column called '{name}'."

    filled = sum(1 for r in rows if str(r.get(name, "")).strip())
    custom.remove(name)
    for r in rows:
        r.pop(name, None)
    write_inventory(rows, custom)

    cat_map = read_categories()
    if any(name in v for v in cat_map.values()):
        for k, v in cat_map.items():
            cat_map[k] = [x for x in v if x != name]
        write_categories(cat_map)

    s = read_col_settings()
    s["hidden"] = [c for c in s["hidden"] if c != name]
    s["order"] = [c for c in s["order"] if c != name]
    s["labels"].pop(name, None)
    write_col_settings(s)

    log_history("delete_column", name, f"removed from {filled} item(s)")
    return True, ""


# ----------------------------------------------------------------------------
# Native "choose folder" dialog (the server runs on the same machine as the UI)
# ----------------------------------------------------------------------------

# Keeps a console window from flashing up when we shell out on Windows.
NO_WINDOW = 0x08000000 if os.name == "nt" else 0


def run_hidden(cmd, timeout=300):
    """Run a helper command without popping up a console window."""
    return subprocess.run(cmd, capture_output=True, text=True,
                          timeout=timeout, creationflags=NO_WINDOW)


def native_picker_available():
    """True if this machine has an OS folder-picker we know how to invoke."""
    if sys.platform == "darwin":
        return shutil.which("osascript") is not None
    if os.name == "nt":
        return shutil.which("powershell") is not None
    return shutil.which("zenity") is not None or shutil.which("kdialog") is not None


def pick_directory_native():
    """
    Open the OS folder chooser and return the selected path, or None if the
    user cancelled or no native dialog is available. Blocks until the user picks.
    """
    try:
        if sys.platform == "darwin":
            script = ('POSIX path of (choose folder with prompt '
                      '"Choose a folder for your inventory data")')
            r = run_hidden(["osascript", "-e", script])
            return r.stdout.strip() if r.returncode == 0 and r.stdout.strip() else None

        if os.name == "nt":
            ps = (
                "Add-Type -AssemblyName System.Windows.Forms; "
                "$f = New-Object System.Windows.Forms.FolderBrowserDialog; "
                "if ($f.ShowDialog() -eq 'OK') { Write-Output $f.SelectedPath }"
            )
            r = run_hidden(["powershell", "-NoProfile", "-WindowStyle", "Hidden",
                            "-Command", ps])
            return r.stdout.strip() if r.returncode == 0 and r.stdout.strip() else None

        # Linux
        for cmd in (["zenity", "--file-selection", "--directory",
                     "--title=Choose a folder for your inventory data"],
                    ["kdialog", "--getexistingdirectory", str(Path.home())]):
            if shutil.which(cmd[0]):
                r = run_hidden(cmd)
                return r.stdout.strip() if r.returncode == 0 and r.stdout.strip() else None
        return None
    except Exception:
        return None


def pick_save_file_native(default_name="inventory_export.csv"):
    """Open the OS 'save as' dialog. Returns a path or None."""
    try:
        if sys.platform == "darwin":
            script = ('POSIX path of (choose file name with prompt '
                      '"Export inventory as CSV" default name "%s")' % default_name)
            r = run_hidden(["osascript", "-e", script])
            return r.stdout.strip() if r.returncode == 0 and r.stdout.strip() else None

        if os.name == "nt":
            ps = (
                "Add-Type -AssemblyName System.Windows.Forms; "
                "$f = New-Object System.Windows.Forms.SaveFileDialog; "
                f"$f.FileName = '{default_name}'; "
                "$f.Filter = 'CSV file|*.csv|All files|*.*'; "
                "if ($f.ShowDialog() -eq 'OK') { Write-Output $f.FileName }"
            )
            r = run_hidden(["powershell", "-NoProfile", "-WindowStyle", "Hidden",
                            "-Command", ps])
            return r.stdout.strip() if r.returncode == 0 and r.stdout.strip() else None

        for cmd in (["zenity", "--file-selection", "--save", "--confirm-overwrite",
                     f"--filename={default_name}"],
                    ["kdialog", "--getsavefilename", default_name]):
            if shutil.which(cmd[0]):
                r = run_hidden(cmd)
                return r.stdout.strip() if r.returncode == 0 and r.stdout.strip() else None
        return None
    except Exception:
        return None


def pick_file_native():
    """Open the OS file chooser for a spreadsheet. Returns a path or None."""
    try:
        if sys.platform == "darwin":
            script = ('POSIX path of (choose file with prompt '
                      '"Choose a spreadsheet to import" of type '
                      '{"csv","xlsx","xlsm","tsv","txt"})')
            r = run_hidden(["osascript", "-e", script])
            return r.stdout.strip() if r.returncode == 0 and r.stdout.strip() else None

        if os.name == "nt":
            ps = (
                "Add-Type -AssemblyName System.Windows.Forms; "
                "$f = New-Object System.Windows.Forms.OpenFileDialog; "
                "$f.Filter = 'Spreadsheets|*.csv;*.xlsx;*.xlsm;*.tsv;*.txt|All files|*.*'; "
                "if ($f.ShowDialog() -eq 'OK') { Write-Output $f.FileName }"
            )
            r = run_hidden(["powershell", "-NoProfile", "-WindowStyle", "Hidden",
                            "-Command", ps])
            return r.stdout.strip() if r.returncode == 0 and r.stdout.strip() else None

        for cmd in (["zenity", "--file-selection",
                     "--title=Choose a spreadsheet to import"],
                    ["kdialog", "--getopenfilename", str(Path.home())]):
            if shutil.which(cmd[0]):
                r = run_hidden(cmd)
                return r.stdout.strip() if r.returncode == 0 and r.stdout.strip() else None
        return None
    except Exception:
        return None


# ----------------------------------------------------------------------------
# Low-level CSV helpers
# ----------------------------------------------------------------------------

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def ensure_files():
    """Create the data folder and empty CSVs if they don't exist yet."""
    d, inv, cats, hist = paths()
    d.mkdir(parents=True, exist_ok=True)
    if not inv.exists():
        write_inventory([], [], sync_view=False)
    if not cats.exists():
        with cats.open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["category_path", "required_fields"])
    if not hist.exists():
        with hist.open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["timestamp", "action", "target", "detail"])


def read_inventory():
    """Return (rows, custom_field_names). Each row is a dict."""
    _, inv, _, _ = paths()
    if not inv.exists():
        return [], []
    with inv.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        header = reader.fieldnames or list(BASE_COLUMNS)
        rows = [dict(r) for r in reader]
    custom = [c for c in header if c not in BASE_COLUMNS]
    # normalise: make sure every row has every column key
    for r in rows:
        for c in header:
            r.setdefault(c, "")
    return rows, custom


def write_inventory(rows, custom_fields, sync_view=True):
    """Write all item rows. custom_fields is the union of extra columns."""
    _, inv, _, _ = paths()
    header = list(BASE_COLUMNS) + list(custom_fields)
    with inv.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=header, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            clean = {k: r.get(k, "") for k in header}
            w.writerow(clean)
    if sync_view:
        write_hierarchy()


def read_categories():
    """Return dict: category_path -> [required_field, ...]."""
    _, _, cats, _ = paths()
    result = {}
    if not cats.exists():
        return result
    with cats.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for r in reader:
            path = (r.get("category_path") or "").strip()
            if not path:
                continue
            req = (r.get("required_fields") or "").strip()
            result[path] = [x.strip() for x in req.split("|") if x.strip()]
    return result


def write_categories(cat_map, sync_view=True):
    _, _, cats, _ = paths()
    with cats.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["category_path", "required_fields"])
        for path in sorted(cat_map.keys()):
            w.writerow([path, "|".join(cat_map[path])])
    if sync_view:
        write_hierarchy()


def log_history(action, target, detail=""):
    _, _, _, hist = paths()
    with hist.open("a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow([now_str(), action, target, detail])


# ----------------------------------------------------------------------------
# Readable per-category mirror (one CSV per top-level category, with sections)
# ----------------------------------------------------------------------------

def safe_filename(name):
    """Turn a category name into something safe for a file name."""
    cleaned = re.sub(r'[<>:"/\\|?*\r\n]', "_", str(name)).strip().rstrip(".")
    return cleaned or "_"


def write_hierarchy():
    """
    Rebuild <data_dir>/by_category/: one CSV per top-level category, each
    containing '=== Full/Category/Path ===' sections with that category's own
    items. Purely a readable view -- inventory.csv remains the master.
    """
    d, _, _, _ = paths()
    out_dir = d / HIER_DIR
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        return

    rows, _ = read_inventory()
    cat_map = read_categories()

    # every category path that exists (stored, implied by items, or an ancestor)
    all_paths = set(cat_map.keys())
    for r in rows:
        p = (r.get("category_path") or "").strip()
        if p:
            all_paths.add(p)
    for p in list(all_paths):
        parts = p.split("/")
        for i in range(1, len(parts)):
            all_paths.add("/".join(parts[:i]))
    all_paths.discard("")

    # items grouped by their exact category
    by_cat = {}
    for r in rows:
        by_cat.setdefault((r.get("category_path") or "").strip(), []).append(r)

    # group category paths under their top-level name
    tops = {}
    for p in all_paths:
        tops.setdefault(p.split("/")[0], []).append(p)
    if by_cat.get(""):
        tops.setdefault(UNCATEGORIZED, [])

    generated = set()
    for top in sorted(tops.keys()):
        sections = sorted(tops[top])
        if top == UNCATEGORIZED:
            sections = [""]
        fname = safe_filename(top) + ".csv"
        target = out_dir / fname
        total = sum(len(by_cat.get(s, [])) for s in sections)
        try:
            with target.open("w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow([HIER_MARKER])
                w.writerow([f"# Category: {top}"])
                w.writerow([f"# Subcategories: {len(sections)}   Items: {total}"])
                w.writerow([f"# Updated: {now_str()}"])
                w.writerow(["# Master data lives in inventory.csv - edits here are overwritten."])
                w.writerow([])
                for sec in sections:
                    items = by_cat.get(sec, [])
                    label = sec if sec else UNCATEGORIZED
                    req = cat_map.get(sec, [])
                    w.writerow([f"=== {label} ==="])
                    if req:
                        w.writerow([f"# required fields: {', '.join(req)}"])
                    # only show custom columns that this section actually uses
                    used = []
                    for it in items:
                        for k, v in it.items():
                            if k not in BASE_COLUMNS and str(v).strip() and k not in used:
                                used.append(k)
                    # show the status/colour columns only where they carry data
                    extra_base = [c for c in ("color", "status", "used_note", "used_date")
                                  if any(str(it.get(c, "")).strip() for it in items)]
                    cols = (["id", "name", "quantity"] + used + extra_base
                            + ["date_added", "last_modified"])
                    w.writerow(cols)
                    if not items:
                        w.writerow(["(no items in this category)"])
                    for it in sorted(items, key=lambda x: (x.get("name") or "").lower()):
                        w.writerow([it.get(c, "") for c in cols])
                    w.writerow([])
            generated.add(fname)
        except Exception:
            continue

    # overview index so the folder explains itself at a glance
    try:
        with (out_dir / "_Overview.csv").open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow([HIER_MARKER])
            w.writerow([f"# Updated: {now_str()}"])
            w.writerow([])
            w.writerow(["category_path", "depth", "direct_items",
                        "items_incl_subcategories", "required_fields", "file"])
            for p in sorted(all_paths):
                direct = len(by_cat.get(p, []))
                deep = sum(len(v) for k, v in by_cat.items()
                           if k == p or k.startswith(p + "/"))
                w.writerow([p, p.count("/") + 1, direct, deep,
                            "|".join(cat_map.get(p, [])),
                            safe_filename(p.split("/")[0]) + ".csv"])
            if by_cat.get(""):
                w.writerow([UNCATEGORIZED, 0, len(by_cat[""]), len(by_cat[""]), "",
                            UNCATEGORIZED + ".csv"])
        generated.add("_Overview.csv")
    except Exception:
        pass

    # drop stale files we generated earlier (never touch files we didn't write)
    try:
        for old in out_dir.glob("*.csv"):
            if old.name in generated:
                continue
            try:
                with old.open("r", encoding="utf-8-sig") as f:
                    first = f.readline()
                if first.startswith(HIER_MARKER[:30]):
                    old.unlink()
            except Exception:
                continue
    except Exception:
        pass


# ----------------------------------------------------------------------------
# Deleted-items archive
# ----------------------------------------------------------------------------

DELETED_COLUMNS = ["deleted_at", "deleted_reason", "category_path", "name",
                   "quantity", "color", "status", "used_note", "used_date",
                   "date_added", "last_modified", "other_fields", "id"]


def archive_deleted(item, reason=""):
    """Append a deleted item to deleted_items.csv, keeping all of its details."""
    f = deleted_path()
    extras = {k: v for k, v in item.items()
              if k not in BASE_COLUMNS and str(v).strip()}
    row = {
        "deleted_at": now_str(),
        "deleted_reason": reason or "",
        "other_fields": " | ".join(f"{k}={v}" for k, v in sorted(extras.items())),
    }
    for c in DELETED_COLUMNS:
        row.setdefault(c, item.get(c, ""))
    try:
        new_file = not f.exists()
        with f.open("a", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=DELETED_COLUMNS, extrasaction="ignore")
            if new_file:
                w.writeheader()
            w.writerow(row)
    except Exception:
        pass


def read_deleted(limit=500):
    f = deleted_path()
    if not f.exists():
        return []
    try:
        with f.open("r", newline="", encoding="utf-8-sig") as fh:
            rows = [dict(r) for r in csv.DictReader(fh)]
    except Exception:
        return []
    rows.reverse()
    return rows[:limit]


# ----------------------------------------------------------------------------
# Order list  (a classic "what to buy" sheet: item, qty, link, price, total)
# ----------------------------------------------------------------------------

ORDER_COLUMNS = ["item", "quantity", "unit_price", "link", "note"]


def parse_number(v):
    """
    Read a number the way a person types it: '12', '12.50', '12,50', '1.234,56',
    '€ 9,99'. Returns a float, or 0.0 when there is nothing usable.
    """
    s = re.sub(r"[^\d,.\-]", "", str(v or "").strip())
    if not s:
        return 0.0
    if "," in s and "." in s:
        # whichever separator comes last is the decimal one
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def money(value):
    """Format a number as 1234.50 -> '1,234.50'."""
    try:
        return f"{float(value):,.2f}"
    except Exception:
        return "0.00"


def read_order():
    """Return (rows, total). Each row gains a computed 'line_total'."""
    f = order_path()
    rows = []
    if f.exists():
        try:
            with f.open("r", newline="", encoding="utf-8-sig") as fh:
                for r in csv.DictReader(fh):
                    rows.append({c: (r.get(c) or "") for c in ORDER_COLUMNS})
        except Exception:
            rows = []
    total = 0.0
    for r in rows:
        qty = parse_number(r.get("quantity")) or 0.0
        price = parse_number(r.get("unit_price"))
        line = qty * price
        r["line_total"] = f"{line:.2f}"
        total += line
    return rows, total


def write_order(rows):
    """Persist the order list. Rows without an item name are dropped."""
    f = order_path()
    clean = []
    for r in rows or []:
        item = str(r.get("item") or "").strip()
        if not item:
            continue
        clean.append({c: str(r.get(c) or "").strip() for c in ORDER_COLUMNS})
    with f.open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=ORDER_COLUMNS, extrasaction="ignore")
        w.writeheader()
        w.writerows(clean)
    return clean


def order_export_rows():
    """(header, body_rows, total) for a clean order-list export."""
    rows, total = read_order()
    header = ["#", "Item", "Quantity", f"Unit price ({CURRENCY})",
              f"Total ({CURRENCY})", "Link", "Note"]
    body = []
    for i, r in enumerate(rows, 1):
        body.append([
            str(i),
            r.get("item", ""),
            r.get("quantity", ""),
            money(parse_number(r.get("unit_price"))),
            money(r.get("line_total", 0)),
            r.get("link", ""),
            r.get("note", ""),
        ])
    return header, body, total


def read_history(limit=500):
    _, _, _, hist = paths()
    if not hist.exists():
        return []
    with hist.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = [dict(r) for r in reader]
    rows.reverse()  # newest first
    return rows[:limit]


# ----------------------------------------------------------------------------
# Domain logic
# ----------------------------------------------------------------------------

def all_category_paths():
    """
    Every category path that exists: those explicitly stored (incl. empty ones)
    plus any implied by an item's category_path or by a child path.
    """
    cat_map = read_categories()
    paths_set = set(cat_map.keys())

    rows, _ = read_inventory()
    for r in rows:
        p = (r.get("category_path") or "").strip()
        if p:
            paths_set.add(p)

    # ensure all ancestors exist
    for p in list(paths_set):
        parts = p.split("/")
        for i in range(1, len(parts)):
            paths_set.add("/".join(parts[:i]))

    paths_set.discard("")
    return paths_set, cat_map


def build_tree():
    paths_set, cat_map = all_category_paths()
    rows, custom = read_inventory()

    counts = {}
    for r in rows:
        p = (r.get("category_path") or "").strip()
        # count item toward its category and all ancestors
        parts = p.split("/") if p else []
        for i in range(1, len(parts) + 1):
            key = "/".join(parts[:i])
            counts[key] = counts.get(key, 0) + 1

    # nested structure
    root = {}
    for p in sorted(paths_set):
        parts = p.split("/")
        node = root
        for i, part in enumerate(parts):
            cur_path = "/".join(parts[: i + 1])
            if part not in node:
                node[part] = {
                    "_path": cur_path,
                    "_children": {},
                    "_count": counts.get(cur_path, 0),
                    "_required": cat_map.get(cur_path, []),
                }
            node = node[part]["_children"]

    def to_list(node):
        out = []
        for name in sorted(node.keys()):
            n = node[name]
            out.append(
                {
                    "name": name,
                    "path": n["_path"],
                    "count": n["_count"],
                    "required_fields": n["_required"],
                    "children": to_list(n["_children"]),
                }
            )
        return out

    return to_list(root)


def build_clean_export(category="", include_sub=True, include_dates=True,
                       include_used=True, include_color=False):
    """
    Produce plain rows for export: real inventory data only, with friendly
    column names and none of the program's internal bookkeeping (ids, etc).
    Returns (header, rows).
    """
    rows, custom = read_inventory()

    selected = []
    for r in rows:
        p = (r.get("category_path") or "").strip()
        if category:
            if include_sub:
                if not (p == category or p.startswith(category + "/")):
                    continue
            elif p != category:
                continue
        selected.append(r)

    used_custom = [c for c in custom
                   if any(str(r.get(c, "")).strip() for r in selected)]

    cols = ["category_path", "name", "quantity"] + used_custom
    if include_color:
        cols.append("color")
    if include_used:
        cols += ["status", "used_note", "used_date"]
    if include_dates:
        cols += ["date_added", "last_modified"]
    cols = [c for c in cols if c not in INTERNAL_COLUMNS]

    labels = {
        "category_path": "Category", "name": "Name", "quantity": "Quantity",
        "color": "Colour", "status": "Status", "used_note": "Used for",
        "used_date": "Used on", "date_added": "Added", "last_modified": "Last modified",
    }
    header = [labels.get(c, c) for c in cols]

    out = []
    for r in sorted(selected, key=lambda x: ((x.get("category_path") or "").lower(),
                                             (x.get("name") or "").lower())):
        row = []
        for c in cols:
            v = r.get(c, "")
            if c == "status":
                v = "Used" if str(v).strip() == "used" else "In stock"
            row.append(v)
        out.append(row)
    return header, out


# ----------------------------------------------------------------------------
# Minimal PDF writer
#
# Writes a clean one-or-more page order list without any third-party library,
# so the app keeps working with nothing but Flask installed. Text uses
# Helvetica; figures use Courier so columns line up exactly when right-aligned.
# ----------------------------------------------------------------------------

PAGE_W, PAGE_H = 595, 842          # A4 in points
MARGIN = 40

# Helvetica is proportional; this is a good enough average for truncation.
_HELV_AVG = 0.52
_COURIER_W = 0.60                  # Courier is monospaced: exact


def _pdf_text(s):
    """Escape a string for a PDF content stream."""
    out = str(s if s is not None else "")
    out = out.replace("\\", r"\\").replace("(", r"\(").replace(")", r"\)")
    return out.replace("\r", " ").replace("\n", " ")


def _fit(text, width, size, avg=_HELV_AVG):
    """Truncate text with an ellipsis so it fits the given width."""
    text = str(text or "")
    max_chars = max(1, int(width / (avg * size)))
    return text if len(text) <= max_chars else text[:max_chars - 1] + "…"


def write_order_pdf(path, title, header_note=""):
    """Render the current order list to a PDF at `path`."""
    rows, total = read_order()

    # column geometry: right edges for the numeric columns
    x_num, x_item = MARGIN, MARGIN + 20
    r_qty, r_unit, r_total = 360, 455, PAGE_W - MARGIN
    item_w = r_qty - x_item - 45

    pages, annots = [], []
    lines, page_links = [], []

    def new_page():
        nonlocal lines, page_links
        if lines:
            pages.append(lines)
            annots.append(page_links)
        lines, page_links = [], []

    def txt(x, y, s, size=9, font="F1"):
        lines.append(f"BT /{font} {size} Tf 1 0 0 1 {x:.1f} {y:.1f} Tm "
                     f"({_pdf_text(s)}) Tj ET")

    def right(x_edge, y, s, size=9, font="F3"):
        s = str(s)
        # use the metric that matches the font actually being drawn
        factor = _COURIER_W if font == "F3" else (0.55 if font == "F2" else _HELV_AVG)
        txt(x_edge - len(s) * factor * size, y, s, size, font)

    def rule(y, x0=MARGIN, x1=PAGE_W - MARGIN, w=0.5, grey=0.75):
        lines.append(f"{grey:.2f} G {w} w {x0} {y:.1f} m {x1} {y:.1f} l S 0 G")

    def table_head(y):
        # the currency is stated once in the footer rather than in every header:
        # the base-14 fonts advance the euro sign badly in some PDF viewers
        txt(x_num, y, "#", 8, "F2")
        txt(x_item, y, "ITEM", 8, "F2")
        right(r_qty, y, "QTY", 8, "F2")
        right(r_unit, y, "UNIT PRICE", 8, "F2")
        right(r_total, y, "TOTAL", 8, "F2")
        rule(y - 4)
        return y - 16

    # ---- first page header ----
    y = PAGE_H - MARGIN
    txt(MARGIN, y - 6, title, 17, "F2")
    y -= 26
    txt(MARGIN, y, datetime.now().strftime("%Y-%m-%d %H:%M"), 9)
    if header_note:
        y -= 13
        txt(MARGIN, y, header_note, 9)
    y -= 18
    y = table_head(y)

    for i, r in enumerate(rows, 1):
        # room needed for this entry (name + optional link/note lines)
        extra = (13 if r.get("link") else 0) + (13 if r.get("note") else 0)
        if y - extra < MARGIN + 70:
            new_page()
            y = PAGE_H - MARGIN
            txt(MARGIN, y - 4, f"{title} (continued)", 11, "F2")
            y -= 22
            y = table_head(y)

        txt(x_num, y, str(i), 9, "F3")
        txt(x_item, y, _fit(r.get("item", ""), item_w, 9), 9, "F1")
        right(r_qty, y, r.get("quantity", "") or "1")
        right(r_unit, y, money(parse_number(r.get("unit_price"))))
        right(r_total, y, money(r.get("line_total", 0)))

        if r.get("link"):
            y -= 12
            shown = _fit(r["link"], PAGE_W - x_item - MARGIN, 7.5)
            lines.append("0.20 0.40 0.75 rg")
            txt(x_item, y, shown, 7.5, "F1")
            lines.append("0 g")
            w = len(shown) * _HELV_AVG * 7.5
            page_links.append((x_item, y - 2, x_item + w, y + 8, r["link"]))
        if r.get("note"):
            y -= 12
            lines.append("0.45 g")
            txt(x_item, y, _fit(r["note"], PAGE_W - x_item - MARGIN, 7.5), 7.5)
            lines.append("0 g")
        y -= 16

    # ---- total ----
    if y < MARGIN + 60:
        new_page()
        y = PAGE_H - MARGIN - 20
    rule(y + 6, x0=r_unit - 60, w=0.8, grey=0.4)
    y -= 8
    txt(x_item, y, "TOTAL", 11, "F2")
    right(r_total, y, f"{money(total)}", 11, "F3")
    rule(y - 6, x0=r_unit - 60, w=0.8, grey=0.4)
    y -= 26
    lines.append("0.45 g")
    txt(MARGIN, y, f"{len(rows)} item(s) - all prices in {CURRENCY}", 8)
    lines.append("0 g")
    new_page()

    # ---- assemble the file ----
    objs = {}
    n_pages = len(pages)
    font_ids = {"F1": 3 + n_pages * 2, "F2": 4 + n_pages * 2, "F3": 5 + n_pages * 2}

    kids = " ".join(f"{3 + i * 2} 0 R" for i in range(n_pages))
    objs[1] = "<< /Type /Catalog /Pages 2 0 R >>"
    objs[2] = f"<< /Type /Pages /Count {n_pages} /Kids [{kids}] >>"

    for i, content in enumerate(pages):
        pid, cid = 3 + i * 2, 4 + i * 2
        annot_objs = ""
        if annots[i]:
            parts = []
            for (x0, y0, x1, y1, url) in annots[i]:
                parts.append(
                    f"<< /Type /Annot /Subtype /Link /Border [0 0 0] "
                    f"/Rect [{x0:.1f} {y0:.1f} {x1:.1f} {y1:.1f}] "
                    f"/A << /Type /Action /S /URI /URI ({_pdf_text(url)}) >> >>")
            annot_objs = " /Annots [" + " ".join(parts) + "]"
        objs[pid] = (f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {PAGE_W} {PAGE_H}] "
                     f"/Resources << /Font << "
                     f"/F1 {font_ids['F1']} 0 R /F2 {font_ids['F2']} 0 R "
                     f"/F3 {font_ids['F3']} 0 R >> >> "
                     f"/Contents {cid} 0 R{annot_objs} >>")
        stream = "\n".join(content)
        objs[cid] = ("<< /Length %d >>\nstream\n%s\nendstream"
                     % (len(stream.encode("cp1252", "replace")) + 1, stream))

    for name, base in (("F1", "Helvetica"), ("F2", "Helvetica-Bold"), ("F3", "Courier")):
        objs[font_ids[name]] = (f"<< /Type /Font /Subtype /Type1 /BaseFont /{base} "
                                f"/Encoding /WinAnsiEncoding >>")

    out = bytearray(b"%PDF-1.4\n")
    offsets = {}
    for num in sorted(objs):
        offsets[num] = len(out)
        out += f"{num} 0 obj\n{objs[num]}\nendobj\n".encode("cp1252", "replace")
    xref_at = len(out)
    top = max(objs) + 1
    out += f"xref\n0 {top}\n".encode()
    out += b"0000000000 65535 f \n"
    for num in range(1, top):
        out += f"{offsets.get(num, 0):010d} 00000 n \n".encode()
    out += (f"trailer\n<< /Size {top} /Root 1 0 R >>\nstartxref\n"
            f"{xref_at}\n%%EOF\n").encode()

    Path(path).write_bytes(bytes(out))
    return len(rows), total


def inherited_required_fields(category_path):
    """Required fields from this category and all its ancestors."""
    cat_map = read_categories()
    fields = []
    parts = category_path.split("/") if category_path else []
    for i in range(1, len(parts) + 1):
        key = "/".join(parts[:i])
        for fld in cat_map.get(key, []):
            if fld not in fields:
                fields.append(fld)
    return fields


def sanitize_field_name(name):
    name = name.strip()
    # keep it Excel-friendly and safe as a CSV column header
    name = re.sub(r"[\r\n,]", " ", name)
    return name


# ----------------------------------------------------------------------------
# Self-update from GitHub
#
# Only ever talks to the one repository configured below (auto-detected from
# this checkout's git remote). Nothing is applied without an explicit request,
# and every file that gets overwritten is backed up first.
# ----------------------------------------------------------------------------

# Files the updater is allowed to write. Anything else in the archive is ignored.
UPDATABLE_SUFFIXES = {".py", ".md", ".txt", ".sh", ".bat", ".vbs",
                      ".applescript", ".gitignore", ""}
UPDATE_SKIP = {".git", ".github", "__pycache__", ".claude", ".DS_Store"}


def github_repo():
    """'owner/name' for the update source: env var, then git remote, then default."""
    env = os.environ.get("INVENTORY_REPO", "").strip()
    if env:
        return env.replace("https://github.com/", "").rstrip("/").removesuffix(".git")
    try:
        r = run_hidden(["git", "-C", str(APP_DIR), "remote", "get-url", "origin"],
                       timeout=10)
        url = (r.stdout or "").strip()
        if url:
            m = re.search(r"github\.com[:/]+([^/]+/[^/\s]+?)(?:\.git)?$", url)
            if m:
                return m.group(1)
    except Exception:
        pass
    return DEFAULT_REPO


def _gh_get(url, timeout=15, raw=False, accept=None):
    """
    GET from GitHub over HTTPS. Returns parsed JSON or bytes.

    The Accept header matters: the zipball/archive endpoint answers 415 for
    'application/octet-stream', so binary downloads ask for '*/*'.
    """
    import urllib.request
    if accept is None:
        accept = "*/*" if raw else "application/vnd.github+json"
    req = urllib.request.Request(url, headers={
        "User-Agent": f"InventoryManager/{APP_VERSION}",
        "Accept": accept,
    })
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = resp.read()
    return data if raw else json.loads(data.decode("utf-8"))


def _version_tuple(v):
    return tuple(int(x) for x in re.findall(r"\d+", str(v))[:4]) or (0,)


def check_for_update():
    """Ask GitHub what the latest version is. Never raises."""
    repo = github_repo()
    info = {"repo": repo, "branch": UPDATE_BRANCH, "current": APP_VERSION,
            "latest": None, "update_available": False, "commits": [], "error": None}
    try:
        # version string in the published app.py
        raw_url = (f"https://raw.githubusercontent.com/{repo}/"
                   f"{UPDATE_BRANCH}/app.py")
        text = _gh_get(raw_url, raw=True).decode("utf-8", "replace")
        m = re.search(r'^APP_VERSION\s*=\s*["\']([^"\']+)["\']', text, re.M)
        info["latest"] = m.group(1) if m else None

        # recent commit subjects, so the user can see what they'd be getting
        try:
            commits = _gh_get(f"https://api.github.com/repos/{repo}/commits"
                              f"?sha={UPDATE_BRANCH}&per_page=5")
            info["commits"] = [
                {"sha": c.get("sha", "")[:7],
                 "message": (c.get("commit", {}).get("message") or "").split("\n")[0],
                 "date": (c.get("commit", {}).get("author", {}) or {}).get("date", "")[:10]}
                for c in commits if isinstance(c, dict)
            ]
        except Exception:
            pass  # rate-limited or private: the version check alone is enough

        if info["latest"]:
            info["update_available"] = (
                _version_tuple(info["latest"]) > _version_tuple(APP_VERSION))
    except Exception as e:
        info["error"] = f"{type(e).__name__}: {e}"
    return info


def apply_update():
    """
    Download the repo archive and replace the program files, backing up
    everything it overwrites. Returns (ok, message, details).
    """
    import io
    import zipfile
    import tempfile

    # fail fast and clearly if we could never write the files anyway
    probe = APP_DIR / ".update_write_test"
    try:
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
    except Exception:
        return False, (f"The program folder is not writable:\n{APP_DIR}\n\n"
                       "Move the app somewhere you own (for example your home "
                       "folder) and try again."), {}

    repo = github_repo()
    url = f"https://api.github.com/repos/{repo}/zipball/{UPDATE_BRANCH}"
    try:
        blob = _gh_get(url, timeout=120, raw=True)
    except Exception as e:
        code = getattr(e, "code", None)
        hint = ""
        if code == 404:
            hint = (f" The repository or branch was not found - check that "
                    f"'{repo}' exists and has a '{UPDATE_BRANCH}' branch.")
        elif code == 403:
            hint = (" GitHub is rate-limiting this machine (60 requests per hour "
                    "without a login). Wait a while and try again.")
        elif code == 401:
            hint = " The repository appears to be private."
        elif isinstance(e, OSError) and code is None:
            hint = " Check your internet connection."
        return False, f"Could not download the update: {e}.{hint}", {}

    try:
        zf = zipfile.ZipFile(io.BytesIO(blob))
    except Exception as e:
        return False, f"The downloaded file was not a valid archive: {e}", {}

    names = zf.namelist()
    if not names:
        return False, "The downloaded archive was empty.", {}
    root = names[0].split("/")[0]

    # sanity check: refuse anything that isn't recognisably this project
    if not any(n == f"{root}/app.py" for n in names):
        return False, "The archive does not contain app.py - update aborted.", {}

    staged = {}
    for n in names:
        rel = n[len(root) + 1:]
        if not rel or n.endswith("/"):
            continue
        parts = Path(rel).parts
        if any(p in UPDATE_SKIP for p in parts):
            continue
        if ".." in parts or Path(rel).is_absolute():
            continue                                   # zip-slip guard
        if Path(rel).suffix not in UPDATABLE_SUFFIXES:
            continue
        staged[rel] = zf.read(n)

    if "app.py" not in staged:
        return False, "app.py was missing from the archive - update aborted.", {}

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = APP_DIR / f"backup_before_update_{stamp}"
    changed, added = [], []
    try:
        backup.mkdir(parents=True, exist_ok=True)
        for rel, content in sorted(staged.items()):
            target = APP_DIR / rel
            if target.exists():
                if target.read_bytes() == content:
                    continue                            # identical, skip
                dest = backup / rel
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(target, dest)
                changed.append(rel)
            else:
                added.append(rel)
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(content)
    except Exception as e:
        return False, f"Update failed while writing files: {e}", {
            "backup": str(backup), "changed": changed, "added": added}

    if not changed and not added:
        try:
            backup.rmdir()
        except Exception:
            pass
        return True, "Already up to date - no files needed changing.", {
            "changed": [], "added": []}

    return True, "Update installed.", {
        "backup": str(backup), "changed": changed, "added": added}


# ----------------------------------------------------------------------------
# API routes
# ----------------------------------------------------------------------------

@APP.route("/api/state")
def api_state():
    with LOCK:
        ensure_files()
        rows, custom = read_inventory()
        tree = build_tree()
        d = str(get_data_dir())
        return jsonify(
            {
                "tree": tree,
                "custom_fields": custom,
                "base_columns": BASE_COLUMNS,
                "data_dir": d,
                "item_count": len(rows),
                "can_browse": native_picker_available(),
                "version": APP_VERSION,
                "dir_from_env": bool(env_data_dir()),
                "columns": read_col_settings(),
                "dir_source": data_dir_source(),
                "app_dir": str(APP_DIR),
            }
        )


@APP.route("/api/items")
def api_items():
    """Items in a given category (optionally including subcategories)."""
    with LOCK:
        ensure_files()
        cat = request.args.get("category", "").strip()
        include_sub = request.args.get("include_sub", "0") == "1"
        search = request.args.get("search", "").strip().lower()
        rows, custom = read_inventory()

        def match_cat(p):
            if not cat:
                return True
            if include_sub:
                return p == cat or p.startswith(cat + "/")
            return p == cat

        result = []
        for r in rows:
            p = (r.get("category_path") or "").strip()
            if not match_cat(p):
                continue
            if search:
                blob = " ".join(str(v) for v in r.values()).lower()
                if search not in blob:
                    continue
            result.append(r)
        return jsonify({"items": result, "custom_fields": custom})


@APP.route("/api/item", methods=["POST"])
def api_item_save():
    """Create or update a single item."""
    with LOCK:
        ensure_files()
        data = request.get_json(force=True)
        rows, custom = read_inventory()

        item_id = data.get("id") or ""
        fields = data.get("fields", {})  # {column: value}
        category = (fields.get("category_path") or "").strip()
        name = (fields.get("name") or "").strip()

        if not name:
            return jsonify({"error": "Name is required."}), 400

        # discover any new custom columns
        for key in fields.keys():
            if key not in BASE_COLUMNS and key not in custom:
                custom.append(sanitize_field_name(key))

        ts = now_str()
        if item_id:
            found = None
            for r in rows:
                if r.get("id") == item_id:
                    found = r
                    break
            if not found:
                return jsonify({"error": "Item not found."}), 404
            old = dict(found)
            for k, v in fields.items():
                found[k] = v
            found["last_modified"] = ts
            changed = [
                k for k in fields
                if str(old.get(k, "")) != str(fields.get(k, ""))
            ]
            log_history("edit_item", f"{category}/{name}",
                        "changed: " + ", ".join(changed) if changed else "no field changes")
        else:
            new_id = uuid.uuid4().hex[:12]
            row = {c: "" for c in BASE_COLUMNS + custom}
            row.update(fields)
            row["id"] = new_id
            row["date_added"] = ts
            row["last_modified"] = ts
            rows.append(row)
            log_history("add_item", f"{category}/{name}", f"id={new_id}")

        write_inventory(rows, custom)
        return jsonify({"ok": True})


@APP.route("/api/item/delete", methods=["POST"])
def api_item_delete():
    with LOCK:
        ensure_files()
        data = request.get_json(force=True)
        item_id = data.get("id")
        rows, custom = read_inventory()
        target = None
        for r in rows:
            if r.get("id") == item_id:
                target = r
                break
        if not target:
            return jsonify({"error": "Item not found."}), 404
        reason = (data.get("reason") or "").strip()
        archive_deleted(target, reason)          # keep a permanent record
        rows = [r for r in rows if r.get("id") != item_id]
        write_inventory(rows, custom)
        log_history("delete_item",
                    f"{target.get('category_path','')}/{target.get('name','')}",
                    f"id={item_id}" + (f", reason: {reason}" if reason else ""))
        return jsonify({"ok": True})


@APP.route("/api/item/status", methods=["POST"])
def api_item_status():
    """Mark an item as used (with a note about where), or put it back in stock."""
    with LOCK:
        ensure_files()
        data = request.get_json(force=True)
        item_id = data.get("id")
        status = (data.get("status") or "").strip().lower()
        note = (data.get("used_note") or "").strip()
        if status not in ("", "used"):
            return jsonify({"error": "Unknown status."}), 400

        rows, custom = read_inventory()
        target = next((r for r in rows if r.get("id") == item_id), None)
        if not target:
            return jsonify({"error": "Item not found."}), 404

        target["status"] = status
        target["used_note"] = note if status == "used" else ""
        target["used_date"] = now_str() if status == "used" else ""
        target["last_modified"] = now_str()
        write_inventory(rows, custom)
        log_history(
            "mark_used" if status == "used" else "return_to_stock",
            f"{target.get('category_path','')}/{target.get('name','')}",
            note or "")
        return jsonify({"ok": True})


@APP.route("/api/items/bulk", methods=["POST"])
def api_items_bulk():
    """Apply one action to several items at once: delete, move, colour, status."""
    with LOCK:
        ensure_files()
        data = request.get_json(force=True)
        ids = [str(i) for i in (data.get("ids") or [])]
        action = (data.get("action") or "").strip()
        if not ids:
            return jsonify({"error": "Nothing selected."}), 400

        rows, custom = read_inventory()
        wanted = set(ids)
        targets = [r for r in rows if r.get("id") in wanted]
        if not targets:
            return jsonify({"error": "None of those items exist any more."}), 404

        ts = now_str()
        names = ", ".join((t.get("name") or "?") for t in targets[:5])
        if len(targets) > 5:
            names += f" +{len(targets) - 5} more"

        if action == "delete":
            reason = (data.get("reason") or "").strip()
            for t in targets:
                archive_deleted(t, reason)
            rows = [r for r in rows if r.get("id") not in wanted]
            write_inventory(rows, custom)
            log_history("delete_items", f"{len(targets)} item(s)",
                        names + (f" - reason: {reason}" if reason else ""))

        elif action == "move":
            new_cat = (data.get("category_path") or "").strip()
            for t in targets:
                t["category_path"] = new_cat
                t["last_modified"] = ts
            write_inventory(rows, custom)
            log_history("move_items", f"{len(targets)} item(s)",
                        f"{names} -> '{new_cat or '(none)'}'")

        elif action == "color":
            color = (data.get("color") or "").strip()
            for t in targets:
                t["color"] = color
                t["last_modified"] = ts
            write_inventory(rows, custom)
            log_history("colour_items", f"{len(targets)} item(s)",
                        f"{names} -> {color or 'no colour'}")

        elif action == "status":
            status = (data.get("status") or "").strip().lower()
            if status not in ("", "used"):
                return jsonify({"error": "Unknown status."}), 400
            note = (data.get("used_note") or "").strip()
            for t in targets:
                t["status"] = status
                t["used_note"] = note if status == "used" else ""
                t["used_date"] = ts if status == "used" else ""
                t["last_modified"] = ts
            write_inventory(rows, custom)
            log_history("mark_used_items" if status else "return_to_stock_items",
                        f"{len(targets)} item(s)", f"{names}. {note}".strip())

        elif action == "rename":
            new_name = (data.get("name") or "").strip()
            if not new_name:
                return jsonify({"error": "A name is required."}), 400
            if len(targets) != 1:
                return jsonify({"error": "Rename works on one item at a time."}), 400
            old = targets[0].get("name", "")
            targets[0]["name"] = new_name
            targets[0]["last_modified"] = ts
            write_inventory(rows, custom)
            log_history("rename_item", targets[0].get("category_path", ""),
                        f"'{old}' -> '{new_name}'")
        else:
            return jsonify({"error": f"Unknown action '{action}'."}), 400

        return jsonify({"ok": True, "count": len(targets)})


@APP.route("/api/columns", methods=["POST"])
def api_columns():
    """Rename, delete, reorder, hide or relabel a column."""
    with LOCK:
        ensure_files()
        data = request.get_json(force=True)
        action = (data.get("action") or "").strip()

        if action == "rename":
            old = (data.get("name") or "").strip()
            new = (data.get("new_name") or "").strip()
            if old in BASE_COLUMNS:
                # built-in columns keep their field name; only the heading changes
                s = read_col_settings()
                if new and new != old:
                    s["labels"][old] = new
                else:
                    s["labels"].pop(old, None)
                write_col_settings(s)
                log_history("relabel_column", old, f"heading -> {new or '(default)'}")
                return jsonify({"ok": True, "relabelled": True})
            ok, err = rename_custom_field(old, new)
            return (jsonify({"ok": True}) if ok else (jsonify({"error": err}), 400))

        if action == "delete":
            ok, err = delete_custom_field((data.get("name") or "").strip())
            return (jsonify({"ok": True}) if ok else (jsonify({"error": err}), 400))

        if action in ("order", "hidden", "settings"):
            s = read_col_settings()
            if isinstance(data.get("order"), list):
                s["order"] = [str(c) for c in data["order"]]
            if isinstance(data.get("hidden"), list):
                s["hidden"] = [str(c) for c in data["hidden"] if c != "name"]
            if isinstance(data.get("labels"), dict):
                s["labels"] = {str(k): str(v) for k, v in data["labels"].items() if v}
            write_col_settings(s)
            return jsonify({"ok": True, "columns": s})

        return jsonify({"error": f"Unknown action '{action}'."}), 400


@APP.route("/api/deleted")
def api_deleted():
    with LOCK:
        ensure_files()
        return jsonify({"deleted": read_deleted(), "file": str(deleted_path())})


@APP.route("/api/order")
def api_order_get():
    with LOCK:
        ensure_files()
        rows, total = read_order()
        return jsonify({"rows": rows, "total": f"{total:.2f}",
                        "total_display": money(total), "currency": CURRENCY})


@APP.route("/api/order", methods=["POST"])
def api_order_save():
    with LOCK:
        ensure_files()
        data = request.get_json(force=True)
        saved = write_order(data.get("rows") or [])
        rows, total = read_order()
        log_history("order_list", "saved", f"{len(saved)} line(s), total {money(total)}")
        return jsonify({"ok": True, "rows": rows,
                        "total_display": money(total), "currency": CURRENCY})


@APP.route("/api/order/add", methods=["POST"])
def api_order_add():
    """Append one inventory item (or a typed name) to the order list."""
    with LOCK:
        ensure_files()
        data = request.get_json(force=True)
        name = (data.get("item") or "").strip()
        if not name:
            return jsonify({"error": "Item name is required."}), 400
        rows, _ = read_order()
        rows.append({"item": name,
                     "quantity": str(data.get("quantity") or "1"),
                     "unit_price": str(data.get("unit_price") or ""),
                     "link": str(data.get("link") or ""),
                     "note": str(data.get("note") or "")})
        write_order(rows)
        rows, total = read_order()
        return jsonify({"ok": True, "count": len(rows), "total_display": money(total)})


@APP.route("/api/order/export", methods=["POST"])
def api_order_export():
    """Export the order list as a clean CSV or as a PDF."""
    with LOCK:
        ensure_files()
        data = request.get_json(silent=True) or {}
        fmt = (data.get("format") or "csv").lower()
        if fmt not in ("csv", "pdf"):
            return jsonify({"error": "Unknown format."}), 400

        header, body, total = order_export_rows()
        if not body:
            return jsonify({"error": "The order list is empty."}), 400

        stamp = datetime.now().strftime("%Y-%m-%d")
        default = f"order_list_{stamp}.{fmt}"
        target = (data.get("path") or "").strip() or pick_save_file_native(default)
        if not target:
            return jsonify({"cancelled": True})
        if not target.lower().endswith("." + fmt):
            target += "." + fmt

        try:
            if fmt == "csv":
                with open(target, "w", newline="", encoding="utf-8-sig") as f:
                    w = csv.writer(f)
                    w.writerow(header)
                    w.writerows(body)
                    w.writerow([])
                    w.writerow(["", "TOTAL", "", "", money(total), "", ""])
                count = len(body)
            else:
                count, total = write_order_pdf(
                    target,
                    (data.get("title") or "Order list").strip() or "Order list",
                    (data.get("note") or "").strip())
        except Exception as e:
            return jsonify({"error": f"Could not write the file: {e}"}), 400

        log_history("order_export", fmt.upper(),
                    f"{count} line(s), total {money(total)} -> {target}")
        return jsonify({"ok": True, "path": target, "count": count,
                        "total_display": money(total)})


@APP.route("/api/export", methods=["POST"])
def api_export():
    """Write a clean CSV of the current data to a location the user picks."""
    with LOCK:
        ensure_files()
        data = request.get_json(silent=True) or {}
        category = (data.get("category") or "").strip()
        header, rows = build_clean_export(
            category=category,
            include_sub=bool(data.get("include_sub", True)),
            include_dates=bool(data.get("include_dates", True)),
            include_used=bool(data.get("include_used", True)),
            include_color=bool(data.get("include_color", False)),
        )
        if not rows:
            return jsonify({"error": "There is nothing to export in that selection."}), 400

        default = (safe_filename(category.replace("/", "-")) or "inventory") + "_export.csv"
        target = (data.get("path") or "").strip() or pick_save_file_native(default)
        if not target:
            return jsonify({"cancelled": True})
        if not target.lower().endswith(".csv"):
            target += ".csv"
        try:
            with open(target, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(header)
                w.writerows(rows)
        except Exception as e:
            return jsonify({"error": f"Could not write the file: {e}"}), 400
        log_history("export", category or "(all)", f"{len(rows)} items -> {target}")
        return jsonify({"ok": True, "path": target, "count": len(rows)})


@APP.route("/api/item/move", methods=["POST"])
def api_item_move():
    with LOCK:
        ensure_files()
        data = request.get_json(force=True)
        item_id = data.get("id")
        new_cat = (data.get("category_path") or "").strip()
        rows, custom = read_inventory()
        target = None
        for r in rows:
            if r.get("id") == item_id:
                target = r
                break
        if not target:
            return jsonify({"error": "Item not found."}), 404
        old_cat = target.get("category_path", "")
        target["category_path"] = new_cat
        target["last_modified"] = now_str()
        write_inventory(rows, custom)
        log_history("move_item", f"{new_cat}/{target.get('name','')}",
                    f"from '{old_cat}' to '{new_cat}'")
        return jsonify({"ok": True})


@APP.route("/api/category", methods=["POST"])
def api_category_add():
    with LOCK:
        ensure_files()
        data = request.get_json(force=True)
        parent = (data.get("parent") or "").strip()
        name = (data.get("name") or "").strip()
        if not name:
            return jsonify({"error": "Category name is required."}), 400
        if "/" in name:
            return jsonify({"error": "Category name cannot contain '/'."}), 400
        path = f"{parent}/{name}" if parent else name
        cat_map = read_categories()
        if path in cat_map:
            return jsonify({"error": "Category already exists."}), 400
        cat_map[path] = []
        write_categories(cat_map)
        log_history("add_category", path, f"parent='{parent}'")
        return jsonify({"ok": True})


@APP.route("/api/category/rename", methods=["POST"])
def api_category_rename():
    with LOCK:
        ensure_files()
        data = request.get_json(force=True)
        old_path = (data.get("path") or "").strip()
        new_name = (data.get("new_name") or "").strip()
        if not old_path or not new_name:
            return jsonify({"error": "Missing data."}), 400
        if "/" in new_name:
            return jsonify({"error": "Category name cannot contain '/'."}), 400

        parts = old_path.split("/")
        parent = "/".join(parts[:-1])
        new_path = f"{parent}/{new_name}" if parent else new_name

        # update categories
        cat_map = read_categories()
        new_map = {}
        for p, req in cat_map.items():
            if p == old_path or p.startswith(old_path + "/"):
                np = new_path + p[len(old_path):]
                new_map[np] = req
            else:
                new_map[p] = req
        write_categories(new_map)

        # update items
        rows, custom = read_inventory()
        for r in rows:
            p = r.get("category_path", "")
            if p == old_path or p.startswith(old_path + "/"):
                r["category_path"] = new_path + p[len(old_path):]
                r["last_modified"] = now_str()
        write_inventory(rows, custom)
        log_history("rename_category", old_path, f"-> {new_path}")
        return jsonify({"ok": True, "new_path": new_path})


@APP.route("/api/category/delete", methods=["POST"])
def api_category_delete():
    with LOCK:
        ensure_files()
        data = request.get_json(force=True)
        path = (data.get("path") or "").strip()
        mode = data.get("mode", "block")  # block | items_only | recursive
        if not path:
            return jsonify({"error": "Missing category."}), 400

        rows, custom = read_inventory()
        affected = [
            r for r in rows
            if r.get("category_path", "") == path
            or r.get("category_path", "").startswith(path + "/")
        ]

        if affected and mode == "block":
            return jsonify(
                {"error": "not_empty", "count": len(affected)}
            ), 409

        if mode == "recursive":
            # archive them exactly like a normal delete, so nothing vanishes
            # without a record
            for t in affected:
                archive_deleted(t, f"category '{path}' deleted")
            affected_ids = {r.get("id") for r in affected}
            rows = [r for r in rows if r.get("id") not in affected_ids]
            write_inventory(rows, custom)

        # remove category + descendants from categories.csv
        cat_map = read_categories()
        cat_map = {
            p: v for p, v in cat_map.items()
            if not (p == path or p.startswith(path + "/"))
        }
        write_categories(cat_map)
        log_history("delete_category", path, f"mode={mode}, items={len(affected)}")
        return jsonify({"ok": True})


@APP.route("/api/category/required", methods=["POST"])
def api_category_required():
    """Set the required custom fields for a category (auto-applied to new items)."""
    with LOCK:
        ensure_files()
        data = request.get_json(force=True)
        path = (data.get("path") or "").strip()
        fields = data.get("fields", [])
        fields = [sanitize_field_name(f) for f in fields if f.strip()]
        cat_map = read_categories()
        cat_map[path] = fields
        write_categories(cat_map)

        # make sure these columns exist in inventory header
        rows, custom = read_inventory()
        changed = False
        for f in fields:
            if f not in BASE_COLUMNS and f not in custom:
                custom.append(f)
                changed = True
        if changed:
            write_inventory(rows, custom)
        log_history("set_required_fields", path, ", ".join(fields))
        return jsonify({"ok": True, "required_fields": fields})


@APP.route("/api/required_for")
def api_required_for():
    with LOCK:
        ensure_files()
        cat = request.args.get("category", "").strip()
        return jsonify({"required_fields": inherited_required_fields(cat)})


@APP.route("/api/history")
def api_history():
    with LOCK:
        ensure_files()
        return jsonify({"history": read_history()})


@APP.route("/api/browse_dir", methods=["POST"])
def api_browse_dir():
    """Open the OS folder picker on the server machine and return the choice."""
    path = pick_directory_native()
    if not path:
        return jsonify({"cancelled": True})
    return jsonify({"path": path})


@APP.route("/api/update/check")
def api_update_check():
    return jsonify(check_for_update())


@APP.route("/api/update/apply", methods=["POST"])
def api_update_apply():
    with LOCK:
        ok, message, details = apply_update()
    payload = {"ok": ok, "message": message}
    payload.update(details)
    return (jsonify(payload), 200 if ok else 400)


@APP.route("/api/update/restart", methods=["POST"])
def api_update_restart():
    """Relaunch the server so freshly updated code takes effect."""
    def _restart():
        import time
        time.sleep(1.0)                      # let this response reach the browser
        try:
            os.execv(sys.executable, [sys.executable] + sys.argv)
        except Exception:
            os._exit(1)
    threading.Thread(target=_restart, daemon=True).start()
    return jsonify({"ok": True})


@APP.route("/api/import_file", methods=["POST"])
def api_import_file():
    """Pick a spreadsheet with the OS dialog and import it (sheets -> subcategories)."""
    data = request.get_json(silent=True) or {}
    path = (data.get("path") or "").strip()
    if not path:
        path = pick_file_native()
    if not path:
        return jsonify({"cancelled": True})
    with LOCK:
        try:
            import import_data  # imported lazily: import_data imports this module
            r = import_data.import_file(path, quiet=True)
        except SystemExit as e:
            return jsonify({"error": str(e)}), 400
        except Exception as e:
            return jsonify({"error": f"Import failed: {e}"}), 400
    return jsonify({"ok": True, "imported": r["added"], "file": r["file"],
                    "categories": [c["path"] for c in r["categories"]]})


@APP.route("/api/import_folder", methods=["POST"])
def api_import_folder():
    """
    Import every spreadsheet in a folder. With preview=true nothing is written -
    it just reports what would happen, so a bulk import can be checked first.
    """
    data = request.get_json(silent=True) or {}
    path = (data.get("path") or "").strip()
    if not path:
        path = pick_directory_native()
    if not path:
        return jsonify({"cancelled": True})

    preview = bool(data.get("preview"))
    with LOCK:
        try:
            import import_data
            r = import_data.import_folder(
                path,
                under=(data.get("under") or "").strip() or None,
                recursive=bool(data.get("recursive")),
                dry_run=preview,
                quiet=True,
            )
        except SystemExit as e:
            return jsonify({"error": str(e)}), 400
        except Exception as e:
            return jsonify({"error": f"Import failed: {e}"}), 400

    return jsonify({
        "ok": True, "preview": preview, "folder": r["folder"],
        "added": r["added"], "skipped": r.get("skipped", 0),
        "message": r.get("message", ""),
        "failed": r.get("failed", []),
        "files": [{"file": f["file"], "added": f["added"],
                   "colored": f["colored"],
                   "categories": [c["path"] for c in f["categories"]]}
                  for f in r["files"]],
    })


@APP.route("/api/set_dir", methods=["POST"])
def api_set_dir():
    with LOCK:
        data = request.get_json(force=True)
        raw = (data.get("path") or "").strip()
        if not raw:
            return jsonify({"error": "Path is required."}), 400
        try:
            p = Path(raw).expanduser()
            p.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            return jsonify({"error": f"Could not use that folder: {e}"}), 400

        scope = (data.get("scope") or "shared").strip()
        old_dir = get_data_dir()
        migrated = False
        # If we're switching to a fresh, empty folder but already have data,
        # copy the CSVs across so the data "moves" with your choice (never deletes source).
        try:
            switching = p.resolve() != old_dir.resolve()
        except Exception:
            switching = True
        if switching and not (p / INVENTORY_CSV).exists() and (old_dir / INVENTORY_CSV).exists():
            try:
                for name in (INVENTORY_CSV, CATEGORIES_CSV, HISTORY_CSV):
                    src = old_dir / name
                    if src.exists():
                        shutil.copy2(src, p / name)
                migrated = True
            except Exception:
                migrated = False

        used = set_data_dir(p, scope)
        ensure_files()
        return jsonify({"ok": True, "data_dir": str(p), "migrated": migrated,
                        "scope": used,
                        "shared_failed": (scope == "shared" and used != "shared")})


@APP.route("/")
def index():
    return Response(INDEX_HTML, mimetype="text/html")


# ----------------------------------------------------------------------------
# Front-end (single-page, embedded)
# ----------------------------------------------------------------------------

INDEX_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Inventory Manager</title>
<style>
  :root{
    --bg:#0f1419;
    --panel:#171d26;
    --panel-2:#1e2530;
    --line:#2a3341;
    --text:#e6ebf1;
    --muted:#8b98a9;
    --accent:#4c9be8;
    --accent-soft:#22344a;
    --danger:#e5646e;
    --ok:#5bbf8f;
    --on-accent:#06121f;   /* text color on top of --accent */
    --radius:8px;
    font-size:15px;
  }
  :root[data-theme="light"]{
    --bg:#f4f6f9;
    --panel:#ffffff;
    --panel-2:#eceff4;
    --line:#d7dde5;
    --text:#1b2431;
    --muted:#5f6b7c;
    --accent:#2f7fd1;
    --accent-soft:#dcebfb;
    --danger:#d1454f;
    --ok:#2e9c6d;
    --on-accent:#ffffff;
  }
  *{box-sizing:border-box;}
  body{
    margin:0;font-family:"Segoe UI",system-ui,-apple-system,sans-serif;
    background:var(--bg);color:var(--text);height:100vh;overflow:hidden;
  }
  button{font-family:inherit;cursor:pointer;}
  .app{display:flex;flex-direction:column;height:100vh;}

  header{
    display:flex;align-items:center;gap:14px;padding:10px 18px;
    background:var(--panel);border-bottom:1px solid var(--line);
  }
  header h1{font-size:16px;font-weight:600;margin:0;letter-spacing:.2px;}
  header .path{font-size:12px;color:var(--muted);margin-left:2px;}
  header .spacer{flex:1;}
  .version-badge{
    background:transparent;border:1px solid transparent;color:var(--muted);
    font-size:11px;padding:2px 7px;border-radius:20px;margin-left:8px;
  }
  .version-badge:hover{border-color:var(--line);color:var(--text);}
  .version-badge.has-update{
    background:var(--accent-soft);border-color:var(--accent);
    color:var(--accent);font-weight:600;
  }
  .upd-row{display:grid;grid-template-columns:70px 1fr;gap:8px;font-size:12px;padding:5px 0;
    border-bottom:1px solid var(--line);}
  .upd-row .sha{color:var(--muted);font-family:ui-monospace,Menlo,Consolas,monospace;}
  .btn{
    background:var(--panel-2);border:1px solid var(--line);color:var(--text);
    padding:7px 12px;border-radius:6px;font-size:13px;transition:.12s;
  }
  .btn:hover{border-color:var(--accent);}
  .btn.primary{background:var(--accent);border-color:var(--accent);color:var(--on-accent);font-weight:600;}
  .btn.primary:hover{filter:brightness(1.08);}
  .btn.ghost{background:transparent;}
  .btn.danger{color:var(--danger);border-color:transparent;background:transparent;}
  .btn.danger:hover{border-color:var(--danger);}
  .btn.small{padding:4px 8px;font-size:12px;}

  main{display:flex;flex:1;min-height:0;}
  .sidebar{
    width:320px;min-width:220px;max-width:520px;background:var(--panel);
    border-right:1px solid var(--line);display:flex;flex-direction:column;
  }
  .sidebar .head{
    display:flex;align-items:center;justify-content:space-between;
    padding:12px 14px 8px;border-bottom:1px solid var(--line);
  }
  .sidebar .head span{font-size:12px;text-transform:uppercase;letter-spacing:1px;color:var(--muted);}
  .tree{flex:1;overflow:auto;padding:6px;}
  .content{flex:1;display:flex;flex-direction:column;min-width:0;}

  .node{user-select:none;}
  .node-row{
    display:flex;align-items:center;gap:4px;padding:5px 6px;border-radius:6px;
    font-size:14px;white-space:nowrap;
  }
  .node-row:hover{background:var(--panel-2);}
  .node-row.selected{background:var(--accent-soft);}
  .twist{
    width:16px;text-align:center;color:var(--muted);font-size:11px;
    flex:0 0 16px;cursor:pointer;
  }
  .twist.empty{visibility:hidden;}
  .node-name{flex:1;overflow:hidden;text-overflow:ellipsis;}
  .node-count{color:var(--muted);font-size:12px;margin-left:6px;}
  .node-actions{display:none;gap:2px;}
  .node-row:hover .node-actions{display:flex;}
  .icon-btn{
    background:transparent;border:none;color:var(--muted);padding:2px 5px;
    border-radius:4px;font-size:13px;line-height:1;
  }
  .icon-btn:hover{background:var(--line);color:var(--text);}
  .children{margin-left:14px;border-left:1px solid var(--line);}

  .toolbar{
    display:flex;align-items:center;gap:10px;padding:12px 18px;
    border-bottom:1px solid var(--line);background:var(--panel);
  }
  .toolbar h2{font-size:15px;margin:0;font-weight:600;}
  .toolbar .crumb{color:var(--muted);font-weight:400;}
  .search{
    background:var(--bg);border:1px solid var(--line);color:var(--text);
    padding:7px 10px;border-radius:6px;font-size:13px;width:220px;
  }
  .search:focus{outline:none;border-color:var(--accent);}
  label.check{display:flex;align-items:center;gap:6px;font-size:12px;color:var(--muted);cursor:pointer;}

  .table-wrap{flex:1;overflow:auto;}
  table{border-collapse:collapse;width:100%;font-size:13px;}
  th,td{
    text-align:left;padding:8px 12px;border-bottom:1px solid var(--line);
    white-space:nowrap;max-width:320px;overflow:hidden;text-overflow:ellipsis;
  }
  th{
    position:sticky;top:0;background:var(--panel-2);color:var(--muted);
    font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.5px;z-index:1;
  }
  tbody tr{cursor:pointer;}
  tbody tr:hover{background:var(--panel-2);}
  .swatch{width:6px;height:30px;border-radius:3px;}
  .dot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:7px;
       vertical-align:middle;box-shadow:0 0 0 1px var(--line) inset;}
  tbody tr.is-used td{opacity:.62;font-style:italic;}

  /* rows carry their own background so the pinned columns can inherit it */
  tbody tr{background:var(--bg);}
  tbody tr:hover{background:var(--panel-2);}
  tbody tr.is-picked{background:var(--accent-soft);}

  /* pinned columns: the checkbox + name stay on the left, actions on the right,
     so with many columns you always know which row you are acting on */
  .col-name{position:sticky;left:0;z-index:2;font-weight:500;
            box-shadow:1px 0 0 var(--line);max-width:none;}
  .col-act{position:sticky;right:0;z-index:2;box-shadow:-1px 0 0 var(--line);text-align:right;}
  tbody td.col-name,tbody td.col-act{background:inherit;}
  th.col-name,th.col-act{background:var(--panel-2);z-index:3;}
  .colhead{display:inline-flex;align-items:center;gap:5px;cursor:context-menu;}
  .colmenu{opacity:0;color:var(--muted);cursor:pointer;padding:0 2px;border-radius:3px;}
  th:hover .colmenu{opacity:1;}
  .colmenu:hover{background:var(--line);color:var(--text);}
  .pickwrap{display:inline-flex;align-items:center;margin-right:9px;vertical-align:middle;}
  .rowpick,#selAll{cursor:pointer;margin:0;}

  .sel-bar{
    display:flex;align-items:center;gap:8px;padding:8px 18px;
    background:var(--accent-soft);border-bottom:1px solid var(--line);font-size:13px;
  }

  .ctx-menu{
    position:fixed;z-index:200;min-width:196px;padding:5px;
    background:var(--panel);border:1px solid var(--line);border-radius:9px;
    box-shadow:0 12px 34px rgba(0,0,0,.35);
  }
  .ctx-menu .ctx-head{
    padding:6px 10px 8px;font-size:11px;color:var(--muted);
    border-bottom:1px solid var(--line);margin-bottom:4px;
    overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:240px;
  }
  .ctx-menu button{
    display:block;width:100%;text-align:left;background:none;border:none;
    color:var(--text);padding:7px 10px;border-radius:6px;font-size:13px;
  }
  .ctx-menu button:hover{background:var(--panel-2);}
  .empty-state{padding:60px 20px;text-align:center;color:var(--muted);}
  .empty-state h3{color:var(--text);font-weight:600;margin:0 0 6px;}

  /* modal */
  .overlay{
    position:fixed;inset:0;background:rgba(4,8,12,.6);display:none;
    align-items:center;justify-content:center;z-index:50;padding:20px;
  }
  .overlay.show{display:flex;}
  .modal{
    background:var(--panel);border:1px solid var(--line);border-radius:12px;
    width:100%;max-width:560px;max-height:88vh;overflow:auto;
    box-shadow:0 24px 60px rgba(0,0,0,.5);
  }
  .modal .m-head{
    display:flex;align-items:center;justify-content:space-between;
    padding:16px 20px;border-bottom:1px solid var(--line);
  }
  .modal .m-head h3{margin:0;font-size:15px;}
  .modal .m-body{padding:18px 20px;}
  .modal .m-foot{
    display:flex;justify-content:flex-end;gap:10px;padding:14px 20px;
    border-top:1px solid var(--line);
  }
  .field{margin-bottom:14px;}
  .field label{display:block;font-size:12px;color:var(--muted);margin-bottom:5px;}
  .field input,.field select,.field textarea{
    width:100%;background:var(--bg);border:1px solid var(--line);color:var(--text);
    padding:8px 10px;border-radius:6px;font-size:14px;font-family:inherit;
  }
  .field input:focus,.field select:focus,.field textarea:focus{outline:none;border-color:var(--accent);}
  .field .req{color:var(--accent);}
  .custom-row{display:flex;gap:8px;align-items:flex-end;margin-bottom:10px;}
  .custom-row .field{flex:1;margin-bottom:0;}
  .muted{color:var(--muted);font-size:12px;}
  .chip{
    display:inline-flex;align-items:center;gap:6px;background:var(--accent-soft);
    color:var(--text);padding:4px 8px;border-radius:20px;font-size:12px;margin:0 6px 6px 0;
  }
  .chip button{background:none;border:none;color:var(--muted);cursor:pointer;font-size:13px;}
  .hist-item{
    display:grid;grid-template-columns:150px 120px 1fr;gap:10px;padding:8px 4px;
    border-bottom:1px solid var(--line);font-size:12px;
  }
  .hist-item .ts{color:var(--muted);}
  .hist-item .act{color:var(--accent);}
  .toast{
    position:fixed;bottom:20px;left:50%;transform:translateX(-50%);
    background:var(--panel-2);border:1px solid var(--line);color:var(--text);
    padding:10px 18px;border-radius:8px;font-size:13px;opacity:0;transition:.25s;
    pointer-events:none;z-index:100;
  }
  .toast.show{opacity:1;}
  .toast.ok{border-color:var(--ok);}
  .toast.err{border-color:var(--danger);}
  ::-webkit-scrollbar{width:10px;height:10px;}
  ::-webkit-scrollbar-thumb{background:var(--line);border-radius:6px;}
  ::-webkit-scrollbar-track{background:transparent;}
</style>
</head>
<body>
<div class="app">
  <header>
    <h1>Inventory Manager</h1>
    <span class="path" id="dataPath"></span>
    <button class="version-badge" id="versionBadge" onclick="openUpdateModal()" title="Check for updates"></button>
    <span class="spacer"></span>
    <button class="btn ghost small" id="themeToggle" onclick="toggleTheme()" title="Toggle light / dark">🌙</button>
    <button class="btn ghost small" onclick="openAppearanceModal()" title="Appearance — light-mode background colour">🎨</button>
    <button class="btn ghost small" onclick="openImportModal()" title="Import a spreadsheet, or every spreadsheet in a folder">Import</button>
    <button class="btn ghost small" onclick="openOrderModal()" title="Build an order list and export it as CSV or PDF">Order list</button>
    <button class="btn ghost small" onclick="openExportModal()" title="Export a clean CSV">Export</button>
    <button class="btn ghost small" onclick="openHistory()">History</button>
    <button class="btn ghost small" onclick="openDirModal()">Data folder</button>
    <button class="btn primary small" onclick="openItemModal()">+ Item</button>
  </header>

  <main>
    <aside class="sidebar">
      <div class="head">
        <span>Categories</span>
        <button class="icon-btn" title="Add top-level category" onclick="addCategory('')">＋</button>
      </div>
      <div class="tree" id="tree"></div>
    </aside>

    <section class="content">
      <div class="toolbar">
        <h2 id="crumb"><span class="crumb">All items</span></h2>
        <span class="spacer" style="flex:1"></span>
        <label class="check"><input type="checkbox" id="includeSub" onchange="loadItems()"> include subcategories</label>
        <label class="check"><input type="checkbox" id="fillRows" onchange="toggleFillRows()"> fill whole row with colour</label>
        <button class="btn small ghost" onclick="openColumnsModal()" title="Rename, reorder, hide or delete columns">Columns</button>
        <input class="search" id="search" placeholder="Search…" oninput="debouncedLoad()">
      </div>
      <div class="sel-bar" id="selBar" style="display:none">
        <b id="selCount">0 selected</b>
        <button class="btn small" id="selRename" onclick="renameSelected()">Rename…</button>
        <button class="btn small" onclick="bulkMove()">Move to…</button>
        <button class="btn small" onclick="bulkColor()">Colour…</button>
        <button class="btn small" onclick="bulkUsed(false)">Mark used</button>
        <button class="btn small" onclick="bulkUsed(true)">Back in stock</button>
        <button class="btn small danger" onclick="bulkDelete()">Delete…</button>
        <span style="flex:1"></span>
        <button class="btn small ghost" onclick="clearSelection()">Clear</button>
      </div>
      <div class="table-wrap" id="tableWrap"></div>
    </section>
  </main>
</div>

<div class="overlay" id="overlay"></div>
<div class="toast" id="toast"></div>

<script>
const $ = s => document.querySelector(s);

// ---------- colour helpers ----------
function hex2rgb(h){
  h=String(h||"").replace("#","");
  if(h.length===3) h=h.split("").map(c=>c+c).join("");
  return [parseInt(h.slice(0,2),16)||0, parseInt(h.slice(2,4),16)||0, parseInt(h.slice(4,6),16)||0];
}
function rgb2hex(r){ return "#"+r.map(v=>Math.max(0,Math.min(255,Math.round(v))).toString(16).padStart(2,"0")).join(""); }
function mix(a,b,t){ const x=hex2rgb(a),y=hex2rgb(b); return rgb2hex(x.map((v,i)=>v+(y[i]-v)*t)); }
function lum(h){ const [r,g,b]=hex2rgb(h); return (0.299*r+0.587*g+0.114*b)/255; }

// ---------- theme (light / dark) ----------
const LIGHT_BG_DEFAULT = "#f4f6f9";
function storedLightBg(){
  try{ return localStorage.getItem("inv_light_bg") || LIGHT_BG_DEFAULT; }catch(e){ return LIGHT_BG_DEFAULT; }
}

// Derive a coherent light palette from one background colour, so panels,
// borders and text stay sensible whatever background is picked.
function applyLightBackground(bg){
  const root=document.documentElement;
  const vars=["--bg","--panel","--panel-2","--line","--text","--muted","--accent-soft"];
  if(root.getAttribute("data-theme")!=="light"){
    vars.forEach(v=>root.style.removeProperty(v));   // dark theme keeps its own palette
    return;
  }
  bg = bg || LIGHT_BG_DEFAULT;
  const dark = lum(bg) < 0.5;                        // someone picked a dark colour
  const text = dark ? "#f2f5f9" : "#1b2431";
  root.style.setProperty("--bg", bg);
  root.style.setProperty("--panel",   dark ? mix(bg,"#ffffff",0.10) : mix(bg,"#ffffff",0.75));
  root.style.setProperty("--panel-2", dark ? mix(bg,"#ffffff",0.05) : mix(bg,"#000000",0.05));
  root.style.setProperty("--line",    dark ? mix(bg,"#ffffff",0.16) : mix(bg,"#000000",0.16));
  root.style.setProperty("--text", text);
  root.style.setProperty("--muted", mix(text,bg,0.45));
  root.style.setProperty("--accent-soft", mix("#2f7fd1",bg,0.80));
}

function applyTheme(t){
  document.documentElement.setAttribute("data-theme", t);
  try{ localStorage.setItem("inv_theme", t); }catch(e){}
  const b=document.getElementById("themeToggle");
  if(b) b.textContent = (t==="light") ? "☀️" : "🌙";
  applyLightBackground(storedLightBg());
}
function toggleTheme(){
  const cur=document.documentElement.getAttribute("data-theme")||"dark";
  applyTheme(cur==="light" ? "dark" : "light");
}
applyTheme((function(){ try{ return localStorage.getItem("inv_theme"); }catch(e){ return null; } })() || "dark");

// ---------- appearance ----------
const LIGHT_PRESETS = [
  ["Cool grey",  "#f4f6f9"], ["Paper white", "#ffffff"],
  ["Warm cream", "#faf5ec"], ["Soft sand",   "#f5f1e8"],
  ["Mint",       "#eef6f1"], ["Sky",         "#eaf2fb"],
  ["Lavender",   "#f2f0fa"], ["Slate",       "#e7ebf0"],
];

function openAppearanceModal(){
  const current = storedLightBg();
  const isLight = document.documentElement.getAttribute("data-theme")==="light";
  const body=`
    ${isLight?"":`<p class="muted">You are in dark mode — switch to light mode with ☀️ to see these take effect.</p>`}
    <div class="field"><label>Light-mode background</label>
      <div id="bgPresets" style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px">
        ${LIGHT_PRESETS.map(([n,c])=>`
          <button class="btn small bgp" data-c="${c}" title="${esc(n)}"
            style="display:flex;align-items:center;gap:6px">
            <span style="width:14px;height:14px;border-radius:3px;border:1px solid var(--line);background:${c};display:inline-block"></span>
            ${esc(n)}</button>`).join("")}
      </div>
      <div style="display:flex;gap:8px;align-items:center">
        <input type="color" id="bgCustom" value="${esc(current)}" style="width:52px;height:34px;padding:2px">
        <span class="muted">custom colour</span>
        <span class="spacer" style="flex:1"></span>
        <button class="btn small" id="bgReset">Reset to default</button>
      </div>
    </div>
    <p class="muted">Panels, borders and text are derived from this colour automatically, so the
    interface stays readable — including if you choose something dark. Dark mode is unaffected.</p>`;
  const foot=`<button class="btn primary" onclick="closeModal()">Done</button>`;
  const m=modalShell("Appearance", body, foot);
  showModal(m);

  const use = c => {
    try{ localStorage.setItem("inv_light_bg", c); }catch(e){}
    if(document.documentElement.getAttribute("data-theme")!=="light") applyTheme("light");
    else applyLightBackground(c);
    m.querySelector("#bgCustom").value = c;
  };
  m.querySelectorAll(".bgp").forEach(b=>b.onclick=()=>use(b.dataset.c));
  m.querySelector("#bgCustom").oninput=e=>use(e.target.value);
  m.querySelector("#bgReset").onclick=()=>use(LIGHT_BG_DEFAULT);
}

let STATE = {tree:[], custom_fields:[], data_dir:"", item_count:0};
let SELECTED = "";           // selected category path ("" = all)
let EXPANDED = new Set();
let SEARCH_TIMER = null;

// ---------- helpers ----------
function toast(msg, kind="ok"){
  const t=$("#toast"); t.textContent=msg; t.className="toast show "+kind;
  setTimeout(()=>t.className="toast",1900);
}
async function api(url, opts){
  const r = await fetch(url, opts);
  const data = await r.json().catch(()=>({}));
  if(!r.ok){ throw data; }
  return data;
}
function esc(s){return (s==null?"":String(s)).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}

// ---------- state / tree ----------
async function loadState(){
  STATE = await api("/api/state");
  $("#dataPath").textContent = "· " + STATE.data_dir;
  paintVersionBadge();
  renderTree();
  loadItems();
}

function renderTree(){
  const el=$("#tree"); el.innerHTML="";
  const allRow=document.createElement("div");
  allRow.className="node-row"+(SELECTED===""?" selected":"");
  allRow.innerHTML=`<span class="twist empty"></span><span class="node-name">All items</span><span class="node-count">${STATE.item_count}</span>`;
  allRow.onclick=()=>{SELECTED="";renderTree();loadItems();};
  el.appendChild(allRow);
  STATE.tree.forEach(n=>el.appendChild(renderNode(n)));
}

function renderNode(node){
  const wrap=document.createElement("div"); wrap.className="node";
  const row=document.createElement("div");
  row.className="node-row"+(SELECTED===node.path?" selected":"");
  const hasKids=node.children.length>0;
  const open=EXPANDED.has(node.path);
  row.innerHTML=`
    <span class="twist ${hasKids?'':'empty'}">${hasKids?(open?'▾':'▸'):''}</span>
    <span class="node-name" title="${esc(node.path)}">${esc(node.name)}</span>
    <span class="node-count">${node.count}</span>
    <span class="node-actions">
      <button class="icon-btn" title="Add subcategory">＋</button>
      <button class="icon-btn" title="Required fields">⚙</button>
      <button class="icon-btn" title="Rename">✎</button>
      <button class="icon-btn" title="Delete">🗑</button>
    </span>`;
  const [addBtn,reqBtn,renBtn,delBtn]=row.querySelectorAll(".node-actions .icon-btn");
  const twist=row.querySelector(".twist");
  twist.onclick=e=>{e.stopPropagation(); if(!hasKids)return; open?EXPANDED.delete(node.path):EXPANDED.add(node.path); renderTree();};
  row.onclick=()=>{SELECTED=node.path;renderTree();loadItems();};
  addBtn.onclick=e=>{e.stopPropagation();addCategory(node.path);};
  reqBtn.onclick=e=>{e.stopPropagation();openRequiredModal(node);};
  renBtn.onclick=e=>{e.stopPropagation();renameCategory(node);};
  delBtn.onclick=e=>{e.stopPropagation();deleteCategory(node);};
  wrap.appendChild(row);
  if(hasKids&&open){
    const kids=document.createElement("div"); kids.className="children";
    node.children.forEach(c=>kids.appendChild(renderNode(c)));
    wrap.appendChild(kids);
  }
  return wrap;
}

// ---------- colour display ----------
let FILL_ROWS = (function(){ try{ return localStorage.getItem("inv_fill_rows")==="1"; }catch(e){ return false; } })();

// black or white text, whichever stays legible on the given colour
function readableOn(hex){
  const h=String(hex||"").replace("#","");
  if(h.length<6) return "inherit";
  const r=parseInt(h.slice(0,2),16), g=parseInt(h.slice(2,4),16), b=parseInt(h.slice(4,6),16);
  return ((0.299*r + 0.587*g + 0.114*b)/255) > 0.6 ? "#111" : "#fff";
}
function toggleFillRows(){
  FILL_ROWS = $("#fillRows").checked;
  try{ localStorage.setItem("inv_fill_rows", FILL_ROWS?"1":"0"); }catch(e){}
  loadItems();
}

// ---------- selection ----------
let SELECTED_IDS = new Set();
let VISIBLE_IDS = [];
let LAST_ITEMS = [], LAST_CUSTOM = [];
let LAST_CLICKED = null;

function toggleSelect(id, force){
  const on = (force===undefined) ? !SELECTED_IDS.has(id) : force;
  if(on) SELECTED_IDS.add(id); else SELECTED_IDS.delete(id);
  LAST_CLICKED = id;
  paintSelection();
}
function selectRange(id){
  // shift-click selects everything between the last click and this row
  const a=VISIBLE_IDS.indexOf(LAST_CLICKED), b=VISIBLE_IDS.indexOf(id);
  if(a<0||b<0){ toggleSelect(id); return; }
  const [lo,hi]=a<b?[a,b]:[b,a];
  for(let i=lo;i<=hi;i++) SELECTED_IDS.add(VISIBLE_IDS[i]);
  paintSelection();
}
function clearSelection(){ SELECTED_IDS.clear(); paintSelection(); }

function paintSelection(){
  document.querySelectorAll("#tableWrap tbody tr").forEach(tr=>{
    const on=SELECTED_IDS.has(tr.dataset.id);
    tr.classList.toggle("is-picked", on);
    const cb=tr.querySelector(".rowpick"); if(cb) cb.checked=on;
  });
  const all=document.querySelector("#selAll");
  if(all) all.checked = VISIBLE_IDS.length>0 && VISIBLE_IDS.every(id=>SELECTED_IDS.has(id));

  const n=SELECTED_IDS.size, bar=$("#selBar");
  if(!bar) return;
  bar.style.display = n ? "flex" : "none";
  if(n) $("#selCount").textContent = `${n} selected`;
  const one=$("#selRename"); if(one) one.style.display = n===1 ? "" : "none";
}

function selectedIds(){ return [...SELECTED_IDS]; }

async function bulk(payload, okMsg){
  try{
    const r=await api("/api/items/bulk",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({ids:selectedIds(), ...payload})});
    toast(`${okMsg} (${r.count})`);
    clearSelection(); loadState();
  }catch(e){ toast(e.error||"Failed","err"); }
}

async function bulkMove(){
  const opts=await categoryOptions("");
  const body=`<p class="muted">${SELECTED_IDS.size} item(s) will be moved.</p>
    <div class="field"><label>Move to category</label>
      <select id="bulkCat"><option value="">— none —</option>${opts}</select></div>`;
  const m=modalShell("Move items", body,
    `<button class="btn ghost" onclick="closeModal()">Cancel</button>
     <button class="btn primary" id="bulkGo">Move</button>`);
  showModal(m);
  m.querySelector("#bulkGo").onclick=()=>{
    const cat=m.querySelector("#bulkCat").value; closeModal();
    bulk({action:"move",category_path:cat}, "Moved");
  };
}
async function bulkColor(){
  const body=`<p class="muted">${SELECTED_IDS.size} item(s) will be recoloured.</p>
    <div class="field"><label>Colour</label>
      <div style="display:flex;gap:8px;align-items:center">
        <input type="color" id="bulkCol" value="#4c9be8" style="width:52px;height:34px;padding:2px">
        <button class="btn small" id="bulkNoCol">No colour</button>
      </div></div>`;
  const m=modalShell("Set colour", body,
    `<button class="btn ghost" onclick="closeModal()">Cancel</button>
     <button class="btn primary" id="bulkGo">Apply</button>`);
  showModal(m);
  m.querySelector("#bulkNoCol").onclick=()=>{ closeModal(); bulk({action:"color",color:""}, "Colour cleared"); };
  m.querySelector("#bulkGo").onclick=()=>{
    const c=m.querySelector("#bulkCol").value; closeModal();
    bulk({action:"color",color:c}, "Colour set");
  };
}
async function bulkDelete(){
  const n=SELECTED_IDS.size;
  const names=(LAST_ITEMS||[]).filter(i=>SELECTED_IDS.has(i.id)).map(i=>i.name);
  const reason=await askText({
    title:`Delete ${n} item(s)?`,
    message:`${esc(names.slice(0,6).join(", "))}${names.length>6?` +${names.length-6} more`:""}<br>`
           +`They are archived in deleted_items.csv and can be reviewed later.`,
    label:"Reason (optional)", placeholder:"e.g. broken, returned, sold",
    okLabel:"Delete"});
  if(reason===null) return;
  bulk({action:"delete",reason}, "Deleted and archived");
}
async function bulkUsed(back){
  if(!SELECTED_IDS.size) return;
  let note="";
  if(back){
    if(!await askConfirm({title:"Back in stock",
      message:`Put ${SELECTED_IDS.size} item(s) back in stock?`, okLabel:"Yes"})) return;
  }else{
    note=await askText({title:`Mark ${SELECTED_IDS.size} item(s) as used`,
      label:"Used for / where", placeholder:"e.g. installed in Lab 2 / given to Ahmed",
      okLabel:"Mark used"});
    if(note===null) return;
  }
  bulk({action:"status",status:back?"":"used",used_note:note}, back?"Back in stock":"Marked as used");
}
async function renameSelected(){
  const id=selectedIds()[0]; if(!id) return;
  const it=(LAST_ITEMS||[]).find(x=>x.id===id);
  const nn=await askText({title:"Rename item", label:"New name",
                          value:it?it.name:"", okLabel:"Rename"});
  if(nn===null || !nn.trim()) return;
  bulk({action:"rename",name:nn.trim()}, "Renamed");
}

// ---------- columns ----------
let ALL_COLUMNS=[], CUSTOM_COLUMNS=[], LAST_COLS=[], COL_LABEL=(c=>c);

function isCustomCol(c){ return CUSTOM_COLUMNS.includes(c); }

async function saveColSettings(patch){
  try{
    await api("/api/columns",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({action:"settings", ...patch})});
    await loadState();
  }catch(e){ toast(e.error||"Failed","err"); }
}

// Save the order of every column currently on screen, with `col` shifted.
async function moveColumn(col, dir){
  const order=LAST_COLS.filter(c=>c!=="name");   // name is pinned, never moves
  const i=order.indexOf(col);
  const j=i+dir;
  if(i<0||j<0||j>=order.length) return;
  [order[i],order[j]]=[order[j],order[i]];
  await saveColSettings({order:["name",...order]});
  toast(`Moved "${COL_LABEL(col)}" ${dir<0?"left":"right"}`);
}

async function hideColumn(col){
  const hidden=[...((STATE.columns||{}).hidden||[])];
  if(!hidden.includes(col)) hidden.push(col);
  await saveColSettings({hidden});
  toast(`"${COL_LABEL(col)}" hidden`);
}

async function renameColumn(col){
  const nn=await askText({
    title:"Rename column",
    message: isCustomCol(col)
      ? "This renames the field for every item and in the CSV file."
      : "This is a built-in column, so only its heading changes here.",
    label:"Column heading", value:COL_LABEL(col), okLabel:"Rename"});
  if(nn===null||!nn.trim()) return;
  try{
    await api("/api/columns",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({action:"rename",name:col,new_name:nn.trim()})});
    toast("Column renamed"); await loadState();
  }catch(e){ toast(e.error||"Failed","err"); }
}

async function deleteColumn(col){
  if(!isCustomCol(col)){
    toast("Built-in columns can't be deleted — hiding it instead","err");
    return hideColumn(col);
  }
  const filled=(LAST_ITEMS||[]).filter(i=>String(i[col]||"").trim()).length;
  const go=await askConfirm({
    title:"Delete column",
    message:`Delete the column <b>${esc(COL_LABEL(col))}</b>?<br><br>`
           +`It currently holds data for <b>${filled}</b> item(s). That data is removed `
           +`from every item and from the CSV file. This cannot be undone.`,
    okLabel:"Delete column", cancelLabel:"Keep it", danger:true});
  if(!go) return;
  try{
    await api("/api/columns",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({action:"delete",name:col})});
    toast("Column deleted"); await loadState();
  }catch(e){ toast(e.error||"Failed","err"); }
}

function openColMenu(x,y,col,cols){
  closeRowMenu();
  const order=cols.filter(c=>c!=="name");
  const i=order.indexOf(col);
  const entries=[
    ["✏️ Rename…",       ()=>renameColumn(col),        true],
    ["◀ Move left",      ()=>moveColumn(col,-1),       i>0],
    ["▶ Move right",     ()=>moveColumn(col,+1),       i>=0 && i<order.length-1],
    ["🚫 Hide column",   ()=>hideColumn(col),          col!=="name"],
    ["🗑 Delete column…",()=>deleteColumn(col),        isCustomCol(col)],
    ["⚙ All columns…",   ()=>openColumnsModal(),       true],
  ].filter(e=>e[2]);

  const menu=document.createElement("div");
  menu.id="rowMenu"; menu.className="ctx-menu";
  menu.innerHTML=`<div class="ctx-head">Column: ${esc(COL_LABEL(col))}`
    + (isCustomCol(col)?"":" (built-in)") + `</div>`
    + entries.map((e,k)=>`<button data-i="${k}">${esc(e[0])}</button>`).join("");
  document.body.appendChild(menu);
  const w=menu.offsetWidth,h=menu.offsetHeight;
  menu.style.left=Math.max(6,Math.min(x,window.innerWidth-w-8))+"px";
  menu.style.top =Math.max(6,Math.min(y,window.innerHeight-h-8))+"px";
  menu.querySelectorAll("button").forEach(b=>b.onclick=()=>{
    closeRowMenu(); entries[+b.dataset.i][1]();
  });
}

async function openColumnsModal(){
  const CS=STATE.columns||{labels:{},hidden:[],order:[]};
  const hidden=CS.hidden||[];
  const known=[...LAST_COLS, ...ALL_COLUMNS.filter(c=>!LAST_COLS.includes(c))]
                .filter((c,i,a)=>a.indexOf(c)===i);
  const body=`<p class="muted">Drag-free ordering: use the arrows. Unticking hides a column
    from the table without touching the data.</p>
    <div id="colList"></div>`;
  const m=modalShell("Columns", body, `<button class="btn primary" onclick="closeModal()">Done</button>`);
  m.style.maxWidth="560px"; showModal(m);

  const draw=()=>{
    m.querySelector("#colList").innerHTML = known.map((c,i)=>`
      <div class="upd-row" style="grid-template-columns:26px 1fr auto;align-items:center">
        <input type="checkbox" class="colvis" data-c="${esc(c)}" ${hidden.includes(c)?"":"checked"}
               ${c==="name"?"disabled":""}>
        <span>${esc(COL_LABEL(c))}${isCustomCol(c)?"":' <span class="muted">(built-in)</span>'}</span>
        <span style="white-space:nowrap">
          <button class="icon-btn" data-mv="-1" data-c="${esc(c)}" ${c==="name"?"disabled":""}>◀</button>
          <button class="icon-btn" data-mv="1"  data-c="${esc(c)}" ${c==="name"?"disabled":""}>▶</button>
          <button class="icon-btn" data-ren="${esc(c)}" title="Rename">✏️</button>
          ${isCustomCol(c)?`<button class="icon-btn" data-del="${esc(c)}" title="Delete column">🗑</button>`:""}
        </span></div>`).join("");
    m.querySelectorAll(".colvis").forEach(cb=>cb.onchange=async()=>{
      const c=cb.dataset.c;
      const h=[...((STATE.columns||{}).hidden||[])].filter(x=>x!==c);
      if(!cb.checked) h.push(c);
      await saveColSettings({hidden:h});
      closeModal(); openColumnsModal();
    });
    m.querySelectorAll("[data-mv]").forEach(b=>b.onclick=async()=>{
      await moveColumn(b.dataset.c, +b.dataset.mv); closeModal(); openColumnsModal();
    });
    m.querySelectorAll("[data-ren]").forEach(b=>b.onclick=async()=>{
      await renameColumn(b.dataset.ren); closeModal(); openColumnsModal();
    });
    m.querySelectorAll("[data-del]").forEach(b=>b.onclick=async()=>{
      await deleteColumn(b.dataset.del); closeModal(); openColumnsModal();
    });
  };
  draw();
}

// ---------- right-click menu ----------
function closeRowMenu(){ const m=$("#rowMenu"); if(m) m.remove(); }
function openRowMenu(x, y, it){
  closeRowMenu();
  const n=SELECTED_IDS.size;
  const used=(it && (it.status||"")==="used");
  const entries=[
    ["✎ Edit…",            ()=>editItem(it.id),                 n===1],
    ["✏️ Rename…",          ()=>renameSelected(),                n===1],
    ["⇄ Move to category…", ()=>bulkMove(),                      true],
    ["🎨 Set colour…",      ()=>bulkColor(),                     true],
    [used?"↩ Return to stock":"✔ Mark as used",
                            ()=>bulkUsed(used),                  true],
    ["🛒 Add to order list", ()=>{ (LAST_ITEMS||[]).filter(i=>SELECTED_IDS.has(i.id))
                                    .forEach(i=>addToOrder(i.name)); }, true],
    ["🗑 Delete…",          ()=>bulkDelete(),                    true],
  ].filter(e=>e[2]);

  const menu=document.createElement("div");
  menu.id="rowMenu"; menu.className="ctx-menu";
  menu.innerHTML = `<div class="ctx-head">${n>1?`${n} items selected`:esc(it?it.name:"")}</div>`
    + entries.map((e,i)=>`<button data-i="${i}">${esc(e[0])}</button>`).join("");
  document.body.appendChild(menu);
  const w=menu.offsetWidth, h=menu.offsetHeight;
  menu.style.left = Math.max(6, Math.min(x, window.innerWidth  - w - 8)) + "px";
  menu.style.top  = Math.max(6, Math.min(y, window.innerHeight - h - 8)) + "px";
  menu.querySelectorAll("button").forEach(b=>b.onclick=()=>{
    closeRowMenu(); entries[+b.dataset.i][1]();
  });
}
document.addEventListener("click", e=>{ if(!e.target.closest("#rowMenu")) closeRowMenu(); });
document.addEventListener("keydown", e=>{
  if(e.key==="Escape"){ closeRowMenu(); if(!$("#overlay").classList.contains("show")) clearSelection(); }
});

// ---------- items ----------
function debouncedLoad(){clearTimeout(SEARCH_TIMER);SEARCH_TIMER=setTimeout(loadItems,220);}

async function loadItems(){
  const cat=SELECTED, sub=$("#includeSub").checked?"1":"0";
  const search=encodeURIComponent($("#search").value.trim());
  const crumb = cat ? cat.split("/").map(esc).join(" › ") : "All items";
  $("#crumb").innerHTML = `<span class="crumb">${crumb}</span>`;
  const data = await api(`/api/items?category=${encodeURIComponent(cat)}&include_sub=${sub}&search=${search}`);
  renderTable(data.items, data.custom_fields);
}

function renderTable(items, customFields){
  const wrap=$("#tableWrap");
  if(!items.length){
    wrap.innerHTML=`<div class="empty-state"><h3>No items here</h3>
      <p>Add an item to this category, or pick another category.</p>
      <button class="btn primary" onclick="openItemModal()">+ Add item</button></div>`;
    return;
  }
  // columns that actually have data in this view
  const usedCustom = customFields.filter(f=>items.some(it=>(it[f]||"").trim()!==""));
  const anyUsed = items.some(it=>(it.status||"")==="used");
  let cols=["name","quantity","category_path",...usedCustom,
            ...(anyUsed?["status","used_note"]:[]),"date_added","last_modified"];

  // column settings: saved order first, then anything new, minus hidden ones
  const CS = (STATE.columns||{labels:{},hidden:[],order:[]});
  const ord=CS.order||[], hidden=CS.hidden||[];
  cols = [...ord.filter(c=>cols.includes(c)), ...cols.filter(c=>!ord.includes(c))];
  cols = cols.filter(c=>c==="name" || !hidden.includes(c));
  ALL_COLUMNS = [...new Set([...cols, ...usedCustom, "quantity","category_path",
                             "status","used_note","date_added","last_modified"])];
  CUSTOM_COLUMNS = customFields.slice();

  const defaultLabels={name:"Name",quantity:"Qty",category_path:"Category",date_added:"Added",
                last_modified:"Modified",status:"Status",used_note:"Used for"};
  const labels=Object.assign({}, defaultLabels, CS.labels||{});
  COL_LABEL = c => labels[c] || c;
  VISIBLE_IDS = items.map(it=>it.id);
  const dataCols = cols.filter(c=>c!=="name");     // name is pinned separately

  // the checkbox and the name share ONE pinned cell, so there is no seam for
  // horizontally scrolled content to show through
  let html="<table><thead><tr>"
    + `<th class="col-name"><label class="pickwrap"><input type="checkbox" id="selAll" title="Select all"></label>`
    + `<span class="colhead" data-col="name">${esc(labels["name"]||"Name")}</span></th>`;
  dataCols.forEach(c=>html+=`<th data-col="${esc(c)}" title="Right-click to rename, move, hide or delete this column">`
    + `<span class="colhead" data-col="${esc(c)}">${esc(labels[c]||c)}<span class="colmenu">⋯</span></span></th>`);
  html+=`<th class="col-act"></th></tr></thead><tbody>`;

  items.forEach(it=>{
    const used=(it.status||"")==="used";
    // colour the whole row, or show it as a dot beside the name
    const rowStyle=(FILL_ROWS&&it.color)
      ? `background:${esc(it.color)};color:${readableOn(it.color)}`
      : "";
    const picked=SELECTED_IDS.has(it.id);
    const dot=it.color
      ? `<span class="dot" style="background:${esc(it.color)}"></span>`
      : `<span class="dot" style="background:transparent"></span>`;
    html+=`<tr data-id="${esc(it.id)}" class="${used?'is-used ':''}${picked?'is-picked':''}" style="${rowStyle}">`;
    html+=`<td class="col-name" title="${esc(it.name||"")}">`
        + `<label class="pickwrap"><input type="checkbox" class="rowpick" data-id="${esc(it.id)}" ${picked?"checked":""}></label>`
        + `${dot}${esc(it.name||"")}</td>`;
    dataCols.forEach(c=>{
      let v=it[c]||"";
      if(c==="status") v = used?"Used":"In stock";
      html+=`<td title="${esc(v)}">${esc(v)}</td>`;
    });
    html+=`<td class="col-act" style="white-space:nowrap">
           <button class="icon-btn" title="Add to order list" data-act="order" data-id="${esc(it.id)}">🛒</button>
           <button class="icon-btn" title="${used?'Return to stock':'Mark as used'}" data-act="used" data-id="${esc(it.id)}">${used?"↩":"✔"}</button>
           <button class="icon-btn" title="More… (or right-click the row)" data-act="menu" data-id="${esc(it.id)}">⋯</button></td>`;
    html+="</tr>";
  });
  html+="</tbody></table>";
  wrap.innerHTML=html;

  // ---- wiring ----
  const byId = id => items.find(x=>x.id===id);

  wrap.querySelectorAll("tbody tr").forEach(tr=>{
    const id=tr.dataset.id;
    tr.onclick=e=>{
      if(e.target.closest("button") || e.target.closest("input")) return;
      if(e.shiftKey){ selectRange(id); return; }
      if(e.metaKey||e.ctrlKey){ toggleSelect(id); return; }
      editItem(id);
    };
    tr.oncontextmenu=e=>{
      e.preventDefault();
      if(!SELECTED_IDS.has(id)){ SELECTED_IDS.clear(); SELECTED_IDS.add(id); paintSelection(); }
      openRowMenu(e.clientX, e.clientY, byId(id));
    };
  });

  wrap.querySelectorAll(".rowpick").forEach(cb=>{
    cb.onclick=e=>{ e.stopPropagation(); toggleSelect(cb.dataset.id, cb.checked); };
  });

  wrap.querySelectorAll("[data-act]").forEach(b=>{
    b.onclick=e=>{
      e.stopPropagation();
      const it=byId(b.dataset.id);
      if(b.dataset.act==="order") addToOrder(it.name);
      else if(b.dataset.act==="used") toggleUsed(it.id,(it.status||"")==="used");
      else {
        const r=b.getBoundingClientRect();
        if(!SELECTED_IDS.has(it.id)){ SELECTED_IDS.clear(); SELECTED_IDS.add(it.id); paintSelection(); }
        openRowMenu(r.left-150, r.bottom+4, it);
      }
    };
  });

  const all=wrap.querySelector("#selAll");
  all.checked = VISIBLE_IDS.length>0 && VISIBLE_IDS.every(id=>SELECTED_IDS.has(id));
  all.onclick=e=>{
    e.stopPropagation();
    if(all.checked) VISIBLE_IDS.forEach(id=>SELECTED_IDS.add(id));
    else VISIBLE_IDS.forEach(id=>SELECTED_IDS.delete(id));
    renderTable(items, customFields);
  };

  // column headers: right-click, or click the ⋯, to manage the column
  wrap.querySelectorAll("th [data-col]").forEach(sp=>{
    const col=sp.dataset.col;
    sp.oncontextmenu=e=>{ e.preventDefault(); e.stopPropagation();
                          openColMenu(e.clientX, e.clientY, col, cols); };
    const dots=sp.querySelector(".colmenu");
    if(dots) dots.onclick=e=>{ e.stopPropagation();
      const r=dots.getBoundingClientRect(); openColMenu(r.left-120, r.bottom+4, col, cols); };
  });
  wrap.querySelectorAll("th").forEach(th=>{
    if(th.dataset.col) th.oncontextmenu=e=>{ e.preventDefault();
      openColMenu(e.clientX, e.clientY, th.dataset.col, cols); };
  });

  LAST_ITEMS=items; LAST_CUSTOM=customFields; LAST_COLS=cols;
  paintSelection();
}

// ---------- modal infra ----------
// Modals can stack: a small ask-dialog opened on top of a bigger modal
// restores it (the same DOM node, so its handlers survive) when it closes.
let MODAL_STACK=[];
function showModal(node, stack){
  const o=$("#overlay");
  if(stack && o.classList.contains("show") && o.firstChild) MODAL_STACK.push(o.firstChild);
  o.innerHTML=""; o.appendChild(node); o.classList.add("show");
}
function closeModal(){
  const o=$("#overlay");
  if(MODAL_STACK.length){ o.innerHTML=""; o.appendChild(MODAL_STACK.pop()); return; }
  o.classList.remove("show"); o.innerHTML="";
}

// ---------- in-app replacements for prompt() / confirm() ----------
// Browsers can block or ignore the native dialogs entirely, which silently
// cancelled every action that relied on them. These always work.
function askText(opts){
  return new Promise(resolve=>{
    const o=opts||{};
    const field = o.options
      ? `<select id="askInput">${o.options.map(v=>
            `<option value="${esc(v)}" ${v===o.value?"selected":""}>${esc(v)}</option>`).join("")}</select>`
      : `<input id="askInput" value="${esc(o.value||"")}" placeholder="${esc(o.placeholder||"")}">`;
    const body=`${o.message?`<p class="muted">${o.message}</p>`:""}
      <div class="field"><label>${esc(o.label||"")}</label>${field}</div>`;
    const m=modalShell(o.title||"", body,
      `<button class="btn ghost" id="askCancel">Cancel</button>
       <button class="btn primary" id="askOk">${esc(o.okLabel||"OK")}</button>`);
    let done=false;
    const finish=v=>{ if(done)return; done=true; closeModal(); resolve(v); };
    showModal(m, true);
    const input=m.querySelector("#askInput");
    input.focus(); if(input.select) input.select();
    input.onkeydown=e=>{
      if(e.key==="Enter"){ e.preventDefault(); finish(input.value); }
      if(e.key==="Escape"){ e.preventDefault(); finish(null); }
    };
    m.querySelector("#askOk").onclick=()=>finish(input.value);
    m.querySelector("#askCancel").onclick=()=>finish(null);
    m.querySelector(".m-head .icon-btn").onclick=()=>finish(null);
  });
}
function askConfirm(opts){
  return new Promise(resolve=>{
    const o=opts||{};
    const m=modalShell(o.title||"Please confirm", `<p>${o.message||""}</p>`,
      `<button class="btn ghost" id="askNo">${esc(o.cancelLabel||"Cancel")}</button>
       <button class="btn ${o.danger?"danger":"primary"}" id="askYes">${esc(o.okLabel||"OK")}</button>`);
    let done=false;
    const finish=v=>{ if(done)return; done=true; closeModal(); resolve(v); };
    showModal(m, true);
    m.querySelector("#askYes").onclick=()=>finish(true);
    m.querySelector("#askNo").onclick=()=>finish(false);
    m.querySelector(".m-head .icon-btn").onclick=()=>finish(false);
    m.querySelector("#askYes").focus();
  });
}
$("#overlay").addEventListener("click",e=>{if(e.target===$("#overlay"))closeModal();});

function modalShell(title, bodyHtml, footHtml){
  const m=document.createElement("div"); m.className="modal";
  m.innerHTML=`<div class="m-head"><h3>${esc(title)}</h3><button class="icon-btn" onclick="closeModal()">✕</button></div>
    <div class="m-body">${bodyHtml}</div><div class="m-foot">${footHtml}</div>`;
  return m;
}

// ---------- categories ----------
async function addCategory(parent){
  const name=await askText({
    title: parent?"New subcategory":"New category",
    message: parent?`It will be created under <b>${esc(parent)}</b>.`:"",
    label:"Category name", placeholder:"e.g. Electronics", okLabel:"Create"});
  if(!name || !name.trim()) return;
  try{ await api("/api/category",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({parent,name:name.trim()})});
    if(parent)EXPANDED.add(parent);
    toast("Category added"); loadState();
  }catch(e){toast(e.error||"Failed","err");}
}
async function renameCategory(node){
  const nn=await askText({title:"Rename category", label:"New name",
                          value:node.name, okLabel:"Rename"});
  if(!nn||!nn.trim()||nn===node.name) return;
  try{ await api("/api/category/rename",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({path:node.path,new_name:nn.trim()})});
    toast("Renamed"); loadState();
  }catch(e){toast(e.error||"Failed","err");}
}
async function deleteCategory(node){
  try{
    await api("/api/category/delete",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({path:node.path,mode:"block"})});
    toast("Category deleted"); if(SELECTED===node.path)SELECTED=""; loadState();
  }catch(e){
    if(e.error==="not_empty"){
      const go=await askConfirm({
        title:"Category is not empty",
        message:`<b>${esc(node.name)}</b> and its subcategories contain <b>${e.count}</b> item(s).`
               +`<br><br>Deleting the category will delete those items too. They are archived in `
               +`deleted_items.csv first.`,
        okLabel:`Delete category and ${e.count} item(s)`, cancelLabel:"Keep everything", danger:true});
      if(!go) return;
      try{
        await api("/api/category/delete",{method:"POST",headers:{"Content-Type":"application/json"},
          body:JSON.stringify({path:node.path,mode:"recursive"})});
        toast("Category and items deleted"); if(SELECTED===node.path)SELECTED=""; loadState();
      }catch(e2){toast(e2.error||"Failed","err");}
    }else{toast(e.error||"Failed","err");}
  }
}

function openRequiredModal(node){
  let fields=[...(node.required_fields||[])];
  const body=`<p class="muted">New items created in <b>${esc(node.name)}</b> (and its subcategories) will automatically get these fields.</p>
    <div id="reqChips" style="margin:12px 0;"></div>
    <div class="custom-row">
      <div class="field" style="margin:0"><label>Add a field name</label>
        <input id="reqInput" placeholder="e.g. serial_number" onkeydown="if(event.key==='Enter'){event.preventDefault();document.getElementById('reqAddBtn').click();}"></div>
      <button class="btn" id="reqAddBtn">Add</button>
    </div>`;
  const foot=`<button class="btn ghost" onclick="closeModal()">Cancel</button>
    <button class="btn primary" id="reqSave">Save</button>`;
  const m=modalShell(`Required fields · ${node.name}`, body, foot);
  showModal(m);
  const renderChips=()=>{
    m.querySelector("#reqChips").innerHTML = fields.length
      ? fields.map((f,i)=>`<span class="chip">${esc(f)}<button data-i="${i}">✕</button></span>`).join("")
      : `<span class="muted">No required fields yet.</span>`;
    m.querySelectorAll("#reqChips .chip button").forEach(b=>b.onclick=()=>{fields.splice(+b.dataset.i,1);renderChips();});
  };
  renderChips();
  m.querySelector("#reqAddBtn").onclick=()=>{
    const v=m.querySelector("#reqInput").value.trim();
    if(v&&!fields.includes(v)){fields.push(v);renderChips();}
    m.querySelector("#reqInput").value="";m.querySelector("#reqInput").focus();
  };
  m.querySelector("#reqSave").onclick=async()=>{
    try{await api("/api/category/required",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({path:node.path,fields})});
      toast("Saved"); closeModal(); loadState();
    }catch(e){toast(e.error||"Failed","err");}
  };
}

// ---------- items ----------
async function categoryOptions(selected){
  // flatten tree paths
  const paths=[];
  (function walk(nodes){nodes.forEach(n=>{paths.push(n.path);walk(n.children);});})(STATE.tree);
  return paths.map(p=>`<option value="${esc(p)}" ${p===selected?"selected":""}>${esc(p)}</option>`).join("");
}

async function openItemModal(){
  await buildItemModal(null);
}
async function editItem(id){
  const data=await api(`/api/items?category=&include_sub=1`);
  const it=data.items.find(x=>x.id===id);
  if(!it){toast("Not found","err");return;}
  await buildItemModal(it);
}

async function buildItemModal(item){
  const isEdit=!!item;
  const startCat = item ? item.category_path : SELECTED;
  const opts = await categoryOptions(startCat);
  let customPairs=[];   // {key,value}
  if(item){
    STATE.custom_fields.forEach(f=>{ if((item[f]||"").trim()!=="") customPairs.push({key:f,value:item[f]}); });
  }

  const body=`
    <div class="field"><label>Name <span class="req">*</span></label>
      <input id="f_name" value="${item?esc(item.name):""}" placeholder="Item name"></div>
    <div class="field"><label>Quantity</label>
      <input id="f_qty" value="${item?esc(item.quantity):"1"}" type="text"></div>
    <div class="field"><label>Category</label>
      <select id="f_cat"><option value="">— none —</option>${opts}</select></div>
    <div class="field"><label>Colour</label>
      <div style="display:flex;gap:8px;align-items:center;">
        <input type="color" id="f_color" value="${item&&item.color?esc(item.color):"#4c9be8"}"
               style="width:52px;padding:2px;height:34px;">
        <button class="btn small" type="button" id="clearColor">No colour</button>
        <span class="muted" id="colorState">${item&&item.color?esc(item.color):"none"}</span>
      </div></div>
    <div class="field"><label>Status</label>
      <select id="f_status">
        <option value="">In stock</option>
        <option value="used" ${item&&item.status==="used"?"selected":""}>Used</option>
      </select></div>
    <div class="field" id="usedWrap" style="display:${item&&item.status==="used"?"block":"none"}">
      <label>Used for / where</label>
      <input id="f_usednote" value="${item?esc(item.used_note||""):""}" placeholder="e.g. installed in Lab 2 / given to Ahmed">
      ${item&&item.used_date?`<div class="muted" style="margin-top:4px">Marked used: ${esc(item.used_date)}</div>`:""}
    </div>
    <hr style="border:none;border-top:1px solid var(--line);margin:16px 0;">
    <div class="muted" style="margin-bottom:8px;">Additional info — add any fields this item needs</div>
    <div id="customFields"></div>
    <div class="custom-row">
      <div class="field" style="margin:0;flex:1.3"><label>Field name</label><input id="newFieldKey" placeholder="e.g. serial_number"></div>
      <div class="field" style="margin:0;max-width:110px"><label>Type</label>
        <select id="newFieldType">
          <option value="text">Text</option>
          <option value="number">Number</option>
          <option value="date">Date</option>
          <option value="notes">Notes</option>
        </select></div>
      <div class="field" style="margin:0;flex:1.3"><label>Value</label><input id="newFieldVal" placeholder="value"></div>
      <button class="btn" id="addFieldBtn">Add</button>
    </div>`;
  const foot=`<button class="btn ghost" onclick="closeModal()">Cancel</button>
    <button class="btn primary" id="saveItem">${isEdit?"Save changes":"Add item"}</button>`;
  const m=modalShell(isEdit?"Edit item":"New item", body, foot);
  showModal(m);

  const renderCustom=()=>{
    const c=m.querySelector("#customFields");
    c.innerHTML = customPairs.map((p,i)=>{
      const val=p.value||"";
      const multiline = /\n/.test(val) || val.length>60;
      const control = multiline
        ? `<textarea data-i="${i}" class="cv" rows="3">${esc(val)}</textarea>`
        : `<input data-i="${i}" class="cv" value="${esc(val)}">`;
      return `<div class="custom-row">
        <div class="field" style="margin:0"><label>${esc(p.key)}</label>${control}</div>
        <button class="btn danger small" data-del="${i}" title="Remove field">✕</button>
      </div>`;
    }).join("");
    c.querySelectorAll(".cv").forEach(inp=>inp.oninput=()=>customPairs[+inp.dataset.i].value=inp.value);
    c.querySelectorAll("[data-del]").forEach(b=>b.onclick=()=>{customPairs.splice(+b.dataset.del,1);renderCustom();});
  };
  renderCustom();

  // swap the "value" input widget to match the chosen field type
  m.querySelector("#newFieldType").onchange=e=>{
    const old=m.querySelector("#newFieldVal");
    const t=e.target.value;
    let el;
    if(t==="notes"){ el=document.createElement("textarea"); el.rows=3; }
    else { el=document.createElement("input"); el.type = (t==="number")?"number":(t==="date")?"date":"text"; }
    el.id="newFieldVal"; el.placeholder="value";
    old.parentElement.replaceChild(el, old);
  };

  m.querySelector("#addFieldBtn").onclick=()=>{
    const k=m.querySelector("#newFieldKey").value.trim();
    const v=m.querySelector("#newFieldVal").value;
    if(!k)return;
    if(customPairs.some(p=>p.key===k)){toast("Field already added","err");return;}
    customPairs.push({key:k,value:v});
    m.querySelector("#newFieldKey").value="";m.querySelector("#newFieldVal").value="";
    renderCustom();
  };

  // auto-apply required fields when category changes
  const applyRequired=async(cat)=>{
    if(!cat)return;
    const r=await api(`/api/required_for?category=${encodeURIComponent(cat)}`);
    (r.required_fields||[]).forEach(f=>{
      if(!customPairs.some(p=>p.key===f)) customPairs.push({key:f,value:""});
    });
    renderCustom();
  };
  m.querySelector("#f_cat").onchange=e=>applyRequired(e.target.value);
  if(!isEdit && startCat) applyRequired(startCat);

  // colour picker: remembers "no colour" as a separate state from a chosen one
  let chosenColor = (item && item.color) ? item.color : "";
  const colorInput=m.querySelector("#f_color"), colorState=m.querySelector("#colorState");
  colorInput.oninput=()=>{ chosenColor=colorInput.value; colorState.textContent=chosenColor; };
  m.querySelector("#clearColor").onclick=()=>{ chosenColor=""; colorState.textContent="none"; };

  // show the "used for" box only when the status is Used
  const statusSel=m.querySelector("#f_status");
  statusSel.onchange=()=>{
    m.querySelector("#usedWrap").style.display = statusSel.value==="used" ? "block" : "none";
  };

  m.querySelector("#saveItem").onclick=async()=>{
    const status=statusSel.value;
    const usedNote=m.querySelector("#f_usednote").value.trim();
    const fields={
      name:m.querySelector("#f_name").value.trim(),
      quantity:m.querySelector("#f_qty").value.trim(),
      category_path:m.querySelector("#f_cat").value,
      color:chosenColor,
      status:status,
      used_note:status==="used"?usedNote:"",
      used_date:status==="used"
        ? ((item&&item.used_date)?item.used_date:new Date().toISOString().slice(0,19).replace("T"," "))
        : ""
    };
    if(!fields.name){toast("Name is required","err");return;}
    customPairs.forEach(p=>{ if(p.key.trim()) fields[p.key.trim()]=p.value; });
    try{
      await api("/api/item",{method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({id:item?item.id:"",fields})});
      toast(isEdit?"Saved":"Item added"); closeModal(); loadState();
    }catch(e){toast(e.error||"Failed","err");}
  };
}

async function deleteItem(id,name){
  const reason=await askText({
    title:"Delete item",
    message:`<b>${esc(name||"this item")}</b> will be archived in deleted_items.csv before it is removed.`,
    label:"Reason (optional)", placeholder:"e.g. broken, returned, sold",
    okLabel:"Delete"});
  if(reason===null) return;      // cancelled
  try{await api("/api/item/delete",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({id,reason})});
    toast("Deleted and archived"); loadState();
  }catch(e){toast(e.error||"Failed","err");}
}

async function toggleUsed(id,isUsed){
  let note="";
  if(!isUsed){
    note=await askText({title:"Mark as used", label:"Used for / where",
      placeholder:"e.g. installed in Lab 2 / given to Ahmed", okLabel:"Mark used"});
    if(note===null) return;
  }else if(!await askConfirm({title:"Back in stock",
      message:"Put this item back in stock?", okLabel:"Yes"})) return;
  try{
    await api("/api/item/status",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({id,status:isUsed?"":"used",used_note:note})});
    toast(isUsed?"Back in stock":"Marked as used"); loadState();
  }catch(e){toast(e.error||"Failed","err");}
}

// ---------- order list ----------
let ORDER_ROWS = [];

// Reads numbers the way people type them: 12.50, 12,50, 1.234,56, "€ 9,99".
// Used for BOTH the line totals and the grand total so they can never disagree.
function parseNum(v){
  let s=String(v==null?"":v).replace(/[^\d,.\-]/g,"");
  if(!s) return 0;
  if(s.includes(",")&&s.includes(".")) s = s.lastIndexOf(",")>s.lastIndexOf(".")
    ? s.replace(/\./g,"").replace(",",".") : s.replace(/,/g,"");
  else if(s.includes(",")) s=s.replace(",",".");
  const f=parseFloat(s); return isNaN(f)?0:f;
}
function lineTotal(r){ return (parseNum(r.quantity)||0) * parseNum(r.unit_price); }
function orderTotal(){ return ORDER_ROWS.reduce((t,r)=>t + lineTotal(r), 0); }
function fmtMoney(n){ return n.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}); }

async function openOrderModal(){
  const data = await api("/api/order");
  ORDER_ROWS = (data.rows||[]).map(r=>({item:r.item,quantity:r.quantity,
                                        unit_price:r.unit_price,link:r.link,note:r.note}));
  const cur = data.currency || "";
  const body = `
    <p class="muted">A classic order list. Add what you need to buy, then export it as CSV or PDF.</p>
    <div id="orderTable"></div>
    <div style="display:flex;gap:8px;margin-top:10px;flex-wrap:wrap">
      <button class="btn small" id="ordAdd">+ Add line</button>
      <button class="btn small" id="ordFromInv">+ From inventory</button>
    </div>
    <div id="ordTotal" style="text-align:right;margin-top:14px;font-size:15px;"></div>
    <hr style="border:none;border-top:1px solid var(--line);margin:14px 0;">
    <div class="field"><label>Title on the exported document</label>
      <input id="ordTitle" value="Order list" placeholder="Order list"></div>
    <div class="field"><label>Note / reference (optional)</label>
      <input id="ordNote" placeholder="e.g. Dept. purchase request #12"></div>`;
  const foot = `<button class="btn ghost" onclick="closeModal()">Close</button>
    <button class="btn" id="ordSave">Save</button>
    <button class="btn" id="ordCsv">Export CSV</button>
    <button class="btn primary" id="ordPdf">Export PDF</button>`;
  const m = modalShell("Order list", body, foot);
  m.style.maxWidth = "820px";
  showModal(m);

  const render = () => {
    const t = m.querySelector("#orderTable");
    if(!ORDER_ROWS.length){
      t.innerHTML = `<p class="muted" style="padding:12px 0">Nothing on the list yet.</p>`;
    } else {
      t.innerHTML = `<table style="width:100%;font-size:13px">
        <thead><tr>
          <th style="width:26px">#</th><th>Item</th><th style="width:70px">Qty</th>
          <th style="width:96px">Unit ${esc(cur)}</th><th style="width:92px">Total</th>
          <th style="width:150px">Link</th><th style="width:26px"></th>
        </tr></thead><tbody>
        ${ORDER_ROWS.map((r,i)=>`<tr style="cursor:default">
          <td>${i+1}</td>
          <td><input class="oi" data-f="item" data-i="${i}" value="${esc(r.item||"")}" style="width:100%"></td>
          <td><input class="oi" data-f="quantity" data-i="${i}" value="${esc(r.quantity||"")}" style="width:100%"></td>
          <td><input class="oi" data-f="unit_price" data-i="${i}" value="${esc(r.unit_price||"")}" style="width:100%" placeholder="0.00"></td>
          <td class="lt" style="text-align:right;font-variant-numeric:tabular-nums"></td>
          <td><input class="oi" data-f="link" data-i="${i}" value="${esc(r.link||"")}" style="width:100%" placeholder="https://…"></td>
          <td><button class="btn danger small" data-del="${i}" title="Remove">✕</button></td>
        </tr>`).join("")}</tbody></table>`;
    }
    // inputs are styled inline to stay compact inside the table
    t.querySelectorAll("input.oi").forEach(inp=>{
      inp.style.background="var(--bg)"; inp.style.border="1px solid var(--line)";
      inp.style.color="var(--text)"; inp.style.borderRadius="5px"; inp.style.padding="5px 7px";
      inp.style.fontFamily="inherit"; inp.style.fontSize="13px";
      inp.oninput=()=>{ ORDER_ROWS[+inp.dataset.i][inp.dataset.f]=inp.value; paintTotals(); };
    });
    t.querySelectorAll("[data-del]").forEach(b=>b.onclick=()=>{
      ORDER_ROWS.splice(+b.dataset.del,1); render();
    });
    paintTotals();
  };

  const paintTotals = () => {
    m.querySelectorAll("#orderTable tbody tr").forEach((tr,i)=>{
      const r=ORDER_ROWS[i]; const cell=tr.querySelector(".lt");
      if(cell) cell.textContent = fmtMoney(lineTotal(r));
    });
    m.querySelector("#ordTotal").innerHTML =
      `<span class="muted">Total </span><b>${esc(cur)} ${fmtMoney(orderTotal())}</b>`;
  };

  render();

  m.querySelector("#ordAdd").onclick=()=>{
    ORDER_ROWS.push({item:"",quantity:"1",unit_price:"",link:"",note:""}); render();
    const ins=m.querySelectorAll('input[data-f="item"]'); if(ins.length) ins[ins.length-1].focus();
  };

  m.querySelector("#ordFromInv").onclick=async()=>{
    const d=await api("/api/items?category=&include_sub=1");
    const names=(d.items||[]).map(x=>x.name).filter(Boolean).sort();
    if(!names.length){ toast("No items in the inventory yet","err"); return; }
    const pick=await askText({title:"Add from inventory", label:"Item",
                              options:names, value:names[0], okLabel:"Add"});
    if(!pick) return;
    ORDER_ROWS.push({item:pick.trim(),quantity:"1",unit_price:"",link:"",note:""}); render();
  };

  const save = async () => {
    const r = await api("/api/order",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({rows:ORDER_ROWS})});
    return r;
  };
  m.querySelector("#ordSave").onclick=async()=>{
    try{ await save(); toast("Order list saved"); }catch(e){ toast(e.error||"Failed","err"); }
  };

  const doExport = async (fmt, btn) => {
    const label=btn.textContent;
    btn.disabled=true; btn.textContent="Choose a location…";
    try{
      await save();                                  // always export what's on screen
      const r = await api("/api/order/export",{method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({format:fmt,
          title:m.querySelector("#ordTitle").value,
          note:m.querySelector("#ordNote").value})});
      if(r.cancelled){ btn.disabled=false; btn.textContent=label; return; }
      toast(`${fmt.toUpperCase()} saved — ${r.count} line(s), total ${r.total_display}`);
    }catch(e){ toast(e.error||"Export failed","err"); }
    btn.disabled=false; btn.textContent=label;
  };
  m.querySelector("#ordCsv").onclick=e=>doExport("csv", e.target);
  m.querySelector("#ordPdf").onclick=e=>doExport("pdf", e.target);
}

async function addToOrder(name){
  try{
    const r=await api("/api/order/add",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({item:name,quantity:"1"})});
    toast(`"${name}" added to the order list (${r.count} line(s))`);
  }catch(e){ toast(e.error||"Failed","err"); }
}

// ---------- export ----------
function openExportModal(){
  const cat=SELECTED;
  const body=`<p class="muted">Writes a clean CSV with just your data — no internal ids or program columns.</p>
    <div class="field"><label>What to export</label>
      <select id="expScope">
        <option value="all">Everything (${STATE.item_count} items)</option>
        ${cat?`<option value="cat" selected>Only "${esc(cat)}" and its subcategories</option>`:""}
      </select></div>
    <label class="check" style="margin:8px 0"><input type="checkbox" id="expDates" checked> include date columns</label>
    <label class="check" style="margin:8px 0"><input type="checkbox" id="expUsed" checked> include status / used-for columns</label>
    <label class="check" style="margin:8px 0"><input type="checkbox" id="expColor"> include colour column</label>`;
  const foot=`<button class="btn ghost" onclick="closeModal()">Cancel</button>
    <button class="btn primary" id="expGo">Choose file &amp; export</button>`;
  const m=modalShell("Export CSV", body, foot); showModal(m);
  m.querySelector("#expGo").onclick=async()=>{
    const btn=m.querySelector("#expGo"); btn.disabled=true; btn.textContent="Choose a location…";
    try{
      const r=await api("/api/export",{method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({
          category:m.querySelector("#expScope").value==="cat"?cat:"",
          include_sub:true,
          include_dates:m.querySelector("#expDates").checked,
          include_used:m.querySelector("#expUsed").checked,
          include_color:m.querySelector("#expColor").checked})});
      if(r.cancelled){ btn.disabled=false; btn.textContent="Choose file & export"; return; }
      toast(`Exported ${r.count} item(s)`); closeModal();
    }catch(e){ toast(e.error||"Export failed","err"); btn.disabled=false; btn.textContent="Choose file & export"; }
  };
}

// ---------- deleted items ----------
async function openDeleted(){
  const data=await api("/api/deleted");
  const rows=data.deleted||[];
  const body = rows.length
    ? `<p class="muted">Archived in ${esc(data.file)}</p>
       <div style="margin-top:10px">${rows.map(d=>`<div class="hist-item">
        <span class="ts">${esc(d.deleted_at)}</span>
        <span class="act">${esc(d.category_path||"—")}</span>
        <span><b>${esc(d.name)}</b> ×${esc(d.quantity)}
        ${d.deleted_reason?` — <span class="muted">${esc(d.deleted_reason)}</span>`:""}
        ${d.other_fields?`<br><span class="muted">${esc(d.other_fields)}</span>`:""}</span></div>`).join("")}</div>`
    : `<p class="muted">Nothing has been deleted yet.</p>`;
  const m=modalShell("Deleted items", body, `<button class="btn primary" onclick="closeModal()">Close</button>`);
  m.style.maxWidth="760px"; showModal(m);
}

async function moveItem(id,current){
  const opts=await categoryOptions(current);
  const body=`<div class="field"><label>Move item to category</label>
    <select id="moveCat"><option value="">— none —</option>${opts}</select></div>`;
  const foot=`<button class="btn ghost" onclick="closeModal()">Cancel</button>
    <button class="btn primary" id="doMove">Move</button>`;
  const m=modalShell("Move item", body, foot); showModal(m);
  m.querySelector("#doMove").onclick=async()=>{
    const cat=m.querySelector("#moveCat").value;
    try{await api("/api/item/move",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({id,category_path:cat})});
      toast("Moved"); closeModal(); loadState();
    }catch(e){toast(e.error||"Failed","err");}
  };
}

// ---------- updates ----------
let UPDATE_INFO = null;

function paintVersionBadge(){
  const b=$("#versionBadge"); if(!b) return;
  const v=STATE.version?("v"+STATE.version):"";
  if(UPDATE_INFO && UPDATE_INFO.update_available){
    b.textContent=`Update available → v${UPDATE_INFO.latest}`;
    b.className="version-badge has-update";
  }else{
    b.textContent=v; b.className="version-badge";
  }
}

async function checkUpdateQuietly(){
  try{
    UPDATE_INFO = await api("/api/update/check");
    paintVersionBadge();
  }catch(e){ /* offline is fine - stay silent */ }
}

async function openUpdateModal(){
  const body=`<div id="updBody"><p class="muted">Checking GitHub…</p></div>`;
  const foot=`<button class="btn ghost" onclick="closeModal()">Close</button>
    <button class="btn primary" id="updBtn" style="display:none">Update now</button>`;
  const m=modalShell("Software update", body, foot); showModal(m);

  let info=UPDATE_INFO;
  try{ info = await api("/api/update/check"); UPDATE_INFO=info; paintVersionBadge(); }
  catch(e){ info=null; }

  const el=m.querySelector("#updBody");
  if(!info || info.error){
    el.innerHTML=`<p>Could not reach GitHub.</p>
      <p class="muted">${esc((info&&info.error)||"Check your internet connection.")}</p>`;
    return;
  }
  const rows=(info.commits||[]).map(c=>`<div class="upd-row">
      <span class="sha">${esc(c.sha)}</span>
      <span>${esc(c.message)}<br><span class="muted">${esc(c.date)}</span></span></div>`).join("");

  el.innerHTML=`
    <div class="field"><label>Installed</label><b>v${esc(info.current)}</b></div>
    <div class="field"><label>Latest on GitHub</label>
      <b>${info.latest?("v"+esc(info.latest)):"unknown"}</b></div>
    <div class="field"><label>Source</label>
      <span class="muted">${esc(info.repo)} · ${esc(info.branch)}</span></div>
    ${info.update_available
      ? `<p>A newer version is available.</p>
         ${rows?`<div style="margin-top:10px"><div class="muted" style="margin-bottom:4px">Recent changes</div>${rows}</div>`:""}
         <p class="muted" style="margin-top:12px">Your inventory data is stored outside the program folder and is not touched.
         Any file that gets replaced is backed up first.</p>`
      : (info.latest
          ? `<p>You are up to date.</p>`
          : `<p>Could not read a version number from the published copy.</p>
             <p class="muted">The version on GitHub may predate this updater. Push the
             current code once and this will start reporting correctly.</p>`)}`;

  const btn=m.querySelector("#updBtn");
  if(info.update_available){
    btn.style.display="";
    btn.onclick=async()=>{
      btn.disabled=true; btn.textContent="Downloading…";
      try{
        const r=await api("/api/update/apply",{method:"POST"});
        const n=(r.changed||[]).length+(r.added||[]).length;
        el.innerHTML=`<p><b>${esc(r.message)}</b></p>
          ${n?`<p class="muted">${n} file(s) updated.</p>`:""}
          ${r.backup?`<p class="muted">Backup: ${esc(r.backup)}</p>`:""}
          <p>Restart the app to use the new version.</p>`;
        btn.textContent="Restart now"; btn.disabled=false;
        btn.onclick=async()=>{
          btn.disabled=true; btn.textContent="Restarting…";
          try{ await api("/api/update/restart",{method:"POST"}); }catch(e){}
          setTimeout(()=>location.reload(), 3000);
        };
      }catch(e){
        el.innerHTML=`<p>Update failed.</p><p class="muted">${esc(e.message||e.error||"Unknown error")}</p>`;
        btn.style.display="none";
      }
    };
  }
}

// ---------- import ----------
function openImportModal(){
  const body=`
    <div class="field"><label>One spreadsheet</label>
      <p class="muted" style="margin:0 0 8px">Each sheet becomes a subcategory, each row an item.</p>
      <button class="btn" id="impFile">Choose a file…</button></div>
    <hr style="border:none;border-top:1px solid var(--line);margin:16px 0;">
    <div class="field"><label>A whole folder</label>
      <p class="muted" style="margin:0 0 8px">Imports every .csv, .tsv and .xlsx in the folder —
      each file becomes its own category. The app's own data files are skipped automatically.</p>
      <label class="check" style="margin:6px 0"><input type="checkbox" id="impRec"> include subfolders (they become parent categories)</label>
      <div class="field" style="margin:8px 0 10px"><label>Put everything under one category (optional)</label>
        <input id="impUnder" placeholder="e.g. Imported 2026"></div>
      <button class="btn" id="impFolder">Choose a folder…</button></div>
    <div id="impResult" style="margin-top:14px"></div>`;
  const foot=`<button class="btn ghost" onclick="closeModal()">Close</button>
    <button class="btn primary" id="impConfirm" style="display:none">Import all</button>`;
  const m=modalShell("Import", body, foot);
  m.style.maxWidth="640px";
  showModal(m);

  const res=m.querySelector("#impResult");
  const confirmBtn=m.querySelector("#impConfirm");
  let pending=null;                  // folder options awaiting confirmation

  m.querySelector("#impFile").onclick=async()=>{
    res.innerHTML=`<p class="muted">Choose a file in the dialog…</p>`;
    try{
      const r=await api("/api/import_file",{method:"POST",headers:{"Content-Type":"application/json"},body:"{}"});
      if(r.cancelled){ res.innerHTML=""; return; }
      res.innerHTML=`<p><b>Imported ${r.imported} item(s)</b> from ${esc(r.file)}</p>
        <p class="muted">${(r.categories||[]).map(esc).join("<br>")}</p>`;
      toast(`Imported ${r.imported} item(s)`); loadState();
    }catch(e){ res.innerHTML=`<p class="muted">${esc(e.error||"Import failed")}</p>`; }
  };

  m.querySelector("#impFolder").onclick=async()=>{
    const opts={preview:true,
                recursive:m.querySelector("#impRec").checked,
                under:m.querySelector("#impUnder").value.trim()};
    res.innerHTML=`<p class="muted">Choose a folder in the dialog…</p>`;
    confirmBtn.style.display="none";
    try{
      const r=await api("/api/import_folder",{method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify(opts)});
      if(r.cancelled){ res.innerHTML=""; return; }
      if(!r.files.length){
        res.innerHTML=`<p>${esc(r.message||"Nothing to import in that folder.")}</p>`;
        return;
      }
      pending={...opts, preview:false, path:r.folder};
      res.innerHTML=`
        <p><b>${r.files.length} file(s) found</b> in ${esc(r.folder)} —
           ${r.added} item(s) would be imported.</p>
        <div style="max-height:220px;overflow:auto;margin-top:8px">
        ${r.files.map(f=>`<div class="upd-row">
            <span class="sha">${f.added}×</span>
            <span>${esc(f.file)}<br><span class="muted">${f.categories.map(esc).join(", ")}
            ${f.colored?` · ${f.colored} coloured`:""}</span></span></div>`).join("")}
        </div>
        ${r.failed&&r.failed.length?`<p class="muted" style="margin-top:8px">Skipped:
          ${r.failed.map(f=>esc(f.file)).join(", ")}</p>`:""}
        <p class="muted" style="margin-top:10px">Nothing has been written yet.</p>`;
      confirmBtn.style.display="";
    }catch(e){ res.innerHTML=`<p class="muted">${esc(e.error||"Import failed")}</p>`; }
  };

  confirmBtn.onclick=async()=>{
    if(!pending) return;
    confirmBtn.disabled=true; confirmBtn.textContent="Importing…";
    try{
      const r=await api("/api/import_folder",{method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify(pending)});
      res.innerHTML=`<p><b>Imported ${r.added} item(s)</b> from ${r.files.length} file(s).</p>`;
      toast(`Imported ${r.added} item(s) from ${r.files.length} file(s)`);
      confirmBtn.style.display="none"; pending=null; loadState();
    }catch(e){ res.innerHTML=`<p class="muted">${esc(e.error||"Import failed")}</p>`; }
    confirmBtn.disabled=false; confirmBtn.textContent="Import all";
  };
}

// ---------- history ----------
async function openHistory(){
  const data=await api("/api/history");
  const rows=data.history||[];
  const body = rows.length
    ? `<div>${rows.map(h=>`<div class="hist-item">
        <span class="ts">${esc(h.timestamp)}</span>
        <span class="act">${esc(h.action)}</span>
        <span>${esc(h.target)}${h.detail?` — <span class="muted">${esc(h.detail)}</span>`:""}</span></div>`).join("")}</div>`
    : `<p class="muted">No changes recorded yet.</p>`;
  const m=modalShell("Change history", body,
    `<button class="btn ghost" onclick="openDeleted()">View deleted items</button>
     <button class="btn primary" onclick="closeModal()">Close</button>`);
  m.style.maxWidth="720px";
  showModal(m);
}

// ---------- data dir ----------
function openDirModal(){
  const canBrowse = STATE.can_browse;
  const body=`<div class="field"><label>Folder for CSV data (e.g. your shared drive)</label>
    <div style="display:flex;gap:8px;align-items:center;">
      <input id="dirInput" style="flex:1;min-width:0" value="${esc(STATE.data_dir)}" placeholder="/path/to/folder or \\\\server\\share\\inventory">
      ${canBrowse?`<button class="btn" id="browseBtn" type="button" style="white-space:nowrap;">Browse…</button>`:``}
    </div></div>
    <p class="muted">Files stored here: inventory.csv, categories.csv, history.csv — all openable in Excel.
    Pick an empty folder and your current data is copied there automatically.</p>
    <label class="check" style="margin:10px 0"><input type="checkbox" id="dirShared" checked>
      remember this for <b>every computer</b> that runs this copy of the app</label>
    <p class="muted">Ticked, the setting is saved in <code>inventory_config.json</code> next to app.py —
    the right choice when the app lives on a shared drive, because every laptop then finds the same data.
    Unticked, it applies only to this computer.</p>
    <p class="muted">Currently taken from: ${esc((STATE.dir_source||{}).where||"—")}</p>
    ${STATE.dir_from_env?`<p class="muted" style="color:var(--accent)">Note: the folder is currently
    fixed by the INVENTORY_DIR environment variable, which overrides this setting for as long as it is set.</p>`:""}`;
  const foot=`<button class="btn ghost" onclick="closeModal()">Cancel</button>
    <button class="btn primary" id="saveDir">Use this folder</button>`;
  const m=modalShell("Data folder", body, foot); showModal(m);
  if(canBrowse){
    m.querySelector("#browseBtn").onclick=async()=>{
      const btn=m.querySelector("#browseBtn"), old=btn.textContent;
      btn.textContent="Opening…"; btn.disabled=true;
      try{
        const r=await api("/api/browse_dir",{method:"POST"});
        if(r.path) m.querySelector("#dirInput").value=r.path;
      }catch(e){ toast("Could not open folder picker","err"); }
      btn.textContent=old; btn.disabled=false;
    };
  }
  m.querySelector("#saveDir").onclick=async()=>{
    const path=m.querySelector("#dirInput").value.trim();
    const scope=m.querySelector("#dirShared").checked?"shared":"user";
    try{const r=await api("/api/set_dir",{method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({path,scope})});
      let msg = r.migrated ? "Data folder set · existing data copied" : "Data folder set";
      if(r.shared_failed) msg += " (couldn't write next to app.py — saved for this computer only)";
      else if(r.scope==="shared") msg += " · shared with every computer";
      toast(msg); closeModal(); loadState();
    }catch(e){toast(e.error||"Failed","err");}
  };
}

$("#fillRows").checked = FILL_ROWS;     // restore the colour-display preference
loadState();
setTimeout(checkUpdateQuietly, 1500);   // background check, silent if offline
</script>
</body>
</html>
"""


def _open_browser(url):
    """Open the default web browser a moment after the server starts."""
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()


if __name__ == "__main__":
    ensure_files()
    HOST = "127.0.0.1"
    # NOTE: not 5000 — on macOS the AirPlay Receiver (Control Center) permanently
    # occupies port 5000, so binding there fails with "address already in use".
    PORT = int(os.environ.get("INVENTORY_PORT", "8765"))
    url = f"http://{HOST}:{PORT}"

    print("=" * 60)
    print("  Department Inventory Manager")
    print("=" * 60)
    print(f"  Data folder : {get_data_dir()}")
    print(f"  Open in your browser: {url}")
    print("  (Press Ctrl+C to stop)")
    print("=" * 60)

    # Flask's dev server would spawn a second process with the reloader and
    # open the browser twice; the reloader is already off (debug=False).
    _open_browser(url)
    try:
        APP.run(host=HOST, port=PORT, debug=False)
    except OSError as e:
        fatal(
            f"Could not start on {url}: {e}\n\n"
            "The port may already be in use (is the app already running?).\n"
            "Set a different port with the INVENTORY_PORT environment variable."
        )
