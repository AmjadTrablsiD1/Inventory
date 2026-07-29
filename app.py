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
APP_VERSION = "1.3.0"

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

CONFIG_FILE = Path.home() / ".inventory_manager_config.json"

INVENTORY_CSV = "inventory.csv"
CATEGORIES_CSV = "categories.csv"
HISTORY_CSV = "history.csv"
DELETED_CSV = "deleted_items.csv"   # every deleted item is archived here
ORDER_CSV = "order_list.csv"        # the working "things to order" list

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


def get_data_dir():
    cfg = load_config()
    d = cfg.get("data_dir") or DEFAULT_DIR
    return Path(d)


def set_data_dir(path):
    cfg = load_config()
    cfg["data_dir"] = str(path)
    save_config(cfg)


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
            count = import_data.import_file(path)
        except SystemExit as e:
            return jsonify({"error": str(e)}), 400
        except Exception as e:
            return jsonify({"error": f"Import failed: {e}"}), 400
    return jsonify({"ok": True, "imported": count, "file": Path(path).name})


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

        set_data_dir(p)
        ensure_files()
        return jsonify({"ok": True, "data_dir": str(p), "migrated": migrated})


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
  tbody tr.is-used td{opacity:.62;font-style:italic;}
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
    <button class="btn ghost small" onclick="importFile()" title="Import a .csv / .xlsx file — each sheet becomes a subcategory">Import file</button>
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
        <input class="search" id="search" placeholder="Search…" oninput="debouncedLoad()">
      </div>
      <div class="table-wrap" id="tableWrap"></div>
    </section>
  </main>
</div>

<div class="overlay" id="overlay"></div>
<div class="toast" id="toast"></div>

<script>
const $ = s => document.querySelector(s);

// ---------- theme (light / dark) ----------
function applyTheme(t){
  document.documentElement.setAttribute("data-theme", t);
  try{ localStorage.setItem("inv_theme", t); }catch(e){}
  const b=document.getElementById("themeToggle");
  if(b) b.textContent = (t==="light") ? "☀️" : "🌙";
}
function toggleTheme(){
  const cur=document.documentElement.getAttribute("data-theme")||"dark";
  applyTheme(cur==="light" ? "dark" : "light");
}
applyTheme((function(){ try{ return localStorage.getItem("inv_theme"); }catch(e){ return null; } })() || "dark");

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
  const cols=["name","quantity","category_path",...usedCustom,
              ...(anyUsed?["status","used_note"]:[]),"date_added","last_modified"];
  const labels={name:"Name",quantity:"Qty",category_path:"Category",date_added:"Added",
                last_modified:"Modified",status:"Status",used_note:"Used for"};
  let html="<table><thead><tr><th style='width:6px;padding:0'></th>";
  cols.forEach(c=>html+=`<th>${esc(labels[c]||c)}</th>`);
  html+="<th></th></tr></thead><tbody>";
  items.forEach(it=>{
    const used=(it.status||"")==="used";
    const swatch=it.color?`background:${esc(it.color)}`:"background:transparent";
    // colour the whole row, or just the stripe at the start
    const rowStyle=(FILL_ROWS&&it.color)
      ? `background:${esc(it.color)};color:${readableOn(it.color)}`
      : "";
    html+=`<tr class="${used?'is-used':''}" style="${rowStyle}" onclick='editItem(${JSON.stringify(it.id)})'>`;
    html+=`<td style="padding:0"><div class="swatch" style="${swatch}"></div></td>`;
    cols.forEach(c=>{
      let v=it[c]||"";
      if(c==="status") v = used?"Used":"In stock";
      html+=`<td title="${esc(v)}">${esc(v)}</td>`;
    });
    html+=`<td style="white-space:nowrap">
           <button class="icon-btn" title="Add to order list" onclick='event.stopPropagation();addToOrder(${JSON.stringify(it.name)})'>🛒</button>
           <button class="icon-btn" title="${used?'Return to stock':'Mark as used'}" onclick='event.stopPropagation();toggleUsed(${JSON.stringify(it.id)},${used?"true":"false"})'>${used?"↩":"✔"}</button>
           <button class="icon-btn" title="Move" onclick='event.stopPropagation();moveItem(${JSON.stringify(it.id)},${JSON.stringify(it.category_path)})'>⇄</button>
           <button class="icon-btn" title="Delete" onclick='event.stopPropagation();deleteItem(${JSON.stringify(it.id)},${JSON.stringify(it.name)})'>🗑</button></td>`;
    html+="</tr>";
  });
  html+="</tbody></table>";
  wrap.innerHTML=html;
}

// ---------- modal infra ----------
function showModal(node){const o=$("#overlay");o.innerHTML="";o.appendChild(node);o.classList.add("show");}
function closeModal(){$("#overlay").classList.remove("show");$("#overlay").innerHTML="";}
$("#overlay").addEventListener("click",e=>{if(e.target===$("#overlay"))closeModal();});

function modalShell(title, bodyHtml, footHtml){
  const m=document.createElement("div"); m.className="modal";
  m.innerHTML=`<div class="m-head"><h3>${esc(title)}</h3><button class="icon-btn" onclick="closeModal()">✕</button></div>
    <div class="m-body">${bodyHtml}</div><div class="m-foot">${footHtml}</div>`;
  return m;
}

// ---------- categories ----------
async function addCategory(parent){
  const name=prompt(parent?`New subcategory under "${parent}":`:"New top-level category:");
  if(!name) return;
  try{ await api("/api/category",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({parent,name:name.trim()})});
    if(parent)EXPANDED.add(parent);
    toast("Category added"); loadState();
  }catch(e){toast(e.error||"Failed","err");}
}
async function renameCategory(node){
  const nn=prompt("Rename category:",node.name);
  if(!nn||nn===node.name) return;
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
      const go=confirm(`"${node.name}" and its subcategories contain ${e.count} item(s).\n\nOK = delete category AND all its items.\nCancel = keep everything.`);
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
  const reason=prompt(`Delete "${name||"this item"}"?\n\nIt will be archived in deleted_items.csv.\nOptionally note why (leave blank to just delete):`);
  if(reason===null) return;      // cancelled
  try{await api("/api/item/delete",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({id,reason})});
    toast("Deleted and archived"); loadState();
  }catch(e){toast(e.error||"Failed","err");}
}

async function toggleUsed(id,isUsed){
  let note="";
  if(!isUsed){
    note=prompt("Mark as used.\n\nWhere is it used / who has it?");
    if(note===null) return;
  }else if(!confirm("Put this item back in stock?")) return;
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
    const pick=prompt("Add which item?\n\n"+names.slice(0,40).join(", ")+(names.length>40?" …":""));
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
async function importFile(){
  toast("Choose a file in the dialog…");
  try{
    const r=await api("/api/import_file",{method:"POST",headers:{"Content-Type":"application/json"},body:"{}"});
    if(r.cancelled) return;
    toast(`Imported ${r.imported} item(s) from ${r.file}`);
    loadState();
  }catch(e){ toast(e.error||"Import failed","err"); }
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
    Pick an empty folder and your current data is copied there automatically.</p>`;
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
    try{const r=await api("/api/set_dir",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({path})});
      toast(r.migrated?"Data folder set · existing data copied":"Data folder set"); closeModal(); loadState();
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
