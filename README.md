# Inventory Manager

A small, self-hosted web app for managing a categorized inventory — units, electronics, tools, anything. It runs locally on your own machine and stores everything as plain **CSV files you can open in Excel**, so your data never depends on the app being around.

Built for a single user or a small team sharing a folder on a network drive.

---

## Why it exists

Most inventory tools force a fixed set of fields on you. This one doesn't:

- Categories nest as deep as you like (`Electronics/Laptops/Docking Stations`).
- Any item can carry **any extra field** — serial number, location, warranty date, notes — and only the items that need it get it.
- A category can define **required fields** that new items in it (and its subcategories) inherit automatically.
- Everything auto-saves. There is no save button.
- Every change is logged with a timestamp.

---

## Features

| | |
|---|---|
| **Dynamic category tree** | Unlimited nesting, rename/move/delete with item-safety checks |
| **Dynamic item fields** | Add fields per item, typed as Text / Number / Date / Notes |
| **Required-field templates** | Per category, inherited by subcategories |
| **Auto-save** | Every change written to CSV immediately |
| **Change history** | Timestamped log of every add, edit, move and delete |
| **Light & dark mode** | Toggle in the header, remembered across restarts |
| **Choose your data folder** | Native folder picker; data is copied along when you switch |
| **Readable CSV mirror** | One CSV per category, subcategories as labelled sections |
| **Spreadsheet importer** | Feed it an `.xlsx` or `.csv` and it maps sheets → subcategories |

---

## Requirements

- **Python 3.8+** (developed on 3.13)
- **Flask** — `pip install flask`
- **openpyxl** — only needed to import `.xlsx` files: `pip install openpyxl`

```bash
python3 -m pip install flask openpyxl
```

---

## Running it

**Option 1 — double-click launcher (macOS, no Terminal window)**

```bash
./build_launcher.sh
```

That produces `Inventory Manager.app` next to `app.py`. Double-click it: it starts the server in the background, installs Flask if missing, and opens your browser. Clicking it again while it's running just reopens the browser instead of starting a second copy.

> macOS may show an "unidentified developer" warning the first time, because the app is built locally rather than signed. Right-click → **Open** once to approve it.

**Option 2 — from the terminal**

```bash
python3 app.py
```

Then open <http://127.0.0.1:8765> (it opens automatically).

> **Note on the port:** the default is **8765**, not the usual 5000, because on modern macOS the AirPlay Receiver permanently occupies port 5000.

---

## Where your data lives

By default everything goes in `~/InventoryData`. Change it with the **Data folder** button in the app (native folder picker), or by setting an environment variable:

```bash
INVENTORY_DIR="/Volumes/share/inventory" python3 app.py
```

If you pick a folder that has no inventory yet, your existing data is **copied there automatically** — the original folder is never deleted.

### File layout

```
InventoryData/
├── inventory.csv        # master data — one row per item
├── categories.csv       # the category tree + required fields
├── history.csv          # timestamped change log
└── by_category/         # auto-generated readable view
    ├── _Overview.csv    # index of every category with item counts
    ├── Electronics.csv
    └── Furniture.csv
```

**`inventory.csv`** is the master copy. Fixed columns `id, category_path, name, quantity, date_added, last_modified`, plus one column for every custom field you've created anywhere.

**`by_category/`** is a human-readable mirror, rebuilt automatically on every change. Each top-level category gets one CSV, with its subcategories as labelled sections — each showing only the columns that section actually uses:

```csv
# Category: Warehouse
# Subcategories: 2   Items: 5

=== Warehouse/Cables ===
id,name,quantity,Length_m,Type,date_added,last_modified
6d65cf...,Ethernet Cat6,40,5,Cat6,2026-01-04 09:12:00,2026-01-04 09:12:00

=== Warehouse/Tools ===
id,name,quantity,Serial Number,Condition,Location,date_added,last_modified
d61586...,Angle Grinder,2,AG-5540,Used,Shelf B1,2026-01-04 09:12:00,2026-01-04 09:12:00
```

`_Overview.csv` lists every category with its depth, direct item count, total including subcategories, and which file it lives in — open that one file and you can see the whole inventory at a glance.

> Files in `by_category/` are **generated output**. Edit them and your changes will be overwritten on the next save — edit through the app, or `inventory.csv`, instead.
>
> CSV files cannot contain multiple sheets (that's an Excel feature), so subcategories are represented as labelled sections rather than sheet tabs.

---

## Importing existing spreadsheets

`import_data.py` loads an existing spreadsheet into the inventory:

| Source | Becomes |
|---|---|
| the **file** | a category, named after the file |
| each **sheet** | a subcategory under it |
| each **row** | an item |
| each **column** | a field on that item |

```bash
# Excel workbook: every sheet becomes a subcategory
python3 import_data.py Warehouse.xlsx

# Preview exactly what would happen, without writing anything
python3 import_data.py Warehouse.xlsx --dry-run

# Import under a specific category instead of the file name
python3 import_data.py stock.csv --category Equipment

# Only one sheet
python3 import_data.py Warehouse.xlsx --sheet Tools

# No arguments → opens a file picker
python3 import_data.py
```

You can also click **Import file** in the app header to do the same thing through a dialog.

It figures out the messy parts for you:

- **Finds the name and quantity columns automatically** — recognises `name`, `item`, `title`, `description`, `qty`, `quantity`, `amount`, and German equivalents (`Bezeichnung`, `Menge`, `Anzahl`).
- **Detects the delimiter** — comma, semicolon, tab or pipe.
- **Handles encodings** — UTF-8, UTF-8-BOM and Latin-1, so umlauts and accents survive.
- Every other column becomes a custom field; blank rows and empty sheets are skipped.
- A plain `.csv` has no sheets, so it becomes a single category named after the file.

Always try `--dry-run` first on a file you care about — it prints the full mapping and writes nothing.

---

## Configuration

| Variable | Default | Purpose |
|---|---|---|
| `INVENTORY_DIR` | `~/InventoryData` | Where the CSV files are stored |
| `INVENTORY_PORT` | `8765` | Port the local server listens on |

Your chosen data folder is remembered in `~/.inventory_manager_config.json`.

---

## Security notes

This is a **local, single-user application with no authentication**. It deliberately binds to `127.0.0.1` only, so it is reachable just from the machine it runs on.

Do not change the host to `0.0.0.0` or put it on a public network. It exposes endpoints that read and write files on the host, and it has no login, no permissions, and no CSRF protection. If you need multiple people to use it, share the **data folder** over a network drive and let each person run their own local copy.

---

## Known limitations

- The **field type** (Text/Number/Date/Notes) chooses the right input widget while you type, but isn't stored — CSV holds plain text, so a date is saved as `2027-01-31`.
- Only one person should write at a time. There's no record locking, so two people saving simultaneously against the same shared folder can overwrite each other.
- The launcher script is macOS-only. On Windows and Linux, run `python3 app.py` (the folder and file pickers work on all three platforms).

---

## Project structure

```
app.py                  # the web app: server, API, and embedded single-page UI
import_data.py          # standalone spreadsheet importer (also used by the app)
launcher.applescript    # source for the macOS no-console launcher
build_launcher.sh       # builds Inventory Manager.app from the above
```

`Inventory Manager.app` is a build artifact and is not committed — run `./build_launcher.sh` to produce it. Your inventory CSVs are never committed either; they live outside the repo (`~/InventoryData` by default).

---

## License

[MIT](LICENSE) — free to use, modify and distribute, with no warranty.
