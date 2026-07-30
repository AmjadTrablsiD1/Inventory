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
| **Custom background** | Choose the light-mode background from presets or any colour |
| **Multi-select & right-click** | Select many rows and move, recolour, mark used or delete in one go |
| **Column management** | Rename, reorder, hide or delete columns — including junk from imports |
| **Choose your data folder** | Native folder picker; data is copied along when you switch |
| **Readable CSV mirror** | One CSV per category, subcategories as labelled sections |
| **Spreadsheet importer** | Feed it an `.xlsx` or `.csv` and it maps sheets → subcategories |
| **Built-in updater** | Checks GitHub for a new version and installs it, with a backup |
| **Item colours** | Pick a colour per item — and imports keep the colours from your spreadsheet |
| **Used, not deleted** | Mark something as used and record where it went, keeping it in the list |
| **Delete archive** | Deleted items are saved to `deleted_items.csv` with date, reason and details |
| **Clean CSV export** | Export just your data, with no internal ids or program columns |
| **Order list** | Build a purchase list with prices and links, export as CSV or PDF |
| **Bulk import** | Import every spreadsheet in a folder at once, with a preview first |

---

Runs on **Windows, macOS and Linux**.

## Requirements

- **Python 3.8+** (developed on 3.13)
- **Flask** — the launchers install it for you on first run

```bash
pip install -r requirements.txt
```

`openpyxl` is only needed to import `.xlsx` files; plain `.csv` import works without it.

> On Windows, install Python from [python.org](https://www.python.org/downloads/) and tick **"Add Python to PATH"** during setup.

---

## Running it

### Windows

Double-click **`Inventory Manager.vbs`**. It starts the app with **no console window**, installs Flask on first run if needed, and opens your browser. Double-clicking it again while it's running just reopens the browser instead of starting a second copy.

If you'd rather see the output — or if your workplace blocks `.vbs` scripts — double-click **`run.bat`** instead. It does the same thing but keeps a console window open, which makes troubleshooting easier.

To put it in your Start menu or on the desktop: right-click `Inventory Manager.vbs` → **Send to** → **Desktop (create shortcut)**.

### macOS

```bash
./build_launcher.sh
```

That produces `Inventory Manager.app` next to `app.py`. Double-click it: it starts the server in the background, installs Flask if missing, and opens your browser.

> macOS may show an "unidentified developer" warning the first time, because the app is built locally rather than signed. Right-click → **Open** once to approve it.

### Any platform — from the terminal

```bash
python app.py
```

(use `python3` on macOS/Linux). Then open <http://127.0.0.1:8765> — it opens automatically.

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
├── deleted_items.csv    # archive of everything you've deleted
├── order_list.csv       # your working "things to order" list
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

# A whole FOLDER: imports every .csv / .tsv / .xlsx inside it
python3 import_data.py ~/Documents/inventory-sheets

# ...including subfolders, which become parent categories
python3 import_data.py ~/Documents/inventory-sheets --recursive

# ...all nested under one category
python3 import_data.py ~/Documents/inventory-sheets --under "Imported 2026"
```

### Importing a whole folder

**Import** in the header offers both: a single file, or **a whole folder**. Folder import takes every `.csv`, `.tsv` and `.xlsx` in the folder and makes each file its own category.

- **Include subfolders** mirrors the folder tree as categories: `Lab2/measuring.csv` becomes `Lab2/measuring`.
- **Put everything under one category** nests the whole import beneath a name you choose.
- You always get a **preview first** — every file, how many items it would add, and which categories it would create. Nothing is written until you press **Import all**.
- The app's own files (`inventory.csv`, `categories.csv`, `history.csv`, `deleted_items.csv`, `order_list.csv` and the generated `by_category/` mirror) are **skipped automatically**, so pointing it at your own data folder can't duplicate your inventory. Excel lock files (`~$…`) and hidden files are skipped too.

From the command line, pass a folder instead of a file — add `--dry-run` to preview.

You can also click **Import file** in the app header to do the same thing through a dialog.

It figures out the messy parts for you:

- **Finds the name and quantity columns automatically** — recognises `name`, `item`, `title`, `description`, `qty`, `quantity`, `amount`, and German equivalents (`Bezeichnung`, `Menge`, `Anzahl`).
- **Detects the delimiter** — comma, semicolon, tab or pipe.
- **Handles encodings** — UTF-8, UTF-8-BOM and Latin-1, so umlauts and accents survive.
- Every other column becomes a custom field; blank rows and empty sheets are skipped.
- A plain `.csv` has no sheets, so it becomes a single category named after the file.
- **Keeps your colours** — a filled row (or a filled column, which fills a cell in every row) carries its colour onto the item. Excel *theme* colours are the one exception: the file stores them as a reference rather than an RGB value, so those come in without a colour. Direct and standard-palette fills work exactly.

Always try `--dry-run` first on a file you care about — it prints the full mapping and writes nothing.

---

## Working with rows

**Selecting.** Every row has a checkbox, and the header checkbox selects everything in view. **Shift-click** a row to select a range, **Ctrl/Cmd-click** to add or remove one. A bar appears above the table showing how many are selected, with actions that apply to all of them: **Rename** (one at a time), **Move to…**, **Colour…**, **Mark used**, **Back in stock** and **Delete**. **Esc** clears the selection.

**Right-click** any row for the same actions as a context menu. Right-clicking inside a selection acts on the whole selection; right-clicking outside it selects that one row first. The **⋯** button on each row opens the same menu if you prefer clicking.

**Nothing runs away from you.** The checkbox and item name are pinned to the left of the table and the action buttons are pinned to the right, so no matter how many custom columns you have or how far you scroll sideways, you can always see which row you are about to act on.

---

## Managing columns

Imported spreadsheets often arrive with unhelpful headings like `333`, or columns you simply don't need. **Right-click any column heading** (or click the **⋯** that appears when you hover it, or use the **Columns** button in the toolbar) to:

- **Rename** it — for a column that came from your data this really renames the field, in the CSV header and on every item, and updates any category that required it. For a built-in column such as *Category* or *Used for* only the heading you see changes; the underlying field keeps its name so nothing else breaks.
- **Move left / Move right** to reorder.
- **Hide** it — the column disappears from the table but the data stays in the file, so you can bring it back any time.
- **Delete** it — only offered for columns that came from your data. It removes that field from every item and from the CSV, and tells you how many items hold data in it first.

The **Columns** dialog lists everything at once with tick boxes for visibility and arrows for order. Settings live in `column_settings.json` next to your data, so they travel with the folder rather than being stuck in one browser.

> The *Name* column can't be hidden or moved — it stays pinned to the left as your anchor for every row.

---

## Appearance

The **🌙 / ☀️** button switches between light and dark mode. The **🎨** button next to it opens **Appearance**, where you can change the **light-mode background colour** — eight presets (paper white, warm cream, mint, sky, lavender, slate …) or any colour you like from the picker, plus **Reset to default**.

Panels, borders, muted text and the selection highlight are all derived from the colour you pick, so the interface stays coherent instead of just one value changing. If you choose something dark, the text switches to light automatically so everything stays readable. Dark mode keeps its own palette and is unaffected.

Your choice is remembered in the browser, along with the theme and the "fill whole row with colour" setting.

---

## Colours, status and deletion

**Colours.** Every item can carry a colour, set with the picker in the item form (or cleared with **No colour**). Colours also survive an import: if rows or columns in your spreadsheet are filled with a colour, each item keeps it.

By default the colour shows as a stripe down the left of the row. Tick **"fill whole row with colour"** in the toolbar to colour the entire row instead — the text automatically switches between black and white so it stays readable on light and dark colours alike. The choice is remembered.

**Used rather than deleted.** When you take something into use, click the **✔** on its row instead of deleting it. You're asked where it went, and the item stays in the inventory marked *Used* — greyed out, with the note and date in the `Status` and `Used for` columns. The **↩** button puts it back in stock. This keeps a truthful record of what you own versus what's in circulation.

**Deletion is archived.** Deleting asks for an optional reason and appends the item to `deleted_items.csv` before removing it — with the timestamp, reason, category, quantity, colour, status and all its custom fields preserved in an `other_fields` column. Nothing is silently lost. View it in the app under **History → View deleted items**.

---

## Exporting

**Export** in the header writes a clean CSV: friendly column headings, no internal `id`, and `Status` rendered as "In stock"/"Used" rather than raw values. You choose:

- everything, or just the selected category and its subcategories
- whether to include the date columns
- whether to include status / used-for columns
- whether to include the colour column

You pick where to save it with a normal Save dialog. This is the file to hand to someone else — unlike `inventory.csv`, it carries nothing the program needs internally.

---

## Order lists

**Order list** in the header opens a classic purchase list: **item, quantity, unit price, line total, link** — with the **sum of all prices** at the bottom, updating as you type.

- **+ Add line** for a free-text entry, **+ From inventory** to pick an existing item.
- The 🛒 button on any inventory row drops it straight onto the list.
- The list is saved in `order_list.csv`, so it survives restarts.
- Prices accept whatever you type: `12.50`, `12,50`, `1.234,56` or `€ 9,99` all work.

Export it as:

- **CSV** — clean columns plus a TOTAL row, ready for a spreadsheet.
- **PDF** — a tidy printable document with your title and reference note, right-aligned figures, clickable links, and the total. It pages automatically for long lists.

The PDF is generated directly by the app, so no extra libraries are needed. Set a different currency with `INVENTORY_CURRENCY` (default `€`).

---

## Updating

The version number in the header doubles as an update button. On startup the app quietly asks GitHub whether a newer version exists; if there is one, the badge lights up as **"Update available"**.

Clicking it shows what you have, what's published, and the recent commit messages. **Update now** downloads the repository archive and replaces the program files, then offers to restart.

Publishing a new version is just:

1. Bump `APP_VERSION` at the top of `app.py`.
2. Commit and push.

Everyone running the app sees the update on their next start.

**What it will and won't do:**

- It only ever fetches from **one repository** — auto-detected from this checkout's `git remote origin`, falling back to a constant in `app.py`. Override with `INVENTORY_REPO="owner/name"`.
- Every file it replaces is copied to `backup_before_update_<timestamp>/` first.
- It only writes files with known source extensions (`.py`, `.md`, `.txt`, `.sh`, `.bat`, `.vbs`, `.applescript`). Anything else in the archive — binaries, `.git`, CI config — is ignored.
- It refuses archives that are corrupt or don't contain `app.py`, and rejects entries with `..` in their paths.
- **Your inventory data is never touched.** It lives outside the program folder.
- Nothing is downloaded or installed without you clicking Update now. If GitHub is unreachable, the check fails silently and the app works normally.

> Because updates run code from the repository, only point `INVENTORY_REPO` at a repository you control.

---

## Configuration

| Variable | Default | Purpose |
|---|---|---|
| `INVENTORY_DIR` | `~/InventoryData` | Where the CSV files are stored. When set, it overrides the folder picked in the app — handy for running a second instance against different data |
| `INVENTORY_PORT` | `8765` | Port the local server listens on |
| `INVENTORY_REPO` | auto-detected | Update source, as `owner/name` |
| `INVENTORY_BRANCH` | `main` | Branch the updater reads from |
| `INVENTORY_CURRENCY` | `€` | Currency shown on order lists and exports |

Your chosen data folder is remembered in `~/.inventory_manager_config.json`.

---

## Security notes

This is a **local, single-user application with no authentication**. It deliberately binds to `127.0.0.1` only, so it is reachable just from the machine it runs on.

Do not change the host to `0.0.0.0` or put it on a public network. It exposes endpoints that read and write files on the host, and it has no login, no permissions, and no CSRF protection. If you need multiple people to use it, share the **data folder** over a network drive and let each person run their own local copy.

---

## Known limitations

- The **field type** (Text/Number/Date/Notes) chooses the right input widget while you type, but isn't stored — CSV holds plain text, so a date is saved as `2027-01-31`.
- Only one person should write at a time. There's no record locking, so two people saving simultaneously against the same shared folder can overwrite each other.
- Linux has no double-click launcher — run `python3 app.py`. The folder and file pickers work on all three platforms (Linux needs `zenity` or `kdialog` installed for them).

---

## Project structure

```
app.py                    # the web app: server, API, and embedded single-page UI
import_data.py            # standalone spreadsheet importer (also used by the app)
requirements.txt          # Python dependencies

Inventory Manager.vbs     # Windows: double-click launcher, no console window
run.bat                   # Windows: same, but with a visible console for troubleshooting

launcher.applescript      # macOS: source for the no-console launcher
build_launcher.sh         # macOS: builds "Inventory Manager.app" from the above
```

### Platform notes

| | Windows | macOS | Linux |
|---|---|---|---|
| Double-click launcher | `Inventory Manager.vbs` | build with `build_launcher.sh` | — |
| Folder / file pickers | PowerShell | AppleScript | zenity or kdialog |
| Runs windowless | `pythonw.exe` | background process | — |

`Inventory Manager.app` is a build artifact and is not committed — run `./build_launcher.sh` to produce it. Your inventory CSVs are never committed either; they live outside the repo (`~/InventoryData` by default).

---

## License

[MIT](LICENSE) — free to use, modify and distribute, with no warranty.
