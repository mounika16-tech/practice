#!/usr/bin/env python3
"""
excel_updater.py

Automates two kinds of edits to an Excel workbook for a given "unique id"
(e.g. a site code like "BRP1E"). Both are driven by one JSON config so the
same run works across many workbooks.

PART 1 -- APPEND  ("append_sheets" in the config)
    Appends a fixed list of new item rows AFTER the last used row of a named
    sheet. Never inserts, never overwrites. Each sheet has its own column
    layout. New cells are highlighted.

PART 2 -- LOOKUP & OVERWRITE  ("update_sheets" in the config)
    In a named sheet, scans a "match column" (e.g. column I) for specific
    values (e.g. sgi_direct_net_1_subnet_5). For every row that matches, it
    writes a defined set of values into other columns (e.g. O, U, V, AC, AD).
    These target cells ARE overwritten on purpose, even if they already hold
    data, and every cell written is highlighted. No other column, row, or
    sheet is touched.

SAFETY
-------
- Part 1 only appends; it pre-checks that every target cell is empty and
  aborts the sheet rather than overwriting anything.
- Part 2 only writes to the explicitly-listed columns, and only on rows whose
  match column exactly equals a listed value. Everything else is left alone.
- Nothing is saved unless all sheets processed cleanly.
- --dry-run shows exactly what would change (including the old value being
  replaced) without saving anything.
- --output writes to a copy, leaving the original file untouched. Recommended
  for the first real run.

--------------------------------------------------------------------------
USAGE

  Preview (no changes written):
    python excel_updater.py --workbook book.xlsx --uid BRP1E \
        --config template_config.json --dry-run

  Write to a copy:
    python excel_updater.py --workbook book.xlsx --uid BRP1E \
        --config template_config.json --output book_updated.xlsx

  Write in place:
    python excel_updater.py --workbook book.xlsx --uid BRP1E \
        --config template_config.json

  Many workbooks at once (CSV with columns: workbook,uid):
    python excel_updater.py --batch batch.csv --config template_config.json

  Run only part 1 or only part 2:
    --only append      (part 1 only)
    --only update      (part 2 only)
--------------------------------------------------------------------------
"""

import argparse
import csv
import json
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Use the bundled copy of openpyxl in ./lib if it is present, so this script
# runs on machines where "pip install" is blocked by a corporate proxy.
# A system-installed openpyxl (if any) still wins; lib/ is only a fallback.
# ---------------------------------------------------------------------------
_LIB = Path(__file__).resolve().parent / "lib"
if _LIB.is_dir():
    sys.path.append(str(_LIB))

try:
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill
except ImportError:
    sys.exit(
        "ERROR: could not import openpyxl.\n"
        "  - If you received a 'lib' folder with this script, keep it in the\n"
        "    SAME folder as excel_updater.py and run the script again.\n"
        "  - Otherwise install it with:  pip install openpyxl\n"
    )

DEFAULT_HIGHLIGHT = "FFFF00"


# ----------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------

def make_fill(hex_color):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")


def render(value, **kwargs):
    """Substitute {uid} / {item} placeholders in a template string."""
    return value.format(**kwargs) if isinstance(value, str) else value


def next_empty_row(ws):
    return ws.max_row + 1 if ws.max_row and ws.max_row > 0 else 1


def norm(v):
    """Normalize a cell value for matching: string, stripped. None -> ''."""
    if v is None:
        return ""
    return str(v).strip()


# ----------------------------------------------------------------------
# PART 1 -- append rows
# ----------------------------------------------------------------------

def resolve_append_rows(sheet_cfg, uid, global_items):
    """
    Turn a sheet's append config into a concrete list of rows, each a dict of
    {column_letter: final_value}. Supports three config forms:

    (a) legacy / item-driven, flat:
            {"A": "{uid}", "B": "dlp-{uid}-ps_{item}"}
        -> one row per entry in the global "items" list.

    (b) item-driven, explicit:
            {"columns": {...}, "items": [...]}   ("items" optional, falls back
        to the global list)

    (c) explicit rows (no items involved):
            {"rows": [{"A": "dlp-{uid}-ong_INET_ADDRESS", "B": "65.66.16.0",
                       "C": 21, "D": "ENT_PREFIX"}, ...]}

    {uid} is substituted in every form. {item} only applies to (a) and (b).
    """
    if "rows" in sheet_cfg:
        return [{c: render(v, uid=uid) for c, v in row.items()}
                for row in sheet_cfg["rows"]]

    col_template = sheet_cfg.get("columns", sheet_cfg)
    items = sheet_cfg.get("items", global_items)
    if not items:
        raise ValueError("Append sheet uses an item template but no items are defined.")
    return [{c: render(t, uid=uid, item=item) for c, t in col_template.items()}
            for item in items]


def resolve_lookups(ws, rows, lookup_cfg, start_row, first_row=1, dry_run=False):
    """
    PART 3 item 2 -- fill a column by copying a value from earlier rows of the
    SAME sheet, matched on a set of key columns.

    For each new row:
      * If its value in 'skip_column' is one of 'skip_values', leave the
        templated value alone. ('skip_values' is empty by default, so every
        row goes through the lookup.)
      * Otherwise, scan EVERY row above the append point (from row 1 --
        no header row is assumed or skipped) and find those whose
        'match_columns' all equal this row's values, and copy their
        'value_column' across. When the same key repeats in the new block,
        the Nth new row takes the Nth match in sheet order.
      * If NO matching row is found above (or the matches run out), the
        value column is left BLANK rather than keeping the template value.

    Existing rows are only READ. They are never modified.

    Returns a list of human-readable notes describing what was resolved.
    """
    value_col = lookup_cfg["value_column"]
    match_cols = lookup_cfg["match_columns"]
    skip_col = lookup_cfg.get("skip_column")
    skip_values = {norm(v) for v in lookup_cfg.get("skip_values", [])}

    # Build an ordered index of the existing data ABOVE the new block.
    index = {}
    for r in range(first_row, start_row):
        key = tuple(norm(ws[f"{c}{r}"].value) for c in match_cols)
        if all(k == "" for k in key):
            continue
        val = ws[f"{value_col}{r}"].value
        if norm(val) == "":
            continue
        index.setdefault(key, []).append((r, val))

    notes = []
    used = {}   # key -> how many matches consumed so far

    for i, row_vals in enumerate(rows):
        target_row = start_row + i

        if skip_col is not None and norm(row_vals.get(skip_col)) in skip_values:
            notes.append(f"    row {target_row}: {skip_col}={row_vals.get(skip_col)} is STATIC "
                         f"-> no lookup, {value_col} stays {row_vals.get(value_col)!r}")
            continue

        key = tuple(norm(row_vals.get(c)) for c in match_cols)
        matches = index.get(key, [])
        n = used.get(key, 0)

        if n < len(matches):
            src_row, src_val = matches[n]
            used[key] = n + 1
            old = row_vals.get(value_col)
            row_vals[value_col] = src_val
            notes.append(f"    row {target_row}: key {key} -> copied {src_val!r} "
                         f"from row {src_row}" +
                         (f" (template had {old!r})" if norm(old) != norm(src_val) else ""))
        else:
            row_vals[value_col] = None
            notes.append(f"    row {target_row}: no{' further' if matches else ''} "
                         f"match above for {match_cols}={key} "
                         f"(found {len(matches)}, needed {n + 1}) -> "
                         f"{value_col} left BLANK")

    return notes


def append_rows_to_sheet(ws, rows, fill, dry_run=False):
    """
    Append the given prepared rows after the sheet's last used row.
    Aborts (raises) if any target cell is not empty -- never overwrites.
    Returns (start_row, end_row).
    """
    if not rows:
        return None, None

    start_row = next_empty_row(ws)

    # Pre-flight: confirm every cell we intend to touch is empty.
    for i, row_vals in enumerate(rows):
        row_num = start_row + i
        for col_letter in row_vals:
            cell = ws[f"{col_letter}{row_num}"]
            if norm(cell.value) != "":
                raise RuntimeError(
                    f"APPEND ABORTED: {ws.title}!{cell.coordinate} is not empty "
                    f"(contains {cell.value!r}). Nothing written to this sheet."
                )

    # Always write into the in-memory workbook so that later blocks see these
    # rows exactly as they would in a real run. --dry-run simply never saves.
    for i, row_vals in enumerate(rows):
        row_num = start_row + i
        for col_letter, value in row_vals.items():
            cell = ws[f"{col_letter}{row_num}"]
            cell.value = value
            cell.fill = fill
        if dry_run:
            preview = "  ".join(f"{c}{row_num}={v!r}" for c, v in row_vals.items())
            print(f"      + {preview}")

    return start_row, start_row + len(rows) - 1


# ----------------------------------------------------------------------
# PART 2 -- find rows by a key column, overwrite specific columns
# ----------------------------------------------------------------------

def update_matching_rows(ws, spec, uid, fill, dry_run=False):
    """
    Scan the WHOLE of spec['match_column'] (every row from row 1 down --
    no header row is assumed, since the data may start anywhere) for each
    key in spec['matches'].
    For every matching row, write that key's column->value map, overwriting
    whatever is there, and highlight those cells.

    Returns (rows_changed, cells_changed, unmatched_keys).
    """
    match_col = spec["match_column"]
    matches = spec["matches"]
    first_row = spec.get("first_row", 1)

    # Normalized lookup so trailing spaces / case don't cause silent misses.
    case_sensitive = spec.get("case_sensitive", True)

    def key_of(s):
        return s if case_sensitive else s.lower()

    lookup = {key_of(norm(k)): v for k, v in matches.items()}
    seen_keys = set()

    rows_changed = 0
    cells_changed = 0

    for row_num in range(first_row, ws.max_row + 1):
        cell_val = key_of(norm(ws[f"{match_col}{row_num}"].value))
        if cell_val == "" or cell_val not in lookup:
            continue

        seen_keys.add(cell_val)
        col_values = lookup[cell_val]
        rows_changed += 1

        for col_letter, tmpl in col_values.items():
            cell = ws[f"{col_letter}{row_num}"]
            new_val = render(tmpl, uid=uid)
            old_val = cell.value

            cell.value = new_val
            cell.fill = fill

            if dry_run:
                if norm(old_val) == norm(new_val):
                    print(f"      = {col_letter}{row_num} already {new_val!r} (would re-highlight)")
                elif norm(old_val) == "":
                    print(f"      + {col_letter}{row_num} (empty) -> {new_val!r}")
                else:
                    print(f"      ~ {col_letter}{row_num} {old_val!r} -> {new_val!r}  [OVERWRITE]")
            cells_changed += 1

    unmatched = [k for k in lookup if k not in seen_keys]
    return rows_changed, cells_changed, unmatched


# ----------------------------------------------------------------------
# driver
# ----------------------------------------------------------------------

def process_workbook(workbook_path, uid, config, output_path=None,
                     dry_run=False, only=None, only_sheets=None):
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    fill = make_fill(config.get("highlight_color", DEFAULT_HIGHLIGHT))
    items = config.get("items", [])
    append_cfg = config.get("append_sheets", {})
    update_cfg = config.get("update_sheets", {})

    wb = load_workbook(workbook_path)
    print(f"\n=== {workbook_path.name}   uid={uid} ===")
    did_something = False

    # ---- PART 1 ----
    if only in (None, "append"):
        for sheet_name, sheet_cfg in append_cfg.items():
            if only_sheets and sheet_name not in only_sheets:
                continue
            if sheet_name not in wb.sheetnames:
                print(f"  [append] '{sheet_name}': not in this workbook, skipped")
                continue
            ws = wb[sheet_name]

            # A sheet may define one block, or a list of blocks applied in order.
            blocks = sheet_cfg if isinstance(sheet_cfg, list) else [sheet_cfg]

            for bi, block in enumerate(blocks):
                label = f"'{sheet_name}'" + (f" block {bi + 1}" if len(blocks) > 1 else "")
                rows = resolve_append_rows(block, uid, items)
                sheet_fill = fill
                if isinstance(block, dict) and block.get("highlight_color"):
                    sheet_fill = make_fill(block["highlight_color"])
                if dry_run:
                    print(f"  [append] {label}:")

                # Part 3 item 2: resolve any values copied from earlier rows.
                lookup_cfg = block.get("lookup") if isinstance(block, dict) else None
                if lookup_cfg:
                    block_start = next_empty_row(ws)
                    notes = resolve_lookups(ws, rows, lookup_cfg, block_start,
                                            first_row=block.get("first_row", 1))
                    warnings = [n for n in notes if "BLANK" in n]
                    if dry_run:
                        for n in notes:
                            print(n)
                    elif warnings:
                        for n in warnings:
                            print(n)

                start, end = append_rows_to_sheet(ws, rows, sheet_fill, dry_run)
                verb = "would append" if dry_run else "appended"
                print(f"  [append] {label}: {verb} {len(rows)} row(s) at rows {start}-{end}")
                did_something = True

    # ---- PART 2 ----
    if only in (None, "update"):
        for sheet_name, spec in update_cfg.items():
            if only_sheets and sheet_name not in only_sheets:
                continue
            if sheet_name not in wb.sheetnames:
                print(f"  [update] '{sheet_name}': not in this workbook, skipped")
                continue
            ws = wb[sheet_name]
            if dry_run:
                print(f"  [update] '{sheet_name}':")
            rows, cells, unmatched = update_matching_rows(ws, spec, uid, fill, dry_run)
            verb = "would update" if dry_run else "updated"
            print(f"  [update] '{sheet_name}': {verb} {rows} row(s), {cells} cell(s) "
                  f"(matched on column {spec['match_column']})")
            if unmatched:
                print(f"           WARNING: no row found for: {', '.join(sorted(unmatched))}")
            did_something = True

    if not did_something:
        print("  (nothing matched -- no changes)")
        return

    if not dry_run:
        save_path = Path(output_path) if output_path else workbook_path
        wb.save(save_path)
        print(f"  SAVED -> {save_path}")
    else:
        print("  DRY RUN -- nothing saved")


def load_config(path):
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    if "append_sheets" not in cfg and "update_sheets" not in cfg:
        raise ValueError("Config must contain 'append_sheets' and/or 'update_sheets'.")
    for name, spec in cfg.get("append_sheets", {}).items():
        blocks = spec if isinstance(spec, list) else [spec]
        for blk in blocks:
            has_rows = isinstance(blk, dict) and "rows" in blk
            has_local_items = isinstance(blk, dict) and blk.get("items")
            if not has_rows and not has_local_items and not cfg.get("items"):
                raise ValueError(
                    f"append_sheets['{name}'] is item-driven but no 'items' list is defined "
                    f"(neither on the sheet nor globally)."
                )
            if isinstance(blk, dict) and "lookup" in blk:
                lk = blk["lookup"]
                for req in ("value_column", "match_columns"):
                    if req not in lk:
                        raise ValueError(
                            f"append_sheets['{name}'] lookup needs '{req}'.")
    for name, spec in cfg.get("update_sheets", {}).items():
        if "match_column" not in spec or "matches" not in spec:
            raise ValueError(f"update_sheets['{name}'] needs 'match_column' and 'matches'.")
    return cfg


def run_batch(batch_csv, config, dry_run, only):
    with open(batch_csv, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        headers = {h.strip() for h in (reader.fieldnames or [])}
        if not {"workbook", "uid"}.issubset(headers):
            raise ValueError("Batch CSV must have columns: workbook,uid")

        errors = []
        for i, row in enumerate(reader, start=2):
            wb_path, uid = row["workbook"].strip(), row["uid"].strip()
            try:
                process_workbook(wb_path, uid, config, dry_run=dry_run, only=only)
            except Exception as e:
                errors.append(f"line {i} ({wb_path} / {uid}): {e}")

        if errors:
            print("\nCompleted WITH ERRORS:", file=sys.stderr)
            for err in errors:
                print(f"  - {err}", file=sys.stderr)
            sys.exit(1)


def main():
    p = argparse.ArgumentParser(description="Append and/or update Excel sheets per unique id.")
    p.add_argument("--config", required=True, help="Path to the JSON config.")
    p.add_argument("--dry-run", action="store_true", help="Preview changes; save nothing.")
    p.add_argument("--only", choices=["append", "update"], help="Run only part 1 or only part 2.")

    p.add_argument("--workbook", help="Workbook path (single-run mode).")
    p.add_argument("--uid", help="Unique id, e.g. BRP1E (single-run mode).")
    p.add_argument("--output", help="Save to a new file instead of overwriting.")
    p.add_argument("--sheets", nargs="+", help="Limit to these sheet names.")

    p.add_argument("--batch", help="CSV with columns: workbook,uid (batch mode).")

    args = p.parse_args()
    config = load_config(args.config)

    if args.batch:
        run_batch(args.batch, config, args.dry_run, args.only)
    elif args.workbook:
        uid = args.uid
        if not uid:
            try:
                uid = input("Enter the unique id (e.g. HAPPY): ").strip()
            except EOFError:
                uid = ""
            if not uid:
                p.error("A unique id is required.")
        process_workbook(args.workbook, uid, config, output_path=args.output,
                         dry_run=args.dry_run, only=args.only, only_sheets=args.sheets)
    else:
        p.error("Provide either --batch <csv>, or --workbook (uid will be requested if omitted).")


if __name__ == "__main__":
    main()
