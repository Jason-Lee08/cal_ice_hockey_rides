#!/usr/bin/env python3
import json
import sys
from typing import Any, Dict, List, Optional, Tuple

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials as ServiceAccountCredentials

from nav_utils import *
from utils import *


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SERVICE_ACCOUNT_FILE = "calicehockey-map-d5de75ad4b3d.json"

def get_service() -> object:
    """
    Authenticate with a Google Cloud Service Account (no user OAuth).
    Requires:
      - service_account.json in the working directory (or point SERVICE_ACCOUNT_FILE to it)
      - the target Google Sheet shared with the service account email (Viewer is enough)
    """
    import os

    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(
            f"Service account key not found at '{SERVICE_ACCOUNT_FILE}'. "
            "Download a JSON key for your service account and set "
            "SERVICE_ACCOUNT_FILE=<path> or place 'service_account.json' next to this script."
        )

    creds = ServiceAccountCredentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return build("sheets", "v4", credentials=creds)


def rgb_from_cell(cell: Dict) -> Optional[Tuple[float, float, float]]:
    """
    Extract an (r,g,b) tuple in [0,1] from a cell's background color, or None if not present.
    Checks backgroundColorStyle.rgbColor first (preferred), then backgroundColor.
    """
    fmt = (cell.get("userEnteredFormat") or {})
    # backgroundColorStyle has priority and preserves theme/alpha better
    style = fmt.get("backgroundColorStyle") or {}
    rgb = style.get("rgbColor")

    if rgb:
        return (
            float(rgb.get("red", 0.0)),
            float(rgb.get("green", 0.0)),
            float(rgb.get("blue", 0.0)),
        )

    col = fmt.get("backgroundColor")
    if col:
        return (
            float(col.get("red", 0.0)),
            float(col.get("green", 0.0)),
            float(col.get("blue", 0.0)),
        )
    return None


def is_nonwhite_highlight(rgb: Optional[Tuple[float, float, float]], tol: float = 0.02) -> bool:
    """True if the cell background is not (approximately) white."""
    if rgb is None:
        return False
    r, g, b = rgb
    return not (abs(r - 1.0) <= tol and abs(g - 1.0) <= tol and abs(b - 1.0) <= tol)


def is_red(rgb: Optional[Tuple[float, float, float]]) -> bool:
    """
    Heuristic for "red-ish" highlights.
    Tune thresholds if your sheet uses a different red:
      - high red, comparatively low green/blue
    """
    if rgb is None:
        return False
    r, g, b = rgb
    return (r == 0.95686275)


def col_to_letter(col_idx_zero_based: int) -> str:
    """0->A, 1->B, ..."""
    n = col_idx_zero_based + 1
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def extract_sheet_grid(service, spreadsheet_id: str, sheet_name: str) -> Dict:
    """
    Fetch grid data (values + formats) for the given sheet name.
    """
    fields = (
        "sheets(properties.title,properties.gridProperties.rowCount,properties.gridProperties.columnCount,"
        "data.rowData.values.userEnteredValue,"
        "data.rowData.values.formattedValue,"
        "data.rowData.values.userEnteredFormat.backgroundColor,"
        "data.rowData.values.userEnteredFormat.backgroundColorStyle)"
    )
    resp = (
        service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            ranges=[sheet_name],
            includeGridData=True,
            fields=fields,
        )
        .execute()
    )
    sheets = resp.get("sheets", [])
    if not sheets:
        raise RuntimeError(f"Sheet named '{sheet_name}' not found or empty.")
    # We asked for a specific range; the first element corresponds to it
    return sheets[0]


def get_cell_value(cell: Dict) -> Optional[str]:
    """
    Prefer formattedValue (looks like what you see in UI), else the raw userEnteredValue.
    Returns None for empty.
    """
    if cell is None:
        return None
    if "formattedValue" in cell:
        val = cell["formattedValue"]
        if val == "":
            return None
        return val
    uev = cell.get("userEnteredValue") or {}
    # Could be stringValue, numberValue, boolValue, formulaValue...
    for key in ("stringValue", "numberValue", "boolValue"):
        if key in uev:
            return str(uev[key])
    return None


def parse_groups(sheet_data: Dict) -> List[Dict]:
    """
    Walk down each column and build carpool groups per the rules.
    """
    data_blocks = sheet_data.get("data", [])
    if not data_blocks:
        return []

    # The first (and only) data block corresponds to the requested range (the whole sheet).
    rows = data_blocks[0].get("rowData", []) or []
    row_count = len(rows)
    # Determine max columns actually present in data
    max_cols = 0
    for r in rows:
        vals = r.get("values", [])
        max_cols = max(max_cols, len(vals))

    groups = []

    for c in range(max_cols):
        # Capacity in row 0
        capacity_cell = rows[0]["values"][c] if (row_count > 0 and len(rows[0].get("values", [])) > c) else {}
        capacity_str = get_cell_value(capacity_cell)
        capacity: Optional[int] = None
        if capacity_str is not None:
            try:
                capacity = int(float(capacity_str))
            except Exception:
                capacity = None  # silently ignore non-numeric

        # Find driver: first highlighted (non-white), non-empty cell from row 1 downwards, excluding red
        driver = None
        driver_row_idx = None
        for r in range(1, row_count):
            row_vals = rows[r].get("values", [])
            cell = row_vals[c] if c < len(row_vals) else {}
            name = get_cell_value(cell)
            if name is None or name.strip() == "":
                # not highlighted & empty — keep scanning (driver could be lower)
                # But if you want to stop on the first blank above the driver area, remove 'continue' and break instead.
                continue
            rgb = rgb_from_cell(cell)
            if is_red(rgb):
                # Red-highlighted potential driver means this person is excluded; but we still have to keep scanning
                # for a valid (non-red) highlighted driver.
                continue
            if is_nonwhite_highlight(rgb):
                driver = name.strip()
                driver_row_idx = r
                break

        if not driver:
            # No valid driver found; skip this column
            continue

        # If the driver cell itself is red (handled above) we'd have skipped; still, guard:
        driver_cell = rows[driver_row_idx]["values"][c]
        if is_red(rgb_from_cell(driver_cell)):
            # Skip entire group
            continue

        # Collect passengers: all non-empty cells beneath the driver until first blank row in this column
        passengers: List[str] = []
        r = driver_row_idx + 1
        while r < row_count:
            row_vals = rows[r].get("values", [])
            cell = row_vals[c] if c < len(row_vals) else {}
            name = get_cell_value(cell)
            if name is None or name.strip() == "":
                break  # groups stop at the first empty cell
            rgb = rgb_from_cell(cell)
            if not is_red(rgb):
                passengers.append(name.strip())
            # If red, skip that person silently
            r += 1

        groups.append(
            {
                "column": col_to_letter(c),
                "capacity": capacity,
                "driver": driver,
                "passengers": passengers,
                "driver_row": driver_row_idx,   # 0-based
                "driver_col": c                 # 0-based
            }
        )


    return groups

def attach_addresses_to_groups(groups: List[Dict], name_to_addr: Dict[str, str]) -> List[Dict]:
    """
    For each group, lookup the driver's address (starting_address)
    and each passenger's address (pickup_addresses).
    Adds two fields to each group dict: 'starting_address' and 'pickup_addresses'.

    Behavior on missing addresses:
      - Driver missing address: FATAL (reported and process exits after all groups are checked)
      - Passenger missing address: LOUD ERROR banner (non-fatal; passenger is skipped)
    """
    fatal_missing: List[str] = []  # collect driver issues to exit after surfacing all
    for g in groups:
        driver_name = g.get("driver", "")
        passengers = g.get("passengers", [])
        column = g.get("column", "?")

        driver_key = normalize_name(driver_name) if driver_name else ""

        # ------- DRIVER ADDRESS LOOKUP --------
        starting_address = None
        for k in name_to_addr.keys():
            if driver_key and driver_key in k:
                starting_address = name_to_addr[k]
                break

        if not starting_address:
            # Loud, visible, red banner; mark as fatal and keep checking others so we show all issues at once
            title = f"Driver address NOT FOUND (Column {column})"
            details = f"""
            Driver: {driver_name or '(missing)'}
            Column: {column}
            Why this matters: Without a starting address, we cannot compute routes for this group.

            How to fix:
              1) Make sure the driver's full name exactly matches the "Full Address & Contact Info" sheet.
              2) Confirm the driver row in "{column}" is not marked red or anything weird with the highlighting
              3) Check for stray whitespace or punctuation; the code normalizes with lowercasing + internal single spaces.

            """
            print_error_banner(title, details, fatal=True)
            fatal_missing.append(driver_name or f"(column {column})")

        # ---- PASSENGER ADDRESS LOOKUP ----
        pickup_addresses = set()
        for p in passengers:
            addr = None
            normal_name = normalize_name(p)
            for k in name_to_addr.keys():
                if normal_name in k:
                    addr = name_to_addr[k]
                    break

            if addr:
                if addr != starting_address:
                    pickup_addresses.add(addr)
            else:
                # Loud, visible, non-fatal banner for passengers (skip them for routing)
                title = "Passenger address NOT FOUND"
                details = f"""
Passenger: {p}
Driver / Column: {driver_name or '(unknown)'} / {column}
Action taken: This passenger will be SKIPPED in route computation.

How to fix:
    1) Ensure the passenger's name exists (column A) and has a valid address (column B)
        in the "Full Address & Contact Info" sheet (row after header).
    2) Confirm the passenger's name spelling matches what's in the address sheet.
    3) Remove red highlight if present; red-marked names are treated as excluded.
                """
                print_error_banner(title, details, fatal=False)
                # also keep the previous lightweight stderr line if you want:
                # print(f"[WARN] No address found for passenger '{p}'", file=sys.stderr)

        g["starting_address"] = starting_address
        g["pickup_addresses"] = list(pickup_addresses)

    # If any driver was missing an address, abort after reporting all problems
    if fatal_missing:
        summary = " / ".join(fatal_missing[:5]) + (" ..." if len(fatal_missing) > 5 else "")
        title = "Cannot continue: one or more drivers have no address"
        details = f"""
        A driver must have a valid starting address to compute routes.
        Affected drivers (sample): {summary}

        Resolve driver address issues and re-run the script.
        """
        print_error_banner(title, details, fatal=True)
        sys.exit(1)

    return groups

def main():
    spreadsheet_id = "1sVtfMNfGRK7UE8JGKcOxsJO6uhMj6V1SGapW-K6xHvI"
    rides_sheet = "Rides Sheet"
    addr_sheet = "Full Address & Contact Info"
    final_destination = "519 18th St, Oakland, CA 94612"

    service = get_service()

    # 1) Parse groups
    sheet_data = extract_sheet_grid(service, spreadsheet_id, rides_sheet)
    groups = parse_groups(sheet_data)

    # 2) Attach addresses
    name_to_addr = build_name_to_address_map(service, spreadsheet_id, addr_sheet, start_row=1)
    groups = attach_addresses_to_groups(groups, name_to_addr)

    # 3) Compute BOTH directions with live traffic
    groups = compute_forward_and_return_routes(groups, final_destination)

    # 4) Attach Google Maps URLs for both
    groups = attach_maps_links_both_directions(groups, start_navigation=False)

    # 5) Write outbound link into driver cell (blue clickable name)
    write_driver_links(service, spreadsheet_id, rides_sheet, groups)

    # 6) Write return link into the cell to the right of the driver
    write_return_links_right_of_driver(service, spreadsheet_id, rides_sheet, groups)

    with open("results.json", "w") as f:
        f.write(json.dumps(groups, indent=2))
    
    LINK = "https://docs.google.com/spreadsheets/d/1sVtfMNfGRK7UE8JGKcOxsJO6uhMj6V1SGapW-K6xHvI/"
    print(f"Completed! Check the google sheet here: {LINK}")

    print_check_results_reminder("results.json")


if __name__ == "__main__":
    main()
