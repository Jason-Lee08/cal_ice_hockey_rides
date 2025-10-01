from googleapiclient.errors import HttpError
import sys
from typing import Dict
from urllib.parse import quote
from urllib.parse import quote

def fetch_values(service, spreadsheet_id: str, range_a1: str):
    """Read values from a sheet range using the values API (faster, no formatting)."""
    try:
        resp = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=range_a1
        ).execute()
        return resp.get("values", [])
    except HttpError as e:
        print(f"[ERROR] Failed to fetch values for {range_a1}: {e}", file=sys.stderr)
        return []


def normalize_name(name: str) -> str:
    """Simple normalization to improve matching between sheets."""
    return " ".join(name.strip().split()).lower()


def build_name_to_address_map(service, spreadsheet_id: str,
                              sheet_name: str = "Full Address & Contact Info",
                              start_row: int = 1) -> Dict[str, str]:
    """
    Build mapping from full name (col A) to address (col B).
    Assumes first row is a header; set start_row accordingly if not.
    """
    # Read A and B columns over a generous row range (adjust if your sheet is huge)
    # A:B gets both columns; we’ll skip the header row.
    rows = fetch_values(service, spreadsheet_id, f"{sheet_name}!A:B")
    name_to_addr: Dict[str, str] = {}

    for i, row in enumerate(rows):
        if i < start_row:
            # Skip header row (if your sheet has no header, set start_row=0)
            continue
        if not row:
            continue

        name = row[0].strip() if len(row) >= 1 and row[0] else ""
        addr = row[1].strip() if len(row) >= 2 and row[1] else ""

        if not name:
            continue  # blank name row
        if not addr:
            # Optionally warn: missing address for a known name
            # print(f"[WARN] No address for name '{name}' on row {i+1}", file=sys.stderr)
            continue

        key = normalize_name(name)
        name_to_addr[key] = addr

    return name_to_addr

def a1_from_row_col(row0: int, col0: int) -> str:
    """Convert 0-based (row, col) to A1 like 'B7'."""
    # col0 -> letters
    n = col0 + 1
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    # rows are 1-based in A1
    return f"{s}{row0 + 1}"

def hyperlink_formula(url: str, label: str) -> str:
    # Use USER_ENTERED so Sheets treats this as a formula:
    # =HYPERLINK("https://...", "Driver Name")
    safe_label = label.replace('"', '”')  # avoid breaking the formula
    safe_url = url.replace('"', '%22')
    return f'=HYPERLINK("{safe_url}","{safe_label}")'

def write_driver_links(service, spreadsheet_id: str, sheet_name: str, groups: list[dict]) -> None:
    """
    For each group with best_route + maps_url, replace the driver's cell value with a HYPERLINK
    that still displays the driver's name. Background highlight is preserved.
    """
    values_body = {"valueInputOption": "USER_ENTERED", "data": []}

    for g in groups:
        url = g.get("maps_url")
        driver = g.get("driver")
        r0 = g.get("driver_row")
        c0 = g.get("driver_col")
        if not (url and driver and isinstance(r0, int) and isinstance(c0, int)):
            continue
        a1 = a1_from_row_col(r0, c0)
        rng = f"'{sheet_name}'!{a1}"
        values_body["data"].append({
            "range": rng,
            "values": [[hyperlink_formula(url, driver)]]
        })

    if not values_body["data"]:
        return

    service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=values_body
    ).execute()

def write_return_links_right_of_driver(service, spreadsheet_id: str, sheet_name: str, groups: list[dict]) -> None:
    """
    Writes a HYPERLINK formula labeled 'Return' into the cell to the RIGHT of the driver's cell
    (same row, col+1). We keep it unhighlighted so the parser ignores it.
    """
    values_body = {"valueInputOption": "USER_ENTERED", "data": []}

    for g in groups:
        url = g.get("maps_return_url")
        r0 = g.get("driver_row")
        c0 = g.get("driver_col")
        if not (url and isinstance(r0, int) and isinstance(c0, int)):
            continue

        # Target: same row, next column to the right
        a1 = a1_from_row_col(r0, c0 + 1)
        rng = f"'{sheet_name}'!{a1}"
        label = "Return"
        values_body["data"].append({
            "range": rng,
            "values": [[hyperlink_formula(url, label)]]
        })

    if not values_body["data"]:
        return

    service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=values_body
    ).execute()

def _qp(s: str) -> str:
    """URL-encode a single component for Google Maps."""
    return quote(s, safe="")

def build_gmaps_directions_url(order_addresses: list[str], start_navigation: bool = False) -> str:
    """
    Create a Google Maps directions URL with fixed waypoint order.
    order_addresses: [start, stop1, stop2, ..., final]
    """
    if not order_addresses or len(order_addresses) < 2:
        return ""

    origin = order_addresses[0]
    destination = order_addresses[-1]
    waypoints_list = order_addresses[1:-1]  # in the exact order we computed

    base = "https://www.google.com/maps/dir/?api=1"
    url = f"{base}&origin={_qp(origin)}&destination={_qp(destination)}&travelmode=driving"
    if waypoints_list:
        waypoints = "%7C".join(_qp(w) for w in waypoints_list)  # '|' encoded as %7C
        url += f"&waypoints={waypoints}"

    if start_navigation:
        # On mobile, this hints to start turn-by-turn navigation.
        url += "&dir_action=navigate"

    return url

def attach_maps_links_both_directions(groups: list[dict], start_navigation: bool = False) -> list[dict]:
    """
    Adds:
      - maps_url        : outbound directions URL (best_route.order)
      - maps_return_url : return directions URL (best_return_route.order)
    """
    for g in groups:
        br = g.get("best_route")
        if br and isinstance(br, dict) and "order" in br and br["order"]:
            g["maps_url"] = build_gmaps_directions_url(br["order"], start_navigation=start_navigation)

        rr = g.get("best_return_route")
        if rr and isinstance(rr, dict) and "order" in rr and rr["order"]:
            g["maps_return_url"] = build_gmaps_directions_url(rr["order"], start_navigation=start_navigation)
    return groups

def attach_maps_links(groups: list[dict], start_navigation: bool = False) -> list[dict]:
    """
    For each group with a computed best_route.order, attach a Google Maps URL.
    """
    for g in groups:
        br = g.get("best_route")
        if br and isinstance(br, dict) and "order" in br and br["order"]:
            g["maps_url"] = build_gmaps_directions_url(br["order"], start_navigation=start_navigation)
    return groups
