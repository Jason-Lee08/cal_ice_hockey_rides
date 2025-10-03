from googleapiclient.errors import HttpError
import sys
from typing import Dict, Tuple, List, Optional
from urllib.parse import quote
from urllib.parse import quote
import os
import shutil

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

def _supports_ansi() -> bool:
    """Best-effort detection for ANSI color support."""
    return sys.stdout.isatty() and (os.name != "nt" or "WT_SESSION" in os.environ or "ANSICON" in os.environ)

def _red(text: str) -> str:
    return f"\033[1;31m{text}\033[0m" if _supports_ansi() else text

def _red_bg(text: str) -> str:
    # white bold on red background for maximum visibility
    return f"\033[1;37;41m{text}\033[0m" if _supports_ansi() else text

def print_error_banner(title: str, details: str, fatal: bool = False) -> None:
    """
    Prints a big, highly visible red banner to stderr.
    If fatal=True, the header says FATAL.
    """
    width = shutil.get_terminal_size(fallback=(100, 20)).columns
    width = max(80, min(width, 140))  # keep it reasonable
    bar = _red_bg(" " * width)
    label = "FATAL" if fatal else "ERROR"
    header = f" {label}: {title} ".center(width, " ")
    header_colored = _red_bg(header)

    lines = []
    for line in details.strip().splitlines():
        # Ensure each content line is padded and colored for readability
        content = f" {line}".ljust(width, " ")
        lines.append(_red_bg(content))

    print(file=sys.stderr)
    print(bar, file=sys.stderr)
    print(header_colored, file=sys.stderr)
    for l in lines:
        print(l, file=sys.stderr)
    print(bar, file=sys.stderr)
    print(file=sys.stderr)

def _yellow_bg(text: str) -> str:
    # black text on yellow background
    return f"\033[30;43m{text}\033[0m" if _supports_ansi() else text

def _cyan_bg(text: str) -> str:
    # black text on cyan background
    return f"\033[30;46m{text}\033[0m" if _supports_ansi() else text

def print_check_results_reminder(path: str) -> None:
    """Print a very visible reminder to check the results.json file."""
    width = shutil.get_terminal_size(fallback=(100, 20)).columns
    width = max(80, min(width, 140))

    bar = _cyan_bg(" " * width)
    header = f" Directions Generated! --> [ACTION REQUIRED] PLEASE CHECK {path} ".center(width, " ")
    header_colored = _cyan_bg(header)

    message = f"""ACTION REQUIRED:
      - Confirm that EVERY driver and passenger you expect
        is included on the ride sheet to ensure everyone gets picked up
    """.strip()

    print(file=sys.stderr)
    print(bar, file=sys.stderr)
    print(header_colored, file=sys.stderr)
    for line in message.splitlines():
        print(_cyan_bg(f" {line}".ljust(width, " ")), file=sys.stderr)
    print(bar, file=sys.stderr)
    print(file=sys.stderr)

def build_name_to_addr_and_display_map(
    service,
    spreadsheet_id: str,
    sheet_name: str = "Full Address & Contact Info",
    start_row: int = 1
) -> Tuple[Dict[str, str], Dict[str, str]]:
    """
    Like build_name_to_address_map, but also returns a map of
    normalized_name -> original display name (exact as in the sheet).
    """
    rows = fetch_values(service, spreadsheet_id, f"{sheet_name}!A:B")
    name_to_addr: Dict[str, str] = {}
    key_to_display: Dict[str, str] = {}

    for i, row in enumerate(rows):
        if i < start_row:
            continue
        if not row:
            continue

        name = row[0].strip() if len(row) >= 1 and row[0] else ""
        addr = row[1].strip() if len(row) >= 2 and row[1] else ""
        if not name or not addr:
            continue

        key = normalize_name(name)
        name_to_addr[key] = addr
        key_to_display[key] = name  # preserve original spelling/casing

    return name_to_addr, key_to_display


def _match_name_to_key(normalized_query: str, name_keys: List[str]) -> Optional[str]:
    """
    Fuzzy-ish match: if normalized_query is contained in a name key,
    or vice-versa, treat it as a match. Returns the matched key or None.
    """
    for k in name_keys:
        if normalized_query in k or k in normalized_query:
            return k
    return None


def find_unassigned_people(
    groups: List[Dict],
    name_to_addr: Dict[str, str],
    key_to_display: Optional[Dict[str, str]] = None
) -> List[Dict[str, str]]:
    """
    Return a list of people from the address sheet who are NOT present
    in any car (neither driver nor passenger).
    Matching uses the same loose substring logic as address attachment.
    Output entries include: {"name": <display or normalized>, "address": <addr>}
    """
    # 1) Collect all names present in car groups (drivers + passengers)
    present_names_norm: List[str] = []
    for g in groups:
        d = g.get("driver")
        if d:
            present_names_norm.append(normalize_name(d))
        for p in g.get("passengers", []) or []:
            present_names_norm.append(normalize_name(p))

    # 2) Resolve each present name to a key in the address map using the same matching strategy
    name_keys = list(name_to_addr.keys())
    assigned_keys: set[str] = set()
    for q in present_names_norm:
        matched = _match_name_to_key(q, name_keys)
        if matched:
            assigned_keys.add(matched)

    # 3) Anything in the address sheet that wasn't assigned is "unassigned"
    unassigned: List[Dict[str, str]] = []
    for key, addr in name_to_addr.items():
        if key not in assigned_keys:
            display = key_to_display.get(key, key) if key_to_display else key
            unassigned.append({"name": display, "address": addr})

    # Sort for nicer output (by name)
    unassigned.sort(key=lambda x: x["name"].lower())
    return unassigned
