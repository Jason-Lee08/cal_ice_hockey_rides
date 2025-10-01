import os
import itertools
import time
import requests
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv

load_dotenv()

GOOGLE_API_KEY = os.environ.get("GOOGLE_MAPS_API_KEY")

def gmaps_distance_matrix(origins: list[str], destinations: list[str],
                          departure_time: int | None = None,
                          traffic_model: str = "best_guess") -> dict:
    """
    Calls Google Distance Matrix with traffic for driving mode.
    origins/destinations: list of address strings (Google will geocode implicitly).
    departure_time: unix epoch seconds; if None, uses 'now' (must be >= now).
    """
    if not GOOGLE_API_KEY:
        raise RuntimeError("Set GOOGLE_MAPS_API_KEY in your environment.")

    url = "https://maps.googleapis.com/maps/api/distancematrix/json"
    params = {
        "origins": "|".join(origins),
        "destinations": "|".join(destinations),
        "mode": "driving",
        "units": "imperial",            # or 'metric'
        "key": GOOGLE_API_KEY,
        "traffic_model": traffic_model, # best_guess / optimistic / pessimistic
    }

    # Must be a future or 'now' timestamp for traffic. We'll use 'now' safely.
    # Google accepts 'now' only for client-side JS; for REST, pass a unix timestamp >= now.
    if departure_time is None:
        departure_time = int(time.time()) + 60  # 1 minute in the future
    params["departure_time"] = departure_time

    r = requests.get(url, params=params, timeout=20)
    r.raise_for_status()
    js = r.json()
    if js.get("status") != "OK":
        raise RuntimeError(f"Distance Matrix error: {js}")
    return js

def build_duration_matrix_with_traffic(addresses: list[str]) -> list[list[float]]:
    """
    Returns durations[i][j] in seconds between all address pairs using live traffic.
    Small carpool groups fit in one call (limits: up to 25x25 elements).
    """
    dm = gmaps_distance_matrix(addresses, addresses)
    rows = dm.get("rows", [])
    durations = []
    for i, row in enumerate(rows):
        elems = row.get("elements", [])
        row_secs = []
        for j, el in enumerate(elems):
            status = el.get("status")
            if status != "OK":
                # Fallback to distance without traffic if trafficDuration absent
                # (some rare cases return only 'duration')
                dur_obj = el.get("duration_in_traffic") or el.get("duration")
                if not dur_obj:
                    raise RuntimeError(f"No duration for pair {i}->{j}: {el}")
            # Prefer duration_in_traffic when available
            dur_obj = el.get("duration_in_traffic") or el.get("duration")
            row_secs.append(float(dur_obj["value"]))
        durations.append(row_secs)
    return durations

def seconds_to_hms(secs: float) -> str:
    secs = int(round(secs))
    h = secs // 3600
    m = (secs % 3600) // 60
    s = secs % 60
    return (f"{h}h {m}m {s}s" if h else f"{m}m {s}s" if m else f"{s}s")

def compute_best_route_with_traffic(group: dict, final_destination: str) -> dict:
    """
    Brute-force permutations with live traffic:
      start -> pickups (any order) -> final_destination
    Uses Google Distance Matrix (driving + traffic).
    """
    start = group.get("starting_address")
    pickups = group.get("pickup_addresses", []) or []
    if not start:
        return {"error": "missing_starting_address"}

    # Full address list: [start] + pickups + [final]
    addresses = [start] + pickups + [final_destination]
    durations = build_duration_matrix_with_traffic(addresses)

    n = len(addresses)
    start_idx = 0
    final_idx = n - 1
    pickup_indices = list(range(1, final_idx))

    # Edge case: no pickups
    if not pickup_indices:
        total = durations[start_idx][final_idx]
        return {
            "order": [start, final_destination],
            "total_seconds": total,
            "pretty_time": seconds_to_hms(total),
        }

    best_time = float("inf")
    best_seq = None

    for perm in itertools.permutations(pickup_indices):
        t = durations[start_idx][perm[0]]
        for i in range(len(perm) - 1):
            t += durations[perm[i]][perm[i+1]]
        t += durations[perm[-1]][final_idx]
        if t < best_time:
            best_time = t
            best_seq = [start_idx] + list(perm) + [final_idx]

    best_order = [addresses[i] for i in best_seq]
    return {
        "order": best_order,
        "total_seconds": best_time,
        "pretty_time": seconds_to_hms(best_time),
    }

def compute_best_route_with_traffic_addresses(start: str, pickups: list[str], final: str) -> dict:
    """
    Brute-force permutations with live traffic:
      start -> pickups (any order) -> final
    Uses Google Distance Matrix (driving + traffic).

    Returns:
      { "order": [start, ..., final], "total_seconds": float, "pretty_time": str }
      or { "error": "...", ... } on failure
    """
    if not start:
        return {"error": "missing_starting_address"}

    addresses = [start] + (pickups or []) + [final]
    durations = build_duration_matrix_with_traffic(addresses)

    n = len(addresses)
    start_idx = 0
    final_idx = n - 1
    pickup_indices = list(range(1, final_idx))

    # No pickups edge case
    if not pickup_indices:
        total = durations[start_idx][final_idx]
        return {
            "order": [start, final],
            "total_seconds": total,
            "pretty_time": seconds_to_hms(total),
        }

    best_time = float("inf")
    best_seq = None
    for perm in itertools.permutations(pickup_indices):
        t = durations[start_idx][perm[0]]
        for i in range(len(perm) - 1):
            t += durations[perm[i]][perm[i+1]]
        t += durations[perm[-1]][final_idx]
        if t < best_time:
            best_time = t
            best_seq = [start_idx] + list(perm) + [final_idx]

    best_order = [addresses[i] for i in best_seq]
    return {
        "order": best_order,
        "total_seconds": best_time,
        "pretty_time": seconds_to_hms(best_time),
    }

def compute_forward_and_return_routes(groups: list[dict], final_destination: str) -> list[dict]:
    """
    For each group, compute:
      - best_route       : start -> pickups -> final_destination
      - best_return_route: final_destination -> pickups -> start
    """
    for g in groups:
        start = g.get("starting_address")
        pickups = g.get("pickup_addresses", []) or []

        fwd = compute_best_route_with_traffic_addresses(start, pickups, final_destination)
        if "error" in fwd:
            g["best_route_error"] = fwd
        else:
            g["best_route"] = fwd

        # Return: swap start/final, recompute (traffic can be asymmetric)
        ret = compute_best_route_with_traffic_addresses(final_destination, pickups, start)
        if "error" in ret:
            g["best_return_route_error"] = ret
        else:
            g["best_return_route"] = ret

    return groups