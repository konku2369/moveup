"""
Shared inventory analysis functions for multi-store comparison.

Pure functions with no Tk dependency. Used by both mainMultiStore.py
(treeview comparison) and mainAnalytics.py (text dashboard).

Previously, identical copies of these functions lived in both windows.
Now there's a single source of truth here.
"""

from typing import Dict, List, Tuple

import pandas as pd


# ---------------------------------------------------------------------------
# Product map builder
# ---------------------------------------------------------------------------

def build_product_map(df: pd.DataFrame) -> Dict[Tuple[str, str], Dict]:
    """Build a map of (Brand, Product Name) -> {type, rooms, qty} from a DataFrame.

    Aggregates across rows: same product in different rooms gets rooms
    concatenated and qty summed.

    Parameters
    ----------
    df : pd.DataFrame
        Mapped inventory DataFrame (must have Brand, Product Name, Type,
        Room, Qty On Hand columns).

    Returns
    -------
    dict
        Keys are (brand, product_name) tuples. Values are dicts with
        'type', 'rooms' (comma-separated string), and 'qty' (int).
    """
    products: Dict[Tuple[str, str], Dict] = {}
    for _, row in df.iterrows():
        brand = str(row.get("Brand", "")).strip()
        name = str(row.get("Product Name", "")).strip()
        ptype = str(row.get("Type", "")).strip()
        room = str(row.get("Room", "")).strip()
        try:
            qty = int(row.get("Qty On Hand", 0))
        except (ValueError, TypeError):
            qty = 0

        key = (brand, name)
        if key not in products:
            products[key] = {"type": ptype, "rooms": room, "qty": qty}
        else:
            existing = products[key]
            existing["qty"] += qty
            if room and room not in existing["rooms"]:
                existing["rooms"] += f", {room}"

    return products


# ---------------------------------------------------------------------------
# Imbalance detection
# ---------------------------------------------------------------------------

def compute_imbalances(
    both_keys: list,
    a_products: Dict,
    b_products: Dict,
    a_name: str,
    b_name: str,
) -> List[Dict]:
    """Find products at both stores where qty is significantly different.

    Flags when ratio >= 2.0x AND abs diff >= 3 units.

    Returns list of dicts sorted by ratio descending, each containing:
    type, brand, name, qty_a, qty_b, ratio (str like "3.2x"),
    overstocked (store name).
    """
    results = []
    for key in both_keys:
        a_info = a_products[key]
        b_info = b_products[key]
        qa, qb = a_info["qty"], b_info["qty"]

        if qa == 0 and qb == 0:
            continue

        bigger = max(qa, qb)
        smaller = max(min(qa, qb), 1)  # avoid div/0
        ratio = bigger / smaller

        if ratio >= 2.0 and abs(qa - qb) >= 3:
            overstocked = a_name if qa > qb else b_name
            results.append({
                "type": a_info["type"],
                "brand": key[0],
                "name": key[1],
                "qty_a": qa,
                "qty_b": qb,
                "ratio": f"{ratio:.1f}x",
                "overstocked": overstocked,
            })

    results.sort(key=lambda r: float(r["ratio"].rstrip("x")), reverse=True)
    return results


# ---------------------------------------------------------------------------
# Transfer recommendations
# ---------------------------------------------------------------------------

def compute_transfer_recs(
    only_a_keys: list,
    only_b_keys: list,
    both_keys: list,
    a_products: Dict,
    b_products: Dict,
    a_name: str,
    b_name: str,
) -> List[Dict]:
    """Generate prioritized transfer recommendations.

    Priority levels:
      HIGH   - product exists at one store only with qty >= 3
      MEDIUM - product exists at one store only with qty 1-2,
               OR both stores have it but ratio >= 3x with diff >= 5
      LOW    - both stores have it, ratio >= 2x, diff >= 3

    Returns list of dicts sorted by priority, each containing:
    priority, type, brand, name, from, to, qty, reason.
    """
    recs: List[Dict] = []

    # Products only at Store A -> recommend transferring some to B
    for key in only_a_keys:
        info = a_products[key]
        qty = info["qty"]
        if qty <= 0:
            continue
        priority = "High" if qty >= 3 else "Medium"
        recs.append({
            "priority": priority,
            "type": info["type"],
            "brand": key[0],
            "name": key[1],
            "from": a_name,
            "to": b_name,
            "qty": qty,
            "reason": f"Not at {b_name}",
            "_sort": (0 if priority == "High" else 1, -qty),
        })

    # Products only at Store B -> recommend transferring some to A
    for key in only_b_keys:
        info = b_products[key]
        qty = info["qty"]
        if qty <= 0:
            continue
        priority = "High" if qty >= 3 else "Medium"
        recs.append({
            "priority": priority,
            "type": info["type"],
            "brand": key[0],
            "name": key[1],
            "from": b_name,
            "to": a_name,
            "qty": qty,
            "reason": f"Not at {a_name}",
            "_sort": (0 if priority == "High" else 1, -qty),
        })

    # Imbalanced products at both stores
    for key in both_keys:
        a_info = a_products[key]
        b_info = b_products[key]
        qa, qb = a_info["qty"], b_info["qty"]
        if qa == 0 and qb == 0:
            continue

        bigger = max(qa, qb)
        smaller = max(min(qa, qb), 1)
        ratio = bigger / smaller
        diff = abs(qa - qb)

        if ratio >= 2.0 and diff >= 3:
            from_store = a_name if qa > qb else b_name
            to_store = b_name if qa > qb else a_name
            transfer_qty = diff // 2  # suggest splitting the excess

            if ratio >= 3.0 and diff >= 5:
                priority = "Medium"
            else:
                priority = "Low"

            recs.append({
                "priority": priority,
                "type": a_info["type"],
                "brand": key[0],
                "name": key[1],
                "from": from_store,
                "to": to_store,
                "qty": transfer_qty,
                "reason": f"Imbalanced ({ratio:.1f}x)",
                "_sort": (1 if priority == "Medium" else 2, -diff),
            })

    # Sort: High first, then Medium, then Low. Within each, higher qty first.
    recs.sort(key=lambda r: r["_sort"])

    # Remove the sort key before returning
    for r in recs:
        del r["_sort"]

    return recs


# ---------------------------------------------------------------------------
# Category breakdown
# ---------------------------------------------------------------------------

def category_breakdown(products: Dict, field: str) -> Dict[str, Dict]:
    """Group products by 'type' or 'brand' -> {count, qty}.

    Parameters
    ----------
    products : dict
        Product map from build_product_map().
    field : str
        Either "type" or "brand".

    Returns
    -------
    dict
        {category_name: {"count": int, "qty": int}}
    """
    stats: Dict[str, Dict] = {}
    for (brand, _), info in products.items():
        key = info["type"] if field == "type" else brand
        if key not in stats:
            stats[key] = {"count": 0, "qty": 0}
        stats[key]["count"] += 1
        stats[key]["qty"] += info["qty"]
    return stats
