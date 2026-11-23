# status_sync.py
# ─────────────────────────────────────────────────────────────
# Update "Action" orders in the Google Sheet from ZR Express.
#   • Delivered  (livre / encaisse / recouvert) → State=Delivered,
#     Time delivered the order = ZR deliveredAt (Algeria time, +0100),
#     Delivery Price, Return Price, Actual Cost = delivery price.
#   • Return (parcel marked isReturn, or back to supplier)  → State=Return,
#     Actual Cost = return price.
#   • Swap: if the customer received another parcel (same phone) that is
#     delivered, the order is marked Delivered with that parcel's ID/tracking.
#   • Only rows whose State is still "Action" are touched.
# Uses the same ZR API key / Google service account as bulk_import.py.
# ─────────────────────────────────────────────────────────────
import re
import time
from datetime import datetime, timedelta

import requests
from gspread.cell import Cell

from bulk_import import _get_gspread_client, ZR_API_KEY, ZR_TENANT_ID, ZR_BASE_URL

DELIVERED_STATES = {"livre", "encaisse", "recouvert"}
RETURNED_STATES = {"recupere_par_fournisseur", "retourne", "retour_fournisseur"}
PENDING_STATE = "action"
MAX_PAGES = 40          # 40 × 100 parcels per direction is plenty
API_DELAY = 0.2

COLS = {
    "state": "state",
    "date": "date & time",
    "phone": "phone",
    "customer id": "customer id",
    "tracking": "tracking",
    "delivery price": "delivery price",
    "return price": "return price",
    "actual cost": "actual cost",
    "time delivered": "time delivered the order",
}


def _headers():
    return {
        "accept": "application/json",
        "Content-Type": "application/json",
        "X-Api-Key": ZR_API_KEY,
        "X-Tenant": ZR_TENANT_ID,
    }


def _base():
    return ZR_BASE_URL.rstrip("/")


def _phone9(value):
    digits = re.sub(r"\D", "", str(value or ""))
    return digits[-9:] if len(digits) >= 9 else ""


def _parcel_phones(parcel):
    cust = parcel.get("customer") or {}
    found = set()
    for m in re.findall(r"\d{9,13}", str(cust)):
        found.add(m[-9:])
    return found


def _to_algeria(ts):
    """ZR timestamps are UTC → 'yyyy-MM-dd HH:mm:ss+0100'."""
    if not ts:
        return ""
    try:
        dt = datetime.strptime(str(ts)[:19].replace("T", " "), "%Y-%m-%d %H:%M:%S")
        return (dt + timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S+0100")
    except Exception:
        return ""


def _state(parcel):
    return str((parcel.get("state") or {}).get("name", "")).lower().strip()


def _is_delivered(parcel):
    return _state(parcel) in DELIVERED_STATES


def _is_return(parcel):
    st = _state(parcel)
    return bool(parcel.get("isReturn")) or st in RETURNED_STATES or "retour" in st


def _parse_date(value):
    try:
        return datetime.strptime(str(value).strip()[:10], "%Y-%m-%d")
    except Exception:
        return None


def _fetch_parcels(oldest_date):
    """All parcels (normal + returns) created after oldest_date - 7 days."""
    limit = (oldest_date - timedelta(days=7)) if oldest_date else None
    url = f"{_base()}/parcels/search"
    parcels = {}
    for is_return in ("false", "true"):
        for page in range(1, MAX_PAGES + 1):
            body = {
                "pageNumber": page,
                "pageSize": 100,
                "orderBy": ["createdAt desc"],
                "advancedFilter": {
                    "filters": [{"field": "isReturn", "operator": "eq", "value": is_return}],
                    "logic": "and",
                },
            }
            r = requests.post(url, headers=_headers(), json=body, timeout=30)
            if r.status_code != 200:
                raise RuntimeError(f"ZR search failed ({r.status_code}): {r.text[:200]}")
            data = r.json()
            items = data.get("items", [])
            for it in items:
                if it.get("id"):
                    parcels[it["id"]] = it
            too_old = False
            if limit and items:
                last = _parse_date(items[-1].get("createdAt"))
                too_old = bool(last and last < limit)
            if not data.get("hasNext") or too_old:
                break
            time.sleep(API_DELAY)
    return parcels


def _detail(parcel_id):
    r = requests.get(f"{_base()}/parcels/{parcel_id}", headers=_headers(), timeout=20)
    time.sleep(API_DELAY)
    return r.json() if r.status_code == 200 else None


def sync_action_orders(sheet_name="Branded Store Orders", tab_name="Orders", apply=False):
    """Returns a dict: {'rows': [...], 'written': int, 'summary': {...}}.
    With apply=False nothing is written (preview)."""
    gc = _get_gspread_client()
    ws = gc.open(sheet_name).worksheet(tab_name)
    data = ws.get_all_values()
    if len(data) <= 1:
        return {"rows": [], "written": 0, "summary": {}}

    headers = [str(h).strip().lower() for h in data[0]]
    missing = [name for name in COLS.values() if name not in headers]
    if missing:
        raise RuntimeError(f"Missing columns in sheet: {missing}")
    c = {k: headers.index(v) for k, v in COLS.items()}
    width = max(c.values()) + 1

    def cell(row, key):
        return str(row[c[key]]).strip() if len(row) > c[key] else ""

    # IDs / trackings already used by other rows (so we never steal a parcel)
    used_by = {}
    action_rows = []
    for i, row in enumerate(data[1:], start=2):
        row = row + [""] * (width - len(row))
        for key in ("customer id", "tracking"):
            v = cell(row, key)
            if v:
                used_by.setdefault(v, set()).add(i)
        if cell(row, "state").lower() == PENDING_STATE:
            action_rows.append((i, row))

    if not action_rows:
        return {"rows": [], "written": 0, "summary": {"action": 0}}

    dates = [d for d in (_parse_date(cell(r, "date")) for _, r in action_rows) if d]
    parcels = _fetch_parcels(min(dates) if dates else None)

    by_tracking = {p.get("trackingNumber"): p for p in parcels.values() if p.get("trackingNumber")}
    by_phone = {}
    for p in parcels.values():
        for ph in _parcel_phones(p):
            by_phone.setdefault(ph, []).append(p)

    results, cells = [], []
    for row_idx, row in action_rows:
        order = row[0]
        cid, trk = cell(row, "customer id"), cell(row, "tracking")
        phone = _phone9(cell(row, "phone"))
        order_date = _parse_date(cell(row, "date"))

        own = parcels.get(cid) or by_tracking.get(trk)

        # other parcels of the same customer that no other sheet row uses
        candidates = []
        for p in by_phone.get(phone, []):
            if own and p["id"] == own["id"]:
                continue
            owners = (used_by.get(p["id"], set()) | used_by.get(p.get("trackingNumber"), set())) - {row_idx}
            if owners:
                continue
            created = _parse_date(p.get("createdAt"))
            swapped = _parse_date(((p.get("swap") or {}).get("swappedAt")))
            if order_date and not ((created and created >= order_date - timedelta(days=1))
                                   or (swapped and swapped >= order_date - timedelta(days=1))):
                continue
            candidates.append(p)
        candidates.sort(key=lambda p: p.get("createdAt") or "", reverse=True)
        delivered_cand = next((p for p in candidates if _is_delivered(p)), None)

        action, target, note = "none", None, ""
        if own and _is_delivered(own):
            action, target = "delivered", own
        elif delivered_cand:
            action, target, note = "delivered", delivered_cand, "swap / other parcel of same customer"
        elif own and _is_return(own):
            action, target = "return", own
        elif not own and candidates:
            action, target, note = "ids", candidates[0], "parcel ID changed in ZR"
        elif not own:
            note = "not found in ZR"
        else:
            note = _state(own)

        update = {}
        when = ""
        if action == "delivered":
            det = _detail(target["id"]) or target
            when = _to_algeria(det.get("deliveredAt") or det.get("lastStateUpdateAt"))
            dp, rp = det.get("deliveryPrice", ""), det.get("returnPrice", "")
            update = {"state": "Delivered", "delivery price": dp, "return price": rp, "actual cost": dp}
            if when and not cell(row, "time delivered"):
                update["time delivered"] = when
        elif action == "return":
            dp, rp = own.get("deliveryPrice", ""), own.get("returnPrice", "")
            update = {"state": "Return", "delivery price": dp, "return price": rp, "actual cost": rp}
        if target is not None:
            if target["id"] != cid:
                update["customer id"] = target["id"]
            if target.get("trackingNumber") and target.get("trackingNumber") != trk:
                update["tracking"] = target["trackingNumber"]

        for key, val in update.items():
            cells.append(Cell(row=row_idx, col=c[key] + 1, value=val))

        results.append({
            "Order": order,
            "ZR status": _state(target or own) if (target or own) else "-",
            "New state": {"delivered": "Delivered", "return": "Return"}.get(action, "Action"),
            "Delivered at": when,
            "Tracking": (target or own or {}).get("trackingNumber", trk),
            "Note": note,
        })

    written = 0
    if apply and cells:
        # re-check that rows are still "Action" right before writing
        fresh = ws.get_all_values()
        still = {i for i, r in enumerate(fresh[1:], start=2)
                 if len(r) > c["state"] and str(r[c["state"]]).strip().lower() == PENDING_STATE}
        cells = [x for x in cells if x.row in still]
        if cells:
            ws.update_cells(cells, value_input_option="USER_ENTERED")
            written = len(cells)

    summary = {
        "action": len(action_rows),
        "delivered": sum(r["New state"] == "Delivered" for r in results),
        "return": sum(r["New state"] == "Return" for r in results),
        "unchanged": sum(r["New state"] == "Action" for r in results),
        "not_found": sum(r["Note"] == "not found in ZR" for r in results),
    }
    return {"rows": results, "written": written, "summary": summary}


if __name__ == "__main__":
    import json, sys
    res = sync_action_orders(apply="--write" in sys.argv)
    print(json.dumps(res["summary"], ensure_ascii=False))
    for r in res["rows"]:
        print(r)
