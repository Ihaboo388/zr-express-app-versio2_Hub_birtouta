# suivi.py
# ─────────────────────────────────────────────────────────────
# Follow-up board for the confirmation agent.
# For every "Action" order it reads the parcel's ZR state history and
# computes the delivery day:
#   • home delivery → counted from the first "sortie_en_livraison"
#   • StopDesk      → counted from arrival at the destination office
# ZR returns the parcel after 3 days, so day 3 = last chance (red).
# The agent can save the result of her call in the "Suivi" column.
# ─────────────────────────────────────────────────────────────
import re
import time
from datetime import datetime, timedelta
from urllib.parse import quote

import requests

from bulk_import import _get_gspread_client
from status_sync import (_fetch_parcels, _headers, _base, _phone9, _parcel_phones,
                         _is_delivered, _is_return, _state, _parse_date, PENDING_STATE)

SUIVI_HEADER = "Suivi"
STORE_NAME = "BuyDZ"
MAX_DAYS = 3

SITUATION_AR = {
    "appel_telephonique": "اتصل به عامل التوصيل",
    "sms_envoye": "أُرسلت له رسالة SMS",
    "ne_repond_pas_1": "لا يرد على الليفرور (1)",
    "ne_repond_pas_2": "لا يرد على الليفرور (2)",
    "ne_repond_pas_3": "لا يرد على الليفرور (3)",
    "commande_anullee": "الزبون ألغى الطلبية",
    "commande_annulee": "الزبون ألغى الطلبية",
    "en_attente_du_client": "في انتظار الزبون",
    "reporter_a_une_date_ulterieure": "أجّل لتاريخ لاحق",
    "rendez-vous-client": "موعد مبرمج مع الزبون",
    "changement_stopdesk": "تغيير المكتب",
}
URGENT_SITUATIONS = {"ne_repond_pas_1", "ne_repond_pas_2", "ne_repond_pas_3",
                     "commande_anullee", "commande_annulee"}

STATE_AR = {
    "commande_recue": "الطلبية عند ZR (لم تُشحن)",
    "pret_a_expedier": "جاهزة للشحن",
    "confirme_au_bureau": "في مكتب ZR",
    "vers_wilaya": "في الطريق إلى الولاية",
    "en_livraison": "عند مكتب التوصيل",
    "sortie_en_livraison": "خرجت مع الليفرور",
    "attente_recuperation_fournisseur": "راجعة (مرتجع)",
}

CALL_RESULTS = ["أجاب وسيستلم ✅", "لا يرد 📵", "أجّل لموعد آخر 📅", "رفض الطلبية ❌", "رقم خاطئ ☎️"]


def _now_dz():
    return datetime.utcnow() + timedelta(hours=1)


def _ts(value):
    try:
        return datetime.strptime(str(value)[:19].replace("T", " "), "%Y-%m-%d %H:%M:%S") + timedelta(hours=1)
    except Exception:
        return None


def _history(parcel_id):
    r = requests.get(f"{_base()}/parcels/{parcel_id}/state-history", headers=_headers(), timeout=20)
    time.sleep(0.15)
    if r.status_code != 200:
        return []
    data = r.json()
    items = data if isinstance(data, list) else data.get("items", [])
    return sorted(items, key=lambda e: e.get("createdAt") or "")


def _start_of_delivery(history, stopdesk):
    """Algeria-time datetime when the 3-day countdown started, or None."""
    if not stopdesk:
        for e in history:
            if (e.get("newState") or {}).get("name") == "sortie_en_livraison":
                return _ts(e.get("createdAt"))
        return None
    # StopDesk: arrival at the destination office = first "confirme_au_bureau"
    # after the last "vers_wilaya"; same-wilaya parcels: first "en_livraison".
    last_dispatch = None
    for i, e in enumerate(history):
        if (e.get("newState") or {}).get("name") == "vers_wilaya":
            last_dispatch = i
    if last_dispatch is not None:
        for e in history[last_dispatch + 1:]:
            if (e.get("newState") or {}).get("name") in ("confirme_au_bureau", "en_livraison"):
                return _ts(e.get("createdAt"))
        return None
    for e in history:
        if (e.get("newState") or {}).get("name") in ("en_livraison", "sortie_en_livraison"):
            return _ts(e.get("createdAt"))
    return None


def _last_situation(history):
    best = None
    for e in history:
        for s in e.get("situations") or []:
            if not best or (s.get("createdAt") or "") > (best.get("createdAt") or ""):
                best = s
    if not best:
        return "", "", None
    slug = best.get("situationSlug") or ""
    return slug, SITUATION_AR.get(slug, best.get("situationName") or slug), _ts(best.get("createdAt"))


def _wa_number(phone):
    p9 = _phone9(phone)
    return "213" + p9 if p9 else ""


def _wa_message(name, product, price, day, stopdesk, place):
    first = (name or "").split(" ")[0]
    hello = f"السلام عليكم {first} 😊 معاك {STORE_NAME}."
    if stopdesk:
        body = (f"طلبيتك ({product}) وصلت لمكتب ZR Express {place}. "
                f"تقدر تجي تستلمها، الثمن {price} دج.")
    else:
        body = f"طلبيتك ({product}) راهي مع عامل التوصيل، الثمن {price} دج. ردّ بالك على التيليفون باش يوصلهالك 🙏"
    if day >= MAX_DAYS:
        body += " ⚠️ اليوم آخر يوم، إذا ما استلمتهاش اليوم ترجع. وقتاش يناسبك نقولو لليفرور يجيك؟"
    return f"{hello} {body}"


def build_suivi(sheet_name="Branded Store Orders", tab_name="Orders"):
    gc = _get_gspread_client()
    ws = gc.open(sheet_name).worksheet(tab_name)
    data = ws.get_all_values()
    headers = [str(h).strip().lower() for h in data[0]]

    def col(name):
        return headers.index(name) if name in headers else None

    idx = {k: col(v) for k, v in {
        "state": "state", "date": "date & time", "name": "full name", "phone": "phone",
        "city": "city", "commune": "communes", "stopdesk": "stopdesk", "hub": "hubs",
        "product": "product sku", "price": "t.p", "cid": "customer id", "trk": "tracking",
        "suivi": SUIVI_HEADER.lower()}.items()}

    def get(row, key):
        i = idx[key]
        return str(row[i]).strip() if i is not None and len(row) > i else ""

    used_ids = set()
    actions = []
    for r_i, row in enumerate(data[1:], start=2):
        if get(row, "cid"): used_ids.add(get(row, "cid"))
        if get(row, "state").lower() == PENDING_STATE:
            actions.append((r_i, row))
    if not actions:
        return []

    dates = [d for d in (_parse_date(get(r, "date")) for _, r in actions) if d]
    parcels = _fetch_parcels(min(dates) if dates else None)
    by_trk = {p.get("trackingNumber"): p for p in parcels.values()}
    by_phone = {}
    for p in parcels.values():
        for ph in _parcel_phones(p):
            by_phone.setdefault(ph, []).append(p)

    now = _now_dz()
    cards = []
    for r_i, row in actions:
        parcel = parcels.get(get(row, "cid")) or by_trk.get(get(row, "trk"))
        if not parcel:
            cands = [p for p in by_phone.get(_phone9(get(row, "phone")), []) if p["id"] not in used_ids]
            parcel = max(cands, key=lambda p: p.get("createdAt") or "") if cands else None
        stopdesk = (get(row, "stopdesk").upper() == "OUI"
                    or (parcel or {}).get("deliveryType") == "pickup-point")
        card = {
            "row": r_i, "order": row[0], "name": get(row, "name"), "phone": get(row, "phone"),
            "product": get(row, "product"), "price": get(row, "price"),
            "wilaya": get(row, "city"), "commune": get(row, "commune"),
            "stopdesk": stopdesk, "hub": get(row, "hub"), "suivi": get(row, "suivi"),
            "tracking": parcel.get("trackingNumber") if parcel else get(row, "trk"),
            "zr_state": _state(parcel) if parcel else "", "day": 0, "start": None,
            "situation": "", "situation_ar": "", "situation_at": None,
            "level": "transit", "priority": 90,
        }
        if not parcel:
            card.update(level="missing", priority=95)
        elif _is_delivered(parcel) or _is_return(parcel):
            continue  # handled by the "Update Status" button
        else:
            st_name = _state(parcel)
            created = _ts(parcel.get("createdAt"))
            if st_name in ("pret_a_expedier", "commande_recue") and created and (now - created).days >= 3:
                card.update(level="stuck", priority=80)
            if st_name in ("sortie_en_livraison", "en_livraison", "confirme_au_bureau", "vers_wilaya"):
                hist = _history(parcel["id"])
                start = _start_of_delivery(hist, stopdesk)
                slug, sit_ar, sit_at = _last_situation(hist)
                card.update(situation=slug, situation_ar=sit_ar, situation_at=sit_at)
                if start:
                    day = (now.date() - start.date()).days + 1
                    level = {1: "day1", 2: "day2"}.get(day, "day3" if day == MAX_DAYS else "late")
                    prio = {"day1": 40, "day2": 30, "day3": 10, "late": 5}[level]
                    if slug in URGENT_SITUATIONS:
                        prio -= 4
                    card.update(day=day, start=start, level=level, priority=prio)
        card["state_ar"] = STATE_AR.get(card["zr_state"], card["zr_state"])
        place = card["hub"] or card["commune"] or card["wilaya"]
        card["wa"] = (f"https://wa.me/{_wa_number(card['phone'])}?text="
                      + quote(_wa_message(card["name"], card["product"], card["price"],
                                          card["day"], stopdesk, place)))
        card["tel"] = "tel:+" + _wa_number(card["phone"])
        cards.append(card)

    cards.sort(key=lambda c: (c["priority"], -(c["day"] or 0)))
    return cards


def save_call_result(row_index, result, sheet_name="Branded Store Orders", tab_name="Orders"):
    """Writes 'YYYY-MM-DD HH:MM — result' into the Suivi column (creates it if missing)."""
    gc = _get_gspread_client()
    ws = gc.open(sheet_name).worksheet(tab_name)
    header = ws.row_values(1)
    lower = [str(h).strip().lower() for h in header]
    if SUIVI_HEADER.lower() in lower:
        col = lower.index(SUIVI_HEADER.lower()) + 1
    else:
        col = len(header) + 1
        if ws.col_count < col:
            ws.add_cols(col - ws.col_count)
        ws.update_cell(1, col, SUIVI_HEADER)
    stamp = _now_dz().strftime("%Y-%m-%d %H:%M")
    value = f"{stamp} — {result}"
    ws.update_cell(row_index, col, value)
    return value
