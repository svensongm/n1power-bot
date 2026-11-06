#!/usr/bin/env python3
"""
N1 POWER bot - cross-border capacity (CBC) and day-ahead power prices to Telegram
================================================================================
* CBC: hourly prices of the daily explicit capacity auctions (SEE + borders around CH).
  Posted as soon as the first results are out (~09:35 CET), then the same pictures are
  updated in place until every border is known; the Excel file follows when complete.
* Prices: hourly day-ahead prices for all markets (like the desk price sheet), a calendar
  for HU, FR and ES with daily base, running month average and weekly averages, and
  calendars of the spreads to HU. Checked from 12:55 CET, posted when HU is out, then
  updated in place until every market is known; the Excel file follows when complete.

Data: Volue Insight API (prices) and ENTSO-E Transparency Platform (capacity auctions,
and prices that Volue does not have). Settings: config.toml. Keys: environment variables.

  python da_cbc_report.py --watch --run-minutes 39   # scheduled mode (GitHub Actions)
  python da_cbc_report.py --now [--date 2026-11-03]  # send everything right away
  python da_cbc_report.py --now --dry-run            # pictures + Excel into ./preview only
  python da_cbc_report.py --find-chat-id             # chats that pressed Start in the bot
"""
from __future__ import annotations

import argparse
import io
import json
import logging
import os
import re
import subprocess
import sys
import threading
import time
import tomllib
import zipfile
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta, timezone
from datetime import time as dtime
from html import escape
from pathlib import Path
from statistics import mean
from zoneinfo import ZoneInfo

import requests

import render as R

log = logging.getLogger("n1power")
TZ = ZoneInfo("Europe/Brussels")                 # CET/CEST market time
UTC = timezone.utc
QH, HOUR, DAY = timedelta(minutes=15), timedelta(hours=1), timedelta(days=1)
BAM_PER_EUR = 1.95583                            # fixed peg of the convertible mark

ENTSOE_URL = os.getenv("ENTSOE_API_URL", "https://web-api.tp.entsoe.eu/api")
VOLUE_AUTH_URL = os.getenv("VOLUE_AUTH_URL", "https://auth.volueinsight.com/oauth2/token")
VOLUE_API_URL = os.getenv("VOLUE_API_URL", "https://api.volueinsight.com")
TELEGRAM_URL = os.getenv("TELEGRAM_API_URL", "https://api.telegram.org")
NBU_URL = os.getenv("NBU_API_URL", "https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange")
HERE = Path(__file__).resolve().parent
STATE_FILE = Path(os.getenv("STATE_FILE", HERE / "state.json"))

# code -> (ENTSO-E EIC, label, Volue area names to try)
AREAS: dict[str, tuple[str, str, list[str]]] = {
    "AL": ("10YAL-KESH-----5", "AL", ["al"]),
    "AT": ("10YAT-APG------L", "AT", ["at"]),
    "BA": ("10YBA-JPCC-----D", "BA", []),
    "BE": ("10YBE----------2", "BE", ["be"]),
    "BG": ("10YCA-BULGARIA-R", "BG", ["bg"]),
    "CH": ("10YCH-SWISSGRIDZ", "CH", ["ch"]),
    "CZ": ("10YCZ-CEPS-----N", "CZ", ["cz"]),
    "DE_LU": ("10Y1001A1001A82H", "DE", ["de"]),
    "ES": ("10YES-REE------0", "ES", ["es"]),
    "FR": ("10YFR-RTE------C", "FR", ["fr"]),
    "GR": ("10YGR-HTSO-----Y", "GR", ["gr"]),
    "HR": ("10YHR-HEP------M", "HR", ["hr"]),
    "HU": ("10YHU-MAVIR----U", "HU", ["hu"]),
    "IT": ("10YIT-GRTN-----B", "IT", ["it"]),          # prices: PUN (Volue); code used for IT capacity
    "IT_NORD": ("10Y1001A1001A73I", "IT-NORD", ["it-nord"]),
    "IT_CNOR": ("10Y1001A1001A70O", "IT-CNOR", ["it-cnor"]),
    "IT_CSUD": ("10Y1001A1001A71M", "IT-CSUD", ["it-csud"]),
    "IT_SUD": ("10Y1001A1001A788", "IT-SUD", ["it-sud"]),
    "IT_CALA": ("10Y1001C--00096J", "IT-CALA", ["it-cala"]),
    "IT_SICI": ("10Y1001A1001A75E", "IT-SICI", ["it-sici"]),
    "IT_SARD": ("10Y1001A1001A74G", "IT-SARD", ["it-sard"]),
    "MD": ("10Y1001A1001A990", "MD", []),
    "ME": ("10YCS-CG-TSO---S", "ME", ["me"]),
    "MK": ("10YMK-MEPSO----8", "MK", ["mk"]),
    "NL": ("10YNL----------L", "NL", ["nl"]),
    "PL": ("10YPL-AREA-----S", "PL", ["pl"]),
    "PT": ("10YPT-REN------W", "PT", ["pt"]),
    "RO": ("10YRO-TEL------P", "RO", ["ro"]),
    "RS": ("10YCS-SERBIATSOV", "RS", ["rs"]),
    "SI": ("10YSI-ELES-----O", "SI", ["si"]),
    "SK": ("10YSK-SEPS-----K", "SK", ["sk"]),
    "TR": ("10YTR-TEIAS----W", "TR", []),
    "UA_IPS": ("10Y1001C--000182", "UA", ["ua", "ua-ips"]),
    "XK": ("10Y1001C--00100H", "XK", []),
}
ALIASES = {"DE": "DE_LU", "UA": "UA_IPS"}


class ReportError(RuntimeError):
    """Stops the run (bad key, bad settings)."""


class SourceError(RuntimeError):
    """One market or border failed; the rest continues."""


# --------------------------------------------------------------------------- small helpers
_SECRETS: list[str] = []


def env(name: str, secret: bool = False) -> str:
    v = (os.getenv(name) or "").strip()
    if secret and len(v) >= 6:
        _SECRETS.append(v)
    return v


def redact(text) -> str:
    text = str(text)
    for s in _SECRETS:
        text = text.replace(s, "***")
    return re.sub(r"(securityToken=)[^&\s]+", r"\1***", text)


def now_local() -> datetime:            # replaced in tests
    return datetime.now(TZ)


def pause(seconds: float) -> None:      # replaced in tests
    time.sleep(seconds)


def day_start(d: date) -> datetime:
    return datetime.combine(d, dtime(0), TZ).astimezone(UTC)


def day_hours(d: date) -> list[datetime]:
    s, e = day_start(d), day_start(d + DAY)
    return [s + i * HOUR for i in range(round((e - s) / HOUR))]


def avg(vals):
    v = [x for x in vals if x is not None]
    return mean(v) if v else None


def norm(code: str) -> str:
    c = str(code).strip().upper().replace("-", "_")
    c = ALIASES.get(c, c)
    if c not in AREAS:
        raise ReportError(f"Unknown area '{code}' in config.toml")
    return c


def label(code: str) -> str:
    return AREAS[code][1]


def blabel(b) -> str:
    return f"{label(b[0])}-{label(b[1])}"            # HU-RS = from HU to RS


def parallel(fn, items, workers: int = 4) -> list:
    if not items:
        return []
    with ThreadPoolExecutor(max_workers=workers) as ex:
        return list(ex.map(fn, items))


# --------------------------------------------------------------------------- ENTSO-E
def _tag(el) -> str:
    return el.tag.rsplit("}", 1)[-1]


def _child(el, name):
    for c in el:
        if _tag(c) == name:
            return c
    return None


def _text(el, name, default=None):
    c = _child(el, name) if el is not None else None
    return c.text.strip() if c is not None and c.text else default


def _dt(s: str) -> datetime:
    s = s.strip()
    return datetime.fromisoformat(s[:-1] + "+00:00" if s.endswith("Z") else s).astimezone(UTC)


_DUR = re.compile(r"^PT(?:(\d+)H)?(?:(\d+)M)?$")


def _res_minutes(res: str) -> int:
    m = _DUR.match(res or "")
    if m and (m.group(1) or m.group(2)):
        return int(m.group(1) or 0) * 60 + int(m.group(2) or 0)
    return {"P1D": 1440, "P7D": 10080}.get(res, 10 ** 6)


def _slots(start, end, res):
    minutes = _res_minutes(res)
    if minutes < 1440:
        step = timedelta(minutes=minutes)
        return [(start + i * step, start + (i + 1) * step) for i in range(round((end - start) / step))]
    if res == "P1D":
        n = max(1, round((end - start) / DAY))
        edges = [start + i * DAY for i in range(n)] + [end]
        return list(zip(edges[:-1], edges[1:]))
    return [(start, end)]


def parse_timeseries(root) -> list[dict]:
    """TimeSeries of an ENTSO-E document as {meta, res, points=[(t0, t1, price, qty)]};
    curve type A03 (repeated values left out) is filled forward."""
    out = []
    for ts in root.iter():
        if _tag(ts) != "TimeSeries":
            continue
        meta = {_tag(c): (c.text or "").strip() for c in ts if len(c) == 0}
        curve = meta.get("curveType", "A01")
        for period in (c for c in ts if _tag(c) == "Period"):
            ti = _child(period, "timeInterval")
            start, end = _dt(_text(ti, "start")), _dt(_text(ti, "end"))
            res = _text(period, "resolution", "PT60M")
            given = {}
            for p in (c for c in period if _tag(c) == "Point"):
                pr, qt = _text(p, "price.amount"), _text(p, "quantity")
                given[int(_text(p, "position"))] = (float(pr) if pr else None, float(qt) if qt else None)
            points, last = [], None
            for i, (t0, t1) in enumerate(_slots(start, end, res), start=1):
                if i in given:
                    last = given[i]
                elif curve != "A03":
                    continue
                if last is not None:
                    points.append((t0, t1, last[0], last[1]))
            out.append({"meta": meta, "res": res, "points": points})
    return out


def _seq(s) -> int:
    try:
        return int(s["meta"].get("classificationSequence_AttributeInstanceComponent.position") or 1)
    except ValueError:
        return 1


def _hourly(d: date, blocks, want_qty=False):
    """(t0, t1, price, qty) blocks -> hourly averages for local day d (first block wins per 15 min)."""
    s, e = day_start(d), day_start(d + DAY)
    nq = round((e - s) / QH)
    pq, qq = [None] * nq, [None] * nq
    for t0, t1, p, q in blocks:
        t = max(t0, s)
        while t < t1 and t < e:
            i = round((t - s) / QH)
            if 0 <= i < nq:
                if pq[i] is None and p is not None:
                    pq[i] = p
                if qq[i] is None and q is not None:
                    qq[i] = q
            t += QH
    return ([avg(pq[i:i + 4]) for i in range(0, nq, 4)], [avg(qq[i:i + 4]) for i in range(0, nq, 4)])


class Entsoe:
    def __init__(self, token: str, http: requests.Session):
        self.token, self.http = token, http

    def query(self, params: dict):
        q = {"securityToken": self.token, **params}
        problem = "no answer"
        for attempt in range(4):
            try:
                r = self.http.get(ENTSOE_URL, params=q, timeout=60)
            except requests.RequestException as e:
                problem = f"network error ({type(e).__name__})"
                pause(4 * (attempt + 1))
                continue
            if r.status_code in (401, 403):
                raise ReportError(f"ENTSO-E refused the API key (HTTP {r.status_code}) - check ENTSOE_API_KEY")
            if r.status_code == 429 or r.status_code >= 500:
                problem = f"HTTP {r.status_code}"
                pause(10 * (attempt + 1))
                continue
            body = r.content
            if body[:2] == b"PK":                       # several documents come zipped: read them all
                with zipfile.ZipFile(io.BytesIO(body)) as z:
                    docs = [z.read(n) for n in z.namelist() if not n.endswith("/")]
            else:
                docs = [body]
            series, reasons, readable = [], [], 0
            for doc in docs:
                try:
                    root = ET.fromstring(doc)
                except ET.ParseError:
                    continue
                readable += 1
                if _tag(root) == "Acknowledgement_MarketDocument":
                    reasons.append(" ".join(e.text.strip() for e in root.iter() if _tag(e) == "text" and e.text))
                else:
                    series += parse_timeseries(root)
            if not readable:
                problem = f"unreadable answer (HTTP {r.status_code})"
                pause(4)
                continue
            if series:
                return series
            reason = " ".join(reasons)
            if "no matching data" in reason.lower() or (not reasons and r.status_code == 200):
                return None
            raise SourceError(f"ENTSO-E: {reason[:160] or f'HTTP {r.status_code}'}")
        raise SourceError(f"ENTSO-E not reachable: {problem}")

    @staticmethod
    def _day_window(d_from: date, d_to: date) -> dict:
        f = "%Y%m%d%H%M"
        return {"periodStart": day_start(d_from).strftime(f), "periodEnd": day_start(d_to).strftime(f)}

    @staticmethod
    def _window(d_from: date, d_to: date) -> dict:
        f = "%Y%m%d%H%M"
        return {"periodStart": (day_start(d_from) - 2 * HOUR).strftime(f),
                "periodEnd": (day_start(d_to) + 2 * HOUR).strftime(f)}

    def day_ahead(self, area: str, d_from: date, d_to: date) -> dict:
        """{day: (hourly prices, currency)} for complete days in [d_from, d_to)."""
        eic = AREAS[area][0]
        base = {"documentType": "A44", "in_Domain": eic, "out_Domain": eic, **self._window(d_from, d_to)}
        try:
            pieces = self.query({**base, "contract_MarketAgreement.type": "A01"})
        except SourceError:
            pieces = None
        if not pieces:
            pieces = self.query(base)
        pieces = [s for s in pieces or [] if s["meta"].get("contract_MarketAgreement.type", "A01") in ("A01", "")
                  and _seq(s) == 1]          # SDAC only: EXAA's 10:15 auction (DE-LU, AT) is sequence 2
        out, d = {}, d_from
        while d < d_to:
            best = None
            for res in sorted({s["res"] for s in pieces}, key=_res_minutes):
                group = sorted((s for s in pieces if s["res"] == res), key=_seq)
                hourly, _ = _hourly(d, [b for s in group for b in s["points"]])
                if all(v is not None for v in hourly):
                    best = (hourly, group[0]["meta"].get("currency_Unit.name") or "EUR")
                    break
            if best:
                out[d] = best
            d += DAY
        return out

    def capacity_range_noted(self, border, d_from: date, d_to: date):
        """Daily explicit auctions for one direction: ({day: (hourly EUR prices, hourly MW, currency)}, note)."""
        params = {"documentType": "A25", "businessType": "B05", "out_Domain": AREAS[border[0]][0],
                  "in_Domain": AREAS[border[1]][0], **self._day_window(d_from, d_to)}
        try:
            pieces = self.query({**params, "contract_MarketAgreement.Type": "A01"})
        except SourceError:
            pieces = self.query({**params, "contract_MarketAgreement.type": "A01"})
        if not pieces:
            return {}, "not on ENTSO-E yet"
        priced = [s for s in pieces
                  if s["meta"].get("contract_MarketAgreement.type", "A01") in ("A01", "")
                  and any(b[2] is not None for b in s["points"])]
        use = [s for s in priced if s["meta"].get("businessType") == "B05"] or priced
        if not use:
            kinds = ", ".join(sorted({s["meta"].get("businessType", "?") for s in pieces}))
            return {}, f"ENTSO-E has it without prices ({kinds})"
        use.sort(key=lambda s: (_res_minutes(s["res"]), _seq(s)))
        cur = use[0]["meta"].get("currency_Unit.name") or "EUR"
        blocks = [b for s in use for b in s["points"]]
        out, d = {}, d_from
        while d < d_to:
            price, mw = _hourly(d, blocks)
            if any(v is not None for v in price):
                if cur == "BAM":
                    price = [None if v is None else v / BAM_PER_EUR for v in price]
                out[d] = (price, mw, cur)
            d += DAY
        return out, ("" if out else "ENTSO-E has it only for other days")

    def capacity_range(self, border, d_from: date, d_to: date) -> dict:
        return self.capacity_range_noted(border, d_from, d_to)[0]

    def capacity(self, border, d: date):
        return self.capacity_range(border, d, d + DAY).get(d)


# --------------------------------------------------------------------------- Volue
class Volue:
    def __init__(self, client_id: str, client_secret: str, http: requests.Session, curves: dict):
        self.cid, self.secret, self.http, self.curves = client_id, client_secret, http, curves
        self._token, self._exp, self._ids = None, 0.0, {}
        self._lock = threading.Lock()
        self.disabled = False

    def _headers(self) -> dict:
        with self._lock:
            if not self._token or time.time() > self._exp - 60:
                r = self.http.post(VOLUE_AUTH_URL, data={"grant_type": "client_credentials"},
                                   auth=(self.cid, self.secret), timeout=30)
                if r.status_code in (400, 401, 403):
                    self.disabled = True
                    raise SourceError("Volue refused the client id / secret")
                r.raise_for_status()
                j = r.json()
                self._token, self._exp = j["access_token"], time.time() + float(j.get("expires_in", 3600))
            return {"Authorization": f"Bearer {self._token}"}

    def _curve_id(self, name: str):
        if name not in self._ids:
            r = self.http.get(f"{VOLUE_API_URL}/api/curves/get", params={"name": name},
                              headers=self._headers(), timeout=30)
            if r.status_code in (400, 404):
                self._ids[name] = None
            else:
                r.raise_for_status()
                j = r.json()
                self._ids[name] = (j[0] if isinstance(j, list) and j else j or {}).get("id")
        return self._ids[name]

    @staticmethod
    def _currency(name: str):
        n = name.lower()
        return "UAH" if "uah" in n else ("EUR" if ("€" in n or "eur" in n) else None)

    def _find(self, area: str):
        """A day-ahead price curve whose name is not the usual one (e.g. UA): search Volue by area."""
        key = ("search", area)
        if key not in self._ids:
            best = None
            for term in dict.fromkeys([label(area)] + [v.upper() for v in AREAS[area][2]]):
                try:
                    r = self.http.get(f"{VOLUE_API_URL}/api/curves", params={"area": term, "category": "PRI",
                                                                             "data_type": "A"},
                                      headers=self._headers(), timeout=30)
                    found = r.json() if r.status_code == 200 else []
                except (requests.RequestException, ValueError):
                    continue
                for c in found if isinstance(found, list) else []:
                    name = str(c.get("name") or "")
                    words = name.lower().split()
                    if "spot" not in words or any(w in ("imb", "id", "intraday", "vwap", "exaa") for w in words):
                        continue
                    cur = self._currency(name)
                    if cur is None or c.get("id") is None:
                        continue
                    rank = (cur != "EUR", str(c.get("frequency", "")).upper() not in ("H", "MIN15"), len(name))
                    if best is None or rank < best[0]:
                        best = (rank, (c["id"], cur, name))
                if best:
                    break
            self._ids[key] = best[1] if best else None
            if best:
                log.info("Volue: %s prices from curve '%s'", label(area), best[1][2])
        return self._ids[key]

    def day_ahead(self, area: str, d_from: date, d_to: date) -> dict:
        names = [self.curves[area]] if area in self.curves else [f"pri {v} spot €/mwh cet h a" for v in AREAS[area][2]]
        target = None
        for name in names:
            cid = self._curve_id(name)
            if cid is not None:
                target = (cid, self._currency(name) or "EUR", name)
                break
        if target is None:
            target = self._find(area)
        if target is None:
            return {}
        cid, cur, _name = target
        r = self.http.get(f"{VOLUE_API_URL}/api/series/{cid}",       # one day extra each side: curves in EET etc.
                          params={"from": (d_from - DAY).isoformat(), "to": (d_to + DAY).isoformat()},
                          headers=self._headers(), timeout=60)
        r.raise_for_status()
        j = r.json() or {}
        step = {"MIN15": QH, "H": HOUR}.get(str(j.get("frequency", "H")).upper(), HOUR)
        blocks = []
        for ms, v in j.get("points") or []:
            if v is not None:
                t = datetime.fromtimestamp(ms / 1000, UTC)
                blocks.append((t, t + step, float(v), None))
        out, d = {}, d_from
        while d < d_to:
            hourly, _ = _hourly(d, blocks)
            if all(v is not None for v in hourly):
                out[d] = (hourly, cur)
            d += DAY
        return out


# --------------------------------------------------------------------------- JAO (optional, free token)
JAO_URL = os.getenv("JAO_API_URL", "https://api.jao.eu/OWSMP")
JAO_CODES = {"DE_LU": ["DE", "D2", "D4", "D7", "D8", "D1"], "UA_IPS": ["UA"]}


def _jao_codes(area: str) -> list[str]:
    return JAO_CODES.get(area, [label(area)])


def _first(d: dict, keys):
    for k in keys:
        v = d.get(k)
        if v not in (None, ""):
            return v
    return None


def _num(v):
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return None


def _utc(v):
    t = datetime.fromisoformat(str(v).strip().replace("Z", "+00:00"))
    return t.replace(tzinfo=UTC) if t.tzinfo is None else t.astimezone(UTC)


def _jao_corridor(a: dict):
    for k in ("corridorCode", "corridor", "corridorName"):
        v = a.get(k)
        if isinstance(v, dict):
            v = v.get("value") or v.get("code") or v.get("name")
        if isinstance(v, str) and "-" in v:
            return v
    parts = str(a.get("identification") or "").split("-")
    return f"{parts[0]}-{parts[1]}" if len(parts) > 2 else None


def _jao_day(a: dict):
    for k in ("marketPeriodStart", "deliveryPeriodStart", "productPeriodStart", "deliveryStart"):
        if a.get(k):
            try:
                return _utc(a[k]).astimezone(TZ).date()
            except ValueError:
                pass
    m = re.search(r"(\d{2})(\d{2})(\d{2})-\d{2}$", str(a.get("identification") or ""))
    if m:
        try:
            return date(2000 + int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def _jao_slot(r: dict, s0: datetime):
    """Sort key of one result row: its start time or start hour (None when the row has neither)."""
    for k, v in r.items():
        kl = k.lower()
        if v in (None, "") or not isinstance(v, (str, int, float)):
            continue
        if "start" in kl or kl in ("from", "periodfrom"):
            try:
                return (_utc(v) - s0).total_seconds() / 60
            except ValueError:
                pass
        if "hour" in kl:
            m = re.match(r"\s*(\d{1,2})(?::(\d{2}))?", str(v))
            if m:
                return int(m.group(1)) * 60 + int(m.group(2) or 0)
    return None


def _jao_hours(a: dict, day: date):
    """(hourly prices, hourly MW, 'EUR') from one daily auction, or None when not (yet) readable."""
    res = [r for r in (a.get("results") or []) if isinstance(r, dict)]
    prods = [p for p in (a.get("products") or []) if isinstance(p, dict)]
    n, s0 = len(day_hours(day)), day_start(day)
    rows = []
    for i, r in enumerate(res):
        p = prods[i] if len(prods) == len(res) else {}
        price = _num(_first(r, ("auctionPrice", "clearingPrice", "marginalPrice", "price")))
        if price is None:
            price = _num(_first(p, ("auctionPrice", "clearingPrice", "price")))
        mw = _num(_first(r, ("allocatedCapacity", "allocated")))
        if mw is None:
            mw = _num(_first(p, ("allocatedCapacity",)))
        slot = _jao_slot(r, s0)
        if slot is None:
            slot = _jao_slot(p, s0)
        rows.append((slot, i, price, mw))
    if not rows:
        return None
    if all(x[0] is not None for x in rows):
        rows.sort(key=lambda x: (x[0], x[1]))
    if len(rows) == 1:
        rows = rows * n                                     # one product for the whole day
    elif len(rows) == 4 * n:                                # quarter-hour products: hourly averages
        rows = [(None, k, avg(x[2] for x in rows[4 * k:4 * k + 4]), avg(x[3] for x in rows[4 * k:4 * k + 4]))
                for k in range(n)]
    if len(rows) != n or any(x[2] is None for x in rows):
        return None
    return [x[2] for x in rows], [x[3] for x in rows], "EUR"


def _jao_shape(a: dict) -> str:
    res = a.get("results") or []
    first = res[0] if res and isinstance(res[0], dict) else {}
    return (f"{a.get('identification')} · auction fields {sorted(a)[:18]} · {len(res)} results, fields "
            f"{sorted(first)[:14]} · first result {json.dumps(first, default=str)[:260]}")


class Jao:
    """JAO Market Data API: published results of the daily explicit auctions. JAO allows one request per second."""

    def __init__(self, token: str, http: requests.Session):
        self.token, self.http = token, http
        self.gap = float(os.getenv("JAO_MIN_GAP", "1.1"))
        self._last, self._lock = 0.0, threading.Lock()
        self.corridors = None
        self.bulk = None               # True: one request answers for all corridors
        self.disabled = False
        self.logged = set()

    def get(self, method: str, params: dict | None = None):
        with self._lock:
            wait = self._last + self.gap - time.monotonic()
            if wait > 0:
                time.sleep(wait)
            try:
                r = self.http.get(f"{JAO_URL}/{method}", params=params or {},
                                  headers={"AUTH_API_KEY": self.token}, timeout=90)
            except requests.RequestException as e:
                raise SourceError(f"JAO not reachable ({type(e).__name__})")
            finally:
                self._last = time.monotonic()
        if r.status_code in (401, 403):
            self.disabled = True
            raise SourceError(f"JAO refused the token (HTTP {r.status_code}) - check the JAO_API_KEY secret")
        if r.status_code >= 400:
            raise SourceError(f"JAO HTTP {r.status_code}")
        try:
            return r.json()
        except ValueError:
            raise SourceError("JAO answer not readable")

    def corridor_list(self) -> list[str]:
        if self.corridors is None:
            data = self.get("getcorridors")
            vals = [(x.get("value") if isinstance(x, dict) else x) for x in (data if isinstance(data, list) else [])]
            self.corridors = [str(v) for v in vals if v]
        return self.corridors

    def corridor(self, b):
        names = {c.upper(): c for c in self.corridor_list()}
        for x in _jao_codes(b[0]):
            for y in _jao_codes(b[1]):
                if f"{x}-{y}" in names:
                    return names[f"{x}-{y}"]
        return None

    def _auctions(self, d_from: date, d_to: date, corridors: list[str]) -> list:
        params = {"horizon": "Daily", "fromdate": (d_from - DAY).isoformat(), "todate": d_to.isoformat()}
        if self.bulk is not False:
            try:
                got = self.get("getauctions", params)
            except SourceError:
                if self.disabled:
                    raise
                got = None
            got = got if isinstance(got, list) else []
            if len({(_jao_corridor(a) or "").upper() for a in got if isinstance(a, dict)}) > 1:
                self.bulk = True
                return got
            self.bulk = False
        out = []
        for c in corridors:
            got = self.get("getauctions", {**params, "corridor": c})
            out += got if isinstance(got, list) else []
        return out

    def daily_range(self, d_from: date, d_to: date, borders) -> dict:
        """{border: {day: (hourly EUR prices, hourly MW, 'EUR')}} for delivery days in [d_from, d_to)."""
        mapping = {}
        for b in borders:
            c = self.corridor(b)
            if c:
                mapping.setdefault(c.upper(), (c, []))[1].append(b)
        if not mapping:
            return {}
        out = {}
        for a in self._auctions(d_from, d_to, [c for c, _ in mapping.values()]):
            if not isinstance(a, dict) or str(a.get("cancelled")).lower() in ("true", "1"):
                continue
            corr = (_jao_corridor(a) or "").upper()
            day = _jao_day(a)
            if corr not in mapping or day is None or not (d_from <= day < d_to):
                continue
            hours = _jao_hours(a, day)
            if hours is None:
                if a.get("results") and any(_first(r, ("auctionPrice",)) is not None for r in a["results"]
                                            if isinstance(r, dict)) and corr not in self.logged:
                    self.logged.add(corr)
                    log.info("JAO: could not read %s - %s", corr, _jao_shape(a))
                continue
            for b in mapping[corr][1]:
                out.setdefault(b, {})[day] = hours
        return out


def test_jao(ctx, d: date) -> int:
    """'test JAO connection': checks the token and shows how JAO's answer looks (no secrets are printed)."""
    j = ctx.jao
    if j is None:
        print("JAO_API_KEY is not set: add it as a GitHub secret and make sure the workflow file passes it on.")
        return 1
    try:
        cors = j.corridor_list()
        print(f"JAO token OK - {len(cors)} corridors: {', '.join(cors)}")
        borders = ctx.cfg["cbc_see"] + ctx.cfg["cbc_ch"]
        on = {blabel(b): j.corridor(b) for b in borders}
        print("Borders read from JAO:", ", ".join(f"{k} ({v})" for k, v in on.items() if v) or "none")
        print("Borders read from ENTSO-E:", ", ".join(k for k, v in on.items() if not v) or "none")
        sample = next((v for v in on.values() if v), None)
        if sample:
            got = j.get("getauctions", {"horizon": "Daily", "corridor": sample,
                                        "fromdate": (d - DAY).isoformat(), "todate": d.isoformat()})
            got = got if isinstance(got, list) else []
            print(f"{sample}: {len(got)} daily auctions found around {d:%d.%m}")
            for a in got[-2:]:
                if isinstance(a, dict):
                    print("  example:", _jao_shape(a), "| delivery day read as", _jao_day(a))
        data = j.daily_range(d, d + DAY, borders)
        print(f"Delivery {d:%d.%m.%Y}: {len(data)} borders read from JAO - {', '.join(sorted(blabel(b) for b in data))}")
        print("One request for all corridors:", "yes" if j.bulk else "no (one request per corridor, 1 per second)")
        for b in sorted(data, key=blabel)[:2]:
            p, mw, _ = data[b][d]
            print(f"  {blabel(b)} hours 1-4: {', '.join(R.fmt(v) for v in p[:4])} EUR/MWh · avg {R.fmt(avg(p))}")
        return 0
    except SourceError as e:
        print(f"JAO: {e}")
        return 1


# --------------------------------------------------------------------------- price store
class Prices:
    """Hourly day-ahead prices per market and day, from the sources in config order."""

    def __init__(self, ctx):
        self.ctx = ctx
        self.cache: dict = {}          # (market, day) -> (hourly, source)
        self._lock = threading.Lock()
        self._fx: dict = {}
        self.converted: dict = {}      # (market, day) -> UAH per EUR used
        self._fx_down = False

    def _uah(self, d: date):
        """Official NBU EUR rate for the day (or the latest one before it)."""
        if d not in self._fx:
            rate = None
            for back in range(5 if not self._fx_down else 0):
                day = d - back * DAY
                try:
                    r = self.ctx.http.get(NBU_URL, params={"valcode": "EUR", "date": day.strftime("%Y%m%d"),
                                                           "json": ""}, timeout=15)
                    j = r.json()
                except Exception:              # NBU not reachable: do not try again in this run
                    self._fx_down = True
                    break
                if j:
                    rate = float(j[0]["rate"])
                    break
            self._fx[d] = rate
        return self._fx[d]

    def get(self, m: str, d_from: date, d_to: date) -> dict:
        days = [d_from + i * DAY for i in range((d_to - d_from).days)]
        for src in self.ctx.cfg["sources"]:
            need = [d for d in days if (m, d) not in self.cache]
            if not need:
                break
            client = self.ctx.volue if src == "volue" else self.ctx.entsoe
            if client is None or (src == "volue" and (client.disabled or not (AREAS[m][2] or m in self.ctx.cfg["volue_curves"]))):
                continue
            try:
                got = client.day_ahead(m, need[0], need[-1] + DAY)
            except (SourceError, requests.RequestException, KeyError, ValueError, TypeError) as e:
                self.ctx.errors[f"{label(m)} {src}"] = redact(e if isinstance(e, SourceError) else type(e).__name__)
                continue
            with self._lock:
                for d, (hourly, cur) in got.items():
                    if d in need:
                        if cur == "UAH":
                            rate = self._uah(d)
                            if not rate:
                                continue
                            hourly = [v / rate for v in hourly]
                            self.converted[(m, d)] = rate
                        elif cur != "EUR":
                            continue
                        self.cache[(m, d)] = (hourly, "Volue" if src == "volue" else "ENTSO-E")
        return {d: self.cache[(m, d)][0] for d in days if (m, d) in self.cache}

    def day(self, m: str, d: date):
        return self.get(m, d, d + DAY).get(d)

    def source(self, m: str, d: date):
        return self.cache.get((m, d), (None, None))[1]


# --------------------------------------------------------------------------- Telegram
class Telegram:
    def __init__(self, token: str, chats: list[str], http: requests.Session):
        self.token, self.chats, self.http = token, chats, http

    def call(self, method: str, **kw) -> dict:
        url = f"{TELEGRAM_URL}/bot{self.token}/{method}"
        problem = "no answer"
        for attempt in range(4):
            try:
                r = self.http.post(url, timeout=60, **kw)
            except requests.RequestException as e:
                problem = f"network error ({type(e).__name__})"
                pause(3 * (attempt + 1))
                continue
            try:
                j = r.json()
            except ValueError:
                j = {}
            if r.status_code == 429:
                pause(float((j.get("parameters") or {}).get("retry_after", 5)) + 1)
                continue
            if r.status_code >= 500:
                problem = f"HTTP {r.status_code}"
                pause(3 * (attempt + 1))
                continue
            if not j.get("ok"):
                raise TelegramRefused(j.get("description") or f"HTTP {r.status_code}", method)
            return j
        raise ReportError(f"Telegram not reachable: {problem}")

    def send_album(self, photos, caption: str) -> list[list[int]]:
        """All pictures as ONE album (one message to forward); the caption shows under it."""
        ids = []
        for chat in self.chats:
            if len(photos) == 1:
                name, png = photos[0]
                j = self.call("sendPhoto", data={"chat_id": chat, "caption": caption, "parse_mode": "HTML"},
                              files={"photo": (f"{name}.png", png, "image/png")})
                ids.append([j["result"]["message_id"]])
                continue
            media, files = [], {}
            for k, (name, png) in enumerate(photos):
                item = {"type": "photo", "media": f"attach://p{k}"}
                if k == 0:
                    item.update(caption=caption, parse_mode="HTML")
                media.append(item)
                files[f"p{k}"] = (f"{name}.png", png, "image/png")
            j = self.call("sendMediaGroup", data={"chat_id": chat, "media": json.dumps(media)}, files=files)
            ids.append([m["message_id"] for m in j["result"]])
        return ids

    @staticmethod
    def _edit(fn) -> None:
        try:
            fn()
        except TelegramRefused as e:
            text = str(e).lower()
            if "not modified" in text:
                return
            if "not found" in text or "can't be edited" in text:
                raise AlbumGone(str(e))
            raise

    def edit_album(self, ids, edits: dict, caption: str) -> None:
        """Replace the changed pictures (by position) and the caption of an album."""
        for c, chat in enumerate(self.chats):
            row = ids[c] if c < len(ids) else []
            if not row:
                raise AlbumGone("album not found")
            for k, (name, png) in sorted(edits.items()):
                media = {"type": "photo", "media": "attach://pic"}
                if k == 0:
                    media.update(caption=caption, parse_mode="HTML")
                self._edit(lambda: self.call("editMessageMedia", data={"chat_id": chat, "message_id": row[k],
                                                                       "media": json.dumps(media)},
                                             files={"pic": (f"{name}.png", png, "image/png")}))
            if 0 not in edits:
                self._edit(lambda: self.call("editMessageCaption", json={"chat_id": chat, "message_id": row[0],
                                                                         "caption": caption, "parse_mode": "HTML"}))

    def delete_album(self, ids) -> None:
        """Remove an album the bot posted earlier (it is replaced by a fuller one)."""
        for c, chat in enumerate(self.chats):
            row = ids[c] if c < len(ids) else []
            if not row:
                continue
            try:
                self.call("deleteMessages", json={"chat_id": chat, "message_ids": row})
            except ReportError:
                for mid in row:
                    try:
                        self.call("deleteMessage", json={"chat_id": chat, "message_id": mid})
                    except ReportError:
                        log.warning("Could not delete the earlier CBC album (message %s)", mid)

    @staticmethod
    def _reply(reply_to, c):
        if reply_to and c < len(reply_to) and reply_to[c]:
            return {"message_id": reply_to[c][0], "allow_sending_without_reply": True}
        return None

    def send_text(self, html: str, reply_to=None, silent: bool = False) -> None:
        for c, chat in enumerate(self.chats):
            body = {"chat_id": chat, "text": html, "parse_mode": "HTML", "disable_web_page_preview": True,
                    "disable_notification": silent}
            if self._reply(reply_to, c):
                body["reply_parameters"] = self._reply(reply_to, c)
            self.call("sendMessage", json=body)

    def send_file(self, name: str, data: bytes, caption: str, reply_to=None, silent: bool = False) -> None:
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        for c, chat in enumerate(self.chats):
            form = {"chat_id": chat, "caption": caption, "parse_mode": "HTML",
                    "disable_notification": "true" if silent else "false"}
            if self._reply(reply_to, c):
                form["reply_parameters"] = json.dumps(self._reply(reply_to, c))
            self.call("sendDocument", data=form, files={"document": (name, data, mime)})


class TelegramRefused(ReportError):
    def __init__(self, description: str, method: str):
        super().__init__(f"Telegram refused {method}: {description}")


class DryRun:
    """Stands in for Telegram: writes the pictures and files to a folder."""

    def __init__(self, folder: Path):
        self.folder, self.n = folder, 0
        folder.mkdir(parents=True, exist_ok=True)

    def send_album(self, photos, caption):
        for name, png in photos:
            (self.folder / f"{name}.png").write_bytes(png)
        print(f"[album {', '.join(n for n, _ in photos)}]\n" + re.sub("<[^>]+>", "", caption))
        self.n += 1
        return [[self.n * 10 + k for k in range(len(photos))]]

    def delete_album(self, ids):
        print(f"[album deleted] {ids}")

    def edit_album(self, ids, edits, caption):
        for k, (name, png) in edits.items():
            (self.folder / f"{name}.png").write_bytes(png)
        print("[album edited]\n" + re.sub("<[^>]+>", "", caption))

    def send_text(self, html, reply_to=None, silent=False):
        print("[message] " + re.sub("<[^>]+>", "", html))

    def send_file(self, name, data, caption, reply_to=None, silent=False):
        (self.folder / name).write_bytes(data)
        print(f"[file {name}] {re.sub('<[^>]+>', '', caption)}")


# --------------------------------------------------------------------------- report pieces
class Ctx:
    def __init__(self, cfg, entsoe, volue, out, http, jao=None):
        self.cfg, self.entsoe, self.volue, self.out, self.http = cfg, entsoe, volue, out, http
        self.jao = jao
        self.errors: dict = {}
        self.notes: dict = {}          # border -> why it is missing
        self.prices = Prices(self)

    def capacity_range(self, border, d_from, d_to) -> dict:
        try:
            data, note = self.entsoe.capacity_range_noted(border, d_from, d_to)
        except SourceError as e:
            self.errors[blabel(border)] = redact(e)
            return {}
        if data:
            self.notes.pop(blabel(border), None)
        elif note:
            self.notes[blabel(border)] = note
        return data

    def capacity(self, border, d):
        return self.capacity_range(border, d, d + DAY).get(d)

    def jao_range(self, d_from, d_to, borders) -> dict:
        if not self.jao or self.jao.disabled or not borders:
            return {}
        try:
            return self.jao.daily_range(d_from, d_to, borders)
        except SourceError as e:
            self.errors["JAO"] = redact(e)
            log.info("JAO: %s - using ENTSO-E", redact(e))
        except (requests.RequestException, ValueError, KeyError, TypeError, AttributeError) as e:
            self.errors["JAO"] = type(e).__name__
            log.info("JAO answer not usable (%s) - using ENTSO-E", type(e).__name__)
        return {}

    def why(self, border) -> str:
        return self.errors.get(blabel(border)) or self.notes.get(blabel(border)) or "not on ENTSO-E yet"


def _peak_idx(d: date, cfg) -> list[int]:
    lo, hi = cfg["peak"]
    return [i for i, h in enumerate(day_hours(d)) if lo <= h.astimezone(TZ).hour < hi]


def _table(d: date, cfg, headers, columns, states, banner, footer) -> dict:
    hours = day_hours(d)
    pk = _peak_idx(d, cfg)
    cols = [c if c else [None] * len(hours) for c in columns]
    return {"date": f"{d:%d/%m}", "weekday": f"{d:%A}", "banner": banner, "headers": headers,
            "states": states, "hours": list(range(1, len(hours) + 1)),
            "rows": [[c[i] for c in cols] for i in range(len(hours))],
            "avg": [avg(c) for c in cols], "peak": [avg(c[i] for i in pk) for c in cols], "footer": footer}


class CbcJob:
    key = "cbc"

    def __init__(self, ctx: Ctx, d: date):
        self.ctx, self.d = ctx, d
        self.groups = [("SEE", ctx.cfg["cbc_see"]), ("CH", ctx.cfg["cbc_ch"])]
        self.groups = [(g, b) for g, b in self.groups if b]
        self.borders = [b for _, bs in self.groups for b in bs]
        self.data: dict = {}
        self.source: dict = {}
        self.inactive: set = set()
        self.checked = False

    def poll(self):
        todo = [b for b in self.borders if b not in self.data]
        for b, days in self.ctx.jao_range(self.d, self.d + DAY, todo).items():      # JAO first
            if self.d in days:
                self.data[b], self.source[b] = days[self.d], "JAO"
        todo = [b for b in todo if b not in self.data]
        for b, r in zip(todo, parallel(lambda b: self.ctx.capacity(b, self.d), todo)):
            if r:
                self.data[b], self.source[b] = r, "ENTSO-E"
        if not self.checked:                    # borders without an auction yesterday count as n/a
            self.checked = True
            todo = [b for b in self.borders if b not in self.data]
            prev_jao = self.ctx.jao_range(self.d - DAY, self.d, todo)
            todo2 = [b for b in todo if self.d - DAY not in prev_jao.get(b, {})]
            prev = parallel(lambda b: self.ctx.capacity(b, self.d - DAY), todo2)
            self.inactive = {b for b, r in zip(todo2, prev) if not r}
        self.inactive -= set(self.data)

    def active(self, borders=None):
        return [b for b in (borders or self.borders) if b in self.data or b not in self.inactive]

    def complete(self):
        return all(b in self.data for b in self.active())

    def ready(self):
        return bool(self.data)

    def signature(self):
        return ",".join(sorted(blabel(b) for b in self.data))

    def table(self, group, borders, final, t) -> dict:
        states, cols, bam = [], [], []
        for b in borders:
            r = self.data.get(b)
            if r:
                states.append("ACTUAL")
                cols.append(r[0])
                if r[2] == "BAM":
                    bam.append(blabel(b))
            else:
                states.append("n/a" if (final or b in self.inactive) else "waiting")
                cols.append(None)
        known = sum(b in self.data for b in borders)
        at = f" at {t:%H:%M}" if t else ""
        banner = f"CBC {group}: {known}/{len(borders)} borders known{at}" + (" · final" if final else "")
        n_jao = sum(self.source.get(b) == "JAO" for b in borders)
        n_ent = sum(self.source.get(b) == "ENTSO-E" for b in borders)
        src = " + ".join(x for x in (f"JAO {n_jao}" if n_jao else "", f"ENTSO-E {n_ent}" if n_ent else "") if x)
        footer = f"Daily explicit capacity auctions, EUR/MWh · source {src or 'JAO / ENTSO-E'} · Peak = hours 9-20"
        if bam:
            footer += f" · {', '.join(bam)} converted from BAM at {BAM_PER_EUR}"
        return _table(self.d, self.ctx.cfg, [blabel(b) for b in borders], cols, states, banner, footer)

    def render(self, final, t):
        out = []
        for group, borders in self.groups:
            tb = self.table(group, borders, final, t)
            known, act = sum(b in self.data for b in borders), len(self.active(borders))
            cap = (f"<b>CBC daily auctions {group} · {self.d:%a %d.%m.%Y}</b>\n"
                   f"{known}/{act} borders known at {t:%H:%M}" + (" · final" if final else " · updating"))
            out.append((f"cbc_{group.lower()}", R.hourly_table_png(tb), cap))
        return out

    def final_caption(self, t):
        return (f"✅ <b>CBC final · {self.d:%a %d.%m.%Y}</b>\n"
                f"{len(self.data)}/{len(self.active())} borders known at {t:%H:%M}")

    def nothing_caption(self):
        return f"CBC {self.d:%a %d.%m.%Y}: no capacity auction results on ENTSO-E by {self.ctx.cfg['cbc_final']:%H:%M}."


class PriceJob:
    key = "da"

    def __init__(self, ctx: Ctx, d: date):
        self.ctx, self.d = ctx, d
        c = ctx.cfg
        self.markets = c["markets"]
        self.data: dict = {}
        self.inactive: set = set()
        self.checked = False

    def start(self, m):
        """When this market is first checked (its usual publication time minus a small margin)."""
        return self.ctx.cfg["price_check_from"].get(m, self.ctx.cfg["price_check_from_default"])

    def waiting(self):
        return [m for m in self.markets if m not in self.data and m not in self.inactive]

    def poll(self, t=None):
        todo = [m for m in self.markets if m not in self.data and (t is None or t.time() >= self.start(m))]
        for m, r in zip(todo, parallel(lambda m: self.ctx.prices.day(m, self.d), todo)):
            if r:
                self.data[m] = r
        if not self.checked:
            self.checked = True
            todo = [m for m in self.markets if m not in self.data]
            prev = parallel(lambda m: self.ctx.prices.day(m, self.d - DAY), todo)
            self.inactive = {m for m, r in zip(todo, prev) if not r}
        self.inactive -= set(self.data)

    def active(self):
        return [m for m in self.markets if m in self.data or m not in self.inactive]

    def complete(self):
        return all(m in self.data for m in self.active())

    def ready(self):
        first = self.ctx.cfg["prices_first"]
        return first in self.data or (first in self.inactive and bool(self.data))

    def signature(self):
        return ",".join(sorted(self.data))

    # ---- hourly table
    def table(self, final, t) -> dict:
        states, cols = [], []
        for m in self.markets:
            if m in self.data:
                states.append("ACTUAL")
                cols.append(self.data[m])
            else:
                states.append("n/a" if (final or m in self.inactive) else "waiting")
                cols.append(None)
        at = f" at {t:%H:%M}" if t else ""
        banner = f"{len(self.data)}/{len(self.markets)} markets ACTUAL{at}" + (" · final" if final else "")
        if not self.data and not final:
            banner = "waiting for day-ahead results · " + self.ctx.cfg["publish_text"]
        via_entsoe = [label(m) for m in self.data if self.ctx.prices.source(m, self.d) == "ENTSO-E"]
        footer = "EUR/MWh · Peak = hours 9-20 · IT = PUN · source Volue" + \
                 (f" (ENTSO-E: {', '.join(via_entsoe)})" if via_entsoe else "")
        if self.ctx.volue is None or self.ctx.volue.disabled:
            footer += " · Volue not available - ENTSO-E used, no PUN"
        fx = [(m, self.ctx.prices.converted[(m, self.d)]) for m in self.data if (m, self.d) in self.ctx.prices.converted]
        if fx:
            footer += " · " + ", ".join(f"{label(m)} from UAH at NBU {R.fmt(r, 4)}" for m, r in fx)
        return _table(self.d, self.ctx.cfg, [label(m) for m in self.markets], cols, states, banner, footer)

    # ---- calendars
    def _stats(self, markets) -> dict:
        """{market: {day: (base, peak)}} from the Monday before the 1st of the month to the delivery day."""
        first = self.d.replace(day=1)
        start = first - timedelta(days=first.weekday())
        out = {}
        for m in markets:
            hist = dict(self.ctx.prices.get(m, start, self.d)) if start < self.d else {}
            if m in self.data:
                hist[self.d] = self.data[m]
            out[m] = {d: (mean(h), avg(h[i] for i in _peak_idx(d, self.ctx.cfg))) for d, h in hist.items()}
        return out

    def _weeks(self):
        d = self.d
        first = d.replace(day=1)
        last = (first + timedelta(days=32)).replace(day=1) - DAY
        start = first - timedelta(days=first.weekday())
        end = last + timedelta(days=6 - last.weekday())
        weeks = [[start + timedelta(days=7 * k + j) for j in range(7)] for k in range(((end - start).days + 1) // 7)]
        return first, last, weeks

    @staticmethod
    def _running(series, first, upto):
        run, acc, x = {}, [], first
        while x <= upto:
            if x in series:
                acc.append(series[x])
                run[x] = mean(acc)
            x += DAY
        return run

    def _block(self, title, stats, spread=None) -> dict:
        """One country: B (base), P (peak 08-20) and MTD (month average of base to that day) per day,
        Mon-Sun B and P per week; spread = (name, {day: value}) adds the spread to HU in small print."""
        d = self.d
        first, last, weeks = self._weeks()
        base = {x: v[0] for x, v in stats.items()}
        peak = {x: v[1] for x, v in stats.items() if v[1] is not None}
        run_b, run_p = self._running(base, first, min(d, last)), self._running(peak, first, min(d, last))
        cells = [[{"day": x.day, "in_month": first <= x <= last, "base": base.get(x) if x <= d else None,
                   "peak": peak.get(x) if x <= d else None, "mtd": run_b.get(x), "target": x == d}
                  for x in wk] for wk in weeks]
        mtd, prev = run_b.get(d), run_b.get(d - DAY) if d > first else None
        footer = f"{d:%B} MTD  B {R.fmt(mtd) or '-'}  P {R.fmt(run_p.get(d)) or '-'}"
        if mtd is not None and prev is not None:
            footer += f"   MTD vs previous day {R.signed(mtd - prev)}"
        block = {"title": title, "kind": "price", "weeks": cells, "mtd": mtd, "mtd_p": run_p.get(d),
                 "dod": (mtd - prev) if (mtd is not None and prev is not None) else None,
                 "code": title.split(" ")[0], "month": f"{d:%B}",
                 "week_base": [avg(base.get(x) for x in wk if x <= d) for wk in weeks],
                 "week_peak": [avg(peak.get(x) for x in wk if x <= d) for wk in weeks],
                 "week_partial": [any(x > d or x not in base for x in wk) for wk in weeks], "footer": footer}
        if spread:
            name, ser = spread
            for wk, days in zip(cells, weeks):
                for cell, x in zip(wk, days):
                    cell["spread"] = ser.get(x) if x <= d else None
            srun = self._running(ser, first, min(d, last))
            block.update(spread_name=name, week_spread=[avg(ser.get(x) for x in wk if x <= d) for wk in weeks],
                         spread_mtd=srun.get(d))
            if srun.get(d) is not None:
                block["footer"] += f"   {name} {R.signed(srun[d])}"
        return block

    def _spread(self, stats, m):
        c = self.ctx.cfg
        hu = c["spread_base"]
        a, b = (m, hu) if m in c["country_first"] else (hu, m)
        return f"{label(a)}{label(b)}", {d: stats[a][d][0] - stats[b][d][0] for d in stats[a] if d in stats[b]}

    def calendars(self):
        """Country calendars with full prices (+ small spread to HU) and the week recap of the spreads."""
        c = self.ctx.cfg
        stats = self._stats(list(dict.fromkeys(c["calendar_countries"] + [c["spread_base"]] + c["spreads"])))
        blocks = [self._block(f"{label(m)} · day-ahead", stats[m], self._spread(stats, m) if m in c["spreads"] else None)
                  for m in c["calendar_countries"]]
        if c["spreads"]:
            first, last, weeks = self._weeks()
            weeks = [wk for wk in weeks if wk[0] <= self.d]
            cols = [(f"W{wk[0].isocalendar()[1]}", f"{wk[0]:%d.%m}" + ("*" if wk[-1] > self.d else "")) for wk in weeks]
            rows = []
            for m in c["spreads"]:
                name, ser = self._spread(stats, m)
                pair = (m, c["spread_base"]) if m in c["country_first"] else (c["spread_base"], m)
                rows.append((name, [avg(ser.get(x) for x in wk if x <= self.d) for wk in weeks]
                             + [avg(v for x, v in ser.items() if first <= x <= self.d)],
                             (label(pair[0]), label(pair[1]))))
            blocks.append({"title": f"Week recap · spreads to {label(c['spread_base'])} (base)", "kind": "recap",
                           "code": label(c["spread_base"]),
                           "columns": cols + [(f"{self.d:%b}", "MTD")], "rows": rows,
                           "footer": "Mon-Sun averages · * week so far · HUxx = HU-xx, DEHU/GRHU/ATHU = country-HU"})
        return blocks

    def render(self, final, t):
        tb = self.table(final, t)
        hu = self.ctx.cfg["spread_base"]
        head = (f"<b>DA prices · {self.d:%a %d.%m.%Y}</b>\n{len(self.data)}/{len(self.active())} markets ACTUAL "
                f"at {t:%H:%M}" + (" · final" if final else " · updating"))
        blocks = self.calendars()
        if hu in self.data:
            i = self.markets.index(hu)
            mtd = next((b.get("mtd") for b in blocks if b["kind"] == "price" and b["title"].startswith(label(hu) + " ")), None)
            head += f"\n{label(hu)} B {R.fmt(tb['avg'][i])} · P {R.fmt(tb['peak'][i])} · {self.d:%B} MTD {R.fmt(mtd) or '-'}"
        cut = self.ctx.cfg["calendar_split_after"]
        countries = self.ctx.cfg["calendar_countries"]
        n1 = countries.index(cut) + 1 if cut in countries else len(blocks)
        pages = [p for p in (blocks[:n1], blocks[n1:]) if p]
        legend = "B = base · P = peak 08-20 · MTD = month average of base to that day"
        out = [("prices_table", R.hourly_table_png(tb), head)]
        for k, page in enumerate(pages, start=1):
            names = [b["title"].split(" ")[0] for b in page if b["kind"] == "price"]
            span = f"{names[0]}–{names[-1]}" if names else ""
            recap = " + spread recap" if any(b["kind"] == "recap" for b in page) else ""
            title = f"Day-ahead calendar {k}/{len(pages)} · {self.d:%B %Y} · {span}{recap} · to {self.d:%a %d.%m.%Y}"
            cap = (f"<b>Calendar {k}/{len(pages)} · {self.d:%B %Y} · {span}{recap}</b>\n{legend}\n"
                   "Week = Mon-Sun, spread to HU in blue (HUxx = HU − xx; DEHU, GRHU, ATHU = country − HU)")
            out.append((f"prices_calendar_{k}", R.calendars_png(title, page, 3, legend + " · Week = Mon-Sun average, "
                                                                 "blue = spread to HU · framed day = delivery day"), cap))
        return out

    def final_caption(self, t):
        return (f"✅ <b>DA prices final · {self.d:%a %d.%m.%Y}</b> · {len(self.data)}/{len(self.active())} markets "
                f"at {t:%H:%M}\n📊 {self.d:%B %Y} file updated (prices + CBC)")

    def nothing_caption(self):
        return f"DA prices {self.d:%a %d.%m.%Y}: nothing published by {self.ctx.cfg['prices_final']:%H:%M}."


def month_workbook(ctx, d: date, t, price_job=None) -> bytes:
    """The month's file, two sheets: Prices (calendars + one hourly table per day, newest first)
    and CBC (daily average per border + one hourly table per day, newest first)."""
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    first = d.replace(day=1)
    days = [first + i * DAY for i in range((d - first).days + 1)]
    pj = price_job or PriceJob(ctx, d)
    if price_job is None:
        pj.poll()
    hist = {m: ctx.prices.get(m, first, d + DAY) for m in pj.markets}
    tables = []
    for day in reversed(days):
        if day == d:
            tables.append(pj.table(pj.complete(), t))
            continue
        j = PriceJob(ctx, day)
        j.data = {m: hist[m][day] for m in pj.markets if day in hist[m]}
        tables.append(j.table(True, None))
    cbc = CbcJob(ctx, d)
    cap = dict(zip(cbc.borders, parallel(lambda b: ctx.capacity_range(b, first, d + DAY), cbc.borders)))
    for b, days in ctx.jao_range(first, d + DAY, cbc.borders).items():         # JAO first where it has the border
        cap[b] = {**cap.get(b, {}), **days}
    daily = [(f"{day:%a %d.%m}", [avg(cap[b][day][0]) if day in cap[b] else None for b in cbc.borders], False)
             for day in days]
    daily.append((f"{d:%B} avg", [avg(vals[k] for _, vals, _b in daily) for k in range(len(cbc.borders))], True))
    cbc_tables = []
    for day in reversed(days):
        j = CbcJob(ctx, day)
        j.data = {b: cap[b][day] for b in j.borders if day in cap[b]}
        cbc_tables.append([j.table(g, bs, True, None) for g, bs in j.groups])
    blocks = pj.calendars()

    def build(wb):
        ws = wb.active
        ws.title = f"Prices {d:%b %Y}"
        ws.cell(1, 1, f"N1 POWER · day-ahead prices · {d:%B %Y} · updated {t:%a %d.%m %H:%M}").font = Font(bold=True, size=14)
        ws.cell(2, 1, "B = base · P = peak 08-20 · MTD = month average of base to that day · Week = Mon-Sun average, "
                      "blue = spread to HU · framed day = latest delivery day. Below: hourly prices, newest day first. "
                      "EUR/MWh.").font = Font(size=9, color="64748B")
        ws.sheet_view.showGridLines = False
        row = 4 + R.xl_cards(ws, 4, blocks, 3) + 1
        for tb in tables:
            row += R.xl_hourly(ws, row, 1, tb) + 1
        for col in range(1, 30):
            ws.column_dimensions[get_column_letter(col)].width = 9.5
        ws2 = wb.create_sheet(f"CBC {d:%b %Y}")
        ws2.cell(1, 1, f"N1 POWER · daily explicit capacity auctions · {d:%B %Y} · updated {t:%a %d.%m %H:%M}").font = \
            Font(bold=True, size=14)
        ws2.cell(2, 1, "EUR/MWh (BA-RS, RS-BA converted from BAM at 1.95583). Top: daily average per border. "
                       "Below: hourly prices, newest day first (SEE left, CH right).")
        ws2.sheet_view.showGridLines = False
        row = 4 + R.xl_matrix(ws2, 4, 1, "Daily average price", "Day", [blabel(b) for b in cbc.borders], daily) + 1
        width_see = len(dict(cbc.groups).get("SEE", [])) + 2
        for pair in cbc_tables:
            used = 0
            for k, tb in enumerate(pair):
                used = max(used, R.xl_hourly(ws2, row, 1 + k * width_see, tb))
            row += used + 1
        for col in range(1, 50):
            ws2.column_dimensions[get_column_letter(col)].width = 8.5
        for sheet in (ws, ws2):                       # print: landscape, all columns on one page width
            sheet.page_setup.orientation = "landscape"
            sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
            sheet.page_setup.fitToWidth, sheet.page_setup.fitToHeight = 1, 0
            sheet.sheet_properties.pageSetUpPr.fitToPage = True
            sheet.print_options.horizontalCentered = True
            sheet.page_margins.left = sheet.page_margins.right = 0.3
    return R.workbook_bytes(build)


# --------------------------------------------------------------------------- state (kept in the repository)
def load_state() -> dict:
    try:
        return json.loads(STATE_FILE.read_text())
    except (OSError, ValueError):
        return {}


def save_state(state: dict, message: str | None = None) -> None:
    keep = sorted(k for k in state if k[:1].isdigit())[-10:] + [k for k in state if not k[:1].isdigit()]
    STATE_FILE.write_text(json.dumps({k: state[k] for k in keep}, indent=1, sort_keys=True))
    if not message or os.getenv("GITHUB_ACTIONS") != "true":
        return
    try:
        git = lambda *a: subprocess.run(["git", *a], cwd=STATE_FILE.parent, capture_output=True, text=True, timeout=60)
        git("add", STATE_FILE.name)
        if git("diff", "--cached", "--quiet").returncode == 0:
            return
        git("commit", "-q", "-m", message)
        for _ in range(3):
            git("pull", "-q", "--rebase", "--autostash")
            if git("push", "-q").returncode == 0:
                return
            pause(5)
        log.warning("Could not save the state file to the repository")
    except Exception as e:  # never stop the bot because of this
        log.warning("Could not save the state file: %s", type(e).__name__)


# --------------------------------------------------------------------------- one album per delivery day
class AlbumGone(ReportError):
    """The album was deleted in Telegram: a new one is posted."""


def log_missing(job, ctx) -> None:
    missing = [b for b in job.borders if b not in job.data]
    text = "; ".join(f"{blabel(b)}: {'no auction yesterday either' if b in job.inactive else ctx.why(b)}"
                     for b in missing)
    if text != getattr(job, "_logged", None):
        job._logged = text
        log.info("CBC %s - %d/%d known%s", job.d, len(job.data), len(job.borders),
                 f". Missing: {text}" if text else "")


def _listing(names: list[str], waiting: list[str]) -> str:
    if len(names) <= 8:
        return ", ".join(names)
    return f"{len(names)} new" + (f" · still waiting: {', '.join(waiting)}" if waiting else " · all in")


class DayReport:
    """Everything for one delivery day in ONE Telegram album (CBC SEE, CBC CH, DA prices table, calendars)
    with the counts in the caption; new data edits the album and a short reply says what came in."""

    def __init__(self, ctx, d: date):
        self.ctx, self.d, self.cfg = ctx, d, ctx.cfg
        self.cbc, self.da = CbcJob(ctx, d), PriceJob(ctx, d)

    def cbc_done(self, t) -> bool:
        return (self.cbc.checked and self.cbc.complete()) or t.time() >= self.cfg["cbc_final"]

    def da_done(self, t) -> bool:
        return (self.da.checked and self.da.complete()) or t.time() >= self.cfg["prices_final"]

    def final(self, t) -> bool:
        return self.cbc_done(t) and self.da_done(t)

    def cbc_due(self, t) -> bool:
        return t.time() >= self.cfg["cbc_start"] and not self.cbc_done(t)

    def da_due(self, t) -> bool:
        return not self.da_done(t) and any(t.time() >= self.da.start(m) for m in self.da.waiting())

    def due(self, t) -> bool:
        return self.cbc_due(t) or self.da_due(t)

    def next_check(self, t):
        """The next time something starts to be due (None when nothing more is coming today)."""
        times = []
        if not self.cbc_done(t) and t.time() < self.cfg["cbc_start"]:
            times.append(self.cfg["cbc_start"])
        if not self.da_done(t):
            times += [self.da.start(m) for m in self.da.waiting() if self.da.start(m) > t.time()]
        return datetime.combine(t.date(), min(times), TZ) if times else None

    def poll(self, t, everything: bool = False) -> None:
        if everything or self.cbc_due(t):
            self.cbc.poll()
            log_missing(self.cbc, self.ctx)
        if everything or self.da_due(t):
            self.da.poll(None if everything else t)

    def known(self):
        return sorted(blabel(b) for b in self.cbc.data), sorted(label(m) for m in self.da.data)

    def counts(self):
        out = [(f"CBC {g}", sum(b in self.cbc.data for b in bs), len(bs)) for g, bs in self.cbc.groups]
        out.append(("DA prices", len(self.da.data), len(self.da.markets)))
        cal = self.cfg["calendar_countries"]
        out.append(("Calendar", sum(m in self.da.data for m in cal), len(cal)))
        return out

    def missing(self) -> list[str]:
        return ([blabel(b) for b in self.cbc.borders if b not in self.cbc.data]
                + [label(m) for m in self.da.markets if m not in self.da.data])

    def caption(self, t, final: bool, snapshot: bool = False, with_prices: bool = True) -> str:
        lines = [f"⚡ <b>Delivery {self.d:%a %d.%m.%Y}</b>"]
        counts = self.counts()
        if not with_prices:
            counts = [c for c in counts if c[0].startswith("CBC")]
        lines += [f"• {name} {k}/{n}" for name, k, n in counts]
        if not with_prices and not final:
            lines.append(f"• DA prices and calendar follow from ~{self.cfg['first_publish']:%H:%M}")
        if snapshot:
            status = f"snapshot {t:%H:%M} · the automatic post updates itself"
        elif final:
            status = "✅ complete" if not self.missing() else f"final {t:%H:%M}"
        else:
            status = f"updated {t:%H:%M} · still filling in"
        lines.append(f"🕑 <i>{status}</i>")
        if (final or snapshot) and self.missing():
            lines.append(f"<i>n/a: {escape(', '.join(self.missing()))}</i>"[:600])
        return "\n".join(lines)

    def new_spreads(self, old_da: set) -> list[str]:
        c, hu = self.cfg, self.cfg["spread_base"]
        out = []
        for m in c["spreads"]:
            if m in self.da.data and hu in self.da.data and not (label(m) in old_da and label(hu) in old_da):
                a, b = (m, hu) if m in c["country_first"] else (hu, m)
                out.append(f"{label(a)}{label(b)}")
        return out

    def note(self, t, old_cbc: set, old_da: set):
        """What came in since the last post, e.g. 'CBC SEE: HU-RS' / 'UA price' / 'Spreads: GRHU'."""
        lines = [f"🔄 <b>Updated {t:%H:%M}</b> · delivery {self.d:%a %d.%m}"]
        for g, bs in self.cbc.groups:
            names = [blabel(b) for b in bs if b in self.cbc.data and blabel(b) not in old_cbc]
            if names:
                waiting = [blabel(b) for b in self.cbc.active(bs) if b not in self.cbc.data]
                lines.append(f"• CBC {g}: {_listing(names, waiting)}")
        new_da = [label(m) for m in self.da.markets if m in self.da.data and label(m) not in old_da]
        if new_da:
            waiting = [label(m) for m in self.da.active() if m not in self.da.data]
            lines.append(f"• {new_da[0]} price" if len(new_da) == 1 else f"• DA prices: {_listing(new_da, waiting)}")
            spreads = self.new_spreads(old_da)
            if spreads:
                lines.append(f"• Spreads: {', '.join(spreads) if len(spreads) <= 8 else f'{len(spreads)} updated'}")
        return "\n".join(lines) if len(lines) > 1 else None

    def photos(self, t, which=("cbc", "da")) -> list[tuple[str, bytes]]:
        out = []
        if "cbc" in which:
            out += [(n, png) for n, png, _ in self.cbc.render(self.cbc_done(t), t)]
        if "da" in which:
            out += [(n, png) for n, png, _ in self.da.render(self.da_done(t), t)]
        return out


def fingerprints(rep: DayReport) -> dict:
    fp = {f"cbc:{blabel(b)}": round(sum(v for v in r[0] if v is not None), 2) for b, r in rep.cbc.data.items()}
    fp.update({f"da:{label(m)}": round(sum(v for v in h if v is not None), 2) for m, h in rep.da.data.items()})
    return fp


def step(rep: DayReport, st: dict, tg, state: dict) -> None:
    t = now_local()
    rep.poll(t)
    cbc_now, da_now = rep.known()
    final = rep.final(t)
    if not st.get("ids"):
        if cbc_now or da_now:                                       # first data of the day: post the album
            full = bool(da_now)                                     # no price yet: CBC pictures only
            photos = rep.photos(t, ("cbc", "da") if full else ("cbc",))
            st["ids"] = tg.send_album(photos, rep.caption(t, final, with_prices=full))
            st["fp"] = fingerprints(rep)
            st.update(names=[n for n, _ in photos], cbc=cbc_now, da=da_now, posted=f"{t:%H:%M}", final=final)
            log.info("Album posted for %s: %s", rep.d, ", ".join(f"{n} {k}/{m}" for n, k, m in rep.counts()))
            save_state(state, f"{rep.d} posted")
        elif final:
            tg.send_text(escape(f"Delivery {rep.d:%a %d.%m.%Y}: nothing was published by "
                                f"{max(rep.cfg['cbc_final'], rep.cfg['prices_final']):%H:%M}."))
            st.update(final=True, excel=True)
            save_state(state, f"{rep.d} nothing published")
    else:
        old_cbc = {str(x).replace(">", "-") for x in st.get("cbc", [])}
        old_da = set(st.get("da", []))
        lost = (old_cbc - set(cbc_now)) | (old_da - set(da_now))     # a source hiccup: wait for the next check
        if lost and not final:
            log.info("Not updating yet, these did not answer this time: %s", ", ".join(sorted(lost)))
            return
        changed = {k for k, old, now in (("cbc", old_cbc, cbc_now), ("da", old_da, da_now)) if set(now) - old}
        if final and not st.get("final"):
            changed |= {"cbc", "da"}                                  # waiting -> n/a, caption 'final'
        fp_now, fp_old = fingerprints(rep), st.get("fp")
        corrected = [] if fp_old is None else sorted(k for k, v in fp_now.items()
                                                     if k in fp_old and abs(fp_old[k] - v) > 0.005)
        if fp_old is None:                                            # first run of this version: redraw once
            changed |= {"cbc", "da"}
        changed |= {k.split(":")[0] for k in corrected}
        has_prices = any(n.startswith("prices") for n in st.get("names", []))
        if da_now and not has_prices:
            # first price result of the day: post the full album and remove the morning CBC-only album
            old_ids = st["ids"]
            photos = rep.photos(t)
            st["ids"] = tg.send_album(photos, rep.caption(t, final))
            st["names"] = [n for n, _ in photos]
            st["fp"] = fp_now
            tg.delete_album(old_ids)
            note = rep.note(t, old_cbc, old_da)
            if note:
                tg.send_text(note, reply_to=st["ids"])
            st.update(cbc=sorted(old_cbc | set(cbc_now)), da=sorted(old_da | set(da_now)), final=final)
            log.info("Full album posted for %s (CBC-only album replaced): %s", rep.d,
                     ", ".join(f"{n} {k}/{m}" for n, k, m in rep.counts()))
            save_state(state, f"{rep.d} full album")
        elif changed:
            caption = rep.caption(t, final, with_prices=has_prices)
            photos = rep.photos(t, changed)
            edits = {st["names"].index(n): (n, png) for n, png in photos if n in st.get("names", [])}
            try:
                tg.edit_album(st["ids"], edits, caption)
            except AlbumGone:
                photos = rep.photos(t, ("cbc", "da") if has_prices else ("cbc",))
                st["ids"] = tg.send_album(photos, caption)
                st["names"] = [n for n, _ in photos]
            note = rep.note(t, old_cbc, old_da)
            if corrected:
                fix = ", ".join(k.split(":", 1)[1] + (" price" if k.startswith("da:") else "") for k in corrected)
                note = (note or f"🔄 <b>Updated {t:%H:%M}</b> · delivery {rep.d:%a %d.%m}") + f"\n• Corrected: {fix}"
            if note:
                tg.send_text(note, reply_to=st["ids"])
            st.update(cbc=sorted(old_cbc | set(cbc_now)), da=sorted(old_da | set(da_now)), final=final, fp=fp_now)
            log.info("Album updated for %s: %s", rep.d, ", ".join(f"{n} {k}/{m}" for n, k, m in rep.counts()))
            save_state(state, f"{rep.d} updated {t:%H:%M}")
    if st.get("final") and not st.get("excel"):
        if rep.cfg["excel"]:
            tg.send_file(f"N1POWER_{rep.d:%Y-%m}.xlsx", month_workbook(rep.ctx, rep.d, t, rep.da),
                         f"📊 <b>{rep.d:%B %Y}</b> · prices and CBC up to {rep.d:%a %d.%m}",
                         reply_to=st.get("ids"), silent=True)
        st["excel"] = True
        save_state(state, f"{rep.d} month file sent")


def sync_repo() -> None:
    """Start from the latest saved state: a queued run can be checked out before the previous run saved it."""
    if os.getenv("GITHUB_ACTIONS") != "true":
        return
    branch = os.getenv("GITHUB_REF_NAME") or "main"
    try:
        git = lambda *a: subprocess.run(["git", *a], cwd=STATE_FILE.parent, capture_output=True, text=True, timeout=60)
        if git("fetch", "-q", "origin", branch).returncode == 0:
            git("checkout", "-q", "-B", branch, f"origin/{branch}")
    except Exception as e:  # never stop the bot because of this
        log.warning("Could not load the latest state: %s", type(e).__name__)


def watch(ctx, tg, run_minutes: float) -> None:
    cfg = ctx.cfg
    sync_repo()
    stop = now_local() + timedelta(minutes=run_minutes)
    d = now_local().date() + DAY
    state = load_state()
    st = state.setdefault(d.isoformat(), {})
    rep = DayReport(ctx, d)
    log.info("Watching delivery day %s", d)
    while True:
        t = now_local()
        if t.date() + DAY != d:
            break
        if st.get("final") and st.get("excel"):
            log.info("Everything for %s is sent", d)
            break
        if rep.due(t) or rep.final(t):
            step(rep, st, tg, state)
            if st.get("final") and st.get("excel"):
                continue
            t = now_local()
        if rep.due(t):
            wake = t + timedelta(seconds=cfg["poll"])
        else:                                   # nothing due: sleep until the next publication time
            wake = rep.next_check(t) or t + timedelta(seconds=cfg["poll"])
        if wake > stop:
            break
        pause(max((wake - now_local()).total_seconds(), 1))
    save_state(state)


def send_now(ctx, tg, d: date) -> None:
    """One snapshot album right away (manual run); the automatic album is not touched."""
    t = now_local()
    rep = DayReport(ctx, d)
    rep.poll(t, everything=True)
    cbc_now, da_now = rep.known()
    if not cbc_now and not da_now:
        tg.send_text(escape(f"Delivery {d:%a %d.%m.%Y}: nothing published yet."))
        return
    full = bool(da_now)
    ids = tg.send_album(rep.photos(t, ("cbc", "da") if full else ("cbc",)),
                        rep.caption(t, False, snapshot=True, with_prices=full))
    if ctx.cfg["excel"]:
        tg.send_file(f"N1POWER_{d:%Y-%m}.xlsx", month_workbook(ctx, d, t, rep.da),
                     f"📊 <b>{d:%B %Y}</b> · prices and CBC up to {d:%a %d.%m}", reply_to=ids, silent=True)


def find_chat_id(token: str, http) -> int:
    """Sends each chat that pressed Start its own chat id in Telegram (nothing personal goes to the run log)."""
    if not token:
        print("Set the TELEGRAM_BOT_TOKEN secret first.")
        return 1
    bot = Telegram(token, [], http)
    updates = bot.call("getUpdates").get("result")
    chats = set()
    for u in updates if isinstance(updates, list) else []:
        msg = u.get("message") or u.get("channel_post") or u.get("my_chat_member") or {}
        if (msg.get("chat") or {}).get("id") is not None:
            chats.add(msg["chat"]["id"])
    if not chats:
        print("No chats found: open the bot in Telegram, press Start, then run this again.")
        return 1
    for cid in chats:
        bot.call("sendMessage", json={"chat_id": cid, "parse_mode": "HTML",
                                      "text": f"Your TELEGRAM_CHAT_ID is <code>{cid}</code>\n"
                                              "Save it as a GitHub secret with that name."})
    print(f"Sent the chat id to {len(chats)} chat(s) in Telegram - check the bot.")
    return 0


# --------------------------------------------------------------------------- settings and main
def _hhmm(s: str) -> dtime:
    h, m = str(s).split(":")
    return dtime(int(h), int(m))


def load_config(path: Path) -> dict:
    with open(path, "rb") as f:
        c = tomllib.load(f)
    sc, pr, ca, cb, tg = (c.get(k, {}) for k in ("schedule", "prices", "calendar", "cbc", "telegram"))
    borders = lambda lst: [tuple(norm(x) for x in b.split(">", 1)) for b in lst]
    peak = str(pr.get("peak", "08:00-20:00")).split("-")
    cfg = {
"cbc_start": _hhmm(sc.get("cbc_start", "09:30")), "cbc_final": _hhmm(sc.get("cbc_final", "17:00")),
        "prices_start": _hhmm(sc.get("prices_start", "12:55")),
        "prices_final": _hhmm(sc.get("prices_final", "17:00")),
        "prices_first": norm(sc.get("prices_first_post_when", "HU")),
        "poll": float(sc.get("check_every_seconds", 30)),
        "early": int(sc.get("check_early_minutes", 5)),
        "run_minutes": float(sc.get("run_minutes", 0)) or None,
        "markets": [norm(m) for m in pr.get("markets", [])],
        "peak": (_hhmm(peak[0]).hour, _hhmm(peak[1]).hour),
        "sources": [s.lower() for s in pr.get("sources", ["volue", "entsoe"])],
        "volue_curves": {norm(k): v for k, v in pr.get("volue_curves", {}).items()},
        "calendar_countries": [norm(m) for m in ca.get("countries", ca.get("price_calendars", ["HU", "FR", "ES"]))],
        "calendar_split_after": norm(ca.get("split_after", "MK")),
        "spread_base": norm(ca.get("spread_base", "HU")),
        "spreads": [norm(m) for m in ca.get("spreads", [])],
        "country_first": [norm(m) for m in ca.get("country_first", [])],
        "cbc_see": borders(cb.get("see", [])), "cbc_ch": borders(cb.get("ch", [])),
        "excel": bool(tg.get("excel_when_final", True)),
    }
    return _price_times(cfg, pr)


def _price_times(cfg: dict, pr: dict) -> dict:
    """Per-market publication times -> when each market is first checked."""
    pub = {str(k).upper(): v for k, v in pr.get("publish", {}).items()}
    default = _hhmm(pub.pop("DEFAULT", f"{cfg['prices_start']:%H:%M}"))
    early = timedelta(minutes=cfg["early"])
    shift = lambda tt: (datetime.combine(date(2000, 1, 1), tt) - early).time()
    cfg["price_check_from"] = {norm(k): shift(_hhmm(v)) for k, v in pub.items()}
    cfg["price_check_from_default"] = shift(default)
    cfg["first_publish"] = min([_hhmm(v) for v in pub.values()] + [default])
    parts = [f"{label(norm(k))} {v}" for k, v in sorted(pub.items(), key=lambda kv: kv[1])]
    cfg["publish_text"] = ", ".join(parts + [f"others {default:%H:%M}"])
    return cfg


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="N1 POWER bot")
    ap.add_argument("--watch", action="store_true", help="scheduled mode: post and update during the day")
    ap.add_argument("--run-minutes", type=float, default=39)
    ap.add_argument("--now", action="store_true", help="send the current state right away")
    ap.add_argument("--date", help="delivery day YYYY-MM-DD for --now (default tomorrow)")
    ap.add_argument("--dry-run", action="store_true", help="write pictures and Excel to ./preview instead of sending")
    ap.add_argument("--find-chat-id", action="store_true")
    ap.add_argument("--test-jao", action="store_true", help="check the JAO token and show how JAO's answer looks")
    ap.add_argument("--config", default=str(HERE / "config.toml"))
    a = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", stream=sys.stdout)

    http = requests.Session()
    http.headers["User-Agent"] = "n1power-bot/2.0"
    token = env("TELEGRAM_BOT_TOKEN", secret=True)
    chats = [c.strip() for c in env("TELEGRAM_CHAT_ID").split(",") if c.strip()]
    if a.find_chat_id:
        return find_chat_id(token, http)
    entsoe_key = env("ENTSOE_API_KEY", secret=True)
    vid, vsecret = env("VOLUE_CLIENT_ID", secret=True), env("VOLUE_CLIENT_SECRET", secret=True)
    dry = a.dry_run or env("DRY_RUN").lower() in ("1", "true", "yes")
    tg = DryRun(HERE / "preview") if dry else Telegram(token, chats, http)
    try:
        cfg = load_config(Path(a.config))
        if not entsoe_key:
            raise ReportError("ENTSOE_API_KEY is not set")
        if not dry and (not token or not chats):
            raise ReportError("TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID must be set")
        volue = Volue(vid, vsecret, http, cfg["volue_curves"]) if vid and vsecret else None
        jao_key = env("JAO_API_KEY", secret=True)
        ctx = Ctx(cfg, Entsoe(entsoe_key, http), volue, tg, http, Jao(jao_key, http) if jao_key else None)
        if a.test_jao:
            return test_jao(ctx, date.fromisoformat(a.date or env("DELIVERY_DATE") or now_local().date().isoformat()))
        if a.watch:
            watch(ctx, tg, cfg["run_minutes"] or a.run_minutes)
        else:
            d = date.fromisoformat(a.date or env("DELIVERY_DATE") or (now_local().date() + DAY).isoformat())
            send_now(ctx, tg, d)
        if ctx.errors:
            log.info("Source notes: %s", "; ".join(f"{k}: {v}" for k, v in list(ctx.errors.items())[:8]))
        return 0
    except Exception as e:
        msg = redact(str(e) if isinstance(e, ReportError) else f"{type(e).__name__}: {e}")
        log.error("Bot failed: %s", msg)
        if not dry and token and chats:
            state = load_state()
            today = now_local().date().isoformat()
            seen = {k: v for k, v in state.get("_problems", {}).items() if k >= today}
            if msg not in seen.get(today, []):
                try:
                    tg.send_text(f"⚠️ <b>N1 POWER bot problem</b>\n{escape(msg[:600])}")
                    seen[today] = seen.get(today, [])[-4:] + [msg]
                    state["_problems"] = seen
                    save_state(state, "problem noted")
                except Exception:
                    pass
        return 1


if __name__ == "__main__":
    sys.exit(main())
