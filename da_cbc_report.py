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
    return f"{label(b[0])}>{label(b[1])}"


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
            if body[:2] == b"PK":
                with zipfile.ZipFile(io.BytesIO(body)) as z:
                    body = z.read(z.namelist()[0])
            try:
                root = ET.fromstring(body)
            except ET.ParseError:
                problem = f"unreadable answer (HTTP {r.status_code})"
                pause(4)
                continue
            if _tag(root) == "Acknowledgement_MarketDocument":
                reason = " ".join(e.text.strip() for e in root.iter() if _tag(e) == "text" and e.text)
                if "no matching data" in reason.lower():
                    return None
                raise SourceError(f"ENTSO-E: {reason[:160] or 'request not accepted'}")
            if r.status_code != 200:
                raise SourceError(f"ENTSO-E HTTP {r.status_code}")
            return parse_timeseries(root)
        raise SourceError(f"ENTSO-E not reachable: {problem}")

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
        pieces = [s for s in pieces or [] if s["meta"].get("contract_MarketAgreement.type", "A01") in ("A01", "")]
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

    def capacity_range(self, border, d_from: date, d_to: date) -> dict:
        """Daily explicit auctions for one direction: {day: (hourly EUR prices, hourly MW, currency)}."""
        params = {"documentType": "A25", "businessType": "B05", "out_Domain": AREAS[border[0]][0],
                  "in_Domain": AREAS[border[1]][0], **self._window(d_from, d_to)}
        try:
            pieces = self.query({**params, "contract_MarketAgreement.Type": "A01"})
        except SourceError:
            pieces = self.query({**params, "contract_MarketAgreement.type": "A01"})
        priced = [s for s in pieces or []
                  if s["meta"].get("contract_MarketAgreement.type", "A01") in ("A01", "")
                  and any(b[2] is not None for b in s["points"])]
        use = [s for s in priced if s["meta"].get("businessType") == "B05"] or priced
        if not use:
            return {}
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
        return out

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

    def day_ahead(self, area: str, d_from: date, d_to: date) -> dict:
        names = [self.curves[area]] if area in self.curves else [f"pri {v} spot €/mwh cet h a" for v in AREAS[area][2]]
        for name in names:
            cid = self._curve_id(name)
            if cid is None:
                continue
            r = self.http.get(f"{VOLUE_API_URL}/api/series/{cid}",
                              params={"from": d_from.isoformat(), "to": d_to.isoformat()},
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
                    out[d] = (hourly, "EUR")
                d += DAY
            return out
        return {}


# --------------------------------------------------------------------------- price store
class Prices:
    """Hourly day-ahead prices per market and day, from the sources in config order."""

    def __init__(self, ctx):
        self.ctx = ctx
        self.cache: dict = {}          # (market, day) -> (hourly, source)
        self._lock = threading.Lock()
        self._fx: dict = {}

    def _uah(self, d: date):
        if d not in self._fx:
            try:
                r = self.ctx.http.get(NBU_URL, params={"valcode": "EUR", "date": d.strftime("%Y%m%d"), "json": ""},
                                      timeout=20)
                self._fx[d] = float(r.json()[0]["rate"])
            except Exception:
                self._fx[d] = None
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

    def send_photos(self, photos) -> list[list[int]]:
        ids = []
        for chat in self.chats:
            row = []
            for name, png, caption in photos:
                j = self.call("sendPhoto", data={"chat_id": chat, "caption": caption, "parse_mode": "HTML"},
                              files={"photo": (f"{name}.png", png, "image/png")})
                row.append(j["result"]["message_id"])
            ids.append(row)
        return ids

    def edit_photos(self, ids: list[list[int]], photos) -> list[list[int]]:
        out = []
        for c, chat in enumerate(self.chats):
            row = list(ids[c]) if c < len(ids) else []
            for k, (name, png, caption) in enumerate(photos):
                media = {"type": "photo", "media": "attach://pic", "caption": caption, "parse_mode": "HTML"}
                try:
                    if k >= len(row):
                        raise TelegramRefused("message to edit not found", "editMessageMedia")
                    self.call("editMessageMedia", data={"chat_id": chat, "message_id": row[k], "media": json.dumps(media)},
                              files={"pic": (f"{name}.png", png, "image/png")})
                except TelegramRefused as e:
                    if "not modified" in str(e):
                        continue
                    if "not found" not in str(e) and "can't be edited" not in str(e):
                        raise
                    j = self.call("sendPhoto", data={"chat_id": chat, "caption": caption, "parse_mode": "HTML"},
                                  files={"photo": (f"{name}.png", png, "image/png")})
                    if k < len(row):
                        row[k] = j["result"]["message_id"]
                    else:
                        row.append(j["result"]["message_id"])
            out.append(row)
        return out

    def send_text(self, html: str) -> None:
        for chat in self.chats:
            self.call("sendMessage", json={"chat_id": chat, "text": html, "parse_mode": "HTML",
                                           "disable_web_page_preview": True})

    def send_file(self, name: str, data: bytes, caption: str) -> None:
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        for chat in self.chats:
            self.call("sendDocument", data={"chat_id": chat, "caption": caption, "parse_mode": "HTML"},
                      files={"document": (name, data, mime)})


class TelegramRefused(ReportError):
    def __init__(self, description: str, method: str):
        super().__init__(f"Telegram refused {method}: {description}")


class DryRun:
    """Stands in for Telegram: writes the pictures and files to a folder."""

    def __init__(self, folder: Path):
        self.folder, self.n = folder, 0
        folder.mkdir(parents=True, exist_ok=True)

    def send_photos(self, photos):
        for name, png, caption in photos:
            (self.folder / f"{name}.png").write_bytes(png)
            print(f"[picture {name}.png] {re.sub('<[^>]+>', '', caption)}")
        self.n += 1
        return [[self.n * 10 + k for k in range(len(photos))]]

    def edit_photos(self, ids, photos):
        self.send_photos(photos)
        return ids

    def send_text(self, html):
        print("[message] " + re.sub("<[^>]+>", "", html))

    def send_file(self, name, data, caption):
        (self.folder / name).write_bytes(data)
        print(f"[file {name}] {re.sub('<[^>]+>', '', caption)}")


# --------------------------------------------------------------------------- report pieces
class Ctx:
    def __init__(self, cfg, entsoe, volue, out, http):
        self.cfg, self.entsoe, self.volue, self.out, self.http = cfg, entsoe, volue, out, http
        self.errors: dict = {}
        self.prices = Prices(self)

    def capacity(self, border, d):
        try:
            return self.entsoe.capacity(border, d)
        except SourceError as e:
            self.errors[blabel(border)] = redact(e)
            return None

    def capacity_range(self, border, d_from, d_to) -> dict:
        try:
            return self.entsoe.capacity_range(border, d_from, d_to)
        except SourceError as e:
            self.errors[blabel(border)] = redact(e)
            return {}


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
        self.inactive: set = set()
        self.checked = False

    def poll(self):
        todo = [b for b in self.borders if b not in self.data]
        for b, r in zip(todo, parallel(lambda b: self.ctx.capacity(b, self.d), todo)):
            if r:
                self.data[b] = r
        if not self.checked:                    # borders without an auction yesterday count as n/a
            self.checked = True
            todo = [b for b in self.borders if b not in self.data]
            prev = parallel(lambda b: self.ctx.capacity(b, self.d - DAY), todo)
            self.inactive = {b for b, r in zip(todo, prev) if not r}
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
        banner = f"CBC {group}: {known}/{len(self.active(borders))} borders known{at}" + (" · final" if final else "")
        footer = "Daily explicit capacity auctions, EUR/MWh · ENTSO-E · Peak = hours 9-20"
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

    def poll(self):
        todo = [m for m in self.markets if m not in self.data]
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
        banner = f"{len(self.data)}/{len(self.active())} markets ACTUAL{at}" + (" · final" if final else "")
        via_entsoe = [label(m) for m in self.data if self.ctx.prices.source(m, self.d) == "ENTSO-E"]
        footer = "EUR/MWh · Peak = hours 9-20 · IT = PUN · source Volue" + \
                 (f" (ENTSO-E: {', '.join(via_entsoe)})" if via_entsoe else "")
        if self.ctx.volue is None or self.ctx.volue.disabled:
            footer += " · Volue not available - ENTSO-E used, no PUN"
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
            footer += f"   DoD {R.signed(mtd - prev)}"
        block = {"title": title, "kind": "price", "weeks": cells, "mtd": mtd,
                 "week_base": [avg(base.get(x) for x in wk if x <= d) for wk in weeks],
                 "week_peak": [avg(peak.get(x) for x in wk if x <= d) for wk in weeks],
                 "week_partial": [any(x > d or x not in base for x in wk) for wk in weeks], "footer": footer}
        if spread:
            name, ser = spread
            srun = self._running(ser, first, min(d, last))
            block.update(spread_name=name, week_spread=[avg(ser.get(x) for x in wk if x <= d) for wk in weeks])
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
                rows.append((name, [avg(ser.get(x) for x in wk if x <= self.d) for wk in weeks]
                             + [avg(v for x, v in ser.items() if first <= x <= self.d)]))
            blocks.append({"title": f"Week recap · spreads to {label(c['spread_base'])} (base)", "kind": "recap",
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
            title = f"Day-ahead calendar {k}/{len(pages)} · {self.d:%B %Y} · {span}{recap} · to {self.d:%a %d.%m}   ({legend})"
            cap = (f"<b>Calendar {k}/{len(pages)} · {self.d:%B %Y} · {span}{recap}</b>\n{legend}\n"
                   "Week = Mon-Sun, spread to HU in blue (HUxx = HU − xx; DEHU, GRHU, ATHU = country − HU)")
            out.append((f"prices_calendar_{k}", R.calendars_png(title, page, 3), cap))
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
        ws.cell(2, 1, "Calendars: B = base, P = peak 08-20, MTD = month average of base to that day, Week = Mon-Sun. "
                      "Below: hourly prices, newest day first. EUR/MWh.")
        row = 4 + R.xl_calendars(ws, 4, blocks, 3) + 1
        for tb in tables:
            row += R.xl_hourly(ws, row, 1, tb) + 1
        for col in range(1, 30):
            ws.column_dimensions[get_column_letter(col)].width = 9.5
        ws2 = wb.create_sheet(f"CBC {d:%b %Y}")
        ws2.cell(1, 1, f"N1 POWER · daily explicit capacity auctions · {d:%B %Y} · updated {t:%a %d.%m %H:%M}").font = \
            Font(bold=True, size=14)
        ws2.cell(2, 1, "EUR/MWh (BA>RS, RS>BA converted from BAM at 1.95583). Top: daily average per border. "
                       "Below: hourly prices, newest day first (SEE left, CH right).")
        row = 4 + R.xl_matrix(ws2, 4, 1, "Daily average price", "Day", [blabel(b) for b in cbc.borders], daily) + 1
        width_see = len(dict(cbc.groups).get("SEE", [])) + 2
        for pair in cbc_tables:
            used = 0
            for k, tb in enumerate(pair):
                used = max(used, R.xl_hourly(ws2, row, 1 + k * width_see, tb))
            row += used + 1
        for col in range(1, 50):
            ws2.column_dimensions[get_column_letter(col)].width = 8.5
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


# --------------------------------------------------------------------------- watcher
def tick(job, st: dict, tg, ctx, final_time: dtime, state: dict) -> None:
    t = now_local()
    job.poll()
    deadline = t.time() >= final_time
    done = job.complete()
    if st.get("final"):
        return
    if not st.get("posted"):
        if job.ready() or (deadline and job.data):
            st["msgs"] = tg.send_photos(job.render(done or deadline, t))
            st.update(posted=True, sig=job.signature(), first_at=f"{t:%H:%M}")
            log.info("%s: first post (%s)", job.key, job.signature() or "-")
            save_state(state, f"{job.key} {job.d} posted")
        elif deadline:
            tg.send_text(escape(job.nothing_caption()))
            st["final"] = True
            save_state(state, f"{job.key} {job.d} nothing published")
            return
    elif job.signature() != st.get("sig") or ((done or deadline) and not st.get("final")):
        st["msgs"] = tg.edit_photos(st.get("msgs", []), job.render(done or deadline, t))
        st["sig"] = job.signature()
        log.info("%s: updated (%d known)", job.key, len(job.data))
        save_state(state)
    if st.get("posted") and (done or deadline) and not st.get("final"):
        if job.key == "cbc" or not ctx.cfg["excel"]:
            tg.send_text(job.final_caption(t))
        else:
            tg.send_file(f"N1POWER_{job.d:%Y-%m}.xlsx", month_workbook(ctx, job.d, t, job), job.final_caption(t))
        st.update(final=True, final_at=f"{t:%H:%M}")
        log.info("%s: final", job.key)
        save_state(state, f"{job.key} {job.d} final")


def watch(ctx, tg, run_minutes: float) -> None:
    cfg = ctx.cfg
    stop = now_local() + timedelta(minutes=run_minutes)
    d = now_local().date() + DAY
    state = load_state()
    jobs = [(CbcJob(ctx, d), cfg["cbc_start"], cfg["cbc_final"]),
            (PriceJob(ctx, d), cfg["prices_start"], cfg["prices_final"])]
    log.info("Watching delivery day %s", d)
    while True:
        t = now_local()
        if t.date() + DAY != d:
            break
        waiting = []
        for job, start, final in jobs:
            st = state.setdefault(d.isoformat(), {}).setdefault(job.key, {})
            if st.get("final"):
                continue
            if t.time() >= start:
                tick(job, st, tg, ctx, final, state)
            if not st.get("final"):
                waiting.append(start)
        if not waiting:
            log.info("Everything for %s is sent", d)
            break
        t = now_local()
        upcoming = [s for s in waiting if s > t.time()]
        if len(upcoming) == len(waiting):          # nothing running yet: sleep until the next window
            wait = (datetime.combine(t.date(), min(upcoming), TZ) - t).total_seconds()
        else:
            wait = cfg["poll"]
        if t + timedelta(seconds=wait) > stop:
            break
        pause(max(wait, 1))
    save_state(state)


def send_now(ctx, tg, d: date) -> None:
    t = now_local()
    cbc, prices = CbcJob(ctx, d), PriceJob(ctx, d)
    cbc.poll()
    prices.poll()
    if cbc.data:
        tg.send_photos(cbc.render(cbc.complete(), t))
    else:
        tg.send_text(escape(f"CBC {d:%a %d.%m.%Y}: no capacity auction results on ENTSO-E yet."))
    if prices.data:
        tg.send_photos(prices.render(prices.complete(), t))
    else:
        tg.send_text(escape(f"DA prices {d:%a %d.%m.%Y}: not published yet."))
    if ctx.cfg["excel"] and (cbc.data or prices.data):
        tg.send_file(f"N1POWER_{d:%Y-%m}.xlsx", month_workbook(ctx, d, t, prices),
                     f"📊 {d:%B %Y} file · prices and CBC up to {d:%a %d.%m}")


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
    return {
        "cbc_start": _hhmm(sc.get("cbc_start", "09:15")), "cbc_final": _hhmm(sc.get("cbc_final", "12:30")),
        "prices_start": _hhmm(sc.get("prices_start", "12:55")),
        "prices_final": _hhmm(sc.get("prices_final", "16:00")),
        "prices_first": norm(sc.get("prices_first_post_when", "HU")),
        "poll": float(sc.get("check_every_seconds", 60)),
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


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="N1 POWER bot")
    ap.add_argument("--watch", action="store_true", help="scheduled mode: post and update during the day")
    ap.add_argument("--run-minutes", type=float, default=39)
    ap.add_argument("--now", action="store_true", help="send the current state right away")
    ap.add_argument("--date", help="delivery day YYYY-MM-DD for --now (default tomorrow)")
    ap.add_argument("--dry-run", action="store_true", help="write pictures and Excel to ./preview instead of sending")
    ap.add_argument("--find-chat-id", action="store_true")
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
        ctx = Ctx(cfg, Entsoe(entsoe_key, http), volue, tg, http)
        if a.watch:
            watch(ctx, tg, a.run_minutes)
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
