"""
geocode.py - latitude / longitude for each community.

The weekly report has no street addresses, only community name, builder and city.
So locations come in three levels, recorded in the Communities tab's "Geo Source":

  1. "City center (approximate)"  - from the bundled ca_places.csv (GeoNames), filled
                                    automatically for every community, offline.
  2. "OpenStreetMap: ..."          - found by name with the app's "Find project locations"
                                    button (needs internet; runs on your computer).
  3. "Manual"                      - coordinates you typed in (Communities tab or a weekly tab).
                                    Never overwritten.
"""
from __future__ import annotations

import csv
import json
import math
import re
import time
import urllib.parse
import urllib.request
from functools import lru_cache
from pathlib import Path

CITY_SOURCE = "City center (approximate)"
MANUAL_SOURCE = "Manual"
OSM_PREFIX = "OpenStreetMap: "
NOMINATIM = "https://nominatim.openstreetmap.org/search"
USER_AGENT = "NewHomeSalesWeeklyImport/1.0 (new-home sales tracker)"

# Nominatim results we accept as a project / neighborhood location (not a whole city)
_OK_TYPES = {"neighbourhood", "suburb", "quarter", "residential", "hamlet", "locality",
             "isolated_dwelling", "village", "construction", "allotments", "park", "yes"}
_CITY_TYPES = {"city", "town", "administrative", "county", "state", "postcode"}


@lru_cache(maxsize=1)
def _places():
    p = Path(__file__).with_name("ca_places.csv")
    out = {}
    with open(p, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            out.setdefault(r["name"].strip().lower(), (float(r["lat"]), float(r["lon"])))
    return out


def city_center(city):
    if not isinstance(city, str) or not city.strip():
        return None
    c = city.strip().lower()
    places = _places()
    if c in places:
        return places[c]
    c2 = re.sub(r"\bmt\b", "mountain", c)
    c2 = re.sub(r"\bst\b", "saint", c2)
    return places.get(c2)


def km(a, b):
    (la1, lo1), (la2, lo2) = a, b
    p = math.pi / 180
    h = (math.sin((la2 - la1) * p / 2) ** 2 +
         math.cos(la1 * p) * math.cos(la2 * p) * math.sin((lo2 - lo1) * p / 2) ** 2)
    return 12742 * math.asin(math.sqrt(h))


def _display(name: str) -> str:
    m = re.match(r"^(.*),\s*The$", name.strip(), re.I)        # "Gables, The" -> "The Gables"
    return f"The {m.group(1)}" if m else name.strip()


def queries(community: str, city: str):
    """Most specific first: the community itself, then its master plan ('... at Luna Park')."""
    name = _display(community)
    out = [(f"{name}, {city}, California", "community")]
    if " at " in name:
        plan = name.split(" at ", 1)[1]
        out.append((f"{plan}, {city}, California", "master plan / neighborhood"))
    return out


def _fetch(q, timeout=20):
    url = NOMINATIM + "?" + urllib.parse.urlencode(
        {"q": q, "format": "jsonv2", "limit": 5, "countrycodes": "us", "addressdetails": 0})
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def pick(results, center, max_km=25):
    """Best result that is a place smaller than a city and near the city center."""
    best = None
    for r in results or []:
        typ = str(r.get("type", "")).lower()
        cat = str(r.get("category", r.get("class", ""))).lower()
        if typ in _CITY_TYPES or (cat == "boundary" and typ == "administrative"):
            continue
        if typ not in _OK_TYPES and cat not in ("place", "landuse", "highway", "leisure"):
            continue
        try:
            pt = (float(r["lat"]), float(r["lon"]))
        except (KeyError, ValueError):
            continue
        if center is not None and km(pt, center) > max_km:
            continue
        score = float(r.get("importance", 0) or 0)
        if best is None or score > best[0]:
            best = (score, pt, r.get("display_name", ""))
    return best


def find_location(community, city, fetch=_fetch, pause=1.1):
    """Return (lat, lon, source) from OpenStreetMap, or None. Respects Nominatim's 1 request/second."""
    center = city_center(city)
    errors, tries = 0, 0
    for q, level in queries(community, city):
        tries += 1
        try:
            res = fetch(q)
        except Exception as e:       # no internet / blocked / rate-limited
            res, errors, last = None, errors + 1, e
        if pause:
            time.sleep(pause)
        hit = pick(res, center)
        if hit:
            return hit[1][0], hit[1][1], f"{OSM_PREFIX}{level}"
    if errors == tries:
        raise ConnectionError(f"Could not reach OpenStreetMap: {last}")
    return None
