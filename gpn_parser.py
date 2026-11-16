"""
gpn_parser.py
Reads the weekly Great Park Neighborhoods (GPN) "General Distribution Report" PDF.

One record per collection: neighborhood, builder, collection, product type,
opening / sold-out dates, total units, sold, weekly gross sales, cans, net
sales, traffic, models-open flag. Each neighborhood block is checked against
its own total row (weekly gross / cans / net).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field, asdict
from datetime import datetime, date

import pdfplumber

TITLE = "Great Park Neighborhoods General Distribution Report"
TYPES = {"SFDC", "SFD", "SFA", "SFDA", "MF", "CONDO", "TH"}
DATE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
SFR = re.compile(r"^\d{3,4}-\d{3,4}$")
# columns after "SF Range", in order
NUM_COLS = ["units", "released", "sold", "reservations", "avail", "remaining", "ros", "closed",
            "gross_sales", "cancellations", "net_sales", "net_res", "wk_closed", "traffic",
            "mo_net", "mo_closed", "ytd_net", "ytd_closed"]

BUILDER_NAMES = {"tay mo": "Taylor Morrison", "toll bros.": "Toll Brothers", "toll bros": "Toll Brothers",
                 "kb": "KB Home", "tripointe": "TRI Pointe", "tri pointe": "TRI Pointe",
                 "lennar": "Lennar", "pulte": "Pulte", "trumark": "Trumark", "shea": "Shea",
                 "brookfield": "Brookfield", "melia": "Melia", "risewell": "Risewell"}


@dataclass
class GPNRecord:
    week_ending: date
    neighborhood: str
    gpn_name: str            # exactly as printed, e.g. "Toll Bros., Elm"
    builder: str             # expanded, e.g. "Toll Brothers"
    collection: str          # e.g. "Elm"
    product: str
    product_type: str
    opening: date | None
    sold_out: date | None
    units: int | None
    sold: int | None
    gross_sales: int
    cancellations: int
    net_sales: int
    traffic: int | None
    models_open: str


@dataclass
class GPNResult:
    week_ending: date | None
    records: list[GPNRecord] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    section_checks: list[dict] = field(default_factory=list)

    def to_frame(self):
        import pandas as pd
        return pd.DataFrame([asdict(r) for r in self.records])


def is_gpn_pdf(path) -> bool:
    with pdfplumber.open(path) as pdf:
        return TITLE in (pdf.pages[0].extract_text() or "")


def _num(tok):
    tok = tok.strip()
    if tok in ("-", "", "N/A"):
        return 0
    m = re.fullmatch(r"\((\d+)\)", tok)          # (1) = -1
    if m:
        return -int(m.group(1))
    try:
        return float(tok) if "." in tok else int(tok.replace(",", ""))
    except ValueError:
        return None


def _date(tok):
    return datetime.strptime(tok, "%m/%d/%Y").date() if DATE.match(tok) else None


def _lines(words, tol=2.0):
    words = sorted(words, key=lambda w: (round(w["top"]), w["x0"]))
    out = []
    for w in words:
        if out and abs(out[-1][0]["top"] - w["top"]) <= tol:
            out[-1].append(w)
        else:
            out.append([w])
    return [sorted(l, key=lambda w: w["x0"]) for l in out]


def parse_gpn(path) -> GPNResult:
    res = GPNResult(week_ending=None)
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            m = re.search(r"Week Ending\s+(\d{2}/\d{2}/\d{4})", text)
            if m and res.week_ending is None:
                res.week_ending = _date(m.group(1))
            words = page.extract_words()
            # Collection column: left of the "Product" header (x of first "Product" header word)
            # Boundary between the (centred) Collection and Product columns: midway between headers
            coll_h = [w for w in words if w["text"] == "Collection"]
            prod_h = [w for w in words if w["text"] == "Product"]
            if coll_h and prod_h:
                cc = (coll_h[0]["x0"] + coll_h[0]["x1"]) / 2
                pc = min(((w["x0"] + w["x1"]) / 2 for w in prod_h), key=lambda x: abs(x - cc))
                prod_x = (cc + pc) / 2 + 4
            else:
                prod_x = 184
            neighborhood, sec = None, None
            for line in _lines(words):
                toks = [w["text"] for w in line]
                joined = " ".join(toks)
                # Neighborhood banner (orange bar) - a short line of title text only
                if len(line) <= 5 and not any(_num(t) is not None and t not in ("-",) for t in toks) \
                        and joined not in ("Project Overview Weekly Monthly YTD",) and "Report" not in joined \
                        and not joined.startswith("Week Ending") and joined[:1].isupper() \
                        and "Collection" not in joined and "Units" not in joined:
                    if joined in ("Great Park Neighborhoods",):
                        neighborhood, sec = None, None       # summary table at the end - skip
                    else:
                        neighborhood = joined
                        sec = {"neighborhood": joined, "gross": 0, "cans": 0, "net": 0, "rows": 0}
                    continue
                if neighborhood is None:
                    continue
                # Section total row: "GREAT PARK NEIGHBORHOOD 559 326 ..."
                if joined.upper().startswith(neighborhood.upper()) and joined.split()[len(neighborhood.split())][:1].isdigit():
                    vals = toks[len(neighborhood.split()):]
                    if len(vals) >= 11:
                        tot = dict(zip(NUM_COLS, [_num(v) for v in vals]))
                        chk = {"neighborhood": neighborhood, "rows": sec["rows"],
                               "gross_pdf": tot["gross_sales"], "gross_parsed": sec["gross"],
                               "cans_pdf": tot["cancellations"], "cans_parsed": sec["cans"],
                               "net_pdf": tot["net_sales"], "net_parsed": sec["net"]}
                        chk["ok"] = (chk["gross_pdf"] == chk["gross_parsed"] and chk["cans_pdf"] == chk["cans_parsed"]
                                     and chk["net_pdf"] == chk["net_parsed"])
                        res.section_checks.append(chk)
                        if not chk["ok"]:
                            res.warnings.append(
                                f"GPN totals mismatch: {neighborhood} gross {sec['gross']}/{tot['gross_sales']}, "
                                f"cans {sec['cans']}/{tot['cancellations']}, net {sec['net']}/{tot['net_sales']}")
                    continue
                # Collection row: has an SF range token followed by 18 numeric tokens
                sf_idx = next((i for i, t in enumerate(toks) if SFR.match(t)), None)
                if sf_idx is None or "," not in joined:
                    continue
                nums = toks[sf_idx + 1: sf_idx + 1 + len(NUM_COLS)]
                if len(nums) < len(NUM_COLS):
                    res.warnings.append(f"GPN: could not read row '{joined[:60]}'")
                    continue
                vals = dict(zip(NUM_COLS, [_num(t) for t in nums]))
                models = toks[sf_idx + 1 + len(NUM_COLS)] if len(toks) > sf_idx + 1 + len(NUM_COLS) else ""
                coll_words = [w["text"] for w in line if w["x1"] < prod_x - 4]
                name = " ".join(coll_words)
                if "," not in name:
                    continue
                b_raw, coll = [x.strip() for x in name.split(",", 1)]
                ttype = next((t for t in toks[len(coll_words):sf_idx] if t in TYPES), "")
                dates = [t for t in toks[len(coll_words):sf_idx] if DATE.match(t) or t == "-"]
                rest = toks[len(coll_words):sf_idx]
                prod = " ".join(t for t in rest[: rest.index(ttype)] if True) if ttype in rest else ""
                rec = GPNRecord(
                    week_ending=res.week_ending, neighborhood=neighborhood, gpn_name=name,
                    builder=BUILDER_NAMES.get(b_raw.lower(), b_raw), collection=coll, product=prod,
                    product_type=ttype, opening=_date(dates[0]) if dates else None,
                    sold_out=_date(dates[1]) if len(dates) > 1 else None,
                    units=vals["units"], sold=vals["sold"], gross_sales=vals["gross_sales"] or 0,
                    cancellations=vals["cancellations"] or 0, net_sales=vals["net_sales"] or 0,
                    traffic=vals["traffic"], models_open=models)
                # The report shows cans as a positive count; net = gross - cans
                if rec.net_sales != rec.gross_sales - rec.cancellations:
                    res.warnings.append(f"GPN: {name} net {rec.net_sales} != gross {rec.gross_sales} - cans {rec.cancellations}")
                res.records.append(rec)
                sec["rows"] += 1
                sec["gross"] += rec.gross_sales
                sec["cans"] += rec.cancellations
                sec["net"] += rec.net_sales
    if res.week_ending is None:
        res.warnings.append("GPN: week-ending date not found")
    for r in res.records:
        r.week_ending = res.week_ending
    return res


if __name__ == "__main__":
    import sys
    r = parse_gpn(sys.argv[1])
    df = r.to_frame()
    print("Week ending", r.week_ending, "| collections", len(df))
    print(df[["neighborhood", "gpn_name", "builder", "collection", "product_type", "units", "sold",
              "gross_sales", "cancellations", "net_sales", "traffic", "sold_out", "models_open"]].to_string())
    for c in r.section_checks:
        print(c)
    for w in r.warnings:
        print("WARN", w)
