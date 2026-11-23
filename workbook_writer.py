"""
workbook_writer.py - renders the database dict into the Excel workbook.
The three weekly views (Gross Sales, Cancellations, Net Sales) are rebuilt from
Weekly_Data on every save; Net Sales cells are formulas (Gross - Cancellations).
"""
from __future__ import annotations

from datetime import timedelta, date

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.table import Table, TableStyleInfo

FONT = "Arial"
F_BASE = Font(name=FONT, size=9)
F_BOLD = Font(name=FONT, size=9, bold=True)
F_HDR = Font(name=FONT, size=9, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
F_NOTE = Font(name=FONT, size=8, italic=True, color="595959")
F_SO = Font(name=FONT, size=7, color="808080")
F_LEGACY = Font(name=FONT, size=9, italic=True, color="7F7F7F")
F_MANUAL = Font(name=FONT, size=9, color="0000FF")
F_GPN = Font(name=FONT, size=9, color="00806B", bold=True)
F_MKT = Font(name=FONT, size=10, bold=True, color="FFFFFF")

FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_MKT = PatternFill("solid", fgColor="2F5597")
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
FILL_SUB = PatternFill("solid", fgColor="EDEDED")
FILL_SO = PatternFill("solid", fgColor="F2F2F2")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
STATUS_FILL = {
    "NEW": PatternFill("solid", fgColor="C6EFCE"),
    "ADDED": PatternFill("solid", fgColor="DDEBF7"),
    "RETURNED": PatternFill("solid", fgColor="FFEB9C"),
    "DROPPED": PatternFill("solid", fgColor="FFC7CE"),
    "SOLD OUT": PatternFill("solid", fgColor="D9D9D9"),
    "ACTIVE - ": PatternFill("solid", fgColor="FCE4D6"),
    "MANUAL": PatternFill("solid", fgColor="E4DFEC"),
    "GPN": PatternFill("solid", fgColor="D9F2EC"),
}
THIN = Side(style="thin", color="A6A6A6")
B_TOP = Border(top=THIN)
CENTER = Alignment(horizontal="center", vertical="center")

HDR_ROW, DATA0 = 4, 5


def _sum_if_any(colL, rows):
    """Blank when no underlying data that week, so pace averages skip it."""
    if not rows:
        return None
    args = ",".join(f"{colL}{x}" for x in rows)
    return f'=IF(COUNT({args})=0,"",SUM({args}))'


def _status_fill(st):
    st = str(st or "")
    for k, f in STATUS_FILL.items():
        if st.startswith(k):
            return f
    return None


def _week_list(weekly):
    wks = sorted(set(pd.to_datetime(weekly["week_ending"]).dt.date))
    if not wks:
        return []
    out, d = [], wks[0]
    while d <= wks[-1]:
        out.append(d)
        d += timedelta(days=7)
    return out


def _layout(comm):
    """Ordered list of ('market', m) / ('section', (cg, sm)) / ('row', community) entries."""
    c = comm.copy()
    for k in ("m_ord", "s_ord", "r_ord"):
        c[k] = pd.to_numeric(c[k], errors="coerce").fillna(9999)
    c["_active"] = c["status"].astype(str).str.match(r"^(NEW|ADDED|RETURNED|ACTIVE)").map({True: 0, False: 1})
    m_order = c.groupby("market")["m_ord"].min().sort_values().index.tolist()
    out = []
    for m in m_order:
        cm = c[c["market"] == m]
        out.append(("market", m))
        s_order = cm.groupby(["county_group", "submarket"], dropna=False)["s_ord"].min().sort_values()
        for (cg, sm), _ in s_order.items():
            cs = cm[(cm["county_group"].fillna("") == ("" if pd.isna(cg) else cg)) &
                    (cm["submarket"].fillna("") == ("" if pd.isna(sm) else sm))]
            cs = cs.sort_values(["_active", "r_ord", "community"])
            out.append(("section", (cg, sm)))
            for r in cs.itertuples():
                out.append(("row", r))
            out.append(("subtotal", sm))
        out.append(("mtotal", m))
    out.append(("gtotal", "SoCal"))
    return out


# Column layouts (key, header, width). Order matches Daniel's 11/13 workbook.
_SUMMARY = [("t4", "4 Wk\nTotal", 13), ("p4", "4 Wk\nPace", 13), ("t8", "8 Wk\nTotal", 13),
            ("p8", "8 Wk\nPace", 13), ("t13", "13 Wk\nTotal", 13), ("p13", "13 Wk\nPace", 13)]
COLS = {
    "base": [("community", "Community", 32), ("builder", "Builder", 18), ("city", "City", 12.3),
             ("lat", "Latitude", 12.5), ("lon", "Longitude", 13.5), ("type", "Type", 6), ("status", "Status", 13), ("first", "First Wk", 9),
             ("last", "Last Rpt", 13), ("units", "TotalUnits", 6), ("total", "Total Sold", 7)]
            + _SUMMARY,
    "net": [("community", "Community", 32), ("builder", "Builder", 18), ("city", "City", 12.3),
            ("lat", "Latitude", 12.5), ("lon", "Longitude", 13.5), ("type", "Type", 6), ("status", "Status", 13), ("label", "Project Label", 29.4),
            ("pace", "Pace Label", 13), ("first", "First Wk", 9), ("last", "Last Rpt", 13),
            ("units", "TotalUnits", 6), ("total", "Total Sold", 7), ("loj", "LOJ Pace", 6)]
           + _SUMMARY,
}


def _cols(kind):
    cols = COLS["net" if kind == "net" else "base"]
    C = {k: i for i, (k, _, _) in enumerate(cols, start=1)}
    return cols, C, len(cols) + 2          # first week column (after one spacer)


def _write_view(ws, kind, layout, weeks, cell_lookup, latest):
    """kind: 'gross' | 'cans' | 'net'"""
    cols, C, WK0 = _cols(kind)
    if kind == "cans":
        cols = [(k, "Total Cans" if k == "total" else h, w) for k, h, w in cols]
    _, _, GWK0 = _cols("gross")            # Net formulas point at the Gross/Cans week columns
    CL = {k: L(i) for k, i in C.items()}
    titles = {"gross": "Gross Sales (Wk's Sales)", "cans": "Cancellations (Wk's Cans)",
              "net": "Net Sales (Gross Sales - Cancellations)"}
    last_col = WK0 + len(weeks) - 1
    ws["A1"] = f"SoCal New Home {titles[kind]}"
    ws["A1"].font = F_TITLE
    ws["E1"] = f"Latest week ending: {latest:%m/%d/%Y}" if latest else ""
    ws["E1"].font = F_BOLD
    legend = ("SOLD OUT = community no longer on the weekly report (kept for history).  "
              "Bold teal = from the Great Park Neighborhoods (GPN) report, which replaces the weekly report for those projects.  "
              "Latitude/Longitude in grey italic = city center (approximate); black = OpenStreetMap or typed in.  "
              "N/R = listed on the report but didn't report that week (Traffic N/A) - left out of totals and pace.  "
              "Status colors: green NEW, blue ADDED (first week tracked), yellow RETURNED, "
              "red DROPPED this week, grey SOLD OUT, orange S/O or TSO per report, purple MANUAL "
              "(not on the weekly report - blue numbers come from the Manual_Entries tab).")
    if kind == "net":
        legend = ("Net = 'Gross Sales' - 'Cancellations' (live formulas).  Grey italic = legacy "
                  "net-only history from New Home Sales Trends v4.  Project Label is yours to edit "
                  "(here, in the Communities tab, or in the app) and is kept on every import.  " + legend)
    ws["A2"] = legend
    ws["A2"].font = F_NOTE

    for i, (_, h, wdt) in enumerate(cols, start=1):
        c = ws.cell(HDR_ROW, i, h)
        c.font = F_HDR
        c.fill = PatternFill("solid", fgColor="BF8F00") if cols[i - 1][0] == "label" else FILL_HDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[L(i)].width = wdt
    ws.column_dimensions[L(len(cols) + 1)].width = 1.5
    prev_m = None
    for j, wk in enumerate(weeks):
        col = WK0 + j
        c = ws.cell(HDR_ROW, col, wk)
        c.number_format = "m/d/yy"
        c.font, c.fill, c.alignment = F_HDR, FILL_HDR, CENTER
        ws.column_dimensions[L(col)].width = 7.5
        if wk.month != prev_m:
            mc = ws.cell(3, col, wk.strftime("%b %Y"))
            mc.font = F_BOLD
            prev_m = wk.month
    ws.row_dimensions[HDR_ROW].height = 26
    ws.cell(3, C["total"], "Summary (last N weeks)").font = F_NOTE

    nw = len(weeks)
    wc = lambda j: L(WK0 + j)
    rngs = {"all": (0, nw - 1), "4": (max(0, nw - 4), nw - 1), "8": (max(0, nw - 8), nw - 1),
            "13": (max(0, nw - 13), nw - 1)}

    def summary(r, bold=False, community_row=False):
        f = F_BOLD if bold else F_BASE
        spec = [("total", "SUM", "all"), ("t4", "SUM", "4"), ("p4", "AVG", "4"), ("t8", "SUM", "8"),
                ("p8", "AVG", "8"), ("t13", "SUM", "13"), ("p13", "AVG", "13")]
        if community_row and "loj" in C:
            spec.append(("loj", "AVG", "all"))      # life-of-job pace: average of all weekly net sales
        for key, fn, k in spec:
            a, b = rngs[k]
            rg = f"{wc(a)}{r}:{wc(b)}{r}"
            v = f"=SUM({rg})" if fn == "SUM" else f'=IFERROR(AVERAGE({rg}),"")'
            c = ws.cell(r, C[key], v)
            c.font = f
            c.number_format = "0.00" if fn == "AVG" else "0;[Red]-0;0"
            c.alignment = CENTER
        if community_row and "pace" in C:
            c = ws.cell(r, C["pace"],
                        f'=IF({CL["loj"]}{r}="","",_xlfn.CONCAT({CL["total"]}{r}," of ",{CL["units"]}{r}," sold  |  ",'
                        f'TEXT({CL["p4"]}{r},"0.00"),"/wk (4 wk)  |  ",'
                        f'TEXT({CL["p13"]}{r},"0.00"),"/wk (13 wk)  |  ",'
                        f'TEXT({CL["loj"]}{r},"0.00"),"/wk (LOJ)"))')
            c.font = F_BASE

    r = DATA0
    sec_rows, sub_rows, mkt_rows = [], [], []
    for typ, obj in layout:
        if typ == "market":
            ws.cell(r, 1, obj)
            for col in range(1, last_col + 1):
                ws.cell(r, col).fill = FILL_MKT
                ws.cell(r, col).font = F_MKT
            sub_rows = []
            r += 1
        elif typ == "section":
            cg, sm = obj
            label = f"{cg} | {sm}" if pd.notna(cg) and cg and cg != sm else str(sm)
            ws.cell(r, 1, label)
            for col in range(1, last_col + 1):
                ws.cell(r, col).fill = FILL_SEC
                ws.cell(r, col).font = F_BOLD
            sec_rows = []
            r += 1
        elif typ == "row":
            cr = obj
            vals = {"community": cr.community, "builder": cr.builder, "city": cr.city,
                    "type": cr.product_type, "status": cr.status,
                    "label": getattr(cr, "project_label", None),
                    "first": cr.first_week, "last": cr.last_reported,
                    "units": cr.units, "lat": cr.latitude, "lon": cr.longitude}
            for key, v in vals.items():
                if key not in C:
                    continue
                v = None if (v is None or (not isinstance(v, str) and pd.isna(v))) else v
                if hasattr(v, "item"):
                    v = v.item()
                c = ws.cell(r, C[key], v)
                c.font = F_BASE
                if key in ("first", "last"):
                    c.number_format = "m/d/yy"
                    c.alignment = CENTER
                if key in ("type", "units"):
                    c.alignment = CENTER
                if key in ("lat", "lon") and v is not None:
                    c.value = float(v)
                    c.number_format = "0.00000000"
                    if str(getattr(cr, "geo_source", "")).startswith("City center"):
                        c.font = F_LEGACY          # grey italic = approximate (city center)
            sf = _status_fill(cr.status)
            if sf:
                ws.cell(r, 1).fill = sf
                ws.cell(r, C["status"]).fill = sf
            for j, wk in enumerate(weeks):
                rec = cell_lookup.get((cr.community_id, wk))
                if rec is None:
                    continue
                c = ws.cell(r, WK0 + j)
                src, txt = rec["source"], rec["legacy_text"]
                if src == "Not on report" or (isinstance(txt, str) and txt.upper() == "SOLD OUT"):
                    c.value, c.font, c.fill = "SOLD OUT", F_SO, FILL_SO
                elif src == "Weekly report" and txt == "N/R":
                    c.value, c.font = "N/R", F_SO          # listed, didn't report: excluded from pace
                elif src in ("Weekly report", "Manual entry", "GPN report"):
                    if kind == "gross":
                        c.value = int(rec["gross_sales"])
                    elif kind == "cans":
                        c.value = int(rec["cancellations"])
                    else:
                        g = L(GWK0 + j)
                        c.value = f"='Gross Sales'!{g}{r}-'Cancellations'!{g}{r}"
                    c.font = F_MANUAL if src == "Manual entry" else (F_GPN if src == "GPN report" else F_BASE)
                    c.number_format = "0;[Red]-0;0"
                else:   # legacy
                    if kind != "net":
                        continue
                    if rec["net_sales"] is not None and pd.notna(rec["net_sales"]):
                        c.value = int(rec["net_sales"])
                        c.number_format = "0;[Red]-0;0"
                        c.font = F_LEGACY
                    elif isinstance(txt, str):
                        c.value, c.font = txt, F_SO
                c.alignment = CENTER
            summary(r, community_row=True)
            sec_rows.append(r)
            r += 1
        elif typ == "subtotal":
            ws.cell(r, 1, f"{obj} Total")
            if sec_rows:
                a, b = sec_rows[0], sec_rows[-1]
                u = CL["units"]
                ws.cell(r, C["units"], f"=SUM({u}{a}:{u}{b})")
                for j in range(nw):
                    rg = f"{wc(j)}{a}:{wc(j)}{b}"
                    ws.cell(r, WK0 + j, f'=IF(COUNT({rg})=0,"",SUM({rg}))').number_format = "0;[Red]-0;0"
            summary(r, bold=True)
            for col in range(1, last_col + 1):
                c = ws.cell(r, col)
                c.fill, c.border = FILL_SUB, B_TOP
                c.font = F_BOLD
                if col >= WK0 or col == C["units"]:
                    c.alignment = CENTER
            sub_rows.append(r)
            r += 2
        elif typ in ("mtotal", "gtotal"):
            rows = sub_rows if typ == "mtotal" else mkt_rows
            ws.cell(r, 1, f"{obj} Total" if typ == "mtotal" else "SoCal Total (all markets)")
            if rows:
                ws.cell(r, C["units"], "=" + "+".join(f"{CL['units']}{x}" for x in rows))
            for j in range(nw):
                ws.cell(r, WK0 + j, _sum_if_any(wc(j), rows)).number_format = "0;[Red]-0;0"
            summary(r, bold=True)
            for col in range(1, last_col + 1):
                c = ws.cell(r, col)
                if typ == "mtotal":
                    c.fill, c.font = FILL_SEC, F_BOLD
                else:
                    c.fill, c.font = FILL_HDR, F_HDR
                if col >= WK0 or col == C["units"]:
                    c.alignment = CENTER
            if typ == "mtotal":
                mkt_rows.append(r)
            r += 2
    # Freeze headers + Community/Builder/City (as in the 11/13 workbook)
    ws.freeze_panes = ws.cell(DATA0, C["lat"])
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.outlinePr.summaryBelow = False


def _plain_table(ws, df, start_row, header_map=None, name=None, widths=None, date_cols=(),
                 input_cols=()):
    cols = list(df.columns)
    for i, col in enumerate(cols, start=1):
        h = header_map.get(col, col) if header_map else col
        c = ws.cell(start_row, i, h)
        c.font, c.fill = F_HDR, (FILL_HDR if col not in input_cols else PatternFill("solid", fgColor="BF8F00"))
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[L(i)].width = (widths or {}).get(col, 14)
    for ri, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        for ci, v in enumerate(row, start=1):
            if v is None or (not isinstance(v, str) and pd.isna(v)):
                v = None
            elif hasattr(v, "item"):
                v = v.item()
            c = ws.cell(ri, ci, v)
            c.font = F_BASE
            if cols[ci - 1] in date_cols:
                c.number_format = "m/d/yyyy"
            if cols[ci - 1] in input_cols:
                c.fill = FILL_INPUT
    ws.freeze_panes = ws.cell(start_row + 1, 1)
    if len(df):
        ref = f"A{start_row}:{L(len(cols))}{start_row + len(df)}"
        if name:
            t = Table(displayName=name, ref=ref)
            t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
            ws.add_table(t)
        else:
            ws.auto_filter.ref = ref


def write_workbook(db, path, header_map):
    comm = db["communities"].copy()
    weekly = db["weekly"].copy()
    weekly["week_ending"] = pd.to_datetime(weekly["week_ending"]).dt.date
    weeks = _week_list(weekly)
    pdf_weeks = weekly.loc[weekly["source"] == "Weekly report", "week_ending"]
    latest = max(pdf_weeks) if len(pdf_weeks) else (weeks[-1] if weeks else None)
    for col in ("first_week", "last_reported"):
        comm[col] = pd.to_datetime(comm[col], errors="coerce").dt.date
    lookup = {(r["community_id"], r["week_ending"]): r for r in weekly.to_dict("records")
              if r["source"] != "Weekly report (superseded by GPN)"}
    layout = _layout(comm)

    wb = Workbook()
    wb.remove(wb.active)
    for name, kind in (("Gross Sales", "gross"), ("Cancellations", "cans"), ("Net Sales", "net")):
        ws = wb.create_sheet(name)
        _write_view(ws, kind, layout, weeks, lookup, latest)

    # Change log
    ws = wb.create_sheet("Change_Log")
    ws["A1"] = "Change Log - new, dropped (likely sold out), returned and flagged communities"
    ws["A1"].font = F_TITLE
    ws["A2"] = ("Rebuilt from Weekly_Data on every import. DROPPED = was on the prior report, "
                "missing this week. ADDED = first week tracked in this database (submarket had no earlier data).")
    ws["A2"].font = F_NOTE
    chg = db["change_log"].copy()
    chg["week_ending"] = pd.to_datetime(chg["week_ending"]).dt.date
    _plain_table(ws, chg.rename(columns={"week_ending": "Week Ending", "market": "Market",
                                         "submarket": "Submarket", "community": "Community",
                                         "builder": "Builder", "event": "Event", "detail": "Detail",
                                         "community_id": "Community ID"}),
                 3, widths={"Community": 32, "Detail": 70, "Submarket": 22, "Community ID": 30},
                 date_cols=("Week Ending",))
    for r in range(4, ws.max_row + 1):
        f = _status_fill(ws.cell(r, 6).value)
        if f:
            ws.cell(r, 6).fill = f

    # Communities
    ws = wb.create_sheet("Communities")
    ws["A1"] = "Communities - master list (never deleted; sold-out projects stay for history)"
    ws["A1"].font = F_TITLE
    ws["A2"] = ("Yellow columns (Project Label, Total Lots, Grand Opening, Tracking, My Notes) are yours to edit and are "
                "preserved on every import. Tracking: Weekly report (default) or Manual (not on the report - "
                "never flagged sold out; enter weeks in Manual_Entries). Other columns refresh from the latest PDF.")
    ws["A2"].font = F_NOTE
    cdf = comm.copy()
    cdf["project_label_saved"] = cdf["project_label"]
    cdf["geo_saved"] = [f"{float(a):.8f},{float(b):.8f}" if pd.notna(a) and pd.notna(b) else None
                        for a, b in zip(cdf["latitude"], cdf["longitude"])]
    for col in ("grand_opening", "last_pdf_week"):
        cdf[col] = pd.to_datetime(cdf[col], errors="coerce").dt.date
    cdf["_lay"] = cdf["community_id"].map(
        {o.community_id: i for i, (t, o) in enumerate(layout) if t == "row"})
    cdf = cdf.sort_values("_lay").drop(columns="_lay")
    _plain_table(ws, cdf, 3, header_map=header_map,
                 widths={"community_id": 30, "community": 32, "builder": 18, "county_group": 22,
                         "submarket": 22, "status": 18, "user_notes": 30, "legacy_name": 26,
                         "project_label": 32},
                 date_cols=("grand_opening", "first_week", "last_reported", "last_pdf_week"),
                 input_cols=("project_label", "latitude", "longitude", "total_lots", "grand_opening",
                             "tracking", "user_notes"))
    for colname in ("latitude", "longitude"):
        idx = list(cdf.columns).index(colname) + 1
        for r in range(4, ws.max_row + 1):
            ws.cell(r, idx).number_format = "0.00000000"
    st_col = list(cdf.columns).index("status") + 1
    for r in range(4, ws.max_row + 1):
        f = _status_fill(ws.cell(r, st_col).value)
        if f:
            ws.cell(r, st_col).fill = f
    for colname in ("m_ord", "s_ord", "r_ord", "project_label_saved", "geo_saved"):
        idx = list(cdf.columns).index(colname) + 1
        ws.column_dimensions[L(idx)].hidden = True

    # Name map
    ws = wb.create_sheet("Name_Map")
    ws["A1"] = "Name_Map - tell the tool that a differently-spelled name is an existing community"
    ws["A1"].font = F_TITLE
    ws["A2"] = ("Add a row when the report renames a project (see POSSIBLE RENAME in Change_Log). "
                "Alias = name exactly as printed in the PDF; Market = LA-Orange-North / Inland Empire / "
                "San Diego-Imperial; Community ID = from the Communities tab. Re-import the PDF afterwards.")
    ws["A2"].font = F_NOTE
    _plain_table(ws, db["name_map"].reindex(columns=["alias", "market", "community_id"]), 3,
                 header_map=header_map, widths={"alias": 36, "market": 20, "community_id": 36},
                 input_cols=("alias", "market", "community_id"))

    # Manual entries (input)
    ws = wb.create_sheet("Manual_Entries")
    ws["A1"] = "Manual_Entries - weekly numbers for communities not on the weekly report"
    ws["A1"].font = F_TITLE
    ws["A2"] = ("For communities with Tracking = Manual (e.g. Toll Brothers). Week Ending = the Sunday of "
                "the report week; Community ID from the Communities tab. Saved here, merged on every import.")
    ws["A2"].font = F_NOTE
    man = db.get("manual")
    man = man if man is not None else pd.DataFrame(columns=["week_ending", "community_id", "gross_sales", "cancellations", "note"])
    man = man.copy()
    man["week_ending"] = pd.to_datetime(man["week_ending"], errors="coerce").dt.date
    _plain_table(ws, man.reindex(columns=["week_ending", "community_id", "gross_sales", "cancellations", "note"]),
                 3, header_map={"week_ending": "Week Ending", "community_id": "Community ID",
                                "gross_sales": "Gross Sales", "cancellations": "Cancellations", "note": "Note"},
                 widths={"community_id": 36, "note": 40}, date_cols=("week_ending",),
                 input_cols=("week_ending", "community_id", "gross_sales", "cancellations", "note"))
    mids = comm.loc[comm["tracking"].astype(str).str.lower().str.startswith("manual"), ["community_id", "community"]]
    ws.cell(3, 7, "Manual-tracked communities (reference)").font = F_BOLD
    for i, rr in enumerate(mids.itertuples(index=False), start=4):
        ws.cell(i, 7, rr.community_id).font = F_BASE
        ws.cell(i, 8, rr.community).font = F_BASE
    ws.column_dimensions["G"].width = 36
    ws.column_dimensions["H"].width = 30

    # Import log
    ws = wb.create_sheet("Import_Log")
    il = db["import_log"].copy()
    il["week_ending"] = pd.to_datetime(il["week_ending"], errors="coerce").dt.date
    _plain_table(ws, il, 1, widths={"warnings": 80, "source_file": 34, "markets": 44,
                                     "totals_check": 30, "imported_at": 16},
                 date_cols=("week_ending",))

    # Weekly data (long / DB-ready)
    ws = wb.create_sheet("Weekly_Data")
    wd = weekly.sort_values(["week_ending", "market", "community_id"])
    _plain_table(ws, wd, 1, name="WeeklyData", widths={"community_id": 34, "source": 18,
                                                        "source_file": 30, "market": 18},
                 date_cols=("week_ending",))

    for key, sheet in (("legacy_report", "Legacy_Match"), ("legacy_compare", "Legacy_Check")):
        if key in db and db[key] is not None and len(db[key]):
            ws = wb.create_sheet(sheet)
            _plain_table(ws, db[key], 1, widths={"legacy_name": 32, "matched_to": 32,
                                                 "community_id": 32, "method": 36, "community": 32})

    # Read me
    ws = wb.create_sheet("Read Me")
    lines = [
        ("SoCal New Home Sales Tracker - weekly import", F_TITLE),
        ("", F_BASE),
        ("How it works", F_BOLD),
        ("1. Each week, open the New Home Sales Weekly Import app, upload this workbook and the new weekly report PDF, review, and click Import.", F_BASE),
        ("2. The tool reads every community's Wk's Sales (Gross) and Wk's Cans (Cancellations); Net Sales = Gross - Cancellations.", F_BASE),
        ("3. Every section is checked against the PDF's own TOTALS row; results are in Import_Log.", F_BASE),
        ("4. Communities are never deleted. One that disappears from the report is flagged DROPPED, then shows SOLD OUT going forward.", F_BASE),
        ("", F_BASE),
        ("Tabs", F_BOLD),
        ("Gross Sales / Cancellations / Net Sales - weekly views, grouped by market and submarket, with 4/8/13-week totals and pace.", F_BASE),
        ("Net Sales also has Project Label (your custom name - edit here, in Communities, or in the app's Project Labels tab), "
         "Pace Label (sold-of-units plus 4-wk, 13-wk and LOJ pace) and LOJ Pace (average of all weekly net sales).", F_BASE),
        ("Change_Log - NEW, ADDED, DROPPED, RETURNED, FLAG S/O, FLAG TSO and POSSIBLE RENAME events by week.", F_BASE),
        ("Communities - master list with status. Yellow columns are yours to edit and are kept on every import.", F_BASE),
        ("Name_Map - aliases for renamed projects, and for GPN names that don't match automatically (alias exactly as GPN prints it, e.g. 'KB, Moonlight').", F_BASE),
        ("GPN report - Great Park Neighborhoods numbers replace the weekly report for the same project and week; shown in bold teal.", F_BASE),
        ("Latitude / Longitude - city center (grey italic) until found on OpenStreetMap (app's Locations tab) or typed in; typed values are kept. Geo Source in Communities says which.", F_BASE),
        ("Manual_Entries - weekly gross/cans for communities you track outside the weekly report (Tracking = Manual).", F_BASE),
        ("Weekly_Data - one row per community per week (the table to load into an online database later).", F_BASE),
        ("Import_Log - one row per import with totals and warnings.", F_BASE),
        ("Legacy_Match / Legacy_Check - how v4 history was matched, and v4 vs PDF for overlapping weeks.", F_BASE),
        ("", F_BASE),
        ("Note: the three weekly tabs are regenerated on each import; put your own notes in the Communities tab, not in those grids.", F_NOTE),
    ]
    for i, (t, f) in enumerate(lines, start=1):
        ws.cell(i, 1, t).font = f
    ws.column_dimensions["A"].width = 130

    wb.active = wb.sheetnames.index("Net Sales")
    for s in wb.worksheets:
        s.sheet_view.tabSelected = (s.title == "Net Sales")
    wb.calculation.fullCalcOnLoad = True   # Excel computes the Net Sales formulas on open
    wb.save(path)
