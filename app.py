"""
New Home Sales Weekly Import - Streamlit app
Run:  streamlit run app.py
"""
from __future__ import annotations

import io
import os
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import sales_db as sdb
from weekly_report_parser import parse_pdf
from gpn_parser import is_gpn_pdf, parse_gpn

st.set_page_config(page_title="New Home Sales Weekly Import", layout="wide")

DEFAULT_DB = os.environ.get("NHS_DB", str(Path(__file__).with_name("New_Home_Sales_Database.xlsx")))
# On Streamlit Community Cloud the app's disk is temporary: work from an uploaded copy and download it.
ON_CLOUD = str(Path(__file__).resolve()).startswith("/mount/src")


def copy_db(db):
    return {k: (v.copy() if isinstance(v, pd.DataFrame) else v) for k, v in db.items()}


def db_bytes(db) -> bytes:
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        tmp = f.name
    sdb.save_db(db, tmp)
    data = Path(tmp).read_bytes()
    os.unlink(tmp)
    return data


EVENT_COLORS = {"NEW": "#C6EFCE", "ADDED": "#DDEBF7", "RETURNED": "#FFEB9C",
                "DROPPED": "#FFC7CE", "POSSIBLE RENAME": "#E4DFEC"}


def style_events(df):
    def f(v):
        for k, c in EVENT_COLORS.items():
            if str(v).startswith(k):
                return f"background-color: {c}"
        return ""
    return df.style.map(f, subset=["event"]) if "event" in df else df


# ----------------------------------------------------------------- sidebar: database
st.sidebar.header("1. Database workbook")
mode = st.sidebar.radio("Where is the database?", ["File on this computer", "Upload a copy"],
                        # online, default to an uploaded copy - unless a workbook ships alongside the app
                        index=1 if (ON_CLOUD and not Path(DEFAULT_DB).exists()) else 0,
                        help="'File on this computer' reads and saves the workbook in place. "
                             "Online (Streamlit Cloud), use 'Upload a copy' and download the result.")
if ON_CLOUD and mode == "File on this computer":
    st.sidebar.warning("This app is running online: changes saved here are wiped when the app restarts. "
                       "Always click Download workbook when you're done.")
if mode == "File on this computer":
    db_path = st.sidebar.text_input("Workbook path", DEFAULT_DB)
    if st.sidebar.button("Load / reload", width="stretch") or "db" not in st.session_state:
        if Path(db_path).exists():
            st.session_state.db = sdb.load_db(db_path)
            st.session_state.db_name = Path(db_path).name
            st.session_state.dirty = False
        else:
            st.session_state.db = sdb.empty_db()
            st.session_state.db_name = Path(db_path).name
            st.sidebar.info("No workbook at that path yet - a new one will be created on import.")
else:
    db_path = None
    up = st.sidebar.file_uploader("Database workbook (.xlsx)", type=["xlsx"])
    if up is not None and st.session_state.get("db_upload_id") != up.file_id:
        st.session_state.db = sdb.load_db(io.BytesIO(up.getvalue()))
        st.session_state.db_name = up.name
        st.session_state.db_upload_id = up.file_id
        st.session_state.dirty = False
    if "db" not in st.session_state:
        st.session_state.db = sdb.empty_db()
        st.session_state.db_name = "New_Home_Sales_Database.xlsx"

db = st.session_state.db
comm, weekly = db["communities"], db["weekly"]
pdf_weeks = sorted(set(weekly.loc[weekly["source"].isin(sdb.PDF_SOURCES), "week_ending"]))
st.sidebar.metric("Communities tracked", len(comm))
st.sidebar.caption(f"Weekly reports imported: {len(pdf_weeks)}"
                   + (f" ({pdf_weeks[0]:%m/%d/%Y} to {pdf_weeks[-1]:%m/%d/%Y})" if pdf_weeks else ""))
if not len(comm):
    st.sidebar.info("No workbook loaded yet.")

def _clear_dirty():
    st.session_state.dirty = False


if len(comm):
    st.sidebar.divider()
    if st.session_state.get("dirty"):
        st.sidebar.warning("You have changes that aren't " + ("saved" if db_path else "downloaded") + " yet.")
    if db_path and st.sidebar.button("Save workbook", type="primary" if st.session_state.get("dirty") else "secondary",
                                     width="stretch"):
        sdb.save_db(st.session_state.db, db_path)
        st.session_state.dirty = False
        st.rerun()
    _snapshot = st.session_state.db
    st.sidebar.download_button("Download workbook", data=lambda: db_bytes(_snapshot),
                               file_name=st.session_state.get("db_name", "New_Home_Sales_Database.xlsx"),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               on_click=_clear_dirty, width="stretch",
                               type="primary" if (st.session_state.get("dirty") and not db_path) else "secondary")
    st.sidebar.caption("The download includes every change made in this session (imports, labels, locations).")

st.title("New Home Sales Weekly Import")
tab_imp, tab_lbl, tab_geo, tab_db, tab_legacy = st.tabs(["Import PDF", "Project Labels", "Locations", "Database", "One-time legacy import"])

# ----------------------------------------------------------------- import tab
with tab_imp:
    st.subheader("2. Upload weekly sales report and/or GPN report PDFs")
    pdfs = st.file_uploader("Weekly sales report and Great Park Neighborhoods (GPN) PDFs - any week, newer or older; "
                            "several can be loaded at once. The report type is detected automatically.",
                            type=["pdf"], accept_multiple_files=True)
    parsed, gpn_parsed = [], []
    for f in pdfs or []:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as t:
            t.write(f.getvalue())
            path = t.name
        if is_gpn_pdf(path):
            gpn_parsed.append((f.name, parse_gpn(path)))
        else:
            parsed.append((f.name, parse_pdf(path)))
        os.unlink(path)

    if parsed or gpn_parsed:
        parsed.sort(key=lambda x: x[1].week_ending or pd.Timestamp.min.date())
        gpn_parsed.sort(key=lambda x: x[1].week_ending or pd.Timestamp.min.date())
        preview = copy_db(db)
        summaries = []
        unreadable = [(n, r) for n, r in parsed if not r.records or r.week_ending is None]
        for name, res in unreadable:
            st.error(f"**{name}**: no community rows could be read from this PDF"
                     + (f" (week ending {res.week_ending:%m/%d/%Y})" if res.week_ending else "")
                     + ". Its layout differs from the reports the tool knows, so it was skipped. "
                     "Send a copy of this PDF so the reader can be updated for it.")
        parsed = [(n, r) for n, r in parsed if r.records and r.week_ending is not None]
        gpn_bad = [(n, r) for n, r in gpn_parsed if not r.records or r.week_ending is None]
        for name, res in gpn_bad:
            st.error(f"**{name}**: no collections could be read from this GPN report, so it was skipped.")
        gpn_parsed = [(n, r) for n, r in gpn_parsed if r.records and r.week_ending is not None]
        for name, res in parsed:
            df = res.to_frame()
            st.markdown(f"#### {name} &mdash; week ending **{res.week_ending:%A, %B %d, %Y}**")
            if res.week_ending in pdf_weeks:
                st.warning("This week is already in the database. Importing again replaces it.")
            elif pdf_weeks and res.week_ending < pdf_weeks[-1]:
                st.info(f"Older than the latest week in the database ({pdf_weeks[-1]:%m/%d/%Y}) - it will be "
                        "backfilled into its place in the timeline. Current community names, units and "
                        "statuses still come from the newest report.")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Communities", len(df))
            c2.metric("Gross sales", int(df.gross_sales.sum()))
            c3.metric("Cancellations", int(df.cancellations.sum()))
            c4.metric("Net sales", int(df.net_sales.sum()))
            checks = pd.DataFrame(res.section_checks)
            ok = int(checks["ok"].sum()) if len(checks) else 0
            if len(checks) and ok == len(checks):
                st.success(f"All {len(checks)} sections match the PDF's own TOTALS rows "
                           "(Wk's Sales and Wk's Cans).")
            else:
                st.error(f"{len(checks) - ok} of {len(checks)} sections do not match the PDF totals - review before importing.")
            for w in res.warnings:
                st.warning(w)
            if res.notes:
                with st.expander(f"{len(res.notes)} section(s) list more projects than the report counts as reporting (sales and cans still match)"):
                    for n in res.notes:
                        st.caption(n)
            with st.expander("Market totals and section checks"):
                st.dataframe(df.groupby("market", sort=False)[["gross_sales", "cancellations", "net_sales"]].sum(),
                             width="stretch")
                st.dataframe(checks, width="stretch", hide_index=True)
            with st.expander("Parsed rows"):
                st.dataframe(df.drop(columns=["week_ending"]), width="stretch", hide_index=True)
            s = sdb.import_report(preview, res, name)
            summaries.append(s)
            ch = s["changes"][["event", "market", "submarket", "community", "builder", "detail"]]
            st.markdown(f"**What changes:** {s['new']} new, {s['added']} added (first tracked), "
                        f"{s['dropped']} dropped (likely sold out), {s['returned']} returned")
            if len(ch):
                st.dataframe(style_events(ch), width="stretch", hide_index=True)
            st.divider()

        for name, res in gpn_parsed:
            df = res.to_frame()
            st.markdown(f"#### GPN &mdash; {name} &mdash; week ending **{res.week_ending:%A, %B %d, %Y}**")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Collections", len(df))
            c2.metric("Gross sales", int(df.gross_sales.sum()))
            c3.metric("Cancellations", int(df.cancellations.sum()))
            c4.metric("Net sales", int(df.net_sales.sum()))
            checks = pd.DataFrame(res.section_checks)
            ok = int(checks["ok"].sum()) if len(checks) else 0
            if len(checks) and ok == len(checks):
                st.success(f"All {len(checks)} neighborhoods match the report's own total rows (gross, cans, net).")
            else:
                st.error(f"{len(checks) - ok} of {len(checks)} neighborhoods do not match the report totals - review before importing.")
            for w in res.warnings:
                st.warning(w)
            s = sdb.import_gpn(preview, res, name)
            m = s["matches"]
            unmatched = m[m["community_id"].isna() & ~m["match"].str.startswith("skipped")]
            st.markdown(f"**Matching:** {int(m['community_id'].notna().sum())} matched to your communities, "
                        f"{s['new']} new GPN-only communities, "
                        f"{int(m['match'].str.startswith('skipped').sum())} long-sold-out collections skipped"
                        + (f", **{len(unmatched)} need a Name_Map row**" if len(unmatched) else ""))
            with st.expander("GPN name -> community matches", expanded=bool(len(unmatched))):
                st.dataframe(m[["gpn_name", "neighborhood", "community", "match"]], width="stretch", hide_index=True)
            if len(s["overrides"]):
                st.markdown("**GPN replaces these weekly report numbers for the week:**")
                st.dataframe(s["overrides"], width="stretch", hide_index=True)
            else:
                st.caption("Where both reports cover a project this week, the numbers agree (or the weekly report isn't imported yet for this week).")
            st.divider()

        st.subheader("3. Import")
        start_new = True
        if not len(db["communities"]):
            st.error("**No database workbook is loaded.** Importing now starts a NEW database that contains only "
                     "these reports - your existing weeks would not be in it. Load your workbook in the sidebar "
                     "first (Upload a copy), then import.")
            start_new = st.checkbox("I want to start a new, empty database with only these reports")
        if st.button("Import into database", type="primary", width="stretch",
                     disabled=not (parsed or gpn_parsed) or not start_new):
            st.session_state.db = preview
            st.session_state.dirty = True
            if db_path:
                Path(db_path).parent.mkdir(parents=True, exist_ok=True)
                sdb.save_db(preview, db_path)
                st.session_state.dirty = False
            st.session_state.just_imported = True
            st.rerun()

    if st.session_state.pop("just_imported", False):
        st.success("Import complete." + (f" Saved to {db_path}." if db_path else
                                         " Use **Download workbook** in the sidebar to keep it."))

# ----------------------------------------------------------------- database tab
with tab_db:
    if not len(comm):
        st.info("Database is empty - import a PDF first.")
    else:
        st.subheader("Status of all communities")
        st.dataframe(comm["status"].value_counts().rename("communities"), width="content")
        chg = db["change_log"]
        if len(chg):
            wk = st.selectbox("Change log for week", sorted(chg["week_ending"].unique(), reverse=True))
            st.dataframe(style_events(chg[chg["week_ending"] == wk]
                                      [["event", "market", "submarket", "community", "builder", "detail"]]),
                         width="stretch", hide_index=True)
        q = st.text_input("Find a community")
        view = comm if not q else comm[comm["community"].str.contains(q, case=False, na=False)]
        st.dataframe(view[["community_id", "market", "submarket", "community", "project_label", "builder", "city", "units",
                           "status", "tracking", "first_week", "last_reported"]],
                     width="stretch", hide_index=True)
        if q and len(view) == 1:
            cid = view.iloc[0]["community_id"]
            hist = weekly[weekly["community_id"] == cid].sort_values("week_ending")
            st.line_chart(hist.set_index("week_ending")[["net_sales"]])
        st.subheader("Export for an online database")
        c1, c2 = st.columns(2)
        c1.download_button("Weekly_Data.csv", weekly.to_csv(index=False).encode(), "weekly_data.csv", "text/csv")
        c2.download_button("Communities.csv", comm.to_csv(index=False).encode(), "communities.csv", "text/csv")

# ----------------------------------------------------------------- project labels tab
with tab_lbl:
    if not len(comm):
        st.info("Database is empty - import a PDF first.")
    else:
        st.markdown("Type your own name for any project in **Project Label** - each change is kept as soon as "
                    "you press Enter or click away, including when you change the filters. Labels show in "
                    "column F of the Net Sales tab. "
                    + ("Click **Save workbook** in the sidebar to write them to the file."
                       if db_path else "Use **Download workbook** in the sidebar to keep them."))
        with st.expander("Load labels from a spreadsheet"):
            st.caption("Any .xlsx with a header row that has **Community** (and ideally Builder and City) and a "
                       "**Project Label** column - for example a copy of the Net Sales tab.")
            lf = st.file_uploader("Label spreadsheet", type=["xlsx"], key="lbl_file")
            overwrite = st.checkbox("Replace labels that are already filled in", value=True)
            if lf is not None and st.button("Apply labels from file"):
                try:
                    rows = sdb.read_label_sheet(io.BytesIO(lf.getvalue()))
                    rep_ = sdb.import_labels(st.session_state.db, rows, overwrite=overwrite)
                    st.session_state.dirty = True
                    st.session_state.lbl_ver = st.session_state.get("lbl_ver", 0) + 1
                    msg = f"Applied {rep_['applied']} of {len(rows)} labels."
                    if rep_["skipped"]:
                        msg += f" {rep_['skipped']} kept their existing label."
                    st.session_state.lbl_msg = msg
                    st.session_state.lbl_unmatched = rep_["unmatched"]
                    st.rerun()
                except Exception as e:
                    st.error(f"Couldn't read that file: {e}")
        if "lbl_msg" in st.session_state:
            st.success(st.session_state.pop("lbl_msg"))
            um = st.session_state.pop("lbl_unmatched", [])
            if um:
                st.warning("Not matched to a community (check the name/builder/city): " + "; ".join(um))

        f1, f2, f3, f4 = st.columns([1.2, 1.6, 1.6, 1])
        mk = f1.selectbox("Market", ["All"] + sorted(comm["market"].dropna().unique()))
        pool = comm if mk == "All" else comm[comm["market"] == mk]
        sm = f2.selectbox("Submarket", ["All"] + sorted(pool["submarket"].dropna().unique()))
        pool = pool if sm == "All" else pool[pool["submarket"] == sm]
        q = f3.text_input("Search community / builder", key="lbl_q")
        if q:
            pool = pool[pool["community"].str.contains(q, case=False, na=False) |
                        pool["builder"].astype(str).str.contains(q, case=False, na=False)]
        only_blank = f4.checkbox("Only unlabeled")
        if only_blank:
            pool = pool[pool["project_label"].isna() | (pool["project_label"].astype(str).str.strip() == "")]
        hide_so = f4.checkbox("Hide sold out", value=False)
        if hide_so:
            pool = pool[~pool["status"].astype(str).str.startswith("SOLD OUT")]
        pool = pool.assign(**{k: pd.to_numeric(pool[k], errors="coerce") for k in ("m_ord", "s_ord", "r_ord")})
        view = pool.sort_values(["m_ord", "s_ord", "r_ord"])[
            ["community_id", "submarket", "community", "builder", "city", "product_type", "status", "project_label"]
        ].reset_index(drop=True)
        st.caption(f"{len(view)} communities shown - "
                   f"{int(comm['project_label'].notna().sum())} of {len(comm)} have a label")
        ekey = f"lbl_editor_{st.session_state.get('lbl_ver', 0)}_{mk}_{sm}_{q}_{only_blank}_{hide_so}"
        st.session_state.lbl_ids = list(view["community_id"])

        def _apply_label_edits(ekey=ekey):
            edits = (st.session_state.get(ekey) or {}).get("edited_rows", {})
            ids = st.session_state.get("lbl_ids", [])
            ch = {ids[int(i)]: (v.get("project_label") or "") for i, v in edits.items()
                  if "project_label" in v and int(i) < len(ids)}
            if ch:
                sdb.set_project_labels(st.session_state.db, ch)
                st.session_state.dirty = True
                # fresh editor next run, so row positions can't drift when a filter hides the row
                st.session_state.lbl_ver = st.session_state.get("lbl_ver", 0) + 1

        st.data_editor(
            view, hide_index=True, width="stretch", height=520, key=ekey, on_change=_apply_label_edits,
            disabled=["community_id", "submarket", "community", "builder", "city", "product_type", "status"],
            column_config={"project_label": st.column_config.TextColumn("Project Label", width="large"),
                           "community_id": None, "product_type": "Type"})


# ----------------------------------------------------------------- locations tab
with tab_geo:
    import geocode as gc
    if not len(comm):
        st.info("Database is empty - import a PDF first.")
    else:
        src = comm["geo_source"].fillna("None").astype(str)
        osm = int(src.str.startswith(gc.OSM_PREFIX).sum())
        man = int((src == gc.MANUAL_SOURCE).sum())
        cc = int((src == gc.CITY_SOURCE).sum())
        st.markdown("The weekly report has no street addresses, so every project starts at its **city center** "
                    "(grey italic in the weekly tabs). The button below looks each one up by name on "
                    "OpenStreetMap; anything you type in yourself is kept and never overwritten.")
        c1, c2, c3 = st.columns(3)
        c1.metric("Typed in (Manual)", man)
        c2.metric("Found on OpenStreetMap", osm)
        c3.metric("City center only", cc)

        st.markdown("##### Find project locations online")
        f1, f2 = st.columns([2, 1])
        scope_mk = f1.selectbox("Market", ["All"] + sorted(comm["market"].dropna().unique()), key="geo_mk")
        active_only = f2.checkbox("Skip sold-out projects", value=True)
        todo = comm[src == gc.CITY_SOURCE]
        if scope_mk != "All":
            todo = todo[todo["market"] == scope_mk]
        if active_only:
            todo = todo[~todo["status"].astype(str).str.startswith("SOLD OUT")]
        st.caption(f"{len(todo)} projects to look up - about {max(1, round(len(todo) * 2.3 / 60))} min "
                   "(OpenStreetMap allows one lookup per second). Needs an internet connection.")
        if st.button("Find locations", disabled=not len(todo)):
            bar = st.progress(0.0)
            found, failed = {}, 0
            fails = 0
            for k, r in enumerate(todo.itertuples()):
                try:
                    hit = gc.find_location(r.community, r.city)
                    fails = 0
                except ConnectionError as e:
                    hit, fails = None, fails + 1
                    if fails >= 3:
                        st.error(f"{e}. Check the internet connection and try again; "
                                 f"{len(found)} location(s) found so far are kept.")
                        break
                if hit:
                    found[r.community_id] = hit
                bar.progress((k + 1) / len(todo), text=f"{k + 1}/{len(todo)} - {r.community}")
                if (k + 1) % 50 == 0 and found and db_path:      # save progress as it goes
                    sdb.set_locations(st.session_state.db, found)
                    sdb.save_db(st.session_state.db, db_path)
            sdb.set_locations(st.session_state.db, found)
            sdb.finalize(st.session_state.db)
            st.session_state.dirty = True
            if db_path:
                sdb.save_db(st.session_state.db, db_path)
                st.session_state.dirty = False
            st.session_state.geo_msg = (f"Found {len(found)} of {len(todo)} on OpenStreetMap; "
                                        f"the rest stay at their city center.")
            st.rerun()
        if "geo_msg" in st.session_state:
            st.success(st.session_state.pop("geo_msg"))

        st.markdown("##### Check or type coordinates")
        st.caption("Tip: in Google Maps, right-click the project and click the coordinates to copy them.")
        q = st.text_input("Search community / builder / city", key="geo_q")
        view = comm if not q else comm[comm["community"].str.contains(q, case=False, na=False) |
                                       comm["builder"].astype(str).str.contains(q, case=False, na=False) |
                                       comm["city"].astype(str).str.contains(q, case=False, na=False)]
        view = view[["community_id", "community", "builder", "city", "latitude", "longitude", "geo_source"]] \
            .assign(latitude=lambda d: pd.to_numeric(d["latitude"], errors="coerce"),
                    longitude=lambda d: pd.to_numeric(d["longitude"], errors="coerce")).reset_index(drop=True)
        gkey = f"geo_ed_{st.session_state.get('geo_ver', 0)}_{q}"
        st.session_state.geo_rows = view[["community_id", "latitude", "longitude"]].to_dict("records")

        def _apply_geo_edits(gkey=gkey):
            edits = (st.session_state.get(gkey) or {}).get("edited_rows", {})
            rows = st.session_state.get("geo_rows", [])
            ch = {}
            for i, v in edits.items():
                i = int(i)
                if i >= len(rows):
                    continue
                r = rows[i]
                lat = sdb._f(v.get("latitude", r["latitude"]))
                lon = sdb._f(v.get("longitude", r["longitude"]))
                ch[r["community_id"]] = (lat, lon, gc.MANUAL_SOURCE)
            if ch:
                sdb.set_locations(st.session_state.db, ch)
                st.session_state.dirty = True
                st.session_state.geo_ver = st.session_state.get("geo_ver", 0) + 1

        st.data_editor(view, hide_index=True, width="stretch", height=420, key=gkey, on_change=_apply_geo_edits,
                       disabled=["community_id", "community", "builder", "city", "geo_source"],
                       column_config={"community_id": None,
                                      "latitude": st.column_config.NumberColumn("Latitude", format="%.8f"),
                                      "longitude": st.column_config.NumberColumn("Longitude", format="%.8f"),
                                      "geo_source": "Source"})
        st.caption("Each coordinate you type is kept as soon as you press Enter (marked Manual).")
        if len(comm):
            pts = comm.assign(lat=pd.to_numeric(comm["latitude"], errors="coerce"),
                              lon=pd.to_numeric(comm["longitude"], errors="coerce")).dropna(subset=["lat", "lon"])
            st.map(pts[["lat", "lon"]], size=40)


# ----------------------------------------------------------------- legacy tab
with tab_legacy:
    st.markdown("Merge net-sales history from the old **New Home Sales Trends** workbook. "
                "Safe to run again: it re-applies the v4 history (it won't duplicate communities, and your "
                "labels and notes are kept). Re-run it once if communities you track outside the weekly report "
                "(e.g. Toll at Great Park) are missing v4 weeks.")
    lf = st.file_uploader("Old trends workbook (.xlsx)", type=["xlsx"], key="legacy")
    if lf is not None and st.button("Merge legacy history"):
        rep, cmp_ = sdb.import_legacy(st.session_state.db, io.BytesIO(lf.getvalue()), lf.name)
        st.session_state.db["legacy_report"], st.session_state.db["legacy_compare"] = rep, cmp_
        st.session_state.dirty = not db_path
        if db_path:
            sdb.save_db(st.session_state.db, db_path)
        st.dataframe(rep["method"].value_counts())
        st.dataframe(rep, width="stretch", hide_index=True)
