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
                        help="'File on this computer' reads and saves the workbook in place.")
if mode == "File on this computer":
    db_path = st.sidebar.text_input("Workbook path", DEFAULT_DB)
    if st.sidebar.button("Load / reload", width="stretch") or "db" not in st.session_state:
        if Path(db_path).exists():
            st.session_state.db = sdb.load_db(db_path)
            st.session_state.db_name = Path(db_path).name
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
    if "db" not in st.session_state:
        st.session_state.db = sdb.empty_db()
        st.session_state.db_name = "New_Home_Sales_Database.xlsx"

db = st.session_state.db
comm, weekly = db["communities"], db["weekly"]
pdf_weeks = sorted(set(weekly.loc[weekly["source"].isin(sdb.PDF_SOURCES), "week_ending"]))
st.sidebar.metric("Communities tracked", len(comm))
st.sidebar.caption(f"Weekly reports imported: {len(pdf_weeks)}"
                   + (f" (latest {pdf_weeks[-1]:%m/%d/%Y})" if pdf_weeks else ""))

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
        if st.button("Import into database", type="primary", width="stretch"):
            st.session_state.db = preview
            st.session_state.out_bytes = db_bytes(preview)
            if db_path:
                Path(db_path).parent.mkdir(parents=True, exist_ok=True)
                sdb.save_db(preview, db_path)
                st.success(f"Saved to {db_path}")
            st.session_state.just_imported = True
            st.rerun()

    if st.session_state.pop("just_imported", False):
        st.success("Import complete. Download the updated workbook below (or open it from its path).")

    if "out_bytes" in st.session_state:
        st.download_button("Download updated workbook", data=st.session_state.out_bytes,
                           file_name=st.session_state.get("db_name", "New_Home_Sales_Database.xlsx"),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width="stretch")

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
        st.markdown("Type your own name for any project in **Project Label**, then click **Save labels**. "
                    "Labels show in column F of the Net Sales tab and are kept on every import. "
                    "(You can also type them straight into the Net Sales or Communities tab in Excel - "
                    "the app picks those up the next time it loads the workbook.)")
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
        edited = st.data_editor(
            view, hide_index=True, width="stretch", height=520, key=f"lbl_editor_{mk}_{sm}_{q}_{only_blank}_{hide_so}",
            disabled=["community_id", "submarket", "community", "builder", "city", "product_type", "status"],
            column_config={"project_label": st.column_config.TextColumn("Project Label", width="large"),
                           "community_id": None, "product_type": "Type"})
        before = dict(zip(view["community_id"], view["project_label"].fillna("").astype(str)))
        after = dict(zip(edited["community_id"], edited["project_label"].fillna("").astype(str)))
        changed = {k: v for k, v in after.items() if v.strip() != before.get(k, "").strip()}
        st.write(f"{len(changed)} unsaved change(s)" if changed else "No unsaved changes")
        if st.button("Save labels", type="primary", disabled=not changed):
            sdb.set_project_labels(st.session_state.db, changed)
            st.session_state.out_bytes = db_bytes(st.session_state.db)
            if db_path:
                sdb.save_db(st.session_state.db, db_path)
            st.session_state.lbl_msg = (f"Saved {len(changed)} label(s)" +
                                        (f" to {db_path}" if db_path else " - download the updated workbook on the Import PDF tab."))
            st.rerun()
        if "lbl_msg" in st.session_state:
            st.success(st.session_state.pop("lbl_msg"))


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
            st.session_state.out_bytes = db_bytes(st.session_state.db)
            if db_path:
                sdb.save_db(st.session_state.db, db_path)
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
        ed = st.data_editor(view, hide_index=True, width="stretch", height=420, key=f"geo_ed_{q}",
                            disabled=["community_id", "community", "builder", "city", "geo_source"],
                            column_config={"community_id": None,
                                           "latitude": st.column_config.NumberColumn("Latitude", format="%.8f"),
                                           "longitude": st.column_config.NumberColumn("Longitude", format="%.8f"),
                                           "geo_source": "Source"})
        chg = {}
        for a, b in zip(view.itertuples(), ed.itertuples()):
            if not (sdb._same(sdb._f(a.latitude), sdb._f(b.latitude)) and sdb._same(sdb._f(a.longitude), sdb._f(b.longitude))):
                chg[b.community_id] = (sdb._f(b.latitude), sdb._f(b.longitude), gc.MANUAL_SOURCE)
        if st.button("Save coordinates", type="primary", disabled=not chg):
            sdb.set_locations(st.session_state.db, chg)
            st.session_state.out_bytes = db_bytes(st.session_state.db)
            if db_path:
                sdb.save_db(st.session_state.db, db_path)
            st.session_state.geo_msg = f"Saved {len(chg)} location(s)."
            st.rerun()
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
        if db_path:
            sdb.save_db(st.session_state.db, db_path)
        st.session_state.out_bytes = db_bytes(st.session_state.db)
        st.dataframe(rep["method"].value_counts())
        st.dataframe(rep, width="stretch", hide_index=True)
