import os
import io
import base64
import yaml
from yaml.loader import SafeLoader
import streamlit as st
import streamlit_authenticator as stauth
import psycopg2
import psycopg2.pool
import pandas as pd
from datetime import date, datetime, timezone, timedelta
import plotly.express as px
import plotly.graph_objects as go
from streamlit_autorefresh import st_autorefresh

# ----------------------------------------------------------------------------
# TIMEZONE — Pakistan Standard Time, fixed UTC+5 (no DST observed)
# ----------------------------------------------------------------------------
PKT = timezone(timedelta(hours=5))


def now_pkt():
    return datetime.now(PKT)


def today_pkt():
    return now_pkt().date()


def parse_dt(iso_str):
    """Parse a stored ISO timestamp; treat old naive timestamps as already-PKT."""
    dt = datetime.fromisoformat(iso_str)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=PKT)
    return dt


# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------
st.set_page_config(page_title="Loss Time Tracker | Artistic Milliners", page_icon="🏭", layout="wide")

DATABASE_URL_SECRET = "DATABASE_URL"  # set this in Streamlit Secrets (Neon connection string)
LOGO_PATH = "logo.jpg"
COMPANY_NAME = "Artistic Milliners"
ALL_UNITS = ["AM-4CT2", "AM-4A"]
DEFAULT_UNIT = "AM-4CT2"  # fallback for any account whose config is missing a "units" list

CATEGORIES = [
    "Trim Issues / Fabric Issue",
    "WIP/Feeding Issues",
    "Cutting Delays",
    "Quality Issue",
    "Maintenance Issue",
    "Changeover Time",
    "Technical Issue",
    "Line Discipline",
    "Other",
]

LINES = [f"Line # {i}" for i in range(1, 7)]

PALETTE = {
    "navy": "#1a3c6e",
    "teal": "#0e8a7d",
    "amber": "#e08e0b",
    "plum": "#8e44ad",
    "rose": "#c0392b",
    "sky": "#2980b9",
    "green": "#27ae60",
    "gray": "#8a94a6",
}
CATEGORY_COLORS = px.colors.qualitative.Bold
CHART_H = 330                     # bigger, readable charts (some scrolling is fine)
CHART_MARGIN = dict(t=40, b=40, l=20, r=20)

# ----------------------------------------------------------------------------
# THEME / CSS
# ----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .stApp { background-color: #ffffff; }
    .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        max-width: 1500px;
    }
    section[data-testid="stSidebar"] { background-color: #f7f9fc; }
    div[data-testid="stForm"] {
        background-color: #fbfcfe;
        border: 1px solid #e3e8f0;
        border-radius: 12px;
        padding: 16px 20px 6px 20px;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #f0f2f6;
        border-radius: 8px 8px 0 0;
        padding: 7px 16px;
        font-weight: 600;
        font-size: 14px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1a3c6e !important;
        color: #ffffff !important;
    }
    div.stButton > button, div.stFormSubmitButton > button {
        border-radius: 8px;
        font-weight: 700;
        padding: 0.35rem 1.2rem;
    }
    [data-testid="stMetricValue"] { font-size: 22px; }
    </style>
    """,
    unsafe_allow_html=True,
)


def section_header(text, color=PALETTE["navy"], emoji=""):
    st.markdown(
        f"<div style='color:{color}; border-bottom:2px solid {color}; "
        f"padding-bottom:4px; margin-top:14px; margin-bottom:8px; "
        f"font-weight:700; font-size:16px;'>{emoji} {text}</div>",
        unsafe_allow_html=True,
    )


def kpi_card(label, value, color1, color2):
    st.markdown(
        f"""
        <div style="background: linear-gradient(135deg,{color1},{color2});
                    padding:14px 18px; border-radius:12px; text-align:center;
                    color:#ffffff; box-shadow: 0 4px 10px rgba(0,0,0,0.12);">
            <div style="font-size:12px; font-weight:700; letter-spacing:0.5px;
                        text-transform:uppercase; opacity:0.92;">{label}</div>
            <div style="font-size:26px; font-weight:800; margin-top:4px;">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data
def _logo_data_uri():
    if not os.path.exists(LOGO_PATH):
        return None
    with open(LOGO_PATH, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    return f"data:image/jpeg;base64,{b64}"


def brand_header(center=False, unit=None, subtitle=None):
    """Single self-contained flex block: logo + wordmark.
    `unit` shows an 'Unit: X' line (used once logged in).
    `subtitle` shows plain text instead (used on the login screen, unit not chosen yet).
    Built as one single-line HTML string on purpose — a blank line inside an
    unsafe_allow_html block makes Streamlit's markdown parser bail out of raw-HTML
    mode partway through and print a stray closing tag as visible text."""
    logo_uri = _logo_data_uri()
    logo_html = f"<img src='{logo_uri}' style='height:52px; width:auto; flex-shrink:0;'/>" if logo_uri else ""
    justify = "center" if center else "flex-start"
    if subtitle:
        extra_line = f"<div style='font-size:13px; color:#666; font-weight:600;'>{subtitle}</div>"
    elif unit:
        extra_line = f"<div style='font-size:13px; color:#666; font-weight:600;'>Unit: {unit}</div>"
    else:
        extra_line = ""
    html = (
        f"<div style='display:flex; align-items:center; gap:16px; justify-content:{justify}; padding:4px 0 10px 0;'>"
        f"{logo_html}"
        f"<div style='line-height:1.35;'>"
        f"<div style='font-size:12px; letter-spacing:1.5px; color:{PALETTE['gray']}; font-weight:700; text-transform:uppercase;'>{COMPANY_NAME}</div>"
        f"<div style='font-size:23px; font-weight:800; color:{PALETTE['navy']};'>Line Loss Time Tracker</div>"
        f"{extra_line}"
        f"</div>"
        f"</div>"
    )
    st.markdown(html, unsafe_allow_html=True)


def fmt_ddmmyyyy(iso_str):
    return datetime.strptime(iso_str, "%Y-%m-%d").strftime("%d-%m-%Y")


def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Loss Time Data")
    return buf.getvalue()


# ----------------------------------------------------------------------------
# AUTHENTICATION
# ----------------------------------------------------------------------------
def load_auth_config():
    try:
        has_secrets = "credentials" in st.secrets
    except Exception:
        has_secrets = False
    if has_secrets:
        credentials = {
            "usernames": {u: dict(v) for u, v in st.secrets["credentials"]["usernames"].items()}
        }
        cookie = dict(st.secrets["cookie"])
    else:
        with open("auth_config.yaml") as f:
            cfg = yaml.load(f, Loader=SafeLoader)
        credentials = cfg["credentials"]
        cookie = cfg["cookie"]
    return credentials, cookie


credentials, cookie_cfg = load_auth_config()


@st.cache_resource
def get_authenticator(_credentials, cookie_name, cookie_key, expiry_days):
    # Leading underscore on _credentials tells Streamlit not to hash this arg (it's a
    # dict). Caching this means bcrypt only hashes every account's password once per
    # app process, not on every single click/filter change — with 14+ accounts that
    # repeated hashing was adding real, noticeable delay to every interaction.
    return stauth.Authenticate(_credentials, cookie_name, cookie_key, expiry_days)


authenticator = get_authenticator(
    credentials, cookie_cfg["name"], cookie_cfg["key"], cookie_cfg.get("expiry_days", 7),
)

if not st.session_state.get("authentication_status"):
    l1, l2, l3 = st.columns([1, 1.3, 1])
    with l2:
        brand_header(center=True, subtitle="SBU AM-4A")
        st.write("")
        authenticator.login(location="main")

auth_status = st.session_state.get("authentication_status")

if auth_status is False:
    st.error("Username or password is incorrect.")
    st.stop()
elif auth_status is None:
    st.stop()

username = st.session_state.get("username")
display_name = st.session_state.get("name")
user_cfg = credentials["usernames"].get(username, {})
user_role = user_cfg.get("role", "viewer")
can_enter_data = user_role == "entry"

# Units this account is allowed to see. Missing "units" key -> single default unit
# (keeps any old-style account config working without edits).
allowed_units = user_cfg.get("units") or [DEFAULT_UNIT]
allowed_units = [u for u in allowed_units if u in ALL_UNITS] or [DEFAULT_UNIT]

with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=60)
    st.markdown(
        f"<div style='font-size:11px; letter-spacing:1.5px; color:{PALETTE['gray']}; "
        f"font-weight:700; text-transform:uppercase;'>{COMPANY_NAME}</div>",
        unsafe_allow_html=True,
    )
    st.markdown(f"**{display_name}**")

    if len(allowed_units) > 1:
        active_unit = st.selectbox("Active Unit", allowed_units, key="active_unit_select")
    else:
        active_unit = allowed_units[0]

    authenticator.logout("Logout", "sidebar")

# ----------------------------------------------------------------------------
# DATABASE HELPERS  (Postgres / Neon) — pooled connections, schema set up once
# ----------------------------------------------------------------------------
def _database_url():
    try:
        return st.secrets[DATABASE_URL_SECRET]
    except Exception:
        st.error(
            "No database connection configured. Add a `DATABASE_URL` entry "
            "(your Neon connection string) to this app's Secrets and reload."
        )
        st.stop()


@st.cache_resource
def get_pool():
    """One pooled set of connections, reused across reruns and sessions —
    avoids a fresh network handshake to Neon on every single query."""
    p = psycopg2.pool.ThreadedConnectionPool(1, 10, _database_url())
    conn = p.getconn()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS loss_time (
                    id SERIAL PRIMARY KEY,
                    entry_date TEXT NOT NULL,
                    line TEXT NOT NULL,
                    category TEXT NOT NULL,
                    reason TEXT,
                    minutes REAL,
                    recorded_at TEXT NOT NULL,
                    recorded_by TEXT,
                    status TEXT DEFAULT 'closed',
                    start_time TEXT,
                    end_time TEXT
                )
                """
            )
            for ddl in [
                "ALTER TABLE loss_time ADD COLUMN IF NOT EXISTS recorded_by TEXT",
                "ALTER TABLE loss_time ADD COLUMN IF NOT EXISTS status TEXT DEFAULT 'closed'",
                "ALTER TABLE loss_time ADD COLUMN IF NOT EXISTS start_time TEXT",
                "ALTER TABLE loss_time ADD COLUMN IF NOT EXISTS end_time TEXT",
                "ALTER TABLE loss_time ADD COLUMN IF NOT EXISTS minutes_lost REAL",
                "ALTER TABLE loss_time ADD COLUMN IF NOT EXISTS workstations_affected REAL",
                "ALTER TABLE loss_time ADD COLUMN IF NOT EXISTS unit TEXT DEFAULT 'AM-4CT2'",
            ]:
                cur.execute(ddl)
        conn.commit()
    finally:
        p.putconn(conn)
    return p


def get_conn():
    return get_pool().getconn()


def release_conn(conn):
    get_pool().putconn(conn)


def add_closed_entry(entry_date, line, category, reason, minutes_lost, workstations_affected, recorded_by, unit):
    total_minutes = round(minutes_lost * workstations_affected, 1)
    now = now_pkt().isoformat(timespec="seconds")
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO loss_time (entry_date, line, category, reason, minutes, minutes_lost, "
                "workstations_affected, recorded_at, recorded_by, status, start_time, end_time, unit) "
                "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s, 'closed', %s, %s, %s)",
                (entry_date, line, category, reason, total_minutes, minutes_lost, workstations_affected,
                 now, recorded_by, now, now, unit),
            )
        conn.commit()
    finally:
        release_conn(conn)
    load_data.clear()


def start_open_issue(line, category, reason, workstations_affected, recorded_by, unit):
    now = now_pkt().isoformat(timespec="seconds")
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO loss_time (entry_date, line, category, reason, minutes, minutes_lost, "
                "workstations_affected, recorded_at, recorded_by, status, start_time, end_time, unit) "
                "VALUES (%s,%s,%s,%s,NULL,NULL,%s,%s,%s, 'open', %s, NULL, %s)",
                (today_pkt().isoformat(), line, category, reason, workstations_affected, now, recorded_by, now, unit),
            )
        conn.commit()
    finally:
        release_conn(conn)
    load_data.clear()


def resolve_issue(entry_id):
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT start_time, workstations_affected FROM loss_time WHERE id=%s", (entry_id,))
            row = cur.fetchone()
            if row:
                start = parse_dt(row[0])
                workstations_affected = row[1] or 1
                end = now_pkt()
                elapsed_minutes = round((end - start).total_seconds() / 60, 1)
                total_minutes = round(elapsed_minutes * workstations_affected, 1)
                cur.execute(
                    "UPDATE loss_time SET end_time=%s, minutes_lost=%s, minutes=%s, status='closed' WHERE id=%s",
                    (end.isoformat(timespec="seconds"), elapsed_minutes, total_minutes, entry_id),
                )
        conn.commit()
    finally:
        release_conn(conn)
    load_data.clear()


@st.cache_data(ttl=15, show_spinner=False)
def load_data(unit):
    """Only pulls rows for the active unit (not the whole table), and caches the
    result for a few seconds so ordinary reruns (typing, filters, switching tabs)
    reuse it instead of hitting Neon every time. Writes below call load_data.clear()
    so the very next rerun after a save always shows fresh data."""
    conn = get_conn()
    try:
        df = pd.read_sql_query(
            "SELECT * FROM loss_time WHERE COALESCE(unit, %s) = %s ORDER BY id DESC",
            conn, params=(DEFAULT_UNIT, unit),
        )
    finally:
        release_conn(conn)
    return df


def delete_entry(entry_id):
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM loss_time WHERE id=%s", (entry_id,))
        conn.commit()
    finally:
        release_conn(conn)
    load_data.clear()



if "selected_date" not in st.session_state:
    st.session_state.selected_date = today_pkt()

st_autorefresh(interval=30 * 60_000, key="global_refresh")  # refresh only every 30 min; otherwise only on manual reload

# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
brand_header(unit=active_unit)
st.markdown("<hr style='margin-top:2px; margin-bottom:10px;'>", unsafe_allow_html=True)

if can_enter_data:
    tab_entry, tab_dashboard = st.tabs(["📝 Data Entry", "📊 Dashboard"])
else:
    tab_dashboard = st.container()
    tab_entry = None

# One query per page render, shared by every tab below (instead of each tab querying separately).
# Scoped to the active unit at the SQL level (and cached for a few seconds) so every
# tab below (entry, today's log, all records, dashboard) only ever sees/writes data
# for that one unit, and ordinary reruns don't re-hit the database each time.
df_all = load_data(active_unit)

# ----------------------------------------------------------------------------
# DATA ENTRY (entry-role only)
# ----------------------------------------------------------------------------
if can_enter_data:
    with tab_entry:
        sub_new, sub_ongoing, sub_today, sub_records = st.tabs(
            ["➕ New Entry", "🔴 Ongoing Issues", "📋 Today's Log", "🗂️ All Records"]
        )

        # ---- Quick log: issue already resolved ----
        with sub_new:
            cdate, _ = st.columns([1, 2])
            with cdate:
                section_header("Date", PALETTE["sky"], "📅")
                st.session_state.selected_date = st.date_input(
                    "Production date", value=st.session_state.selected_date,
                    format="DD-MM-YYYY", label_visibility="collapsed",
                )

            section_header("Log a Resolved Loss Time Event", PALETTE["amber"], "⏱️")
            st.caption("Use this for issues that are already over. For an issue still happening, use the **Ongoing Issues** tab instead.")

            if "entry_form_key" not in st.session_state:
                st.session_state.entry_form_key = 0
            fk = st.session_state.entry_form_key

            c1, c2 = st.columns(2)
            with c1:
                line = st.selectbox("Select Line", LINES, key=f"line_{fk}")
            with c2:
                category = st.selectbox("Select Reason Category", CATEGORIES, key=f"category_{fk}")
            reason = st.text_area("Detailed Reason", height=80, key=f"reason_{fk}",
                                   placeholder="e.g. Waiting for trim from store...")
            c3, c4 = st.columns(2)
            with c3:
                minutes_lost = st.number_input(
                    "Minutes Lost", min_value=0.0, step=1.0, value=None, format="%.0f",
                    placeholder="Enter minutes lost", key=f"minutes_lost_{fk}",
                )
            with c4:
                workstations_affected = st.number_input(
                    "No. of Work-stations Affected", min_value=1, step=1, value=None, format="%d",
                    placeholder="Enter no. of work-stations", key=f"workstations_{fk}",
                )

            if minutes_lost is not None and workstations_affected is not None:
                total_preview = minutes_lost * workstations_affected
                st.markdown(
                    f"<div style='background:#f0f2f6; border-radius:10px; padding:10px 14px; "
                    f"margin:6px 0; font-size:14px;'>"
                    f"<b>Total Minutes Lost</b> = {minutes_lost:.0f} × {workstations_affected:.0f} "
                    f"= <b style='color:{PALETTE['rose']};'>{total_preview:.0f}</b></div>",
                    unsafe_allow_html=True,
                )
            else:
                st.caption("Total Minutes Lost will be calculated here once both fields above are filled in.")

            if st.button("Submit", use_container_width=True, type="primary", key=f"submit_btn_{fk}"):
                if minutes_lost is None or minutes_lost <= 0:
                    st.error("Please enter minutes lost greater than 0.")
                elif workstations_affected is None or workstations_affected < 1:
                    st.error("Please enter the number of work-stations affected.")
                elif not reason.strip():
                    st.error("Please enter a detailed reason.")
                else:
                    add_closed_entry(
                        st.session_state.selected_date.isoformat(),
                        line, category, reason.strip(), minutes_lost, workstations_affected, display_name,
                        active_unit,
                    )
                    total_minutes = minutes_lost * workstations_affected
                    st.success(f"Recorded: {line} | {category} | {total_minutes:.0f} total min "
                               f"({minutes_lost:.0f} × {workstations_affected:.0f})")
                    st.session_state.entry_form_key += 1
                    st.rerun()

        # ---- Ongoing issues: start now / resolve later ----
        with sub_ongoing:
            section_header("Start a New Ongoing Issue", PALETTE["rose"], "🔴")
            st.caption("Log it the moment it starts. It will show live on the Dashboard until you resolve it.")
            with st.form("start_form", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    o_line = st.selectbox("Select Line", LINES, key="o_line")
                with c2:
                    o_category = st.selectbox("Select Reason Category", CATEGORIES, key="o_cat")
                o_reason = st.text_area("Detailed Reason", height=70, key="o_reason",
                                         placeholder="e.g. Main sewing machine motor failure...")
                o_workstations = st.number_input(
                    "No. of Work-stations Affected", min_value=1, step=1, value=None, format="%d",
                    placeholder="Enter no. of work-stations", key="o_workstations",
                )
                start_submitted = st.form_submit_button("🔴 Start Issue", use_container_width=True, type="primary")
                if start_submitted:
                    if not o_reason.strip():
                        st.error("Please enter a detailed reason.")
                    elif o_workstations is None or o_workstations < 1:
                        st.error("Please enter the number of work-stations affected.")
                    else:
                        start_open_issue(o_line, o_category, o_reason.strip(), o_workstations, display_name, active_unit)
                        st.success(f"Started: {o_line} | {o_category} — now live on the Dashboard.")
                        st.rerun()

            section_header("Currently Ongoing", PALETTE["rose"], "⏳")
            open_df = df_all[df_all["status"] == "open"].copy()
            if open_df.empty:
                st.info("No ongoing issues right now.")
            else:
                now = now_pkt()
                for _, row in open_df.iterrows():
                    elapsed = round((now - parse_dt(row["start_time"])).total_seconds() / 60)
                    ws = row["workstations_affected"] or 1
                    running_total = elapsed * ws
                    ic1, ic2 = st.columns([5, 1])
                    with ic1:
                        st.markdown(
                            f"🔴 **{row['line']}** — {row['category']} — *{row['reason']}* "
                            f"— running **{elapsed} min** × **{ws:.0f} work-stations** "
                            f"= **{running_total:.0f} total min so far** "
                            f"(started {parse_dt(row['start_time']).strftime('%H:%M')})"
                        )
                    with ic2:
                        if st.button("Resolve", key=f"resolve_{row['id']}"):
                            resolve_issue(row["id"])
                            st.rerun()

        # ---- Today's log ----
        with sub_today:
            df_today = df_all[(df_all["entry_date"] == st.session_state.selected_date.isoformat())
                               & (df_all["status"] == "closed")]

            if not df_today.empty:
                kpi_card("Total Lost Minutes Today", f"{df_today['minutes'].sum():,.0f}",
                         PALETTE["rose"], "#e67e22")
                st.write("")
                disp = df_today[["id", "line", "category", "reason", "minutes_lost",
                                  "workstations_affected", "minutes", "entry_date"]].copy()
                disp["entry_date"] = disp["entry_date"].apply(fmt_ddmmyyyy)
                disp = disp.rename(columns={
                    "minutes_lost": "Minutes Lost", "workstations_affected": "Work-stations Affected",
                    "minutes": "Total Minutes Lost",
                })
                st.dataframe(disp, use_container_width=True, hide_index=True, height=260)
                st.download_button(
                    "⬇️ Download Today's Records as Excel", to_excel_bytes(disp),
                    file_name=f"loss_time_today_{st.session_state.selected_date.isoformat()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                with st.expander("🗑️ Delete an entry (mistake correction)"):
                    del_id = st.selectbox(
                        "Select entry ID to delete",
                        df_today["id"].tolist(),
                        format_func=lambda i: f"#{i} — "
                        f"{df_today.loc[df_today['id']==i,'line'].values[0]} / "
                        f"{df_today.loc[df_today['id']==i,'category'].values[0]} / "
                        f"{df_today.loc[df_today['id']==i,'minutes'].values[0]:.0f} min",
                    )
                    if st.button("Delete Selected Entry", type="secondary"):
                        delete_entry(del_id)
                        st.rerun()
            else:
                st.info("No resolved entries yet for this date.")

        # ---- All records ----
        with sub_records:
            closed_df = df_all[df_all["status"] == "closed"]
            if closed_df.empty:
                st.info("No records yet.")
            else:
                section_header("Filters", PALETTE["sky"], "🔎")
                date_options = sorted(closed_df["entry_date"].unique(), reverse=True)
                date_label_map = {fmt_ddmmyyyy(d): d for d in date_options}
                rf1, rf2, rf3 = st.columns(3)
                with rf1:
                    rdate = st.multiselect("Date", list(date_label_map.keys()), key="rec_date")
                with rf2:
                    rline = st.multiselect("Line", LINES, key="rec_line")
                with rf3:
                    rcat = st.multiselect("Category", CATEGORIES, key="rec_cat")

                rdf = closed_df.copy()
                if rdate:
                    rdf = rdf[rdf["entry_date"].isin([date_label_map[l] for l in rdate])]
                if rline:
                    rdf = rdf[rdf["line"].isin(rline)]
                if rcat:
                    rdf = rdf[rdf["category"].isin(rcat)]

                section_header("All Records", PALETTE["navy"], "🗂️")
                disp_all = rdf.sort_values("recorded_at", ascending=False)[
                    ["id", "entry_date", "line", "category", "reason", "minutes_lost",
                     "workstations_affected", "minutes", "recorded_by"]
                ].copy()
                disp_all["entry_date"] = disp_all["entry_date"].apply(fmt_ddmmyyyy)
                disp_all = disp_all.rename(columns={
                    "minutes_lost": "Minutes Lost", "workstations_affected": "Work-stations Affected",
                    "minutes": "Total Minutes Lost",
                })
                st.dataframe(disp_all, use_container_width=True, hide_index=True, height=340)

                dl1, dl2 = st.columns(2)
                with dl1:
                    st.download_button("⬇️ Download as Excel", to_excel_bytes(disp_all),
                                        file_name="loss_time_data.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with dl2:
                    st.download_button("⬇️ Download as CSV", disp_all.to_csv(index=False).encode("utf-8"),
                                        file_name="loss_time_data.csv", mime="text/csv")

# ----------------------------------------------------------------------------
# DASHBOARD (everyone) — live alert banner + 4 charts, gentle scrolling OK
# ----------------------------------------------------------------------------
with tab_dashboard:
    df = df_all
    open_df = df[df["status"] == "open"].copy()

    # ---- Live ongoing-issue alert banner ----
    if not open_df.empty:
        now = now_pkt()
        st.markdown(
            f"<div style='background:#fdecea; border:1px solid {PALETTE['rose']}; "
            f"border-radius:10px; padding:10px 16px; margin-bottom:10px;'>"
            f"<span style='color:{PALETTE['rose']}; font-weight:800; font-size:14px;'>"
            f"🔴 {len(open_df)} ONGOING ISSUE{'S' if len(open_df) > 1 else ''}</span>",
            unsafe_allow_html=True,
        )
        for _, row in open_df.iterrows():
            elapsed = round((now - parse_dt(row["start_time"])).total_seconds() / 60)
            ws = row["workstations_affected"] or 1
            total_so_far = round(elapsed * ws)
            st.markdown(
                f"<div style='padding:3px 0 3px 4px; font-size:14px;'>"
                f"🔴 <b>{row['line']}</b> — {row['category']} — running <b>{elapsed} min</b> "
                f"({total_so_far} minutes lost)</div>",
                unsafe_allow_html=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

    closed_df = df[df["status"] == "closed"]

    if closed_df.empty:
        st.info("No resolved data yet. Charts will appear here once loss-time entries are logged.")
    else:
        all_dates = sorted(pd.to_datetime(closed_df["entry_date"]).dt.date.unique())
        min_date, max_date = all_dates[0], max(all_dates[-1], today_pkt())

        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            date_range = st.date_input(
                "Filter by Date (from – to)",
                value=(today_pkt(), today_pkt()),
                min_value=min_date, max_value=max_date,
                format="DD-MM-YYYY", key="dash_date_range",
            )
        with fc2:
            line_filter = st.multiselect("Filter by Line", LINES)
        with fc3:
            category_filter = st.multiselect("Filter by Category", CATEGORIES)

        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_d, end_d = date_range
        elif isinstance(date_range, tuple) and len(date_range) == 1:
            start_d = end_d = date_range[0]
        else:
            start_d = end_d = date_range

        fdf = closed_df.copy()
        fdf_dates = pd.to_datetime(fdf["entry_date"]).dt.date
        fdf = fdf[(fdf_dates >= start_d) & (fdf_dates <= end_d)]
        if line_filter:
            fdf = fdf[fdf["line"].isin(line_filter)]
        if category_filter:
            fdf = fdf[fdf["category"].isin(category_filter)]

        if fdf.empty:
            st.warning("No records match the selected filters.")
        else:
            worst_cat = fdf.groupby("category")["minutes"].sum().idxmax()
            k1, k2 = st.columns(2)
            with k1:
                kpi_card("Total Lost Minutes", f"{fdf['minutes'].sum():,.0f}", PALETTE["rose"], "#e67e22")
            with k2:
                kpi_card("Top Loss Category", worst_cat, PALETTE["plum"], "#3f51b5")

            # ---- Row 1: Line-wise | Category donut ----
            r1c1, r1c2 = st.columns(2)
            with r1c1:
                section_header("Line-wise Lost Minutes", PALETTE["sky"], "📶")
                line_agg = fdf.groupby("line")["minutes"].sum().reindex(LINES).fillna(0).reset_index()
                fig = px.bar(line_agg, x="line", y="minutes", text_auto=".0f", color="line",
                             color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_traces(marker_line_width=0, textfont_size=12)
                fig.update_layout(showlegend=False, yaxis_title="Minutes", xaxis_title="",
                                   plot_bgcolor="white", paper_bgcolor="white",
                                   height=CHART_H, margin=CHART_MARGIN)
                st.plotly_chart(fig, use_container_width=True)

            with r1c2:
                section_header("Category-wise Lost Minutes", PALETTE["amber"], "🍩")
                total_minutes = fdf["minutes"].sum()
                cat_agg = fdf.groupby("category")["minutes"].sum().sort_values(ascending=False).reset_index()
                fig2 = go.Figure(data=[go.Pie(
                    labels=cat_agg["category"], values=cat_agg["minutes"], hole=0.5,
                    marker=dict(colors=CATEGORY_COLORS, line=dict(color="#ffffff", width=2)),
                    texttemplate="%{value:.0f} (%{percent})", textposition="outside", textfont_size=11,
                    hovertemplate="%{label}<br>%{value:.0f} min (%{percent})<extra></extra>",
                )])
                fig2.update_layout(
                    showlegend=True, legend=dict(orientation="h", y=-0.2, font=dict(size=10)),
                    paper_bgcolor="white", height=CHART_H, margin=CHART_MARGIN,
                    annotations=[dict(text=f"{total_minutes:.0f}<br>min", x=0.5, y=0.5,
                                       font_size=15, showarrow=False)],
                )
                st.plotly_chart(fig2, use_container_width=True)

            # ---- Row 2: Top 5 Loss Time Reasons | Pareto ----
            r2c1, r2c2 = st.columns(2)
            with r2c1:
                section_header("Top 5 Loss Time Reasons", PALETTE["plum"], "🏆")
                total_minutes_all = fdf["minutes"].sum()
                top5 = (
                    fdf.groupby("category")["minutes"].sum()
                    .sort_values(ascending=False).head(5).reset_index()
                )
                top5["pct"] = top5["minutes"] / total_minutes_all * 100
                top5["label"] = top5.apply(lambda r: f"{r['minutes']:.0f} ({r['pct']:.1f}%)", axis=1)
                top5_plot = top5.sort_values("minutes", ascending=False)  # largest first so it renders at the top
                fig_top5 = px.bar(
                    top5_plot, x="minutes", y="category", orientation="h",
                    text="label", color="category",
                    color_discrete_sequence=px.colors.qualitative.Set2,
                )
                fig_top5.update_traces(marker_line_width=0, textfont_size=12, width=0.55)
                fig_top5.update_layout(
                    showlegend=False, yaxis_title="", xaxis_title="Minutes",
                    plot_bgcolor="white", paper_bgcolor="white",
                    height=CHART_H, margin=CHART_MARGIN,
                )
                st.plotly_chart(fig_top5, use_container_width=True)


            with r2c2:
                section_header("Pareto Chart — Loss Time Categories", PALETTE["rose"], "📊")
                pareto = fdf.groupby("category")["minutes"].sum().sort_values(ascending=False).reset_index()
                pareto["cum_pct"] = pareto["minutes"].cumsum() / pareto["minutes"].sum() * 100
                fig3 = go.Figure()
                fig3.add_bar(x=pareto["category"], y=pareto["minutes"], name="Minutes Lost", width=0.5,
                              marker=dict(color=pareto["minutes"], colorscale="Tealgrn", line=dict(width=0)),
                              text=pareto["minutes"].round(0), textposition="outside", textfont_size=10)
                fig3.add_trace(go.Scatter(x=pareto["category"], y=pareto["cum_pct"], name="Cumulative %",
                                           yaxis="y2", mode="lines+markers",
                                           line=dict(color=PALETTE["rose"], width=2.5), marker=dict(size=6)))
                fig3.add_hline(y=80, line_dash="dot", line_color="gray", yref="y2")
                fig3.update_layout(
                    yaxis=dict(title="Minutes", tickfont_size=9),
                    yaxis2=dict(title="Cum %", overlaying="y", side="right", range=[0, 110], tickfont_size=9),
                    xaxis_tickangle=-35, xaxis_tickfont_size=9,
                    legend=dict(orientation="h", y=1.18, font=dict(size=10)),
                    plot_bgcolor="white", paper_bgcolor="white", height=CHART_H, margin=CHART_MARGIN,
                )
                st.plotly_chart(fig3, use_container_width=True)

