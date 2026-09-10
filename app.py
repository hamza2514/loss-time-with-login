import os
import io
import base64
import yaml
from yaml.loader import SafeLoader
import streamlit as st
import streamlit_authenticator as stauth
import sqlite3
import pandas as pd
from datetime import date, datetime
import plotly.express as px
import plotly.graph_objects as go
from streamlit_autorefresh import st_autorefresh

# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------
st.set_page_config(page_title="Loss Time Tracker | AM-4CT2", page_icon="🏭", layout="wide")

DB_PATH = "loss_time.db"
LOGO_PATH = "logo.jpg"
COMPANY_NAME = "Artistic Milliners"
UNIT_NAME = "AM-4CT2"

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


def brand_header(center=False):
    """Single self-contained flex block: logo + wordmark. Avoids column-layout clipping."""
    logo_uri = _logo_data_uri()
    logo_html = f"<img src='{logo_uri}' style='height:52px; width:auto; flex-shrink:0;'/>" if logo_uri else ""
    justify = "center" if center else "flex-start"
    st.markdown(
        f"""
        <div style="display:flex; align-items:center; gap:16px; justify-content:{justify};
                    padding:4px 0 10px 0;">
            {logo_html}
            <div style="line-height:1.35;">
                <div style="font-size:12px; letter-spacing:1.5px; color:{PALETTE['gray']};
                            font-weight:700; text-transform:uppercase;">{COMPANY_NAME}</div>
                <div style="font-size:23px; font-weight:800; color:{PALETTE['navy']};">
                    Line Loss Time Tracker</div>
                <div style="font-size:13px; color:#666; font-weight:600;">
                    Unit: {UNIT_NAME}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


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

authenticator = stauth.Authenticate(
    credentials, cookie_cfg["name"], cookie_cfg["key"], cookie_cfg.get("expiry_days", 7),
)

if not st.session_state.get("authentication_status"):
    l1, l2, l3 = st.columns([1, 1.3, 1])
    with l2:
        brand_header(center=True)
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
user_role = credentials["usernames"].get(username, {}).get("role", "viewer")
can_enter_data = user_role == "entry"

with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=60)
    st.markdown(f"**{display_name}**")
    st.caption(f"Role: {'Data Entry + Dashboard' if can_enter_data else 'Dashboard (view only)'}")
    authenticator.logout("Logout", "sidebar")

# ----------------------------------------------------------------------------
# DATABASE HELPERS
# ----------------------------------------------------------------------------
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS loss_time (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
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
    conn.commit()
    cols = [r[1] for r in conn.execute("PRAGMA table_info(loss_time)").fetchall()]
    for col, ddl in [
        ("recorded_by", "ALTER TABLE loss_time ADD COLUMN recorded_by TEXT"),
        ("status", "ALTER TABLE loss_time ADD COLUMN status TEXT DEFAULT 'closed'"),
        ("start_time", "ALTER TABLE loss_time ADD COLUMN start_time TEXT"),
        ("end_time", "ALTER TABLE loss_time ADD COLUMN end_time TEXT"),
    ]:
        if col not in cols:
            conn.execute(ddl)
            conn.commit()
    return conn


def add_closed_entry(entry_date, line, category, reason, minutes, recorded_by):
    conn = get_conn()
    now = datetime.now().isoformat(timespec="seconds")
    conn.execute(
        "INSERT INTO loss_time (entry_date, line, category, reason, minutes, recorded_at, "
        "recorded_by, status, start_time, end_time) VALUES (?,?,?,?,?,?,?, 'closed', ?, ?)",
        (entry_date, line, category, reason, minutes, now, recorded_by, now, now),
    )
    conn.commit()
    conn.close()


def start_open_issue(line, category, reason, recorded_by):
    conn = get_conn()
    now = datetime.now().isoformat(timespec="seconds")
    conn.execute(
        "INSERT INTO loss_time (entry_date, line, category, reason, minutes, recorded_at, "
        "recorded_by, status, start_time, end_time) VALUES (?,?,?,?,NULL,?,?, 'open', ?, NULL)",
        (date.today().isoformat(), line, category, reason, now, recorded_by, now),
    )
    conn.commit()
    conn.close()


def resolve_issue(entry_id):
    conn = get_conn()
    row = conn.execute("SELECT start_time FROM loss_time WHERE id=?", (entry_id,)).fetchone()
    if row:
        start = datetime.fromisoformat(row[0])
        end = datetime.now()
        minutes = round((end - start).total_seconds() / 60, 1)
        conn.execute(
            "UPDATE loss_time SET end_time=?, minutes=?, status='closed' WHERE id=?",
            (end.isoformat(timespec="seconds"), minutes, entry_id),
        )
        conn.commit()
    conn.close()


def load_data():
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM loss_time ORDER BY id DESC", conn)
    conn.close()
    return df


def delete_entry(entry_id):
    conn = get_conn()
    conn.execute("DELETE FROM loss_time WHERE id=?", (entry_id,))
    conn.commit()
    conn.close()


if "selected_date" not in st.session_state:
    st.session_state.selected_date = date.today()

st_autorefresh(interval=15_000, key="global_refresh")  # keeps ongoing-issue timers & dashboard live

# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
brand_header()
st.markdown("<hr style='margin-top:2px; margin-bottom:10px;'>", unsafe_allow_html=True)

if can_enter_data:
    tab_entry, tab_dashboard = st.tabs(["📝 Data Entry", "📊 Dashboard"])
else:
    tab_dashboard = st.container()
    tab_entry = None

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
            with st.form("entry_form", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    line = st.selectbox("Select Line", LINES)
                with c2:
                    category = st.selectbox("Select Reason Category", CATEGORIES)
                reason = st.text_area("Detailed Reason", height=80,
                                       placeholder="e.g. Waiting for trim from store...")
                minutes = st.number_input("Minutes Lost", min_value=0.0, step=1.0, format="%.0f")
                submitted = st.form_submit_button("Submit", use_container_width=True, type="primary")

                if submitted:
                    if minutes <= 0:
                        st.error("Please enter minutes lost greater than 0.")
                    elif not reason.strip():
                        st.error("Please enter a detailed reason.")
                    else:
                        add_closed_entry(
                            st.session_state.selected_date.isoformat(),
                            line, category, reason.strip(), minutes, display_name,
                        )
                        st.success(f"Recorded: {line} | {category} | {minutes:.0f} min")
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
                start_submitted = st.form_submit_button("🔴 Start Issue", use_container_width=True, type="primary")
                if start_submitted:
                    if not o_reason.strip():
                        st.error("Please enter a detailed reason.")
                    else:
                        start_open_issue(o_line, o_category, o_reason.strip(), display_name)
                        st.success(f"Started: {o_line} | {o_category} — now live on the Dashboard.")
                        st.rerun()

            section_header("Currently Ongoing", PALETTE["rose"], "⏳")
            df_all = load_data()
            open_df = df_all[df_all["status"] == "open"].copy()
            if open_df.empty:
                st.info("No ongoing issues right now.")
            else:
                now = datetime.now()
                for _, row in open_df.iterrows():
                    elapsed = round((now - datetime.fromisoformat(row["start_time"])).total_seconds() / 60)
                    ic1, ic2 = st.columns([5, 1])
                    with ic1:
                        st.markdown(
                            f"🔴 **{row['line']}** — {row['category']} — *{row['reason']}* "
                            f"— running **{elapsed} min** (started {datetime.fromisoformat(row['start_time']).strftime('%H:%M')})"
                        )
                    with ic2:
                        if st.button("Resolve", key=f"resolve_{row['id']}"):
                            resolve_issue(row["id"])
                            st.rerun()

        # ---- Today's log ----
        with sub_today:
            df_all = load_data()
            df_today = df_all[(df_all["entry_date"] == st.session_state.selected_date.isoformat())
                               & (df_all["status"] == "closed")]

            if not df_today.empty:
                kpi_card("Total Lost Minutes Today", f"{df_today['minutes'].sum():,.0f}",
                         PALETTE["rose"], "#e67e22")
                st.write("")
                disp = df_today[["id", "line", "category", "reason", "minutes", "entry_date"]].copy()
                disp["entry_date"] = disp["entry_date"].apply(fmt_ddmmyyyy)
                st.dataframe(disp, use_container_width=True, hide_index=True, height=260)
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
            df_all = load_data()
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
                    ["id", "entry_date", "line", "category", "reason", "minutes", "recorded_by"]
                ].copy()
                disp_all["entry_date"] = disp_all["entry_date"].apply(fmt_ddmmyyyy)
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
    df = load_data()
    open_df = df[df["status"] == "open"].copy()

    # ---- Live ongoing-issue alert banner ----
    if not open_df.empty:
        now = datetime.now()
        st.markdown(
            f"<div style='background:#fdecea; border:1px solid {PALETTE['rose']}; "
            f"border-radius:10px; padding:10px 16px; margin-bottom:10px;'>"
            f"<span style='color:{PALETTE['rose']}; font-weight:800; font-size:14px;'>"
            f"🔴 {len(open_df)} ONGOING ISSUE{'S' if len(open_df) > 1 else ''}</span>",
            unsafe_allow_html=True,
        )
        for _, row in open_df.iterrows():
            elapsed = round((now - datetime.fromisoformat(row["start_time"])).total_seconds() / 60)
            st.markdown(
                f"<div style='padding:3px 0 3px 4px; font-size:14px;'>"
                f"🔴 <b>{row['line']}</b> — {row['category']} — running <b>{elapsed} min</b></div>",
                unsafe_allow_html=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

    closed_df = df[df["status"] == "closed"]

    if closed_df.empty:
        st.info("No resolved data yet. Charts will appear here once loss-time entries are logged.")
    else:
        date_options = sorted(closed_df["entry_date"].unique(), reverse=True)
        date_label_map = {fmt_ddmmyyyy(d): d for d in date_options}

        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            date_filter_labels = st.multiselect("Filter by Date", list(date_label_map.keys()))
        with fc2:
            line_filter = st.multiselect("Filter by Line", LINES)
        with fc3:
            category_filter = st.multiselect("Filter by Category", CATEGORIES)

        fdf = closed_df.copy()
        if date_filter_labels:
            selected_iso = [date_label_map[lbl] for lbl in date_filter_labels]
            fdf = fdf[fdf["entry_date"].isin(selected_iso)]
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

            # ---- Row 2: Pareto (full width) ----
            section_header("Pareto Chart — Loss Time Categories", PALETTE["rose"], "📊")
            pareto = fdf.groupby("category")["minutes"].sum().sort_values(ascending=False).reset_index()
            pareto["cum_pct"] = pareto["minutes"].cumsum() / pareto["minutes"].sum() * 100
            fig3 = go.Figure()
            fig3.add_bar(x=pareto["category"], y=pareto["minutes"], name="Minutes Lost",
                          marker=dict(color=pareto["minutes"], colorscale="Tealgrn", line=dict(width=0)),
                          text=pareto["minutes"].round(0), textposition="outside", textfont_size=11)
            fig3.add_trace(go.Scatter(x=pareto["category"], y=pareto["cum_pct"], name="Cumulative %",
                                       yaxis="y2", mode="lines+markers",
                                       line=dict(color=PALETTE["rose"], width=2.5), marker=dict(size=7)))
            fig3.add_hline(y=80, line_dash="dot", line_color="gray", yref="y2")
            fig3.update_layout(
                yaxis=dict(title="Minutes"), yaxis2=dict(title="Cumulative %", overlaying="y", side="right", range=[0, 110]),
                xaxis_tickangle=-25, legend=dict(orientation="h", y=1.15, font=dict(size=11)),
                plot_bgcolor="white", paper_bgcolor="white", height=CHART_H, margin=CHART_MARGIN,
            )
            st.plotly_chart(fig3, use_container_width=True)

            # ---- Row 3: Heatmap (full width, all 6 lines forced visible) ----
            section_header("Line × Category Heatmap", PALETTE["teal"], "🔥")
            heat = fdf.pivot_table(index="line", columns="category", values="minutes",
                                    aggfunc="sum", fill_value=0).reindex(LINES)
            fig4 = px.imshow(heat, text_auto=".0f", aspect="auto", color_continuous_scale="Sunsetdark")
            fig4.update_yaxes(tickmode="array", tickvals=list(range(len(LINES))), ticktext=LINES)
            fig4.update_layout(xaxis_tickangle=-20, paper_bgcolor="white",
                                height=CHART_H + 20, margin=CHART_MARGIN, font_size=11)
            fig4.update_traces(xgap=3, ygap=3)
            st.plotly_chart(fig4, use_container_width=True)
