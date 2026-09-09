import os
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
}
CATEGORY_COLORS = px.colors.qualitative.Bold
CHART_H = 260          # compact chart height so everything fits one screen
CHART_MARGIN = dict(t=35, b=30, l=10, r=10)

# ----------------------------------------------------------------------------
# COMPACT LIGHT THEME / CSS
# ----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .stApp { background-color: #ffffff; }
    .block-container {
        padding-top: 1.1rem;
        padding-bottom: 1rem;
        max-width: 1400px;
    }
    section[data-testid="stSidebar"] { background-color: #f7f9fc; }
    div[data-testid="stForm"] {
        background-color: #fbfcfe;
        border: 1px solid #e3e8f0;
        border-radius: 12px;
        padding: 14px 18px 4px 18px;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #f0f2f6;
        border-radius: 8px 8px 0 0;
        padding: 6px 14px;
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
        padding: 0.35rem 1.1rem;
    }
    h1, h2, h3 { margin-top: 0.2rem; margin-bottom: 0.2rem; }
    [data-testid="stMetricValue"] { font-size: 22px; }
    </style>
    """,
    unsafe_allow_html=True,
)


def section_header(text, color=PALETTE["navy"], emoji=""):
    st.markdown(
        f"<div style='color:{color}; border-bottom:2px solid {color}; "
        f"padding-bottom:3px; margin-top:4px; margin-bottom:6px; "
        f"font-weight:700; font-size:16px;'>{emoji} {text}</div>",
        unsafe_allow_html=True,
    )


def kpi_card(label, value, color1, color2):
    st.markdown(
        f"""
        <div style="background: linear-gradient(135deg,{color1},{color2});
                    padding:12px 16px; border-radius:12px; text-align:center;
                    color:#ffffff; box-shadow: 0 4px 10px rgba(0,0,0,0.12);">
            <div style="font-size:12px; font-weight:700; letter-spacing:0.5px;
                        text-transform:uppercase; opacity:0.92;">{label}</div>
            <div style="font-size:24px; font-weight:800; margin-top:4px;">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def fmt_ddmmyyyy(iso_str):
    return datetime.strptime(iso_str, "%Y-%m-%d").strftime("%d-%m-%Y")


# ----------------------------------------------------------------------------
# AUTHENTICATION
# ----------------------------------------------------------------------------
def load_auth_config():
    """Prefer Streamlit Cloud secrets; fall back to local YAML for laptop testing."""
    try:
        has_secrets = "credentials" in st.secrets
    except Exception:
        has_secrets = False
    if has_secrets:
        credentials = {
            "usernames": {
                u: dict(v) for u, v in st.secrets["credentials"]["usernames"].items()
            }
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
    credentials,
    cookie_cfg["name"],
    cookie_cfg["key"],
    cookie_cfg.get("expiry_days", 7),
)

# ---- Login screen ----
if not st.session_state.get("authentication_status"):
    l1, l2, l3 = st.columns([1, 1.1, 1])
    with l2:
        if os.path.exists(LOGO_PATH):
            lc1, lc2, lc3 = st.columns([1, 1, 1])
            with lc2:
                st.image(LOGO_PATH, width=90)
        st.markdown(
            f"<h3 style='text-align:center; color:{PALETTE['navy']};'>Line Loss Time Tracker</h3>"
            f"<p style='text-align:center; color:#666; margin-top:-8px;'>Unit: {UNIT_NAME}</p>",
            unsafe_allow_html=True,
        )
        authenticator.login(location="main")

auth_status = st.session_state.get("authentication_status")

if auth_status is False:
    st.error("Username or password is incorrect.")
    st.stop()
elif auth_status is None:
    st.stop()

# ---- Logged in ----
username = st.session_state.get("username")
display_name = st.session_state.get("name")
user_role = credentials["usernames"].get(username, {}).get("role", "viewer")
can_enter_data = user_role == "entry"

with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=70)
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
            minutes REAL NOT NULL,
            recorded_at TEXT NOT NULL,
            recorded_by TEXT
        )
        """
    )
    conn.commit()
    cols = [r[1] for r in conn.execute("PRAGMA table_info(loss_time)").fetchall()]
    if "recorded_by" not in cols:
        conn.execute("ALTER TABLE loss_time ADD COLUMN recorded_by TEXT")
        conn.commit()
    return conn


def add_entry(entry_date, line, category, reason, minutes, recorded_by):
    conn = get_conn()
    conn.execute(
        "INSERT INTO loss_time (entry_date, line, category, reason, minutes, recorded_at, recorded_by) "
        "VALUES (?,?,?,?,?,?,?)",
        (entry_date, line, category, reason, minutes,
         datetime.now().isoformat(timespec="seconds"), recorded_by),
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


SLOT_EDGES_MIN = [510, 570, 630, 690, 750, 810, 870, 930, 990, 1050]
SLOT_LABELS = [
    "08:30-09:30", "09:30-10:30", "10:30-11:30", "11:30-12:30", "12:30-13:30",
    "13:30-14:30", "14:30-15:30", "15:30-16:30", "16:30-17:30",
]


def assign_time_slot(recorded_at_str):
    t = datetime.fromisoformat(recorded_at_str)
    mins = t.hour * 60 + t.minute
    for i in range(len(SLOT_EDGES_MIN) - 1):
        if SLOT_EDGES_MIN[i] <= mins < SLOT_EDGES_MIN[i + 1]:
            return SLOT_LABELS[i]
    return "Outside Shift"


if "selected_date" not in st.session_state:
    st.session_state.selected_date = date.today()

# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
h_logo, h_title = st.columns([1, 8])
with h_logo:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=55)
with h_title:
    st.markdown(
        f"<span style='color:{PALETTE['navy']}; font-size:22px; font-weight:800;'>🏭 Line Loss Time Tracker</span>"
        f"&nbsp;&nbsp;<span style='background:{PALETTE['teal']}; color:white; font-weight:700; "
        f"padding:2px 10px; border-radius:14px; font-size:12px;'>Unit: {UNIT_NAME}</span>",
        unsafe_allow_html=True,
    )
st.markdown("<hr style='margin-top:6px; margin-bottom:6px;'>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# MAIN NAVIGATION — tabs shown depend on role
# ----------------------------------------------------------------------------
if can_enter_data:
    tab_entry, tab_dashboard = st.tabs(["📝 Data Entry", "📊 Dashboard"])
else:
    tab_dashboard = st.container()
    tab_entry = None

# ----------------------------------------------------------------------------
# DATA ENTRY (entry-role only) — organised in its own sub-tabs
# ----------------------------------------------------------------------------
if can_enter_data:
    with tab_entry:
        sub_new, sub_today = st.tabs(["➕ New Entry", "📋 Today's Log"])

        with sub_new:
            cdate, cform = st.columns([1, 2])
            with cdate:
                section_header("Date", PALETTE["sky"], "📅")
                st.session_state.selected_date = st.date_input(
                    "Production date", value=st.session_state.selected_date,
                    format="DD-MM-YYYY", label_visibility="collapsed",
                )

            section_header("Log a Loss Time Event", PALETTE["amber"], "⏱️")
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
                        add_entry(
                            st.session_state.selected_date.isoformat(),
                            line, category, reason.strip(), minutes, display_name,
                        )
                        st.success(f"Recorded: {line} | {category} | {minutes:.0f} min")
                        st.rerun()

        with sub_today:
            df_all = load_data()
            df_today = df_all[df_all["entry_date"] == st.session_state.selected_date.isoformat()]

            if not df_today.empty:
                kpi_card("Total Lost Minutes Today", f"{df_today['minutes'].sum():,.0f}",
                         PALETTE["rose"], "#e67e22")
                st.write("")
                disp = df_today[["id", "line", "category", "reason", "minutes", "entry_date"]].copy()
                disp["entry_date"] = disp["entry_date"].apply(fmt_ddmmyyyy)
                st.dataframe(disp, use_container_width=True, hide_index=True, height=230)
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
                st.info("No entries yet for this date.")

# ----------------------------------------------------------------------------
# DASHBOARD (everyone) — split into sub-tabs so each screen fits a laptop
# ----------------------------------------------------------------------------
with tab_dashboard:
    st_autorefresh(interval=15_000, key="dashboard_refresh")
    df = load_data()

    if df.empty:
        st.info("No data recorded yet. Entries will appear here live as they're submitted.")
    else:
        date_options = sorted(df["entry_date"].unique(), reverse=True)
        date_label_map = {fmt_ddmmyyyy(d): d for d in date_options}

        fc1, fc2 = st.columns(2)
        with fc1:
            date_filter_labels = st.multiselect("Filter by Date", list(date_label_map.keys()))
        with fc2:
            line_filter = st.multiselect("Filter by Line", LINES)

        fdf = df.copy()
        if date_filter_labels:
            selected_iso = [date_label_map[lbl] for lbl in date_filter_labels]
            fdf = fdf[fdf["entry_date"].isin(selected_iso)]
        if line_filter:
            fdf = fdf[fdf["line"].isin(line_filter)]

        if fdf.empty:
            st.warning("No records match the selected filters.")
        else:
            sub_summary, sub_details = st.tabs(["📈 Summary", "🔍 Details & Records"])

            # ---------------- SUMMARY ----------------
            with sub_summary:
                worst_cat = fdf.groupby("category")["minutes"].sum().idxmax()
                k1, k2 = st.columns(2)
                with k1:
                    kpi_card("Total Lost Minutes", f"{fdf['minutes'].sum():,.0f}", PALETTE["rose"], "#e67e22")
                with k2:
                    kpi_card("Top Loss Category", worst_cat, PALETTE["plum"], "#3f51b5")

                st.write("")
                c1, c2 = st.columns(2)
                with c1:
                    section_header("Line-wise Lost Minutes", PALETTE["sky"], "📶")
                    line_agg = fdf.groupby("line")["minutes"].sum().reindex(LINES).fillna(0).reset_index()
                    fig = px.bar(line_agg, x="line", y="minutes", text_auto=".0f", color="line",
                                 color_discrete_sequence=px.colors.qualitative.Set2)
                    fig.update_traces(marker_line_width=0, textfont_size=11)
                    fig.update_layout(showlegend=False, yaxis_title="Minutes", xaxis_title="",
                                       plot_bgcolor="white", paper_bgcolor="white",
                                       height=CHART_H, margin=CHART_MARGIN)
                    st.plotly_chart(fig, use_container_width=True)

                with c2:
                    section_header("Category-wise Lost Minutes", PALETTE["amber"], "🍩")
                    total_minutes = fdf["minutes"].sum()
                    cat_agg = fdf.groupby("category")["minutes"].sum().sort_values(ascending=False).reset_index()
                    fig2 = go.Figure(data=[go.Pie(
                        labels=cat_agg["category"], values=cat_agg["minutes"], hole=0.5,
                        marker=dict(colors=CATEGORY_COLORS, line=dict(color="#ffffff", width=2)),
                        texttemplate="%{value:.0f} (%{percent})", textposition="outside",
                        textfont_size=10,
                        hovertemplate="%{label}<br>%{value:.0f} min (%{percent})<extra></extra>",
                    )])
                    fig2.update_layout(
                        showlegend=True, legend=dict(orientation="h", y=-0.25, font=dict(size=9)),
                        paper_bgcolor="white", height=CHART_H, margin=CHART_MARGIN,
                        annotations=[dict(text=f"{total_minutes:.0f}<br>min", x=0.5, y=0.5,
                                           font_size=13, showarrow=False)],
                    )
                    st.plotly_chart(fig2, use_container_width=True)

                section_header("Pareto Chart — Loss Time Categories", PALETTE["rose"], "📊")
                pareto = fdf.groupby("category")["minutes"].sum().sort_values(ascending=False).reset_index()
                pareto["cum_pct"] = pareto["minutes"].cumsum() / pareto["minutes"].sum() * 100
                fig3 = go.Figure()
                fig3.add_bar(x=pareto["category"], y=pareto["minutes"], name="Minutes Lost",
                              marker=dict(color=pareto["minutes"], colorscale="Tealgrn", line=dict(width=0)),
                              text=pareto["minutes"].round(0), textposition="outside", textfont_size=10)
                fig3.add_trace(go.Scatter(x=pareto["category"], y=pareto["cum_pct"], name="Cumulative %",
                                           yaxis="y2", mode="lines+markers",
                                           line=dict(color=PALETTE["rose"], width=2), marker=dict(size=6)))
                fig3.add_hline(y=80, line_dash="dot", line_color="gray", yref="y2")
                fig3.update_layout(
                    yaxis=dict(title="Minutes"), yaxis2=dict(title="Cum %", overlaying="y", side="right", range=[0, 110]),
                    xaxis_tickangle=-25, legend=dict(orientation="h", y=1.2, font=dict(size=10)),
                    plot_bgcolor="white", paper_bgcolor="white", height=CHART_H + 30,
                    margin=dict(t=45, b=60, l=10, r=10),
                )
                st.plotly_chart(fig3, use_container_width=True)

            # ---------------- DETAILS ----------------
            with sub_details:
                c3, c4 = st.columns(2)
                with c3:
                    section_header("Line × Category Heatmap", PALETTE["teal"], "🔥")
                    heat = fdf.pivot_table(index="line", columns="category", values="minutes",
                                            aggfunc="sum", fill_value=0).reindex(LINES)
                    fig4 = px.imshow(heat, text_auto=".0f", aspect="auto", color_continuous_scale="Sunsetdark")
                    fig4.update_layout(xaxis_tickangle=-25, paper_bgcolor="white",
                                        height=CHART_H, margin=CHART_MARGIN, font_size=9)
                    fig4.update_traces(xgap=2, ygap=2)
                    st.plotly_chart(fig4, use_container_width=True)

                with c4:
                    section_header("Loss Minutes by Hour of Day", PALETTE["green"], "🕒")
                    hdf = fdf.copy()
                    hdf["time_slot"] = hdf["recorded_at"].apply(assign_time_slot)
                    slot_order = SLOT_LABELS + ["Outside Shift"]
                    hourly = hdf.groupby("time_slot")["minutes"].sum().reindex(slot_order).fillna(0).reset_index()
                    fig5 = px.bar(hourly, x="time_slot", y="minutes", text_auto=".0f",
                                   color="minutes", color_continuous_scale="Bluyl")
                    fig5.update_traces(marker_line_width=0, textfont_size=10)
                    fig5.update_layout(yaxis_title="Minutes", xaxis_title="", xaxis_tickangle=-25,
                                        coloraxis_showscale=False, plot_bgcolor="white", paper_bgcolor="white",
                                        height=CHART_H, margin=CHART_MARGIN, font_size=9)
                    st.plotly_chart(fig5, use_container_width=True)

                section_header("All Records", PALETTE["navy"], "🗂️")
                disp_all = fdf.sort_values("recorded_at", ascending=False)[
                    ["id", "entry_date", "line", "category", "reason", "minutes", "recorded_by"]
                ].copy()
                disp_all["entry_date"] = disp_all["entry_date"].apply(fmt_ddmmyyyy)
                st.dataframe(disp_all, use_container_width=True, hide_index=True, height=230)
                st.download_button("⬇️ Download filtered data as CSV",
                                    fdf.to_csv(index=False).encode("utf-8"),
                                    file_name="loss_time_data.csv", mime="text/csv")
