import os
import io
import re
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
DEFAULT_UNIT = "AM-4CT2"  # legacy rows saved before units existed belong to this unit

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


def _legacy_accounts():
    """Old-style accounts from Secrets [credentials] (or local auth_config.yaml).
    Only read once, to seed the database the first time the new version runs."""
    try:
        if "credentials" in st.secrets:
            return {u: dict(v) for u, v in st.secrets["credentials"]["usernames"].items()}
    except Exception:
        pass
    if os.path.exists("auth_config.yaml"):
        with open("auth_config.yaml") as f:
            cfg = yaml.load(f, Loader=SafeLoader) or {}
        return (cfg.get("credentials") or {}).get("usernames") or {}
    return {}


def _admin_from_secrets():
    try:
        if "admin" in st.secrets:
            a = dict(st.secrets["admin"])
            if a.get("username") and a.get("password"):
                return a
    except Exception:
        pass
    return None


def _seed_defaults(conn):
    """One-time setup: default units/lines, existing accounts, and the admin account."""
    def hash_pw(p):
        p = str(p)
        return p if stauth.Hasher.is_hash(p) else stauth.Hasher.hash(p)

    with conn.cursor() as cur:
        cur.execute("SELECT COUNT(*) FROM units")
        if cur.fetchone()[0] == 0:
            for u in ("AM-4CT2", "AM-4A"):
                cur.execute("INSERT INTO units (name) VALUES (%s) ON CONFLICT DO NOTHING", (u,))
                for i in range(1, 7):
                    cur.execute("INSERT INTO lines (unit, name) VALUES (%s,%s) ON CONFLICT DO NOTHING",
                                (u, f"Line # {i}"))

        cur.execute("SELECT COUNT(*) FROM app_users")
        if cur.fetchone()[0] == 0:
            for uname, v in _legacy_accounts().items():
                role = v.get("role", "viewer")
                role = role if role in ("entry", "viewer") else "viewer"
                units = list(v.get("units") or ["AM-4CT2"])
                cur.execute(
                    "INSERT INTO app_users (username, name, password_hash, role, units) "
                    "VALUES (%s,%s,%s,%s,%s) ON CONFLICT (username) DO NOTHING",
                    (str(uname).lower(), v.get("name", uname), hash_pw(v.get("password", "")), role, units),
                )

    conn.commit()
    _ensure_admin(conn)


def _ensure_admin(conn):
    """Create the admin from Secrets [admin] if no admin exists yet. If that username
    already exists (e.g. your old 'hamza' account) it is upgraded to admin.
    Returns True if it created/upgraded one."""
    a = _admin_from_secrets()
    if not a:
        return False
    with conn.cursor() as cur:
        cur.execute("SELECT 1 FROM app_users WHERE role='admin' LIMIT 1")
        if cur.fetchone():
            return False
        pw = str(a["password"])
        cur.execute(
            "INSERT INTO app_users (username, name, password_hash, role, units, active) "
            "VALUES (%s,%s,%s,'admin','{}',TRUE) "
            "ON CONFLICT (username) DO UPDATE SET role='admin', name=EXCLUDED.name, "
            "password_hash=EXCLUDED.password_hash, units='{}', active=TRUE",
            (str(a["username"]).lower(), a.get("name", a["username"]),
             pw if stauth.Hasher.is_hash(pw) else stauth.Hasher.hash(pw)),
        )
    conn.commit()
    return True


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
            cur.execute("CREATE TABLE IF NOT EXISTS units (id SERIAL PRIMARY KEY, name TEXT UNIQUE NOT NULL, active BOOLEAN NOT NULL DEFAULT TRUE)")
            cur.execute("CREATE TABLE IF NOT EXISTS lines (id SERIAL PRIMARY KEY, unit TEXT NOT NULL, name TEXT NOT NULL, active BOOLEAN NOT NULL DEFAULT TRUE, UNIQUE (unit, name))")
            cur.execute(
                "CREATE TABLE IF NOT EXISTS app_users (username TEXT PRIMARY KEY, name TEXT NOT NULL, "
                "password_hash TEXT NOT NULL, role TEXT NOT NULL DEFAULT 'viewer', units TEXT[] NOT NULL DEFAULT '{}', "
                "active BOOLEAN NOT NULL DEFAULT TRUE)"
            )
        conn.commit()
        _seed_defaults(conn)
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



# ----------------------------------------------------------------------------
# USERS / UNITS / LINES  (stored in the database — managed from the Admin tab)
# ----------------------------------------------------------------------------
USERNAME_RE = re.compile(r"^[a-z0-9_.-]{3,30}$")
MIN_PASSWORD_LEN = 6


def db_query(sql, params=None):
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchall()
    finally:
        release_conn(conn)


def db_run(sql, params=None):
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, params)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        release_conn(conn)


@st.cache_data(ttl=300, show_spinner=False)
def load_auth_users():
    """Active accounts in the shape streamlit-authenticator expects. Passwords are
    already bcrypt hashes in the DB, so nothing is hashed on rerun (fast)."""
    rows = db_query("SELECT username, name, password_hash, role, units FROM app_users WHERE active")
    return {"usernames": {
        r[0]: {"name": r[1], "password": r[2], "role": r[3], "units": list(r[4] or [])} for r in rows
    }}


@st.cache_data(ttl=300, show_spinner=False)
def load_all_users():
    rows = db_query("SELECT username, name, role, units, active FROM app_users ORDER BY (role <> 'admin'), role, username")
    return [{"username": r[0], "name": r[1], "role": r[2], "units": list(r[3] or []), "active": r[4]} for r in rows]


@st.cache_data(ttl=300, show_spinner=False)
def load_units():
    rows = db_query("SELECT name, active FROM units ORDER BY id")
    return [{"name": r[0], "active": r[1]} for r in rows]


@st.cache_data(ttl=300, show_spinner=False)
def load_lines():
    rows = db_query("SELECT unit, name, active FROM lines ORDER BY id")
    return [{"unit": r[0], "name": r[1], "active": r[2]} for r in rows]


def clear_admin_caches():
    for fn in (load_auth_users, load_all_users, load_units, load_lines):
        fn.clear()


def ensure_admin_account():
    """Runs only while no admin exists: picks up a newly added Secrets [admin]
    without needing the app to restart."""
    conn = get_conn()
    try:
        created = _ensure_admin(conn)
    finally:
        release_conn(conn)
    if created:
        clear_admin_caches()
    return created


def lines_for(unit, only_active):
    return [l["name"] for l in load_lines() if l["unit"] == unit and (l["active"] or not only_active)]


def add_user(username, name, password, role, units):
    """Returns an error string, or None on success."""
    username = (username or "").strip().lower()
    name = (name or "").strip()
    if not USERNAME_RE.match(username):
        return "Username must be 3–30 characters: lowercase letters, digits, dot, dash or underscore."
    if not name:
        return "Please enter the person's name."
    if len(password or "") < MIN_PASSWORD_LEN:
        return f"Password must be at least {MIN_PASSWORD_LEN} characters."
    if role not in ("entry", "viewer"):
        return "Invalid role."
    if not units:
        return "Please assign at least one unit."
    if any(u["username"] == username for u in load_all_users()):
        return "That username already exists."
    try:
        db_run(
            "INSERT INTO app_users (username, name, password_hash, role, units) VALUES (%s,%s,%s,%s,%s)",
            (username, name, stauth.Hasher.hash(password), role, list(units)),
        )
    except psycopg2.errors.UniqueViolation:
        return "That username already exists."
    clear_admin_caches()
    return None


def update_user(username, name, role, units, active, new_password=None):
    if not (name or "").strip():
        return "Name cannot be empty."
    if new_password and len(new_password) < MIN_PASSWORD_LEN:
        return f"New password must be at least {MIN_PASSWORD_LEN} characters."
    if role != "admin" and not units:
        return "Please assign at least one unit."
    if role == "admin":
        # admin: role/units/active are fixed (admin sees every unit, can't be locked out)
        db_run("UPDATE app_users SET name=%s WHERE username=%s", (name.strip(), username))
    else:
        db_run(
            "UPDATE app_users SET name=%s, role=%s, units=%s, active=%s WHERE username=%s",
            (name.strip(), role, list(units), bool(active), username),
        )
    if new_password:
        db_run("UPDATE app_users SET password_hash=%s WHERE username=%s",
               (stauth.Hasher.hash(new_password), username))
    clear_admin_caches()
    return None


def add_unit(name, n_lines):
    name = (name or "").strip()
    if not (2 <= len(name) <= 20):
        return "Unit name must be 2–20 characters."
    if any(u["name"].lower() == name.lower() for u in load_units()):
        return "That unit already exists."
    try:
        db_run("INSERT INTO units (name) VALUES (%s)", (name,))
    except psycopg2.errors.UniqueViolation:
        return "That unit already exists."
    for i in range(1, int(n_lines) + 1):
        db_run("INSERT INTO lines (unit, name) VALUES (%s,%s) ON CONFLICT DO NOTHING", (name, f"Line # {i}"))
    clear_admin_caches()
    return None


def set_unit_active(name, active):
    db_run("UPDATE units SET active=%s WHERE name=%s", (bool(active), name))
    clear_admin_caches()


def add_line(unit, name):
    name = (name or "").strip()
    if not (1 <= len(name) <= 30):
        return "Line name must be 1–30 characters."
    if any(l["unit"] == unit and l["name"].lower() == name.lower() for l in load_lines()):
        return "That line already exists in this unit."
    try:
        db_run("INSERT INTO lines (unit, name) VALUES (%s,%s)", (unit, name))
    except psycopg2.errors.UniqueViolation:
        return "That line already exists in this unit."
    clear_admin_caches()
    return None


def set_line_active(unit, name, active):
    db_run("UPDATE lines SET active=%s WHERE unit=%s AND name=%s", (bool(active), unit, name))
    clear_admin_caches()


def _flash(kind, msg):
    st.session_state["_flash"] = (kind, msg)


def render_admin():
    """Admin-only page: users, units, lines."""
    flash = st.session_state.pop("_flash", None)
    if flash:
        (st.success if flash[0] == "ok" else st.error)(flash[1])

    units_all = load_units()
    unit_names = [u["name"] for u in units_all]
    active_unit_names = [u["name"] for u in units_all if u["active"]]
    t_users, t_units, t_lines = st.tabs(["👥 Users", "🏭 Units", "📏 Lines"])

    # ---------------- USERS ----------------
    with t_users:
        section_header("Add User", PALETTE["teal"], "➕")
        ctr = st.session_state.setdefault("_au_ctr", 0)
        with st.form(f"add_user_form_{ctr}"):
            a1, a2 = st.columns(2)
            with a1:
                n_username = st.text_input("Username (login)", placeholder="e.g. ayesha")
                n_password = st.text_input("Password", type="password")
            with a2:
                n_name = st.text_input("Full name (shown in the app)")
                n_role = st.selectbox(
                    "Access", ["entry", "viewer"],
                    format_func=lambda r: "Data entry + Dashboard" if r == "entry" else "Dashboard only",
                )
            n_units = st.multiselect("Units this user can access", active_unit_names)
            if st.form_submit_button("Add User", type="primary"):
                err = add_user(n_username, n_name, n_password, n_role, n_units)
                if err:
                    st.error(err)
                else:
                    st.session_state["_au_ctr"] = ctr + 1
                    _flash("ok", f"User '{n_username.strip().lower()}' added. Share the password with them directly — it can't be viewed later.")
                    st.rerun()

        section_header("Existing Users", PALETTE["navy"], "👥")
        for u in load_all_users():
            is_admin_row = u["role"] == "admin"
            status = "" if u["active"] else "  ·  ⛔ inactive"
            units_txt = "all units" if is_admin_row else ", ".join(u["units"]) or "no units"
            with st.expander(f"{u['username']}  —  {u['name']}  ·  {u['role']}  ·  {units_txt}{status}"):
                with st.form(f"uf_{u['username']}"):
                    e1, e2 = st.columns(2)
                    with e1:
                        e_name = st.text_input("Full name", value=u["name"])
                        e_pw = st.text_input("New password (leave blank to keep current)", type="password")
                    with e2:
                        if is_admin_row:
                            st.caption("Admin account: always has access to every unit and can't be deactivated.")
                            e_role, e_units, e_active = "admin", [], True
                        else:
                            e_role = st.selectbox(
                                "Access", ["entry", "viewer"], index=0 if u["role"] == "entry" else 1,
                                format_func=lambda r: "Data entry + Dashboard" if r == "entry" else "Dashboard only",
                            )
                            e_units = st.multiselect(
                                "Units", unit_names, default=[x for x in u["units"] if x in unit_names],
                            )
                            e_active = st.checkbox("Account active", value=u["active"])
                    if st.form_submit_button("Save changes"):
                        err = update_user(u["username"], e_name, e_role, e_units, e_active, e_pw or None)
                        if err:
                            st.error(err)
                        else:
                            _flash("ok", f"Saved changes for '{u['username']}'.")
                            st.rerun()

    # ---------------- UNITS ----------------
    with t_units:
        section_header("Add Unit", PALETTE["teal"], "➕")
        uctr = st.session_state.setdefault("_unit_ctr", 0)
        with st.form(f"add_unit_form_{uctr}"):
            b1, b2 = st.columns(2)
            with b1:
                nu_name = st.text_input("Unit name", placeholder="e.g. AM-15B")
            with b2:
                nu_lines = st.number_input("Create this many lines now (Line # 1 … N)", min_value=0, max_value=60, value=6, step=1)
            if st.form_submit_button("Add Unit", type="primary"):
                err = add_unit(nu_name, nu_lines)
                if err:
                    st.error(err)
                else:
                    st.session_state["_unit_ctr"] = uctr + 1
                    _flash("ok", f"Unit '{nu_name.strip()}' added. Assign users to it from the Users tab.")
                    st.rerun()

        section_header("Existing Units", PALETTE["navy"], "🏭")
        st.caption("Units can be deactivated (hidden from everyone but kept in history) — not deleted, so past records stay intact.")
        for u in units_all:
            n_act = len(lines_for(u["name"], True))
            r1, r2, r3 = st.columns([3, 2, 2])
            r1.markdown(f"**{u['name']}**  ·  {n_act} active line{'s' if n_act != 1 else ''}")
            r2.markdown("🟢 Active" if u["active"] else "⛔ Inactive")
            if r3.button("Deactivate" if u["active"] else "Reactivate", key=f"ut_{u['name']}"):
                set_unit_active(u["name"], not u["active"])
                _flash("ok", f"Unit '{u['name']}' {'deactivated' if u['active'] else 'reactivated'}.")
                st.rerun()

    # ---------------- LINES ----------------
    with t_lines:
        if not unit_names:
            st.info("Add a unit first.")
        else:
            sel_unit = st.selectbox("Unit", unit_names, key="admin_line_unit")
            unit_lines = [l for l in load_lines() if l["unit"] == sel_unit]

            section_header(f"Add Line to {sel_unit}", PALETTE["teal"], "➕")
            with st.form(f"add_line_form_{sel_unit}_{len(unit_lines)}"):
                nl_name = st.text_input("Line name", value=f"Line # {len(unit_lines) + 1}")
                if st.form_submit_button("Add Line", type="primary"):
                    err = add_line(sel_unit, nl_name)
                    if err:
                        st.error(err)
                    else:
                        _flash("ok", f"'{nl_name.strip()}' added to {sel_unit}.")
                        st.rerun()

            section_header(f"Lines in {sel_unit}", PALETTE["navy"], "📏")
            st.caption("Deactivated lines disappear from the entry forms but stay in dashboards and records.")
            if not unit_lines:
                st.info("No lines yet for this unit.")
            for l in unit_lines:
                r1, r2, r3 = st.columns([3, 2, 2])
                r1.markdown(f"**{l['name']}**")
                r2.markdown("🟢 Active" if l["active"] else "⛔ Inactive")
                if r3.button("Deactivate" if l["active"] else "Reactivate", key=f"lt_{sel_unit}_{l['name']}"):
                    set_line_active(sel_unit, l["name"], not l["active"])
                    _flash("ok", f"'{l['name']}' {'deactivated' if l['active'] else 'reactivated'}.")
                    st.rerun()


# ----------------------------------------------------------------------------
# AUTHENTICATION  (accounts live in the database; see Admin tab)
# ----------------------------------------------------------------------------
def load_cookie_cfg():
    try:
        if "cookie" in st.secrets:
            return dict(st.secrets["cookie"])
    except Exception:
        pass
    if os.path.exists("auth_config.yaml"):
        with open("auth_config.yaml") as f:
            return yaml.load(f, Loader=SafeLoader)["cookie"]
    st.error("No [cookie] section found in Secrets.")
    st.stop()


cookie_cfg = load_cookie_cfg()
credentials = load_auth_users()  # cached; passwords are already bcrypt hashes -> no hashing per rerun
if not any(v.get("role") == "admin" for v in credentials["usernames"].values()) and ensure_admin_account():
    credentials = load_auth_users()

authenticator = stauth.Authenticate(
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
is_admin = user_role == "admin"
can_enter_data = user_role in ("entry", "admin")

# Units this account may see: admin = every active unit; others = their assigned active units.
_active_units = [u["name"] for u in load_units() if u["active"]]
if is_admin:
    allowed_units = _active_units
else:
    allowed_units = [u for u in (user_cfg.get("units") or []) if u in _active_units]

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
    elif allowed_units:
        active_unit = allowed_units[0]
    else:
        active_unit = None

    authenticator.logout("Logout", "sidebar")

if active_unit is None:
    brand_header()
    if is_admin:
        st.warning("No active units yet. Add one below.")
        render_admin()
    else:
        st.warning("No unit is assigned to your account yet. Please contact the admin.")
    st.stop()

active_lines = lines_for(active_unit, True)    # shown in entry forms
all_lines = lines_for(active_unit, False)      # used for filters/charts (keeps history of retired lines)

if "selected_date" not in st.session_state:
    st.session_state.selected_date = today_pkt()

st_autorefresh(interval=30 * 60_000, key="global_refresh")  # refresh only every 30 min; otherwise only on manual reload

# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
brand_header(unit=active_unit)
st.markdown("<hr style='margin-top:2px; margin-bottom:10px;'>", unsafe_allow_html=True)

_labels = (["📝 Data Entry"] if can_enter_data else []) + ["📊 Dashboard"] + (["⚙️ Admin"] if is_admin else [])
_tabs = iter([st.container()] if len(_labels) == 1 else st.tabs(_labels))
tab_entry = next(_tabs) if can_enter_data else None
tab_dashboard = next(_tabs)
tab_admin = next(_tabs) if is_admin else None

if is_admin:
    with tab_admin:
        render_admin()

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
                line = st.selectbox("Select Line", active_lines, key=f"line_{fk}")
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
                if not line:
                    st.error("No line is set up for this unit yet. Please ask the admin to add lines.")
                elif minutes_lost is None or minutes_lost <= 0:
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
                    o_line = st.selectbox("Select Line", active_lines, key="o_line")
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
                    if not o_line:
                        st.error("No line is set up for this unit yet. Please ask the admin to add lines.")
                    elif not o_reason.strip():
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
                    rline = st.multiselect("Line", all_lines, key="rec_line")
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
            line_filter = st.multiselect("Filter by Line", all_lines)
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
                chart_lines = all_lines + sorted(set(fdf["line"]) - set(all_lines))
                line_agg = fdf.groupby("line")["minutes"].sum().reindex(chart_lines).fillna(0).reset_index()
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

