# app_fixed.py
# uKids Kids Availability Form
#
# Window:  Monday 05:00 SAST  →  Sunday 23:00 SAST
# Target:  Sunday of that same week
#
# Sheet tabs managed automatically:
#   "uKids Kids SB"  — master child list (Family Code | Name & Surname | Age)
#   "Master Log"     — every submission ever appended (never modified)
#   "22 Jun 2026"    — auto-created after window closes; latest submission per child, flattened
#
# Service columns are always fixed: Morning | Evening  (no ServiceDates sheet needed)

import re
import time
import random
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None

try:
    import gspread
    from gspread.exceptions import APIError, WorksheetNotFound
except Exception:
    gspread = None
    class APIError(Exception):
        pass
    class WorksheetNotFound(Exception):
        pass


# ─────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────
TZ_NAME   = "Africa/Johannesburg"
TAB_SB        = "uKids Kids SB"
TAB_FINAL     = "Final Schedule"
MAX_KIDS      = 5
SERVICE_LABELS = ["Morning", "Evening"]  # fixed service columns — no sheet lookup needed

TEAL   = "#5BC4C0"
ORANGE = "#E8724A"
PURPLE = "#7B4FA6"
YELLOW = "#F5C842"
CREAM  = "#FAF6EE"
WHITE  = "#FFFFFF"
DARK   = "#2A2A2A"
MUTED  = "#7A6F6F"


def weekly_tab_name(service_date_str: str) -> str:
    """'2026-06-22'  ->  '22 Jun 2026'"""
    try:
        dt = datetime.strptime(service_date_str, "%Y-%m-%d")
        return f"{dt.day} {dt.strftime('%b %Y')}"
    except Exception:
        return service_date_str


# ─────────────────────────────────────────────────────────────
# Page config & CSS
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="uKids Availability", layout="centered")

st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap');
  html, body, [class*="css"] {{ font-family: 'Nunito', sans-serif; }}
  :root {{
    --teal:{TEAL}; --orange:{ORANGE}; --purple:{PURPLE};
    --yellow:{YELLOW}; --cream:{CREAM}; --white:{WHITE};
    --dark:{DARK}; --muted:{MUTED};
  }}
  .stApp {{ background:var(--cream); }}
  section[data-testid="stSidebar"] {{ display:none; }}

  /* ── BLACK text on everything by default ── */
  html, body, .stApp, .stApp * {{
    color: var(--dark) !important;
  }}

  /* ── WHITE text on coloured backgrounds — must come AFTER the black rule ── */
  .ukids-hero *,
  .info-banner, .info-banner *,
  .closed-banner, .closed-banner * {{
    color: var(--white) !important;
  }}

  /* ── Buttons always white text ── */
  .stButton > button, .stButton > button * {{
    color: var(--white) !important;
  }}

  /* ── Section pill (yellow background) stays dark ── */
  .section-pill {{ color: var(--dark) !important; }}

  .ukids-hero {{
    background:var(--teal); border-radius:20px;
    padding:34px 28px 28px; margin-bottom:28px;
    text-align:center; position:relative; overflow:hidden;
  }}
  .ukids-hero::before {{
    content:''; position:absolute; top:-40px; right:-40px;
    width:160px; height:160px; background:var(--orange); border-radius:50%; opacity:0.28;
  }}
  .ukids-hero::after {{
    content:''; position:absolute; bottom:-30px; left:-30px;
    width:110px; height:110px; background:var(--purple); border-radius:50%; opacity:0.22;
  }}
  .ukids-hero-content {{ position:relative; z-index:2; }}
  .ukids-hero h1 {{
    font-size:2rem; font-weight:900; color:var(--white);
    margin:0 0 6px; letter-spacing:-0.5px; text-shadow:0 2px 8px rgba(0,0,0,0.15);
  }}
  .ukids-hero p {{ font-size:1rem; color:var(--white); opacity:0.93; margin:0; }}

  .ukids-card {{
    background:var(--white); border-radius:14px;
    padding:20px 20px 16px; margin-bottom:16px;
    box-shadow:0 2px 8px rgba(0,0,0,0.06); border-top:4px solid var(--teal);
  }}
  .ukids-card-orange {{ border-top-color:var(--orange); }}
  .ukids-card-purple {{ border-top-color:var(--purple); }}
  .ukids-card h3 {{
    font-size:1.05rem; font-weight:800; color:var(--dark);
    margin:0 0 12px; text-transform:uppercase; letter-spacing:0.04em;
  }}

  .info-banner {{
    background:linear-gradient(135deg, var(--purple) 0%, #9B6FC6 100%);
    border-radius:12px; padding:14px 18px;
    font-size:0.88rem; color:var(--white); margin-bottom:16px; line-height:1.75;
  }}
  .info-banner strong {{ font-weight:800; }}

  .closed-banner {{
    background:linear-gradient(135deg, var(--purple) 0%, #9B6FC6 100%);
    border-radius:12px; padding:16px 18px; font-size:0.95rem;
    color:var(--white); margin-bottom:16px; line-height:1.7; text-align:center;
  }}

  .kid-block {{
    background:#F7F5FF; border:1.5px solid #D6CCF0;
    border-radius:12px; padding:14px 16px 10px; margin-bottom:14px;
  }}
  .kid-name {{ font-size:1.05rem; font-weight:800; color:var(--purple); }}
  .kid-age  {{ font-size:0.82rem; color:var(--muted); margin-bottom:8px; }}

  .child-block {{
    background:#FFF8F0; border:1.5px solid #FCDAB8;
    border-radius:12px; padding:16px 16px 10px; margin-bottom:14px;
  }}
  .child-block-label {{
    font-size:0.75rem; font-weight:800; color:var(--orange);
    text-transform:uppercase; letter-spacing:0.06em; margin-bottom:8px;
  }}

  hr {{ border-color:#EDE8DC; margin:10px 0; }}

  .sticky-submit {{
    position:sticky; bottom:0; z-index:999;
    background:var(--cream); padding:10px 0 4px; border-top:2px solid #EDE8DC;
  }}

  .stButton > button {{
    width:100%; height:52px; font-size:1.05rem; font-weight:800;
    border-radius:12px; background:var(--teal) !important;
    color:var(--white) !important; border:none !important;
    letter-spacing:0.02em; transition:background 0.2s, transform 0.1s;
  }}
  .stButton > button:hover {{ background:#3AADA9 !important; transform:translateY(-1px); }}
  .stButton > button:active {{ transform:translateY(0); }}

  .section-pill {{
    display:inline-block; background:var(--yellow); color:var(--dark);
    font-size:0.7rem; font-weight:800; text-transform:uppercase;
    letter-spacing:0.08em; padding:3px 10px; border-radius:20px; margin-bottom:10px;
  }}

  @media (max-width:520px) {{
    div[data-testid="column"] {{ width:100% !important; flex:0 0 100% !important; }}
    .ukids-hero h1 {{ font-size:1.55rem; }}
  }}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="ukids-hero">
  <div class="ukids-hero-content">
    <h1>uKids Availability</h1>
    <p>Let us know which services your children can serve at this week.</p>
  </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# Secrets
# ─────────────────────────────────────────────────────────────
def _get_secret_any(*paths):
    try:
        cur = st.secrets
    except Exception:
        return None
    for path in paths:
        c, ok = cur, True
        for k in path:
            if k in c:
                c = c[k]
            else:
                ok = False
                break
        if ok:
            return c
    return None


ADMIN_KEY = str(_get_secret_any(["ADMIN_KEY"], ["general", "ADMIN_KEY"]) or "")


def is_sheets_enabled() -> bool:
    if gspread is None:
        return False
    return bool(
        _get_secret_any(["gcp_service_account"], ["general", "gcp_service_account"])
        and _get_secret_any(["GSHEET_ID"], ["general", "GSHEET_ID"])
    )


if not is_sheets_enabled():
    st.error("Google Sheets is not configured. Add GSHEET_ID and [gcp_service_account] to Secrets.")
    st.stop()


# ─────────────────────────────────────────────────────────────
# Google Sheets connection
# ─────────────────────────────────────────────────────────────
def gs_retry(func, *args, **kwargs):
    for attempt in range(6):
        try:
            return func(*args, **kwargs)
        except APIError as e:
            status = getattr(getattr(e, "response", None), "status_code", None)
            if status in (429, 500, 502, 503):
                time.sleep(min(15, (2 ** attempt) + random.random()))
                continue
            raise


@st.cache_resource
def get_spreadsheet():
    sa       = _get_secret_any(["gcp_service_account"], ["general", "gcp_service_account"])
    sheet_id = _get_secret_any(["GSHEET_ID"], ["general", "GSHEET_ID"])
    if not sa or not sheet_id:
        raise RuntimeError("Missing GSHEET_ID or gcp_service_account in secrets.")
    sa = dict(sa)
    pk = sa.get("private_key", "")
    if isinstance(pk, str):
        pk = pk.replace("\\n", "\n").strip()
        if not pk.endswith("\n"):
            pk += "\n"
        sa["private_key"] = pk
    gc = gspread.service_account_from_dict(sa)
    return gs_retry(gc.open_by_key, sheet_id)


def ensure_worksheet(sh, title: str, rows: int = 2000, cols: int = 50):
    try:
        return sh.worksheet(title)
    except WorksheetNotFound:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)


def make_unique_headers(header: list) -> list:
    counts: dict = {}
    out = []
    for h in header:
        base = str(h).strip() or "Unnamed"
        counts[base] = counts.get(base, 0) + 1
        out.append(base if counts[base] == 1 else f"{base}_{counts[base]}")
    return out


def ws_get_values_and_df(ws):
    values = gs_retry(ws.get_all_values)
    if not values:
        return [], [], pd.DataFrame()
    header_raw    = values[0]
    header_unique = make_unique_headers(header_raw)
    rows          = values[1:]
    return header_raw, header_unique, pd.DataFrame(rows, columns=header_unique)


def ws_ensure_header(ws, desired_header: list) -> list:
    header = gs_retry(ws.row_values, 1)
    if not header:
        gs_retry(ws.update, "1:1", [desired_header])
        return desired_header
    missing = [c for c in desired_header if c not in header]
    if missing:
        header = header + missing
        gs_retry(ws.update, "1:1", [header])
    return header




# ─────────────────────────────────────────────────────────────
# Sheet writing — Master Log + auto weekly summary
# ─────────────────────────────────────────────────────────────
TAB_MASTER = "Master Log"
BASE_COLS  = ["timestamp", "Service date", "Family Code", "Name", "Age"]
FLAT_COLS  = ["timestamp", "Service date", "Service", "Name", "Age"]


def append_to_master_log(rows_to_write: list, service_labels: list):
    """
    Appends every submission raw to 'Master Log'.
    Columns: timestamp | Service date | Family Code | Name | Age | <service labels>
    Every submit = new rows, nothing is ever overwritten here.
    """
    sh             = get_spreadsheet()
    desired_header = BASE_COLS + service_labels
    ws             = ensure_worksheet(sh, TAB_MASTER, rows=20000, cols=len(desired_header) + 5)
    ws_ensure_header(ws, desired_header)
    rows_as_lists  = [[r.get(c, "") for c in desired_header] for r in rows_to_write]
    gs_retry(ws.append_rows, rows_as_lists)


def build_weekly_summary(service_date_str: str):
    """
    Reads Master Log, filters to this week, keeps only the LATEST
    submission per child (by timestamp), flattens to one row per child
    per Yes service, and writes to a tab named e.g. '22 Jun 2026'.
    Format: timestamp | Service date | Service | Name | Age

    Service labels are derived directly from the Master Log columns
    so this function needs no external inputs beyond the date.
    """
    sh       = get_spreadsheet()
    tab_name = weekly_tab_name(service_date_str)

    ws_master = ensure_worksheet(sh, TAB_MASTER, rows=20000, cols=100)
    _, _, df  = ws_get_values_and_df(ws_master)

    if df.empty or "Service date" not in df.columns:
        return

    # Filter to this week only
    week_df = df[df["Service date"].str.strip() == service_date_str].copy()
    if week_df.empty:
        return

    # Derive service label columns from the Master Log header
    # (everything that isn't a base column)
    service_labels = [c for c in week_df.columns if c not in BASE_COLS]

    # Keep only the latest submission per child
    if "timestamp" in week_df.columns:
        week_df = week_df.sort_values("timestamp")
    week_df = week_df.drop_duplicates(subset=["Name"], keep="last")

    # Flatten to one row per child per Yes service
    rows_out = []
    for _, row in week_df.iterrows():
        for svc in service_labels:
            if str(row.get(svc, "")).strip().lower() == "yes":
                rows_out.append([
                    row.get("timestamp", ""),
                    row.get("Service date", ""),
                    svc,
                    row.get("Name", ""),
                    row.get("Age", ""),
                ])

    ws_week = ensure_worksheet(sh, tab_name, rows=2000, cols=10)
    ws_week.clear()
    gs_retry(ws_week.update, "A1", [FLAT_COLS])
    if rows_out:
        gs_retry(ws_week.append_rows, rows_out)


def maybe_build_weekly_summary(service_date_str: str):
    """
    Called on every page load when the window is closed.
    Checks if the weekly summary tab exists and has data —
    builds it automatically if not. Safe to call repeatedly.
    """
    sh       = get_spreadsheet()
    tab_name = weekly_tab_name(service_date_str)

    try:
        ws_existing = sh.worksheet(tab_name)
        values      = gs_retry(ws_existing.get_all_values)
        if values and len(values) > 1:
            return  # Already built, nothing to do
    except WorksheetNotFound:
        pass  # Tab doesn't exist yet — build it

    build_weekly_summary(service_date_str)



# ─────────────────────────────────────────────────────────────
# Data loading — session_state based, fetch once per session
# ─────────────────────────────────────────────────────────────
def _raw_fetch_sb() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_SB, rows=4000, cols=20)
    _, _, df = ws_get_values_and_df(ws)
    df["Family Code"]    = df["Family Code"].astype(str).str.strip()
    df["Name & Surname"] = df["Name & Surname"].astype(str).str.strip()
    df["Age"]            = df["Age"].astype(str).str.strip()
    return df


def load_all_data(force: bool = False):
    if not force and st.session_state.get("data_loaded"):
        return True, ""

    try:
        sb = _raw_fetch_sb()
    except Exception as e:
        return False, str(e)

    miss = {"Family Code", "Name & Surname", "Age"} - set(sb.columns)
    if miss:
        return False, f"Sheet '{TAB_SB}' missing columns: {', '.join(sorted(miss))}"

    st.session_state["sb_df"]       = sb
    st.session_state["data_loaded"] = True
    return True, ""


# ─────────────────────────────────────────────────────────────
# Time helpers
# ─────────────────────────────────────────────────────────────
def get_now_sast() -> datetime:
    if ZoneInfo:
        return datetime.now(ZoneInfo(TZ_NAME))
    return datetime.utcnow() + timedelta(hours=2)


def get_window():
    now               = get_now_sast()
    days_since_monday = now.weekday()
    monday            = now.replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=days_since_monday)
    opening_dt        = monday.replace(hour=5,  minute=0, second=0, microsecond=0)
    deadline_dt       = monday.replace(hour=23, minute=0, second=0, microsecond=0) + timedelta(days=6)  # Sunday 23:00
    sunday_dt         = monday + timedelta(days=6)
    service_date_str  = sunday_dt.strftime("%Y-%m-%d")
    is_open           = opening_dt <= now < deadline_dt
    return is_open, service_date_str, opening_dt, deadline_dt, now


def format_time_remaining(delta_seconds: float) -> str:
    secs = max(0, int(delta_seconds))
    hrs  = secs // 3600
    mins = (secs % 3600) // 60
    return f"{hrs}h {mins}m" if hrs > 0 else f"{mins}m"


# ─────────────────────────────────────────────────────────────
# Pure data helpers
# ─────────────────────────────────────────────────────────────
def _clean_cell(x) -> str:
    s = str(x).strip()
    return "" if (s == "" or s.lower() == "nan") else s


def _is_truthy(v) -> bool:
    return str(v).strip().lower() in ("1", "true", "yes", "y", "t")


def _safe_parse_date_ymd(s: str) -> datetime:
    try:
        return datetime.strptime(str(s).strip(), "%Y-%m-%d")
    except Exception:
        return datetime(1900, 1, 1)


def age_display(age_raw: str) -> str:
    age_str = _clean_cell(age_raw)
    if not age_str:
        return ""
    if "Born" in age_str or "Age" in age_str:
        return age_str
    try:
        age_int = int(float(age_str))
        return f"Age {age_int} (Born {datetime.now().year - age_int})"
    except Exception:
        return age_str


def age_to_number(age_raw: str) -> str:
    age_str = _clean_cell(age_raw)
    if not age_str:
        return ""
    if age_str.isdigit():
        return age_str
    m = re.search(r"\d+", age_str)
    return m.group() if m else age_str


# ─────────────────────────────────────────────────────────────
# Registration helpers
# ─────────────────────────────────────────────────────────────
def is_duplicate_name(name: str, df: pd.DataFrame) -> bool:
    existing = df["Name & Surname"].str.strip().str.lower().tolist()
    return name.strip().lower() in existing


def get_next_family_code(df: pd.DataFrame) -> int:
    codes = pd.to_numeric(df["Family Code"], errors="coerce").dropna()
    return int(codes.max()) + 1 if not codes.empty else 1


def get_family_code_for_sibling(sibling_name: str, df: pd.DataFrame) -> str:
    rows = df[df["Name & Surname"].str.strip() == sibling_name.strip()]
    if rows.empty:
        return ""
    return _clean_cell(rows.iloc[0].get("Family Code", ""))


def register_children_batch(children: list, family_code: str):
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_SB, rows=4000, cols=20)
    ws_ensure_header(ws, ["Family Code", "Name & Surname", "Age"])

    _, _, latest_df = ws_get_values_and_df(ws)
    if "Name & Surname" not in latest_df.columns:
        existing_names = set()
    else:
        existing_names = set(latest_df["Name & Surname"].str.strip().str.lower().tolist())

    rows_to_add, skipped = [], []
    for child in children:
        name_clean = child["name"].strip()
        if not name_clean:
            continue
        if name_clean.lower() in existing_names:
            skipped.append(name_clean)
        else:
            rows_to_add.append([family_code, name_clean, child["age"]])
            existing_names.add(name_clean.lower())

    if rows_to_add:
        gs_retry(ws.append_rows, rows_to_add)

    st.session_state["data_loaded"] = False
    return len(rows_to_add), skipped


# ═════════════════════════════════════════════════════════════
# LOAD DATA
# ═════════════════════════════════════════════════════════════
with st.spinner("Loading..."):
    ok, err = load_all_data()

if not ok:
    st.error(f"Could not load data from Google Sheets: {err}")
    st.stop()

sb_df = st.session_state["sb_df"]


# ═════════════════════════════════════════════════════════════
# SECTION 1 — AVAILABILITY FORM
# ═════════════════════════════════════════════════════════════
is_open, service_date_key, opening_dt, deadline_dt, now_sast = get_window()
remaining_seconds = (deadline_dt - now_sast).total_seconds() if is_open else 0

if is_open:
    st.markdown(f"""
<div class="info-banner">
  Submitting availability for <strong>{weekly_tab_name(service_date_key)}</strong> (Sunday service)<br>
  Form open from <strong>{opening_dt.strftime('%A %d %b, %H:%M')}</strong>
  until <strong>{deadline_dt.strftime('%A %d %b, %H:%M')}</strong> SAST<br>
  Time remaining: <strong>{format_time_remaining(remaining_seconds)}</strong>
  &nbsp;&mdash;&nbsp; You can submit more than once.
</div>
""", unsafe_allow_html=True)

    if st.button("Refresh timer"):
        st.rerun()

    st.markdown("---")
    st.markdown("**Generate weekly summary early**")
    st.caption(
        "Builds the summary sheet from all responses received so far. "
        "Run this any time during the week — or wait until the deadline and it happens automatically."
    )
    if st.button("Generate summary now", key="btn_generate_open"):
        try:
            with st.spinner(f"Building '{weekly_tab_name(service_date_key)}' from Master Log..."):
                build_weekly_summary(service_date_key)
            st.success(
                f"Done! Sheet **'{weekly_tab_name(service_date_key)}'** updated with latest responses."
            )
        except Exception as e:
            st.error(f"Could not generate summary: {e}")
    st.markdown("---")

    date_labels = SERVICE_LABELS

    child_names = sorted({
        n for n in sb_df["Name & Surname"].tolist()
        if n and str(n).strip().lower() not in ("", "nan")
    })

    def get_family_code_for_child(name: str) -> str:
        rows = sb_df[sb_df["Name & Surname"] == name]
        return _clean_cell(rows.iloc[0].get("Family Code", "")) if not rows.empty else ""

    def get_kids_info(family_code: str) -> list:
        rows = sb_df[sb_df["Family Code"] == family_code].copy()
        out  = []
        for idx, row in rows.reset_index(drop=True).iterrows():
            kid_name = _clean_cell(row.get("Name & Surname", ""))
            age_val  = _clean_cell(row.get("Age", ""))
            if kid_name:
                out.append({"slot": idx + 1, "name": kid_name, "age": age_val})
        return out

    st.markdown("""
<div class="ukids-card">
  <span class="section-pill">Step 1</span>
  <h3>Submit your child's availability</h3>
</div>
""", unsafe_allow_html=True)

    selected_child = st.selectbox(
        "Search for one child from your family",
        options=[""] + child_names,
        index=0,
        key="main_child_select",
    )

    if not selected_child:
        st.caption("Select one child to load your whole family's form.")

    else:
        family_code = get_family_code_for_child(selected_child)
        if not family_code:
            st.error("Could not find a family code for that child.")

        else:
            kids_info = get_kids_info(family_code)
            if not kids_info:
                st.warning("No children found for this family in the sheet.")

            else:
                st.markdown(f"""
<div class="ukids-card ukids-card-purple">
  <span class="section-pill">Step 2</span>
  <h3>Select services — {weekly_tab_name(service_date_key)}</h3>
</div>
""", unsafe_allow_html=True)

                kids_selected_map: dict = {}

                for k in kids_info:
                    slot   = k["slot"]
                    k_name = k["name"]
                    k_age  = age_display(k["age"])

                    st.markdown(f"""
<div class="kid-block">
  <div class="kid-name">{k_name}</div>
  <div class="kid-age">{k_age}</div>
</div>
""", unsafe_allow_html=True)
                    st.caption("Which services can this child attend?")

                    # Display label includes the date so parents know exactly
                    # which Sunday they're confirming — stored value stays "Morning"/"Evening"
                    sunday_display = weekly_tab_name(service_date_key)
                    selected_labels = []
                    for svc in date_labels:
                        display_label = f"{svc} — {sunday_display}"
                        if st.checkbox(display_label, key=f"svc_{slot}_{service_date_key}_{svc}"):
                            selected_labels.append(svc)
                    kids_selected_map[slot] = set(selected_labels)
                    st.divider()

                st.markdown("**Review**")
                for k in kids_info:
                    count = len(kids_selected_map.get(k["slot"], set()))
                    st.write(f"**{k['name']}** — {count} service{'s' if count != 1 else ''} selected")

                st.markdown('<div class="sticky-submit">', unsafe_allow_html=True)
                submitted = st.button("Submit availability")
                st.markdown("</div>", unsafe_allow_html=True)

                if submitted:
                    is_still_open, _, _, _, _ = get_window()
                    if not is_still_open:
                        st.error("The form has just closed. Submissions are no longer accepted.")
                    else:
                        now_iso       = datetime.utcnow().isoformat() + "Z"
                        rows_to_write = []

                        for k in kids_info:
                            slot    = k["slot"]
                            row_map = {
                                "timestamp":    now_iso,
                                "Service date": service_date_key,
                                "Family Code":  family_code,
                                "Name":         k["name"],
                                "Age":          age_to_number(k["age"]),
                            }
                            for svc in date_labels:
                                row_map[svc] = "Yes" if svc in kids_selected_map.get(slot, set()) else "No"
                            rows_to_write.append(row_map)

                        try:
                            with st.spinner("Saving..."):
                                append_to_master_log(rows_to_write, date_labels)
                            n = len(rows_to_write)
                            st.success(
                                f"Submitted. Availability saved for "
                                f"{n} child{'ren' if n != 1 else ''}."
                            )
                            st.balloons()
                        except Exception as e:
                            st.error(f"Failed to save: {e}")

else:
    now  = get_now_sast()
    wday = now.weekday()
    days_to_next_monday = (7 - wday) % 7 or 7
    next_monday   = (now + timedelta(days=days_to_next_monday)).replace(hour=5, minute=0, second=0, microsecond=0)
    next_sunday   = next_monday + timedelta(days=6)
    next_deadline = next_monday + timedelta(days=6, hours=18)  # Sunday 23:00

    # Automatically build the weekly summary tab if it hasn't been created yet
    try:
        maybe_build_weekly_summary(service_date_key)
    except Exception:
        pass  # Never block the UI if this fails

    st.markdown(f"""
<div class="closed-banner">
  <strong>The availability form is currently closed.</strong><br><br>
  The next window opens <strong>{next_monday.strftime('%A %d %b at %H:%M')}</strong> SAST<br>
  and closes <strong>{next_deadline.strftime('%A %d %b at %H:%M')}</strong> SAST<br>
  for the <strong>{next_sunday.strftime('%d %b %Y')}</strong> Sunday service.
</div>
""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("**Generate weekly summary early**")
    st.caption(
        "The summary sheet is created automatically at the deadline, but you can "
        "generate it early at any point during the week to see current responses."
    )
    if st.button("Generate summary now", key="btn_generate_early"):
        try:
            with st.spinner(f"Building '{weekly_tab_name(service_date_key)}' from Master Log..."):
                build_weekly_summary(service_date_key)
            st.success(
                f"Done! Sheet **'{weekly_tab_name(service_date_key)}'** has been updated "
                "with the latest responses. You can run this again at any time."
            )
        except Exception as e:
            st.error(f"Could not generate summary: {e}")



# ═════════════════════════════════════════════════════════════
# SECTION 2 — FAMILY REGISTRATION
# ═════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("""
<div class="ukids-card ukids-card-orange">
  <span class="section-pill">New to uKids serving?</span>
  <h3>Register your family</h3>
</div>
""", unsafe_allow_html=True)

with st.expander("Open registration form", expanded=False):
    st.caption(
        "Use this section if any of your children are not yet on the list. "
        "You can register your whole family at once. "
        "Children already on the list will be skipped automatically."
    )

    num_children = st.number_input(
        "How many children are you registering?",
        min_value=1, max_value=MAX_KIDS, value=1, step=1,
        key="reg_num_children",
    )

    reg_children = []
    for i in range(int(num_children)):
        ordinal = ["First", "Second", "Third", "Fourth", "Fifth"][i]
        st.markdown(
            f'<div class="child-block"><div class="child-block-label">{ordinal} child</div>',
            unsafe_allow_html=True,
        )
        c_name = st.text_input(
            "Full name (Name & Surname)",
            placeholder="e.g. Levi Van der Vyver",
            key=f"reg_child_name_{i}",
        )
        c_birth = st.number_input(
            "Birth year",
            min_value=2000, max_value=datetime.now().year,
            value=datetime.now().year - 8, step=1,
            key=f"reg_child_birth_{i}",
        )
        c_age = str(datetime.now().year - int(c_birth))
        if c_name.strip():
            st.caption(f"Will be stored as: **{age_display(c_age)}**")
        st.markdown("</div>", unsafe_allow_html=True)
        reg_children.append({"name": c_name.strip(), "age": c_age})

    st.markdown("---")
    st.markdown("**Is any sibling from your family already on the list?**")
    st.caption(
        "If a sibling is already registered, select their name so the new children "
        "are linked to the same family. Otherwise a new family code will be created."
    )

    link_choice = st.radio(
        "Family link",
        options=["Yes — pick an existing sibling", "No — this is a brand-new family"],
        key="reg_link_choice",
        label_visibility="collapsed",
    )

    resolved_fam_code = None
    chosen_sibling    = None

    if link_choice == "Yes — pick an existing sibling":
        reg_child_list = sorted({
            n for n in sb_df["Name & Surname"].tolist()
            if n and str(n).strip().lower() not in ("", "nan")
        })
        chosen_sibling = st.selectbox(
            "Select an existing sibling",
            options=[""] + reg_child_list,
            key="reg_sibling",
        )
        if chosen_sibling:
            resolved_fam_code = get_family_code_for_sibling(chosen_sibling, sb_df)
            if resolved_fam_code:
                st.success(f"New children will be added to Family Code **{resolved_fam_code}**.")
            else:
                st.error("Could not find a family code for that sibling. Try the 'new family' option.")
    else:
        resolved_fam_code = str(get_next_family_code(sb_df))
        st.info(f"A new Family Code **{resolved_fam_code}** will be assigned to your family.")

    all_names_filled = all(c["name"] for c in reg_children)
    sibling_ok = (
        link_choice != "Yes — pick an existing sibling"
        or (chosen_sibling and resolved_fam_code)
    )
    reg_ready = all_names_filled and sibling_ok and bool(resolved_fam_code)

    if st.button("Register family", disabled=not reg_ready, key="btn_register"):
        valid_children = [c for c in reg_children if c["name"]]
        all_dupes = [c["name"] for c in valid_children if is_duplicate_name(c["name"], sb_df)]
        non_dupes = [c for c in valid_children if not is_duplicate_name(c["name"], sb_df)]

        if all_dupes and not non_dupes:
            st.error(
                f"All the names you entered are already on the list: "
                f"**{', '.join(all_dupes)}**. "
                "If you believe this is a different child with the same name, "
                "please contact **Ps Some** to resolve this manually."
            )
        else:
            try:
                with st.spinner("Registering..."):
                    written, skipped = register_children_batch(valid_children, resolved_fam_code)
                    load_all_data(force=True)
                    sb_df = st.session_state["sb_df"]

                if written > 0:
                    names_written = [
                        c["name"] for c in valid_children
                        if c["name"].lower() not in [s.lower() for s in skipped]
                    ]
                    st.success(
                        f"Registered {written} child{'ren' if written != 1 else ''}: "
                        f"**{', '.join(names_written)}** (Family Code {resolved_fam_code}). "
                        "They now appear in the availability form above."
                    )
                    st.balloons()

                if skipped:
                    st.warning(
                        f"Already on the list (skipped): **{', '.join(skipped)}**. "
                        "If you think this is a different child with the same name, "
                        "please contact **Ps Some**."
                    )
            except Exception as e:
                st.error(f"Registration failed: {e}")
