# app_fixed.py
# uKids Kids Availability Form
# - Family surname dropdown
# - Auto-detect kids (Kid #1, Kid #2, Kid #3...) from "uKids Kids SB"
# - For EACH kid: Morning + Evening checkbox questions (mobile friendly)
# - OUTPUT: Appends ONE ROW PER KID into "uKids Kids responses"
#           columns: timestamp | Availability month | Family Surname | Serving kid | <service labels...>

import time
import random
from datetime import datetime

import pandas as pd
import streamlit as st

# ✅ Timezone-aware deadlines
try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None  # fallback

# Optional: Google Sheets libs.
try:
    import gspread
    from gspread.exceptions import APIError, WorksheetNotFound
except Exception:
    gspread = None

    class APIError(Exception):
        ...

    class WorksheetNotFound(Exception):
        ...


# ──────────────────────────────────────────────────────────────────────────────
# UI CONFIG + mobile tweaks
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="uKids Kids Availability Form", page_icon="🗓️", layout="centered")
st.title("🗓️ uKids Kids Availability Form")

st.markdown(
    """
<style>
  .stButton > button { width: 100%; height: 48px; font-size: 16px; }
  @media (max-width: 520px){
    div[data-testid="column"] { width: 100% !important; flex: 0 0 100% !important; }
    pre, code { font-size: 15px; line-height: 1.35; }
  }
  .sticky-submit {
    position: sticky; bottom: 0; z-index: 999;
    background: #fff; padding: 10px 0; border-top: 1px solid #eee;
  }
</style>
""",
    unsafe_allow_html=True,
)

# ──────────────────────────────────────────────────────────────────────────────
# Google Sheets tabs
# ──────────────────────────────────────────────────────────────────────────────
TAB_RESPONSES = "uKids Kids responses"   # ✅ output tab (rows per kid)
TAB_SB = "uKids Kids SB"                 # ✅ your family/kids SB
TAB_DEADLINES = "Deadlines"
TAB_DATES = "Kids & Guys ServiceDates"

# ──────────────────────────────────────────────────────────────────────────────
# Secrets helpers
# ──────────────────────────────────────────────────────────────────────────────
def _get_secret_any(*paths):
    """Try multiple secret paths, return the first value found."""
    try:
        cur = st.secrets
    except Exception:
        return None
    for path in paths:
        c = cur
        ok = True
        for k in path:
            if k in c:
                c = c[k]
            else:
                ok = False
                break
        if ok:
            return c
    return None


def get_admin_key() -> str:
    v = _get_secret_any(["ADMIN_KEY"], ["general", "ADMIN_KEY"])
    return str(v) if v else ""


ADMIN_KEY = get_admin_key()


def is_sheets_enabled() -> bool:
    if gspread is None:
        return False
    sa = _get_secret_any(["gcp_service_account"], ["general", "gcp_service_account"])
    sid = _get_secret_any(["GSHEET_ID"], ["general", "GSHEET_ID"])
    return bool(sa and sid)


if not is_sheets_enabled():
    st.error("Google Sheets is not configured. Add GSHEET_ID and [gcp_service_account] to Secrets.")
    st.stop()


# ──────────────────────────────────────────────────────────────────────────────
# Google Sheets helpers
# ──────────────────────────────────────────────────────────────────────────────
def gs_retry(func, *args, **kwargs):
    for attempt in range(5):
        try:
            return func(*args, **kwargs)
        except APIError as e:
            status = getattr(getattr(e, "response", None), "status_code", None)
            if status in (429, 500, 502, 503):
                time.sleep(min(10, (2**attempt) + random.random()))
                continue
            raise


@st.cache_resource
def get_spreadsheet():
    """
    Open the spreadsheet and return gspread Spreadsheet object.
    Includes private_key newline fixer.
    """
    sa = _get_secret_any(["gcp_service_account"], ["general", "gcp_service_account"])
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


def ws_get_df(ws) -> pd.DataFrame:
    values = gs_retry(ws.get_all_values)
    if not values:
        return pd.DataFrame()
    header, rows = values[0], values[1:]
    return pd.DataFrame(rows, columns=header)


def ws_ensure_header(ws, desired_header: list[str]) -> list[str]:
    header = gs_retry(ws.row_values, 1)
    if not header:
        gs_retry(ws.update, "1:1", [desired_header])
        return desired_header
    missing = [c for c in desired_header if c not in header]
    if missing:
        header = header + missing
        gs_retry(ws.update, "1:1", [header])
    return header


@st.cache_data(ttl=30, show_spinner=False)
def fetch_sb_df() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_SB, rows=4000, cols=50)
    return ws_get_df(ws)


@st.cache_data(ttl=30, show_spinner=False)
def fetch_deadlines_df() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_DEADLINES, rows=500, cols=10)
    return ws_get_df(ws)


@st.cache_data(ttl=30, show_spinner=False)
def fetch_service_dates_df() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_DATES, rows=4000, cols=20)
    return ws_get_df(ws)


@st.cache_data(ttl=30, show_spinner=False)
def fetch_responses_df() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_RESPONSES, rows=8000, cols=400)
    return ws_get_df(ws)


def append_response_row(desired_header: list[str], row_map: dict):
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_RESPONSES, rows=8000, cols=max(400, len(desired_header) + 10))
    header = ws_ensure_header(ws, desired_header)
    row = [row_map.get(col, "") for col in header]
    gs_retry(ws.append_row, row)


def clear_caches():
    for fn in (fetch_sb_df, fetch_deadlines_df, fetch_service_dates_df, fetch_responses_df):
        try:
            fn.clear()
        except Exception:
            pass
    try:
        st.cache_data.clear()
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────
# Time helpers
# ─────────────────────────────────────────────────────────────
def get_now_in_tz(tz_name: str) -> datetime:
    if ZoneInfo is None:
        return datetime.utcnow()
    return datetime.now(ZoneInfo(tz_name))


def add_one_month(dt: datetime) -> datetime:
    y, m = dt.year, dt.month
    if m == 12:
        y2, m2 = y + 1, 1
    else:
        y2, m2 = y, m + 1
    if dt.tzinfo:
        return datetime(y2, m2, 1, tzinfo=dt.tzinfo)
    return datetime(y2, m2, 1)


def get_target_month_key(now_local: datetime) -> str:
    return add_one_month(now_local).strftime("%Y-%m")


def parse_deadline_local(deadline_local: str, tz_name: str) -> datetime:
    dt_naive = datetime.strptime(deadline_local, "%Y-%m-%d %H:%M")
    if ZoneInfo is None:
        return dt_naive
    return dt_naive.replace(tzinfo=ZoneInfo(tz_name))


def format_minutes_remaining(delta_seconds: float) -> str:
    mins = max(0, int(delta_seconds // 60))
    hrs = mins // 60
    rem_m = mins % 60
    return f"{hrs}h {rem_m}m" if hrs > 0 else f"{rem_m}m"


def _safe_parse_date_ymd(s: str) -> datetime:
    try:
        return datetime.strptime(str(s).strip(), "%Y-%m-%d")
    except Exception:
        return datetime(1900, 1, 1)


def _is_truthy_service_day(v) -> bool:
    s = str(v).strip().lower()
    return s in ("1", "true", "yes", "y", "t")


def _display_date_only(label: str) -> str:
    s = str(label).strip()
    s = s.replace("Morning Service", "").replace("Evening Service", "")
    s = s.replace("Morning", "").replace("Evening", "")
    s = s.replace("Service", "")
    return " ".join(s.split()).strip(" -")


def _build_display_map(labels: list[str]) -> dict:
    display_map, used = {}, set()
    for lbl in labels:
        base = _display_date_only(lbl)
        disp = base
        i = 2
        while disp in used:
            disp = f"{base} ({i})"
            i += 1
        used.add(disp)
        display_map[disp] = lbl
    return display_map


# ─────────────────────────────────────────────────────────────
# Load config
# ─────────────────────────────────────────────────────────────
try:
    sb_df = fetch_sb_df()
    deadlines_df = fetch_deadlines_df()
    service_dates_all = fetch_service_dates_df()
except Exception as e:
    st.error(f"Failed to load config from Google Sheets: {e}")
    st.stop()

# Validate required columns
for df, tab, needed in [
    (deadlines_df, TAB_DEADLINES, {"month", "deadline_local", "timezone"}),
    (service_dates_all, TAB_DATES, {"target_month", "date", "label", "is_service_day"}),
]:
    miss = needed - set(df.columns)
    if miss:
        st.error(f"Google Sheet tab '{tab}' is missing columns: {', '.join(sorted(miss))}")
        st.stop()

# Validate SB columns
sb_needed = {"Family Surname", "Kid #1"}
miss_sb = sb_needed - set(sb_df.columns)
if miss_sb:
    st.error(f"Google Sheet tab '{TAB_SB}' is missing columns: {', '.join(sorted(miss_sb))}")
    st.stop()

# Clean deadlines + dates
deadlines_df["month"] = deadlines_df["month"].astype(str).str.strip()
deadlines_df["deadline_local"] = deadlines_df["deadline_local"].astype(str).str.strip()
deadlines_df["timezone"] = deadlines_df["timezone"].astype(str).str.strip()

service_dates_all["target_month"] = service_dates_all["target_month"].astype(str).str.strip()
service_dates_all["date"] = service_dates_all["date"].astype(str).str.strip()
service_dates_all["label"] = service_dates_all["label"].astype(str).str.strip()
service_dates_all["is_service_day"] = service_dates_all["is_service_day"].astype(str).str.strip()

# Base TZ
BASE_TZ = "Africa/Johannesburg"
try:
    tz0 = str(deadlines_df["timezone"].iloc[0]).strip()
    if tz0:
        BASE_TZ = tz0
except Exception:
    pass

now_base = get_now_in_tz(BASE_TZ)
target_month_key = get_target_month_key(now_base)

# Filter month services
month_dates = service_dates_all[
    (service_dates_all["target_month"] == target_month_key)
    & (service_dates_all["is_service_day"].map(_is_truthy_service_day))
].copy()

if month_dates.empty:
    st.markdown(
        f"""
        ## 🔒 This month’s availability form is not open yet.

        No service dates were found for **{target_month_key}**.
        """
    )
    st.stop()

month_dates["_sort"] = month_dates["date"].map(_safe_parse_date_ymd)
month_dates = month_dates.sort_values("_sort").drop(columns=["_sort"])

# Service label columns we must write to the sheet
date_labels = month_dates["label"].astype(str).tolist()

# Split (relies on "morning"/"evening" in label text)
morning_labels = [l for l in date_labels if "morning" in l.lower()]
evening_labels = [l for l in date_labels if "evening" in l.lower()]

morning_display_map = _build_display_map(morning_labels)
evening_display_map = _build_display_map(evening_labels)
morning_options = list(morning_display_map.keys())
evening_options = list(evening_display_map.keys())


def get_deadline_for_target_month(deadlines: pd.DataFrame, month_key: str):
    match = deadlines[deadlines["month"] == month_key]
    if match.empty:
        return None, BASE_TZ
    row = match.iloc[0]
    tz_name = str(row["timezone"]).strip() or BASE_TZ
    dl = parse_deadline_local(str(row["deadline_local"]).strip(), tz_name)
    return dl, tz_name


deadline_dt, deadline_tz = get_deadline_for_target_month(deadlines_df, target_month_key)

# Closed check
is_closed = True
if deadline_dt is not None:
    now_local = get_now_in_tz(deadline_tz)
    is_closed = (deadline_dt - now_local).total_seconds() <= 0

if is_closed:
    target_month_name = datetime.strptime(target_month_key, "%Y-%m").strftime("%B")
    st.markdown(f"## 🔒 {target_month_name} availability submissions are now closed.")
    st.stop()

# Info header
now_local = get_now_in_tz(deadline_tz)
remaining_seconds = (deadline_dt - now_local).total_seconds()

st.info(
    f"🗓️ Submitting availability for **{target_month_key}**.\n\n"
    f"⏳ Form closes at **{deadline_dt.strftime('%Y-%m-%d %H:%M')}** ({deadline_tz}). "
    f"Time remaining: **{format_minutes_remaining(remaining_seconds)}**\n\n"
    f"🔁 You can submit more than once — we will use your most recent submission."
)

if st.button("Refresh timer"):
    st.rerun()

# ─────────────────────────────────────────────────────────────
# Build family -> kids mapping
# ─────────────────────────────────────────────────────────────
sb = sb_df.copy()
sb["Family Surname"] = sb["Family Surname"].astype(str).str.strip()

kid_cols = [c for c in sb.columns if str(c).strip().lower().startswith("kid #")]
kid_cols = sorted(
    kid_cols,
    key=lambda x: int("".join([ch for ch in x if ch.isdigit()]) or "999")
)

families = sorted({f for f in sb["Family Surname"].tolist() if f and f.lower() != "nan"})


def get_kids_for_family(family: str) -> list[str]:
    if not family:
        return []
    row = sb[sb["Family Surname"] == family]
    if row.empty:
        return []
    r0 = row.iloc[0]
    kids = []
    for c in kid_cols:
        v = str(r0.get(c, "")).strip()
        if v and v.lower() != "nan":
            kids.append(v)
    return kids


# ─────────────────────────────────────────────────────────────
# Session state
# ─────────────────────────────────────────────────────────────
if "answers" not in st.session_state:
    st.session_state.answers = {}
answers = st.session_state.answers

# ─────────────────────────────────────────────────────────────
# Form UI
# ─────────────────────────────────────────────────────────────
st.subheader("Your details")
answers["FAMILY_SURNAME"] = st.selectbox("Family surname", options=[""] + families, index=0)

if not answers.get("FAMILY_SURNAME"):
    st.caption("Select your family surname to continue.")
    st.stop()

kids = get_kids_for_family(answers.get("FAMILY_SURNAME", ""))

if not kids:
    st.warning("No kids found for this family in 'uKids Kids SB'. Please fill in Kid #1 (and Kid #2/#3 if applicable).")
    st.stop()

st.subheader(f"Availability for {target_month_key}")

# We collect selections per kid: kid_name -> set(actual_label)
kids_selected_map: dict[str, set[str]] = {}

for idx, kid_name in enumerate(kids, start=1):
    st.markdown(f"## 👦 {kid_name}")

    # Morning
    st.markdown("### Which morning services are you available?")
    m1, m2 = st.columns(2)
    with m1:
        if st.button(f"Select all mornings ({kid_name})"):
            for opt in morning_options:
                st.session_state[f"k{idx}_morn_{target_month_key}_{opt}"] = True
    with m2:
        if st.button(f"Clear mornings ({kid_name})"):
            for opt in morning_options:
                st.session_state[f"k{idx}_morn_{target_month_key}_{opt}"] = False

    chosen_m = []
    for opt in morning_options:
        key = f"k{idx}_morn_{target_month_key}_{opt}"
        if st.checkbox(opt, key=key):
            chosen_m.append(opt)

    st.divider()

    # Evening
    st.markdown("### Which evening services are you available?")
    e1, e2 = st.columns(2)
    with e1:
        if st.button(f"Select all evenings ({kid_name})"):
            for opt in evening_options:
                st.session_state[f"k{idx}_eve_{target_month_key}_{opt}"] = True
    with e2:
        if st.button(f"Clear evenings ({kid_name})"):
            for opt in evening_options:
                st.session_state[f"k{idx}_eve_{target_month_key}_{opt}"] = False

    chosen_e = []
    for opt in evening_options:
        key = f"k{idx}_eve_{target_month_key}_{opt}"
        if st.checkbox(opt, key=key):
            chosen_e.append(opt)

    # Convert display -> actual labels
    selected_morning_labels = {morning_display_map[d] for d in chosen_m if d in morning_display_map}
    selected_evening_labels = {evening_display_map[d] for d in chosen_e if d in evening_display_map}
    kids_selected_map[kid_name] = selected_morning_labels.union(selected_evening_labels)

    st.divider()

# Review
st.subheader("Review")
st.write(f"**Family:** {answers.get('FAMILY_SURNAME')}")
for kid_name in kids:
    st.write(f"- **{kid_name}:** {len(kids_selected_map.get(kid_name, set()))} services selected")

# ─────────────────────────────────────────────────────────────
# Submit (sticky)
# ─────────────────────────────────────────────────────────────
st.markdown('<div class="sticky-submit">', unsafe_allow_html=True)
submitted = st.button("Submit")
st.markdown("</div>", unsafe_allow_html=True)

if submitted:
    # Hard deadline check
    now_check = get_now_in_tz(deadline_tz)
    if (deadline_dt - now_check).total_seconds() <= 0:
        st.error("Form is closed.")
        st.stop()

    family = answers.get("FAMILY_SURNAME") or ""
    if not family:
        st.error("Please select your family surname.")
        st.stop()

    # ✅ One timestamp shared for all kids in this submission (like your screenshot)
    now_iso = datetime.utcnow().isoformat() + "Z"

    # ✅ Desired header EXACTLY like your example + service columns
    desired_header = ["timestamp", "Availability month", "Family Surname", "Serving kid"] + date_labels

    try:
        # ✅ Append ONE ROW PER KID
        for kid_name in kids:
            selected = kids_selected_map.get(kid_name, set())

            # Yes/No for each service column (same format as guys)
            yes_map = {lbl: ("Yes" if lbl in selected else "No") for lbl in date_labels}

            row_map = {
                "timestamp": now_iso,
                "Availability month": target_month_key,
                "Family Surname": family,
                "Serving kid": kid_name,
            }
            row_map.update(yes_map)

            append_response_row(desired_header, row_map)

        clear_caches()
        st.success("Submission saved to Google Sheets (one row per child).")

    except Exception as e:
        st.error(f"Failed to save submission: {e}")

# ─────────────────────────────────────────────────────────────
# Admin (optional)
# ─────────────────────────────────────────────────────────────
with st.expander("Admin"):
    st.caption("Mode: Google Sheets (kids tabs)")

    if not ADMIN_KEY:
        st.info("Admin key is not set (ADMIN_KEY). Admin view is open to anyone with the link.")

    key = st.text_input("Enter admin key to access exports", type="password")
    if ADMIN_KEY and key != ADMIN_KEY:
        if key:
            st.error("Incorrect admin key.")
    else:
        st.success("Admin unlocked.")
        try:
            responses_df = fetch_responses_df()
        except Exception as e:
            st.error(f"Could not load responses: {e}")
            responses_df = pd.DataFrame()

        st.write(f"Total submissions (all months): **{len(responses_df)}**")
        if not responses_df.empty:
            st.dataframe(responses_df, use_container_width=True)
