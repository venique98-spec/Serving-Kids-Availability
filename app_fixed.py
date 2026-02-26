# app_fixed.py
# uKids Kids Availability Form
# ✅ SB input supports: Family Surname | Kid #1 | Age | Kid #2 | Age | Kid #3 | Age ...
# ✅ Age is NOT shown on the form
# ✅ Age IS written to output (one row per child) in "uKids Kids responses"
# ✅ Deadlines pulled from "Kids Deadlines"

import time
import random
from datetime import datetime

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
        ...

    class WorksheetNotFound(Exception):
        ...


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

TAB_RESPONSES = "uKids Kids responses"
TAB_SB = "uKids Kids SB"
TAB_DEADLINES = "Kids Deadlines"
TAB_DATES = "Kids & Guys ServiceDates"


def _get_secret_any(*paths):
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


def is_sheets_enabled() -> bool:
    if gspread is None:
        return False
    sa = _get_secret_any(["gcp_service_account"], ["general", "gcp_service_account"])
    sid = _get_secret_any(["GSHEET_ID"], ["general", "GSHEET_ID"])
    return bool(sa and sid)


if not is_sheets_enabled():
    st.error("Google Sheets is not configured. Add GSHEET_ID and [gcp_service_account] to Secrets.")
    st.stop()


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


def make_unique_headers(header: list[str]) -> list[str]:
    counts = {}
    out = []
    for h in header:
        base = str(h).strip() or "Unnamed"
        counts[base] = counts.get(base, 0) + 1
        out.append(base if counts[base] == 1 else f"{base}_{counts[base]}")
    return out


def ws_get_values_and_df(ws) -> tuple[list[str], list[str], pd.DataFrame]:
    values = gs_retry(ws.get_all_values)
    if not values:
        return [], [], pd.DataFrame()
    header_raw = values[0]
    header_unique = make_unique_headers(header_raw)
    rows = values[1:]
    df = pd.DataFrame(rows, columns=header_unique)
    return header_raw, header_unique, df


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
def fetch_sb_raw_unique_df():
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_SB, rows=4000, cols=50)
    return ws_get_values_and_df(ws)


@st.cache_data(ttl=30, show_spinner=False)
def fetch_deadlines_df() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_DEADLINES, rows=500, cols=10)
    _, _, df = ws_get_values_and_df(ws)
    return df


@st.cache_data(ttl=30, show_spinner=False)
def fetch_service_dates_df() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_DATES, rows=4000, cols=20)
    _, _, df = ws_get_values_and_df(ws)
    return df


def append_response_row(desired_header: list[str], row_map: dict):
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_RESPONSES, rows=8000, cols=max(400, len(desired_header) + 10))
    header = ws_ensure_header(ws, desired_header)
    row = [row_map.get(col, "") for col in header]
    gs_retry(ws.append_row, row)


def clear_caches():
    for fn in (fetch_sb_raw_unique_df, fetch_deadlines_df, fetch_service_dates_df):
        try:
            fn.clear()
        except Exception:
            pass
    try:
        st.cache_data.clear()
    except Exception:
        pass


def get_now_in_tz(tz_name: str) -> datetime:
    if ZoneInfo is None:
        return datetime.utcnow()
    return datetime.now(ZoneInfo(tz_name))


def add_one_month(dt: datetime) -> datetime:
    y, m = dt.year, dt.month
    y2, m2 = (y + 1, 1) if m == 12 else (y, m + 1)
    return datetime(y2, m2, 1, tzinfo=dt.tzinfo) if dt.tzinfo else datetime(y2, m2, 1)


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
    return str(v).strip().lower() in ("1", "true", "yes", "y", "t")


def _display_date_only(label: str) -> str:
    s = str(label).strip()
    s = s.replace("Morning Service", "").replace("Evening Service", "")
    s = s.replace("Morning", "").replace("Evening", "").replace("Service", "")
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
    sb_header_raw, sb_header_unique, sb_df = fetch_sb_raw_unique_df()
    deadlines_df = fetch_deadlines_df()
    service_dates_all = fetch_service_dates_df()
except Exception as e:
    st.error(f"Failed to load config from Google Sheets: {e}")
    st.stop()

# Required columns
for needed, tab, df in [
    ({"month", "deadline_local", "timezone"}, TAB_DEADLINES, deadlines_df),
    ({"target_month", "date", "label", "is_service_day"}, TAB_DATES, service_dates_all),
]:
    miss = needed - set(df.columns)
    if miss:
        st.error(f"Google Sheet tab '{tab}' is missing columns: {', '.join(sorted(miss))}")
        st.stop()

if "Family Surname" not in sb_df.columns:
    st.error(f"Google Sheet tab '{TAB_SB}' must include column 'Family Surname'.")
    st.stop()


def build_kid_age_pairs(header_raw: list[str], header_unique: list[str]) -> list[tuple[str, str | None]]:
    pairs = []
    for i, h in enumerate(header_raw):
        hs = str(h).strip().lower()
        if hs.startswith("kid #"):
            kid_col = header_unique[i]
            age_col = None
            if i + 1 < len(header_raw) and str(header_raw[i + 1]).strip().lower() == "age":
                age_col = header_unique[i + 1]
            pairs.append((kid_col, age_col))
    return pairs


kid_age_pairs = build_kid_age_pairs(sb_header_raw, sb_header_unique)
if not kid_age_pairs:
    st.error(f"'{TAB_SB}' must include Kid columns like 'Kid #1' (and optional 'Age' next to each kid).")
    st.stop()

# Clean
deadlines_df["month"] = deadlines_df["month"].astype(str).str.strip()
deadlines_df["deadline_local"] = deadlines_df["deadline_local"].astype(str).str.strip()
deadlines_df["timezone"] = deadlines_df["timezone"].astype(str).str.strip()

service_dates_all["target_month"] = service_dates_all["target_month"].astype(str).str.strip()
service_dates_all["date"] = service_dates_all["date"].astype(str).str.strip()
service_dates_all["label"] = service_dates_all["label"].astype(str).str.strip()
service_dates_all["is_service_day"] = service_dates_all["is_service_day"].astype(str).str.strip()

BASE_TZ = "Africa/Johannesburg"
try:
    tz0 = str(deadlines_df["timezone"].iloc[0]).strip()
    if tz0:
        BASE_TZ = tz0
except Exception:
    pass

now_base = get_now_in_tz(BASE_TZ)
target_month_key = get_target_month_key(now_base)

month_dates = service_dates_all[
    (service_dates_all["target_month"] == target_month_key)
    & (service_dates_all["is_service_day"].map(_is_truthy_service_day))
].copy()

if month_dates.empty:
    st.markdown(f"## 🔒 This month’s availability form is not open yet.\n\nNo service dates for **{target_month_key}**.")
    st.stop()

month_dates["_sort"] = month_dates["date"].map(_safe_parse_date_ymd)
month_dates = month_dates.sort_values("_sort").drop(columns=["_sort"])

date_labels = month_dates["label"].astype(str).tolist()
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
if deadline_dt is None:
    st.error(f"No deadline found for month {target_month_key} in '{TAB_DEADLINES}'.")
    st.stop()

now_local = get_now_in_tz(deadline_tz)
if (deadline_dt - now_local).total_seconds() <= 0:
    target_month_name = datetime.strptime(target_month_key, "%Y-%m").strftime("%B")
    st.markdown(f"## 🔒 {target_month_name} availability submissions are now closed.")
    st.stop()

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
# Family -> kids (slot-safe)
# ─────────────────────────────────────────────────────────────
sb = sb_df.copy()
sb["Family Surname"] = sb["Family Surname"].astype(str).str.strip()
families = sorted({f for f in sb["Family Surname"].tolist() if f and f.lower() != "nan"})


def _clean_cell(x) -> str:
    s = str(x).strip()
    return "" if (s == "" or s.lower() == "nan") else s


def get_kids_info_for_family(family: str) -> list[dict]:
    row = sb[sb["Family Surname"] == family]
    if row.empty:
        return []
    r0 = row.iloc[0]
    out = []
    for slot_idx, (kid_col, age_col) in enumerate(kid_age_pairs, start=1):
        kid_name = _clean_cell(r0.get(kid_col, ""))
        if not kid_name:
            continue
        age_val = _clean_cell(r0.get(age_col, "")) if age_col else ""
        out.append({"slot": slot_idx, "name": kid_name, "age": age_val})
    return out


# ─────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────
st.subheader("Your details")
family = st.selectbox("Family surname", options=[""] + families, index=0)
if not family:
    st.caption("Select your family surname to continue.")
    st.stop()

kids_info = get_kids_info_for_family(family)
if not kids_info:
    st.warning("No kids found for this family in 'uKids Kids SB'. Please fill in Kid #1 (and Age if you want).")
    st.stop()

st.subheader(f"Availability for {target_month_key}")

kids_selected_map: dict[int, set[str]] = {}

for k in kids_info:
    slot = k["slot"]
    kid_name = k["name"]

    st.markdown(f"## {kid_name}")
    st.caption("Select all services this child is available for.")

    st.markdown("### Which morning services are you available?")
    m1, m2 = st.columns(2)
    with m1:
        if st.button(f"Select all mornings ({kid_name})"):
            for opt in morning_options:
                st.session_state[f"slot{slot}_morn_{target_month_key}_{opt}"] = True
    with m2:
        if st.button(f"Clear mornings ({kid_name})"):
            for opt in morning_options:
                st.session_state[f"slot{slot}_morn_{target_month_key}_{opt}"] = False

    chosen_m = []
    for opt in morning_options:
        key = f"slot{slot}_morn_{target_month_key}_{opt}"
        if st.checkbox(opt, key=key):
            chosen_m.append(opt)

    st.divider()

    st.markdown("### Which evening services are you available?")
    e1, e2 = st.columns(2)
    with e1:
        if st.button(f"Select all evenings ({kid_name})"):
            for opt in evening_options:
                st.session_state[f"slot{slot}_eve_{target_month_key}_{opt}"] = True
    with e2:
        if st.button(f"Clear evenings ({kid_name})"):
            for opt in evening_options:
                st.session_state[f"slot{slot}_eve_{target_month_key}_{opt}"] = False

    chosen_e = []
    for opt in evening_options:
        key = f"slot{slot}_eve_{target_month_key}_{opt}"
        if st.checkbox(opt, key=key):
            chosen_e.append(opt)

    selected_morning_labels = {morning_display_map[d] for d in chosen_m if d in morning_display_map}
    selected_evening_labels = {evening_display_map[d] for d in chosen_e if d in evening_display_map}
    kids_selected_map[slot] = selected_morning_labels.union(selected_evening_labels)

    st.divider()

st.subheader("Review")
st.write(f"**Family:** {family}")
for k in kids_info:
    slot = k["slot"]
    st.write(f"- **{k['name']}:** {len(kids_selected_map.get(slot, set()))} services selected")

st.markdown('<div class="sticky-submit">', unsafe_allow_html=True)
submitted = st.button("Submit")
st.markdown("</div>", unsafe_allow_html=True)

if submitted:
    now_check = get_now_in_tz(deadline_tz)
    if (deadline_dt - now_check).total_seconds() <= 0:
        st.error("Form is closed.")
        st.stop()

    now_iso = datetime.utcnow().isoformat() + "Z"

    # ✅ EXACT ORDER like your screenshot
    desired_header = ["timestamp", "Availability month", "Family Surname", "Serving kid", "Age"] + date_labels

    try:
        for k in kids_info:
            slot = k["slot"]
            kid_name = k["name"]
            age_val = k.get("age", "")

            selected = kids_selected_map.get(slot, set())
            yes_map = {lbl: ("Yes" if lbl in selected else "No") for lbl in date_labels}

            row_map = {
                "timestamp": now_iso,
                "Availability month": target_month_key,
                "Family Surname": family,
                "Serving kid": kid_name,
                "Age": age_val,   # ✅ output only, not shown on form
            }
            row_map.update(yes_map)

            append_response_row(desired_header, row_map)

        clear_caches()
        st.success("Submission saved to Google Sheets.")

    except Exception as e:
        st.error(f"Failed to save submission: {e}")
