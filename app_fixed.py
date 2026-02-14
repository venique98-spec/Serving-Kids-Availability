# app_kids_adapted.py
import time
import random
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st

# ✅ Timezone-aware deadlines
try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None  # fallback

# Optional: Google Sheets libs. If missing, the app runs in Local CSV mode.
try:
    import gspread
    from gspread.exceptions import APIError, WorksheetNotFound
except Exception:  # still run in local mode if not installed
    gspread = None

    class APIError(Exception):
        ...

    class WorksheetNotFound(Exception):
        ...


# ──────────────────────────────────────────────────────────────────────────────
# UI CONFIG + mobile tweaks
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Kids Availability Form", page_icon="🗓️", layout="centered")
st.title("🗓️ Kids Availability Form")

st.markdown(
    """
<style>
  .stButton > button { width: 100%; height: 48px; font-size: 16px; }
  label[data-baseweb="radio"] { padding: 6px 0; }
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
# Google Sheets (one spreadsheet, tabs per your new structure)
# ──────────────────────────────────────────────────────────────────────────────
TAB_RESPONSES = "Responses"
TAB_KIDSBASE = "Kidsbase"
TAB_DEADLINES = "Deadlines"
TAB_KIDS_DATES = "KidsServingdates"

BASE_TZ_FALLBACK = "Africa/Johannesburg"

# Candidate column names (supports slight variations)
DIRECTOR_COL_CANDIDATES = ["Director", "Director Name", "director"]
KID_COL_CANDIDATES = ["Serving Kid", "Serving child", "Serving Child", "Kid", "Child", "Name", "Serving Girl"]
DATE_TAB_REQUIRED_CANDIDATES = {
    "target_month": ["target_month", "target month", "month", "Target Month", "TargetMonth"],
    "date": ["date", "Date", "service_date", "service date", "Service Date"],
    "label": ["label", "Label", "service_label", "Service Label", "service label"],
    "is_service_day": ["is_service_day", "is service day", "is_service", "Is Service Day", "is_service_day?"],
}


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


SHEETS_MODE = is_sheets_enabled()
if not SHEETS_MODE:
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
    Open the single spreadsheet and return the gspread Spreadsheet object.
    Includes a robust private_key newline fixer to avoid RSA key parse errors.
    """
    sa = _get_secret_any(["gcp_service_account"], ["general", "gcp_service_account"])
    sheet_id = _get_secret_any(["GSHEET_ID"], ["general", "GSHEET_ID"])

    if not sa or not sheet_id:
        raise RuntimeError("Missing GSHEET_ID or gcp_service_account in secrets.")

    sa = dict(sa)  # copy so we can safely modify
    pk = sa.get("private_key", "")
    if isinstance(pk, str):
        pk = pk.replace("\\n", "\n").strip()
        if not pk.endswith("\n"):
            pk += "\n"
        sa["private_key"] = pk

    gc = gspread.service_account_from_dict(sa)
    sh = gs_retry(gc.open_by_key, sheet_id)
    return sh


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
    if not header:
        return pd.DataFrame()
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
def fetch_kidsbase_df() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_KIDSBASE, rows=4000, cols=10)
    return ws_get_df(ws)


@st.cache_data(ttl=30, show_spinner=False)
def fetch_deadlines_df() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_DEADLINES, rows=500, cols=10)
    return ws_get_df(ws)


@st.cache_data(ttl=30, show_spinner=False)
def fetch_kids_service_dates_df() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_KIDS_DATES, rows=4000, cols=10)
    return ws_get_df(ws)


@st.cache_data(ttl=30, show_spinner=False)
def fetch_responses_df() -> pd.DataFrame:
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_RESPONSES, rows=8000, cols=200)
    return ws_get_df(ws)


def append_response_row(desired_header: list[str], row_map: dict):
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_RESPONSES, rows=8000, cols=max(200, len(desired_header) + 10))
    header = ws_ensure_header(ws, desired_header)
    row = [row_map.get(col, "") for col in header]
    gs_retry(ws.append_row, row)


def clear_caches():
    for fn in (fetch_kidsbase_df, fetch_deadlines_df, fetch_kids_service_dates_df, fetch_responses_df):
        try:
            fn.clear()
        except Exception:
            pass
    try:
        st.cache_data.clear()
    except Exception:
        pass


# ──────────────────────────────────────────────────────────────────────────────
# Helpers: column detection / normalization
# ──────────────────────────────────────────────────────────────────────────────
def _norm_col(s: str) -> str:
    return str(s).strip().lower()


def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    if df is None or df.empty:
        return None
    norm_map = {_norm_col(c): c for c in df.columns}
    for cand in candidates:
        k = _norm_col(cand)
        if k in norm_map:
            return norm_map[k]
    return None


def require_col(df: pd.DataFrame, candidates: list[str], friendly_name: str, tab_name: str) -> str:
    col = pick_col(df, candidates)
    if not col:
        st.error(
            f"Tab '{tab_name}' is missing a required column for {friendly_name}. "
            f"Accepted names: {', '.join(candidates)}"
        )
        st.stop()
    return col


def require_date_tab_cols(df: pd.DataFrame, tab_name: str) -> dict:
    out = {}
    for key, cands in DATE_TAB_REQUIRED_CANDIDATES.items():
        col = pick_col(df, cands)
        if not col:
            st.error(
                f"Tab '{tab_name}' is missing required column '{key}'. "
                f"Accepted names: {', '.join(cands)}"
            )
            st.stop()
        out[key] = col
    return out


# ──────────────────────────────────────────────────────────────────────────────
# Time helpers
# ──────────────────────────────────────────────────────────────────────────────
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
    """In Feb -> target is Mar, in Mar -> target is Apr, etc."""
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
    if hrs > 0:
        return f"{hrs}h {rem_m}m"
    return f"{rem_m}m"


def _safe_parse_date_ymd(s: str) -> datetime:
    try:
        return datetime.strptime(str(s).strip(), "%Y-%m-%d")
    except Exception:
        return datetime(1900, 1, 1)


# ──────────────────────────────────────────────────────────────────────────────
# Reporting
# ──────────────────────────────────────────────────────────────────────────────
def build_human_report(
    target_month_key: str,
    director: str,
    kid_name: str,
    date_labels: list[str],
    answers: dict,
) -> str:
    lines = [
        f"Availability month: {target_month_key}",
        f"Director: {director or '—'}",
        f"Serving Kid: {kid_name or '—'}",
        "Availability:",
    ]
    for lbl in date_labels:
        val = (answers.get(lbl) or "No").title()
        lines.append(f"{lbl}: {val}")
    return "\n".join(lines)


# ─────────────────────────────────────────────────────────────
# Load config from Google Sheets (new tabs)
# ─────────────────────────────────────────────────────────────
try:
    kidsbase_df = fetch_kidsbase_df()
    deadlines_df = fetch_deadlines_df()
    kids_dates_df = fetch_kids_service_dates_df()
except Exception as e:
    st.error(f"Failed to load config from Google Sheets: {e}")
    st.stop()

# Detect needed columns
director_col = require_col(kidsbase_df, DIRECTOR_COL_CANDIDATES, "Director", TAB_KIDSBASE)
kid_col = require_col(kidsbase_df, KID_COL_CANDIDATES, "Serving Kid name", TAB_KIDSBASE)

# Deadlines required columns
for col in ["month", "deadline_local", "timezone"]:
    if col not in deadlines_df.columns:
        st.error(
            f"Tab '{TAB_DEADLINES}' is missing columns: {', '.join([c for c in ['month','deadline_local','timezone'] if c not in deadlines_df.columns])}"
        )
        st.stop()

date_cols = require_date_tab_cols(kids_dates_df, TAB_KIDS_DATES)

# Clean kids base
kidsbase_df[director_col] = kidsbase_df[director_col].astype(str).str.strip()
kidsbase_df[kid_col] = kidsbase_df[kid_col].astype(str).str.strip()
kidsbase_df = kidsbase_df[(kidsbase_df[director_col] != "") & (kidsbase_df[kid_col] != "")].drop_duplicates()

# Clean deadlines
deadlines_df["month"] = deadlines_df["month"].astype(str).str.strip()
deadlines_df["deadline_local"] = deadlines_df["deadline_local"].astype(str).str.strip()
deadlines_df["timezone"] = deadlines_df["timezone"].astype(str).str.strip()

# Clean kids dates
kids_dates_df[date_cols["target_month"]] = kids_dates_df[date_cols["target_month"]].astype(str).str.strip()
kids_dates_df[date_cols["date"]] = kids_dates_df[date_cols["date"]].astype(str).str.strip()
kids_dates_df[date_cols["label"]] = kids_dates_df[date_cols["label"]].astype(str).str.strip()
kids_dates_df[date_cols["is_service_day"]] = kids_dates_df[date_cols["is_service_day"]].astype(str).str.strip()

# Build director->kids map
serving_map = (
    kidsbase_df.groupby(director_col)[kid_col]
    .apply(lambda s: sorted({x for x in s if x}))
    .to_dict()
)

# Base timezone (prefer first row timezone)
BASE_TZ = BASE_TZ_FALLBACK
try:
    tz0 = str(deadlines_df["timezone"].iloc[0]).strip()
    if tz0:
        BASE_TZ = tz0
except Exception:
    pass

now_base = get_now_in_tz(BASE_TZ)
target_month_key = get_target_month_key(now_base)

# Filter service dates for target month (KidsServingdates)
month_dates = kids_dates_df[
    (kids_dates_df[date_cols["target_month"]] == target_month_key)
    & (kids_dates_df[date_cols["is_service_day"]] == "1")
].copy()

if month_dates.empty:
    st.markdown(
        f"""
        ## 🔒 This month’s availability form is not open yet.

        No service dates were found for **{target_month_key}**.

        Please contact your director.
        """
    )
    st.stop()

month_dates["_sort"] = month_dates[date_cols["date"]].map(_safe_parse_date_ymd)
month_dates = month_dates.sort_values("_sort").drop(columns=["_sort"])

date_labels = month_dates[date_cols["label"]].astype(str).tolist()

# Deadline for target month
def get_deadline_for_target_month(deadlines: pd.DataFrame, month_key: str):
    tz_guess = BASE_TZ
    match = deadlines[deadlines["month"] == month_key]
    if match.empty:
        return None, tz_guess
    row = match.iloc[0]
    tz_name = str(row["timezone"]).strip() or tz_guess
    dl = parse_deadline_local(str(row["deadline_local"]).strip(), tz_name)
    return dl, tz_name


deadline_dt, deadline_tz = get_deadline_for_target_month(deadlines_df, target_month_key)

# Closed if missing deadline or past deadline
is_closed = True
if deadline_dt is not None:
    now_local = get_now_in_tz(deadline_tz)
    is_closed = (deadline_dt - now_local).total_seconds() <= 0

if is_closed:
    target_month_dt = datetime.strptime(target_month_key, "%Y-%m")
    target_month_name = target_month_dt.strftime("%B")
    next_cycle_dt = add_one_month(target_month_dt)
    next_cycle_name = next_cycle_dt.strftime("%B")
    open_date_str = target_month_dt.strftime("1 %B")

    st.markdown(
        f"""
        ## 🔒 {target_month_name} availability submissions are now closed.

        If you have not submitted your dates, please contact your director.

        <hr style="margin:30px 0;">

        <div style="
            color: #1e7e34;
            font-size: 24px;
            font-weight: 600;
        ">
            {next_cycle_name} availability submissions will open on {open_date_str}.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

# Countdown + policy note (no minimum requirement for kids)
now_local = get_now_in_tz(deadline_tz)
remaining_seconds = (deadline_dt - now_local).total_seconds()
st.markdown('<meta http-equiv="refresh" content="60">', unsafe_allow_html=True)

st.info(
    f"🗓️ Submitting availability for **{target_month_key}**.\n\n"
    f"⏳ Form closes at **{deadline_dt.strftime('%Y-%m-%d %H:%M')}** ({deadline_tz}). "
    f"Time remaining: **{format_minutes_remaining(remaining_seconds)}**\n\n"
    f"🔁 You are welcome to submit this form more than once. "
    f"We will use your most recent submission for scheduling. "
    f"Please remember to send a screenshot of your final submission to your director."
)

# ─────────────────────────────────────────────────────────────
# UI state
# ─────────────────────────────────────────────────────────────
if "answers" not in st.session_state:
    st.session_state.answers = {}
answers = st.session_state.answers

# ─────────────────────────────────────────────────────────────
# Form UI
# ─────────────────────────────────────────────────────────────
st.subheader("Your details")
directors = sorted([d for d in serving_map.keys() if d])
answers["Q1"] = st.selectbox("Please select your director’s name", options=[""] + directors, index=0)

if answers.get("Q1"):
    kids = serving_map.get(answers["Q1"], [])
    answers["Q2"] = st.selectbox("Please select your name", options=[""] + kids, index=0)
else:
    answers["Q2"] = ""

st.subheader(f"Availability for {target_month_key}")

radio_options = ["Yes", "No"]
for lbl in date_labels:
    saved = answers.get(lbl)
    idx = radio_options.index(saved) if saved in radio_options else None
    choice = st.radio(
        f"Are you available {lbl}?",
        options=radio_options,
        index=idx,
        key=f"avail_{target_month_key}_{lbl}",
        horizontal=False,
    )
    answers[lbl] = choice

# Review
st.subheader("Review")
c1, c2 = st.columns(2)
with c1:
    st.metric("Director", answers.get("Q1") or "—")
with c2:
    st.metric("Name", answers.get("Q2") or "—")

# ─────────────────────────────────────────────────────────────
# Submit (sticky)
# ─────────────────────────────────────────────────────────────
errors = {}
st.markdown('<div class="sticky-submit">', unsafe_allow_html=True)
submitted = st.button("Submit")
st.markdown("</div>", unsafe_allow_html=True)

if submitted:
    # Hard deadline check on submit too
    now_check = get_now_in_tz(deadline_tz)
    if (deadline_dt - now_check).total_seconds() <= 0:
        target_month_dt = datetime.strptime(target_month_key, "%Y-%m")
        target_month_name = target_month_dt.strftime("%B")
        next_cycle_dt = add_one_month(target_month_dt)
        next_cycle_name = next_cycle_dt.strftime("%B")
        open_date_str = target_month_dt.strftime("1 %B")

        st.markdown(
            f"""
            ## 🔒 {target_month_name} availability submissions are now closed.

            If you have not submitted your dates, please contact your director.

            <hr style="margin:30px 0;">

            <div style="
                color: #1e7e34;
                font-size: 24px;
                font-weight: 600;
            ">
                {next_cycle_name} availability submissions will open on {open_date_str}.
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.stop()

    if not answers.get("Q1"):
        errors["Q1"] = "Please select a director."
    if not answers.get("Q2"):
        errors["Q2"] = "Please select your name."

    if errors:
        for msg in errors.values():
            st.error(msg)
    else:
        now = datetime.utcnow().isoformat() + "Z"

        # Keep output columns consistent (nice for exports)
        OUT_DIRECTOR_COL = "Director"
        OUT_KID_COL = "Serving Kid"

        row_map = {
            "timestamp": now,
            "Availability month": target_month_key,
            OUT_DIRECTOR_COL: answers.get("Q1") or "",
            OUT_KID_COL: answers.get("Q2") or "",
        }
        for lbl in date_labels:
            row_map[lbl] = (answers.get(lbl) or "No").title()

        desired_header = ["timestamp", "Availability month", OUT_DIRECTOR_COL, OUT_KID_COL] + date_labels

        try:
            append_response_row(desired_header, row_map)
            clear_caches()
            st.success("Submission saved to Google Sheets.")
        except Exception as e:
            st.error(f"Failed to save submission: {e}")

        report_text = build_human_report(
            target_month_key=target_month_key,
            director=answers.get("Q1") or "",
            kid_name=answers.get("Q2") or "",
            date_labels=date_labels,
            answers=answers,
        )
        st.markdown("### 📄 Screenshot-friendly report (text)")
        st.code(report_text, language=None)
        st.download_button(
            "Download report as .txt",
            data=report_text.encode("utf-8"),
            file_name=f"Availability_{target_month_key}_{(answers.get('Q2') or 'name').replace(' ', '_')}.txt",
            mime="text/plain",
        )

# ─────────────────────────────────────────────────────────────
# Admin: exports + non-responders + diagnostics
# ─────────────────────────────────────────────────────────────
def compute_nonresponders(kidsbase: pd.DataFrame, responses: pd.DataFrame) -> pd.DataFrame:
    # Canonical output columns
    OUT_DIRECTOR_COL = "Director"
    OUT_KID_COL = "Serving Kid"

    if kidsbase is None or kidsbase.empty:
        return pd.DataFrame(columns=[OUT_DIRECTOR_COL, OUT_KID_COL])

    sb = kidsbase[[director_col, kid_col]].copy()
    sb.rename(columns={director_col: OUT_DIRECTOR_COL, kid_col: OUT_KID_COL}, inplace=True)
    sb[OUT_DIRECTOR_COL] = sb[OUT_DIRECTOR_COL].astype(str).str.strip()
    sb[OUT_KID_COL] = sb[OUT_KID_COL].astype(str).str.strip()
    sb = sb[(sb[OUT_DIRECTOR_COL] != "") & (sb[OUT_KID_COL] != "")].drop_duplicates()

    if responses is None or responses.empty:
        out = sb.copy()
        out["Responded"] = False
        out["Last submission"] = ""
        return out

    # Detect response columns (support older headers too)
    r_dir = pick_col(responses, ["Director", OUT_DIRECTOR_COL, "director"])
    r_kid = pick_col(responses, ["Serving Kid", OUT_KID_COL, "Serving Girl", "Kid", "Name", "Child"])
    if not r_dir or not r_kid:
        # If responses exist but headers are unusual, show none as responded (safer)
        out = sb.copy()
        out["Responded"] = False
        out["Last submission"] = ""
        return out

    cols = list(responses.columns)
    ts_col = "timestamp" if "timestamp" in cols else (cols[0] if cols else "timestamp")

    use_cols = [c for c in [r_dir, r_kid, ts_col] if c in responses.columns]
    resp = responses[use_cols].copy()
    if ts_col not in resp.columns:
        resp[ts_col] = ""
    resp.rename(columns={r_dir: OUT_DIRECTOR_COL, r_kid: OUT_KID_COL, ts_col: "Last submission"}, inplace=True)

    resp[OUT_DIRECTOR_COL] = resp[OUT_DIRECTOR_COL].astype(str).str.strip()
    resp[OUT_KID_COL] = resp[OUT_KID_COL].astype(str).str.strip()
    resp = resp.sort_values("Last submission").drop_duplicates(subset=[OUT_DIRECTOR_COL, OUT_KID_COL], keep="last")

    merged = sb.merge(resp, on=[OUT_DIRECTOR_COL, OUT_KID_COL], how="left")
    merged["Responded"] = merged["Last submission"].notna() & (merged["Last submission"] != "")
    nonresp = merged[~merged["Responded"]].copy()
    return nonresp.sort_values([OUT_DIRECTOR_COL, OUT_KID_COL]).reset_index(drop=True)


with st.expander("Admin"):
    st.caption(f"Mode: Google Sheets | Tabs: {TAB_RESPONSES}, {TAB_KIDSBASE}, {TAB_DEADLINES}, {TAB_KIDS_DATES}")
    if not ADMIN_KEY:
        st.info("To protect exports, set an ADMIN_KEY in Streamlit Secrets (optional).")

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

        st.write(f"Total submissions: **{len(responses_df)}**")
        if not responses_df.empty:
            st.dataframe(responses_df, use_container_width=True)
            try:
                import openpyxl  # noqa
                out = BytesIO()
                with pd.ExcelWriter(out, engine="openpyxl") as xw:
                    responses_df.to_excel(xw, index=False, sheet_name="Responses")
                st.download_button(
                    "Download all responses",
                    data=out.getvalue(),
                    file_name="Kids_availability_responses.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception:
                st.download_button(
                    "Download all responses",
                    data=responses_df.to_csv(index=False).encode("utf-8"),
                    file_name="Kids_availability_responses.csv",
                    mime="text/csv",
                )
        else:
            st.warning("No submissions yet.")

        st.markdown("### ❌ Non-responders")
        nonresp_df = compute_nonresponders(kidsbase_df, responses_df)

        all_directors = ["All"] + sorted(kidsbase_df[director_col].unique().tolist())
        sel_dir = st.selectbox("Filter by director", options=all_directors, index=0)
        view_df = nonresp_df if sel_dir == "All" else nonresp_df[nonresp_df["Director"] == sel_dir]
        total_expected = len(kidsbase_df[[director_col, kid_col]].dropna().drop_duplicates())
        st.write(f"Non-responders shown: **{len(view_df)}**  |  Total expected pairs: **{total_expected}**")
        st.dataframe(view_df[["Director", "Serving Kid"]], use_container_width=True)

        st.divider()
        st.markdown("#### 🔍 Secrets / Sheets check")
        try:
            s = st.secrets
            gsa = s.get("gcp_service_account", {})
            gs_id = s.get("GSHEET_ID") or s.get("general", {}).get("GSHEET_ID")
            st.write(
                {
                    "SHEETS_MODE": SHEETS_MODE,
                    "GSHEET_ID_present": bool(gs_id),
                    "client_email": gsa.get("client_email", "(missing)"),
                    "private_key_present": bool(gsa.get("private_key")),
                    "private_key_starts_with": (gsa.get("private_key", "")[:30] if gsa else ""),
                    "gspread_installed": gspread is not None,
                    "tabs_expected": [TAB_RESPONSES, TAB_KIDSBASE, TAB_DEADLINES, TAB_KIDS_DATES],
                    "kidsbase_detected_columns": {"director_col": director_col, "kid_col": kid_col},
                    "kidsservingdates_detected_columns": date_cols,
                }
            )
            sh = get_spreadsheet()
            ensure_worksheet(sh, TAB_RESPONSES, rows=8000, cols=200)
            ensure_worksheet(sh, TAB_KIDSBASE, rows=4000, cols=10)
            ensure_worksheet(sh, TAB_DEADLINES, rows=500, cols=10)
            ensure_worksheet(sh, TAB_KIDS_DATES, rows=4000, cols=10)
            st.success(f"✅ Auth OK. Opened sheet: {sh.title}")
        except Exception as e:
            st.error(f"❌ Diagnostics failed: {e}")
