# app_fixed.py
# Weekly uKids Kids Availability Form
#
# Registration flow:
# - Family registers all their children at once (up to 5 kids per submission)
# - Family is either brand-new (auto Family Code) or linked via an existing sibling
# - Duplicate check per child name; duplicates blocked with a "contact Ps Some" message
# - Age stored as plain number; displayed as "Age 11 (Born 2015)" in the UI
#
# Availability flow:
# - Parent selects one child -> app loads all siblings via Family Code
# - Parent selects available services per child and submits
#
# SHEET TABS REQUIRED:
# "uKids Kids SB"          : Family Code | Name & Surname | Age
# "Kids Deadlines"         : service_date | Opening date | deadline_local | timezone
# "Kids & Guys ServiceDates": service_date | date | label | is_service_day
# "uKids Kids responses"   : timestamp | Service date | Family Code | Name | Age | <service cols>
# "Final Kids serving dates": timestamp | Service date | Service | Name | Age

import re
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
        pass

    class WorksheetNotFound(Exception):
        pass


# ─────────────────────────────────────────────────────────────
# Page config & CSS
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="uKids Availability", layout="centered")

st.markdown(
    """
<style>
  @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800&display=swap');
  html, body, [class*="css"] { font-family: 'Nunito', sans-serif; }

  :root {
    --green:    #2D6A4F;
    --green-lt: #52B788;
    --green-bg: #D8F3DC;
    --amber:    #F4A261;
    --text:     #1B1B1B;
    --muted:    #6B7280;
    --card:     #FFFFFF;
    --bg:       #F0FAF4;
  }

  .stApp { background: var(--bg); }

  .ukids-hero {
    background: linear-gradient(135deg, var(--green) 0%, var(--green-lt) 100%);
    border-radius: 16px;
    padding: 28px 24px 22px;
    margin-bottom: 24px;
    color: #fff;
  }
  .ukids-hero h1 { font-size: 1.9rem; font-weight: 800; margin: 0 0 4px; color: #fff; }
  .ukids-hero p  { font-size: 0.95rem; margin: 0; opacity: 0.88; }

  .ukids-card {
    background: var(--card);
    border-radius: 12px;
    padding: 20px 20px 16px;
    margin-bottom: 18px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
    border-left: 4px solid var(--green-lt);
  }
  .ukids-card-amber { border-left-color: var(--amber); }
  .ukids-card h3 { font-size: 1.05rem; font-weight: 700; color: var(--green); margin: 0 0 10px; }

  .kid-name { font-size: 1.05rem; font-weight: 700; color: var(--text); }
  .kid-age  { font-size: 0.82rem; color: var(--muted); margin-bottom: 6px; }

  .info-pill {
    background: var(--green-bg);
    border-radius: 8px; padding: 12px 16px;
    font-size: 0.88rem; color: var(--green);
    margin-bottom: 14px; line-height: 1.7;
  }

  .sticky-submit {
    position: sticky; bottom: 0; z-index: 999;
    background: var(--bg);
    padding: 10px 0 4px;
    border-top: 1px solid #d1fae5;
  }

  .stButton > button {
    width: 100%; height: 48px; font-size: 1rem;
    font-weight: 700; border-radius: 10px;
    background: var(--green) !important;
    color: #fff !important; border: none !important;
    transition: background 0.2s;
  }
  .stButton > button:hover { background: var(--green-lt) !important; }

  .child-block {
    background: #f8fdfb;
    border: 1px solid #d1fae5;
    border-radius: 10px;
    padding: 16px 16px 10px;
    margin-bottom: 14px;
  }
  .child-block-label {
    font-size: 0.78rem; font-weight: 700;
    color: var(--green); text-transform: uppercase;
    letter-spacing: 0.05em; margin-bottom: 8px;
  }

  @media (max-width: 520px) {
    div[data-testid="column"] { width: 100% !important; flex: 0 0 100% !important; }
    .ukids-hero h1 { font-size: 1.5rem; }
  }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="ukids-hero">
  <h1>uKids Availability</h1>
  <p>Let us know which services your children can serve at this week.</p>
</div>
""",
    unsafe_allow_html=True,
)

TAB_RESPONSES = "uKids Kids responses"
TAB_SB        = "uKids Kids SB"
TAB_DEADLINES = "Kids Deadlines"
TAB_DATES     = "Kids & Guys ServiceDates"
TAB_FINAL     = "Final Kids serving dates"

MAX_CHILDREN_PER_REGISTRATION = 5


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
# Google Sheets helpers
# ─────────────────────────────────────────────────────────────
def gs_retry(func, *args, **kwargs):
    for attempt in range(5):
        try:
            return func(*args, **kwargs)
        except APIError as e:
            status = getattr(getattr(e, "response", None), "status_code", None)
            if status in (429, 500, 502, 503):
                time.sleep(min(10, (2 ** attempt) + random.random()))
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


def append_multiple_rows(sheet_title: str, desired_header: list, row_maps: list):
    sh     = get_spreadsheet()
    ws     = ensure_worksheet(sh, sheet_title, rows=12000, cols=max(100, len(desired_header) + 10))
    header = ws_ensure_header(ws, desired_header)
    values = [[row.get(col, "") for col in header] for row in row_maps]
    if values:
        gs_retry(ws.append_rows, values)


@st.cache_data(ttl=300, show_spinner=False)
def fetch_sb_df():
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_SB, rows=4000, cols=20)
    _, _, df = ws_get_values_and_df(ws)
    return df


@st.cache_data(ttl=300, show_spinner=False)
def fetch_deadlines_df():
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_DEADLINES, rows=500, cols=10)
    _, _, df = ws_get_values_and_df(ws)
    return df


@st.cache_data(ttl=300, show_spinner=False)
def fetch_service_dates_df():
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_DATES, rows=4000, cols=20)
    _, _, df = ws_get_values_and_df(ws)
    return df


@st.cache_data(ttl=300, show_spinner=False)
def fetch_responses_df():
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, TAB_RESPONSES, rows=12000, cols=300)
    _, _, df = ws_get_values_and_df(ws)
    return df


# ─────────────────────────────────────────────────────────────
# Pure helpers
# ─────────────────────────────────────────────────────────────
def get_now_in_tz(tz_name: str) -> datetime:
    if ZoneInfo is None:
        return datetime.utcnow()
    return datetime.now(ZoneInfo(tz_name))


def parse_local_dt(value: str, tz_name: str) -> datetime:
    dt_naive = datetime.strptime(str(value).strip(), "%Y-%m-%d %H:%M")
    if ZoneInfo is None:
        return dt_naive
    return dt_naive.replace(tzinfo=ZoneInfo(tz_name))


def format_time_remaining(delta_seconds: float) -> str:
    mins  = max(0, int(delta_seconds // 60))
    hrs   = mins // 60
    rem_m = mins % 60
    return f"{hrs}h {rem_m}m" if hrs > 0 else f"{rem_m}m"


def _safe_parse_date_ymd(s: str) -> datetime:
    try:
        return datetime.strptime(str(s).strip(), "%Y-%m-%d")
    except Exception:
        return datetime(1900, 1, 1)


def _is_truthy(v) -> bool:
    return str(v).strip().lower() in ("1", "true", "yes", "y", "t")


def _clean_cell(x) -> str:
    s = str(x).strip()
    return "" if (s == "" or s.lower() == "nan") else s


def age_display(age_raw: str) -> str:
    """'11' -> 'Age 11 (Born 2015)'. Already-formatted strings pass through."""
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
    """'Age 11 (Born 2015)' -> '11'. Plain numbers pass through."""
    age_str = _clean_cell(age_raw)
    if not age_str:
        return ""
    if age_str.isdigit():
        return age_str
    m = re.search(r"\d+", age_str)
    return m.group() if m else age_str


def get_currently_open_week(deadlines: pd.DataFrame, base_tz: str):
    candidates = []
    for _, row in deadlines.iterrows():
        service_date = str(row.get("service_date", "")).strip()
        tz_name      = str(row.get("timezone", "")).strip() or base_tz
        opening_raw  = str(row.get("Opening date", "")).strip()
        deadline_raw = str(row.get("deadline_local", "")).strip()
        if not service_date or not opening_raw or not deadline_raw:
            continue
        try:
            opening_dt  = parse_local_dt(opening_raw, tz_name)
            deadline_dt = parse_local_dt(deadline_raw, tz_name)
            now_local   = get_now_in_tz(tz_name)
        except Exception:
            continue
        if opening_dt <= now_local < deadline_dt:
            candidates.append((service_date, opening_dt, deadline_dt, tz_name))
    if not candidates:
        return None, None, None, base_tz
    candidates.sort(key=lambda x: x[2])
    return candidates[0]


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


def register_children_batch(children: list, family_code: str) -> int:
    """
    Writes all children in `children` (list of dicts with name/age) to uKids Kids SB.
    Skips any whose name is already in the sheet (checked fresh).
    Returns the count of rows actually written.
    """
    sh     = get_spreadsheet()
    ws     = ensure_worksheet(sh, TAB_SB, rows=4000, cols=20)
    ws_ensure_header(ws, ["Family Code", "Name & Surname", "Age"])

    # Fresh read to get the latest names for duplicate detection
    _, _, latest_df = ws_get_values_and_df(ws)
    if "Name & Surname" not in latest_df.columns:
        existing_names = set()
    else:
        existing_names = set(latest_df["Name & Surname"].str.strip().str.lower().tolist())

    rows_to_add   = []
    skipped_names = []
    for child in children:
        name_clean = child["name"].strip()
        if name_clean.lower() in existing_names:
            skipped_names.append(name_clean)
            continue
        rows_to_add.append([family_code, name_clean, child["age"]])
        existing_names.add(name_clean.lower())  # prevent within-batch duplicates

    if rows_to_add:
        gs_retry(ws.append_rows, rows_to_add)

    fetch_sb_df.clear()
    return len(rows_to_add), skipped_names


# ─────────────────────────────────────────────────────────────
# Admin panel
# ─────────────────────────────────────────────────────────────
def rebuild_final_schedule():
    sh       = get_spreadsheet()
    ws_raw   = ensure_worksheet(sh, TAB_RESPONSES, rows=12000, cols=300)
    ws_final = ensure_worksheet(sh, TAB_FINAL,     rows=12000, cols=20)
    _, _, df = ws_get_values_and_df(ws_raw)
    header_final = ["timestamp", "Service date", "Service", "Name", "Age"]
    if df.empty:
        ws_final.clear()
        gs_retry(ws_final.update, "A1", [header_final])
        return 0
    base_cols    = ["timestamp", "Service date", "Family Code", "Name", "Age"]
    service_cols = [c for c in df.columns if c not in base_cols]
    rows_out     = []
    for _, row in df.iterrows():
        timestamp    = row.get("timestamp", "")
        service_date = row.get("Service date", "")
        name         = row.get("Name", "")
        age          = row.get("Age", "")
        for service in service_cols:
            if str(row.get(service, "")).strip().lower() == "yes":
                rows_out.append([timestamp, service_date, service, name, age])
    ws_final.clear()
    gs_retry(ws_final.update, "A1", [header_final])
    if rows_out:
        gs_retry(ws_final.append_rows, rows_out)
    return len(rows_out)


def show_admin_panel():
    with st.expander("Admin"):
        if not ADMIN_KEY:
            st.info("Admin key is not set. Admin tools are open to anyone with the link.")
            unlocked = True
        else:
            key      = st.text_input("Enter admin key", type="password")
            unlocked = key == ADMIN_KEY
            if key and not unlocked:
                st.error("Incorrect admin key.")
            elif unlocked:
                st.success("Admin unlocked.")

        if unlocked:
            if st.button("Rebuild Final Kids Schedule"):
                try:
                    count = rebuild_final_schedule()
                    try:
                        fetch_responses_df.clear()
                    except Exception:
                        pass
                    st.success(f"Final Kids serving dates rebuilt. Rows written: {count}")
                except Exception as e:
                    st.error(f"Failed to rebuild final schedule: {e}")

            if st.button("Preview raw submissions count"):
                try:
                    responses_df = fetch_responses_df()
                    st.write(f"Raw submission rows: **{len(responses_df)}**")
                except Exception as e:
                    st.error(f"Could not load raw submissions: {e}")


# ─────────────────────────────────────────────────────────────
# Load config from Sheets
# ─────────────────────────────────────────────────────────────
try:
    sb_df             = fetch_sb_df()
    deadlines_df      = fetch_deadlines_df()
    service_dates_all = fetch_service_dates_df()
except Exception as e:
    st.error(f"Failed to load config from Google Sheets: {e}")
    show_admin_panel()
    st.stop()

for needed, tab, df in [
    ({"service_date", "deadline_local", "timezone", "Opening date"}, TAB_DEADLINES, deadlines_df),
    ({"service_date", "date", "label", "is_service_day"},            TAB_DATES,     service_dates_all),
    ({"Family Code", "Name & Surname", "Age"},                       TAB_SB,        sb_df),
]:
    miss = needed - set(df.columns)
    if miss:
        st.error(f"Sheet tab '{tab}' is missing columns: {', '.join(sorted(miss))}")
        show_admin_panel()
        st.stop()

# Normalise
deadlines_df["service_date"]   = deadlines_df["service_date"].astype(str).str.strip()
deadlines_df["deadline_local"] = deadlines_df["deadline_local"].astype(str).str.strip()
deadlines_df["timezone"]       = deadlines_df["timezone"].astype(str).str.strip()
deadlines_df["Opening date"]   = deadlines_df["Opening date"].astype(str).str.strip()

service_dates_all["service_date"]   = service_dates_all["service_date"].astype(str).str.strip()
service_dates_all["date"]           = service_dates_all["date"].astype(str).str.strip()
service_dates_all["label"]          = service_dates_all["label"].astype(str).str.strip()
service_dates_all["is_service_day"] = service_dates_all["is_service_day"].astype(str).str.strip()

sb_df["Family Code"]    = sb_df["Family Code"].astype(str).str.strip()
sb_df["Name & Surname"] = sb_df["Name & Surname"].astype(str).str.strip()
sb_df["Age"]            = sb_df["Age"].astype(str).str.strip()

BASE_TZ = "Africa/Johannesburg"
try:
    tz0 = str(deadlines_df["timezone"].iloc[0]).strip()
    if tz0:
        BASE_TZ = tz0
except Exception:
    pass

service_date_key, opening_dt, deadline_dt, deadline_tz = get_currently_open_week(deadlines_df, BASE_TZ)


# ═════════════════════════════════════════════════════════════
# SECTION 1 — FAMILY REGISTRATION
# Always shown so new families can register even when the
# availability window is closed.
# ═════════════════════════════════════════════════════════════
st.markdown(
    """
<div class="ukids-card ukids-card-amber">
  <h3>Register a new child or family</h3>
</div>
""",
    unsafe_allow_html=True,
)

with st.expander("Open registration form", expanded=False):
    st.caption(
        "Use this section if any of your children are not yet on the list. "
        "You can register your whole family at once. "
        "Children already on the list will be skipped automatically."
    )

    # ── How many children? ──
    num_children = st.number_input(
        "How many children are you registering?",
        min_value=1,
        max_value=MAX_CHILDREN_PER_REGISTRATION,
        value=1,
        step=1,
        key="reg_num_children",
    )

    # ── Per-child inputs ──
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
            min_value=2000,
            max_value=datetime.now().year,
            value=datetime.now().year - 8,
            step=1,
            key=f"reg_child_birth_{i}",
        )
        c_age = str(datetime.now().year - int(c_birth))
        if c_name.strip():
            st.caption(f"Will be stored as: **{age_display(c_age)}**")
        st.markdown("</div>", unsafe_allow_html=True)
        reg_children.append({"name": c_name.strip(), "age": c_age})

    # ── Family linking ──
    st.markdown("---")
    st.markdown("**Is this family already partly on the list?**")
    st.caption(
        "If a sibling is already registered, pick their name so the new children "
        "are added under the same family. Otherwise a new family code will be created."
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
        child_names_for_reg = sorted(
            {
                n for n in sb_df["Name & Surname"].tolist()
                if n and str(n).strip().lower() not in ("", "nan")
            }
        )
        chosen_sibling = st.selectbox(
            "Select an existing sibling",
            options=[""] + child_names_for_reg,
            key="reg_sibling",
        )
        if chosen_sibling:
            resolved_fam_code = get_family_code_for_sibling(chosen_sibling, sb_df)
            if resolved_fam_code:
                st.success(f"New children will be added to Family Code **{resolved_fam_code}**.")
            else:
                st.error("Could not find a family code for that sibling. Try the 'new family' option.")
    else:
        next_code         = get_next_family_code(sb_df)
        resolved_fam_code = str(next_code)
        st.info(f"A new Family Code **{resolved_fam_code}** will be assigned to your family.")

    # ── Validate before enabling button ──
    all_names_filled = all(c["name"] for c in reg_children)
    sibling_resolved = (
        link_choice != "Yes — pick an existing sibling"
        or (chosen_sibling and resolved_fam_code)
    )
    reg_ready = all_names_filled and sibling_resolved and resolved_fam_code

    if st.button("Register family", disabled=not reg_ready, key="btn_register"):
        valid_children = [c for c in reg_children if c["name"]]

        # Check each name for duplicates before writing
        duplicate_names = [
            c["name"] for c in valid_children
            if is_duplicate_name(c["name"], sb_df)
        ]
        non_duplicate_children = [
            c for c in valid_children
            if not is_duplicate_name(c["name"], sb_df)
        ]

        if duplicate_names and not non_duplicate_children:
            # Every child entered is already on the list
            st.error(
                f"All the names you entered are already on the list: "
                f"**{', '.join(duplicate_names)}**. "
                "If you believe a child with the same name is being confused with yours, "
                "please contact **Ps Some** to resolve this manually."
            )
        else:
            try:
                written, skipped = register_children_batch(valid_children, resolved_fam_code)
                sb_df = fetch_sb_df()  # refresh local copy

                if written > 0:
                    names_written = [
                        c["name"] for c in valid_children
                        if c["name"].lower() not in [s.lower() for s in skipped]
                    ]
                    st.success(
                        f"Registered {written} child{'ren' if written != 1 else ''}: "
                        f"**{', '.join(names_written)}** "
                        f"(Family Code {resolved_fam_code}). "
                        "They now appear in the availability form below."
                    )
                    st.balloons()

                if skipped:
                    st.warning(
                        f"The following were already on the list and were skipped: "
                        f"**{', '.join(skipped)}**. "
                        "If you think this is a different child with the same name, "
                        "please contact **Ps Some**."
                    )

            except Exception as e:
                st.error(f"Registration failed: {e}")


# ═════════════════════════════════════════════════════════════
# SECTION 2 — AVAILABILITY FORM
# ═════════════════════════════════════════════════════════════

# Gate: form window must be open
if not service_date_key:
    st.markdown("## No availability form is currently open.")
    show_admin_panel()
    st.stop()

week_dates = service_dates_all[
    (service_dates_all["service_date"] == service_date_key)
    & (service_dates_all["is_service_day"].map(_is_truthy))
].copy()

if week_dates.empty:
    st.markdown(f"## No service dates found for **{service_date_key}**.")
    show_admin_panel()
    st.stop()

week_dates["_sort"] = week_dates["date"].map(_safe_parse_date_ymd)
week_dates          = week_dates.sort_values("_sort").drop(columns=["_sort"])
date_labels         = week_dates["label"].astype(str).tolist()

now_local         = get_now_in_tz(deadline_tz)
remaining_seconds = (deadline_dt - now_local).total_seconds()

st.markdown(
    f"""
<div class="info-pill">
  Submitting availability for <strong>{service_date_key}</strong><br>
  Form opened &nbsp;<strong>{opening_dt.strftime('%d %b %Y %H:%M')}</strong> ({deadline_tz})<br>
  Closes at &nbsp;<strong>{deadline_dt.strftime('%d %b %Y %H:%M')}</strong>
  &nbsp;&mdash;&nbsp;<strong>{format_time_remaining(remaining_seconds)}</strong> remaining<br>
  You can submit more than once.
</div>
""",
    unsafe_allow_html=True,
)

if st.button("Refresh timer"):
    st.rerun()

# Refresh sb_df so a just-registered child appears in the dropdown
try:
    sb_df = fetch_sb_df()
    sb_df["Family Code"]    = sb_df["Family Code"].astype(str).str.strip()
    sb_df["Name & Surname"] = sb_df["Name & Surname"].astype(str).str.strip()
    sb_df["Age"]            = sb_df["Age"].astype(str).str.strip()
except Exception:
    pass

child_names = sorted(
    {
        name for name in sb_df["Name & Surname"].tolist()
        if name and str(name).strip().lower() not in ("", "nan")
    }
)


def get_family_code_for_child(child_name: str) -> str:
    rows = sb_df[sb_df["Name & Surname"] == child_name]
    if rows.empty:
        return ""
    return _clean_cell(rows.iloc[0].get("Family Code", ""))


def get_kids_info_for_family_code(family_code: str) -> list:
    rows = sb_df[sb_df["Family Code"] == family_code].copy()
    if rows.empty:
        return []
    out = []
    for idx, row in rows.reset_index(drop=True).iterrows():
        kid_name = _clean_cell(row.get("Name & Surname", ""))
        age_val  = _clean_cell(row.get("Age", ""))
        if not kid_name:
            continue
        out.append({"slot": idx + 1, "name": kid_name, "age": age_val})
    return out


st.markdown(
    """
<div class="ukids-card">
  <h3>Select your child</h3>
</div>
""",
    unsafe_allow_html=True,
)

selected_child = st.selectbox(
    "Search for one child from your family",
    options=[""] + child_names,
    index=0,
    key="main_child_select",
)

if not selected_child:
    st.caption("Select one child to load your family's full availability form.")
    show_admin_panel()
    st.stop()

family_code = get_family_code_for_child(selected_child)

if not family_code:
    st.error("Could not find a family code for the selected child.")
    show_admin_panel()
    st.stop()

kids_info = get_kids_info_for_family_code(family_code)

if not kids_info:
    st.warning("No children found for this family in 'uKids Kids SB'.")
    show_admin_panel()
    st.stop()

st.markdown(
    f"""
<div class="ukids-card">
  <h3>Availability for {service_date_key}</h3>
</div>
""",
    unsafe_allow_html=True,
)

kids_selected_map: dict = {}

for k in kids_info:
    slot     = k["slot"]
    kid_name = k["name"]
    kid_age  = age_display(k["age"])

    st.markdown(
        f"""
<div class="ukids-card">
  <div class="kid-name">{kid_name}</div>
  <div class="kid-age">{kid_age}</div>
</div>
""",
        unsafe_allow_html=True,
    )
    st.caption("Which services can this child attend?")

    selected_labels = []
    for service_label in date_labels:
        key = f"slot{slot}_svc_{service_date_key}_{service_label}"
        if st.checkbox(service_label, key=key):
            selected_labels.append(service_label)

    kids_selected_map[slot] = set(selected_labels)
    st.divider()


# Summary
st.markdown("### Review")
for k in kids_info:
    count = len(kids_selected_map.get(k["slot"], set()))
    label = "service" if count == 1 else "services"
    st.write(f"**{k['name']}** — {count} {label} selected")

st.markdown('<div class="sticky-submit">', unsafe_allow_html=True)
submitted = st.button("Submit availability")
st.markdown("</div>", unsafe_allow_html=True)

if submitted:
    now_check = get_now_in_tz(deadline_tz)
    if not (opening_dt <= now_check < deadline_dt):
        st.error("The form has closed. Submissions are no longer accepted.")
        st.stop()

    now_iso        = datetime.utcnow().isoformat() + "Z"
    desired_header = ["timestamp", "Service date", "Family Code", "Name", "Age"] + date_labels
    rows_to_write  = []

    for k in kids_info:
        slot              = k["slot"]
        selected_services = kids_selected_map.get(slot, set())
        row_map = {
            "timestamp":    now_iso,
            "Service date": service_date_key,
            "Family Code":  family_code,
            "Name":         k["name"],
            "Age":          age_to_number(k["age"]),
        }
        for service_label in date_labels:
            row_map[service_label] = "Yes" if service_label in selected_services else "No"
        rows_to_write.append(row_map)

    try:
        append_multiple_rows(TAB_RESPONSES, desired_header, rows_to_write)
        n = len(rows_to_write)
        st.success(
            f"Submitted. Availability saved for "
            f"{n} child{'ren' if n != 1 else ''}."
        )
        st.balloons()
    except Exception as e:
        st.error(f"Failed to save submission: {e}")


show_admin_panel()
