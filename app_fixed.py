# app_fixed.py
# Weekly uKids Kids Availability Form
#
# Parent selects/searches one child.
# App loads all siblings with the same Family Code.
#
# RAW OUTPUT:
# uKids Kids responses
# timestamp | Service date | Family Code | Name | Age | service columns as Yes/No
#
# FINAL OUTPUT:
# Final Kids serving dates
# timestamp | Service date | Service | Name | Age

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


st.set_page_config(page_title="uKids Kids Availability Form", page_icon="🗓️", layout="centered")
st.title("🗓️ uKids Kids Availability Form")

st.markdown(
    """
<style>
  .stButton > button { width: 100%; height: 48px; font-size: 16px; }
  @media (max-width: 520px){
    div[data-testid="column"] { width: 100% !important; flex: 0 0 100% !important; }
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
TAB_FINAL = "Final Kids serving dates"


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
    sa = _get_secret_any(["gcp_service_account"], ["general", "gcp_service_account"])
    sheet_id = _get_secret_any(["GSHEET_ID"], ["general", "GSHEET_ID"])

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


def ws_get_values_and_df(ws):
    values = gs_retry(ws.get_all_values)
    if not values:
        return [], [], pd.DataFrame()

    header_raw = values[0]
    header_unique = make_unique_headers(header_raw)
    rows = values[1:]
    return header_raw, header_unique, pd.DataFrame(rows, columns=header_unique)


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


def append_multiple_rows(sheet_title: str, desired_header: list[str], row_maps: list[dict]):
    sh = get_spreadsheet()
    ws = ensure_worksheet(sh, sheet_title, rows=12000, cols=max(100, len(desired_header) + 10))
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


def get_now_in_tz(tz_name: str) -> datetime:
    if ZoneInfo is None:
        return datetime.utcnow()
    return datetime.now(ZoneInfo(tz_name))


def parse_local_dt(value: str, tz_name: str) -> datetime:
    dt_naive = datetime.strptime(str(value).strip(), "%Y-%m-%d %H:%M")
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


def _is_truthy(v) -> bool:
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


def _clean_cell(x) -> str:
    s = str(x).strip()
    return "" if (s == "" or s.lower() == "nan") else s


def get_currently_open_week(deadlines: pd.DataFrame, base_tz: str):
    candidates = []

    for _, row in deadlines.iterrows():
        service_date = str(row.get("service_date", "")).strip()
        tz_name = str(row.get("timezone", "")).strip() or base_tz
        opening_raw = str(row.get("Opening date", "")).strip()
        deadline_raw = str(row.get("deadline_local", "")).strip()

        if not service_date or not opening_raw or not deadline_raw:
            continue

        try:
            opening_dt = parse_local_dt(opening_raw, tz_name)
            deadline_dt = parse_local_dt(deadline_raw, tz_name)
            now_local = get_now_in_tz(tz_name)
        except Exception:
            continue

        if opening_dt <= now_local < deadline_dt:
            candidates.append((service_date, opening_dt, deadline_dt, tz_name))

    if not candidates:
        return None, None, None, base_tz

    candidates.sort(key=lambda x: x[2])
    return candidates[0]


def rebuild_final_schedule():
    sh = get_spreadsheet()
    ws_raw = ensure_worksheet(sh, TAB_RESPONSES, rows=12000, cols=300)
    ws_final = ensure_worksheet(sh, TAB_FINAL, rows=12000, cols=20)

    _, _, df = ws_get_values_and_df(ws_raw)

    header_final = ["timestamp", "Service date", "Service", "Name", "Age"]

    if df.empty:
        ws_final.clear()
        gs_retry(ws_final.update, "A1", [header_final])
        return 0

    base_cols = ["timestamp", "Service date", "Family Code", "Name", "Age"]
    service_cols = [c for c in df.columns if c not in base_cols]

    rows_out = []

    for _, row in df.iterrows():
        timestamp = row.get("timestamp", "")
        service_date = row.get("Service date", "")
        name = row.get("Name", "")
        age = row.get("Age", "")

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
            st.info("Admin key is not set. Admin tools are currently open to anyone with the link.")
            unlocked = True
        else:
            key = st.text_input("Enter admin key", type="password")
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


# Load config
try:
    sb_df = fetch_sb_df()
    deadlines_df = fetch_deadlines_df()
    service_dates_all = fetch_service_dates_df()
except Exception as e:
    st.error(f"Failed to load config from Google Sheets: {e}")
    show_admin_panel()
    st.stop()


for needed, tab, df in [
    ({"service_date", "deadline_local", "timezone", "Opening date"}, TAB_DEADLINES, deadlines_df),
    ({"service_date", "date", "label", "is_service_day"}, TAB_DATES, service_dates_all),
    ({"Family Code", "Name & Surname", "Age"}, TAB_SB, sb_df),
]:
    miss = needed - set(df.columns)
    if miss:
        st.error(f"Google Sheet tab '{tab}' is missing columns: {', '.join(sorted(miss))}")
        show_admin_panel()
        st.stop()


deadlines_df["service_date"] = deadlines_df["service_date"].astype(str).str.strip()
deadlines_df["deadline_local"] = deadlines_df["deadline_local"].astype(str).str.strip()
deadlines_df["timezone"] = deadlines_df["timezone"].astype(str).str.strip()
deadlines_df["Opening date"] = deadlines_df["Opening date"].astype(str).str.strip()

service_dates_all["service_date"] = service_dates_all["service_date"].astype(str).str.strip()
service_dates_all["date"] = service_dates_all["date"].astype(str).str.strip()
service_dates_all["label"] = service_dates_all["label"].astype(str).str.strip()
service_dates_all["is_service_day"] = service_dates_all["is_service_day"].astype(str).str.strip()

sb_df["Family Code"] = sb_df["Family Code"].astype(str).str.strip()
sb_df["Name & Surname"] = sb_df["Name & Surname"].astype(str).str.strip()
sb_df["Age"] = sb_df["Age"].astype(str).str.strip()

BASE_TZ = "Africa/Johannesburg"
try:
    tz0 = str(deadlines_df["timezone"].iloc[0]).strip()
    if tz0:
        BASE_TZ = tz0
except Exception:
    pass

service_date_key, opening_dt, deadline_dt, deadline_tz = get_currently_open_week(deadlines_df, BASE_TZ)

if not service_date_key:
    st.markdown("## 🔒 No availability form is currently open.")
    show_admin_panel()
    st.stop()


week_dates = service_dates_all[
    (service_dates_all["service_date"] == service_date_key)
    & (service_dates_all["is_service_day"].map(_is_truthy))
].copy()

if week_dates.empty:
    st.markdown(
        f"""
        ## 🔒 This week’s availability form is not open yet.

        No service dates were found for **{service_date_key}**.
        """
    )
    show_admin_panel()
    st.stop()

week_dates["_sort"] = week_dates["date"].map(_safe_parse_date_ymd)
week_dates = week_dates.sort_values("_sort").drop(columns=["_sort"])

date_labels = week_dates["label"].astype(str).tolist()

morning_labels = [l for l in date_labels if "morning" in l.lower()]
evening_labels = [l for l in date_labels if "evening" in l.lower()]

morning_display_map = _build_display_map(morning_labels)
evening_display_map = _build_display_map(evening_labels)

morning_options = list(morning_display_map.keys())
evening_options = list(evening_display_map.keys())

now_local = get_now_in_tz(deadline_tz)
remaining_seconds = (deadline_dt - now_local).total_seconds()

st.info(
    f"🗓️ Submitting availability for **{service_date_key}**.\n\n"
    f"📬 Form opened at **{opening_dt.strftime('%Y-%m-%d %H:%M')}** ({deadline_tz}).\n\n"
    f"⏳ Form closes at **{deadline_dt.strftime('%Y-%m-%d %H:%M')}** ({deadline_tz}). "
    f"Time remaining: **{format_minutes_remaining(remaining_seconds)}**\n\n"
    f"🔁 You can submit more than once."
)

if st.button("Refresh timer"):
    st.rerun()


child_names = sorted(
    {
        name for name in sb_df["Name & Surname"].tolist()
        if name and str(name).strip().lower() != "nan"
    }
)


def get_family_code_for_child(child_name: str) -> str:
    rows = sb_df[sb_df["Name & Surname"] == child_name]
    if rows.empty:
        return ""
    return _clean_cell(rows.iloc[0].get("Family Code", ""))


def get_kids_info_for_family_code(family_code: str) -> list[dict]:
    rows = sb_df[sb_df["Family Code"] == family_code].copy()
    if rows.empty:
        return []

    out = []
    for idx, row in rows.reset_index(drop=True).iterrows():
        kid_name = _clean_cell(row.get("Name & Surname", ""))
        age_val = _clean_cell(row.get("Age", ""))
        if not kid_name:
            continue
        out.append({"slot": idx + 1, "name": kid_name, "age": age_val})
    return out


st.subheader("Your details")
selected_child = st.selectbox("Select one child from your family", options=[""] + child_names, index=0)

if not selected_child:
    st.caption("Search for and select one child to continue.")
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

st.subheader(f"Availability for {service_date_key}")

kids_selected_map: dict[int, set[str]] = {}

for k in kids_info:
    slot = k["slot"]
    kid_name = k["name"]

    st.markdown(f"## {kid_name}")
    st.caption("Select all services this child is available for.")

    if morning_options:
        st.markdown("### Which morning services are you available?")
        m1, m2 = st.columns(2)
        with m1:
            if st.button(f"Select all mornings ({kid_name})"):
                for opt in morning_options:
                    st.session_state[f"slot{slot}_morn_{service_date_key}_{opt}"] = True
        with m2:
            if st.button(f"Clear mornings ({kid_name})"):
                for opt in morning_options:
                    st.session_state[f"slot{slot}_morn_{service_date_key}_{opt}"] = False

        chosen_m = []
        for opt in morning_options:
            key = f"slot{slot}_morn_{service_date_key}_{opt}"
            if st.checkbox(opt, key=key):
                chosen_m.append(opt)
    else:
        chosen_m = []

    if evening_options:
        st.divider()
        st.markdown("### Which evening services are you available?")
        e1, e2 = st.columns(2)
        with e1:
            if st.button(f"Select all evenings ({kid_name})"):
                for opt in evening_options:
                    st.session_state[f"slot{slot}_eve_{service_date_key}_{opt}"] = True
        with e2:
            if st.button(f"Clear evenings ({kid_name})"):
                for opt in evening_options:
                    st.session_state[f"slot{slot}_eve_{service_date_key}_{opt}"] = False

        chosen_e = []
        for opt in evening_options:
            key = f"slot{slot}_eve_{service_date_key}_{opt}"
            if st.checkbox(opt, key=key):
                chosen_e.append(opt)
    else:
        chosen_e = []

    selected_morning_labels = {morning_display_map[d] for d in chosen_m if d in morning_display_map}
    selected_evening_labels = {evening_display_map[d] for d in chosen_e if d in evening_display_map}
    kids_selected_map[slot] = selected_morning_labels.union(selected_evening_labels)

    st.divider()


st.subheader("Review")
st.write(f"**Selected child:** {selected_child}")
for k in kids_info:
    st.write(f"- **{k['name']}:** {len(kids_selected_map.get(k['slot'], set()))} services selected")

st.markdown('<div class="sticky-submit">', unsafe_allow_html=True)
submitted = st.button("Submit")
st.markdown("</div>", unsafe_allow_html=True)

if submitted:
    now_check = get_now_in_tz(deadline_tz)
    if not (opening_dt <= now_check < deadline_dt):
        st.error("Form is currently closed.")
        st.stop()

    now_iso = datetime.utcnow().isoformat() + "Z"
    desired_header = ["timestamp", "Service date", "Family Code", "Name", "Age"] + date_labels

    rows_to_write = []
    for k in kids_info:
        slot = k["slot"]
        selected_services = kids_selected_map.get(slot, set())

        row_map = {
            "timestamp": now_iso,
            "Service date": service_date_key,
            "Family Code": family_code,
            "Name": k["name"],
            "Age": k["age"],
        }

        for service_label in date_labels:
            row_map[service_label] = "Yes" if service_label in selected_services else "No"

        rows_to_write.append(row_map)

    try:
        append_multiple_rows(TAB_RESPONSES, desired_header, rows_to_write)
        st.success(f"Submission saved to Google Sheets. Rows added: {len(rows_to_write)}")
    except Exception as e:
        st.error(f"Failed to save submission: {e}")


show_admin_panel()
