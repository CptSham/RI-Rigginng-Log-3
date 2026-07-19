"""
Rowing Ireland Rigging Log
A Streamlit dashboard for logging boat rigging data, backed by a Google Sheet
so the data lives in one shared, live spreadsheet that any coach can also
open directly in Google Sheets.
"""

import datetime as dt
import json
import os
import uuid

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

# ---------------------------------------------------------------------------
# Config: boat classes and the metrics/notes we track per seat
# ---------------------------------------------------------------------------

BOAT_CLASSES = {
    "1x": {"seats": 1, "type": "scull", "cox": False, "label": "Single scull"},
    "2x": {"seats": 2, "type": "scull", "cox": False, "label": "Double scull"},
    "2-": {"seats": 2, "type": "sweep", "cox": False, "label": "Coxless pair"},
    "2+": {"seats": 2, "type": "sweep", "cox": True, "label": "Coxed pair"},
    "4x": {"seats": 4, "type": "scull", "cox": False, "label": "Quad scull"},
    "4-": {"seats": 4, "type": "sweep", "cox": False, "label": "Coxless four"},
    "4+": {"seats": 4, "type": "sweep", "cox": True, "label": "Coxed four"},
    "8+": {"seats": 8, "type": "sweep", "cox": True, "label": "Eight"},
}

# (column key, display label, unit, default placeholder)
FIELD_DEFS = [
    ("span_spread", "Span/Spread", "cm", 86.0),
    ("oarlock_height", "Oarlock height", "cm", 16.5),
    ("heel_seat", "Heel-seat distance", "cm", 6.0),
    ("toe_work", "Toe-work distance", "cm", 80.0),
    ("stretcher_angle", "Stretcher angle", "°", 42.0),
    ("fore_aft_pitch", "Fore/aft pitch", "°", 4.0),
    ("lateral_pitch", "Lateral pitch", "°", 0.0),
    ("oar_length", "Oar full length", "cm", 374.0),
    ("oar_inboard", "Oar inboard", "cm", 114.0),
]

NOTE_DEFS = [
    ("seat_height", "Seat height"),
    ("rigger_wedges", "Rigger wedges"),
    ("shoe_wedges", "Shoe wedges"),
]

BOATS_SHEET = "Boats"
BOATS_HEADER = ["boat_id", "name", "class", "sheet_name"]
HISTORY_HEADER = (
    ["date", "note", "seat", "seat_label", "side", "rower_name"]
    + [key for key, *_ in FIELD_DEFS]
    + [key for key, _ in NOTE_DEFS]
)

PORT_COLOR = "#3F9E6D"   # bow side
STROKE_COLOR = "#C4453B"  # stroke side


# ---------------------------------------------------------------------------
# Google Sheets connection
# ---------------------------------------------------------------------------

def _get_service_account_info():
    """Load Google service account credentials from st.secrets (Streamlit
    Community Cloud) or, failing that, from a GCP_SERVICE_ACCOUNT_JSON
    environment variable (Hugging Face Spaces and similar hosts)."""
    try:
        if "gcp_service_account" in st.secrets:
            return dict(st.secrets["gcp_service_account"])
    except Exception:
        pass
    raw = os.environ.get("GCP_SERVICE_ACCOUNT_JSON")
    if raw:
        return json.loads(raw)
    raise RuntimeError(
        "No Google service account credentials found. Set st.secrets "
        "['gcp_service_account'] (Streamlit Cloud) or the "
        "GCP_SERVICE_ACCOUNT_JSON environment variable (Hugging Face Spaces)."
    )


def _get_spreadsheet_id():
    try:
        if "spreadsheet_id" in st.secrets:
            return st.secrets["spreadsheet_id"]
    except Exception:
        pass
    val = os.environ.get("SPREADSHEET_ID")
    if val:
        return val
    raise RuntimeError(
        "No spreadsheet ID found. Set st.secrets['spreadsheet_id'] "
        "(Streamlit Cloud) or the SPREADSHEET_ID environment variable "
        "(Hugging Face Spaces)."
    )


@st.cache_resource(show_spinner=False)
def get_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        _get_service_account_info(), scopes=scopes
    )
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    client = get_client()
    return client.open_by_key(_get_spreadsheet_id())


def ensure_boats_sheet(sh):
    try:
        ws = sh.worksheet(BOATS_SHEET)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=BOATS_SHEET, rows=200, cols=len(BOATS_HEADER))
        ws.append_row(BOATS_HEADER)
    return ws


def seat_label(i, total):
    if i == 1:
        return "Bow"
    if i == total:
        return "Stroke"
    return str(i)


def safe_sheet_name(name, existing_titles):
    base = "".join(c for c in name if c not in "[]:*?/\\").strip()[:80] or "Boat"
    candidate = base
    n = 2
    while candidate in existing_titles:
        candidate = f"{base} ({n})"[:95]
        n += 1
    return candidate


# ---------------------------------------------------------------------------
# Data access
# ---------------------------------------------------------------------------

@st.cache_data(ttl=30, show_spinner=False)
def load_boats():
    sh = get_spreadsheet()
    ws = ensure_boats_sheet(sh)
    records = ws.get_all_records()
    return records


def create_boat(name, cls):
    sh = get_spreadsheet()
    boats_ws = ensure_boats_sheet(sh)
    existing_titles = {w.title for w in sh.worksheets()}
    sheet_name = safe_sheet_name(f"{name}", existing_titles)
    boat_id = uuid.uuid4().hex[:10]

    hist_ws = sh.add_worksheet(title=sheet_name, rows=1000, cols=len(HISTORY_HEADER))
    hist_ws.append_row(HISTORY_HEADER)

    boats_ws.append_row([boat_id, name, cls, sheet_name])
    load_boats.clear()
    return boat_id


def delete_boat(boat_id, sheet_name):
    sh = get_spreadsheet()
    boats_ws = ensure_boats_sheet(sh)
    cell = boats_ws.find(boat_id)
    if cell:
        boats_ws.delete_rows(cell.row)
    try:
        sh.del_worksheet(sh.worksheet(sheet_name))
    except gspread.WorksheetNotFound:
        pass
    load_boats.clear()


@st.cache_data(ttl=15, show_spinner=False)
def load_history(sheet_name):
    sh = get_spreadsheet()
    ws = sh.worksheet(sheet_name)
    records = ws.get_all_records()
    return pd.DataFrame(records)


def append_session(sheet_name, date_str, note, rows):
    sh = get_spreadsheet()
    ws = sh.worksheet(sheet_name)
    ws.append_rows(rows, value_input_option="USER_ENTERED")
    load_history.clear()


# ---------------------------------------------------------------------------
# UI: boat list
# ---------------------------------------------------------------------------

def render_seat_badge(seat_num, label, side, boat_type):
    if boat_type == "scull":
        color = "#C79A4B"
        side_text = "P+S"
    elif side == "port":
        color = PORT_COLOR
        side_text = "BOW SIDE"
    else:
        color = STROKE_COLOR
        side_text = "STROKE SIDE"
    st.markdown(
        f"""
        <div style="border:1px solid {color}; border-radius:8px; padding:6px 10px;
                    display:inline-block; margin:2px; font-family:monospace; font-size:12px;">
            <span style="color:{color}; font-weight:600;">Seat {seat_num}</span>
            <span style="color:#888;"> · {label} · {side_text}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def boats_page():
    st.title("🚣 Rowing Ireland Rigging Log")
    st.caption("Live rigging data, backed by a shared Google Sheet.")

    boats = load_boats()

    if not boats:
        st.info("No boats yet — add your first shell below.")
    else:
        cols = st.columns(3)
        for i, b in enumerate(boats):
            with cols[i % 3]:
                with st.container(border=True):
                    st.subheader(b["name"])
                    cfg = BOAT_CLASSES.get(b["class"], {})
                    st.caption(f"{b['class']} · {cfg.get('label', '')}")
                    if st.button("Open", key=f"open_{b['boat_id']}", use_container_width=True):
                        st.session_state["selected_boat"] = b
                        st.session_state["page"] = "detail"
                        st.rerun()
                    with st.expander("Delete boat"):
                        st.warning("This permanently deletes the boat's full rigging history.")
                        if st.button("Confirm delete", key=f"del_{b['boat_id']}"):
                            delete_boat(b["boat_id"], b["sheet_name"])
                            st.rerun()

    st.divider()
    with st.expander("➕ Add a boat", expanded=not boats):
        with st.form("add_boat_form"):
            name = st.text_input("Boat name", placeholder="e.g. Varsity 8, Bow #2")
            cls = st.selectbox(
                "Class",
                list(BOAT_CLASSES.keys()),
                format_func=lambda k: f"{k} — {BOAT_CLASSES[k]['label']}",
            )
            submitted = st.form_submit_button("Create boat")
            if submitted:
                if not name.strip():
                    st.error("Give the boat a name first.")
                else:
                    create_boat(name.strip(), cls)
                    st.rerun()


# ---------------------------------------------------------------------------
# UI: boat detail (log a session + view history)
# ---------------------------------------------------------------------------

def boat_detail_page():
    boat = st.session_state["selected_boat"]
    cfg = BOAT_CLASSES[boat["class"]]

    if st.button("← All boats"):
        st.session_state["page"] = "list"
        st.rerun()

    st.title(boat["name"])
    st.caption(f"{cfg['label']} · {boat['class']} · {'sweep' if cfg['type'] == 'sweep' else 'scull'}")

    tab_log, tab_history = st.tabs(["Log a session", "History"])

    # ---- Log a session ----
    with tab_log:
        with st.form("session_form"):
            c1, c2 = st.columns([1, 2])
            with c1:
                date_val = st.date_input("Session date", value=dt.date.today())
            with c2:
                note = st.text_input("Note (optional)", placeholder="e.g. windy, raised pitch")

            seat_inputs = []
            for i in range(1, cfg["seats"] + 1):
                label = seat_label(i, cfg["seats"])
                st.markdown(f"#### Seat {i} · {label}")
                default_side = "port" if i % 2 == 1 else "starboard"
                cols = st.columns(2)
                rower_name = cols[0].text_input("Rower name", key=f"name_{i}")
                if cfg["type"] == "sweep":
                    side = cols[1].radio(
                        "Side", ["port", "starboard"],
                        index=0 if default_side == "port" else 1,
                        format_func=lambda s: "Bow side" if s == "port" else "Stroke side",
                        key=f"side_{i}", horizontal=True,
                    )
                else:
                    side = "both"

                metric_cols = st.columns(3)
                values = {}
                for idx, (key, mlabel, unit, placeholder) in enumerate(FIELD_DEFS):
                    with metric_cols[idx % 3]:
                        values[key] = st.number_input(
                            f"{mlabel} ({unit})", value=None,
                            placeholder=str(placeholder), step=0.1,
                            key=f"{key}_{i}", format="%.1f",
                        )

                note_cols = st.columns(3)
                for idx, (key, nlabel) in enumerate(NOTE_DEFS):
                    with note_cols[idx]:
                        values[key] = st.text_input(nlabel, key=f"{key}_{i}")

                seat_inputs.append((i, label, side, rower_name, values))

                if cfg["cox"] and i == cfg["seats"]:
                    st.markdown("#### Coxswain")
                    cox_name = st.text_input("Cox name", key="cox_name")

            submitted = st.form_submit_button("Save session", type="primary")
            if submitted:
                rows = []
                for i, label, side, rower_name, values in seat_inputs:
                    row = [
                        date_val.isoformat(), note, i, label,
                        "" if side == "both" else side, rower_name,
                    ]
                    row += [values[key] if values[key] is not None else "" for key, *_ in FIELD_DEFS]
                    row += [values[key] for key, _ in NOTE_DEFS]
                    rows.append(row)
                if cfg["cox"]:
                    cox_row = [date_val.isoformat(), note, cfg["seats"] + 1, "Cox", "", st.session_state.get("cox_name", "")]
                    cox_row += [""] * (len(FIELD_DEFS) + len(NOTE_DEFS))
                    rows.append(cox_row)

                append_session(boat["sheet_name"], date_val.isoformat(), note, rows)
                st.success(f"Session saved for {date_val.isoformat()}.")
                st.rerun()

    # ---- History ----
    with tab_history:
        df = load_history(boat["sheet_name"])
        if df.empty:
            st.info("No sessions logged yet.")
        else:
            seat_options = sorted(df["seat"].unique().tolist())
            chosen_seat = st.selectbox(
                "Seat", seat_options,
                format_func=lambda s: f"Seat {s} ({seat_label(int(s), cfg['seats'])})"
            )
            seat_df = df[df["seat"] == chosen_seat].sort_values("date")

            display_cols = ["date", "note", "rower_name", "side"] + \
                [k for k, *_ in FIELD_DEFS] + [k for k, _ in NOTE_DEFS]
            display_cols = [c for c in display_cols if c in seat_df.columns]
            st.dataframe(seat_df[display_cols], use_container_width=True, hide_index=True)

            metric_options = {mlabel: key for key, mlabel, unit, ph in FIELD_DEFS}
            chosen_metric_label = st.selectbox("Chart a metric over time", list(metric_options.keys()))
            chosen_metric = metric_options[chosen_metric_label]
            chart_df = seat_df[["date", chosen_metric]].copy()
            chart_df[chosen_metric] = pd.to_numeric(chart_df[chosen_metric], errors="coerce")
            chart_df = chart_df.dropna().set_index("date")
            if not chart_df.empty:
                st.line_chart(chart_df)

            st.caption("This data lives in your Google Sheet — open it directly any time for pivot tables or backups.")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    st.set_page_config(page_title="Rowing Ireland Rigging Log", page_icon="🚣", layout="wide")

    if "page" not in st.session_state:
        st.session_state["page"] = "list"

    if st.session_state["page"] == "detail" and "selected_boat" in st.session_state:
        boat_detail_page()
    else:
        boats_page()


if __name__ == "__main__":
    main()
