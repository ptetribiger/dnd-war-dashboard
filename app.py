
import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
import pandas as pd

# ---------------------------------
# App configuration & introduction
# ---------------------------------
st.set_page_config(page_title="DND War Dashboard (Web)", layout="wide")
st.title("DND War Dashboard — Streamlit Web App")

st.markdown(
    """
Upload **DND.xlsm** (or .xlsx) to work with your campaign data.

**Notes**
- Macros (VBA) are preserved when you download the file, but not executed here.
- Use the tabs to view/edit data, then download the updated workbook.
"""
)

# ---------------------------------
# File uploader
# ---------------------------------
uploaded = st.file_uploader(
    "Upload your workbook", type=["xlsm", "xlsx"], accept_multiple_files=False
)

if not uploaded:
    st.info("Upload your Excel file to begin.")
    st.stop()

# ---------------------------------
# Load workbook (preserve VBA)
# ---------------------------------
try:
    buf_in = BytesIO(uploaded.read())
    wb = load_workbook(buf_in, read_only=False, keep_vba=True, data_only=True)
except Exception as e:
    st.error(f"Failed to load workbook: {e}")
    st.stop()

# ---------------------------------
# Helpers
# ---------------------------------
def get_ws(wb, name: str):
    """Return worksheet or None if missing."""
    try:
        return wb[name]
    except KeyError:
        return None

def find_header_row(ws, key="RegionName"):
    """Find header row index by matching `key` in column A (case-insensitive)."""
    if ws is None:
        return None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip().lower() == key.lower():
            return r
    return None

def make_unique_columns(cols):
    """
    Ensure all column names are strings, not empty, and unique:
    - Replace None/blank with 'ColN'
    - Deduplicate with ' (1)', ' (2)', ...
    """
    norm = [str(c).strip() if (c is not None and str(c).strip() != "") else None for c in cols]
    filled = [c if c is not None else f"Col{i+1}" for i, c in enumerate(norm)]
    seen = {}
    unique = []
    for c in filled:
        if c not in seen:
            seen[c] = 0
            unique.append(c)
        else:
            seen[c] += 1
            unique.append(f"{c} ({seen[c]})")
    return unique

def sheet_to_dataframe_by_key(ws, header_key="RegionName"):
    """
    Build a DataFrame from a sheet that uses a header row identified by `header_key` in col A.
    """
    if ws is None:
        return None, "Worksheet not found."
    h = find_header_row(ws, header_key)
    if h is None:
        return None, f"Header row not found (looking for '{header_key}' in column A)."
    headers = [ws.cell(row=h, column=c).value for c in range(1, ws.max_column + 1)]
    columns = make_unique_columns(headers)

    rows = []
    for r in range(h + 1, ws.max_row + 1):
        row = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if all(v is None for v in row):
            break
        rows.append(row)

    df = pd.DataFrame(rows, columns=columns)
    return df, None

def sheet_to_dataframe_first_row(ws):
    """
    Build a DataFrame assuming row 1 is the header (used for Monsters in the updated workbook).
    """
    if ws is None:
        return None, "Worksheet not found."
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    columns = make_unique_columns(headers)
    rows = []
    for r in range(2, ws.max_row + 1):
        row = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if not any(row):
            continue
        rows.append(row)
    df = pd.DataFrame(rows, columns=columns)
    return df, None

# ---------------------------------
# Tabs
# ---------------------------------
tab_dash, tab_terr, tab_recon, tab_mon, tab_upg, tab_events, tab_save = st.tabs(
    ["Dashboard", "Territories", "Recon", "Monsters", "Upgrades", "Events", "Save / Download"]
)

# ---------------------------------
# Dashboard (label:value from WarDashboard)
# ---------------------------------
with tab_dash:
    ws = get_ws(wb, "WarDashboard")
    data_map = {}
    if ws is None:
        st.warning("Sheet 'WarDashboard' not found.")
    else:
        for r in range(1, ws.max_row + 1):
            key = ws.cell(row=r, column=1).value
            val = ws.cell(row=r, column=2).value
            if key:
                data_map[str(key)] = val

        # Handle potential #N/A / blanks gracefully
        def safe(v):
            return "" if v in (None, "#N/A") else v

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Control %", safe(data_map.get("Total Control %", 0)))
        c2.metric("Total Income/day", safe(data_map.get("Total Income/day", 0)))
        c3.metric("Counterattack Risk", safe(data_map.get("Counterattack Risk", 0)))

        st.write("**Next Attack:**", safe(data_map.get("Next Attack", "")))
        st.write("**Attack Type:**", safe(data_map.get("Attack Type", "")))

# ---------------------------------
# Territories (editable)
# ---------------------------------
with tab_terr:
    ws = get_ws(wb, "Territories")
    if ws is None:
        st.warning("Sheet 'Territories' not found.")
    else:
        df_terr, err = sheet_to_dataframe_by_key(ws, header_key="RegionName")
        if err:
            st.warning(err)
        else:
            st.caption("Edit the table below, then click **Apply changes to workbook**.")
            edited = st.data_editor(df_terr, use_container_width=True, num_rows="dynamic")

            if st.button("Apply changes to workbook"):
                # Re-find header (defensive)
                h2 = find_header_row(ws, "RegionName")
                start_row = h2 + 1 if h2 else 2

                # Align outgoing columns with worksheet max width
                max_cols = max(len(edited.columns), ws.max_column)

                # Clear existing rows up to a safe margin
                for r in range(start_row, start_row + len(edited) + 10):
                    for c in range(1, max_cols + 1):
                        ws.cell(row=r, column=c).value = None

                # Write edited rows in the order of edited DF columns
                for i, row in edited.iterrows():
                    for c, col_name in enumerate(edited.columns, start=1):
                        ws.cell(row=start_row + i, column=c).value = row[col_name]

                st.success("Territories updated (remember to use **Save / Download**).")

# ---------------------------------
# Recon (read-only table; headers sanitized)
# ---------------------------------
with tab_recon:
    ws = get_ws(wb, "Recon")
    if ws is None:
        st.warning("Sheet 'Recon' not found.")
    else:
        df_recon, err = sheet_to_dataframe_by_key(ws, header_key="RegionName")
        if err:
            st.warning(err)
        else:
            st.dataframe(df_recon, use_container_width=True)

# ---------------------------------
# Monsters (first-row header table in the updated workbook)
# ---------------------------------
with tab_mon:
    ws = get_ws(wb, "Monsters")
    if ws is None:
        st.warning("Sheet 'Monsters' not found.")
    else:
        df_mon, err = sheet_to_dataframe_first_row(ws)
        if err:
            st.warning(err)
        else:
            q = st.text_input("Filter by name")
            view = df_mon
            if q:
                if "Name" in df_mon.columns:
                    view = df_mon[df_mon["Name"].astype(str).str.contains(q, case=False, na=False)]
                else:
                    st.info("No 'Name' column detected to filter on.")
            st.dataframe(view, use_container_width=True)

# ---------------------------------
# Upgrades: list tiers & event codes
# ---------------------------------
with tab_upg:
    ws = get_ws(wb, "Upgrade Systems")
    if ws is None:
        st.warning("Sheet 'Upgrade Systems' not found.")
    else:
        weapons, militia, event_types = [], [], []
        for r in range(1, ws.max_row + 1):
            a = ws.cell(row=r, column=1).value
            if not isinstance(a, str):
                continue
            text = a.strip()
            if text.lower().startswith("weapon tier"):
                block = [ws.cell(row=rr, column=1).value for rr in range(r, r + 8)]
                weapons.append("; ".join(str(x) for x in block if x))
            elif text.lower().startswith("militia tier"):
                block = [ws.cell(row=rr, column=1).value for rr in range(r, r + 8)]
                militia.append("; ".join(str(x) for x in block if x))
            elif text.isupper() and "_" in text:
                event_types.append(text)

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Weapon tiers")
            if weapons:
                for w in weapons:
                    st.write("- ", w)
            else:
                st.info("(none detected)")

        with col2:
            st.subheader("Militia tiers")
            if militia:
                for m in militia:
                    st.write("- ", m)
            else:
                st.info("(none detected)")

        st.subheader("Event codes detected")
        st.write(", ".join(sorted(set(event_types))) or "(none)")

# ---------------------------------
# Events (append rows to TerritoryEvents)
# ---------------------------------
with tab_events:
    # Regions for dropdown from Territories
    ws_t = get_ws(wb, "Territories")
    regions = []
    if ws_t:
        h = find_header_row(ws_t, "RegionName")
        if h:
            for r in range(h + 1, ws_t.max_row + 1):
                name = ws_t.cell(row=r, column=1).value
                rest = [ws_t.cell(row=r, column=c).value for c in range(1, ws_t.max_column + 1)]
                if all(v is None for v in rest):
                    break
                if name:
                    regions.append(str(name))

    # Event type list from Upgrade Systems
    ws_u = get_ws(wb, "Upgrade Systems")
    event_types = []
    if ws_u:
        for r in range(1, ws_u.max_row + 1):
            a = ws_u.cell(row=r, column=1).value
            if isinstance(a, str):
                t = a.strip()
                if t.isupper() and "_" in t:
                    event_types.append(t)

    # Ensure TerritoryEvents exists with header
    ws_e = get_ws(wb, "TerritoryEvents")
    if not ws_e:
        ws_e = wb.create_sheet("TerritoryEvents")
        ws_e.cell(row=1, column=1).value = "RegionName"
        ws_e.cell(row=1, column=2).value = "EventType"
        ws_e.cell(row=1, column=3).value = "Value"
        ws_e.cell(row=1, column=4).value = "Notes"

    st.write("Append an event to 'TerritoryEvents'.")
    region = st.selectbox("Region", options=regions or [""], index=0 if regions else None)
    etype = st.selectbox("Event Type", options=sorted(set(event_types)) or [""], index=0 if event_types else None)
    value = st.number_input("Value", value=0.0, step=1.0)
    notes = st.text_input("Notes", value="")

    if st.button("Add Event"):
        target_row = ws_e.max_row + 1
        ws_e.cell(row=target_row, column=1).value = region
        ws_e.cell(row=target_row, column=2).value = etype
        ws_e.cell(row=target_row, column=3).value = float(value)
        ws_e.cell(row=target_row, column=4).value = notes
        st.success("Event appended (remember to use **Save / Download**).")

# ---------------------------------
# Save & Download (preserve macros)
# ---------------------------------
with tab_save:
    st.write("Download the updated workbook (macros preserved if present).")
    try:
        out = BytesIO()
        wb.save(out)
        out.seek(0)

        # choose MIME based on extension
        is_xlsm = uploaded.name.lower().endswith(".xlsm")
        mime = (
            "application/vnd.ms-excel.sheet.macroEnabled.12"
            if is_xlsm
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            "Download updated workbook",
            data=out,
            file_name=uploaded.name,
            mime=mime,
        )
    except Exception as e:
        st.error(f"Failed to prepare download: {e}")
