import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
import pandas as pd

# ----------------------------
# App configuration & header
# ----------------------------
st.set_page_config(page_title="DND War Dashboard (Web)", layout="wide")
st.title("DND War Dashboard — Streamlit Web App")

st.markdown(
    """
This app turns your Excel workbook into a browser UI. Upload **DND.xlsm** (or .xlsx) below.

**Notes**
- Macros (VBA) are **preserved on save** (round-trip) but **not executed** here.
- Use the tabs to view/edit data and then **Download updated workbook**.
"""
)

# ----------------------------
# File uploader
# ----------------------------
uploaded = st.file_uploader(
    "Upload your workbook", type=["xlsm", "xlsx"], accept_multiple_files=False
)

if not uploaded:
    st.info("Upload your Excel file to begin.")
    st.stop()

# Load workbook from memory buffer, preserving VBA if present
try:
    buf_in = BytesIO(uploaded.read())
    wb = load_workbook(buf_in, read_only=False, keep_vba=True, data_only=True)
except Exception as e:
    st.error(f"Failed to load workbook: {e}")
    st.stop()

# Safe sheet getter
def get_ws(wb, name: str):
    """Return worksheet or None if missing."""
    try:
        return wb[name]
    except KeyError:
        return None

# Utility: find header row by a marker (e.g., "RegionName" in col A)
def find_header_row(ws, key="RegionName"):
    if ws is None:
        return None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip().lower() == key.lower():
            return r
    return None

# ----------------------------
# Tabs
# ----------------------------
tab_dash, tab_terr, tab_recon, tab_mon, tab_upg, tab_events, tab_save = st.tabs(
    ["Dashboard", "Territories", "Recon", "Monsters", "Upgrades", "Events", "Save / Download"]
)

# ----------------------------
# Dashboard
# ----------------------------
with tab_dash:
    ws = get_ws(wb, "WarDashboard")
    data_map = {}
    if ws is None:
        st.warning("Sheet 'WarDashboard' not found in your workbook.")
    else:
        for r in range(1, ws.max_row + 1):
            k = ws.cell(row=r, column=1).value
            v = ws.cell(row=r, column=2).value
            if k:
                data_map[str(k)] = v

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Control %", data_map.get("Total Control %", 0))
        c2.metric("Total Income/day", data_map.get("Total Income/day", 0))
        c3.metric("Counterattack Risk", data_map.get("Counterattack Risk", 0))

        st.write("**Next Attack:**", data_map.get("Next Attack", ""))
        st.write("**Attack Type:**", data_map.get("Attack Type", ""))

# ----------------------------
# Territories (view + editor)
# ----------------------------
with tab_terr:
    ws = get_ws(wb, "Territories")
    if ws is None:
        st.warning("Sheet 'Territories' not found.")
    else:
        h = find_header_row(ws, "RegionName")
        if h is None:
            st.warning("Could not find the header row in 'Territories' (looking for 'RegionName' in column A).")
        else:
            headers = [ws.cell(row=h, column=c).value for c in range(1, ws.max_column + 1)]
            rows = []
            for r in range(h + 1, ws.max_row + 1):
                row = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
                if all(v is None for v in row):
                    break
                rows.append(row)
            df_terr = pd.DataFrame(rows, columns=headers)

            st.caption("Edit the table below, then click **Apply changes to workbook**.")
            edited = st.data_editor(df_terr, use_container_width=True, num_rows="dynamic")

            if st.button("Apply changes to workbook"):
                # Re-locate header (defensive)
                h2 = find_header_row(ws, "RegionName")
                start_row = h2 + 1 if h2 else 2

                # Clear existing rows up to a safe margin
                max_cols = max(len(edited.columns), ws.max_column)
                for r in range(start_row, start_row + len(edited) + 10):
                    for c in range(1, max_cols + 1):
                        ws.cell(row=r, column=c).value = None

                # Write edited rows
                for i, row in edited.iterrows():
                    for c, col_name in enumerate(edited.columns, start=1):
                        ws.cell(row=start_row + i, column=c).value = row[col_name]

                st.success("Territories updated (remember to use **Save / Download** tab).")

# ----------------------------
# Recon
# ----------------------------
with tab_recon:
    ws = get_ws(wb, "Recon")
    if ws is None:
        st.warning("Sheet 'Recon' not found.")
    else:
        h = find_header_row(ws, "RegionName")
        if h is None:
            st.warning("Header row not found in 'Recon' (looking for 'RegionName' in column A).")
        else:
            headers = [ws.cell(row=h, column=c).value for c in range(1, ws.max_column + 1)]
            rows = []
            for r in range(h + 1, ws.max_row + 1):
                row = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
                if all(v is None for v in row):
                    break
                rows.append(row)
            df = pd.DataFrame(rows, columns=headers)
            st.dataframe(df, use_container_width=True)

# ----------------------------
# Monsters (simple search)
# ----------------------------
with tab_mon:
    st.write("Quick text search across the 'Monsters' sheet (shows matching rows).")
    q = st.text_input("Search name or text")
    ws = get_ws(wb, "Monsters")

    if ws is None:
        st.warning("Sheet 'Monsters' not found.")
    else:
        matches = []
        if q:
            ql = q.lower()
            # scan first 20 columns as a text row
            for r in range(1, ws.max_row + 1):
                row = [ws.cell(row=r, column=c).value for c in range(1, 21)]
                text = " | ".join("" if x is None else str(x) for x in row)
                if ql in text.lower():
                    matches.append(text)

        if not q:
            st.info("Type a query above to search.")
        elif not matches:
            st.info("No matches found.")
        else:
            # show up to 150 matches
            st.text("\n".join(matches[:150]))

# ----------------------------
# Upgrades (extract tiers + event codes)
# ----------------------------
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

# ----------------------------
# Events (append rows to TerritoryEvents)
# ----------------------------
with tab_events:
    # Regions list from Territories
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

    # Event type list
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

# ----------------------------
# Save & Download (round-trip)
# ----------------------------
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
