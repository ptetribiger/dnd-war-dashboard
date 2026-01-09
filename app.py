
import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
import pandas as pd

st.set_page_config(page_title="DND War Dashboard (Web)", layout="wide")
st.title("DND War Dashboard — Streamlit Web App")

st.markdown("""
This app turns your Excel workbook into a browser UI. Upload **DND.xlsm** below.
- Macros (VBA) are **preserved on save** (round-trip) but **not executed** in this web app.
- Use the tabs to view/edit data and then **Download updated workbook**.
""")

uploaded = st.file_uploader("Upload your workbook", type=["xlsm", "xlsx"], accept_multiple_files=False)

if not uploaded:
    st.info("Upload your Excel file to begin.")
    st.stop()

# Load workbook from memory buffer, preserving VBA if present
buf_in = BytesIO(uploaded.read())
wb = load_workbook(buf_in, read_only=False, keep_vba=True, data_only=True)

# ---------- helpers ----------
def find_header_row(ws, key="RegionName"):
    for r in range(1, ws.max_row+1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip().lower() == key.lower():
            return r
    return None

# Dashboard
tab_dash, tab_terr, tab_recon, tab_mon, tab_upg, tab_events, tab_save = st.tabs(
    ["Dashboard", "Territories", "Recon", "Monsters", "Upgrades", "Events", "Save / Download"]
)

with tab_dash:
    ws = wb.get("WarDashboard")
    data_map = {}
    if ws:
        for r in range(1, ws.max_row+1):
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

# Territories view + editor
with tab_terr:
    ws = wb.get("Territories")
    df_terr = None
    if ws:
        h = find_header_row(ws)
        if h:
            headers = [ws.cell(row=h, column=c).value for c in range(1, ws.max_column+1)]
            rows = []
            for r in range(h+1, ws.max_row+1):
                row = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column+1)]
                if all(v is None for v in row):
                    break
                rows.append(row)
            df_terr = pd.DataFrame(rows, columns=headers)
            st.caption("Edit the table below, then click **Apply changes to workbook**.")
            edited = st.data_editor(df_terr, use_container_width=True, num_rows="dynamic")
            if st.button("Apply changes to workbook"):
                # write back edited DF
                # re-locate header row (in case of changes)
                h2 = find_header_row(ws)
                start_row = h2 + 1 if h2 else 2
                # clear existing rows up to a safe limit
                for r in range(start_row, start_row + len(edited) + 10):
                    for c in range(1, ws.max_column+1):
                        ws.cell(row=r, column=c).value = None
                # write edited rows
                for i, row in edited.iterrows():
                    for c, col_name in enumerate(edited.columns, start=1):
                        ws.cell(row=start_row + i, column=c).value = row[col_name]
                st.success("Territories updated (remember to Download updated workbook).")
        else:
            st.warning("Could not find the header row in 'Territories'.")
    else:
        st.warning("'Territories' sheet not found.")

# Recon
with tab_recon:
    ws = wb.get("Recon")
    if ws:
        h = find_header_row(ws)
        if h:
            headers = [ws.cell(row=h, column=c).value for c in range(1, ws.max_column+1)]
            rows = []
            for r in range(h+1, ws.max_row+1):
                row = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column+1)]
                if all(v is None for v in row):
                    break
                rows.append(row)
            df = pd.DataFrame(rows, columns=headers)
            st.dataframe(df, use_container_width=True)
        else:
            st.warning("Header row not found in 'Recon'.")
    else:
        st.warning("'Recon' sheet not found.")

# Monsters quick search (simple text search across rows)
with tab_mon:
    st.write("Quick text search across the 'Monsters' sheet (shows matching rows).")
    q = st.text_input("Search name or text")
    ws = wb.get("Monsters")
    if ws:
        matches = []
        if q:
            ql = q.lower()
            for r in range(1, ws.max_row+1):
                row = [ws.cell(row=r, column=c).value for c in range(1, 20)]
                text = " | ".join("" if x is None else str(x) for x in row)
                if ql in text.lower():
                    matches.append(text)
        st.text("\n".join(matches[:150]) if matches else ("Type a query above" if not q else "No matches"))
    else:
        st.warning("'Monsters' sheet not found.")

# Upgrades: weapon/militia tiers + event codes (basic extraction)
with tab_upg:
    ws = wb.get("Upgrade Systems")
    if ws:
        weapons, militia, event_types = [], [], []
        for r in range(1, ws.max_row+1):
            a = ws.cell(row=r, column=1).value
            if not isinstance(a, str):
                continue
            text = a.strip()
            if text.lower().startswith("weapon tier"):
                block = [ws.cell(row=rr, column=1).value for rr in range(r, r+8)]
                weapons.append("; ".join(str(x) for x in block if x))
            elif text.lower().startswith("militia tier"):
                block = [ws.cell(row=rr, column=1).value for rr in range(r, r+8)]
                militia.append("; ".join(str(x) for x in block if x))
            elif text.isupper() and "_" in text:
                event_types.append(text)
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Weapon tiers")
            for w in weapons:
                st.write("- ", w)
        with col2:
            st.subheader("Militia tiers")
            for m in militia:
                st.write("- ", m)
        st.subheader("Event codes detected")
        st.write(", ".join(sorted(set(event_types))) or "(none)")
    else:
        st.warning("'Upgrade Systems' sheet not found.")

# Events appender
with tab_events:
    ws_t = wb.get("Territories")
    regions = []
    if ws_t:
        h = find_header_row(ws_t)
        if h:
            for r in range(h+1, ws_t.max_row+1):
                name = ws_t.cell(row=r, column=1).value
                rest = [ws_t.cell(row=r, column=c).value for c in range(1, ws_t.max_column+1)]
                if all(v is None for v in rest):
                    break
                if name:
                    regions.append(str(name))
    ws_u = wb.get("Upgrade Systems")
    event_types = []
    if ws_u:
        for r in range(1, ws_u.max_row+1):
            a = ws_u.cell(row=r, column=1).value
            if isinstance(a, str):
                t = a.strip()
                if t.isupper() and "_" in t:
                    event_types.append(t)
    ws_e = wb.get("TerritoryEvents")
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
        st.success("Event appended (remember to Download updated workbook).")

# Save/Download
with tab_save:
    st.write("Download the updated workbook (macros preserved if present).")
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    st.download_button(
        "Download updated workbook",
        data=out,
        file_name=uploaded.name,
        mime="application/vnd.ms-excel.sheet.macroEnabled.12" if uploaded.name.lower().endswith(".xlsm") else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
