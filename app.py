#!/usr/bin/env python3
"""
Interview Tracker — Streamlit web app (deployable to a public URL).

Local run:   streamlit run app.py
Deploy:      push app.py + interview_automation.py + requirements.txt
             (+ the .streamlit/ folder) to GitHub, deploy on share.streamlit.io.

The heavy lifting lives in interview_automation.py; this file is the UI.
"""

import os
import tempfile
from datetime import date

import pandas as pd
import streamlit as st

import interview_automation as engine

DEFAULT_HR = ["Aaruni", "Kishor", "Vartika", "Garima", "George", "Musarat"]
BRAND = "#A100FF"        # Accenture purple
BRAND_DARK = "#7500C0"
HR_COL = "HR interviewer"
XLSX_MIME = ("application/vnd.openxmlformats-officedocument."
             "spreadsheetml.sheet")

st.set_page_config(page_title="Interview Tracker", page_icon="🗂",
                   layout="centered")

# --------------------------------------------------------------------------- #
# Look & feel
# --------------------------------------------------------------------------- #
st.markdown(
    f"""
    <style>
      .block-container {{ padding-top: 1.6rem; max-width: 820px; }}

      /* --- hero banner --- */
      .banner {{
          position: relative; overflow: hidden;
          background: linear-gradient(135deg, {BRAND} 0%, {BRAND_DARK} 100%);
          color:#fff; padding: 30px 34px; border-radius: 18px;
          margin-bottom: 24px; box-shadow: 0 10px 28px rgba(161,0,255,.28); }}
      .banner::after {{
          content: ">"; position: absolute; right: -6px; top: -54px;
          font-size: 230px; font-weight: 800; line-height: 1;
          color: rgba(255,255,255,.14); pointer-events: none; }}
      .banner h1 {{ font-size: 28px; margin: 0; font-weight: 800;
                    position: relative; z-index: 1; }}
      .banner p  {{ color:#F0DCFF; margin: 6px 0 0; font-size: 14.5px;
                    position: relative; z-index: 1; }}

      /* --- left-accent panel titles --- */
      .panel-title {{
          border-left: 4px solid {BRAND}; padding-left: 12px;
          font-size: 1.05rem; font-weight: 700; margin-bottom: 2px;
          color: #26313D; }}

      /* card-like bordered containers */
      div[data-testid="stVerticalBlockBorderWrapper"] {{
          border-radius: 16px; box-shadow: 0 1px 3px rgba(161,0,255,.10); }}

      /* --- buttons: purple gradient --- */
      .stButton > button, .stDownloadButton > button {{
          width: 100%; border-radius: 10px; font-weight: 700; padding: .55rem 1rem;
          border: none; color: #fff;
          background: linear-gradient(135deg, {BRAND}, {BRAND_DARK});
          transition: filter .15s, box-shadow .15s; }}
      .stButton > button:hover, .stDownloadButton > button:hover {{
          filter: brightness(1.08);
          box-shadow: 0 4px 14px rgba(161,0,255,.30); }}
      .stButton > button:disabled {{
          background: #E3E3E8 !important; color: #9AA0A6 !important; }}
      .stButton > button[kind="primary"] {{ font-size: 1.05rem; padding: .7rem; }}

      /* --- metric tiles --- */
      div[data-testid="stMetric"] {{
          background:#F6EBFF; border:1px solid #E4C6FF; border-radius:14px;
          padding:14px 16px; }}
      div[data-testid="stMetricValue"] {{ color:{BRAND_DARK}; font-weight:800; }}

      section[data-testid="stFileUploaderDropzone"] {{ border-radius: 12px; }}
    </style>
    """,
    unsafe_allow_html=True,
)


def panel_title(text):
    st.markdown(f'<div class="panel-title">{text}</div>', unsafe_allow_html=True)


st.markdown(
    """<div class="banner">
         <h1>🗂 Interview Tracker</h1>
         <p>Import candidates · auto-assign HR (POC) · download a ready tracker</p>
       </div>""",
    unsafe_allow_html=True,
)

with st.expander("How it works"):
    st.markdown(
        "1. **Keeps** rows whose Interview Status is one of the 6 valid values.\n"
        "2. **Scheduled** skill/final interviews are kept only if dated the day you pick.\n"
        "3. **SAP / Oracle / Salesforce / Agentforce** are split evenly across your HR.\n"
        "4. **All other skills** are split evenly too.\n"
        "5. Each in-scope row gets a **POC** (the assigned HR).\n"
        "6. An **Updates** dropdown is added for dispositions.\n\n"
        "Out-of-scope rows stay in the file (greyed, POC blank) — nothing is lost."
    )


# --------------------------------------------------------------------------- #
# Helper
# --------------------------------------------------------------------------- #
def build_output(upload, names, run_date):
    with tempfile.TemporaryDirectory() as d:
        ip = os.path.join(d, "in.xlsx")
        op = os.path.join(d, "out.xlsx")
        with open(ip, "wb") as f:
            f.write(upload.getbuffer())
        _, stats = engine.process_file(ip, op, names, run_date)
        with open(op, "rb") as f:
            return stats, f.read()


# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
with st.container(border=True):
    panel_title("👥 &nbsp;1 · HR interviewers")
    st.caption("These become the POC options. Edit names, or use the ＋ row to "
               "add and the 🗑 to remove.")
    edited = st.data_editor(
        pd.DataFrame({HR_COL: DEFAULT_HR}),
        num_rows="dynamic", hide_index=True, use_container_width=True,
        column_config={HR_COL: st.column_config.TextColumn(
            "HR interviewer", width="large")},
        key="hr_editor",
    )
    names = [str(n).strip() for n in edited[HR_COL].tolist()
             if pd.notna(n) and str(n).strip()]

col_a, col_b = st.columns(2)
with col_a:
    with st.container(border=True):
        panel_title("📅 &nbsp;2 · Date")
        run_date = st.date_input("Interview date", value=date.today(),
                                 help="Scheduled skill/final interviews are kept "
                                      "only if dated this day.")
with col_b:
    with st.container(border=True):
        panel_title("📁 &nbsp;3 · File")
        upload = st.file_uploader("Candidate Excel (.xlsx)",
                                  type=["xlsx", "xlsm"],
                                  label_visibility="collapsed")

# Small live readiness hint.
ready = len(names) >= 2 and len(set(names)) == len(names) and upload is not None
if not ready:
    bits = []
    if len(names) < 2:
        bits.append("add at least 2 HR names")
    elif len(set(names)) != len(names):
        bits.append("remove duplicate HR names")
    if upload is None:
        bits.append("upload a file")
    st.info("To run: " + " and ".join(bits) + ".")

run = st.button("Run  ▶", type="primary", disabled=not ready)

# --------------------------------------------------------------------------- #
# Run
# --------------------------------------------------------------------------- #
if run:
    try:
        with st.spinner("Assigning POCs and building your tracker…"):
            stats, data = build_output(upload, names, run_date)
        stem = os.path.splitext(upload.name)[0]
        st.session_state["result"] = {
            "stats": stats, "data": data, "name": f"{stem}_Tracker.xlsx"}
    except Exception as exc:                            # surface, don't crash
        st.session_state.pop("result", None)
        st.error(f"Could not process file: {type(exc).__name__}: {exc}")

# --------------------------------------------------------------------------- #
# Results (persist across reruns so the download button keeps working)
# --------------------------------------------------------------------------- #
res = st.session_state.get("result")
if res:
    stats = res["stats"]
    st.divider()
    top = st.columns([3, 2])
    with top[0]:
        st.markdown(f"### ✅ Tracker ready")
        st.caption(f"📅 For {stats['today'].strftime('%A, %d %b %Y')}")
    with top[1]:
        st.download_button("↓ Download tracker", data=res["data"],
                           file_name=res["name"], mime=XLSX_MIME,
                           type="primary")

    c1, c2, c3 = st.columns(3)
    c1.metric("Rows in file", stats["n_rows"])
    c2.metric("In scope · got POC", stats["in_scope"],
              help="Valid status, and (for scheduled rows) dated the chosen day.")
    c3.metric("Parked · no POC", stats["parked"],
              help="Kept in the file, greyed, with POC left blank.")

    dist = stats["distribution"]
    totals = [d["total"] for d in dist]

    panel_title("Interviews per HR")
    chart_df = pd.DataFrame(
        {"HR": [d["hr"] for d in dist], "Interviews": totals}).set_index("HR")
    st.bar_chart(chart_df, horizontal=True, color=BRAND)
    if totals and max(totals) - min(totals) <= 1:
        st.caption("✓ Evenly balanced across HR (totals differ by at most one).")

    panel_title("Breakdown by skill")
    matrix = {sk: [d["by_skill"].get(sk, 0) for d in dist]
              for sk in stats["skills"]}
    table = pd.DataFrame(matrix, index=[d["hr"] for d in dist])
    table["Total"] = totals
    table.index.name = "HR"
    st.dataframe(
        table, use_container_width=True,
        column_config={"Total": st.column_config.ProgressColumn(
            "Total", format="%d", min_value=0,
            max_value=max(totals) if totals else 1)},
    )
