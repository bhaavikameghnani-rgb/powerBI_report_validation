"""Power BI Validator — redesigned Streamlit UI.

Run locally: python -m streamlit run app_powerbi.py
Requires powerbi_validator_core.py in the same folder.
UI styling/layout has been refreshed; validation logic remains in the core module.
"""

import os
import shutil
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st

import powerbi_validator_core as core

st.set_page_config(page_title="Power BI Validator", page_icon="📊", layout="wide")

# Presentation-only styling. Validation functions and report-generation calls are unchanged.
st.markdown(
    """
    <style>
    :root { --ink:#172b4d; --muted:#66758b; --teal:#087f8c; --teal-dark:#07636d;
            --line:#dce5ed; --surface:#ffffff; --canvas:#f4f7fb; }
    .stApp { background:var(--canvas); }
    [data-testid="stHeader"] { background:rgba(244,247,251,.92); }
    .block-container { max-width:1440px; padding-top:1.6rem; padding-bottom:3rem; }
    h1,h2,h3 { color:var(--ink); letter-spacing:-.025em; }
    h1 { font-size:2.05rem !important; font-weight:760 !important; }
    h2 { font-size:1.35rem !important; }
    h3 { font-size:1.05rem !important; }
    [data-testid="stCaptionContainer"] { color:var(--muted); }
    .hero { background:linear-gradient(115deg,#102a43 0%,#145c70 65%,#087f8c 100%);
            border-radius:20px; padding:1.6rem 1.8rem; margin:.3rem 0 1.35rem 0;
            color:white; box-shadow:0 10px 28px rgba(16,42,67,.13); }
    .hero-kicker { color:#bcebf0; text-transform:uppercase; letter-spacing:.13em;
                   font-size:.72rem; font-weight:750; margin-bottom:.45rem; }
    .hero-title { color:white; font-size:1.85rem; font-weight:780; line-height:1.2; }
    .hero-copy { color:#e2f3f5; margin-top:.55rem; max-width:920px; font-size:.96rem; }
    .section-label { color:#52657b; text-transform:uppercase; letter-spacing:.1em;
                     font-size:.74rem; font-weight:750; margin:.3rem 0 .65rem; }
    div[data-testid="stVerticalBlockBorderWrapper"] { border-color:var(--line); border-radius:16px; }
    div[data-testid="stFileUploader"] { background:#fbfdff; border:1px dashed #b8c8d8;
        border-radius:12px; padding:.45rem; }
    div.stButton > button[kind="primary"], div.stDownloadButton > button {
        background:var(--teal); color:white; border:0; border-radius:10px;
        font-weight:700; min-height:2.7rem; padding:.55rem 1.1rem; }
    div.stButton > button[kind="primary"]:hover, div.stDownloadButton > button:hover {
        background:var(--teal-dark); color:white; border:0; }
    div[data-testid="stMetric"] { background:white; border:1px solid var(--line);
        padding:1rem 1.05rem; border-radius:14px; box-shadow:0 2px 7px rgba(23,43,77,.025); }
    div[data-testid="stMetricLabel"] { color:var(--muted); font-weight:650; }
    div[data-testid="stMetricValue"] { color:var(--ink); font-weight:760; }
    div[data-testid="stTabs"] button { font-weight:650; }
    div[data-testid="stDataFrame"] { border:1px solid var(--line); border-radius:10px; overflow:hidden; }
    .hint { color:var(--muted); font-size:.86rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
      <div class="hero-kicker">Report quality assurance</div>
      <div class="hero-title">Power BI Report Validator</div>
      <div class="hero-copy">Compare source and destination PBIX files across model objects and report pages.
      Add exported reports when you also want a visual and KPI comparison. Your existing validation engine
      and Excel report workflow are retained.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="section-label">01 · Upload reports</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2, gap="large")
with col1:
    with st.container(border=True):
        st.markdown("### ◉ Source report")
        st.caption("The original or baseline Power BI report")
        src_file = st.file_uploader("Power BI file (.pbix)", type=["pbix"], key="src_file")
        st.markdown("<div class='hint'>Optional visual comparison input</div>", unsafe_allow_html=True)
        src_visual = st.file_uploader(
            "Exported PDF or screenshot", type=["pdf", "png", "jpg", "jpeg"], key="src_visual"
        )
with col2:
    with st.container(border=True):
        st.markdown("### ◉ Destination report")
        st.caption("The migrated or updated Power BI report")
        dst_file = st.file_uploader("Power BI file (.pbix)", type=["pbix"], key="dst_file")
        st.markdown("<div class='hint'>Optional visual comparison input</div>", unsafe_allow_html=True)
        dst_visual = st.file_uploader(
            "Exported PDF or screenshot", type=["pdf", "png", "jpg", "jpeg"], key="dst_visual"
        )

left, right = st.columns([1, 3])
with left:
    run = st.button("▶  Run validation", type="primary", disabled=not (src_file and dst_file), use_container_width=True)
with right:
    if not (src_file and dst_file):
        st.info("Upload both PBIX files to enable validation.", icon="ℹ️")
    elif bool(src_visual) != bool(dst_visual):
        st.warning("For visual comparison, upload an export for both reports. Model comparison can run without them.", icon="⚠️")


def _save_upload(uploaded_file, tmp_dir):
    path = os.path.join(tmp_dir, uploaded_file.name)
    with open(path, "wb") as f:
        f.write(uploaded_file.getvalue())
    return path


def _highlight_status(df):
    if "Status" not in df.columns:
        return df.style

    def _row_style(row):
        color = "#ffe0e0" if row["Status"] in ("Mismatch", "Missing in Destination", "Extra in Destination") else ""
        return [f"background-color: {color}"] * len(row)

    return df.style.apply(_row_style, axis=1)


if run:
    tmp_dir = tempfile.mkdtemp(prefix="powerbi_validator_")
    try:
        with st.spinner("Parsing PBIX data models and report pages…"):
            src_path = _save_upload(src_file, tmp_dir)
            dst_path = _save_upload(dst_file, tmp_dir)
            src_meta = core.parse_workbook(src_path)
            dst_meta = core.parse_workbook(dst_path)
            diff = core.compare_metadata(src_meta, dst_meta)

        for meta, label in [(src_meta, "Source"), (dst_meta, "Destination")]:
            if meta.get("notes"):
                st.warning(f"{label}: {meta['notes']}")

        page_results = None
        if src_visual and dst_visual:
            with st.spinner("Running visual/KPI comparison…"):
                sv_path = _save_upload(src_visual, tmp_dir)
                dv_path = _save_upload(dst_visual, tmp_dir)
                page_results = core.run_visual_comparison(sv_path, dv_path, tmp_dir)

        with st.spinner("Building Excel report…"):
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = os.path.join(tmp_dir, f"powerbi_comparison_{ts}.xlsx")
            core.build_excel_report(src_meta, dst_meta, diff, out_path, page_results)
            with open(out_path, "rb") as f:
                report_bytes = f.read()

        st.success("Validation complete.")
        st.markdown('<div class="section-label">02 · Validation overview</div>', unsafe_allow_html=True)

        counts = {k: len(v) for k, v in diff.items()}
        if counts:
            metric_cols = st.columns(len(counts))
            for c, (k, n) in zip(metric_cols, counts.items()):
                c.metric(k.replace("_Mismatch", "").replace("_", " "), n)

        with st.expander("Parsed item counts · sanity check"):
            count_rows = []
            for label, meta in [("Source", src_meta), ("Destination", dst_meta)]:
                count_rows.append({
                    "File": label,
                    "Tables": len(meta["tables"]),
                    "Measures": len(meta["measures"]),
                    "Calc Columns": len(meta["calc_columns"]),
                    "Relationships": len(meta["relationships"]),
                    "Pages": len(meta["pages"]),
                })
            st.dataframe(pd.DataFrame(count_rows), use_container_width=True, hide_index=True)
            st.caption(
                "If these counts are 0 (or identical between files) but you know the reports differ, "
                "the difference may be visual/formatting-only. Use visual comparison for that."
            )

        st.download_button(
            "⬇  Download full Excel report",
            data=report_bytes,
            file_name=f"powerbi_comparison_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown('<div class="section-label">03 · Detailed results</div>', unsafe_allow_html=True)
        tab_map = {
            "Tables": "Table_Mismatch",
            "Columns": "Column_Mismatch",
            "Measures": "Measure_Mismatch",
            "Calc Columns": "CalcColumn_Mismatch",
            "Calc Tables": "CalcTable_Mismatch",
            "Relationships": "Relationship_Mismatch",
            "Power Query": "PowerQuery_Mismatch",
            "Parameters": "Parameter_Mismatch",
            "RLS": "RLS_Mismatch",
            "Pages": "Page_Mismatch",
        }
        tab_names = list(tab_map.keys())
        if page_results is not None:
            tab_names.append("Visual Comparison")
        tabs = st.tabs(tab_names)

        for tab, name in zip(tabs, tab_names):
            with tab:
                if name == "Visual Comparison":
                    for pr in page_results:
                        is_bad = pr["pixel_changed_pct"] > 2 or pr["text_status"] == "Mismatch"
                        badge = "🔴 MISMATCH" if is_bad else "🟢 MATCH"
                        with st.container(border=True):
                            st.markdown(f"#### {badge} · Page {pr['page']}")
                            a, b = st.columns(2)
                            a.metric("Pixel change", f"{pr['pixel_changed_pct']}%")
                            b.metric("Text status", pr["text_status"])
                            if pr.get("diff_img_path") and os.path.exists(pr["diff_img_path"]):
                                st.image(pr["diff_img_path"], caption="Pixel-difference image", use_container_width=True)
                else:
                    df = diff.get(tab_map[name], pd.DataFrame())
                    if df.empty:
                        st.success("No differences found in this category.")
                    else:
                        st.dataframe(_highlight_status(df), use_container_width=True, hide_index=True)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
