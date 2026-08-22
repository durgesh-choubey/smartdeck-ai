"""
SmartDeck AI -- turns the bundled Excel workbook into an updated PowerPoint
deck from the bundled template, in one click.

Deliberately simple, single-file app:
  1. Edit the data (Info / KPIs / Highlights / Chart) in-app, or in the Excel
     file directly and re-upload it.
  2. Click "Update PPT" -- regenerates the deck from the template.
  3. Download the .pptx, or switch to Fullscreen Preview to review the data
     and the resulting slide content at full width first.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from src.generator import DeckData, generate_pptx

ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = ROOT / "assets" / "template.pptx"
DEFAULT_DATA_PATH = ROOT / "data" / "sample_data.xlsx"

PRIMARY = "#0078D4"

st.set_page_config(page_title="SmartDeck AI · PULSE", page_icon="📽️", layout="wide")

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.5rem; max-width: 1200px; }
    .sd-card { background:#FFFFFF; border:1px solid #E5E7EB; border-radius:14px;
               padding:1.1rem 1.4rem; box-shadow:0 1px 2px rgba(0,0,0,0.04), 0 4px 12px rgba(0,0,0,0.04);
               margin-bottom:1rem; }
    .sd-header-title { font-size:1.4rem; font-weight:700; color:#323130; margin:0; }
    .sd-header-sub { font-size:0.85rem; color:#605E5C; margin:0.15rem 0 0 0; }
    .sd-slide-card { background:#FFFFFF; border:1px solid #E5E7EB; border-radius:12px;
                      padding:1.2rem 1.4rem; margin-bottom:0.9rem; }
    .sd-slide-label { font-size:0.72rem; font-weight:700; color:#0078D4; text-transform:uppercase;
                       letter-spacing:0.04em; margin-bottom:0.3rem; }
    .sd-slide-title { font-size:1.15rem; font-weight:700; color:#323130; margin:0 0 0.5rem 0; }
    .sd-slide-body { font-size:0.95rem; color:#323130; white-space:pre-line; line-height:1.6; }
    .stButton > button { background:#0078D4 !important; color:#FFFFFF !important;
                          border-radius:8px !important; border:none !important; font-weight:600 !important; }
    .stDownloadButton > button { background:#107C10 !important; color:#FFFFFF !important;
                                  border-radius:8px !important; border:none !important; font-weight:600 !important; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
def _load_default_dataframes() -> dict[str, pd.DataFrame]:
    return {
        "Info": pd.read_excel(DEFAULT_DATA_PATH, sheet_name="Info"),
        "KPIs": pd.read_excel(DEFAULT_DATA_PATH, sheet_name="KPIs"),
        "Highlights": pd.read_excel(DEFAULT_DATA_PATH, sheet_name="Highlights"),
        "Chart": pd.read_excel(DEFAULT_DATA_PATH, sheet_name="Chart", skiprows=2),
    }


def _blank_dataframes() -> dict[str, pd.DataFrame]:
    """Field names stay visible (so it's clear what each row is for) but every
    value is empty -- this is the app's starting state, so the first thing a
    POC audience sees is a genuinely blank template, not sample numbers."""
    info_fields = ["report_title", "report_subtitle", "company", "date", "prepared_by"]
    return {
        "Info": pd.DataFrame({"Field": info_fields, "Value": [""] * len(info_fields)}),
        "KPIs": pd.DataFrame({"Metric": pd.Series(dtype="str"), "Value": pd.Series(dtype="str")}),
        "Highlights": pd.DataFrame({"Highlight": pd.Series(dtype="str")}),
        "Chart": pd.DataFrame({"Category": pd.Series(dtype="str"), "Value": pd.Series(dtype="float")}),
    }


if "sd_dataframes" not in st.session_state:
    st.session_state.sd_dataframes = _blank_dataframes()
if "sd_chart_title" not in st.session_state:
    st.session_state.sd_chart_title = ""
if "sd_fullscreen" not in st.session_state:
    st.session_state.sd_fullscreen = False
if "sd_preview_data" not in st.session_state:
    # What the Preview section renders. Deliberately NOT recomputed from the
    # live editors on every rerun -- it only changes when Update PPT (or
    # Reset to Blank) is clicked, so the preview always matches the actual
    # generated .pptx rather than showing in-progress, not-yet-applied edits.
    st.session_state.sd_preview_data = DeckData()
if "sd_version" not in st.session_state:
    # Bumped on every Update/Reset. Used as part of the download button's
    # widget key so Streamlit treats it as a brand-new widget each time the
    # underlying data changes, rather than reusing a previous widget
    # instance that could otherwise serve a stale cached file.
    st.session_state.sd_version = 0
if "sd_data_version" not in st.session_state:
    # Bumped whenever sd_dataframes is set PROGRAMMATICALLY (upload, Load
    # Sample Data, Reset to Blank) rather than via direct user edits. This
    # is essential: st.data_editor widgets remember their own committed
    # state once a `key` is set, and silently ignore a new dataframe passed
    # in on a later rerun -- worse, they hand the OLD content straight back,
    # overwriting the fresh data you just tried to load. Folding this
    # counter into each editor's `key=` forces a genuinely new widget
    # instance whenever we set data from outside the widget, so the new
    # data actually shows up (and doesn't get silently reverted).
    st.session_state.sd_data_version = 0


def _dataframes_to_deck_data() -> DeckData:
    dfs = st.session_state.sd_dataframes
    info = {str(r["Field"]).strip(): str(r["Value"]) for _, r in dfs["Info"].iterrows() if pd.notna(r["Field"])}
    kpis = [(str(r["Metric"]).strip(), str(r["Value"])) for _, r in dfs["KPIs"].iterrows() if pd.notna(r["Metric"])]
    highlights = [str(r["Highlight"]).strip() for _, r in dfs["Highlights"].iterrows() if pd.notna(r["Highlight"])]
    chart_df = dfs["Chart"].dropna(subset=["Category", "Value"])
    return DeckData(
        info=info,
        kpis=kpis,
        highlights=highlights,
        chart_title=st.session_state.sd_chart_title,
        chart_categories=[str(c) for c in chart_df["Category"].tolist()],
        chart_values=[float(v) for v in chart_df["Value"].tolist()],
    )


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown(
    """
    <div class="sd-card" style="display:flex; align-items:center; justify-content:space-between;">
        <div>
            <p class="sd-header-title">📽️ SmartDeck AI</p>
            <p class="sd-header-sub">AI PowerPoint Generator — PULSE Automation Platform</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

col_toggle, _ = st.columns([1, 4])
with col_toggle:
    st.session_state.sd_fullscreen = st.toggle("🔎 Fullscreen Preview", value=st.session_state.sd_fullscreen)


# ---------------------------------------------------------------------------
# Edit data (hidden while in Fullscreen Preview, to keep that view uncluttered)
# ---------------------------------------------------------------------------
if not st.session_state.sd_fullscreen:
    with st.container(border=True):
        st.markdown("**Excel Data Source**")

        col_upload, col_sample = st.columns([3, 1])
        with col_upload:
            uploaded = st.file_uploader("Replace with your own workbook (optional)", type=["xlsx"])
        with col_sample:
            st.write("")  # vertical alignment spacer
            if st.button("Load Sample Data", use_container_width=True):
                st.session_state.sd_dataframes = _load_default_dataframes()
                st.session_state.sd_chart_title = "Revenue by Region ($K)"
                st.session_state.sd_data_version += 1
                st.rerun()

        if uploaded is not None:
            try:
                st.session_state.sd_dataframes = {
                    "Info": pd.read_excel(uploaded, sheet_name="Info"),
                    "KPIs": pd.read_excel(uploaded, sheet_name="KPIs"),
                    "Highlights": pd.read_excel(uploaded, sheet_name="Highlights"),
                    "Chart": pd.read_excel(uploaded, sheet_name="Chart", skiprows=2),
                }
                st.session_state.sd_data_version += 1
                st.success("Workbook loaded. Review the tabs below, then click Update PPT.")
            except Exception as exc:
                st.error(f"Couldn't read that workbook -- make sure it has Info/KPIs/Highlights/Chart tabs. ({exc})")

        v = st.session_state.sd_data_version
        tab_info, tab_kpis, tab_highlights, tab_chart = st.tabs(["Info", "KPIs", "Highlights", "Chart"])
        with tab_info:
            st.session_state.sd_dataframes["Info"] = st.data_editor(
                st.session_state.sd_dataframes["Info"], num_rows="fixed", use_container_width=True, key=f"editor_info_{v}"
            )
        with tab_kpis:
            st.session_state.sd_dataframes["KPIs"] = st.data_editor(
                st.session_state.sd_dataframes["KPIs"], num_rows="dynamic", use_container_width=True, key=f"editor_kpis_{v}"
            )
        with tab_highlights:
            st.session_state.sd_dataframes["Highlights"] = st.data_editor(
                st.session_state.sd_dataframes["Highlights"], num_rows="dynamic", use_container_width=True, key=f"editor_highlights_{v}"
            )
        with tab_chart:
            st.session_state.sd_chart_title = st.text_input("Chart title", value=st.session_state.sd_chart_title, key=f"chart_title_{v}")
            st.session_state.sd_dataframes["Chart"] = st.data_editor(
                st.session_state.sd_dataframes["Chart"], num_rows="dynamic", use_container_width=True, key=f"editor_chart_{v}"
            )

    col_update, col_download, col_reset = st.columns([1, 1, 1])
    with col_update:
        if st.button("🔄 Update PPT", type="primary", use_container_width=True):
            data = _dataframes_to_deck_data()
            st.session_state.sd_preview_data = data
            st.session_state.sd_version += 1
            st.success("PPT updated from the current data.")
    with col_download:
        # Always regenerated fresh, right here, directly from sd_preview_data --
        # the exact same source of truth the Preview section below uses. No
        # "only if None" shortcut and no reliance on a value set earlier in a
        # previous run, so the two can never drift apart.
        current_pptx = generate_pptx(str(TEMPLATE_PATH), st.session_state.sd_preview_data)
        st.download_button(
            "⬇ Download PPT",
            data=current_pptx,
            file_name="smartdeck_report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            key=f"download_btn_v{st.session_state.sd_version}",
        )
    with col_reset:
        if st.button("🧹 Reset to Blank", use_container_width=True):
            st.session_state.sd_dataframes = _blank_dataframes()
            st.session_state.sd_chart_title = ""
            st.session_state.sd_preview_data = DeckData()
            st.session_state.sd_version += 1
            st.session_state.sd_data_version += 1
            st.rerun()


# ---------------------------------------------------------------------------
# Preview -- only reflects the last generated PPT (see sd_preview_data above),
# not whatever's currently being edited in the tabs. Full width in Fullscreen
# Preview mode.
# ---------------------------------------------------------------------------
data = st.session_state.sd_preview_data
is_blank = not (data.info or data.kpis or data.highlights or data.chart_categories)

st.markdown("### Preview")
st.caption("Reflects the last generated .pptx -- click Update PPT to refresh it with your current edits.")


def _slide_card(label: str, title: str, body: str) -> None:
    st.markdown(
        f"""
        <div class="sd-slide-card">
            <div class="sd-slide-label">{label}</div>
            <div class="sd-slide-title">{title}</div>
            <div class="sd-slide-body">{body}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


if is_blank:
    st.info("Nothing generated yet. Fill in the data above (or click **Load Sample Data**) and click **Update PPT** to see a preview here.")
elif st.session_state.sd_fullscreen:
    _slide_card("Slide 1 · Title", data.info.get("report_title", ""), data.info.get("report_subtitle", ""))
    _slide_card("Slide 2 · Key Metrics", "Key Metrics", data.kpi_list_text().replace("\n", "<br>"))
    _slide_card("Slide 3 · Highlights", "Highlights", data.highlights_list_text().replace("\n", "<br>"))
    if data.chart_categories:
        fig = px.bar(x=data.chart_categories, y=data.chart_values, labels={"x": "", "y": ""},
                      title=data.chart_title, color_discrete_sequence=[PRIMARY])
        fig.update_layout(height=420, margin=dict(l=10, r=10, t=50, b=10), paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF")
        st.plotly_chart(fig, use_container_width=True)
    st.markdown("#### Excel Data")
    for name, df in st.session_state.sd_dataframes.items():
        st.caption(name)
        st.dataframe(df, use_container_width=True)
else:
    preview_col1, preview_col2 = st.columns(2)
    with preview_col1:
        _slide_card("Slide 1 · Title", data.info.get("report_title", ""), data.info.get("report_subtitle", ""))
        _slide_card("Slide 2 · Key Metrics", "Key Metrics", data.kpi_list_text().replace("\n", "<br>"))
    with preview_col2:
        _slide_card("Slide 3 · Highlights", "Highlights", data.highlights_list_text().replace("\n", "<br>"))
        if data.chart_categories:
            fig = px.bar(x=data.chart_categories, y=data.chart_values, labels={"x": "", "y": ""},
                          title=data.chart_title, color_discrete_sequence=[PRIMARY])
            fig.update_layout(height=280, margin=dict(l=10, r=10, t=40, b=10), paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF")
            st.plotly_chart(fig, use_container_width=True)