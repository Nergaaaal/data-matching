"""
Data Matching — Match records across multiple files.
"""

import io
import csv
import html as html_lib
from typing import List, Optional

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows


# ── Page Config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Data Matching",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

    .stApp {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        background: linear-gradient(145deg, #f0f2f5 0%, #e8ecf1 50%, #f0f2f5 100%);
    }

    /* ── Brand header ── */
    .brand {
        text-align: center;
        padding: 1.8rem 0 1.2rem;
    }
    .brand h1 {
        font-size: 1.8rem;
        font-weight: 800;
        color: #1a1a2e;
        letter-spacing: -0.03em;
        margin: 0;
    }
    .brand-line {
        width: 40px;
        height: 3px;
        background: linear-gradient(90deg, #8b5cf6, #3b82f6);
        border-radius: 2px;
        margin: 0.5rem auto 0;
    }

    /* ── Glass card ── */
    .glass-card {
        background: rgba(255,255,255,0.85);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(255,255,255,0.6);
        border-radius: 16px;
        padding: 1.5rem;
        margin-bottom: 1.2rem;
        box-shadow: 0 4px 24px rgba(0,0,0,0.04);
    }

    /* ── File info row ── */
    .file-info-row {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        padding: 0.4rem 0;
    }
    .fi-dot {
        width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0;
    }
    .fi-dot-a { background: #8b5cf6; }
    .fi-dot-b { background: #3b82f6; }
    .fi-dot-x { background: #10b981; }
    .fi-name {
        font-weight: 600; font-size: 0.85rem; color: #1f2937;
    }
    .fi-meta {
        font-size: 0.78rem; color: #9ca3af; white-space: nowrap;
    }

    /* ── Loaded banner ── */
    .loaded-banner {
        background: linear-gradient(135deg, #f0fdf4, #dcfce7);
        border: 1px solid #bbf7d0;
        border-radius: 10px;
        padding: 0.7rem 1rem;
        font-size: 0.85rem;
        color: #166534;
        font-weight: 500;
        margin-bottom: 0.8rem;
        display: flex;
        align-items: center;
        gap: 0.4rem;
    }

    /* ── Stats ── */
    .stats-row {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 1rem;
        margin: 1.5rem 0;
    }
    .stat-box {
        background: rgba(255,255,255,0.9);
        border: 1px solid rgba(255,255,255,0.6);
        border-radius: 14px;
        padding: 1.4rem 1rem;
        text-align: center;
        box-shadow: 0 2px 12px rgba(0,0,0,0.03);
        transition: all 0.3s ease;
    }
    .stat-box:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0,0,0,0.07);
    }
    .stat-icon { font-size: 1.4rem; margin-bottom: 0.3rem; }
    .stat-value {
        font-size: 2.2rem; font-weight: 800;
        letter-spacing: -0.02em; line-height: 1.1;
    }
    .stat-label {
        font-size: 0.75rem; color: #9ca3af;
        text-transform: uppercase; letter-spacing: 0.08em;
        margin-top: 0.3rem; font-weight: 600;
    }
    .c-dark  { color: #1a1a2e; }
    .c-green { color: #059669; }
    .c-red   { color: #ef4444; }
    .c-blue  { color: #3b82f6; }

    /* ── Separator ── */
    .sep {
        height: 1px;
        background: linear-gradient(90deg, transparent, #d1d5db, transparent);
        margin: 1.2rem 0;
    }

    /* ── Section label ── */
    .section-label {
        font-size: 0.95rem; font-weight: 700; color: #374151;
        margin-bottom: 0.8rem;
    }

    /* ── Footer ── */
    .footer {
        text-align: center; padding: 2rem 0 1rem;
        color: #c9cdd3; font-size: 0.78rem;
    }

    /* ── Streamlit overrides ── */
    .stDownloadButton > button,
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #1a1a2e 0%, #2d2d50 100%) !important;
        color: #fff !important; border: none !important;
        border-radius: 12px !important; padding: 0.75rem 2rem !important;
        font-weight: 700 !important; font-size: 0.95rem !important;
        box-shadow: 0 4px 14px rgba(26,26,46,0.2) !important;
        transition: all 0.3s cubic-bezier(0.4,0,0.2,1) !important;
    }
    .stDownloadButton > button:hover,
    .stButton > button[kind="primary"]:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 24px rgba(26,26,46,0.35) !important;
    }
    .stButton > button:not([kind="primary"]) {
        border-radius: 10px !important; font-weight: 600 !important;
    }
    div[data-testid="stExpander"] {
        border: 1px solid rgba(0,0,0,0.05) !important;
        border-radius: 12px !important;
        background: rgba(255,255,255,0.5);
    }
    .stSelectbox > div > div { border-radius: 10px !important; }
    .stTextArea textarea {
        border-radius: 10px !important;
        font-family: 'Inter', monospace !important;
    }
    div[data-testid="stDataFrame"] {
        border-radius: 12px; overflow: hidden;
        box-shadow: 0 2px 12px rgba(0,0,0,0.04);
    }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)


# ── Helpers ─────────────────────────────────────────────────────────────────

def detect_delimiter(uploaded_file) -> str:
    uploaded_file.seek(0)
    sample = uploaded_file.read(8192)
    uploaded_file.seek(0)
    if isinstance(sample, bytes):
        sample = sample.decode("utf-8", errors="ignore")
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        return dialect.delimiter
    except csv.Error:
        return ","


def read_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    try:
        if name.endswith(".csv"):
            delimiter = detect_delimiter(uploaded_file)
            uploaded_file.seek(0)
            try:
                df = pd.read_csv(uploaded_file, sep=delimiter)
            except Exception:
                uploaded_file.seek(0)
                try:
                    df = pd.read_csv(uploaded_file, sep=delimiter, on_bad_lines="skip")
                except TypeError:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, sep=delimiter, error_bad_lines=False)
            if len(df.columns) == 1:
                col_name = df.columns[0]
                for alt_sep in [";", "\t", "|", ","]:
                    if alt_sep == delimiter:
                        continue
                    pat = alt_sep if alt_sep != "|" else "\\|"
                    if df[col_name].astype(str).str.contains(pat, na=False).mean() > 0.5:
                        uploaded_file.seek(0)
                        try:
                            df2 = pd.read_csv(uploaded_file, sep=alt_sep, on_bad_lines="skip")
                        except TypeError:
                            uploaded_file.seek(0)
                            df2 = pd.read_csv(uploaded_file, sep=alt_sep, error_bad_lines=False)
                        if len(df2.columns) > 1:
                            return df2
            return df
        elif name.endswith((".xlsx", ".xls")):
            return pd.read_excel(uploaded_file)
        else:
            st.error(f"Unsupported format: {uploaded_file.name}")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Error reading **{uploaded_file.name}**: {e}")
        return pd.DataFrame()


def normalize(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().str.lower()


def perform_matching(df_a, col_a, df_b, col_b, manual_ids=None):
    a = df_a[[col_a]].copy()
    a["_key"] = normalize(a[col_a])
    b = df_b[[col_b]].copy()
    b["_key"] = normalize(b[col_b])
    if manual_ids:
        manual_df = pd.DataFrame({"_key": [s.strip().lower() for s in manual_ids], col_b: manual_ids})
        b = pd.concat([b, manual_df], ignore_index=True)
    merged = pd.merge(
        a.drop_duplicates(subset="_key"),
        b.drop_duplicates(subset="_key"),
        on="_key", how="outer", suffixes=("_A", "_B"),
    )
    result = pd.DataFrame()
    result["File A Value"] = merged[col_a].where(merged[col_a].notna(), "—")
    result["File B Value"] = merged[col_b].where(merged[col_b].notna(), "—")
    result["Match"] = ((merged[col_a].notna()) & (merged[col_b].notna())).astype(int)
    return result.sort_values("Match", ascending=False).reset_index(drop=True)


def generate_styled_xlsx(df, match_pct):
    wb = Workbook()
    ws = wb.active
    ws.title = "Match Results"
    hf = PatternFill(start_color="1A1A2E", end_color="1A1A2E", fill_type="solid")
    hfont = Font(name="Inter", bold=True, color="FFFFFF", size=11)
    mf = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    nf = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
    dfont = Font(name="Inter", size=10, color="1F2937")
    tb = Border(bottom=Side(style="thin", color="E5E7EB"))
    for ri, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            if ri == 1:
                cell.font, cell.fill = hfont, hf
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.font, cell.border = dfont, tb
                if ci == 3:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.fill = mf if val == 1 else nf
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 35
    ws.column_dimensions["C"].width = 14
    sr = len(df) + 3
    ws.cell(row=sr, column=1, value="Match Rate:").font = Font(name="Inter", bold=True, size=11, color="1A1A2E")
    ws.cell(row=sr, column=2, value=f"{match_pct:.1f}%").font = Font(name="Inter", bold=True, size=11, color="059669")
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def render_file_chips(files_info, compact=False):
    """Render file info using native Streamlit — no HTML leaks."""
    dot_cls = ["fi-dot-a", "fi-dot-b"] + ["fi-dot-x"] * 10
    for idx, fi in enumerate(files_info):
        d = dot_cls[idx] if idx < len(dot_cls) else "fi-dot-x"
        name = html_lib.escape(fi["name"])
        disp = name if len(name) <= 40 else name[:37] + "…"
        meta = f"{fi['rows']:,} rows" if compact else f"{fi['rows']:,} rows · {fi['cols']} cols"
        st.markdown(
            f'<div class="file-info-row">'
            f'<div class="fi-dot {d}"></div>'
            f'<span class="fi-name" title="{name}">{disp}</span>'
            f'<span class="fi-meta">{meta}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )


# ══════════════════════════════════════════════════════════════════════════════
#                              STATE
# ══════════════════════════════════════════════════════════════════════════════

if "phase" not in st.session_state:
    st.session_state.phase = "upload"


# ── Brand header (always visible) ───────────────────────────────────────────
st.markdown("""
<div class="brand">
    <h1>Data Matching</h1>
    <div class="brand-line"></div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#                         PHASE 1: UPLOAD
# ══════════════════════════════════════════════════════════════════════════════

if st.session_state.phase == "upload":

    # Check if we already have stored data (coming back from configure/results)
    has_stored = "df_a" in st.session_state and "df_b" in st.session_state

    if has_stored:
        st.markdown(
            '<div class="loaded-banner">✅ Files already loaded. You can re-upload or continue with existing data.</div>',
            unsafe_allow_html=True,
        )
        render_file_chips(st.session_state.all_files_info)
        st.markdown('<div class="sep"></div>', unsafe_allow_html=True)

    # File uploaders
    col_a, col_b = st.columns(2)
    with col_a:
        file_a = st.file_uploader("File A", type=["csv", "xlsx", "xls"], key="file_a")
    with col_b:
        file_b = st.file_uploader("File B", type=["csv", "xlsx", "xls"], key="file_b")

    # Additional files
    with st.expander("➕ Upload additional files (up to 5)", expanded=False):
        extra_files_list = []
        ecols = st.columns(3)
        for i in range(5):
            with ecols[i % 3]:
                ef = st.file_uploader(
                    f"Extra {i+1}", type=["csv", "xlsx", "xls"],
                    key=f"extra_{i}", label_visibility="collapsed",
                )
                if ef:
                    extra_files_list.append(ef)

    st.markdown('<div class="sep"></div>', unsafe_allow_html=True)

    # Manual IDs
    st.markdown(
        '<div class="section-label">Manual IDs '
        '<span style="color:#9ca3af; font-weight:400; font-size:0.82rem;">(optional)</span></div>',
        unsafe_allow_html=True,
    )
    mc1, mc2 = st.columns([4, 1])
    with mc1:
        default_manual = "\n".join(st.session_state.get("manual_ids", []))
        manual_text = st.text_area(
            "Enter IDs", height=110, placeholder="One ID per line…",
            key="manual_ids_area", label_visibility="collapsed",
            value=default_manual if has_stored and default_manual else "",
        )
    with mc2:
        parsed_ids: List[str] = []
        if manual_text and manual_text.strip():
            parsed_ids = [l.strip() for l in manual_text.strip().splitlines() if l.strip()]
            st.metric("IDs", len(parsed_ids))

    st.markdown('<div class="sep"></div>', unsafe_allow_html=True)

    # -- Proceed logic --
    new_files_uploaded = file_a is not None and file_b is not None

    if new_files_uploaded or has_stored:
        if st.button("▶  Continue to Matching", use_container_width=True, type="primary"):
            # If new files uploaded, re-read them
            if new_files_uploaded:
                df_a = read_file(file_a)
                df_b = read_file(file_b)
                all_info = [
                    {"name": file_a.name, "rows": len(df_a), "cols": len(df_a.columns)},
                    {"name": file_b.name, "rows": len(df_b), "cols": len(df_b.columns)},
                ]
                for ef in extra_files_list:
                    df_extra = read_file(ef)
                    if not df_extra.empty:
                        df_b = pd.concat([df_b, df_extra], ignore_index=True)
                        all_info.append({"name": ef.name, "rows": len(df_extra), "cols": len(df_extra.columns)})

                if df_a.empty or df_b.empty:
                    st.error("⚠️ Could not read one of the files.")
                    st.stop()

                st.session_state.df_a = df_a
                st.session_state.df_b = df_b
                st.session_state.all_files_info = all_info
                st.session_state.file_a_name = file_a.name
                st.session_state.file_b_name = file_b.name

            # Store manual IDs
            st.session_state.manual_ids = parsed_ids
            st.session_state.phase = "configure"
            st.rerun()
    else:
        st.markdown("""
        <div style="text-align:center; padding:2rem; color:#b0b5bc;">
            <div style="font-size:2.5rem; opacity:0.4; margin-bottom:0.3rem;">📂</div>
            <p style="font-size:0.9rem; font-weight:500; color:#9ca3af;">Upload at least two files to start</p>
        </div>
        """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#                      PHASE 2: CONFIGURE
# ══════════════════════════════════════════════════════════════════════════════

elif st.session_state.phase == "configure":

    df_a = st.session_state.df_a
    df_b = st.session_state.df_b

    if st.button("← Back"):
        st.session_state.phase = "upload"
        st.rerun()

    render_file_chips(st.session_state.all_files_info)

    st.markdown('<div class="sep"></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-label">Select matching columns</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        sel_a = st.selectbox(
            f"Column from {st.session_state.file_a_name}",
            options=df_a.columns.tolist(), key="col_a_sel",
        )
    with c2:
        sel_b = st.selectbox(
            f"Column from {st.session_state.file_b_name}",
            options=df_b.columns.tolist(), key="col_b_sel",
        )

    with st.expander("👀 Preview", expanded=False):
        p1, p2 = st.columns(2)
        with p1:
            st.dataframe(df_a[[sel_a]].head(8), use_container_width=True, hide_index=True)
        with p2:
            st.dataframe(df_b[[sel_b]].head(8), use_container_width=True, hide_index=True)

    st.markdown('<div class="sep"></div>', unsafe_allow_html=True)

    if st.button("▶  Run Match", use_container_width=True, type="primary"):
        with st.spinner("Matching…"):
            manual = st.session_state.get("manual_ids", [])
            result = perform_matching(df_a, sel_a, df_b, sel_b, manual or None)
            st.session_state.result_df = result
            st.session_state.phase = "results"
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
#                          PHASE 3: RESULTS
# ══════════════════════════════════════════════════════════════════════════════

elif st.session_state.phase == "results":

    result_df = st.session_state.result_df

    if st.button("← Back"):
        st.session_state.phase = "configure"
        st.session_state.pop("result_df", None)
        st.rerun()

    render_file_chips(st.session_state.all_files_info, compact=True)

    total = len(result_df)
    matched = int(result_df["Match"].sum())
    unmatched = total - matched
    pct = (matched / total * 100) if total else 0
    pc = "c-green" if pct >= 70 else ("c-blue" if pct >= 40 else "c-red")

    st.markdown(f"""
    <div class="stats-row">
        <div class="stat-box">
            <div class="stat-icon">📊</div>
            <div class="stat-value c-dark">{total:,}</div>
            <div class="stat-label">Total Records</div>
        </div>
        <div class="stat-box">
            <div class="stat-icon">✅</div>
            <div class="stat-value c-green">{matched:,}</div>
            <div class="stat-label">Matched</div>
        </div>
        <div class="stat-box">
            <div class="stat-icon">❌</div>
            <div class="stat-value c-red">{unmatched:,}</div>
            <div class="stat-label">Unmatched</div>
        </div>
        <div class="stat-box">
            <div class="stat-icon">🎯</div>
            <div class="stat-value {pc}">{pct:.1f}%</div>
            <div class="stat-label">Match Rate</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.dataframe(
        result_df.style.applymap(
            lambda v: "background-color:#ecfdf5;color:#059669;font-weight:600;"
            if v == 1
            else ("background-color:#fef2f2;color:#ef4444;font-weight:600;" if v == 0 else ""),
            subset=["Match"],
        ),
        use_container_width=True, hide_index=True, height=460,
    )

    xlsx = generate_styled_xlsx(result_df, pct)
    st.download_button(
        "⬇  Download Match Report (.xlsx)", data=xlsx,
        file_name="match_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


st.markdown('<div class="footer">Data Matching · All processing happens locally</div>', unsafe_allow_html=True)
