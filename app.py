"""
app.py — Streamlit UI for the Ledger Excel Formatter.
Corporate formal design — clean, professional, light theme.
"""

import io
from pathlib import Path

import pandas as pd
import streamlit as st

from daybook_parser import parse_daybook, load_ledger_set

# ─── Page config ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Daybook Analyser | SGV",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

LEDGER_PATH = Path(__file__).parent / "ledger.txt"

# ─── Corporate CSS ───────────────────────────────────────────────────────────

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    /* ── Global ── */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        color: #1c1c1c;
    }
    .stApp {
        background-color: #f5f5f5;
    }

    /* ── Top header bar ── */
    .top-bar {
        background: #1c1c1c;
        padding: 0 28px;
        height: 60px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin: -1rem -1rem 0 -1rem;
        border-bottom: 4px solid #c0392b;
    }
    .top-bar-right {
        font-size: 0.72rem;
        color: #aaaaaa;
        letter-spacing: 0.06em;
        text-transform: uppercase;
    }
    .top-bar-title {
        font-size: 0.85rem;
        font-weight: 600;
        color: #ffffff;
        letter-spacing: 0.08em;
        text-transform: uppercase;
    }

    /* ── Page heading ── */
    .page-heading {
        padding: 22px 0 14px 0;
        border-bottom: 2px solid #c0392b;
        margin-bottom: 20px;
    }
    .page-heading h1 {
        font-size: 1.5rem;
        font-weight: 700;
        color: #1c1c1c;
        margin: 0 0 3px 0;
        letter-spacing: -0.01em;
    }
    .page-heading p {
        font-size: 0.82rem;
        color: #666666;
        margin: 0;
    }

    /* ── Metric strip ── */
    .metric-strip {
        display: flex;
        gap: 12px;
        margin: 16px 0 20px 0;
        flex-wrap: wrap;
    }
    .metric-card {
        flex: 1;
        min-width: 140px;
        background: #ffffff;
        border: 1px solid #ddd;
        border-top: 3px solid #1c1c1c;
        border-radius: 4px;
        padding: 14px 18px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.07);
    }
    .metric-card.green { border-top-color: #c0392b; }
    .metric-card.red   { border-top-color: #7b0d0d; }
    .metric-card .m-label {
        font-size: 0.67rem;
        font-weight: 600;
        color: #888888;
        text-transform: uppercase;
        letter-spacing: 0.07em;
    }
    .metric-card .m-value {
        font-size: 1.35rem;
        font-weight: 700;
        color: #1c1c1c;
        margin-top: 4px;
        line-height: 1.1;
    }

    /* ── Filter panel ── */
    .filter-panel {
        background: #ffffff;
        border: 1px solid #ddd;
        border-left: 4px solid #c0392b;
        border-radius: 4px;
        padding: 16px 20px;
        margin-bottom: 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .section-label {
        font-size: 0.7rem;
        font-weight: 700;
        color: #1c1c1c;
        text-transform: uppercase;
        letter-spacing: 0.09em;
        margin-bottom: 10px;
        padding-bottom: 8px;
        border-bottom: 1px solid #eeeeee;
    }

    /* ── Table container ── */
    .table-container {
        background: #ffffff;
        border: 1px solid #ddd;
        border-radius: 4px;
        padding: 16px 20px 8px 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background: #ffffff !important;
        border-right: 1px solid #ddd !important;
    }
    [data-testid="stSidebar"] .stMarkdown h2 {
        font-size: 0.82rem !important;
        color: #1c1c1c !important;
        text-transform: uppercase;
        letter-spacing: 0.07em;
        font-weight: 700;
    }

    /* ── Form labels ── */
    .stSelectbox label, .stMultiSelect label, .stFileUploader label {
        font-size: 0.73rem !important;
        font-weight: 600 !important;
        color: #444444 !important;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    [data-testid="stFileUploader"] {
        background: #fff8f8;
        border: 1.5px dashed #c0392b;
        border-radius: 4px;
        padding: 8px;
    }

    /* ── Buttons ── */
    .stButton > button {
        background: #1c1c1c;
        color: #ffffff;
        border: none;
        border-radius: 3px;
        font-size: 0.78rem;
        font-weight: 600;
        padding: 0.45rem 1.1rem;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        transition: background 0.15s;
    }
    .stButton > button:hover { background: #c0392b; }

    .stDownloadButton > button {
        background: #ffffff;
        color: #1c1c1c !important;
        border: 1.5px solid #1c1c1c;
        border-radius: 3px;
        font-size: 0.76rem;
        font-weight: 600;
        padding: 0.4rem 1rem;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        transition: all 0.15s;
    }
    .stDownloadButton > button:hover {
        background: #c0392b;
        color: #fff !important;
        border-color: #c0392b;
    }

    /* ── Misc ── */
    hr { border-color: #dddddd; }
    [data-testid="stInfo"] {
        background: #fff5f5;
        border-left: 4px solid #c0392b;
        color: #7b0d0d;
        border-radius: 0;
    }
    [data-testid="stExpander"] {
        background: #ffffff;
        border: 1px solid #dddddd;
        border-radius: 4px;
    }
    .block-container { padding-top: 0.5rem !important; }

    /* ── Badges ── */
    .badge-cc {
        display:inline-block; background:#DEEAF1; color:#1f4e79;
        font-size:0.65rem; font-weight:700; padding:2px 10px;
        border-radius:2px; letter-spacing:0.05em; text-transform:uppercase;
        border: 1px solid #9dc3e6;
    }
    .badge-cr {
        display:inline-block; background:#FFF2CC; color:#7f6000;
        font-size:0.65rem; font-weight:700; padding:2px 10px;
        border-radius:2px; letter-spacing:0.05em; text-transform:uppercase;
        border: 1px solid #ffd966;
    }
    /* Divider line under logo header */
    .logo-divider {
        height: 4px;
        background: #c0392b;
        margin: 0 -1rem 0 -1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─── Navigation bar ──────────────────────────────────────────────────────────

# ─── Top bar ─────────────────────────────────────────────────────────────────

LOGO_PATH = Path(__file__).parent / "logo.png"

st.markdown(
    '<div class="top-bar">'
    '<span class="top-bar-title">Daybook Ledger Analyser</span>'
    '<span class="top-bar-right">Finance &amp; Accounts &nbsp;|&nbsp; Internal Use Only</span>'
    '</div>',
    unsafe_allow_html=True,
)

# ─── Logo + Page Heading ──────────────────────────────────────────────────────

_logo_col, _title_col = st.columns([1, 7])
with _logo_col:
    if LOGO_PATH.exists():
        st.markdown(
            "<div style='display:flex;justify-content:center;align-items:center;height:100%;padding-top:12px;'>",
            unsafe_allow_html=True,
        )
        st.image(str(LOGO_PATH), width=88)
        st.markdown("</div>", unsafe_allow_html=True)
with _title_col:
    st.markdown(
        """
        <div class="page-heading">
            <h1>Daybook Ledger Report</h1>
            <p>Normalised transaction data &mdash; Cost Centres &amp; Creditors extracted from Tally daybook export</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ─── Colour Scheme Definitions ────────────────────────────────────────────────
# Each entry: (label, cc_web, cr_web, cc_excel, cr_excel)
# Dark brand colours are converted to light readable tints for row backgrounds.
COLOUR_SCHEMES = {
    "SGV Brand (Default)":      ("SGV Brand",    "#fde8e8", "#f0eeee", "#F4CCCC", "#D9D9D9"),
    "Charcoal / Steel Blue":    ("Charcoal/Blue","#D6E0EF", "#E0E0E0", "#BBCCE3", "#D0D0D0"),
    "Deep Teal / Ice Blue":     ("Teal/Ice",     "#D6EAE9", "#E6F2F1", "#C0DFDF", "#D4ECF0"),
    "Burgundy / Warm Grey":     ("Burg./Grey",   "#EEDDE0", "#EFEDEE", "#E0C0C5", "#EFEDEE"),
    "Forest Green / Pale Sage": ("Forest/Sage",  "#D6E8DA", "#E8F0EB", "#C0D9C5", "#D5E8DC"),
    "No Colour":                ("No Colour",    "",        "",        "#FFFFFF", "#FFFFFF"),
}

# ─── Sidebar ───────────────────────────────────────────────────────────

with st.sidebar:
    # ─ Colour Scheme ─────────────────────────────────────────
    st.markdown("## Row Colour Scheme")
    scheme_key = st.selectbox(
        "Pick a colour theme",
        options=list(COLOUR_SCHEMES.keys()),
        index=0,
        key="colour_scheme",
        label_visibility="collapsed",
    )
    _scheme = COLOUR_SCHEMES[scheme_key]
    _label, _cc_web, _cr_web, _cc_xl, _cr_xl = _scheme

    # Live swatch preview
    st.markdown(
        f"""
        <div style="display:flex;gap:8px;margin:4px 0 16px 0;">
            <div style="flex:1;background:{_cc_web or '#ffffff'};border:1px solid #ccc;
                        border-radius:3px;padding:6px 8px;font-size:0.68rem;
                        font-weight:600;color:#333;text-align:center;">Cost Centre</div>
            <div style="flex:1;background:{_cr_web or '#ffffff'};border:1px solid #ccc;
                        border-radius:3px;padding:6px 8px;font-size:0.68rem;
                        font-weight:600;color:#333;text-align:center;">Creditor</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("---")

    # ─ Ledger Configuration ────────────────────────────────
    st.markdown("## Ledger Configuration")
    st.caption("Names used to detect transaction header rows. One ledger per line.")

    # Load into session_state once (survives reruns within the same session).
    if "ledger_text" not in st.session_state:
        st.session_state.ledger_text = (
            LEDGER_PATH.read_text(encoding="utf-8", errors="ignore")
            if LEDGER_PATH.exists()
            else ""
        )

    edited = st.text_area(
        "Ledger Names",
        st.session_state.ledger_text,
        height=350,
        key="ledger_editor",
        label_visibility="collapsed",
    )

    count = len([ln for ln in edited.splitlines() if ln.strip()])
    st.markdown(
        f"<div style='font-size:0.74rem;color:#888;margin-bottom:8px;'>"
        f"{count} ledgers loaded</div>",
        unsafe_allow_html=True,
    )

    col_s, col_d = st.columns(2)
    with col_s:
        if st.button("Apply", use_container_width=True,
                     help="Updates the ledger list for this session"):
            st.session_state.ledger_text = edited
            try:
                LEDGER_PATH.write_text(edited, encoding="utf-8")
            except OSError:
                pass
            st.success("Applied ✓")
    with col_d:
        st.download_button(
            "⬇ Save File",
            data=edited.encode("utf-8"),
            file_name="ledger.txt",
            mime="text/plain",
            use_container_width=True,
        )


# ─── File Upload ─────────────────────────────────────────────────────────────

uploaded_file = st.file_uploader(
    "Upload Tally Daybook Excel (.xlsx)",
    type=["xlsx", "xls"],
    help="Data must start from Row 8 in the standard Tally daybook layout.",
)

if uploaded_file:

    # Raw preview
    with st.expander("View Raw Excel Data (first 30 rows)", expanded=False):
        try:
            raw_df = pd.read_excel(
                uploaded_file, header=None, skiprows=6, nrows=30, engine="openpyxl"
            )
            st.dataframe(raw_df, use_container_width=True)
            uploaded_file.seek(0)
        except Exception as e:
            st.warning(f"Could not preview: {e}")
            uploaded_file.seek(0)

    # ── Parse ─────────────────────────────────────────────────────────────────
    with st.spinner("Processing daybook data…"):
        try:
            df = parse_daybook(uploaded_file, LEDGER_PATH)
        except Exception as e:
            st.error(f"Parsing error: {e}")
            st.stop()

    if df.empty:
        st.warning(
            "No ledger transactions found. Ensure ledger names in the sidebar "
            "match the Particulars column in your Excel file."
        )
        st.stop()

    # ── Summary Metrics ───────────────────────────────────────────────────────
    cc_total     = df.loc[df["Type"] == "Cost Centre", "Amount"].dropna().sum()
    cr_total     = df.loc[df["Type"] == "Creditor",    "Amount"].dropna().sum()
    n_txn        = df["Vch No."].nunique()
    n_ledgers    = df["Ledger"].nunique()
    n_rows       = len(df)

    st.markdown(
        f"""
        <div class="metric-strip">
            <div class="metric-card">
                <div class="m-label">Total Records</div>
                <div class="m-value">{n_rows:,}</div>
            </div>
            <div class="metric-card">
                <div class="m-label">Transactions</div>
                <div class="m-value">{n_txn:,}</div>
            </div>
            <div class="metric-card">
                <div class="m-label">Ledgers</div>
                <div class="m-value">{n_ledgers:,}</div>
            </div>
            <div class="metric-card green">
                <div class="m-label">Cost Centre Total (₹)</div>
                <div class="m-value">{cc_total:,.2f}</div>
            </div>
            <div class="metric-card red">
                <div class="m-label">Creditor Total (₹)</div>
                <div class="m-value">{cr_total:,.2f}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Filters ───────────────────────────────────────────────────────────────
    st.markdown('<div class="filter-panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">🔍 Filter Records</div>', unsafe_allow_html=True)

    fc1, fc2, fc3, fc4 = st.columns(4)
    all_dates   = sorted(df["Date"].dropna().unique().tolist())
    all_ledgers = sorted(df["Ledger"].dropna().unique().tolist())
    all_vchs    = sorted(df["Vch Type"].dropna().unique().tolist())
    all_types   = sorted(df["Type"].dropna().unique().tolist())

    with fc1:
        sel_dates   = st.multiselect("Date", all_dates, placeholder="All dates")
    with fc2:
        sel_ledgers = st.multiselect("Ledger", all_ledgers, placeholder="All ledgers")
    with fc3:
        sel_vchs    = st.multiselect("Voucher Type", all_vchs, placeholder="All types")
    with fc4:
        sel_types   = st.multiselect("Entry Type", all_types, placeholder="All entries")

    st.markdown("</div>", unsafe_allow_html=True)

    filtered = df.copy()
    if sel_dates:
        filtered = filtered[filtered["Date"].isin(sel_dates)]
    if sel_ledgers:
        filtered = filtered[filtered["Ledger"].isin(sel_ledgers)]
    if sel_vchs:
        filtered = filtered[filtered["Vch Type"].isin(sel_vchs)]
    if sel_types:
        filtered = filtered[filtered["Type"].isin(sel_types)]

    # ── Data Table ────────────────────────────────────────────────────────────
    st.markdown('<div class="table-container">', unsafe_allow_html=True)
    st.markdown(
        f'<div class="section-label">Normalised Daybook — {len(filtered):,} entries</div>',
        unsafe_allow_html=True,
    )

    def _row_colour(row):
        if row["Type"] == "Cost Centre" and _cc_web:
            return [f"background-color: {_cc_web}"] * len(row)
        elif row["Type"] == "Creditor" and _cr_web:
            return [f"background-color: {_cr_web}"] * len(row)
        return [""] * len(row)

    styled = (
        filtered.style
        .apply(_row_colour, axis=1)
        .format({"Amount": lambda x: f"₹ {x:,.2f}" if pd.notna(x) else "—"})
        .set_properties(**{"font-size": "12.5px", "border": "1px solid #e5e5e5"})
        .set_table_styles([
            {"selector": "th", "props": [
                ("background-color", "#1c1c1c"),
                ("color", "#ffffff"),
                ("font-size", "11px"),
                ("font-weight", "600"),
                ("text-transform", "uppercase"),
                ("letter-spacing", "0.05em"),
                ("padding", "10px 12px"),
                ("border", "1px solid #c0392b"),
            ]},
        ])
    )

    st.dataframe(styled, use_container_width=True, height=520)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Legend ────────────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style="display:flex;gap:16px;margin:8px 0 20px 0;align-items:center;">
            <span style="font-size:0.72rem;color:#666666;font-weight:600;text-transform:uppercase;
                         letter-spacing:0.05em;">Legend:</span>
            <span style="display:inline-block;background:{cc_swatch};color:#333333;
                font-size:0.65rem;font-weight:700;padding:2px 12px;border-radius:2px;
                border:1px solid #cccccc;letter-spacing:0.05em;text-transform:uppercase;"
            >&#9632; &nbsp;Cost Centre</span>
            <span style="display:inline-block;background:{cr_swatch};color:#333333;
                font-size:0.65rem;font-weight:700;padding:2px 12px;border-radius:2px;
                border:1px solid #cccccc;letter-spacing:0.05em;text-transform:uppercase;"
            >&#9632; &nbsp;Creditor</span>
            <span style="font-size:0.68rem;color:#aaa;margin-left:4px;">({_label})</span>
        </div>
        """.format(cc_swatch=_cc_web or "#ffffff", cr_swatch=_cr_web or "#ffffff", _label=_label),
        unsafe_allow_html=True,
    )

    # ── Downloads ─────────────────────────────────────────────────────────────
    st.markdown('<div class="section-label" style="margin-top:8px">Export Data</div>', unsafe_allow_html=True)
    dl1, dl2, dl3 = st.columns([1, 1, 4])

    # CSV
    csv_bytes = filtered.to_csv(index=False).encode("utf-8-sig")
    with dl1:
        st.download_button("⬇ Download CSV", data=csv_bytes,
                           file_name="daybook_normalised.csv", mime="text/csv",
                           use_container_width=True)

    # Excel
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="xlsxwriter") as writer:
        filtered.to_excel(writer, index=False, sheet_name="Daybook")
        wb  = writer.book
        ws  = writer.sheets["Daybook"]

        # Formats
        hdr_fmt  = wb.add_format({"bold": True, "bg_color": "#1c1c1c",
                                  "font_color": "#ffffff", "border": 1,
                                  "align": "center", "font_size": 9,
                                  "text_wrap": True})
        money_fmt = wb.add_format({"num_format": "#,##0.00", "border": 1,
                                   "font_size": 9})
        cc_fmt    = wb.add_format({"bg_color": _cc_xl, "border": 1, "font_size": 9})
        cr_fmt    = wb.add_format({"bg_color": _cr_xl, "border": 1, "font_size": 9})
        cell_fmt  = wb.add_format({"border": 1, "font_size": 9})

        col_widths = {
            "Date": 13, "Ledger": 42, "Vch Type": 12, "Vch No.": 22,
            "Particulars": 38, "Amount": 15, "Type": 13,
        }

        # Write header
        for ci, cn in enumerate(filtered.columns):
            ws.write(0, ci, cn, hdr_fmt)
            ws.set_column(ci, ci, col_widths.get(cn, 18))

        # Write data rows with colour coding
        for ri, (_, dr) in enumerate(filtered.iterrows(), start=1):
            row_type = dr.get("Type", "")
            base_fmt = cc_fmt if row_type == "Cost Centre" else \
                       (cr_fmt if row_type == "Creditor" else cell_fmt)
            for ci, cn in enumerate(filtered.columns):
                val = dr[cn]
                fmt = money_fmt if cn == "Amount" else base_fmt
                ws.write(ri, ci, (None if (isinstance(val, float) and
                                           pd.isna(val)) else val), fmt)

        # Freeze header row
        ws.freeze_panes(1, 0)

    with dl2:
        st.download_button("⬇ Download Excel", data=excel_buf.getvalue(),
                           file_name="daybook_normalised.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

else:
    st.markdown("<br>", unsafe_allow_html=True)
    st.info(
        "**Upload a Tally daybook Excel file to begin.**  \n"
        "Expected layout: data from Row 8 · Columns: Date · Particulars · "
        "Sub-Amount · Dr/Cr · Vch Type · Vch No. · Debit Amount · Credit Amount"
    )
