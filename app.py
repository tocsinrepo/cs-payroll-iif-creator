"""
Cornerstones Payroll IIF Generator — Streamlit App
Upload the Paycom source xlsx, enter pay period details, download 4 QuickBooks IIF files.
"""

import os
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

import streamlit as st

# ── Path setup ────────────────────────────────────────────────────────────────
APP_DIR  = Path(__file__).parent
REF1     = APP_DIR / "reference" / "Reference 1 Class Table.xlsx"
REF2     = APP_DIR / "reference" / "Reference 2 Class Table FINAL CHECK.xlsx"

# Add app dir to path so generate_iif is importable
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Cornerstones Payroll IIF Generator",
    page_icon="📊",
    layout="wide",
)

# ── Header ────────────────────────────────────────────────────────────────────
st.title("📊 Cornerstones Payroll IIF Generator")
st.caption(
    "Upload the Paycom source workbook, fill in pay period details, "
    "and download the four QuickBooks IIF journal entry files."
)
st.divider()

# ── Session state init ────────────────────────────────────────────────────────
if "result" not in st.session_state:
    st.session_state.result = None   # dict: {success, log, files}

# ── Layout: two columns ───────────────────────────────────────────────────────
col_left, col_right = st.columns([1, 1], gap="large")

with col_left:
    st.subheader("1 · Source File")
    uploaded = st.file_uploader(
        "Upload source XLSX",
        type=["xlsx"],
        label_visibility="collapsed",
        help="e.g.  FY26 - PD 20260313 Payroll.xlsx",
    )
    if uploaded:
        st.success(f"✔  {uploaded.name}  ({uploaded.size / 1024:.0f} KB)")

with col_right:
    st.subheader("2 · Pay Period Details")
    c1, c2, c3 = st.columns(3)
    pay_date_val   = c1.date_input("Pay Date", value=date.today())
    pay_period_num = c2.number_input("Pay Period #", min_value=1, max_value=26, value=18, step=1)
    fiscal_year_num = c3.number_input("Fiscal Year (2-digit)", min_value=24, max_value=35, value=26, step=1)

    prior_pay_date_val = st.date_input(
        "Prior Pay Date  **optional — for cash line memo)*",
        value=None,
        help="Leave blank to use the current pay date in the memo.",
    )

    with st.expander("⚙ Advanced — Extra Class Overrides", expanded=False):
        extra_class_text = st.text_area(
            "One override per line:  CODE=Full QB Class Path",
            height=80,
            placeholder="22000=20000 Resource Development:22000 Volunteers",
            help="Use to fix any class code that is missing or incorrectly mapped.",
        )

st.divider()

# ── Generate button ───────────────────────────────────────────────────────────
generate_disabled = uploaded is None
generate_btn = st.button(
    "⚡  Generate IIF Files",
    type="primary",
    disabled=generate_disabled,
    use_container_width=True,
)

if generate_disabled:
    st.caption("⬆ Upload a source file to enable generation.")

# ── Run generation ────────────────────────────────────────────────────────────
if generate_btn and uploaded:
    st.session_state.result = None

    with st.spinner("Generating IIF files — this takes a few seconds…"):
        try:
            from generate_iif import run_programmatic
        except ImportError as e:
            st.error(f"Could not import generate_iif: {e}")
            st.stop()

        # Build pay date string
        pd = pay_date_val
        pay_date_str = f"{pd.month}/{pd.day}/{pd.year}"

        # Prior pay date string
        prior_str = ""
        if prior_pay_date_val:
            pp = prior_pay_date_val
            prior_str = f"{pp.month}/{pp.day}/{pp.year}"

        # Extra class overrides
        extra_classes = []
        if extra_class_text.strip():
            extra_classes = [ln.strip() for ln in extra_class_text.strip().splitlines() if ln.strip()]

        with tempfile.TemporaryDirectory() as tmpdir:
            # Save uploaded file
            src_path = os.path.join(tmpdir, uploaded.name)
            with open(src_path, "wb") as f:
                f.write(uploaded.getvalue())

            success, log, files = run_programmatic(
                source_path            = src_path,
                class_ref_path         = str(REF1),
                class_final_check_path = str(REF2),
                pay_date_str           = pay_date_str,
                pay_period             = int(pay_period_num),
                fiscal_year            = int(fiscal_year_num),
                output_dir             = tmpdir,
                prior_pay_date_str     = prior_str,
                extra_classes          = extra_classes or None,
            )
            st.session_state.result = {
                "success": success,
                "log":     log,
                "files":   files,
            }

# ── Results ───────────────────────────────────────────────────────────────────
res = st.session_state.result
if res is not None:
    st.divider()

    if res["success"]:
        st.success("✅  All Pivot Check amounts match — files are safe to import into QuickBooks.")
    else:
        st.error("❌  Generation failed or Pivot Check mismatch — **do NOT import these files.**")

    # Download buttons
    files = res["files"]
    if files:
        st.subheader("3 · Download IIF Files")
        dl_cols = st.columns(4)
        labels = {
            "Output 1": "📥 Output 1 — Payroll JE",
            "Output 2": "📥 Output 2 — Taxes JE",
            "Output 3": "📥 Output 3 — 401K JE",
            "Output 4": "📥 Output 4 — Payroll Fee JE",
        }
        for i, (fname, content) in enumerate(sorted(files.items())):
            prefix = fname[:8]   # "Output 1", "Output 2" …
            label  = labels.get(prefix, f"📥 {fname}")
            dl_cols[i % 4].download_button(
                label      = label,
                data       = content,
                file_name  = fname,
                mime       = "text/plain",
                key        = f"dl_{fname}",
                use_container_width = True,
            )

    # Generation log
    with st.expander("📋 Generation Log", expanded=(not res["success"])):
        st.code(res["log"], language="")

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.caption("Cornerstones, Inc. · Internal payroll processing tool · Not for distribution")
