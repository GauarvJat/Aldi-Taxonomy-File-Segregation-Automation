"""
Aldi DG  —  Data Segregation Tool  (Streamlit / Web version)
─────────────────────────────────────────────────────────────
Run locally : streamlit run aldi_dg_streamlit.py
Deploy      : Push repo to GitHub → connect to streamlit.io
               Set main file to:   aldi_dg_streamlit.py

requirements.txt (must be in the GitHub repo ROOT alongside this file):
  streamlit>=1.35.0
  pandas>=2.0.0
  xlsxwriter>=3.1.0
  xlrd>=2.0.1
  openpyxl>=3.1.0
"""

from __future__ import annotations

import io
import zipfile
from pathlib import PurePath

import streamlit as st

# ── Safe imports with clear error message ─────────────────────────────────────
_missing: list[str] = []

try:
    import pandas as pd
except ImportError:
    _missing.append("pandas")

try:
    import xlsxwriter  # noqa: F401
except ImportError:
    _missing.append("xlsxwriter")

if _missing:
    st.error(
        f"**Missing package(s): `{'`, `'.join(_missing)}`**\n\n"
        "Make sure your GitHub repo root contains a file named **`requirements.txt`** "
        "with this content:\n\n"
        "```\nstreamlit>=1.35.0\npandas>=2.0.0\nxlsxwriter>=3.1.0\n"
        "xlrd>=2.0.1\nopenpyxl>=3.1.0\n```\n\n"
        "After committing and pushing, open Streamlit Cloud → "
        "**Manage app** → **Reboot app**."
    )
    st.stop()


# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════

PLATFORM_CONFIG: dict[str, list[str]] = {
    "AdServer":    ["CAMPAIGN", "PLACEMENT", "CREATIVE"],
    "Paid Social": ["CAMPAIGN", "PLACEMENT", "CREATIVE"],
    "Programmatic":["CAMPAIGN", "PLACEMENT", "PLACEMENTGROUP"],  # CREATIVE = PLACEMENTGROUP
    "Search":      ["CAMPAIGN", "PLACEMENT", "CREATIVE"],
}

# Maps compact / alternate spellings → canonical platform name
PLATFORM_ALIASES: dict[str, str] = {
    "adserver":     "AdServer",
    "ad server":    "AdServer",
    "paidsocial":   "Paid Social",
    "paid social":  "Paid Social",
    "programmatic": "Programmatic",
    "search":       "Search",
}

# Longest first → prevents "PLACEMENT" matching before "PLACEMENTGROUP"
KNOWN_LEVELS: list[str] = ["PLACEMENTGROUP", "CAMPAIGN", "PLACEMENT", "CREATIVE"]

RED_HEX = "#FF0000"


# ═══════════════════════════════════════════════════════════════════════════════
#  FILENAME PARSER
# ═══════════════════════════════════════════════════════════════════════════════

def parse_filename(filename: str) -> tuple[str | None, str | None, str | None]:
    """
    Accepts two patterns:
      A)  <Channel> <LEVEL> - <Date>.xlsx   (original format with " - " separator)
      B)  <Channel> <LEVEL> <Date>.xlsx     (space-only format)

    Returns (platform_raw, level, date) or (None, None, None).
    """
    stem = PurePath(filename).stem.strip()

    # Pattern A — split on " - "
    if " - " in stem:
        try:
            name_part, date = stem.rsplit(" - ", 1)
            for level in KNOWN_LEVELS:
                if name_part.upper().endswith(level):
                    platform_raw = name_part[: -len(level)].strip()
                    return platform_raw, level, date.strip()
        except ValueError:
            pass

    # Pattern B — space-separated pivot token
    parts = stem.split()
    if len(parts) >= 3:
        for i, token in enumerate(parts):
            if token.upper() in KNOWN_LEVELS:
                if i == 0 or i == len(parts) - 1:
                    continue
                return (
                    " ".join(parts[:i]).strip(),
                    token.upper(),
                    " ".join(parts[i + 1:]).strip(),
                )

    return None, None, None


def resolve_platform(platform_raw: str) -> str | None:
    return PLATFORM_ALIASES.get(platform_raw.lower())


# ═══════════════════════════════════════════════════════════════════════════════
#  CORE PROCESSING
# ═══════════════════════════════════════════════════════════════════════════════

def get_nc_col_range(df: pd.DataFrame) -> tuple[int | None, int | None]:
    """
    Returns (first_nc_idx, last_nc_idx + 1) as 0-based column positions.
    Use as df.iloc[:, start:end] to select the full NC range.
    """
    nc_indices = [i for i, col in enumerate(df.columns) if "NC" in str(col)]
    if not nc_indices:
        return None, None
    return nc_indices[0], nc_indices[-1] + 1


def process_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    1. Fill every blank/empty cell inside the NC column range with 'Is missing'.
    2. Rename duplicate Market column to Market_MK.
    """
    start, end = get_nc_col_range(df)

    if start is not None:
        nc_slice = df.iloc[:, start:end].copy()
        nc_slice = nc_slice.replace("", pd.NA)
        nc_slice = nc_slice.fillna("Is missing")
        df.iloc[:, start:end] = nc_slice

    market_cols = [c for c in df.columns if str(c).lower() == "market"]
    if len(market_cols) > 1:
        df.rename(columns={market_cols[1]: "Market_MK"}, inplace=True)

    return df


def _apply_nc_formatting(
    writer:     pd.ExcelWriter,
    df:         pd.DataFrame,
    sheet_name: str,
) -> None:
    """
    Add red conditional-format RULES via xlsxwriter for every column in the
    NC range.  Excel evaluates the rule natively — no cell-by-cell iteration,
    no openpyxl read-back.  Fast regardless of row count.
    """
    start, end = get_nc_col_range(df)
    if start is None:
        return

    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]
    max_row   = len(df)   # data rows; header sits at row index 0

    red_fmt = workbook.add_format({"bg_color": RED_HEX, "font_color": "#000000"})
    bad_texts = ["Is missing", "Invalid", "Invalid value"]

    for col_idx in range(start, end):
        for bad_text in bad_texts:
            worksheet.conditional_format(
                1, col_idx, max_row, col_idx,  # row1, col1, row2, col2 (0-based)
                {
                    "type":     "text",
                    "criteria": "containing",
                    "value":    bad_text,
                    "format":   red_fmt,
                },
            )


# ═══════════════════════════════════════════════════════════════════════════════
#  ORCHESTRATION
# ═══════════════════════════════════════════════════════════════════════════════

def run_segregation(
    uploaded_files: list,
    progress_cb=None,
) -> tuple[dict[str, bytes], dict]:
    """
    Process uploaded .xlsx files entirely in memory.
    Returns (output_files, stats).
    """
    stats: dict = {
        "files_created":       0,
        "tabs_created":        0,
        "platforms_processed": 0,
        "skipped":             [],
    }
    inputs: dict = {}

    # ── Phase 1: group files ───────────────────────────────────────────────────
    for uf in uploaded_files:
        platform_raw, level, date = parse_filename(uf.name)

        if not platform_raw:
            stats["skipped"].append(f"{uf.name}  — filename not parseable")
            continue
        platform = resolve_platform(platform_raw)
        if not platform:
            stats["skipped"].append(f"{uf.name}  — unknown channel: '{platform_raw}'")
            continue
        if level not in PLATFORM_CONFIG[platform]:
            stats["skipped"].append(f"{uf.name}  — '{level}' not valid for {platform}")
            continue

        inputs.setdefault(platform, {}).setdefault(date, {})[level] = uf

    output_files: dict[str, bytes] = {}
    total_combos = sum(len(d) for d in inputs.values())
    combo_idx    = 0

    # ── Phase 2: process ──────────────────────────────────────────────────────
    for platform, dates in inputs.items():
        stats["platforms_processed"] += 1

        for date, levels in dates.items():
            combo_idx += 1
            if progress_cb:
                progress_cb(combo_idx / max(total_combos, 1))

            market_data: dict = {}

            for level, uf in levels.items():
                try:
                    uf.seek(0)
                    df = pd.read_excel(
                        io.BytesIO(uf.read()),
                        keep_default_na=False,
                        na_values=[""],
                    )
                except Exception as exc:
                    stats["skipped"].append(f"{uf.name}  — could not read: {exc}")
                    continue

                if df.empty or "Market" not in df.columns:
                    stats["skipped"].append(
                        f"{uf.name}  — empty or missing 'Market' column"
                    )
                    continue

                df = process_dataframe(df)

                for market, group in df.groupby("Market"):
                    market_data.setdefault(str(market), {})[level] = group

            # ── One xlsx per market, written with xlsxwriter ───────────────────
            for market, levels_dict in market_data.items():
                safe_market     = str(market).replace("/", "-").replace("\\", "-")
                output_filename = f"{platform}_{safe_market}_{date}.xlsx"

                buf = io.BytesIO()
                with pd.ExcelWriter(
                    buf,
                    engine="xlsxwriter",
                    engine_kwargs={"options": {"strings_to_urls": False}},
                ) as writer:
                    for level in PLATFORM_CONFIG[platform]:
                        if level not in levels_dict:
                            continue
                        level_df = levels_dict[level]
                        level_df.to_excel(writer, sheet_name=level, index=False)
                        stats["tabs_created"] += 1
                        _apply_nc_formatting(writer, level_df, level)

                output_files[output_filename] = buf.getvalue()
                stats["files_created"] += 1

    return output_files, stats


def build_zip(output_files: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, data in output_files.items():
            zf.writestr(fname, data)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ═══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Aldi DG — Data Segregation",
    page_icon="🛒",
    layout="centered",
)

st.markdown("""
<style>
html, body, [data-testid="stAppViewContainer"] { background-color: #0A0F1E !important; }
[data-testid="stAppViewContainer"] > .main     { background-color: #0A0F1E; }
[data-testid="stHeader"]  { background: transparent !important; }
[data-testid="stToolbar"] { display: none; }
footer { visibility: hidden; }

.aldi-hero {
    background: linear-gradient(135deg, #0D1B2A 0%, #1A2744 100%);
    border-bottom: 4px solid #00A6E2;
    border-radius: 12px;
    padding: 32px 36px 26px;
    margin-bottom: 28px;
}
.aldi-hero-name { font-size: 2rem; font-weight: 900; color: #FFFFFF; margin: 0 0 4px; }
.aldi-hero-name span { color: #00A6E2; }
.aldi-hero-sub  { font-size: .95rem; color: #7A9BB5; margin: 0 0 18px; }
.aldi-tags { display: flex; flex-wrap: wrap; gap: 8px; }
.aldi-tag  {
    background: rgba(0,166,226,.1); border: 1px solid rgba(0,166,226,.3);
    color: #7DCFEF; font-size: .78rem; font-weight: 600;
    padding: 3px 12px; border-radius: 20px;
}
.aldi-label {
    font-size: .72rem; font-weight: 800; letter-spacing: 1.5px;
    text-transform: uppercase; color: #7A9BB5; margin-bottom: 6px;
}

[data-testid="stFileUploader"] section {
    background: #0D1B2A !important;
    border: 1.5px dashed #1E3A5C !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploader"] section:hover { border-color: #00A6E2 !important; }
[data-testid="stFileUploader"] label { color: #7A9BB5 !important; }

.stButton > button {
    background: #00A6E2 !important; color: #FFFFFF !important;
    border: none !important; border-radius: 8px !important;
    font-weight: 800 !important; font-size: 1rem !important;
    padding: 12px 0 !important; width: 100% !important;
}
.stButton > button:hover    { background: #0088C0 !important; }
.stButton > button:disabled { opacity: .45 !important; }

.stDownloadButton > button {
    background: #FF6200 !important; color: #FFFFFF !important;
    border: none !important; border-radius: 8px !important;
    font-weight: 800 !important; font-size: .95rem !important;
    padding: 12px 0 !important; width: 100% !important;
}
.stDownloadButton > button:hover { background: #D95200 !important; }

[data-testid="metric-container"] {
    background: #0D1B2A; border: 1px solid #1A3A5C;
    border-radius: 10px; padding: 14px !important;
}
[data-testid="metric-container"] label { color: #7A9BB5 !important; font-size: .8rem !important; }
[data-testid="stMetricValue"]          { color: #00A6E2 !important; }

[data-testid="stAlert"]    { background: #0D1B2A !important; border: 1px solid #1A3A5C !important; border-radius: 8px !important; color: #7A9BB5 !important; }
[data-testid="stExpander"] { background: #0D1B2A !important; border: 1px solid #1A3A5C !important; border-radius: 10px !important; }
.streamlit-expanderHeader  { color: #7A9BB5 !important; font-size: .9rem !important; }

[data-testid="stProgressBar"] > div > div { background: #00A6E2 !important; }

hr   { border-color: #1A3A5C !important; }
p, li, span { color: #7A9BB5; }
strong      { color: #D0E4F0 !important; }
code        {
    background: #060C16 !important; color: #7DCFEF !important;
    border-radius: 4px !important; font-size: .85rem !important;
}

.aldi-footer {
    text-align: center; color: #1A3A5C; font-size: .78rem;
    padding: 24px 0 8px; border-top: 1px solid #1A3A5C; margin-top: 40px;
}
</style>
""", unsafe_allow_html=True)


# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="aldi-hero">
    <div class="aldi-hero-name">🛒 Aldi <span>DG</span></div>
    <div class="aldi-hero-sub">Data Segregation Tool &nbsp;·&nbsp; v1.0</div>
    <div class="aldi-tags">
        <span class="aldi-tag">● AdServer</span>
        <span class="aldi-tag">● Paid Social</span>
        <span class="aldi-tag">● Programmatic</span>
        <span class="aldi-tag">● Search</span>
    </div>
</div>
""", unsafe_allow_html=True)


# ── File naming guide ─────────────────────────────────────────────────────────
with st.expander("📋  File Naming Convention — click to expand"):
    st.markdown("""
Each uploaded file **must** follow one of these patterns:
```
<Channel> <LEVEL> - <Date>.xlsx      ← with separator  (original)
<Channel> <LEVEL> <Date>.xlsx        ← space only
```
**Valid examples:**
```
AdServer CAMPAIGN - 06-04-2026.xlsx
Paid Social PLACEMENT - 06-04-2026.xlsx
Programmatic PLACEMENTGROUP - 06-04-2026.xlsx
Search CREATIVE - 06-04-2026.xlsx
```

| Channel | Valid Levels |
|---|---|
| AdServer | CAMPAIGN · PLACEMENT · CREATIVE |
| Paid Social | CAMPAIGN · PLACEMENT · CREATIVE |
| Search | CAMPAIGN · PLACEMENT · CREATIVE |
| Programmatic | CAMPAIGN · PLACEMENT · PLACEMENTGROUP |

> **Programmatic note:** Use `PLACEMENTGROUP` — it maps to CREATIVE for this channel.

**What this tool does:**
- Fills every blank cell in the NC column range with `"Is missing"`
- Highlights `"Is missing"` / `"Invalid"` cells in **red** (NC range only)
- Splits rows by `Market` → one `.xlsx` file per market
- One tab per campaign level per output file
""")


# ── Upload ────────────────────────────────────────────────────────────────────
st.markdown('<div class="aldi-label">Upload Input Files</div>', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    label="Drop your .xlsx files here or click to browse",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if uploaded_files:
    valid_count   = 0
    invalid_names = []
    for uf in uploaded_files:
        p, l, d = parse_filename(uf.name)
        if p and resolve_platform(p) and l:
            valid_count += 1
        else:
            invalid_names.append(uf.name)

    col_a, col_b = st.columns(2)
    col_a.metric("Files uploaded",     len(uploaded_files))
    col_b.metric("Files ready to run", valid_count)

    if invalid_names:
        with st.expander(f"⚠️  {len(invalid_names)} file(s) with unrecognised names"):
            for name in invalid_names:
                st.markdown(f"- `{name}`")

st.markdown("<br>", unsafe_allow_html=True)


# ── Run ───────────────────────────────────────────────────────────────────────
run_clicked = st.button("▶   Run Segregation", disabled=not bool(uploaded_files))

if run_clicked and uploaded_files:
    progress_bar = st.progress(0, text="Initialising…")

    def _update(fraction: float) -> None:
        pct = min(int(fraction * 100), 99)
        progress_bar.progress(pct, text=f"Processing…  {pct}%")

    with st.spinner("Running segregation — please wait…"):
        try:
            output_files, stats = run_segregation(uploaded_files, progress_cb=_update)
        except Exception as exc:
            st.error(f"Fatal error: {exc}")
            st.stop()

    progress_bar.progress(100, text="Done ✅")
    st.markdown("---")

    # ── Results ───────────────────────────────────────────────────────────────
    st.markdown("### ✅ Segregation Complete")
    c1, c2, c3 = st.columns(3)
    c1.metric("Platforms processed",  stats["platforms_processed"])
    c2.metric("Output files created", stats["files_created"])
    c3.metric("Tabs written",         stats["tabs_created"])

    if stats["skipped"]:
        with st.expander(f"⚠️  {len(stats['skipped'])} file(s) skipped", expanded=True):
            for msg in stats["skipped"]:
                st.markdown(f"- `{msg}`")

    # ── Download ──────────────────────────────────────────────────────────────
    if output_files:
        st.markdown("---")
        st.markdown('<div class="aldi-label">Download Output</div>', unsafe_allow_html=True)

        if len(output_files) == 1:
            fname, data = next(iter(output_files.items()))
            st.download_button(
                label     = f"⬇️  Download  {fname}",
                data      = data,
                file_name = fname,
                mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            zip_bytes = build_zip(output_files)
            st.download_button(
                label     = f"⬇️  Download All  ({len(output_files)} files)  as ZIP",
                data      = zip_bytes,
                file_name = "AldiDG_Segregated_Output.zip",
                mime      = "application/zip",
            )
            with st.expander(f"📂  Files in the ZIP  ({len(output_files)})"):
                for fname in sorted(output_files.keys()):
                    st.markdown(f"- `{fname}`")
    else:
        st.warning(
            "No output files were generated. "
            "Check your filenames match the required pattern above."
        )


# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="aldi-footer">© Aldi DG &nbsp;·&nbsp; Internal Use Only</div>',
    unsafe_allow_html=True,
)
