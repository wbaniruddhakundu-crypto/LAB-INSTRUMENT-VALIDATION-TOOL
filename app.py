"""Precision Validation Tool — Streamlit edition.

Upload Siemens/Atellica analyzer PDF reports (SID "CH PRECISION DAYn Ln") and
auto-generate one precision Excel per analyte — either from a built-in template
or populated straight into your own uploaded template, with live formulas.
"""

from __future__ import annotations

import streamlit as st

from pdf_parser import extract_text_from_pdf, parse_analyzer_text
from excel_export import (
    AnalyteInfo,
    ResultRecord,
    WorklistInfo,
    analyte_filename,
    build_workbook,
    build_zip,
    parse_template,
    populate_template,
)

st.set_page_config(page_title="Precision Validation Tool", page_icon="🧪", layout="wide")


@st.cache_data(show_spinner=False)
def _parse_pdf(file_bytes: bytes, precision_type: str, force_day: int | None):
    text = extract_text_from_pdf(file_bytes)
    parsed = parse_analyzer_text(text, precision_type)
    out = []
    for r in parsed:
        day = force_day if (precision_type == "inter" and force_day is not None) else r.day_number
        out.append(
            ResultRecord(
                analyte_name=r.analyte_name,
                level=r.level,
                day_number=(day if precision_type == "inter" else 0),
                replication=r.replication,
                precision_type=precision_type,
                value=r.value,
                date=r.date,
                unit=r.unit,
            )
        )
    return out


@st.cache_data(show_spinner=False)
def _parse_template_cached(file_bytes: bytes):
    return parse_template(file_bytes)


@st.cache_data(show_spinner=False)
def _build_one(analyte_name: str, unit: str, levels: int,
               records: list[ResultRecord], worklist: WorklistInfo,
               template_bytes: bytes | None) -> bytes:
    analyte = AnalyteInfo(
        name=analyte_name, levels=levels, unit=unit,
        control_info=[{"lot": "", "limits": ""} for _ in range(levels)],
    )
    if template_bytes is not None:
        tpl = parse_template(template_bytes)
        return populate_template(tpl, analyte, records)
    return build_workbook(analyte, worklist, records)


# ── Header ───────────────────────────────────────────────────────────────────

st.title("🧪 Precision Validation Tool")
st.caption(
    "Upload your analyzer PDF reports day by day. Analytes, days and replicates "
    "are detected automatically and each analyte's precision Excel is generated "
    "to match your template — no manual entry."
)

# ── Lab info ─────────────────────────────────────────────────────────────────

with st.expander("Lab Information (optional — appears in the generated Excel header)"):
    c1, c2, c3 = st.columns(3)
    session_name = c1.text_input("Run / Session Name", "Precision Run")
    location = c2.text_input("Location", "")
    instrument = c3.text_input("Instrument", "")
    c4, c5 = st.columns(2)
    hodcol = c4.text_input("HOD / COL", "")
    technician = c5.text_input("Technician", "")

lab_meta = {
    "location": location, "instrument": instrument,
    "hodcol": hodcol, "technician": technician,
}

# ── Template upload ──────────────────────────────────────────────────────────

st.subheader("Template Excel (optional)")
st.caption(
    "Upload your precision template and the data fills into it — respecting "
    "however many days / replicates / control levels it has. Without one, a "
    "standard sheet is generated."
)
template_file = st.file_uploader("Template (.xlsx)", type=["xlsx"], key="template")
template_bytes: bytes | None = None
template_obj = None
if template_file is not None:
    template_bytes = template_file.getvalue()
    try:
        template_obj = _parse_template_cached(template_bytes)
        if not template_obj.blocks:
            st.error("Couldn't find a DATE / REPLICATION grid in this file.")
            template_bytes = None
        else:
            st.success(
                f"Template loaded: {len(template_obj.blocks)} control(s) · "
                f"up to {template_obj.max_days} day(s) · {template_obj.max_reps} replicate(s)."
            )
    except Exception as e:  # noqa: BLE001
        st.error(f"Couldn't read the template: {e}")
        template_bytes = None

# ── Inter-day uploads ────────────────────────────────────────────────────────

st.subheader("Inter-Day Precision")
default_days = template_obj.max_days if template_obj else 5
num_days = st.number_input(
    "Number of days", min_value=1, max_value=30,
    value=int(default_days) if default_days else 5, step=1,
)
st.caption("Upload each day's PDF(s). Each day becomes one row; replicates fill across the columns.")

records: list[ResultRecord] = []
day_cols = st.columns(min(int(num_days), 5))
for d in range(1, int(num_days) + 1):
    col = day_cols[(d - 1) % len(day_cols)]
    files = col.file_uploader(
        f"Day {d}", type=["pdf"], accept_multiple_files=True, key=f"day_{d}"
    )
    for f in files or []:
        records.extend(_parse_pdf(f.getvalue(), "inter", d))

# ── Intra-day upload ─────────────────────────────────────────────────────────

st.subheader("Intra-Day Precision")
st.caption("Upload the within-run PDF(s) containing all replicates (e.g. 10/15/20 reps).")
intra_files = st.file_uploader(
    "Intra-day PDF(s)", type=["pdf"], accept_multiple_files=True, key="intra"
)
for f in intra_files or []:
    records.extend(_parse_pdf(f.getvalue(), "intra", None))

# ── Aggregate + downloads ────────────────────────────────────────────────────

records = [r for r in records if r.value is not None]

if not records:
    st.info("Upload at least one analyzer PDF to detect analytes and generate Excel files.")
    st.stop()

analyte_names = sorted({r.analyte_name for r in records})

units: dict[str, str] = {}
analytes: list[AnalyteInfo] = []
for name in analyte_names:
    rs = [r for r in records if r.analyte_name == name]
    levels = max((r.level for r in rs), default=1)
    unit = next((r.unit for r in rs if r.unit), "")
    units[name] = unit
    analytes.append(AnalyteInfo(
        name=name, levels=levels, unit=unit,
        control_info=[{"lot": "", "limits": ""} for _ in range(levels)],
    ))

inter = [r for r in records if r.precision_type == "inter"]
intra = [r for r in records if r.precision_type == "intra"]
worklist = WorklistInfo(
    name=session_name,
    num_days=max((r.day_number for r in inter), default=1),
    reps_per_day=max((r.replication for r in inter), default=1),
    intra_day_reps=max((r.replication for r in intra), default=1),
    precision_type=("both" if inter and intra else "intra" if intra else "inter"),
    analytes=analytes,
    metadata=lab_meta,
)

st.divider()
left, right = st.columns([3, 1])
left.subheader(f"Detected Analytes ({len(analyte_names)})")
left.caption(f"{len(records)} results populated. Download one analyte or grab them all as a ZIP.")

zip_bytes = build_zip(worklist, records, template_obj)
right.download_button(
    "⬇️ Download All (ZIP)",
    data=zip_bytes,
    file_name=f"{session_name or 'Precision'}_All_Analytes.zip",
    mime="application/zip",
    use_container_width=True,
)

grid = st.columns(4)
for i, name in enumerate(analyte_names):
    a = next(a for a in analytes if a.name == name)
    count = sum(1 for r in records if r.analyte_name == name)
    data = _build_one(name, units.get(name, ""), a.levels, records, worklist, template_bytes)
    grid[i % 4].download_button(
        f"📄 {name}  ({count})",
        data=data,
        file_name=analyte_filename(name),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_{name}",
        use_container_width=True,
    )
