"""Parse Siemens/Atellica analyzer Assay Report PDFs into precision results.

Port of the original TypeScript pdf-parser. Each measurement is a separate
"CH PRECISION DAYn" block whose own replicate line always restarts at "1", so
the replicate number is assigned by COUNTING sequential measurements per
(analyte, day, level). The populated value is the reportable Result (the number
before the unit), e.g. "2.5 g/dL" -> 2.5, "32 U/L" -> 32.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import pdfplumber


@dataclass
class ParsedResult:
    analyte_name: str
    level: int
    day_number: int
    replication: int
    precision_type: str
    value: float | None
    unit: str
    date: str = ""


# The analyte gray heading is always the line right after the page banner
# "Order Time: All  Sample Type: All  Analyzer: All", followed by "Result".
ANALYTE_ANCHOR = re.compile(r"Order\s*Time:.*Analyzer:\s*All", re.IGNORECASE)


def extract_text_from_pdf(file) -> str:
    """Extract text, reconstructing lines from each word's vertical position so
    the analyte heading, SID line and result line stay on their own lines.

    `file` may be a path, a file-like object, or raw bytes.
    """
    import io

    if isinstance(file, (bytes, bytearray)):
        file = io.BytesIO(file)

    full_text: list[str] = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            words = page.extract_words(use_text_flow=False)
            line_map: dict[int, list[tuple[float, str]]] = {}
            for w in words:
                text = w.get("text", "").strip()
                if not text:
                    continue
                # Bucket by rounded vertical position (top grows downward).
                y = round(w["top"] / 2) * 2
                line_map.setdefault(y, []).append((w["x0"], text))

            for y in sorted(line_map.keys()):  # top -> bottom = reading order
                items = sorted(line_map[y], key=lambda i: i[0])
                line_text = " ".join(s for _, s in items).strip()
                if line_text:
                    full_text.append(line_text)
            full_text.append("")  # page break

    return "\n".join(full_text)


# Single-letter fragments (H, I, L) are layout artifacts, not real analytes.
SKIP_ANALYTES = {"H", "I", "L"}


def _read_analyte_name(lines: list[str], anchor_idx: int) -> str:
    name = ""
    for j in range(anchor_idx + 1, min(anchor_idx + 5, len(lines))):
        if re.match(r"^Result$", lines[j], re.IGNORECASE) or re.match(
            r"^Patient\s+Name", lines[j], re.IGNORECASE
        ):
            break
        name += lines[j]
    return re.sub(r"\s+", "", name)


def parse_analyzer_text(text: str, precision_type: str = "inter") -> list[ParsedResult]:
    results: list[ParsedResult] = []
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    # Template capacity: 5 replication columns (inter) / 20 (intra). Extra
    # measurements are dropped, keeping the first N to match the grid.
    max_reps = 20 if precision_type == "intra" else 5

    current_analyte = ""
    rep_counter: dict[str, int] = {}

    for i, line in enumerate(lines):
        # Analyte gray heading (line right after the page banner).
        if ANALYTE_ANCHOR.search(line):
            current_analyte = _read_analyte_name(lines, i)
            continue

        day_m = re.search(r"CH\s*PRECISION\s*DAY\s*(\d+)", line, re.IGNORECASE)
        if day_m and current_analyte and current_analyte not in SKIP_ANALYTES:
            day = int(day_m.group(1))
            level: int | None = None
            value: float | None = None
            unit = ""
            date_m = re.search(r"(\d{2}/\d{2}/\d{4})", line)
            date = date_m.group(1) if date_m else ""

            same_line_level = re.search(r"DAY\s*\d+\s+L\s*(\d+)", line, re.IGNORECASE)
            if same_line_level:
                level = int(same_line_level.group(1))

            # Scan the block (until the next CH PRECISION line) for value + level.
            for k in range(i + 1, min(i + 9, len(lines))):
                lk = lines[k]
                if re.search(r"CH\s*PRECISION\s*DAY\s*\d+", lk, re.IGNORECASE):
                    break

                if value is None:
                    # Result row: "***** ***** 2.5 g/dL 281.0098 Low" (unit optional)
                    vm = re.search(
                        r"\*+\s+\*+\s+([\d.]+)(?:\s+([A-Za-z\u00b5%][A-Za-z\u00b5/%]*))?",
                        lk,
                    )
                    if not vm:
                        # Fallback replicate row: "1 2.5 g/dL 281.0098 06/03/2026 6:56 PM"
                        vm = re.search(
                            r"^\d+\s+([\d.]+)\s+([A-Za-z\u00b5%][A-Za-z\u00b5/%]*)\s+[-\d.]",
                            lk,
                        )
                    if vm:
                        value = float(vm.group(1))
                        unit = vm.group(2) or ""

                if level is None:
                    lm = re.match(r"^L\s*(\d+)\b", lk) or re.search(r"\bL\s*(\d+)\s+\d", lk)
                    if lm:
                        level = int(lm.group(1))

                if not date:
                    dm = re.search(r"(\d{2}/\d{2}/\d{4})", lk)
                    if dm:
                        date = dm.group(1)

            if level is None:
                level = 1

            if value is not None:
                key = f"{current_analyte}|{day}|{level}"
                rep = rep_counter.get(key, 0) + 1
                if rep > max_reps:
                    continue
                rep_counter[key] = rep
                results.append(
                    ParsedResult(
                        analyte_name=current_analyte,
                        level=level,
                        day_number=day if precision_type == "inter" else 0,
                        replication=rep,
                        precision_type=precision_type,
                        value=value,
                        unit=unit,
                        date=date,
                    )
                )

    return results
