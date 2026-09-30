"""PDF parser for the Lab Instrument Validation Tool.

Supports the existing CH PRECISION format plus Siemens/Atellica Assay Report
formats for both INTER RUN and INTRA PRECISION.

The numeric Result is extracted; RLU/MAU is never used as the result.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Optional

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


ANALYTE_ANCHOR = re.compile(r"Order\s*Time:.*Analyzer:\s*All", re.IGNORECASE)
SKIP_ANALYTES = {"H", "I", "L"}

# Supports:
# INTER RUN DAY-2 ...
# INTER RUN L1-DAY-2 ...
# INTER RUN L-1-DAY-2 ...
INTER_ROW_RE = re.compile(
    r"^INTER\s+RUN\s+"
    r"(?:L\s*-?\s*(\d+)\s*-\s*)?"
    r"DAY\s*-\s*(\d+)\s+"
    r"([-+]?(?:\d+(?:\.\d*)?|\.\d+))\s+"
    r"([A-Za-zµμ%]+(?:/[A-Za-zµμ%]+)?)\s+"
    r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)\s+"
    r"(\d{2}/\d{2}/\d{4})"
    r"(?:\s+(\d{1,2}:\d{2}\s*[AP]M))?\s*$",
    re.IGNORECASE,
)

INTRA_ROW_RE = re.compile(
    r"^INTRA\s+PRECISION\s+L\s*-\s*(\d+)\s+"
    r"([-+]?(?:\d+(?:\.\d*)?|\.\d+))\s+"
    r"([A-Za-zµμ%]+(?:/[A-Za-zµμ%]+)?)\s+"
    r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)\s+"
    r"(\d{2}/\d{2}/\d{4})"
    r"(?:\s+(\d{1,2}:\d{2}\s*[AP]M))?\s*$",
    re.IGNORECASE,
)


def extract_text_from_pdf(file) -> str:
    if isinstance(file, (bytes, bytearray)):
        file = io.BytesIO(file)

    full_text: list[str] = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            words = page.extract_words(use_text_flow=False)
            line_map: dict[int, list[tuple[float, str]]] = {}

            for w in words:
                word = w.get("text", "").strip()
                if not word:
                    continue
                y = round(float(w["top"]) / 2) * 2
                line_map.setdefault(y, []).append((float(w["x0"]), word))

            for y in sorted(line_map):
                items = sorted(line_map[y], key=lambda item: item[0])
                line = " ".join(text for _, text in items).strip()
                if line:
                    full_text.append(line)
            full_text.append("")

    return "\n".join(full_text)


def _normalize_lines(text: str) -> list[str]:
    raw = [line.strip() for line in text.splitlines() if line.strip()]
    lines: list[str] = []
    i = 0

    while i < len(raw):
        line = raw[i]

        if (
            re.fullmatch(r"\d{2}/\d{2}/\d{4}", line)
            and i + 2 < len(raw)
            and re.match(
                r"^(?:INTER\s+RUN|INTRA\s+PRECISION)",
                raw[i + 1],
                re.I,
            )
            and re.fullmatch(r"\d{1,2}:\d{2}\s*[AP]M", raw[i + 2], re.I)
        ):
            lines.append(f"{raw[i + 1]} {line} {raw[i + 2]}")
            i += 3
            continue

        if (
            re.search(r"\d{2}/\d{2}/\d{4}$", line)
            and i + 1 < len(raw)
            and re.fullmatch(r"\d{1,2}:\d{2}\s*[AP]M", raw[i + 1], re.I)
        ):
            line += " " + raw[i + 1]
            i += 1

        lines.append(line)
        i += 1

    return lines


def _is_new_assay_report(text: str) -> bool:
    return bool(
        re.search(r"\bINTER\s+RUN\b", text, re.I)
        or re.search(r"\bINTRA\s+PRECISION\b", text, re.I)
    )


def _looks_like_result_header(line: str) -> bool:
    return bool(
        re.search(r"\bPatient\s+Name\b", line, re.I)
        and re.search(r"\bResult\b", line, re.I)
    )


def _looks_like_non_analyte(line: str) -> bool:
    return bool(
        re.search(
            r"^(Assay Report|Software Version|Order Time|Printed:|Max Lab|Page\s+\d+)",
            line,
            re.I,
        )
        or re.search(r"\bPatient\s+Name\b", line, re.I)
    )


def _find_analyte_before_header(lines: list[str], header_idx: int) -> str:
    for j in range(header_idx - 1, max(-1, header_idx - 5), -1):
        candidate = lines[j].strip()
        if not candidate or _looks_like_non_analyte(candidate):
            continue
        if re.match(r"^(Assay Report|Software Version|Order Time)", candidate, re.I):
            continue
        # The analyte is the standalone label immediately before the table header.
        return re.sub(r"\s+", "", candidate)
    return ""


def _parse_new_assay_report(text: str, precision_type: str) -> list[ParsedResult]:
    lines = _normalize_lines(text)
    results: list[ParsedResult] = []
    current_analyte = ""
    rep_counter: dict[tuple[str, int, str], int] = {}

    for i, line in enumerate(lines):
        # Generic analyte detection: use the report table structure, not a
        # fixed analyte-name pattern such as *_A.
        if _looks_like_result_header(line):
            candidate = _find_analyte_before_header(lines, i)
            if candidate:
                current_analyte = candidate
            continue

        if precision_type.lower() == "inter":
            match = INTER_ROW_RE.match(line)
            if not match or not current_analyte:
                continue

            explicit_level = match.group(1)
            day = int(match.group(2))
            value = float(match.group(3))
            unit = match.group(4)
            date = match.group(5)

            level = int(explicit_level) if explicit_level else day
            key = (current_analyte, day, f"inter-{level}")
            replication = rep_counter.get(key, 0) + 1
            rep_counter[key] = replication

            if replication > 25:
                continue

            results.append(
                ParsedResult(
                    analyte_name=current_analyte,
                    level=level,
                    day_number=day,
                    replication=replication,
                    precision_type="inter",
                    value=value,
                    unit=unit,
                    date=date,
                )
            )

        else:
            match = INTRA_ROW_RE.match(line)
            if not match or not current_analyte:
                continue

            level = int(match.group(1))
            value = float(match.group(2))
            unit = match.group(3)
            date = match.group(4)
            key = (current_analyte, level, "intra")
            replication = rep_counter.get(key, 0) + 1
            rep_counter[key] = replication

            if replication > 25:
                continue

            results.append(
                ParsedResult(
                    analyte_name=current_analyte,
                    level=level,
                    day_number=0,
                    replication=replication,
                    precision_type="intra",
                    value=value,
                    unit=unit,
                    date=date,
                )
            )

    return results


def _read_analyte_name(lines: list[str], anchor_idx: int) -> str:
    name = ""
    for j in range(anchor_idx + 1, min(anchor_idx + 5, len(lines))):
        if re.match(r"^Result$", lines[j], re.IGNORECASE) or re.match(
            r"^Patient\s+Name", lines[j], re.IGNORECASE
        ):
            break
        name += lines[j]
    return re.sub(r"\s+", "", name)


def _parse_old_ch_precision(text: str, precision_type: str) -> list[ParsedResult]:
    results: list[ParsedResult] = []
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    max_reps = 25
    current_analyte = ""
    rep_counter: dict[str, int] = {}

    for i, line in enumerate(lines):
        if ANALYTE_ANCHOR.search(line):
            current_analyte = _read_analyte_name(lines, i)
            continue

        day_m = re.search(r"CH\s*PRECISION\s*DAY\s*(\d+)", line, re.I)
        if not day_m or not current_analyte or current_analyte in SKIP_ANALYTES:
            continue

        day = int(day_m.group(1))
        level: Optional[int] = None
        value: Optional[float] = None
        unit = ""
        date_m = re.search(r"(\d{2}/\d{2}/\d{4})", line)
        date = date_m.group(1) if date_m else ""

        same_line_level = re.search(r"DAY\s*\d+\s+L\s*(\d+)", line, re.I)
        if same_line_level:
            level = int(same_line_level.group(1))

        for k in range(i + 1, min(i + 9, len(lines))):
            lk = lines[k]
            if re.search(r"CH\s*PRECISION\s*DAY\s*\d+", lk, re.I):
                break

            if value is None:
                vm = re.search(
                    r"\*+\s+\*+\s+([-+]?[\d.]+)(?:\s+([A-Za-zµ%][A-Za-zµ/%]*))?",
                    lk,
                )
                if not vm:
                    vm = re.search(
                        r"^\d+\s+([-+]?[\d.]+)\s+([A-Za-zµ%][A-Za-zµ/%]*)\s+[-\d.]",
                        lk,
                    )
                if vm:
                    value = float(vm.group(1))
                    unit = vm.group(2) or ""

            if level is None:
                lm = re.match(r"^L\s*(\d+)\b", lk) or re.search(
                    r"\bL\s*(\d+)\s+\d", lk
                )
                if lm:
                    level = int(lm.group(1))

            if not date:
                dm = re.search(r"(\d{2}/\d{2}/\d{4})", lk)
                if dm:
                    date = dm.group(1)

        if level is None:
            level = 1
        if value is None:
            continue

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


def parse_analyzer_text(text: str, precision_type: str = "inter") -> list[ParsedResult]:
    precision_type = precision_type.lower().strip()

    if precision_type not in {"inter", "intra"}:
        raise ValueError("precision_type must be 'inter' or 'intra'")

    if _is_new_assay_report(text):
        return _parse_new_assay_report(text, precision_type)

    return _parse_old_ch_precision(text, precision_type)
