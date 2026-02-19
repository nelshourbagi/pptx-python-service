import base64
import io
import re
import traceback
from typing import Dict, List, Any, Optional, Tuple

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from pptx import Presentation
from pptx.table import Table

app = FastAPI()


# ─────────────────────────────────────────────────────────────
# Global error handler (so Railway returns JSON, not plain 500)
# ─────────────────────────────────────────────────────────────

@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    return JSONResponse(
        status_code=500,
        content={
            "error": str(exc),
            "type": exc.__class__.__name__,
            "traceback": traceback.format_exc(),
        },
    )


# ─────────────────────────────────────────────────────────────
# Request / Response Models
# ─────────────────────────────────────────────────────────────

class CanonicalCell(BaseModel):
    raw: Any
    formatted: str


class CanonicalRow(BaseModel):
    row_key: str
    cells: Dict[str, CanonicalCell]


class CanonicalTableSection(BaseModel):
    headers: List[str]
    rows: List[CanonicalRow]


class CanonicalTableMeta(BaseModel):
    # e.g. { "Top": "Nov 30, 2025", "Bottom": "Dec 31, 2025" }
    section_dates: Optional[Dict[str, str]] = None


class CanonicalTable(BaseModel):
    headers: List[str]
    rows: List[CanonicalRow]
    sections: Optional[Dict[str, CanonicalTableSection]] = None
    meta: Optional[CanonicalTableMeta] = None


class CanonicalData(BaseModel):
    meta: Dict[str, Any]
    kpis: Dict[str, CanonicalCell]
    tables: Dict[str, CanonicalTable]


class Binding(BaseModel):
    id: str
    type: str


class MappingConfig(BaseModel):
    bindings: List[Binding]


class ProcessRequest(BaseModel):
    template_base64: str
    canonical_data: CanonicalData
    mapping_config: MappingConfig


class ProcessResponse(BaseModel):
    output_base64: str
    errors: List[str]


# ─────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────

def get_shape_alt_text(shape) -> str:
    try:
        desc = shape._element.xpath(".//*[local-name()='cNvPr']/@descr")
        if desc:
            return desc[0]
    except Exception:
        pass

    try:
        return shape.name or ""
    except Exception:
        return ""


def _norm(s: Any) -> str:
    if s is None:
        return ""
    s = str(s)

    # Replace non-breaking spaces etc.
    s = s.replace("\u00A0", " ")

    # Drop footnote/superscript-ish trailing digits if they are appended to a word
    # e.g. "Index2" -> "Index"
    s = re.sub(r"([A-Za-z])(\d+)\s*$", r"\1", s)

    # Collapse whitespace
    s = re.sub(r"\s+", " ", s).strip()
    return s


def update_text_shape(shape, formatted_value: str):
    if not shape.has_text_frame:
        return

    tf = shape.text_frame
    if not tf.paragraphs:
        return

    para = tf.paragraphs[0]

    if para.runs:
        para.runs[0].text = formatted_value
        for run in para.runs[1:]:
            run._r.getparent().remove(run._r)
    else:
        para.text = formatted_value

    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)


def _is_section_header(cell_text: str) -> bool:
    return _norm(cell_text).lower().startswith("net returns")


def _extract_trailing_paren_value(text: str) -> Optional[str]:
    m = re.search(r"\(([^)]*)\)\s*$", str(text).strip())
    return m.group(1).strip() if m else None


def _update_cell_text_preserve_first_run(cell, new_text: str) -> None:
    if not cell.text_frame or not cell.text_frame.paragraphs:
        return
    tf = cell.text_frame
    p0 = tf.paragraphs[0]
    if p0.runs:
        p0.runs[0].text = new_text
        for r in p0.runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p0.text = new_text

    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)


def _update_section_header_date(cell, new_date: str) -> None:
    if not cell.text_frame:
        return
    full_text = cell.text_frame.text or ""
    updated_text = re.sub(r"\([^)]*\)\s*$", f"({new_date})", full_text.strip())
    if updated_text != full_text.strip():
        _update_cell_text_preserve_first_run(cell, updated_text)


def _header_matches(canonical_h: str, ppt_h: str) -> bool:
    """
    Partial-friendly header matching with parenthetical qualifier stripping.
    """
    ch = _norm(canonical_h).lower()
    ph = _norm(ppt_h).lower()
    if not ch or not ph:
        return False
    if ch == ph:
        return True
    if ch in ph or ph in ch:
        return True

    # Strip trailing parenthetical qualifiers and retry
    ch2 = re.sub(r"\s*\([^)]*\)\s*$", "", ch).strip()
    ph2 = re.sub(r"\s*\([^)]*\)\s*$", "", ph).strip()
    if not ch2 or not ph2:
        return False
    return (ch2 == ph2) or (ch2 in ph2) or (ph2 in ch2)


def _score_header_row(table: Table, row_idx: int, canonical_headers: List[str]) -> Tuple[int, List[str]]:
    """
    Score a candidate header row. Returns (score, raw_header_texts_list).
    Header texts read from columns 1..end (skip col0).
    """
    if row_idx < 0 or row_idx >= len(table.rows):
        return (0, [])

    ppt_headers = [table.rows[row_idx].cells[i].text for i in range(1, len(table.rows[row_idx].cells))]
    score = 0

    for ch in canonical_headers:
        for ph in ppt_headers:
            if _header_matches(ch, ph):
                score += 1
                break

    return (score, ppt_headers)


def _detect_header_row_idx(table: Table, sh_row_idx: int, canonical_headers: List[str], scan_depth: int = 4) -> Tuple[int, int, List[str]]:
    """
    IMPORTANT: in your PPT, the header row is often the SAME row as the section header (Net Returns ...),
    because col0 has the title/date, and col1..n have the column headers.

    So we scan:
      sh_row_idx .. sh_row_idx+scan_depth
    and choose the highest scoring row.

    Returns:
      (best_row_idx, best_score, chosen_header_texts)
    """
    best_idx = sh_row_idx
    best_score = -1
    best_headers: List[str] = []

    last_row = min(len(table.rows) - 1, sh_row_idx + scan_depth)
    for ridx in range(sh_row_idx, last_row + 1):
        score, headers = _score_header_row(table, ridx, canonical_headers)
        if score > best_score:
            best_score = score
            best_idx = ridx
            best_headers = headers

    # If nothing matched at all, fallback to sh_row_idx (safer than sh_row_idx+1 for your layout)
    if best_score <= 0:
        fallback_score, fallback_headers = _score_header_row(table, sh_row_idx, canonical_headers)
        return (sh_row_idx, fallback_score, fallback_headers)

    return (best_idx, best_score, best_headers)


def _try_parse_float(x: Any) -> Optional[float]:
    if x is None:
        return None
    s = str(x).strip()
    if s == "":
        return None
    # allow commas in data
    s2 = s.replace(",", "")
    try:
        return float(s2)
    except Exception:
        return None


def _count_decimals_from_text(text: str, default: int = 2) -> int:
    """
    Infer decimals from an existing cell text like "12.34%" or "1,189".
    """
    t = text.strip()
    # remove symbols
    t = t.replace("%", "")
    t = re.sub(r"[^\d\.\-]", "", t)
    if "." in t:
        return len(t.split(".")[-1])
    return default


def _format_number_like_ppt(existing_text: str, raw_value: Any) -> str:
    """
    Format using the existing PPT cell text as the 'style guide'.

    Rules:
    - If existing has '%': treat as percent
      - if input is 0.5544 -> 55.44% (assume fraction)
      - if input is 55.44 -> 55.44% (assume already percent)
    - If existing has currency symbol: add symbol, commas, inferred decimals
    - If existing has commas or looks like integer: commas, inferred decimals
    - Otherwise: return raw as string
    """
    ex = (existing_text or "").strip()

    # preserve dash style
    if raw_value is None:
        return ex if ex else "-"
    rv_str = str(raw_value).strip()
    if rv_str == "":
        return ex if ex else "-"

    num = _try_parse_float(raw_value)
    if num is None:
        # non-numeric values: just write as-is
        return rv_str

    # percent
    if "%" in ex:
        decimals = _count_decimals_from_text(ex, default=2)

        # Heuristic: if abs(num) <= 1.5 -> likely fraction; else likely already percent
        pct = num * 100.0 if abs(num) <= 1.5 else num
        return f"{pct:.{decimals}f}%"

    # currency symbols
    currency_match = re.search(r"[\$\£\€\¥]", ex)
    if currency_match:
        symbol = currency_match.group(0)
        decimals = _count_decimals_from_text(ex, default=2)
        return f"{symbol}{num:,.{decimals}f}"

    # numbers with commas / plain numbers
    if "," in ex:
        decimals = _count_decimals_from_text(ex, default=0)
        return f"{num:,.{decimals}f}"

    # integer-looking
    if re.fullmatch(r"-?\d+", ex):
        return f"{int(round(num))}"

    # default numeric
    decimals = _count_decimals_from_text(ex, default=2)
    return f"{num:.{decimals}f}"


def _set_cell_text_preserve_style(cell, new_text: str) -> None:
    """
    Preserve formatting by updating only the first run of the first paragraph.
    """
    if cell.text_frame and cell.text_frame.paragraphs:
        p = cell.text_frame.paragraphs[0]
        if p.runs:
            p.runs[0].text = new_text
        else:
            p.text = new_text


def update_table_shape(table: Table, canonical_table: CanonicalTable, errors: List[str], binding_id: str):
    """
    Section-aware table updater.

    Your specific table layout:
    - The "Net Returns (USD) (...)" row is ALSO the column header row,
      because col0 has the title/date and col1..n contain the headers.
    - So we detect the header row by scanning sh_row_idx..sh_row_idx+4.

    Updates:
    - Match row labels in col0 against canonical row_key (normalized)
    - Match column headers in the chosen header row against canonical col_key (partial-friendly)
    - Apply formatting based on existing PPT cell text (percent/currency/commas)
    """

    if len(table.rows) < 2:
        errors.append(f"Table '{binding_id}' has fewer than 2 rows")
        return

    total_rows = len(table.rows)

    # Detect section header row indices
    section_header_indices: List[int] = []
    for row_idx in range(total_rows):
        cell0_text = table.rows[row_idx].cells[0].text
        if _is_section_header(cell0_text):
            section_header_indices.append(row_idx)

    has_sections = (
        len(section_header_indices) > 0
        and canonical_table.sections
        and len(canonical_table.sections) > 0
    )

    updated_cells_count = 0
    section_debug_info: Dict[str, Any] = {}

    # ── Flat fallback (no sections detected) ───────────────────
    if not has_sections:
        headers = [table.rows[0].cells[i].text for i in range(len(table.rows[0].cells))]
        if len(headers) < 2:
            errors.append(f"Table '{binding_id}': expected at least 2 columns")
            return

        col_lookup: Dict[str, int] = {}
        for idx in range(1, len(headers)):
            h = headers[idx]
            if _norm(h):
                col_lookup[_norm(h)] = idx

        canonical_row_lookup = {_norm(r.row_key): r for r in canonical_table.rows}

        for row_idx in range(1, total_rows):
            row = table.rows[row_idx]
            row_label = _norm(row.cells[0].text)
            if not row_label:
                continue

            cr = canonical_row_lookup.get(row_label)
            if not cr:
                continue

            for ck, cv in cr.cells.items():
                target_col_idx = None
                ck_norm = _norm(ck)
                if ck_norm in col_lookup:
                    target_col_idx = col_lookup[ck_norm]
                else:
                    for ppt_h_norm, idx2 in col_lookup.items():
                        if _header_matches(ck_norm, ppt_h_norm):
                            target_col_idx = idx2
                            break

                if target_col_idx is None:
                    continue

                cell = row.cells[target_col_idx]
                existing = cell.text
                new_text = _format_number_like_ppt(existing, cv.raw)
                _set_cell_text_preserve_style(cell, new_text)
                updated_cells_count += 1

        if updated_cells_count == 0:
            errors.append(f"DEBUG {binding_id}: 0 cells updated (flat)")
        return

    # ── Section-aware update ───────────────────────────────────
    section_dates: Dict[str, str] = {}
    if canonical_table.meta and canonical_table.meta.section_dates:
        section_dates = canonical_table.meta.section_dates

    # Resolve section labels in order (Top, Bottom, ...)
    positional_labels = ["Top", "Bottom", "Section3", "Section4"]
    resolved_sections: List[str] = []
    for idx, _ in enumerate(section_header_indices):
        resolved_sections.append(positional_labels[idx] if idx < len(positional_labels) else f"Section{idx+1}")

    for sec_idx, sh_row_idx in enumerate(section_header_indices):
        section_label = resolved_sections[sec_idx]
        sec_data = canonical_table.sections.get(section_label) if canonical_table.sections else None

        # Update date in section header (col0) if provided
        if section_label in section_dates:
            _update_section_header_date(table.rows[sh_row_idx].cells[0], section_dates[section_label])

        if not sec_data:
            continue

        # Determine where this section ends
        next_sh = section_header_indices[sec_idx + 1] if (sec_idx + 1) < len(section_header_indices) else total_rows

        # Detect the correct header row for this section (can be same as sh_row_idx)
        chosen_header_row_idx, chosen_header_score, chosen_headers = _detect_header_row_idx(
            table=table,
            sh_row_idx=sh_row_idx,
            canonical_headers=sec_data.headers,
            scan_depth=4,
        )

        # Data rows start after the chosen header row
        data_row_start = chosen_header_row_idx + 1

        # Build column lookup using the chosen header row texts
        sec_col_lookup: Dict[str, int] = {}
        # chosen_headers corresponds to columns 1..n
        for offset, h in enumerate(chosen_headers, start=1):
            hn = _norm(h)
            if hn:
                sec_col_lookup[hn] = offset

        def _find_col_idx(canonical_header: str) -> Optional[int]:
            ch_norm = _norm(canonical_header)
            if ch_norm in sec_col_lookup:
                return sec_col_lookup[ch_norm]
            for ppt_h_norm, idx2 in sec_col_lookup.items():
                if _header_matches(ch_norm, ppt_h_norm):
                    return idx2
            return None

        sec_row_lookup = {_norm(r.row_key): r for r in sec_data.rows}

        # Collect sample row labels for debugging
        ppt_rows_sample: List[str] = []
        for ridx in range(data_row_start, min(next_sh, data_row_start + 12)):
            rl = _norm(table.rows[ridx].cells[0].text)
            if rl and not _is_section_header(rl):
                ppt_rows_sample.append(rl)

        # Update cells
        for ridx in range(data_row_start, next_sh):
            row = table.rows[ridx]
            row_label = _norm(row.cells[0].text)
            if not row_label or _is_section_header(row_label):
                continue

            cr = sec_row_lookup.get(row_label)
            if not cr:
                continue

            for ck, cv in cr.cells.items():
                col_idx = _find_col_idx(ck)
                if col_idx is None:
                    continue

                cell = row.cells[col_idx]
                existing = cell.text
                new_text = _format_number_like_ppt(existing, cv.raw)
                _set_cell_text_preserve_style(cell, new_text)
                updated_cells_count += 1

        # Debug per section if needed
        section_debug_info[section_label] = {
            "chosen_header_row_idx": chosen_header_row_idx,
            "chosen_header_score": chosen_header_score,
            "chosen_headers(sample)": [_norm(x) for x in chosen_headers][:10],
            "ppt_rows(sample)": ppt_rows_sample[:8],
            "canonical_rows(sample)": [_norm(r.row_key) for r in sec_data.rows][:8],
            "canonical_headers(sample)": [_norm(h) for h in sec_data.headers][:10],
        }

    # Emit debug only when no updates happened
    if updated_cells_count == 0:
        errors.append(
            f"DEBUG {binding_id}: 0 cells updated (sectioned). "
            f"section_headers={section_header_indices} resolved={resolved_sections} "
            f"header_choice={section_debug_info} version=fmt-v1"
        )


# ─────────────────────────────────────────────────────────────
# Endpoints
# ─────────────────────────────────────────────────────────────

@app.post("/", response_model=ProcessResponse)
async def process_pptx(request: ProcessRequest):
    errors: List[str] = []

    try:
        template_bytes = base64.b64decode(request.template_base64)
        prs = Presentation(io.BytesIO(template_bytes))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid PPTX template: {e}")

    text_bindings = {b.id for b in request.mapping_config.bindings if b.type == "text"}
    table_bindings = {b.id for b in request.mapping_config.bindings if b.type == "table"}

    for slide in prs.slides:
        for shape in slide.shapes:
            alt_text = get_shape_alt_text(shape).strip()
            if not alt_text:
                continue

            if alt_text in text_bindings:
                if alt_text in request.canonical_data.kpis:
                    kpi = request.canonical_data.kpis[alt_text]
                    update_text_shape(shape, str(kpi.formatted))
                else:
                    errors.append(f"KPI binding '{alt_text}' not found")

            elif alt_text in table_bindings:
                if not shape.has_table:
                    errors.append(f"Binding '{alt_text}' is not a table")
                    continue

                if alt_text in request.canonical_data.tables:
                    tbl = request.canonical_data.tables[alt_text]
                    update_table_shape(shape.table, tbl, errors, alt_text)
                else:
                    errors.append(f"Table binding '{alt_text}' not found")

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)

    return ProcessResponse(
        output_base64=base64.b64encode(output.read()).decode("utf-8"),
        errors=errors
    )


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.get("/version")
async def version():
    return {"version": "fmt-v1"}

