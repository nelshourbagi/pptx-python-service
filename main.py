import base64
import io
import re
import traceback
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from typing import Dict, List, Any, Optional, Tuple

from pptx import Presentation
from pptx.table import Table

app = FastAPI()

# ─────────────────────────────────────────────────────────────
# Global error handler (so Railway returns useful JSON)
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
    format_hint: Optional[str] = None  # <-- new


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
    """Normalize strings for matching: trim, collapse whitespace, keep case stable."""
    if s is None:
        return ""
    txt = str(s).replace("\u00A0", " ").strip()
    txt = re.sub(r"\s+", " ", txt)
    return txt


def _is_section_header(cell_text: str) -> bool:
    return _norm(cell_text).lower().startswith("net returns")


def _extract_trailing_paren_value(text: str) -> Optional[str]:
    m = re.search(r"\(([^)]*)\)\s*$", _norm(text))
    return m.group(1).strip() if m else None


def _update_cell_text_preserve_first_run(cell, new_text: str) -> None:
    """
    Preserve formatting by updating only the first run of the first paragraph,
    removing extra runs/paragraphs.
    """
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
    """
    Replace ONLY the trailing (date) part, even if original is split across runs.
    """
    if not cell.text_frame:
        return
    full_text = cell.text_frame.text or ""
    updated_text = re.sub(r"\([^)]*\)\s*$", f"({new_date})", _norm(full_text))
    if updated_text and updated_text != _norm(full_text):
        _update_cell_text_preserve_first_run(cell, updated_text)


def update_text_shape(shape, text_value: str):
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    if not tf.paragraphs:
        return

    para = tf.paragraphs[0]

    if para.runs:
        para.runs[0].text = text_value
        for run in para.runs[1:]:
            run._r.getparent().remove(run._r)
    else:
        para.text = text_value

    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)


# ─────────────────────────────────────────────────────────────
# Formatting (Option A)
# ─────────────────────────────────────────────────────────────

def _try_float(x: Any) -> Optional[float]:
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = _norm(x)
    if not s:
        return None
    # Remove commas for parsing, allow negatives
    s2 = s.replace(",", "")
    # If user accidentally included a % sign, strip it for float parsing
    s2 = s2.replace("%", "")
    try:
        return float(s2)
    except Exception:
        return None


def _format_value(raw: Any, format_hint: Optional[str]) -> str:
    """
    Applies presentation formatting.
    Option A: percent hints expect decimals (0.1532 => 15.32%)
    """
    # Default: if it's already a string, return as-is
    if format_hint is None:
        return _norm(raw)

    hint = _norm(format_hint).lower()
    if hint in ("", "raw", "none"):
        return _norm(raw)

    if hint == "text":
        return _norm(raw)

    val = _try_float(raw)
    if val is None:
        return _norm(raw)

    try:
        if hint == "percent_2dp":
            return f"{val * 100:.2f}%"
        if hint == "percent_1dp":
            return f"{val * 100:.1f}%"
        if hint == "number_0dp":
            return f"{val:,.0f}"
        if hint == "number_2dp":
            return f"{val:,.2f}"
        if hint == "currency_0dp":
            return f"${val:,.0f}"
        if hint == "currency_2dp":
            return f"${val:,.2f}"
    except Exception:
        # Any unexpected formatting issue falls back to raw
        return _norm(raw)

    # Unknown hint -> raw
    return _norm(raw)


# ─────────────────────────────────────────────────────────────
# Header detection / matching
# ─────────────────────────────────────────────────────────────

def _header_matches(canonical_h: str, ppt_h: str) -> bool:
    """Partial-friendly match with parenthetical stripping."""
    ch = _norm(canonical_h).lower()
    ph = _norm(ppt_h).lower()
    if not ch or not ph:
        return False
    if ch == ph or ch in ph or ph in ch:
        return True
    ch2 = re.sub(r"\s*\([^)]*\)\s*$", "", ch).strip()
    ph2 = re.sub(r"\s*\([^)]*\)\s*$", "", ph).strip()
    return bool(ch2 and ph2 and (ch2 == ph2 or ch2 in ph2 or ph2 in ch2))


def _score_header_row(table: Table, row_idx: int, canonical_headers: List[str]) -> Tuple[int, List[str]]:
    """Return (score, headers_list) for candidate header row. Headers list is normalized text from cols 1..end."""
    if row_idx < 0 or row_idx >= len(table.rows):
        return (0, [])
    ppt_headers = [_norm(c.text) for c in table.rows[row_idx].cells[1:]]  # skip col0
    score = 0
    for ch in canonical_headers:
        for ph in ppt_headers:
            if _header_matches(ch, ph):
                score += 1
                break
    return (score, ppt_headers)


def _detect_header_row_idx(table: Table, sh_row_idx: int, canonical_headers: List[str], scan_depth: int = 4) -> Tuple[int, int, List[str]]:
    """
    Scan rows sh_row_idx+1..sh_row_idx+scan_depth and choose highest score.
    Returns (best_row_idx, best_score, chosen_headers).
    """
    best_idx = min(sh_row_idx + 1, len(table.rows) - 1)
    best_score = 0
    best_headers: List[str] = []

    last_row = min(len(table.rows) - 1, sh_row_idx + scan_depth)
    for ridx in range(sh_row_idx + 1, last_row + 1):
        score, headers = _score_header_row(table, ridx, canonical_headers)
        if score > best_score:
            best_score = score
            best_idx = ridx
            best_headers = headers

    if not best_headers and 0 <= best_idx < len(table.rows):
        best_headers = [_norm(c.text) for c in table.rows[best_idx].cells[1:]]
    return (best_idx, best_score, best_headers)


# ─────────────────────────────────────────────────────────────
# Table updater (section-aware + formatting)
# ─────────────────────────────────────────────────────────────

def update_table_shape(table: Table, canonical_table: CanonicalTable, errors: List[str], binding_id: str) -> None:
    """
    Section-aware table updater:
    - Detect section header rows (col 0 starts with "Net Returns")
    - Resolve sections Top/Bottom etc
    - Detect correct column-header row by scanning rows after each section header
    - Update cells by row_label (col 0) + header match (col headers)
    - Apply formatting per-cell using CanonicalCell.format_hint
    """

    if len(table.rows) < 2:
        errors.append(f"Table '{binding_id}' has fewer than 2 rows")
        return

    total_rows = len(table.rows)

    # Find section header rows in PPT
    section_header_indices: List[int] = []
    for ridx in range(total_rows):
        t0 = _norm(table.rows[ridx].cells[0].text)
        if _is_section_header(t0):
            section_header_indices.append(ridx)

    has_sections = (
        len(section_header_indices) > 0
        and canonical_table.sections is not None
        and len(canonical_table.sections) > 0
    )

    # --- Flat fallback (no section handling)
    if not has_sections:
        headers = [_norm(c.text) for c in table.rows[0].cells]
        if len(headers) < 2:
            errors.append(f"Table '{binding_id}': expected at least 2 columns")
            return

        col_lookup: Dict[str, int] = {}
        for idx in range(1, len(headers)):
            h = _norm(headers[idx])
            if h:
                col_lookup[h] = idx

        canonical_row_lookup = {_norm(r.row_key): r for r in canonical_table.rows}

        updated = 0
        for ridx in range(1, total_rows):
            row = table.rows[ridx]
            row_label = _norm(row.cells[0].text)
            if not row_label:
                continue
            cr = canonical_row_lookup.get(row_label)
            if not cr:
                continue

            for ppt_h, col_idx in col_lookup.items():
                # find canonical cell whose header matches this ppt_h
                chosen: Optional[CanonicalCell] = None
                for ck, cv in cr.cells.items():
                    if _header_matches(ck, ppt_h):
                        chosen = cv
                        break
                if not chosen:
                    continue

                new_text = _format_value(chosen.raw, chosen.format_hint)
                cell = row.cells[col_idx]
                _update_cell_text_preserve_first_run(cell, new_text)
                updated += 1

        if updated == 0:
            errors.append(f"DEBUG {binding_id}: 0 cells updated (flat).")
        return

    # --- Section-aware update
    section_dates: Dict[str, str] = {}
    if canonical_table.meta and canonical_table.meta.section_dates:
        section_dates = canonical_table.meta.section_dates

    positional_labels = ["Top", "Bottom", "Section3", "Section4"]
    resolved_sections: List[str] = []

    # Resolve section labels by matching date (if possible), else by position
    for idx, sh_row_idx in enumerate(section_header_indices):
        cell_text = _norm(table.rows[sh_row_idx].cells[0].text)
        header_date = _extract_trailing_paren_value(cell_text)

        resolved = None
        if header_date and section_dates:
            for sec_label, sec_date in section_dates.items():
                if _norm(sec_date) == _norm(header_date):
                    resolved = sec_label
                    break

        if not resolved:
            resolved = positional_labels[idx] if idx < len(positional_labels) else f"Section{idx+1}"

        resolved_sections.append(resolved)

    updated_cells_count = 0
    section_debug_info: Dict[str, Any] = {}

    for sec_idx, sh_row_idx in enumerate(section_header_indices):
        section_label = resolved_sections[sec_idx]
        sec_data = canonical_table.sections.get(section_label) if canonical_table.sections else None

        # Update date in section header cell
        if section_label in section_dates:
            _update_section_header_date(table.rows[sh_row_idx].cells[0], section_dates[section_label])

        # Determine end of this section
        next_sh = section_header_indices[sec_idx + 1] if sec_idx + 1 < len(section_header_indices) else total_rows

        if not sec_data:
            continue

        # Detect header row for this section (scan after section title row)
        chosen_header_row_idx, chosen_header_score, chosen_headers = _detect_header_row_idx(
            table, sh_row_idx, sec_data.headers, scan_depth=4
        )

        # Data starts after header row
        data_row_start = chosen_header_row_idx + 1

        # Build lookup of PPT header -> column index (cols 1..)
        sec_col_lookup: Dict[str, int] = {}
        for i, h in enumerate(chosen_headers, start=1):
            hn = _norm(h)
            if hn:
                sec_col_lookup[hn] = i

        # Row lookup by normalized name
        sec_row_lookup = {_norm(r.row_key): r for r in sec_data.rows}

        # Debug samples
        ppt_rows_sample: List[str] = []
        for ridx in range(data_row_start, next_sh):
            if ridx >= total_rows:
                break
            rl = _norm(table.rows[ridx].cells[0].text)
            if rl and not _is_section_header(rl):
                ppt_rows_sample.append(rl)
            if len(ppt_rows_sample) >= 8:
                break

        section_debug_info[section_label] = {
            "chosen_header_row_idx": chosen_header_row_idx,
            "chosen_header_score": chosen_header_score,
            "chosen_headers(sample)": chosen_headers[:10],
            "ppt_rows(sample)": ppt_rows_sample[:8],
            "canonical_rows(sample)": [ _norm(r.row_key) for r in sec_data.rows[:8] ],
            "canonical_headers(sample)": [ _norm(h) for h in sec_data.headers[:10] ],
        }

        # Helper: find ppt column index for a canonical header
        def _find_col_idx(canonical_header: str) -> Optional[int]:
            ch = _norm(canonical_header)
            # exact match
            if ch in sec_col_lookup:
                return sec_col_lookup[ch]
            # partial match
            for ppt_h, idx_col in sec_col_lookup.items():
                if _header_matches(ch, ppt_h):
                    return idx_col
            return None

        # Update cells
        for ridx in range(data_row_start, next_sh):
            if ridx >= total_rows:
                break
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

                new_text = _format_value(cv.raw, cv.format_hint)
                cell = row.cells[col_idx]
                _update_cell_text_preserve_first_run(cell, new_text)
                updated_cells_count += 1

    if updated_cells_count == 0:
        errors.append(
            f"DEBUG {binding_id}: 0 cells updated (sectioned). "
            f"section_headers={section_header_indices} "
            f"resolved={resolved_sections} "
            f"header_choice={section_debug_info}"
        )


# ─────────────────────────────────────────────────────────────
# Endpoints
# ─────────────────────────────────────────────────────────────

@app.get("/health")
async def health():
    return {"status": "ok"}

@app.get("/version")
async def version():
    return {"version": "format-v1"}

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
            alt_text = _norm(get_shape_alt_text(shape))
            if not alt_text:
                continue

            if alt_text in text_bindings:
                if alt_text in request.canonical_data.kpis:
                    kpi = request.canonical_data.kpis[alt_text]
                    # Apply formatting at last mile
                    rendered = _format_value(kpi.raw, kpi.format_hint)
                    update_text_shape(shape, rendered)
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

