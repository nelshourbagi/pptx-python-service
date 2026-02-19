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

VERSION = "header-scan-v4"


# ─────────────────────────────────────────────────────────────
# Global exception handler (Railway returns JSON, not plain 500)
# ─────────────────────────────────────────────────────────────
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    return JSONResponse(
        status_code=500,
        content={
            "error": str(exc),
            "type": exc.__class__.__name__,
            "traceback": traceback.format_exc(),
            "version": VERSION,
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

def _norm(s: Any) -> str:
    if s is None:
        return ""
    return str(s).strip()


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
    # Section header row begins with "Net Returns ..."
    return _norm(cell_text).lower().startswith("net returns")


def _extract_trailing_paren_value(text: str) -> Optional[str]:
    m = re.search(r"\(([^)]*)\)\s*$", _norm(text))
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
    ch = _norm(canonical_h).lower()
    ph = _norm(ppt_h).lower()
    if not ch or not ph:
        return False
    if ch == ph:
        return True
    if ch in ph or ph in ch:
        return True

    # Strip trailing "(...)" and retry
    ch2 = re.sub(r"\s*\([^)]*\)\s*$", "", ch).strip()
    ph2 = re.sub(r"\s*\([^)]*\)\s*$", "", ph).strip()
    return bool(ch2 and ph2 and (ch2 == ph2 or ch2 in ph2 or ph2 in ch2))


def _score_header_row(table: Table, row_idx: int, canonical_headers: List[str]) -> Tuple[int, List[str]]:
    """
    Returns (score, headers_list) where headers_list are texts from columns 1..end.
    NOTE: python-pptx cells do NOT support slicing; index iteration only.
    """
    if row_idx < 0 or row_idx >= len(table.rows):
        return (0, [])

    row = table.rows[row_idx]
    headers: List[str] = []
    for ci in range(1, len(row.cells)):
        headers.append(_norm(row.cells[ci].text))

    score = 0
    for ch in canonical_headers:
        for ph in headers:
            if _header_matches(ch, ph):
                score += 1
                break

    return (score, headers)


def _detect_header_row_idx(table: Table, sh_row_idx: int, canonical_headers: List[str], scan_depth: int = 4) -> Tuple[int, int, List[str]]:
    """
    IMPORTANT FIX: include sh_row_idx itself as candidate header row.
    Scan rows sh_row_idx..sh_row_idx+scan_depth and choose the highest score.
    Returns (best_row_idx, best_score, chosen_headers)
    """
    best_idx = sh_row_idx
    best_score = 0
    best_headers: List[str] = []

    last_row = min(len(table.rows) - 1, sh_row_idx + scan_depth)
    for ridx in range(sh_row_idx, last_row + 1):
        score, headers = _score_header_row(table, ridx, canonical_headers)
        if score > best_score:
            best_score = score
            best_idx = ridx
            best_headers = headers

    if not best_headers and best_idx < len(table.rows):
        _, best_headers = _score_header_row(table, best_idx, canonical_headers)

    return (best_idx, best_score, best_headers)


def update_table_shape(table: Table, canonical_table: CanonicalTable, errors: List[str], binding_id: str):
    if len(table.rows) < 2:
        errors.append(f"Table '{binding_id}' has fewer than 2 rows")
        return

    total_rows = len(table.rows)

    # Locate section header rows by col0 starting "Net Returns"
    section_header_indices: List[int] = []
    for r in range(total_rows):
        t0 = _norm(table.rows[r].cells[0].text)
        if _is_section_header(t0):
            section_header_indices.append(r)

    has_sections = bool(section_header_indices) and canonical_table.sections and len(canonical_table.sections) > 0

    # ── Flat fallback ────────────────────────────────────────────
    if not has_sections:
        headers = []
        for ci in range(len(table.rows[0].cells)):
            headers.append(_norm(table.rows[0].cells[ci].text))

        col_lookup: Dict[str, int] = {}
        for idx in range(1, len(headers)):
            h = _norm(headers[idx])
            if h:
                col_lookup[h] = idx

        canonical_row_lookup = {_norm(r.row_key): r for r in canonical_table.rows}

        updated = 0
        for row_idx in range(1, total_rows):
            row = table.rows[row_idx]
            row_label = _norm(row.cells[0].text)
            if not row_label:
                continue

            cr = canonical_row_lookup.get(row_label)
            if not cr:
                continue

            for col_header, col_idx in col_lookup.items():
                cell_data = cr.cells.get(col_header)
                if cell_data is None:
                    for ck, cv in cr.cells.items():
                        if _header_matches(ck, col_header):
                            cell_data = cv
                            break
                if cell_data is None:
                    continue

                cell = row.cells[col_idx]
                if cell.text_frame and cell.text_frame.paragraphs:
                    p = cell.text_frame.paragraphs[0]
                    if p.runs:
                        p.runs[0].text = cell_data.formatted
                    else:
                        p.text = cell_data.formatted
                    updated += 1

        if updated == 0:
            errors.append(f"DEBUG {binding_id}: 0 cells updated (flat). headers(sample)={headers[1:6]}")
        return

    # ── Section-aware update ─────────────────────────────────────
    section_dates: Dict[str, str] = {}
    if canonical_table.meta and canonical_table.meta.section_dates:
        section_dates = canonical_table.meta.section_dates

    positional_labels = ["Top", "Bottom", "Section3", "Section4"]

    # Resolve section labels by matching the trailing date in PPT header if possible
    resolved_sections: List[str] = []
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

        # Update section title date text
        if section_label in section_dates:
            _update_section_header_date(table.rows[sh_row_idx].cells[0], section_dates[section_label])

        if not sec_data:
            section_debug_info[section_label] = {"reason": "no canonical section data"}
            continue

        # Detect header row (includes sh_row_idx itself)
        col_header_row_idx, best_score, chosen_headers = _detect_header_row_idx(
            table, sh_row_idx, sec_data.headers, scan_depth=4
        )

        data_row_start = col_header_row_idx + 1
        next_sh = section_header_indices[sec_idx + 1] if sec_idx + 1 < len(section_header_indices) else total_rows

        # Map header text -> column index (chosen_headers align with columns 1..N)
        sec_col_lookup: Dict[str, int] = {}
        for j, h in enumerate(chosen_headers, start=1):
            hn = _norm(h)
            if hn:
                sec_col_lookup[hn] = j

        def _find_col_idx(canonical_header: str) -> Optional[int]:
            ch = _norm(canonical_header)
            if ch in sec_col_lookup:
                return sec_col_lookup[ch]
            for ppt_h, idx_col in sec_col_lookup.items():
                if _header_matches(ch, ppt_h):
                    return idx_col
            return None

        sec_row_lookup = {_norm(r.row_key): r for r in sec_data.rows}

        # Debug samples
        ppt_rows_sample: List[str] = []
        for rr in range(data_row_start, min(next_sh, data_row_start + 10)):
            rl = _norm(table.rows[rr].cells[0].text)
            if rl and not _is_section_header(rl):
                ppt_rows_sample.append(rl)

        section_debug_info[section_label] = {
            "chosen_header_row_idx": col_header_row_idx,
            "chosen_header_score": best_score,
            "chosen_headers(sample)": chosen_headers[:10],
            "ppt_rows(sample)": ppt_rows_sample[:10],
            "canonical_rows(sample)": [_norm(r.row_key) for r in sec_data.rows[:10]],
            "canonical_headers(sample)": [_norm(h) for h in sec_data.headers[:10]],
        }

        # Update cells
        for row_idx in range(data_row_start, next_sh):
            row = table.rows[row_idx]
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
                if cell.text_frame and cell.text_frame.paragraphs:
                    p = cell.text_frame.paragraphs[0]
                    if p.runs:
                        p.runs[0].text = cv.formatted
                    else:
                        p.text = cv.formatted
                    updated_cells_count += 1

    if updated_cells_count == 0:
        errors.append(
            f"DEBUG {binding_id}: 0 cells updated (sectioned). "
            f"section_headers={section_header_indices} resolved={resolved_sections} "
            f"header_choice={section_debug_info} version={VERSION}"
        )


# ─────────────────────────────────────────────────────────────
# Endpoints
# ─────────────────────────────────────────────────────────────

@app.get("/health")
async def health():
    return {"status": "ok", "version": VERSION}


@app.get("/version")
async def version():
    return {"version": VERSION}


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
                    update_text_shape(shape, kpi.formatted)
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

