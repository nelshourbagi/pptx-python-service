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

# Change this any time you redeploy, so you can confirm Railway is live-updated
SERVICE_VERSION = "header-scan-v2"


# ─────────────────────────────────────────────────────────────
# Global exception handler (so Railway won't return plain text 500)
# ─────────────────────────────────────────────────────────────
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    return JSONResponse(
        status_code=500,
        content={
            "error": str(exc),
            "type": exc.__class__.__name__,
            "traceback": traceback.format_exc(),
            "version": SERVICE_VERSION,
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
    """
    Normalize text for matching:
    - string cast
    - trim
    - collapse whitespace
    - remove common footnote/superscript artifacts
    """
    if s is None:
        return ""
    t = str(s).strip()
    t = re.sub(r"\s+", " ", t)

    # Remove Unicode superscripts ¹²³⁴⁵⁶⁷⁸⁹⁰ and common footnote markers
    t = re.sub(r"[¹²³⁴⁵⁶⁷⁸⁹⁰]+$", "", t).strip()
    t = re.sub(r"[\*\u2020\u2021]+$", "", t).strip()  # *, †, ‡ at end

    return t


def _header_matches(canonical_h: str, ppt_h: str) -> bool:
    """
    Partial-friendly header matching with parenthetical stripping.
    Examples:
      "3 year (annl)" matches "3 year"
      "Current AUM ($m)" matches "Current AUM"
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
    return bool(ch2 and ph2 and (ch2 == ph2 or ch2 in ph2 or ph2 in ch2))


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
    # Detect section header rows by checking if column-0 text starts with 'Net Returns'
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
    """
    Replace trailing (date) in the section header cell.
    """
    if not cell.text_frame:
        return
    full_text = cell.text_frame.text or ""
    updated_text = re.sub(r"\([^)]*\)\s*$", f"({new_date})", full_text.strip())
    if updated_text != full_text.strip():
        _update_cell_text_preserve_first_run(cell, updated_text)


def _score_header_row(table: Table, row_idx: int, canonical_headers: List[str]) -> Tuple[int, List[str]]:
    """
    Score a candidate header row by how many canonical headers match.
    Returns (score, candidate_headers_list).
    """
    if row_idx < 0 or row_idx >= len(table.rows):
        return 0, []
    candidate_headers = [_norm(c.text) for c in table.rows[row_idx].cells[1:]]  # skip col 0
    score = 0
    for ch in canonical_headers:
        for ph in candidate_headers:
            if _header_matches(ch, ph):
                score += 1
                break
    return score, candidate_headers


def _detect_header_row_idx(table: Table, sh_row_idx: int, canonical_headers: List[str], scan_depth: int = 4) -> Tuple[int, int, List[str]]:
    """
    Scan rows sh_row_idx+1 .. sh_row_idx+scan_depth and pick best header row.
    Returns: (best_row_idx, best_score, chosen_headers)
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

    if not best_headers and best_idx < len(table.rows):
        best_headers = [_norm(c.text) for c in table.rows[best_idx].cells[1:]]

    return best_idx, best_score, best_headers


def update_table_shape(table: Table, canonical_table: CanonicalTable, errors: List[str], binding_id: str):
    """
    Section-aware table updater:
    - Detects section header rows (col 0 starts with "Net Returns")
    - Chooses the REAL header row by scanning a few rows after the section header
    - Updates by matching row_key (col 0) and col_key (header text)
    - Updates section header dates via canonical_table.meta.section_dates
    - Falls back to flat matching if section data isn't present
    """
    if len(table.rows) < 2:
        errors.append(f"Table '{binding_id}' has fewer than 2 rows")
        return

    total_rows = len(table.rows)

    # detect section headers
    section_header_indices: List[int] = []
    for row_idx in range(total_rows):
        if _is_section_header(table.rows[row_idx].cells[0].text):
            section_header_indices.append(row_idx)

    has_sections = bool(section_header_indices) and canonical_table.sections and len(canonical_table.sections) > 0

    # ── Flat fallback (no sections) ────────────────────────────
    if not has_sections:
        headers = [_norm(c.text) for c in table.rows[0].cells]
        if len(headers) < 2:
            errors.append(f"Table '{binding_id}': expected at least 2 columns")
            return

        col_lookup: Dict[str, int] = {}
        for idx in range(1, len(headers)):
            h = headers[idx]
            if h:
                col_lookup[_norm(h)] = idx

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

            for ck, cv in cr.cells.items():
                # find col index by match
                col_idx = None
                ck_norm = _norm(ck)
                if ck_norm in col_lookup:
                    col_idx = col_lookup[ck_norm]
                else:
                    for ppt_h, idx in col_lookup.items():
                        if _header_matches(ck_norm, ppt_h):
                            col_idx = idx
                            break
                if col_idx is None:
                    continue

                cell = row.cells[col_idx]
                if cell.text_frame and cell.text_frame.paragraphs:
                    p = cell.text_frame.paragraphs[0]
                    if p.runs:
                        p.runs[0].text = cv.formatted
                    else:
                        p.text = cv.formatted
                    updated += 1

        if updated == 0:
            errors.append(f"DEBUG {binding_id}: 0 cells updated (flat). headers={headers[:10]}")
        return

    # ── Section-aware update ───────────────────────────────────
    section_dates: Dict[str, str] = {}
    if canonical_table.meta and canonical_table.meta.section_dates:
        section_dates = canonical_table.meta.section_dates

    # resolve section labels by position (Top/Bottom) unless matched by date
    positional_labels = ["Top", "Bottom", "Section3", "Section4"]
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

        # update header date text
        if section_label in section_dates:
            _update_section_header_date(table.rows[sh_row_idx].cells[0], section_dates[section_label])

        if not sec_data:
            section_debug_info[section_label] = {"skipped": "no canonical section data"}
            continue

        # find correct header row by scanning
        chosen_header_row_idx, chosen_score, chosen_headers = _detect_header_row_idx(
            table,
            sh_row_idx,
            sec_data.headers,
            scan_depth=4,
        )

        # next section header boundary
        next_sh = section_header_indices[sec_idx + 1] if sec_idx + 1 < len(section_header_indices) else total_rows
        data_row_start = chosen_header_row_idx + 1

        # build lookup from chosen headers
        sec_col_lookup: Dict[str, int] = {}
        for idx2, h in enumerate(chosen_headers, start=1):
            hn = _norm(h)
            if hn:
                sec_col_lookup[hn] = idx2

        # row lookup
        sec_row_lookup = {_norm(r.row_key): r for r in sec_data.rows}

        # update rows
        for row_idx in range(data_row_start, next_sh):
            row = table.rows[row_idx]
            row_label = _norm(row.cells[0].text)
            if not row_label or _is_section_header(row_label):
                continue

            cr = sec_row_lookup.get(row_label)
            if not cr:
                continue

            for ck, cv in cr.cells.items():
                ck_norm = _norm(ck)
                col_idx = None
                if ck_norm in sec_col_lookup:
                    col_idx = sec_col_lookup[ck_norm]
                else:
                    for ppt_h, idx3 in sec_col_lookup.items():
                        if _header_matches(ck_norm, ppt_h):
                            col_idx = idx3
                            break
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

        section_debug_info[section_label] = {
            "chosen_header_row_idx": chosen_header_row_idx,
            "chosen_header_score": chosen_score,
            "chosen_headers_sample": chosen_headers[:10],
        }

    if updated_cells_count == 0:
        # include helpful debug if nothing changed
        ppt_rows_sample: List[str] = []
        for r in range(0, min(total_rows, 20)):
            rl = _norm(table.rows[r].cells[0].text)
            if rl and not _is_section_header(rl):
                ppt_rows_sample.append(rl)
        errors.append(
            f"DEBUG {binding_id}: 0 cells updated (sectioned). "
            f"section_headers={section_header_indices}, resolved={resolved_sections}. "
            f"header_choice={section_debug_info}. "
            f"PPT_rows(sample)={ppt_rows_sample[:10]}. "
            f"Canonical_rows(sample)={[_norm(x.row_key) for x in (list(canonical_table.sections.values())[0].rows[:10] if canonical_table.sections else [])]}"
        )


# ─────────────────────────────────────────────────────────────
# Endpoints
# ─────────────────────────────────────────────────────────────

@app.get("/health")
async def health():
    return {"status": "ok", "version": SERVICE_VERSION}


@app.get("/version")
async def version():
    return {"version": SERVICE_VERSION}


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
        errors=errors,
    )
