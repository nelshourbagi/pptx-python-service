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

# Bump this whenever you change logic so you can verify Railway updated.
SERVICE_VERSION = "header-scan-v2a"


# ─────────────────────────────────────────────────────────────
# Global error handler (so Railway returns useful JSON)
# ─────────────────────────────────────────────────────────────
@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
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
    """Normalize text for robust matching."""
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    # Normalize whitespace
    s = s.replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _header_matches(canonical_h: str, ppt_h: str) -> bool:
    """
    Partial-friendly header matching with parenthetical qualifier stripping.
    - Exact match after normalization
    - Substring contains either direction
    - Retry after stripping trailing "(...)" qualifiers
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


def get_shape_alt_text(shape) -> str:
    """Extract alt text (Description) from a shape; fallback to shape name."""
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
    """Replace the entire text content of a shape while preserving formatting."""
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
    # In your PPT table, these rows are "titles" like "Net Returns (USD) (Nov 30, 2025)"
    return _norm(cell_text).lower().startswith("net returns")


def _extract_trailing_paren_value(text: str) -> Optional[str]:
    # Extract the final "(...)" content at the end of the string
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
    Replace ONLY the trailing (date) part, even if the original text is split across runs.
    We read full cell text, transform it, then write back into the first run.
    """
    if not cell.text_frame:
        return
    full_text = _norm(cell.text_frame.text or "")
    updated_text = re.sub(r"\([^)]*\)\s*$", f"({new_date})", full_text)
    if updated_text != full_text:
        _update_cell_text_preserve_first_run(cell, updated_text)


def _score_header_row(table: Table, row_idx: Any, canonical_headers: List[str]) -> Tuple[int, List[str]]:
    """
    Score a candidate header row.
    Returns (score, headers_list) where headers_list are the raw texts from columns 1..end.
    Defensive: if row_idx isn't an int, returns (0, []).
    """
    if not isinstance(row_idx, int):
        return (0, [])
    if row_idx < 0 or row_idx >= len(table.rows):
        return (0, [])

    # candidate headers from columns 1..end (skip col0)
    ppt_headers_raw = [table.rows[row_idx].cells[i].text for i in range(1, len(table.rows[row_idx].cells))]
    score = 0

    # Compare canonical headers vs candidate headers
    for ch in canonical_headers:
        matched = False
        for ph in ppt_headers_raw:
            if _header_matches(ch, ph):
                matched = True
                break
        if matched:
            score += 1

    return (score, ppt_headers_raw)


def _detect_header_row_idx(
    table: Table,
    sh_row_idx: Any,
    canonical_headers: List[str],
    scan_depth: int = 4,
) -> Tuple[int, int, List[str]]:
    """
    Scan rows sh_row_idx+1..sh_row_idx+scan_depth and choose highest score.
    Returns (best_row_idx, best_score, chosen_headers_raw).
    Defensive: ensures all indices are ints and never raises.
    """
    if not isinstance(sh_row_idx, int):
        # Fallback: safest
        return (1, 0, [])

    best_idx: int = int(sh_row_idx) + 1
    best_score: int = 0
    best_headers_raw: List[str] = []

    last_row = min(len(table.rows) - 1, int(sh_row_idx) + int(scan_depth))
    for ridx in range(int(sh_row_idx) + 1, int(last_row) + 1):
        score, headers_raw = _score_header_row(table, ridx, canonical_headers)
        if score > best_score:
            best_score = score
            best_idx = ridx
            best_headers_raw = headers_raw

    # If best_headers_raw ended empty, at least capture the chosen row's headers
    if not best_headers_raw and 0 <= best_idx < len(table.rows):
        try:
            best_headers_raw = [
                table.rows[best_idx].cells[i].text
                for i in range(1, len(table.rows[best_idx].cells))
            ]
        except Exception:
            best_headers_raw = []

    return (best_idx, best_score, best_headers_raw)


def update_table_shape(table: Table, canonical_table: CanonicalTable, errors: List[str], binding_id: str):
    """
    Section-aware table updater:
    - Detects section header rows (col 0 starts with "Net Returns")
    - Uses canonical_table.sections["Top"/"Bottom"] if present
    - Detects the correct column header row by scanning a few rows after each section header
    - Updates by matching:
        row_key == text in col 0  (normalized)
        col_key == header text in column header row (cols 1..end) (partial-friendly)
    - Updates section header date using canonical_table.meta.section_dates if available
    - Falls back to flat matching if no sections detected / provided
    """
    if len(table.rows) < 2:
        errors.append(f"Table '{binding_id}' has fewer than 2 rows")
        return

    total_rows = len(table.rows)

    # Detect section header row indices
    section_header_indices: List[int] = []
    for row_idx in range(total_rows):
        cell0_text = _norm(table.rows[row_idx].cells[0].text)
        if _is_section_header(cell0_text):
            section_header_indices.append(row_idx)

    has_sections = (
        len(section_header_indices) > 0
        and canonical_table.sections
        and len(canonical_table.sections) > 0
    )

    # ── Flat fallback (no sections) ────────────────────────────
    if not has_sections:
        headers0 = [_norm(c.text) for c in table.rows[0].cells]
        if len(headers0) < 2:
            errors.append(f"Table '{binding_id}': expected at least 2 columns")
            return

        col_lookup: Dict[str, int] = {}
        for idx in range(1, len(headers0)):
            h = headers0[idx]
            if h:
                col_lookup[h] = idx

        canonical_row_lookup = {_norm(r.row_key): r for r in canonical_table.rows}

        updated_cells_count = 0
        for row_idx in range(1, total_rows):
            row = table.rows[row_idx]
            row_label = _norm(row.cells[0].text)
            if not row_label:
                continue
            cr = canonical_row_lookup.get(row_label)
            if not cr:
                continue

            for col_header, col_idx in col_lookup.items():
                # find matching canonical key
                cell_data = None
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
                    updated_cells_count += 1

        if updated_cells_count == 0:
            # minimal debug
            errors.append(f"DEBUG {binding_id}: 0 cells updated (flat). version={SERVICE_VERSION}")
        return

    # ── Section-aware update ───────────────────────────────────
    section_dates: Dict[str, str] = {}
    if canonical_table.meta and canonical_table.meta.section_dates:
        section_dates = canonical_table.meta.section_dates or {}

    # Determine section label per header row:
    # Prefer matching by date in the header text (deterministic).
    # Fallback to positional labeling if needed.
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
            resolved = positional_labels[idx] if idx < len(positional_labels) else f"Section{idx + 1}"

        resolved_sections.append(resolved)

    updated_cells_count = 0
    section_debug_info: Dict[str, Any] = {}

    for sec_idx, sh_row_idx in enumerate(section_header_indices):
        section_label = resolved_sections[sec_idx]
        sec_data = canonical_table.sections.get(section_label) if canonical_table.sections else None

        # Update date in section header cell (col 0)
        if section_label in section_dates:
            _update_section_header_date(table.rows[sh_row_idx].cells[0], section_dates[section_label])

        # Determine next section boundary (or end of table)
        next_sh = section_header_indices[sec_idx + 1] if sec_idx + 1 < len(section_header_indices) else total_rows

        if not sec_data:
            section_debug_info[section_label] = {
                "skipped": True,
                "reason": "no_canonical_section_data",
            }
            continue

        # Detect correct header row by scanning a few rows after section header
        canonical_headers = sec_data.headers or canonical_table.headers or []
        chosen_header_row_idx, chosen_header_score, chosen_headers_raw = _detect_header_row_idx(
            table=table,
            sh_row_idx=sh_row_idx,
            canonical_headers=canonical_headers,
            scan_depth=4,
        )

        # Data starts after the chosen header row
        data_row_start = chosen_header_row_idx + 1

        # Build normalized PPT header → column index lookup from chosen header row
        # chosen_headers_raw corresponds to columns 1..end, so start index at 1
        sec_col_lookup: Dict[str, int] = {}
        for i, h_raw in enumerate(chosen_headers_raw, start=1):
            hn = _norm(h_raw)
            if hn:
                sec_col_lookup[hn] = i

        # helper to find best matching column index for a canonical header
        def _find_col_idx(canonical_header: str) -> Optional[int]:
            chn = _norm(canonical_header)
            if not chn:
                return None
            # exact normalized match
            if chn in sec_col_lookup:
                return sec_col_lookup[chn]
            # partial-friendly match
            for ppt_h, idx_col in sec_col_lookup.items():
                if _header_matches(chn, ppt_h):
                    return idx_col
            return None

        # Build canonical row lookup (normalized row_key)
        sec_row_lookup = {_norm(r.row_key): r for r in sec_data.rows}

        # Collect some PPT row labels for debug
        ppt_row_labels_norm: List[str] = []
        for ridx in range(data_row_start, next_sh):
            if ridx < 0 or ridx >= total_rows:
                continue
            rl = _norm(table.rows[ridx].cells[0].text)
            if rl and not _is_section_header(rl):
                ppt_row_labels_norm.append(rl)

        # Update rows in this section
        for ridx in range(data_row_start, next_sh):
            if ridx < 0 or ridx >= total_rows:
                continue
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

                # Ensure column exists
                if col_idx < 0 or col_idx >= len(row.cells):
                    continue

                cell = row.cells[col_idx]
                if cell.text_frame and cell.text_frame.paragraphs:
                    p = cell.text_frame.paragraphs[0]
                    if p.runs:
                        p.runs[0].text = cv.formatted
                    else:
                        p.text = cv.formatted
                    updated_cells_count += 1

        # Store per-section debug info (only used if zero updates overall)
        chosen_headers_norm = [_norm(x) for x in chosen_headers_raw][:10]
        section_debug_info[section_label] = {
            "chosen_header_row_idx": chosen_header_row_idx,
            "chosen_header_score": chosen_header_score,
            "chosen_headers(sample)": chosen_headers_norm,
            "ppt_rows(sample)": ppt_row_labels_norm[:10],
            "canonical_rows(sample)": [_norm(r.row_key) for r in sec_data.rows][:10],
            "canonical_headers(sample)": [_norm(h) for h in (sec_data.headers or [])][:10],
        }

    # ── Debug output when no cells updated across all sections ──
    if updated_cells_count == 0:
        # Build a single debug string safely (no inline slicing/chaining that can crash)
        section_headers_copy = list(section_header_indices)
        resolved_copy = list(resolved_sections)

        # Flatten minimal samples
        errors.append(
            "DEBUG {bid}: 0 cells updated (sectioned). "
            "section_headers={sh} resolved={res} header_choice={hc} version={ver}".format(
                bid=binding_id,
                sh=section_headers_copy,
                res=resolved_copy,
                hc=section_debug_info,
                ver=SERVICE_VERSION,
            )
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
        errors=errors,
    )


@app.get("/health")
async def health():
    return {"status": "ok", "version": SERVICE_VERSION}


@app.get("/version")
async def version():
    return {"version": SERVICE_VERSION}

