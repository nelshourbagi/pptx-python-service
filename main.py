import base64
import io
import re
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import Dict, List, Any, Optional
from pptx import Presentation
from pptx.table import Table

app = FastAPI()


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

def _norm(s: str) -> str:
    """
    Normalize strings for matching:
    - convert NBSP to space
    - collapse all whitespace (including newlines) to single space
    - strip
    """
    s = (s or "").replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


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
    # Extract the final "(...)" content at the end of the string
    m = re.search(r"\(([^)]*)\)\s*$", _norm(text))
    return _norm(m.group(1)) if m else None


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
    Robustly replace ONLY the trailing (date) part, even if original text
    is split across runs. We read full cell text, transform it, then write
    it back into the first run to preserve formatting.
    """
    if not cell.text_frame:
        return
    full_text = cell.text_frame.text or ""
    full_text_norm = full_text.strip()
    updated_text = re.sub(r"\([^)]*\)\s*$", f"({new_date})", full_text_norm)
    if updated_text != full_text_norm:
        _update_cell_text_preserve_first_run(cell, updated_text)


def _set_cell_text_preserve_first_run(cell, new_text: str) -> bool:
    """
    Set cell text (preserving formatting) and return True if it changed.
    """
    if not cell.text_frame or not cell.text_frame.paragraphs:
        return False
    tf = cell.text_frame
    p0 = tf.paragraphs[0]

    # current visible text (normalized)
    current = _norm(tf.text or "")
    target = _norm(new_text)

    if current == target:
        return False

    if p0.runs:
        p0.runs[0].text = new_text
        for r in p0.runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p0.text = new_text

    # remove extra paragraphs
    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)

    return True


def update_table_shape(table: Table, canonical_table: CanonicalTable, errors: List[str], binding_id: str):
    """
    Section-aware table updater:
    - Detects section header rows (col 0 starts with "Net Returns")
    - Uses canonical_table.sections["Top"/"Bottom"] if present
    - Column header row = row immediately after section header row
    - Data starts = row after column header row
    - Updates by matching (normalized):
        row_key == text in col 0
        col_key == header text in column header row (cols 1..end)
    - Updates section header date using canonical_table.meta.section_dates if available
    - Falls back to flat matching if no sections detected / provided
    - Emits DEBUG message if 0 cells updated
    """

    if len(table.rows) < 2:
        errors.append(f"Table '{binding_id}' has fewer than 2 rows")
        return

    total_rows = len(table.rows)
    updated_cells_count = 0

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

        for row_idx in range(1, total_rows):
            row = table.rows[row_idx]
            row_label = _norm(row.cells[0].text)
            if not row_label:
                continue

            cr = canonical_row_lookup.get(row_label)
            if not cr:
                continue

            for col_header_norm, col_idx in col_lookup.items():
                cell_data = None

                # Try direct then normalized key match
                if col_header_norm in cr.cells:
                    cell_data = cr.cells[col_header_norm]
                else:
                    for ck, cv in cr.cells.items():
                        if _norm(ck) == col_header_norm:
                            cell_data = cv
                            break

                if cell_data is None:
                    continue

                cell = row.cells[col_idx]
                if _set_cell_text_preserve_first_run(cell, cell_data.formatted):
                    updated_cells_count += 1

        if updated_cells_count == 0:
            ppt_headers_sample = headers[1:11]
            ppt_rows_sample = []
            for rix in range(1, min(total_rows, 11)):
                ppt_rows_sample.append(_norm(table.rows[rix].cells[0].text))
            canonical_rows_sample = [_norm(r.row_key) for r in canonical_table.rows[:10]]
            canonical_headers_sample = [_norm(h) for h in canonical_table.headers[:10]]
            errors.append(
                f"DEBUG {binding_id}: 0 cells updated (flat). "
                f"PPT_headers(sample)={ppt_headers_sample}. PPT_rows(sample)={ppt_rows_sample}. "
                f"Canonical_rows(sample)={canonical_rows_sample}. Canonical_headers(sample)={canonical_headers_sample}."
            )
        return

    # ── Section-aware update ───────────────────────────────────
    section_dates: Dict[str, str] = {}
    if canonical_table.meta and canonical_table.meta.section_dates:
        section_dates = canonical_table.meta.section_dates

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
            resolved = positional_labels[idx] if idx < len(positional_labels) else f"Section{idx+1}"

        resolved_sections.append(resolved)

    # Track samples for debug
    ppt_headers_sample: List[str] = []
    ppt_rows_sample: List[str] = []
    canonical_rows_sample: List[str] = []
    canonical_headers_sample: List[str] = []

    for sec_idx, sh_row_idx in enumerate(section_header_indices):
        section_label = resolved_sections[sec_idx]
        sec_data = canonical_table.sections.get(section_label) if canonical_table.sections else None

        # Update date in section header cell (col 0)
        if section_label in section_dates:
            _update_section_header_date(table.rows[sh_row_idx].cells[0], section_dates[section_label])

        # Column header row is the row immediately after section header
        col_header_row_idx = sh_row_idx + 1
        if col_header_row_idx >= total_rows:
            errors.append(
                f"Table '{binding_id}': section '{section_label}' header at row {sh_row_idx} has no column header row after it"
            )
            continue

        # Next section header (or end of table)
        next_sh = section_header_indices[sec_idx + 1] if sec_idx + 1 < len(section_header_indices) else total_rows
        data_row_start = col_header_row_idx + 1

        # Read headers for this section (normalized)
        sec_headers = [_norm(c.text) for c in table.rows[col_header_row_idx].cells]
        sec_col_lookup: Dict[str, int] = {}
        for idx2 in range(1, len(sec_headers)):
            h = _norm(sec_headers[idx2])
            if h:
                sec_col_lookup[h] = idx2

        if not sec_data:
            continue

        sec_row_lookup = {_norm(r.row_key): r for r in sec_data.rows}

        # Capture samples from the first processed section for debug
        if not ppt_headers_sample:
            ppt_headers_sample = sec_headers[1:11]
            canonical_headers_sample = [_norm(h) for h in sec_data.headers[:10]]
            canonical_rows_sample = [_norm(r.row_key) for r in sec_data.rows[:10]]

        # Update rows in this section
        for row_idx in range(data_row_start, next_sh):
            row = table.rows[row_idx]
            row_label = _norm(row.cells[0].text)
            if not row_label or _is_section_header(row_label):
                continue

            if len(ppt_rows_sample) < 10:
                ppt_rows_sample.append(row_label)

            cr = sec_row_lookup.get(row_label)
            if not cr:
                continue

            for col_header_norm, col_idx in sec_col_lookup.items():
                cell_data = None

                # Try direct then normalized key match
                if col_header_norm in cr.cells:
                    cell_data = cr.cells[col_header_norm]
                else:
                    for ck, cv in cr.cells.items():
                        if _norm(ck) == col_header_norm:
                            cell_data = cv
                            break

                if cell_data is None:
                    continue

                cell = row.cells[col_idx]
                if _set_cell_text_preserve_first_run(cell, cell_data.formatted):
                    updated_cells_count += 1

    if updated_cells_count == 0:
        errors.append(
            f"DEBUG {binding_id}: 0 cells updated (sectioned). "
            f"section_headers={section_header_indices}, resolved={resolved_sections}. "
            f"PPT_headers(sample)={ppt_headers_sample}. PPT_rows(sample)={ppt_rows_sample}. "
            f"Canonical_rows(sample)={canonical_rows_sample}. Canonical_headers(sample)={canonical_headers_sample}."
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


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.get("/version")
async def version():
    return {"version": "norm-debug-v1"}

