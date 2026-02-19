import base64
import io
import re
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import Dict, List, Any, Optional, Tuple
from pptx import Presentation
from pptx.table import Table

app = FastAPI()

# ─────────────────────────────────────────────────────────────
# Version (use this to confirm Railway is running latest code)
# ─────────────────────────────────────────────────────────────

@app.get("/version")
async def version():
    return {"version": "section-row-is-header-v1"}


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
    type: str  # "text" | "table"


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

_SUPERSCRIPT_MAP = str.maketrans({
    "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4",
    "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9",
})

def _norm(s: Any) -> str:
    """Generic normalization for headers and values."""
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\u00a0", " ")  # non-breaking space
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _norm_key(s: Any) -> str:
    """Normalization for row keys (handles superscripts / whitespace / case)."""
    s = _norm(s)
    s = s.translate(_SUPERSCRIPT_MAP)
    s = s.lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _header_matches(canonical_h: str, ppt_h: str) -> bool:
    """
    Partial-friendly header matching with parenthetical qualifier stripping.
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

    # Strip trailing parentheticals and retry
    ch2 = re.sub(r"\s*\([^)]*\)\s*$", "", ch).strip()
    ph2 = re.sub(r"\s*\([^)]*\)\s*$", "", ph).strip()
    if not ch2 or not ph2:
        return False
    return (ch2 == ph2) or (ch2 in ph2) or (ph2 in ch2)

def get_shape_alt_text(shape) -> str:
    """Extract alt text (Description) from a shape."""
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
    """Replace text while preserving formatting (first run of first paragraph)."""
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    if not tf.paragraphs:
        return

    p0 = tf.paragraphs[0]
    if p0.runs:
        p0.runs[0].text = formatted_value
        for r in p0.runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p0.text = formatted_value

    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)

def _is_section_header(cell_text: str) -> bool:
    """
    Your table uses a row like:
      "Net Returns (USD) (Nov 30, 2025)"
    as the start of a section.
    """
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
    Replace ONLY the trailing (date) portion in the section header cell text.
    Works even if text is split across runs.
    """
    if not cell.text_frame:
        return
    full_text = cell.text_frame.text or ""
    updated_text = re.sub(r"\([^)]*\)\s*$", f"({new_date})", _norm(full_text))
    if updated_text and updated_text != _norm(full_text):
        _update_cell_text_preserve_first_run(cell, updated_text)


# ─────────────────────────────────────────────────────────────
# Table updater
# ─────────────────────────────────────────────────────────────

def update_table_shape(table: Table, canonical_table: CanonicalTable, errors: List[str], binding_id: str):
    """
    Section-aware table updater for your Slide 2 layout.

    Key rule for this template:
    - The section header row (where col0 starts with "Net Returns") ALSO contains
      the true column headers in columns 1..end (e.g. "2025 YTD", "3 year (annl)", ...).
      So for each section:
        col_header_row_idx = sh_row_idx
        data starts at sh_row_idx + 1

    Falls back to generic flat update if no sections are detected/provided.
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

    # ── Flat fallback (generic table: row0 headers, col0 row label) ──
    if not has_sections:
        header_row_idx = 0
        ppt_headers = [_norm(c.text) for c in table.rows[header_row_idx].cells[1:]]  # skip col0
        ppt_header_norms = [_norm(h) for h in ppt_headers]

        # Build lookup (normalized header -> column index)
        col_lookup: Dict[str, int] = {}
        for idx, h in enumerate(ppt_header_norms, start=1):
            if h:
                col_lookup[h] = idx

        canonical_rows = { _norm_key(r.row_key): r for r in canonical_table.rows }

        updated = 0
        for row_idx in range(1, total_rows):
            row = table.rows[row_idx]
            row_label = _norm_key(row.cells[0].text)
            if not row_label:
                continue
            cr = canonical_rows.get(row_label)
            if not cr:
                continue

            for ck, cv in cr.cells.items():
                # Find matching column index using partial-friendly match
                ck_norm = _norm(ck)
                col_idx = None
                if ck_norm in col_lookup:
                    col_idx = col_lookup[ck_norm]
                else:
                    for ph, i in col_lookup.items():
                        if _header_matches(ck_norm, ph):
                            col_idx = i
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
            errors.append(
                f"DEBUG {binding_id}: 0 cells updated (flat). "
                f"PPT_headers(sample)={ppt_headers[:10]} | "
                f"PPT_rows(sample)={[ _norm(table.rows[i].cells[0].text) for i in range(1, min(total_rows, 12)) ]} | "
                f"Canonical_rows(sample)={[r.row_key for r in canonical_table.rows[:10]]} | "
                f"Canonical_headers(sample)={canonical_table.headers[:10]}"
            )
        return

    # ── Section-aware update ──
    section_dates: Dict[str, str] = {}
    if canonical_table.meta and canonical_table.meta.section_dates:
        section_dates = canonical_table.meta.section_dates

    # Resolve section labels deterministically:
    # - If header has a trailing (date) matching section_dates -> use that key
    # - Else positional (Top, Bottom, Section3...)
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
    section_debug: List[Dict[str, Any]] = []

    for sec_idx, sh_row_idx in enumerate(section_header_indices):
        section_label = resolved_sections[sec_idx]
        sec_data = canonical_table.sections.get(section_label) if canonical_table.sections else None

        # Update the as-of date in the section header cell col0 (optional)
        if section_label in section_dates:
            _update_section_header_date(table.rows[sh_row_idx].cells[0], section_dates[section_label])

        # For THIS template: section header row IS the header row
        col_header_row_idx = sh_row_idx
        data_row_start = col_header_row_idx + 1

        # Determine where section ends (next section header or end)
        next_sh = section_header_indices[sec_idx + 1] if sec_idx + 1 < len(section_header_indices) else total_rows

        # Read headers for this section from the section header row cells[1:]
        chosen_headers = [_norm(c.text) for c in table.rows[col_header_row_idx].cells[1:]]  # skip col0 title
        chosen_headers_norm = [_norm(h) for h in chosen_headers]

        # Build lookup (normalized header -> column index)
        sec_col_lookup: Dict[str, int] = {}
        for idx2, h in enumerate(chosen_headers_norm, start=1):
            if h:
                sec_col_lookup[h] = idx2

        if not sec_data:
            section_debug.append({
                "section": section_label,
                "col_header_row_idx": col_header_row_idx,
                "note": "no canonical section data"
            })
            continue

        # Canonical row lookup for this section (normalized)
        sec_row_lookup = { _norm_key(r.row_key): r for r in sec_data.rows }

        # Helper: find column index for a canonical header using partial match
        def find_col_idx(canonical_header: str) -> Optional[int]:
            ch = _norm(canonical_header)
            if ch in sec_col_lookup:
                return sec_col_lookup[ch]
            for ph, i in sec_col_lookup.items():
                if _header_matches(ch, ph):
                    return i
            return None

        # Update rows in this section
        for row_idx in range(data_row_start, next_sh):
            row = table.rows[row_idx]
            row_label_raw = row.cells[0].text
            row_label = _norm_key(row_label_raw)
            if not row_label:
                continue
            if _is_section_header(row_label_raw):
                continue

            cr = sec_row_lookup.get(row_label)
            if not cr:
                continue

            for ck, cv in cr.cells.items():
                col_idx = find_col_idx(ck)
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

        section_debug.append({
            "section": section_label,
            "col_header_row_idx": col_header_row_idx,
            "headers_sample": chosen_headers[:10],
            "rows_range": [data_row_start, next_sh - 1],
        })

    if updated_cells_count == 0:
        # Samples for debugging
        sample_ppt_headers = []
        if section_header_indices:
            sample_ppt_headers = [_norm(c.text) for c in table.rows[section_header_indices[0]].cells[1:]][:10]

        sample_ppt_rows = []
        # Show first ~10 row labels after first section header
        if section_header_indices:
            start = section_header_indices[0] + 1
            for i in range(start, min(total_rows, start + 12)):
                sample_ppt_rows.append(_norm(table.rows[i].cells[0].text))

        sample_canonical_rows = []
        sample_canonical_headers = []
        if canonical_table.sections:
            for _, sd in canonical_table.sections.items():
                sample_canonical_rows = [r.row_key for r in sd.rows[:10]]
                sample_canonical_headers = [h for h in sd.headers[:10]]
                break

        errors.append(
            f"DEBUG {binding_id}: 0 cells updated (sectioned). "
            f"section_headers={section_header_indices} | "
            f"resolved={resolved_sections} | "
            f"header_row_is_section_row=true | "
            f"section_debug={section_debug} | "
            f"PPT_headers(sample)={sample_ppt_headers} | "
            f"PPT_rows(sample)={sample_ppt_rows} | "
            f"Canonical_rows(sample)={sample_canonical_rows} | "
            f"Canonical_headers(sample)={sample_canonical_headers}"
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
        errors=errors,
    )

@app.get("/health")
async def health():
    return {"status": "ok"}
