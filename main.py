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
    return cell_text.strip().lower().startswith("net returns")


def _extract_trailing_paren_value(text: str) -> Optional[str]:
    # Extract the final "(...)" content at the end of the string
    m = re.search(r"\(([^)]*)\)\s*$", text.strip())
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
    Robustly replace ONLY the trailing (date) part, even if original text
    is split across runs. We read full cell text, transform it, then write
    it back into the first run to preserve formatting.
    """
    if not cell.text_frame:
        return
    full_text = cell.text_frame.text or ""
    updated_text = re.sub(r"\([^)]*\)\s*$", f"({new_date})", full_text.strip())
    if updated_text != full_text.strip():
        _update_cell_text_preserve_first_run(cell, updated_text)


def update_table_shape(table: Table, canonical_table: CanonicalTable, errors: List[str], binding_id: str):
    """
    Section-aware table updater:
    - Detects section header rows (col 0 starts with "Net Returns")
    - Uses canonical_table.sections["Top"/"Bottom"] if present
    - Column header row = row immediately after section header row
    - Data starts = row after column header row
    - Updates by matching:
        row_key == text in col 0
        col_key == header text in column header row (cols 1..end)
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
        cell0_text = table.rows[row_idx].cells[0].text.strip()
        if _is_section_header(cell0_text):
            section_header_indices.append(row_idx)

    has_sections = (
        len(section_header_indices) > 0
        and canonical_table.sections
        and len(canonical_table.sections) > 0
    )

    # ── Flat fallback (no sections) ────────────────────────────
    if not has_sections:
        headers = [c.text.strip() for c in table.rows[0].cells]
        if len(headers) < 2:
            errors.append(f"Table '{binding_id}': expected at least 2 columns")
            return

        col_lookup: Dict[str, int] = {}
        for idx in range(1, len(headers)):
            h = headers[idx].strip()
            if h:
                col_lookup[h] = idx

        canonical_row_lookup = {r.row_key.strip(): r for r in canonical_table.rows}

        for row_idx in range(1, total_rows):
            row = table.rows[row_idx]
            row_label = row.cells[0].text.strip()
            if not row_label:
                continue
            cr = canonical_row_lookup.get(row_label)
            if not cr:
                continue

            for col_header, col_idx in col_lookup.items():
                cell_data = cr.cells.get(col_header)
                if cell_data is None:
                    # try trimmed match
                    for ck, cv in cr.cells.items():
                        if ck.strip() == col_header:
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
        cell_text = table.rows[sh_row_idx].cells[0].text.strip()
        header_date = _extract_trailing_paren_value(cell_text)

        resolved = None
        if header_date and section_dates:
            for sec_label, sec_date in section_dates.items():
                if sec_date.strip() == header_date.strip():
                    resolved = sec_label
                    break

        if not resolved:
            resolved = positional_labels[idx] if idx < len(positional_labels) else f"Section{idx+1}"

        resolved_sections.append(resolved)

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

        # Read headers for this section
        sec_headers = [c.text.strip() for c in table.rows[col_header_row_idx].cells]
        sec_col_lookup: Dict[str, int] = {}
        for idx2 in range(1, len(sec_headers)):
            h = sec_headers[idx2].strip()
            if h:
                sec_col_lookup[h] = idx2

        if not sec_data:
            continue

        sec_row_lookup = {r.row_key.strip(): r for r in sec_data.rows}

        # Update rows in this section
        for row_idx in range(data_row_start, next_sh):
            row = table.rows[row_idx]
            row_label = row.cells[0].text.strip()
            if not row_label or _is_section_header(row_label):
                continue

            cr = sec_row_lookup.get(row_label)
            if not cr:
                continue

            for col_header, col_idx in sec_col_lookup.items():
                cell_data = cr.cells.get(col_header)
                if cell_data is None:
                    for ck, cv in cr.cells.items():
                        if ck.strip() == col_header:
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
