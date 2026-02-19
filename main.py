import base64
import io
import re
import traceback
from typing import Dict, List, Any, Optional

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from pptx import Presentation
from pptx.table import Table

app = FastAPI()

SERVICE_VERSION = "section-header-direct-v2"


# ─────────────────────────────────────────────────────────────
# Global exception handler (so Supabase sees real Python error)
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
# Models
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
    s = str(s)
    s = s.replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _header_matches(canonical_h: str, ppt_h: str) -> bool:
    ch = _norm(canonical_h).lower()
    ph = _norm(ppt_h).lower()

    if not ch or not ph:
        return False

    if ch == ph:
        return True

    if ch in ph or ph in ch:
        return True

    # remove parenthetical qualifiers and retry
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


def _is_section_header(text: str) -> bool:
    return _norm(text).lower().startswith("net returns")


# ─────────────────────────────────────────────────────────────
# Table updater (Section header row contains the column headers)
# IMPORTANT: do NOT slice pptx collections (row.cells[1:] breaks)
# ─────────────────────────────────────────────────────────────

def update_table_shape(table: Table, canonical_table: CanonicalTable, errors: List[str], binding_id: str):
    total_rows = len(table.rows)
    updated = 0

    # Find section header rows in the PPT (col 0 starts with "Net Returns")
    section_header_indices: List[int] = []
    for i in range(total_rows):
        try:
            cell0 = table.rows[i].cells[0].text
        except Exception:
            cell0 = ""
        if _is_section_header(cell0):
            section_header_indices.append(i)

    if not section_header_indices or not canonical_table.sections:
        errors.append(
            f"DEBUG {binding_id}: no sections detected or canonical_table.sections missing. "
            f"ppt_section_headers={section_header_indices} version={SERVICE_VERSION}"
        )
        return

    # Use canonical section order as provided (Top/Bottom)
    canonical_section_labels = list(canonical_table.sections.keys())

    header_choice_debug: Dict[str, Any] = {}

    for sec_idx, sh_row_idx in enumerate(section_header_indices):
        section_label = canonical_section_labels[sec_idx] if sec_idx < len(canonical_section_labels) else f"Section{sec_idx+1}"
        sec_data = canonical_table.sections.get(section_label)
        if not sec_data:
            continue

        # In this PPT layout: section header row IS the header row
        header_row_idx = sh_row_idx

        # Build col_lookup from that same row, columns 1..end
        header_row = table.rows[header_row_idx]
        ncols = len(header_row.cells)

        ppt_headers: List[str] = []
        col_lookup: Dict[str, int] = {}

        for j in range(1, ncols):
            h = _norm(header_row.cells[j].text)
            ppt_headers.append(h)
            if h:
                col_lookup[h] = j

        # Data runs until next section header or end
        next_sh = section_header_indices[sec_idx + 1] if sec_idx + 1 < len(section_header_indices) else total_rows
        data_start = header_row_idx + 1

        # Canonical rows lookup
        canonical_rows = {_norm(r.row_key): r for r in sec_data.rows}

        # Collect some ppt row labels for debug
        ppt_rows_sample: List[str] = []
        for r_i in range(data_start, min(next_sh, data_start + 8)):
            rl = _norm(table.rows[r_i].cells[0].text)
            if rl and not _is_section_header(rl):
                ppt_rows_sample.append(rl)

        # Update cells
        for r_i in range(data_start, next_sh):
            row = table.rows[r_i]
            row_label = _norm(row.cells[0].text)
            if not row_label or _is_section_header(row_label):
                continue

            cr = canonical_rows.get(row_label)
            if not cr:
                continue

            for ck, cv in cr.cells.items():
                col_idx = None

                # match canonical header to PPT header using tolerant matching
                for ppt_h, idx_col in col_lookup.items():
                    if _header_matches(ck, ppt_h):
                        col_idx = idx_col
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

        header_choice_debug[section_label] = {
            "header_row_idx": header_row_idx,
            "ppt_headers(sample)": ppt_headers[:10],
            "ppt_rows(sample)": ppt_rows_sample[:6],
            "canonical_headers(sample)": [_norm(h) for h in sec_data.headers][:10],
            "canonical_rows(sample)": list(canonical_rows.keys())[:6],
        }

    if updated == 0:
        errors.append(
            f"DEBUG {binding_id}: 0 cells updated (direct-header mode). "
            f"sections={section_header_indices} header_debug={header_choice_debug} "
            f"version={SERVICE_VERSION}"
        )


# ─────────────────────────────────────────────────────────────
# Endpoints
# ─────────────────────────────────────────────────────────────

@app.post("/", response_model=ProcessResponse)
async def process_pptx(request: ProcessRequest):
    errors: List[str] = []

    template_bytes = base64.b64decode(request.template_base64)
    prs = Presentation(io.BytesIO(template_bytes))

    table_bindings = {b.id for b in request.mapping_config.bindings if b.type == "table"}

    for slide in prs.slides:
        for shape in slide.shapes:
            alt = get_shape_alt_text(shape).strip()
            if not alt:
                continue

            if alt in table_bindings:
                if not shape.has_table:
                    errors.append(f"Binding '{alt}' is not a table")
                    continue

                canonical_table = request.canonical_data.tables.get(alt)
                if not canonical_table:
                    errors.append(f"Table binding '{alt}' not found in canonical data")
                    continue

                update_table_shape(shape.table, canonical_table, errors, alt)

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
