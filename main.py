import base64
import io
import re
import traceback
from typing import Dict, List, Any, Optional

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from pptx import Presentation
from pptx.table import Table

app = FastAPI()

SERVICE_VERSION = "section-header-direct-v1"


# ─────────────────────────────────────────────────────────────
# Global exception handler
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

    ch2 = re.sub(r"\s*\([^)]*\)\s*$", "", ch).strip()
    ph2 = re.sub(r"\s*\([^)]*\)\s*$", "", ph).strip()

    return ch2 == ph2


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


# ─────────────────────────────────────────────────────────────
# Table Updater (FINAL FIX)
# ─────────────────────────────────────────────────────────────

def update_table_shape(table: Table, canonical_table: CanonicalTable, errors: List[str], binding_id: str):

    total_rows = len(table.rows)
    updated = 0

    # Detect section header rows
    section_rows = []
    for i in range(total_rows):
        if _norm(table.rows[i].cells[0].text).lower().startswith("net returns"):
            section_rows.append(i)

    if not section_rows or not canonical_table.sections:
        errors.append(f"DEBUG {binding_id}: no sections detected")
        return

    section_labels = list(canonical_table.sections.keys())

    for idx, sh_row_idx in enumerate(section_rows):

        section_label = section_labels[idx] if idx < len(section_labels) else None
        if not section_label:
            continue

        sec_data = canonical_table.sections.get(section_label)
        if not sec_data:
            continue

        # Header row is the section row itself
        header_row_idx = sh_row_idx

        # Extract header texts from same row (skip col0)
        header_cells = table.rows[header_row_idx].cells
        ppt_headers = [_norm(c.text) for c in header_cells[1:]]

        col_lookup: Dict[str, int] = {}
        for col_i, h in enumerate(ppt_headers, start=1):
            col_lookup[h] = col_i

        next_section = section_rows[idx + 1] if idx + 1 < len(section_rows) else total_rows
        data_start = header_row_idx + 1

        canonical_rows = {_norm(r.row_key): r for r in sec_data.rows}

        for r_i in range(data_start, next_section):
            row = table.rows[r_i]
            row_label = _norm(row.cells[0].text)
            if not row_label:
                continue

            cr = canonical_rows.get(row_label)
            if not cr:
                continue

            for ck, cv in cr.cells.items():

                col_idx = None
                for ppt_h, i_col in col_lookup.items():
                    if _header_matches(ck, ppt_h):
                        col_idx = i_col
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
            f"DEBUG {binding_id}: 0 cells updated (direct-header mode) version={SERVICE_VERSION}"
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
            if alt in table_bindings and shape.has_table:
                canonical_table = request.canonical_data.tables.get(alt)
                if canonical_table:
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

