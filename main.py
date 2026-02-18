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

def _norm(s: str) -> str:
    s = (s or "").replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

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
    return bool(ch2 and ph2 and (ch2 == ph2 or ch2 in ph2 or ph2 in ch2))

def _score_header_row(table: Table, row_idx: int, canonical_headers: List[str]) -> Tuple[int, List[str]]:
    if row_idx < 0 or row_idx >= len(table.rows):
        return (0, [])
    ppt_headers = [_norm(c.text) for c in table.rows[row_idx].cells[1:]]
    score = 0
    for ch in canonical_headers:
        for ph in ppt_headers:
            if _header_matches(ch, ph):
                score += 1
                break
    return (score, ppt_headers)

def _detect_header_row_idx(table: Table, sh_row_idx: int, canonical_headers: List[str]) -> Tuple[int, int, List[str]]:
    best_idx = sh_row_idx + 1
    best_score = 0
    best_headers = []
    last_row = min(len(table.rows) - 1, sh_row_idx + 4)
    for ridx in range(sh_row_idx + 1, last_row + 1):
        score, headers = _score_header_row(table, ridx, canonical_headers)
        if score > best_score:
            best_score = score
            best_idx = ridx
            best_headers = headers
    if not best_headers and best_idx < len(table.rows):
        best_headers = [_norm(c.text) for c in table.rows[best_idx].cells[1:]]
    return best_idx, best_score, best_headers

def _is_section_header(text: str) -> bool:
    return _norm(text).lower().startswith("net returns")

def _extract_trailing_paren_value(text: str) -> Optional[str]:
    m = re.search(r"\(([^)]*)\)\s*$", _norm(text))
    return _norm(m.group(1)) if m else None

def _set_cell(cell, value: str) -> bool:
    if not cell.text_frame or not cell.text_frame.paragraphs:
        return False
    current = _norm(cell.text_frame.text)
    target = _norm(value)
    if current == target:
        return False
    p = cell.text_frame.paragraphs[0]
    if p.runs:
        p.runs[0].text = value
        for r in p.runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p.text = value
    return True

# ─────────────────────────────────────────────────────────────
# Table Update Logic
# ─────────────────────────────────────────────────────────────

def update_table_shape(table: Table, canonical_table: CanonicalTable, errors: List[str], binding_id: str):

    updated_cells = 0
    total_rows = len(table.rows)

    section_header_indices = []
    for i in range(total_rows):
        if _is_section_header(table.rows[i].cells[0].text):
            section_header_indices.append(i)

    if not section_header_indices or not canonical_table.sections:
        return

    section_dates = canonical_table.meta.section_dates if canonical_table.meta else {}
    debug_info = {}

    for idx, sh_idx in enumerate(section_header_indices):
        section_label = list(canonical_table.sections.keys())[idx]
        sec_data = canonical_table.sections.get(section_label)

        header_row_idx, score, chosen_headers = _detect_header_row_idx(
            table, sh_idx, sec_data.headers
        )

        debug_info[section_label] = {
            "row": header_row_idx,
            "score": score,
            "headers": chosen_headers[:10]
        }

        next_section = section_header_indices[idx + 1] if idx + 1 < len(section_header_indices) else total_rows
        data_start = header_row_idx + 1

        col_lookup = {}
        for col_i, h in enumerate(chosen_headers, start=1):
            col_lookup[_norm(h)] = col_i

        row_lookup = {_norm(r.row_key): r for r in sec_data.rows}

        for r_i in range(data_start, next_section):
            row_label = _norm(table.rows[r_i].cells[0].text)
            if not row_label:
                continue
            canonical_row = row_lookup.get(row_label)
            if not canonical_row:
                continue

            for ck, cv in canonical_row.cells.items():
                col_idx = None
                for ppt_h, idx2 in col_lookup.items():
                    if _header_matches(ck, ppt_h):
                        col_idx = idx2
                        break
                if col_idx is None:
                    continue

                if _set_cell(table.rows[r_i].cells[col_idx], cv.formatted):
                    updated_cells += 1

    if updated_cells == 0:
        errors.append(
            f"DEBUG {binding_id}: 0 cells updated. header_choice={debug_info}"
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
            alt = shape._element.xpath(".//*[local-name()='cNvPr']/@descr")
            if not alt:
                continue
            binding_id = alt[0]
            if binding_id in table_bindings and shape.has_table:
                canonical_table = request.canonical_data.tables.get(binding_id)
                if canonical_table:
                    update_table_shape(shape.table, canonical_table, errors, binding_id)

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
    return {"version": "header-scan-v2"}

