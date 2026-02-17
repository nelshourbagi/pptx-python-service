import base64
import io
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import Dict, List, Any
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


class CanonicalTable(BaseModel):
    headers: List[str]
    rows: List[CanonicalRow]


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
        desc = shape._element.xpath(
            ".//*[local-name()='cNvPr']/@descr"
        )
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


def update_table_shape(table: Table, canonical_table: CanonicalTable, errors: List[str], binding_id: str):
    if len(table.rows) < 2:
        errors.append(f"Table '{binding_id}' has fewer than 2 rows")
        return

    headers = [cell.text.strip() for cell in table.rows[0].cells]

    if not headers or headers[0].lower() != "year":
        errors.append(f"Table '{binding_id}' first column must be 'Year'")
        return

    col_lookup = {}
    for i, h in enumerate(headers):
        if i == 0:
            continue
        col_lookup[h.strip()] = i

    for row in table.rows[1:]:
        year_val = row.cells[0].text.strip()

        match = None
        for r in canonical_table.rows:
            if r.row_key.strip() == year_val:
                match = r
                break

        if not match:
            continue

        for header, col_index in col_lookup.items():
            if header in match.cells:
                cell_data = match.cells[header]
                cell = row.cells[col_index]

                if cell.text_frame and cell.text_frame.paragraphs:
                    para = cell.text_frame.paragraphs[0]
                    if para.runs:
                        para.runs[0].text = cell_data.formatted
                    else:
                        para.text = cell_data.formatted


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
