"""Editable Word table continuations; no fake automatic page break discovery.

- Native mode retains ONE table and repeats the first header row when it spans pages.
- Explicit mode emits multiple physical tables, with repeated headers and a
  '续表' heading linked by native REF field to the first table's SEQ caption.
- We do not claim Word inserts '续表' at arbitrary natural page boundaries.
"""
from __future__ import annotations
from typing import Any, Callable
from docx.oxml import OxmlElement
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from cross_references import q, NS, field_runs, Target


def table_segments(rows: list[list[str]], settings: dict | None = None) -> list[list[list[str]]]:
    cfg=settings or {'mode':'repeat-header'}
    if cfg.get('mode','repeat-header')=='repeat-header':
        return [rows]
    size=cfg['rows_per_part']
    return [rows[i:i+size] for i in range(0,len(rows),size)]


def build_editable_table(doc: Any, headers: list[str], rows: list[list[str]], *,
                         variant: str, border_fn: Callable[..., Any],
                         table_id: str, part_number: int) -> Any:
    table=doc.add_table(rows=1,cols=len(headers))
    table.autofit=False
    if variant=='three-line':table.style='学术三线表'
    elif variant=='grid':table.style='Table Grid'
    width=doc.sections[0].page_width-doc.sections[0].left_margin-doc.sections[0].right_margin
    cell_width=round(width/len(headers))
    for col in table.columns:col.width=cell_width
    for i,data in enumerate([headers]+rows):
        row=table.rows[0] if i==0 else table.add_row()
        props=row._tr.get_or_add_trPr()
        if props.find(q('cantSplit')) is None:props.append(OxmlElement('w:cantSplit'))
        if i==0 and props.find(q('tblHeader')) is None:props.append(OxmlElement('w:tblHeader'))
        for cell,value in zip(row.cells,data):
            cell.width=cell_width;cell.text=value
            for p in cell.paragraphs:
                p.style='表格正文'
                if i==0:
                    for r in p.runs:r.bold=True
    border_fn(table._tbl,variant)
    # A machine-readable marker does not affect visible appearance; used to
    # check that continuations are not miscounted in the list of tables.
    tblpr=table._tbl.tblPr
    cap=tblpr.find(q('tblCaption'))
    if cap is None:cap=OxmlElement('w:tblCaption');tblpr.append(cap)
    cap.set(q('val'),f'CP:{table_id}:part:{part_number}')
    return table


def continuation_caption(paragraph: Any,target: Target,title: str) -> None:
    paragraph.add_run('续表')
    paragraph._p.extend(field_runs('REF '+target.name+r' \h',target.cached))
    paragraph.add_run('  '+title)


def add_table_note(doc: Any, content: str) -> Any:
    p=doc.add_paragraph(content,style='表格正文')
    # Table notes must not inherit ordinary main-text first-line indent blindly.
    p.paragraph_format.first_line_indent=Cm(0.74)
    p.paragraph_format.line_spacing=1.0
    p.paragraph_format.keep_together=True
    return p


def audit_long_tables(document: Any) -> dict[str,Any]:
    from collections import defaultdict
    tables=list(document.iter(q('tbl')))
    errors=[];warnings=[];grouped=defaultdict(list)
    for index,table in enumerate(tables,1):
        rows=table.findall(q('tr'))
        first=rows[0] if rows else None
        has_header=bool(first is not None and first.find('w:trPr/w:tblHeader',NS) is not None)
        if not has_header:warnings.append({'code':'missing-repeat-header','table_index':index})
        for rid,row in enumerate(rows,1):
            if row.find('w:trPr/w:cantSplit',NS) is None:
                warnings.append({'code':'splittable-row','table_index':index,'row_index':rid})
        meta=table.find('w:tblPr/w:tblCaption',NS)
        val=meta.get(q('val'),'') if meta is not None else ''
        if val.startswith('CP:'):
            import re
            mat=re.fullmatch(r'CP:([A-Za-z][A-Za-z0-9_.-]{0,63}):part:([1-9][0-9]*)',val)
            if not mat:errors.append({'code':'invalid-continuation-tag','table_index':index})
            else:grouped[mat.group(1)].append((int(mat.group(2)),index))
    for key,parts in grouped.items():
        ranks=[p[0] for p in parts]
        if ranks!=list(range(1,len(ranks)+1)):
            errors.append({'code':'nonsequential-table-parts','table_id':key,'parts':ranks})
        if len(parts)>1:
            for part,index in parts[1:]:
                # A continued caption should directly precede each physical table.
                node=tables[index-1].getprevious()
                expected=False
                if node is not None and node.tag==q('p'):
                    s=''.join(t.text or '' for t in node.iter(q('t')))
                    expected=s.startswith('续表') and any((i.text or '').strip().startswith('REF ') for i in node.iter(q('instrText')))
                if not expected:
                    errors.append({'code':'missing-linked-continuation-caption','table_id':key,'part':part})
    return {'tables':len(tables),'groups':dict(grouped),'errors':errors,
            'warnings':warnings,'review_required':bool(errors or warnings),
            'note':'Physical pagination requires rendering; repeat-header mode cannot auto-generate page-specific 续表 captions.'}
