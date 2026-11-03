"""Neutral academic code/table OOXML styles for editable Chinese Word documents.

Generic defaults, not a claim of compliance with any university template.
The caller must select a target: this module never scans prose to guess code.
"""
from __future__ import annotations
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
def q(local: str) -> str:
    return '{' + W + '}' + local

BLACK = '000000'
WHITE = 'FFFFFF'
CODE_GRAY = 'F5F5F5'
TABLE_STYLE_ID = 'AcademicThreeLine'
TABLE_STYLE_NAME = '学术三线表'
TABLE_VARIANTS = ('three-line', 'grid', 'none')
CODE_STYLE_NAMES = ('代码段', 'Code Block', 'CodeBlock')

PPR_ORDER = ['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr',
             'suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku',
             'wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi',
             'adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents',
             'suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap',
             'outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange']
TRPR_ORDER = ['cnfStyle','divId','gridBefore','gridAfter','wBefore','wAfter','cantSplit',
              'trHeight','tblHeader','tblCellSpacing','jc','hidden','ins','del','trPrChange']
TCPR_ORDER = ['cnfStyle','tcW','gridSpan','hMerge','vMerge','tcBorders','shd',
              'noWrap','tcMar','textDirection','tcFitText','vAlign','hideMark','headers','cellIns','cellDel','cellMerge','tcPrChange']
TBLPR_ORDER = ['tblStyle','tblpPr','tblOverlap','bidiVisual','tblStyleRowBandSize',
               'tblStyleColBandSize','tblW','jc','tblCellSpacing','tblInd','tblBorders',
               'shd','tblLayout','tblCellMar','tblLook','tblCaption','tblDescription','tblPrChange']

def ordered(parent, tag: str, order: list[str]):
    found = parent.find(q(tag))
    if found is not None:
        return found
    found = etree.Element(q(tag))
    rank = order.index(tag)
    idx = next((i for i, child in enumerate(parent)
                if etree.QName(child).localname in order
                and order.index(etree.QName(child).localname) > rank), len(parent))
    parent.insert(idx, found)
    return found

def sort_ooxml_properties(parent, order):
    original=list(parent)
    positions={k:i for i,k in enumerate(order)}
    # Stable, defined elements sorted according to the OOXML schema ordering.
    parent[:]=sorted(original,key=lambda e:positions.get(etree.QName(e).localname,len(order)))


def drop(parent, tag: str):
    for item in list(parent.findall(q(tag))):
        parent.remove(item)

def border(parent, edge, *, sz=8, show=True):
    elem = etree.SubElement(parent, q(edge))
    elem.set(q('val'), 'single' if show else 'nil')
    if show:
        elem.set(q('sz'), str(sz))  # OOXML border size in eighths of a point.
        elem.set(q('color'), BLACK)
    return elem

def set_code_style(style, *, fill=CODE_GRAY):
    ppr = ordered(style, 'pPr', ['name','aliases','basedOn','next','link','autoRedefine',
                                 'hidden','uiPriority','semiHidden','unhideWhenUsed',
                                 'qFormat','locked','personal','personalCompose','personalReply',
                                 'rsid','pPr','rPr','tblPr','trPr','tcPr'])
    drop(ppr, 'pBdr')
    shd = ordered(ppr, 'shd', PPR_ORDER)
    shd.attrib.clear()
    shd.set(q('val'), 'clear')
    shd.set(q('color'), 'auto')
    shd.set(q('fill'), fill)
    ind = ordered(ppr, 'ind', PPR_ORDER)
    ind.attrib.pop(q('hanging'), None)
    ind.attrib.pop(q('hangingChars'), None)
    ind.set(q('firstLine'), '0');ind.set(q('firstLineChars'),'0')
    rpr = ordered(style, 'rPr', ['name','aliases','basedOn','next','link','autoRedefine',
                                 'hidden','uiPriority','semiHidden','unhideWhenUsed',
                                 'qFormat','locked','personal','personalCompose','personalReply',
                                 'rsid','pPr','rPr','tblPr','trPr','tcPr'])
    col = rpr.find(q('color'))
    if col is None:
        col = etree.SubElement(rpr,q('color'))
    col.attrib.clear();col.set(q('val'),BLACK)
    drop(rpr,'shd');drop(rpr,'highlight')
    sort_ooxml_properties(ppr,PPR_ORDER)


def set_code_paragraph(p, *, fill=CODE_GRAY):
    """Only explicit code style paragraphs; preserve semantic content and indentation."""
    ppr = p.find(q('pPr'))
    if ppr is None:
        ppr = etree.Element(q('pPr'));p.insert(0,ppr)
    drop(ppr, 'pBdr')
    shd = ordered(ppr, 'shd', PPR_ORDER)
    shd.attrib.clear()
    shd.set(q('val'), 'clear');shd.set(q('color'), 'auto');shd.set(q('fill'),fill)
    sort_ooxml_properties(ppr,PPR_ORDER)
    # No forced change to inline code formatting or code/text runs.


def create_academic_table_style(styles):
    """Add a semantic reusable Word table style, without replacing Table Grid."""
    entry = next((s for s in styles.findall(q('style'))
                  if s.get(q('styleId')) == TABLE_STYLE_ID),None)
    if entry is None:
        entry = etree.SubElement(styles,q('style'))
        entry.set(q('type'),'table');entry.set(q('styleId'),TABLE_STYLE_ID)
        entry.set(q('customStyle'),'1')
        name=etree.SubElement(entry,q('name'));name.set(q('val'),TABLE_STYLE_NAME)
        based=etree.SubElement(entry,q('basedOn'));based.set(q('val'),'TableNormal')
    tblpr=entry.find(q('tblPr'))
    if tblpr is None: tblpr=etree.SubElement(entry,q('tblPr'))
    b=ordered(tblpr,'tblBorders',TBLPR_ORDER)
    for child in list(b):b.remove(child)
    for edge in ('top','left','bottom','right','insideH','insideV'):
        border(b,edge,sz=12,show=edge in ('top','bottom'))
    drop(tblpr,'shd')
    # Avoid changing other pre-existing style features or user-defined styles.
    return entry


def normalize_neutral_styles(styles):
    """Apply black ink/white paper to bundled generic paragraph styles only."""
    for st in styles.findall(q('style')):
        if st.get(q('type')) != 'paragraph':continue
        name=st.find(q('name'))
        title=name.get(q('val'),'') if name is not None else ''
        sid=st.get(q('styleId'),'')
        is_code=(title in CODE_STYLE_NAMES or sid in CODE_STYLE_NAMES)
        if is_code:
            set_code_style(st)
            continue
        ppr=st.find(q('pPr'))
        if ppr is not None:
            drop(ppr,'shd')
            # Generic academic mode avoids ornamental title, note and quote rules.
            drop(ppr,'pBdr')
        rpr=st.find(q('rPr'))
        if rpr is not None:
            col=rpr.find(q('color'))
            if col is not None:
                col.attrib.clear();col.set(q('val'),BLACK)
            drop(rpr,'shd');drop(rpr,'highlight')
    create_academic_table_style(styles)


def set_academic_table(table, variant='three-line', *, mark_header=True):
    """Three horizontal rules; explicit grid/none alternatives for exceptional data.

    Changes only formatting in the selected table, never text, merges, or structure.
    """
    if variant not in TABLE_VARIANTS:
        raise ValueError('Unsupported table style: '+str(variant))
    tblpr=table.find(q('tblPr'))
    if tblpr is None:
        tblpr=etree.Element(q('tblPr'));table.insert(0,tblpr)
    st=ordered(tblpr,'tblStyle',TBLPR_ORDER)
    st.set(q('val'),TABLE_STYLE_ID if variant=='three-line' else ('TableGrid' if variant=='grid' else 'TableNormal'))
    drop(tblpr,'shd')
    borders=ordered(tblpr,'tblBorders',TBLPR_ORDER)
    for child in list(borders):borders.remove(child)
    for edge in ('top','left','bottom','right','insideH','insideV'):
        enable=(variant=='grid' or (variant=='three-line' and edge in ('top','bottom')))
        border(borders,edge,sz=4 if variant=='grid' else 12,show=enable)
    rows=table.findall(q('tr'))
    for i,tr in enumerate(rows):
        trpr=tr.find(q('trPr'))
        if trpr is None:
            trpr=etree.Element(q('trPr'));tr.insert(0,trpr)
        if mark_header and i==0:
            hdr=ordered(trpr,'tblHeader',TRPR_ORDER)
            hdr.set(q('val'),'true')
        sort_ooxml_properties(trpr,TRPR_ORDER)
        for cell in tr.findall(q('tc')):
            tcpr=cell.find(q('tcPr'))
            if tcpr is None:
                tcpr=etree.Element(q('tcPr'));cell.insert(0,tcpr)
            drop(tcpr,'shd');drop(tcpr,'tcBorders')
            if variant=='three-line' and i==0:
                cb=ordered(tcpr,'tcBorders',TCPR_ORDER)
                border(cb,'bottom',sz=8,show=True)
            sort_ooxml_properties(tcpr,TCPR_ORDER)
    sort_ooxml_properties(tblpr,TBLPR_ORDER)
    return len(rows)
