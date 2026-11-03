"""Native Word heading numbering profiles (OOXML multilevel numbering).

Profiles are editorial defaults, *not* claims that any school mandates them.
Existing documents are never automatically renumbered by mere audit/import.
"""
from __future__ import annotations

import re
from hashlib import sha256
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}

def q(name: str) -> str:
    return '{' + W + '}' + name

# format, Word field pattern, separation after number label
SCHEMES = {
    'chinese-formal': (
        ('chineseCounting', '%1、', 'nothing'),
        ('chineseCounting', '（%2）', 'nothing'),
        ('decimal', '%3.', 'space'),
    ),
    'academic-decimal': (
        ('decimal', '%1', 'space'),
        ('decimal', '%1.%2', 'space'),
        ('decimal', '%1.%2.%3', 'space'),
    ),
    'academic-chapter': (
        ('decimal', '第%1章', 'space'),
        ('decimal', '%1.%2', 'space'),
        ('decimal', '%1.%2.%3', 'space'),
    ),
    'humanities': (
        ('chineseCounting', '第%1章', 'space'),
        ('chineseCounting', '%2、', 'nothing'),
        ('chineseCounting', '（%3）', 'nothing'),
    ),
}
DOCUMENT_TYPE_SCHEME = {
    'general': 'chinese-formal',
    'formal-report': 'chinese-formal',
    'teaching': 'chinese-formal',
    'project-report': 'chinese-formal',
    'research-report': 'academic-decimal',
    'technical-report': 'academic-decimal',
    'thesis-natural': 'academic-decimal',
    'thesis-chapter': 'academic-chapter',
    'thesis-humanities': 'humanities',
}
NUMBERING_CHOICES = ('auto', 'none', *SCHEMES)
DOCUMENT_TYPES = tuple(DOCUMENT_TYPE_SCHEME)
HEADER_STYLES = {'Heading1': 1, 'Heading2': 2, 'Heading3': 3}
SKIP_LABELS = frozenset({
    '摘要', '中文摘要', 'abstract', '目录', '目 录', 'contents', '参考文献',
    'bibliography', '致谢', '后记', '缩略语', '符号和缩略语说明', '插图目录',
    '附表目录', '附录', '附录a', '附录b', '学位论文声明',
})
# Existing manually numbered heading text is NOT edited: avoid accidental doubles.
MANUAL_PREFIX = re.compile(
    r'^\s*(?:第\s*[一二三四五六七八九十百千零〇两\d]+\s*章'
    r'|[一二三四五六七八九十百零〇两]+\s*[、．.]'
    r'|[（(]\s*[一二三四五六七八九十百零〇两\d]+\s*[）)]'
    r'|\d+(?:\.\d+)+(?:[.．、])?\s*'
    r'|\d+[.．、]\s*)'
)


def select_scheme(document_type: str = 'general', numbering: str = 'auto') -> str:
    if numbering not in NUMBERING_CHOICES:
        raise ValueError('Unknown numbering scheme: ' + numbering)
    if document_type not in DOCUMENT_TYPES:
        raise ValueError('Unknown document_type: ' + document_type)
    return DOCUMENT_TYPE_SCHEME[document_type] if numbering == 'auto' else numbering


def text_of(p) -> str:
    return ''.join(e.text or '' for e in p.iter(q('t')))


def is_special_heading(text: str) -> bool:
    label = ''.join(text.strip().split()).casefold()
    return label in {''.join(x.split()).casefold() for x in SKIP_LABELS} or bool(
        re.match(r'^附录\s*[A-ZＡ-Ｚ](?:$|\s|[、：:])', text.strip(), re.I))


def has_manual_number(text: str) -> bool:
    return MANUAL_PREFIX.match(text) is not None


def numbered_paragraph(p, level: int, num_id: int) -> bool:
    """Assign native, editable Word list numbering without touching any text run."""
    ppr = p.find(q('pPr'))
    if ppr is None:
        ppr = etree.Element(q('pPr'))
        p.insert(0, ppr)
    old = ppr.find(q('numPr'))
    if old is not None:
        existing_ilvl = old.find(q('ilvl'))
        existing_num = old.find(q('numId'))
        if (existing_num is not None and existing_ilvl is not None
                and existing_num.get(q('val')) == str(num_id)
                and existing_ilvl.get(q('val')) == str(level - 1)):
            return False
        ppr.remove(old)
    numpr = etree.Element(q('numPr'))
    ilvl = etree.SubElement(numpr, q('ilvl'))
    ilvl.set(q('val'), str(level - 1))
    num = etree.SubElement(numpr, q('numId'))
    num.set(q('val'), str(num_id))
    # Document schema: numPr follows pStyle/keepNext/.../widowControl.
    before = {'pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl'}
    idx = 0
    for i, child in enumerate(ppr):
        if etree.QName(child).localname in before:
            idx = i + 1
    ppr.insert(idx, numpr)
    return True


def suppress_number(p) -> bool:
    """Protect existing special/non-numbered headings from profile style linkage."""
    ppr = p.find(q('pPr'))
    if ppr is None:
        ppr = etree.Element(q('pPr')); p.insert(0, ppr)
    existing = ppr.find(q('numPr'))
    if existing is not None:
        id_el = existing.find(q('numId'))
        if id_el is not None and id_el.get(q('val')) == '0':
            return False
    return numbered_paragraph(p, 1, 0)



def remove_manual_prefix_if_safe(p) -> bool:
    """Opt-in: remove prefix in ordinary heading text runs; do not touch fields.

    Unlike assigning p.text in python-docx, this preserves each existing run,
    emphasis and non-text XML. Reject complex paragraph constructs entirely.
    """
    full=text_of(p)
    match=MANUAL_PREFIX.match(full)
    if match is None or not full[match.end():].strip():
        return False
    # Do not rewrite runs containing hyperlinks, bookmarks, images, field codes,
    # changes, inline controls or non-text components.
    for el in p.iter():
        if etree.QName(el).localname in {'hyperlink','sdt','fldChar','instrText',
                'drawing','pict','tab','br','bookmarkStart','bookmarkEnd',
                'ins','del','moveFrom','moveTo'}:
            return False
    text_nodes=p.findall('./w:r/w:t',NS)
    if ''.join((node.text or '') for node in text_nodes)!=full:
        return False
    remaining=match.end()
    for node in text_nodes:
        if remaining<=0: break
        data=node.text or ''
        if len(data)<=remaining:
            node.text='';remaining-=len(data)
        else:
            node.text=data[remaining:]
            remaining=0
            if node.text.startswith(' ') or node.text.endswith(' '):
                node.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    return remaining==0

def get_num(p) -> tuple[str | None, str | None]:
    el = p.find('w:pPr/w:numPr', NS)
    if el is None:
        return (None, None)
    id_el, ilvl = el.find(q('numId')), el.find(q('ilvl'))
    return (id_el.get(q('val')) if id_el is not None else None,
            ilvl.get(q('val')) if ilvl is not None else None)


def _find_abstract_by_name(root, name):
    for entry in root.findall(q('abstractNum')):
        n = entry.find(q('name'))
        if n is not None and n.get(q('val')) == name:
            return entry
    return None


def ensure_multilevel_numbering(root, scheme: str) -> int:
    """Create or reuse a uniquely named abstract definition and its concrete num.

    This expects a normal Word docx numbering part. The existing list numbering
    and all original abstractNum definitions are retained unchanged.
    """
    if scheme not in SCHEMES:
        raise ValueError('Unsupported scheme: ' + scheme)
    name = 'CPDOCX-HEADING-' + scheme
    abstract = _find_abstract_by_name(root, name)
    if abstract is None:
        ids = [int(e.get(q('abstractNumId'))) for e in root.findall(q('abstractNum'))]
        abstract_id = max(ids, default=-1) + 1
        abstract = etree.Element(q('abstractNum'))
        abstract.set(q('abstractNumId'), str(abstract_id))
        nsid = etree.SubElement(abstract, q('nsid'))
        nsid.set(q('val'), sha256(name.encode()).hexdigest()[:8].upper())
        m = etree.SubElement(abstract, q('multiLevelType')); m.set(q('val'), 'multilevel')
        aname = etree.SubElement(abstract, q('name')); aname.set(q('val'), name)
        for idx, (fmt, label, suffix) in enumerate(SCHEMES[scheme]):
            lvl = etree.SubElement(abstract, q('lvl')); lvl.set(q('ilvl'),str(idx))
            start = etree.SubElement(lvl, q('start')); start.set(q('val'),'1')
            num_fmt = etree.SubElement(lvl,q('numFmt'));num_fmt.set(q('val'),fmt)
            if idx > 0:
                restart=etree.SubElement(lvl,q('lvlRestart'));restart.set(q('val'),str(idx))
            pstyle=etree.SubElement(lvl,q('pStyle'));pstyle.set(q('val'),f'Heading{idx+1}')
            suff=etree.SubElement(lvl,q('suff'));suff.set(q('val'),suffix)
            lt=etree.SubElement(lvl,q('lvlText'));lt.set(q('val'),label)
            just=etree.SubElement(lvl,q('lvlJc'));just.set(q('val'),'left')
            ppr=etree.SubElement(lvl,q('pPr'))
            ind=etree.SubElement(ppr,q('ind'))
            ind.set(q('left'),'0');ind.set(q('hanging'),'0')
        # OOXML numbering part places abstractNum entries before num instances.
        first_num = root.find(q('num'))
        root.insert(root.index(first_num) if first_num is not None else len(root),abstract)
    aid = abstract.get(q('abstractNumId'))
    for num in root.findall(q('num')):
        ref=num.find(q('abstractNumId'))
        if ref is not None and ref.get(q('val')) == aid:
            return int(num.get(q('numId')))
    ids = [int(e.get(q('numId'))) for e in root.findall(q('num'))]
    num_id = max(ids,default=0)+1
    new = etree.SubElement(root,q('num'));new.set(q('numId'),str(num_id))
    ref=etree.SubElement(new,q('abstractNumId'));ref.set(q('val'),aid)
    return num_id


def recognized_scheme(root, num_id: str | int) -> str | None:
    entry=root.find('w:num[@w:numId="'+str(num_id)+'"]', NS)
    if entry is None:return None
    ref=entry.find(q('abstractNumId'))
    if ref is None:return None
    abstract=root.find('w:abstractNum[@w:abstractNumId="'+ref.get(q('val'))+'"]',NS)
    if abstract is None:return None
    n=abstract.find(q('name'))
    if n is None:return None
    raw=n.get(q('val'),'')
    prefix='CPDOCX-HEADING-'
    return raw[len(prefix):] if raw.startswith(prefix) else None


def analyze_heading_paragraphs(docroot, number_root=None):
    entries=[];last_level=0; jumps=0
    for p in docroot.findall('./w:body/w:p',NS):
        st=p.find('w:pPr/w:pStyle',NS)
        sid=st.get(q('val')) if st is not None else ''
        level=HEADER_STYLES.get(sid)
        if level is None: continue
        val=text_of(p)
        n,ilvl=get_num(p)
        special=is_special_heading(val)
        manual=has_manual_number(val)
        if not special:
            if level-last_level>1:jumps+=1
            last_level=level
        entries.append({'level':level,'style':sid,'text':val,'special':special,
                        'manual_number':manual,'num_id':n,'ilvl':ilvl,
                        'numbering':recognized_scheme(number_root,n) if number_root is not None and n not in (None,'0') else None})
    return entries,jumps
