"""Editable Word bibliography citations and figure/table/equation cross-references.

Creation paths insert native SEQ and REF fields, with bookmarks around
*only* the numbered portion of their targets. Existing-document repair requires
explicitly authored [[CITE:key]] / [[FIG:key]] / [[TABLE:key]] / [[EQ:key]]
markers and already-established target bookmarks; numeric-looking prose is
never guessed or rewritten. All cached values are provisional until host Word
updates fields. Avoid replacing unrelated field codes, headings, or hyperlinks.
"""
from __future__ import annotations
from copy import deepcopy
from dataclasses import dataclass
from hashlib import sha256
import re
from typing import Any, Mapping
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML = 'http://www.w3.org/XML/1998/namespace'
NS = {'w': W}

def q(name: str) -> str:
    return '{' + W + '}' + name

class ReferenceError(ValueError):
    pass

# Semicolon separates document keys, not display numbers. Keep these labels
# ASCII so they remain stable across mixed operating systems and editors.
KEY_RE = re.compile(r'^[A-Za-z][A-Za-z0-9_.-]{0,63}$')
TOKEN_RE = re.compile(r'\[\[(CITE|FIG|TABLE|EQ|PAGE):([A-Za-z0-9_.;@\-]+)\]\]')
ANY_TOKEN_RE = re.compile(r'\[\[(?:CITE|FIG|TABLE|EQ|PAGE):.*?\]\]')
BOOKMARK_PREFIX = {'bibliography': '_CP_BIB_', 'figure': '_CP_FIG_',
                   'table': '_CP_TAB_', 'equation': '_CP_EQ_'}
DISPLAY_LABEL = {'figure': '图', 'table': '表', 'equation': '式'}
SEQ_NAME = {'bibliography': 'CPBibliography', 'figure': 'CPFigure',
            'table': 'CPTable', 'equation': 'CPEquation'}
TOKEN_KIND = {'CITE': 'bibliography', 'FIG': 'figure',
              'TABLE': 'table', 'EQ': 'equation', 'PAGE': 'page'}

@dataclass(frozen=True)
class Target:
    kind: str
    key: str
    name: str
    cached: str
    title: str = ''


def bookmark_name(kind: str, key: str) -> str:
    if kind not in BOOKMARK_PREFIX:
        raise ReferenceError('Unknown target kind: ' + str(kind))
    if not KEY_RE.fullmatch(key):
        raise ReferenceError('Target key must be ASCII [A-Za-z][A-Za-z0-9_.-]* (max 64 characters).')
    # Include type *and* key; no cross-type collisions even with equal ids.
    return BOOKMARK_PREFIX[kind] + sha256((kind + ':' + key).encode()).hexdigest()[:20]


def new_run(text: str = '', *, rpr: Any = None, superscript: bool = False) -> Any:
    r = etree.Element(q('r'))
    if rpr is not None:
        r.append(deepcopy(rpr))
    if superscript:
        rp = r.find(q('rPr'))
        if rp is None:
            rp = etree.Element(q('rPr'));r.insert(0,rp)
        old=rp.find(q('vertAlign'))
        if old is not None:rp.remove(old)
        vert=etree.SubElement(rp,q('vertAlign'));vert.set(q('val'),'superscript')
    if text:
        t=etree.SubElement(r,q('t'));t.text=text
        if text.startswith(' ') or text.endswith(' ') or '\t' in text:
            t.set('{'+XML+'}space','preserve')
    return r


def field_runs(instruction: str, fallback: str, *, superscript: bool = False,
               rpr: Any = None) -> list[Any]:
    """Construct complex Word field with editable instruction and cached result."""
    begin=new_run(rpr=rpr,superscript=superscript)
    e=etree.SubElement(begin,q('fldChar'));e.set(q('fldCharType'),'begin')
    instr=new_run(rpr=rpr,superscript=superscript)
    e=etree.SubElement(instr,q('instrText'));e.set('{'+XML+'}space','preserve')
    e.text=' '+instruction+' '
    separator=new_run(rpr=rpr,superscript=superscript)
    e=etree.SubElement(separator,q('fldChar'));e.set(q('fldCharType'),'separate')
    cached=new_run(fallback,rpr=rpr,superscript=superscript)
    ending=new_run(rpr=rpr,superscript=superscript)
    e=etree.SubElement(ending,q('fldChar'));e.set(q('fldCharType'),'end')
    return [begin,instr,separator,cached,ending]


def mark_begin(name: str, numeric_id: int) -> Any:
    e=etree.Element(q('bookmarkStart'))
    e.set(q('id'),str(numeric_id));e.set(q('name'),name)
    return e

def mark_end(numeric_id: int) -> Any:
    e=etree.Element(q('bookmarkEnd'));e.set(q('id'),str(numeric_id))
    return e


class Registry:
    def __init__(self, document: Any):
        self.document=document
        self.targets: dict[tuple[str,str],Target] = {}
        self.occupied={e.get(q('name')) for e in document.iter(q('bookmarkStart'))}
        self.max_id=max([int(e.get(q('id'))) for e in document.iter(q('bookmarkStart'))
                         if (e.get(q('id')) or '').isdigit()]+[0])
        self.bookmarks_added=0

    def _allocate(self, kind: str, key: str, cached: str, title: str) -> Target:
        k=(kind,key)
        if k in self.targets:
            raise ReferenceError(f'Duplicate {kind} id: {key}')
        name=bookmark_name(kind,key)
        if name in self.occupied:
            raise ReferenceError('Existing bookmark collides with generated key: '+name)
        self.max_id+=1
        self.occupied.add(name)
        self.bookmarks_added+=1
        target=Target(kind,key,name,cached,title)
        self.targets[k]=target
        return target

    def add_number(self, paragraph: Any, kind: str, key: str, position: int,
                   *, chapter: int | None = None, sep: str = '-') -> Target:
        """Add a caption number, returning an auditable target for REF fields.

        Caller handles prefix and title. For chapter-scoped labels, the bookmark
        spans STYLEREF chapter + separator + SEQ chapter-local number, so
        references update after reordering in Word. Generic prose uses global SEQ.
        """
        if position<=0 or (chapter is not None and chapter<=0):
            raise ReferenceError('Caption cached numbers must be positive.')
        cached=f'{chapter}{sep}{position}' if chapter is not None else str(position)
        target=self._allocate(kind,key,cached,'')
        p=paragraph._p if hasattr(paragraph,'_p') else paragraph
        p.append(mark_begin(target.name,self.max_id))
        if chapter is not None:
            p.extend(field_runs(r'SEQ CPChapter \c',str(chapter)))
            p.append(new_run(sep))
        seq=r'SEQ '+SEQ_NAME[kind]+r' \* ARABIC'
        # Chapter-local SEQ counters are reset explicitly at each numbered
        # Heading 1, not with \s 1 (which LibreOffice may ignore).
        p.extend(field_runs(seq,str(position)))
        p.append(mark_end(self.max_id))
        return target

    def add_bibliography(self, paragraph: Any, key: str, position: int) -> Target:
        target=self._allocate('bibliography',key,str(position),'')
        p=paragraph._p if hasattr(paragraph,'_p') else paragraph
        p.append(new_run('['))
        p.append(mark_begin(target.name,self.max_id))
        p.extend(field_runs(r'SEQ '+SEQ_NAME['bibliography']+r' \* ARABIC',str(position)))
        p.append(mark_end(self.max_id))
        p.append(new_run(']  '))
        return target

    def add_author_year(self, paragraph: Any, key: str, short_label: str) -> Target:
        """Anchor a hidden author/year label, leaving the visible entry editable.

        REF references the deliberately hidden bookmark content, *not* the
        entire bibliography paragraph. Host refresh is still mandatory.
        """
        target=self._allocate('bibliography',key,short_label,'')
        p=paragraph._p if hasattr(paragraph,'_p') else paragraph
        rpr=etree.Element(q('rPr'))
        etree.SubElement(rpr,q('vanish'))
        p.append(mark_begin(target.name,self.max_id))
        p.append(new_run(short_label,rpr=rpr))
        p.append(mark_end(self.max_id))
        return target


def _parse_cluster(payload: str) -> tuple[list[str], str | None]:
    """Parse explicit citation keys; a page span is valid only for one source."""
    parts=payload.split(';')
    if not parts or any(not p for p in parts):
        raise ReferenceError('Empty citation cluster.')
    keys=[];page=None
    for i,part in enumerate(parts):
        key,sep,loc=part.partition('@')
        if not KEY_RE.fullmatch(key):
            raise ReferenceError('Invalid citation key: '+key)
        if sep:
            if len(parts)!=1 or not re.fullmatch(r'\d+(?:-\d+)?',loc):
                raise ReferenceError('Citation page locator requires one key and digits or a page range.')
            page=loc
        keys.append(key)
    if len(set(keys))!=len(keys):
        raise ReferenceError('Duplicate source in a single citation cluster.')
    return keys,page


def plan_citations(blocks: list[dict[str,Any]]) -> list[str]:
    """Return first-citation order, excluding code, bibliography and captions."""
    ordered=[];seen=set()
    for block in blocks:
        if block['type'] not in ('paragraph','note'):
            continue
        for match in TOKEN_RE.finditer(block.get('text','')):
            if match.group(1)!='CITE':continue
            for k in _parse_cluster(match.group(2))[0]:
                if k not in seen:
                    seen.add(k);ordered.append(k)
    return ordered


def _range_groups(keys: list[str], targets: Mapping[tuple[str,str],Target]) -> list[list[str]]:
    """Sort by first-occurrence number and group consecutive bibliographic IDs.

    The displayed endpoints are REF fields. Their *grouping* is generated at
    authoring time, so inserting/reordering source paragraphs requires rerunning
    the bibliography generator and a host field refresh.
    """
    numbered=[]
    for key in keys:
        target=targets.get(('bibliography',key))
        if target is None:raise ReferenceError('Unresolved bibliography id: '+key)
        if not target.cached.isdecimal():raise ReferenceError('Numeric citation needs numbered bibliography targets.')
        numbered.append((int(target.cached),key))
    numbered.sort()
    groups=[]
    for n,key in numbered:
        if groups and n==groups[-1][-1][0]+1:groups[-1].append((n,key))
        else:groups.append([(n,key)])
    return [[key for _,key in group] for group in groups]


def marker_nodes(marker: str, value: str, targets: Mapping[tuple[str,str],Target],
                 *, citation_mode: str = 'numeric-superscript', rpr: Any = None) -> list[Any]:
    kind=TOKEN_KIND[marker]
    if kind=='bibliography':
        keys,page=_parse_cluster(value)
        if citation_mode=='author-year':
            out=[new_run('（',rpr=rpr)]
            for i,key in enumerate(keys):
                target=targets.get(('bibliography',key))
                if target is None:raise ReferenceError('Unresolved bibliography id: '+key)
                if i:out.append(new_run('；',rpr=rpr))
                out.extend(field_runs('REF '+target.name+r' \h',target.cached,rpr=rpr))
            out.append(new_run('）',rpr=rpr))
            if page:out.append(new_run(page,rpr=rpr,superscript=True))
            return out
        if citation_mode not in ('numeric-superscript','numeric-inline'):
            raise ReferenceError('Unsupported citation mode.')
        superscript=citation_mode=='numeric-superscript'
        groups=_range_groups(keys,targets)
        out=[new_run('[',rpr=rpr,superscript=superscript)]
        for idx,group in enumerate(groups):
            if idx:out.append(new_run(',',rpr=rpr,superscript=superscript))
            # A range of two or more consecutive sources has two live REF
            # endpoints; singletons retain their normal one-field form.
            endpoint_keys=group if len(group)==1 else (group[0],group[-1])
            for j,key in enumerate(endpoint_keys):
                if j:out.append(new_run('-',rpr=rpr,superscript=superscript))
                ref=targets[('bibliography',key)]
                out.extend(field_runs('REF '+ref.name+r' \h',ref.cached,rpr=rpr,
                                      superscript=superscript))
        out.append(new_run(']',rpr=rpr,superscript=superscript))
        if page:out.append(new_run(page,rpr=rpr,superscript=True))
        return out
    if not KEY_RE.fullmatch(value):
        raise ReferenceError('Invalid cross-reference id: '+value)
    target_type='figure' if kind=='figure' or kind=='page' and ('figure',value) in targets else kind
    if kind=='page':
        choices=[v for (t,k),v in targets.items() if k==value]
        if len(choices)!=1:
            raise ReferenceError('PAGE:key requires exactly one matching target id: '+value)
        ref=choices[0]
        return field_runs('PAGEREF '+ref.name+r' \h','1',rpr=rpr)
    target=targets.get((target_type,value))
    if target is None:raise ReferenceError(f'Unresolved {kind} id: {value}')
    prefix=DISPLAY_LABEL[kind]
    if kind=='equation':
        return [new_run(prefix+'（',rpr=rpr),
                *field_runs('REF '+target.name+r' \h',target.cached,rpr=rpr),
                new_run('）',rpr=rpr)]
    return [new_run(prefix,rpr=rpr),
            *field_runs('REF '+target.name+r' \h',target.cached,rpr=rpr)]


def replace_markers(paragraphs: list[Any], targets: Mapping[tuple[str,str],Target],
                    *, citation_mode: str = 'numeric-superscript',
                    strict: bool = True) -> dict[str,Any]:
    """Modify only individual plain-text runs; never flatten rich paragraph XML.

    If a marker is split across runs or a target is unknown, strict mode errors
    out *before* caller writes the result. Non-strict mode returns an audit issue
    without rewriting problematic runs. Existing unrelated REF fields survive.
    """
    report={'replaced':0,'unresolved':[],'split_run_markers':0}
    for para in paragraphs:
        p=para._p if hasattr(para,'_p') else para
        # Forbid markers inside code and captions; no auto-classification of code.
        style=p.find('w:pPr/w:pStyle',NS)
        if style is not None and style.get(q('val')) in ('CodeBlock','代码段'):
            continue
        runs=list(p.findall(q('r')))
        full=''.join(t.text or '' for r in runs for t in r.findall(q('t')))
        has_any=ANY_TOKEN_RE.findall(full)
        handled=0
        for r in runs:
            texts=r.findall(q('t'))
            if len(texts)!=1 or not TOKEN_RE.search(texts[0].text or ''):
                continue
            # Complex run could contain breaks/tab/drawing/fields; skip safely.
            if any(c.tag not in (q('rPr'),q('t')) for c in r):
                report['unresolved'].append({'reason':'complex-run','fragment':(texts[0].text or '')[:60]})
                continue
            source=texts[0].text or ''
            new_nodes=[]
            rpr=r.find(q('rPr'))
            previous=0
            try:
                for m in TOKEN_RE.finditer(source):
                    if m.start()>previous:
                        new_nodes.append(new_run(source[previous:m.start()],rpr=rpr))
                    new_nodes.extend(marker_nodes(m.group(1),m.group(2),targets,
                                                  citation_mode=citation_mode,rpr=rpr))
                    previous=m.end()
                if previous<len(source):
                    new_nodes.append(new_run(source[previous:],rpr=rpr))
            except ReferenceError as exc:
                report['unresolved'].append({'reason':str(exc),'fragment':source[:60]})
                continue
            parent=r.getparent()
            if parent is None:raise ReferenceError('Detached run')
            loc=parent.index(r)
            parent.remove(r)
            for node in new_nodes:
                parent.insert(loc,node);loc+=1
            n=len(TOKEN_RE.findall(source))
            report['replaced']+=n;handled+=n
        remaining=ANY_TOKEN_RE.findall(''.join(t.text or '' for t in p.iter(q('t'))))
        if remaining:
            report['split_run_markers']+=len(remaining)
            report['unresolved'].append({'reason':'unconverted/split marker',
                                         'fragment':remaining[0][:60]})
    if strict and report['unresolved']:
        raise ReferenceError('Reference markers need review: '+str(report['unresolved'][:4]))
    return report


def audit_fields(document: Any, *, citation_mode: str = "numeric-superscript") -> dict[str,Any]:
    """Check REF targets and citation superscripts; report no fake field refresh."""
    book=[b.get(q('name')) for b in document.iter(q('bookmarkStart'))]
    counts={x:book.count(x) for x in set(book) if x}
    issues=[];warnings=[]
    for name,n in counts.items():
        if n>1:issues.append({'code':'duplicate-bookmark','name':name,'count':n})
    all_b=set(book)
    ref_count=0;pageref_count=0;citations=0;manual=0
    for instr in document.iter(q('instrText')):
        text=' '.join((instr.text or '').strip().split())
        m=re.match(r'^(REF|PAGEREF)\s+([A-Za-z0-9_]+)(?:\s|$)',text,re.I)
        if m:
            field_kind,name=m.groups()
            if field_kind.upper()=='REF':ref_count+=1
            else:pageref_count+=1
            if name not in all_b:
                issues.append({'code':'dangling-reference','bookmark':name})
            if name.startswith(BOOKMARK_PREFIX['bibliography']):
                citations+=1
                parent=instr.getparent()
                has_superscript=(parent is not None and parent.find('w:rPr/w:vertAlign',NS) is not None
                    and parent.find('w:rPr/w:vertAlign',NS).get(q('val'))=='superscript')
                if citation_mode=='numeric-superscript' and not has_superscript:
                    warnings.append({'code':'citation-not-superscript','bookmark':name,
                                     'note':'Generated numeric superscript citation should be raised.'})
                elif citation_mode!='numeric-superscript' and has_superscript:
                    warnings.append({'code':'unexpected-superscript-citation','bookmark':name})
    rem=[m.group(0) for node in document.iter(q('t')) for m in ANY_TOKEN_RE.finditer(node.text or '')]
    if rem:issues.append({'code':'unresolved-markers','count':len(rem),'examples':rem[:3]})
    # A marker can also be split across runs (no single w:t contains it).
    if b'[[' in etree.tostring(document) and not rem:
        for p in document.iter(q('p')):
            text=''.join(t.text or '' for t in p.iter(q('t')))
            if ANY_TOKEN_RE.search(text):
                issues.append({'code':'split-unresolved-markers','example':text[:100]});break
    # Paragraph position matters for sequential bibliography: moving an early
    # citation after a different citation can invalidate its order even though
    # all REF fields remain intact. Never silently reorder user's entries.
    bib_prefix=BOOKMARK_PREFIX['bibliography']
    first_mentions=[];seen_mentions=set();listed=[]
    for para in document.iter(q('p')):
        for child in para.iter():
            if child.tag==q('instrText'):
                m=re.match(r'\s*REF\s+('+re.escape(bib_prefix)+r'[A-Fa-f0-9]+)',child.text or '')
                if m:
                    name=m.group(1)
                    if name not in seen_mentions:
                        first_mentions.append(name);seen_mentions.add(name)
            elif child.tag==q('bookmarkStart'):
                name=child.get(q('name'),'')
                if name.startswith(bib_prefix):listed.append(name)
    if citation_mode!='author-year' and first_mentions != [name for name in listed if name in seen_mentions]:
        warnings.append({'code':'citation-order-mismatch',
                         'note':'Bibliography order differs from first in-text citation order; review/reorder explicitly.'})
    unused=[name for name in listed if name not in seen_mentions]
    if unused:warnings.append({'code':'uncited-bibliography-entries','count':len(unused),
                               'note':'May be an intentional further-reading list.'})
    targets={kind:sum(n.startswith(prefix) for n in all_b)
             for kind,prefix in BOOKMARK_PREFIX.items()}
    return {'targets':targets,'ref_fields':ref_count,'page_fields':pageref_count,
            'citation_ref_fields':citations,'errors':issues,'warnings':warnings,
            'review_required':bool(issues or warnings),
            'field_refresh':'pending Word field update; cached values are not authoritative'}
