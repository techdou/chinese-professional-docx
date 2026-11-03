#!/usr/bin/env python3
"""Create, inspect and conservatively repair Chinese professional DOCX/DOTX files.

No network access, macro execution, in-place writes or implicit overwrites.
Formatting repairs patch OOXML parts; they do not rebuild existing documents.
"""
from __future__ import annotations

import argparse
from collections import Counter
from copy import deepcopy
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path, PurePosixPath
import re
import sys
from typing import Any
from zipfile import BadZipFile, ZipFile, ZIP_DEFLATED

from lxml import etree

from bibliography_7714 import (STANDARD as DEFAULT_REFERENCE_STANDARD, STANDARDS as REFERENCE_STANDARDS,
                               CITATION_MODES, BibliographyError, audit_bibliography,
                               format_entry, assign_year_suffix, citation_label)
from academic_styles import (TABLE_VARIANTS, TABLE_STYLE_ID, CODE_STYLE_NAMES,
                             set_academic_table, set_code_style, set_code_paragraph,
                             create_academic_table_style)
from cross_references import (ReferenceError, Registry, KEY_RE, TOKEN_RE, ANY_TOKEN_RE,
                              BOOKMARK_PREFIX, Target, bookmark_name, plan_citations, _parse_cluster,
                              replace_markers, audit_fields, q as ref_q, NS as REF_NS)
from heading_numbering import (SCHEMES, NUMBERING_CHOICES, DOCUMENT_TYPES,
                              HEADER_STYLES, select_scheme, text_of,
                              has_manual_number, is_special_heading, get_num,
                              remove_manual_prefix_if_safe,
                              numbered_paragraph, suppress_number,
                              ensure_multilevel_numbering, recognized_scheme,
                              analyze_heading_paragraphs)

ROOT = Path(__file__).resolve().parents[1]
PROFILE = json.loads((ROOT / 'assets/style-profile.json').read_text(encoding='utf-8'))
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
XML = 'http://www.w3.org/XML/1998/namespace'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
DOCX_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
DOTX_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml'
SCENES = {'generic': 'templates/base.docx', 'competition': 'examples/competition.docx',
          'handout': 'examples/handout.docx', 'report': 'examples/report.docx',
          'components': 'examples/components.docx'}
MAX_TOTAL = 128 * 1024 * 1024
MAX_XML = 16 * 1024 * 1024
MANUAL_LIST = re.compile(
    r'^(?:[•●▪■]|[-+]\s+|(?:\d{1,3}|[一二三四五六七八九十]+)[、)）]'
    r'|\d{1,3}\.(?!\d)|[一二三四五六七八九十]+\.'
    r'|[（(](?:\d{1,3}|[一二三四五六七八九十]+)[)）])\s*')


class DocumentError(ValueError):
    """Actionable validation or safety error."""


def q(name: str) -> str:
    return '{' + W + '}' + name


def canon(value: str) -> str:
    return re.sub(r'\s+', '', value).casefold()


def attrs(element: Any) -> dict[str, str]:
    return {} if element is None else {etree.QName(k).localname: v for k, v in element.attrib.items()}


def parse_xml(data: bytes) -> Any:
    if len(data) > MAX_XML or b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
        raise DocumentError('XML exceeds the size limit or declares a forbidden DTD/entity.')
    try:
        return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True,
                                                      remove_blank_text=False, huge_tree=False))
    except etree.XMLSyntaxError as exc:
        raise DocumentError('Invalid OOXML: ' + str(exc)) from exc


def xml_bytes(root: Any) -> bytes:
    return etree.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def write_new(path: Path, content: bytes) -> None:
    """Exclusive creation: neither a source nor an existing output is overwritten."""
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open('xb') as f:
            f.write(content)
    except FileExistsError as exc:
        raise DocumentError(f'Output already exists; choose a new filename: {path}') from exc


class Package:
    def __init__(self, path: Path):
        self.path = Path(path)
        if self.path.suffix.lower() not in ('.docx', '.dotx'):
            raise DocumentError('Only unencrypted .docx/.dotx are supported; convert .doc separately.')
        try:
            with ZipFile(self.path) as z:
                infos = z.infolist()
                names = [i.filename for i in infos]
                if len(names) != len(set(names)):
                    raise DocumentError('Duplicate ZIP members are not supported.')
                if len(names) > 10000 or sum(i.file_size for i in infos) > MAX_TOTAL:
                    raise DocumentError('Document exceeds the 128 MiB uncompressed safety limit.')
                for name in names:
                    parts = PurePosixPath(name).parts
                    if name.startswith('/') or '\\' in name or '..' in parts:
                        raise DocumentError('Unsafe ZIP member path: ' + name)
                    if 'vbaproject' in name.casefold():
                        raise DocumentError('Macro-bearing files are not supported, even if renamed.')
                self.infos = infos
                self.data = {i.filename: z.read(i) for i in infos}
        except (BadZipFile, RuntimeError, OSError) as exc:
            raise DocumentError('Unable to read the unencrypted OOXML package: ' + str(exc)) from exc
        for name in ('[Content_Types].xml', 'word/document.xml', 'word/styles.xml'):
            if name not in self.data:
                raise DocumentError('Missing required OOXML part: ' + name)
        ct = parse_xml(self.data['[Content_Types].xml'])
        cts = [el.get('ContentType') for el in ct if el.get('PartName') == '/word/document.xml']
        expected = DOTX_CT if self.path.suffix.lower() == '.dotx' else DOCX_CT
        if cts != [expected]:
            raise DocumentError('File extension and OOXML main content type do not match.')
        self.styles = parse_xml(self.data['word/styles.xml'])
        self.document = parse_xml(self.data['word/document.xml'])
        if self.document.tag != q('document'):
            raise DocumentError('Strict/unknown OOXML namespace: use a compatible Word export first.')
        self.style_map = {e.get(q('styleId')): e for e in self.styles.findall('w:style', NS)
                          if e.get(q('type')) == 'paragraph'}
        self.default = next((k for k, v in self.style_map.items()
                             if v.get(q('default')) == '1'), 'Normal')
        self._verify_chains()

    def _verify_chains(self) -> None:
        for sid in self.style_map:
            self.chain(sid)

    def chain(self, sid: str) -> list[Any]:
        chain, seen = [], set()
        while sid in self.style_map:
            if sid in seen:
                raise DocumentError('Cyclic paragraph style inheritance: ' + sid)
            seen.add(sid)
            st = self.style_map[sid]
            chain.append(st)
            base = st.find('w:basedOn', NS)
            sid = base.get(q('val'), '') if base is not None else ''
        return list(reversed(chain))

    def sid(self, paragraph: Any) -> str:
        el = paragraph.find('w:pPr/w:pStyle', NS)
        return el.get(q('val'), self.default) if el is not None else self.default

    def keys(self, sid: str) -> set[str]:
        result = {canon(sid)}
        st = self.style_map.get(sid)
        name = st.find('w:name', NS) if st is not None else None
        if name is not None:
            result.add(canon(name.get(q('val'), '')))
        return result

    def prop(self, sid: str, path: str, paragraph: Any = None) -> dict[str, str]:
        """Merge property attributes through defaults, base styles and direct formatting."""
        group, _, rest = path.partition('/')
        default_path = 'w:docDefaults/w:' + group[2:] + 'Default/' + path
        result = attrs(self.styles.find(default_path, NS))
        for st in self.chain(sid):
            result.update(attrs(st.find(path, NS)))
        if paragraph is not None:
            direct = path if group == 'w:pPr' else 'w:pPr/w:rPr/' + rest
            result.update(attrs(paragraph.find(direct, NS)))
        return result

    def body_ids(self, extras: list[str] | None = None) -> set[str]:
        names = {canon(n) for n in PROFILE['body_style_keys'] + (extras or [])}
        result = {sid for sid in self.style_map if self.keys(sid) & names}
        result.add(self.default)
        # A named derivative of Body Text is prose; a derivative of Normal alone is ambiguous.
        for sid in self.style_map:
            if any(st.get(q('styleId')) in result - {self.default, 'Normal'} for st in self.chain(sid)):
                result.add(sid)
        # A named special style stays special even when based on Body Text.
        special_keys = {canon(x) for x in PROFILE['special_style_keys']}
        structural = set()
        for sid in self.style_map:
            keys = set().union(*(self.keys(st.get(q('styleId'))) for st in self.chain(sid)))
            outline = self.prop(sid, 'w:pPr/w:outlineLvl').get('val')
            num = self.prop(sid, 'w:pPr/w:numPr/w:numId').get('val')
            if (keys & special_keys or any(k.startswith(('heading', 'toc', 'list')) for k in keys)
                    or (outline is not None and int(outline) < 9) or (num is not None and num != '0')):
                structural.add(sid)
        result -= structural
        if extras:
            if any(canon(x) in self.keys(s) for x in extras for s in structural):
                raise DocumentError('--body-style cannot override a structural heading/list/special style.')
            missing = [x for x in extras if not any(canon(x) in self.keys(s) for s in self.style_map)]
            if missing:
                raise DocumentError('Unknown --body-style name/ID: ' + ', '.join(missing))
        return result

    def text(self, paragraph: Any) -> str:
        return ''.join(e.text or '' for e in paragraph.iter(q('t')))

    def classify(self, p: Any, bodies: set[str]) -> str:
        ancestors = {etree.QName(e).localname for e in p.iterancestors()}
        if 'tc' in ancestors:
            return 'table'
        if 'txbxContent' in ancestors:
            return 'textbox'
        if ancestors & {'ins', 'del', 'moveFrom', 'moveTo'}:
            return 'tracked'
        if 'sdt' in ancestors:
            return 'content-control'
        sid = self.sid(p)
        outline = self.prop(sid, 'w:pPr/w:outlineLvl', p).get('val')
        if outline is not None and int(outline) < 9:
            return 'heading'
        num = self.prop(sid, 'w:pPr/w:numPr/w:numId', p).get('val')
        if num is not None and num != '0':
            return 'list'
        special = {canon(x) for x in PROFILE['special_style_keys']}
        for st in self.chain(sid):
            keys = self.keys(st.get(q('styleId')))
            if any(k.startswith(('heading', 'toc', 'list')) for k in keys) or keys & special:
                return 'special'
        text = self.text(p)
        if p.find('.//w:drawing', NS) is not None or p.find('.//w:pict', NS) is not None:
            return 'mixed-drawing'
        if p.xpath('.//*[local-name()="oMath" or local-name()="oMathPara"]'):
            return 'equation'
        if not text.strip():
            return 'empty-body' if sid in bodies else 'empty'
        if MANUAL_LIST.match(text.lstrip()):
            return 'ambiguous-manual-list'
        return 'body' if sid in bodies else 'unclassified'

    def assert_editable(self) -> None:
        if any(n.startswith('_xmlsignatures/') for n in self.data):
            raise DocumentError('Digitally signed file: obtain an unsigned working copy before editing.')
        for name, data in self.data.items():
            if not re.fullmatch(r'word/(?:document|header\d*|footer\d*|footnotes|endnotes|comments\d*)\.xml', name):
                continue
            story = self.document if name == 'word/document.xml' else parse_xml(data)
            if story.xpath('.//w:ins|.//w:del|.//w:moveFrom|.//w:moveTo|'
                           './/w:pPrChange|.//w:rPrChange|.//w:sectPrChange', namespaces=NS):
                raise DocumentError('Tracked changes detected in ' + name +
                                    '. Resolve the review workflow before repair.')
        if 'word/settings.xml' in self.data:
            settings = parse_xml(self.data['word/settings.xml'])
            for el in settings.findall('w:documentProtection', NS):
                if el.get(q('enforcement'), '1') not in ('0', 'false', 'off'):
                    raise DocumentError('Document protection is enabled; use an authorized editable copy.')

    def serialized(self, changed: dict[str, bytes]) -> bytes:
        if not changed:
            return self.path.read_bytes()
        buf = BytesIO()
        with ZipFile(buf, 'w', compression=ZIP_DEFLATED) as z:
            for info in self.infos:
                z.writestr(info, changed.get(info.filename, self.data[info.filename]))
        return buf.getvalue()


def ensure(parent: Any, tag: str) -> Any:
    found = parent.find('w:' + tag, NS)
    if found is None:
        found = etree.SubElement(parent, q(tag))
    return found


# WordprocessingML order for elements that this script inserts into pPr.
PPR_ORDER = ['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr',
             'suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku',
             'wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi',
             'adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents',
             'suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap',
             'outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange']


def ensure_ppr(parent: Any) -> Any:
    ppr = parent.find('w:pPr', NS)
    if ppr is None:
        ppr = etree.Element(q('pPr'))
        if parent.tag == q('p'):
            parent.insert(0, ppr)
        else:
            rpr = parent.find('w:rPr', NS)
            parent.insert(parent.index(rpr) if rpr is not None else len(parent), ppr)
    return ppr


def set_indent(parent: Any, values: dict[str, str], *, replace: bool = False) -> None:
    ppr = ensure_ppr(parent)
    ind = ppr.find('w:ind', NS)
    if ind is None:
        ind = etree.Element(q('ind'))
        rank = PPR_ORDER.index('ind')
        index = next((i for i, e in enumerate(ppr)
                      if etree.QName(e).localname in PPR_ORDER
                      and PPR_ORDER.index(etree.QName(e).localname) > rank), len(ppr))
        ppr.insert(index, ind)
    if replace:
        ind.attrib.clear()
    else:
        for key in ('hanging', 'hangingChars'):
            ind.attrib.pop(q(key), None)
    for key, value in values.items():
        ind.set(q(key), str(value))


def expected_fallback(pkg: Package, sid: str, p: Any = None) -> int:
    # size in half-points × 10 twips/half-point × two characters.
    return round(float(pkg.prop(sid, 'w:rPr/w:sz', p).get('val', '24')) * 20)


def body_issues(pkg: Package, sid: str, p: Any = None) -> list[tuple[str, str]]:
    ind = pkg.prop(sid, 'w:pPr/w:ind', p)
    issues = []
    if ind.get('firstLineChars') != '200':
        issues.append(('first-line-chars', 'Expected firstLineChars=200 (two characters).'))
    if any(float(ind.get(k, 0)) != 0 for k in ('hanging', 'hangingChars')):
        issues.append(('hanging-conflict', 'Body prose has an inherited/direct hanging indent.'))
    if ind.get('firstLine') != str(expected_fallback(pkg, sid, p)):
        issues.append(('fallback-indent', 'Twip fallback does not match two characters at the paragraph style size.'))
    return issues


def leading_padding_nodes(paragraph: Any) -> list[tuple[Any, Any, int]]:
    """Find safely editable prefix nodes in document order; stop at complex content."""
    found = []
    for child in paragraph:
        if child.tag in (q('pPr'), q('bookmarkStart'), q('bookmarkEnd'), q('proofErr')):
            continue
        if child.tag != q('r'):
            break  # A hyperlink/field/content control is a semantic boundary.
        for node in child:
            if node.tag == q('rPr'):
                continue
            if node.tag == q('tab'):
                found.append((child, node, 1))
            elif node.tag == q('t'):
                text = node.text or ''
                count = len(text) - len(text.lstrip(' \u3000\t'))
                if count:
                    found.append((child, node, count))
                if text[count:]:
                    return found
            else:
                return found
    return found


def strip_leading_padding(paragraph: Any) -> int:
    removed = 0
    for run, node, count in leading_padding_nodes(paragraph):
        if node.tag == q('tab'):
            run.remove(node)
        else:
            node.text = (node.text or '')[count:]
        removed += count
    return removed


def audit(path: Path, extras: list[str] | None = None, *,
          scope: str = 'generic', details: bool = False) -> dict[str, Any]:
    if scope not in ('generic', 'indent'):
        raise DocumentError('Audit scope must be generic or indent.')
    pkg = Package(path)
    bodies = pkg.body_ids(extras)
    errors, warnings = [], []
    for sid in sorted(bodies & pkg.style_map.keys()):
        for code, message in body_issues(pkg, sid):
            target = warnings if code == 'fallback-indent' else errors
            target.append({'code':code,'location':'style:' + sid,'message':message})
    counts: Counter[str] = Counter()
    paragraph_details = []
    for i, p in enumerate(pkg.document.findall('.//w:body//w:p', NS), 1):
        role = pkg.classify(p, bodies)
        counts[role] += 1
        if details:
            sid = pkg.sid(p)
            style = pkg.style_map.get(sid)
            name = style.find('w:name', NS) if style is not None else None
            paragraph_details.append({'index': i, 'style_id': sid,
                'style_name': name.get(q('val'), sid) if name is not None else sid,
                'role': role, 'effective_indent': pkg.prop(sid, 'w:pPr/w:ind', p)})
        if role != 'body':
            if role in ('unclassified', 'ambiguous-manual-list', 'mixed-drawing', 'equation', 'content-control'):
                warnings.append({'code':'manual-review','location':f'paragraph:{i}',
                                 'message':'Not automatically treated as prose: ' + role})
            continue
        for code, message in body_issues(pkg, pkg.sid(p), p):
            (warnings if code == 'fallback-indent' else errors).append(
                {'code':code,'location':f'paragraph:{i}','message':message})
        text = pkg.text(p)
        if leading_padding_nodes(p):
            warnings.append({'code':'manual-padding','location':f'paragraph:{i}',
                             'message':'Inspect leading spaces/Tab; not removed without explicit approval.'})
        if p.find('.//w:br', NS) is not None:
            warnings.append({'code':'soft-break','location':f'paragraph:{i}',
                             'message':'Shift+Enter is not a new paragraph; inspect the intended paragraph boundaries.'})
    if scope == 'generic':
        n = pkg.default
        checks = [(pkg.prop(n,'w:rPr/w:sz').get('val') == '24', 'normal-font-size', 'Default prose size differs from 12 pt.'),
                  (pkg.prop(n,'w:rPr/w:rFonts').get('eastAsia') in ('宋体','SimSun'), 'chinese-font','Default Chinese font differs from Songti/SimSun.'),
                  (all(pkg.prop(n,'w:rPr/w:rFonts').get(k) == 'Times New Roman' for k in ('ascii','hAnsi')), 'latin-font','Default Latin font differs from Times New Roman.'),
                  (pkg.prop(n,'w:pPr/w:spacing').get('line') == '360' and pkg.prop(n,'w:pPr/w:spacing').get('lineRule') == 'auto','line-spacing','Default prose line spacing differs from 1.5 lines.'),
                  (pkg.prop(n,'w:pPr/w:jc').get('val') == 'both','alignment','Default prose is not justified.')]
        for passed, code, message in checks:
            if not passed:
                warnings.append({'code':code,'location':'style:' + n,'message':message})
        for i, sect in enumerate(pkg.document.findall('.//w:sectPr', NS), 1):
            mar = attrs(sect.find('w:pgMar', NS)); size = attrs(sect.find('w:pgSz', NS))
            if any(abs(int(mar.get(k,0)) - v) > 2 for k,v in PROFILE['page']['margins_twips'].items()):
                warnings.append({'code':'page-margins','location':f'section:{i}', 'message':'Margins differ from the generic 2.5 cm profile; may be intentional.'})
            if any(abs(int(size.get(k,0))-v) > 2 for k,v in zip(('w','h'),PROFILE['page']['size_twips'])):
                warnings.append({'code':'page-size','location':f'section:{i}', 'message':'Page size differs from portrait A4; may be intentional.'})
        for sid, expected in PROFILE['headings'].items():
            if sid in pkg.style_map and pkg.prop(sid,'w:rPr/w:sz').get('val') != str(expected['size_half_points']):
                warnings.append({'code':'heading-size','location':'style:' + sid, 'message':'Heading size differs from generic v1.1; do not silently reset an official template.'})
    try:
        pkg.assert_editable()
    except DocumentError as exc:
        warnings.append({'code':'edit-blocked','location':'package','message':str(exc)})
    external = 0
    for name,b in pkg.data.items():
        if name.endswith('.rels'):
            external += sum(e.get('TargetMode') == 'External' for e in parse_xml(b))
    if external:
        warnings.append({'code':'external-relationships','location':'package', 'message':f'{external} external relationships; not followed or fetched.'})
    result = {'schema_version': '2.0', 'file': Path(path).name, 'profile': PROFILE['id'],
            'scope': scope, 'sha256': sha256(Path(path).read_bytes()).hexdigest(),
            'status': 'fail' if errors else 'pass-with-warnings' if warnings else 'pass',
            'paragraph_roles': dict(counts), 'errors': errors, 'warnings': warnings,
            'review_required': bool(warnings),
            'checks': {'indent': 'fail' if errors else 'pass',
                'role_coverage': 'partial' if any(w['code'] == 'manual-review' for w in warnings) else 'recognized-roles',
                'font_availability': 'not-checked', 'official_compliance': 'not-evaluated'},
            'visual_review': 'not-performed-by-this-command',
            'coverage': 'Body style/indent and optional generic-profile checks; not full rendering or official compliance.'}
    if details:
        result['paragraphs'] = paragraph_details  # No source prose is copied into reports.
    return result


def freeze_indent(pkg: Package, sid: str, parent: Any) -> None:
    """Shield non-prose from a later change of a shared base style."""
    p = parent if parent.tag == q('p') else None
    old = pkg.prop(sid, 'w:pPr/w:ind', p)
    if 'firstLineChars' not in old:
        old['firstLineChars'] = '0'
    if 'firstLine' not in old:
        old['firstLine'] = '0'
    set_indent(parent, old, replace=True)


def fix_indent(src: Path, out: Path, extras: list[str] | None = None,
               strip_manual_padding: bool = False) -> dict[str, Any]:
    src, out = Path(src), Path(out)
    if out.suffix.lower() != src.suffix.lower():
        raise DocumentError('Repair must preserve .docx/.dotx type; use convert for a type change.')
    if src.resolve() == out.resolve() or out.exists():
        raise DocumentError('Repair requires a new output path, never the original/existing file.')
    pkg = Package(src); pkg.assert_editable(); bodies = pkg.body_ids(extras)
    # Capture classifications and old properties BEFORE changing any inheritance.
    all_paras = list(pkg.document.findall('.//w:body//w:p', NS))
    roles = [pkg.classify(p, bodies) for p in all_paras]
    needs_change = any(body_issues(pkg,s) for s in bodies & pkg.style_map.keys())
    needs_change |= any(body_issues(pkg,pkg.sid(p),p) for p,r in zip(all_paras,roles) if r == 'body')
    if strip_manual_padding:
        needs_change |= any(leading_padding_nodes(p) for p,r in zip(all_paras,roles) if r == 'body')
    if not needs_change:
        write_new(out, src.read_bytes())
        return {'changed':False,'body_paragraphs':roles.count('body'),'output':out.name}
    changes: dict[str,bytes] = {}
    # Freeze descendant special styles to protect future special paragraphs as well.
    snapshots = [(sid, pkg.prop(sid,'w:pPr/w:ind')) for sid in pkg.style_map if sid not in bodies
                 and any(st.get(q('styleId')) in bodies for st in pkg.chain(sid))]
    for sid, ind in snapshots:
        if 'firstLineChars' not in ind: ind['firstLineChars']='0'
        if 'firstLine' not in ind: ind['firstLine']='0'
        set_indent(pkg.style_map[sid], ind, replace=True)
    for p, role in zip(all_paras,roles):
        if role not in ('body', 'empty-body') and pkg.sid(p) in bodies:
            freeze_indent(pkg,pkg.sid(p),p)
    # Header/footer and notes are not body prose, even when styled Normal.
    other_xml = {}
    for name,b in pkg.data.items():
        if re.fullmatch(r'word/(?:header\d*|footer\d*|footnotes|endnotes|comments\d*)\.xml',name):
            tree = parse_xml(b)
            for p in tree.findall('.//w:p',NS):
                if pkg.sid(p) in bodies:
                    freeze_indent(pkg,pkg.sid(p),p)
            other_xml[name]=tree
    for sid in sorted(bodies & pkg.style_map.keys()):
        set_indent(pkg.style_map[sid], {'firstLineChars':'200','firstLine':str(expected_fallback(pkg,sid))})
    removed_padding = 0
    for p,role in zip(all_paras,roles):
        if role != 'body': continue
        sid = pkg.sid(p)
        set_indent(p, {'firstLineChars':'200','firstLine':str(expected_fallback(pkg,sid,p))})
        if strip_manual_padding:
            removed_padding += strip_leading_padding(p)
    changes['word/styles.xml']=xml_bytes(pkg.styles)
    changes['word/document.xml']=xml_bytes(pkg.document)
    changes.update({name:xml_bytes(tree) for name,tree in other_xml.items()})
    write_new(out,pkg.serialized(changes))
    return {'changed':True,'body_paragraphs':roles.count('body'), 'manual_padding_removed':removed_padding,
            'output':out.name,'changed_parts':sorted(k for k,v in changes.items() if v != pkg.data[k]),
            'non_body_policy':'preserve existing indentation; unclassified paragraphs require review'}


def _style_ids(pkg: Package, labels: tuple[str,...] | list[str]) -> set[str]:
    allowed={canon(value) for value in labels}
    return {sid for sid in pkg.style_map if pkg.keys(sid) & allowed}


def format_headings(src: Path, out: Path, *, scheme: str = 'chinese-formal',
                    convert_manual: bool = False) -> dict[str,Any]:
    """Apply native multilevel numbering to unnumbered Heading1-3 only.

    Never deletes manual prefixes, overwrites existing list ids, or modifies
    other paragraph styles. Special front-/back-matter headings stay unnumbered.
    """
    src,out=Path(src),Path(out)
    if scheme not in SCHEMES: raise DocumentError('Choose an explicit heading numbering scheme.')
    if src.suffix.lower()!=out.suffix.lower():
        raise DocumentError('Heading formatting must preserve DOCX/DOTX extension.')
    if src.resolve()==out.resolve() or out.exists():
        raise DocumentError('Use a new output path; source is never overwritten.')
    pkg=Package(src);pkg.assert_editable()
    paras=pkg.document.findall('./w:body/w:p',NS)
    report={'formatted':0,'already':0,'converted_manual':0,'skipped_special':0,'skipped_manual':0,
            'skipped_existing_numbering':0,'heading_paragraphs':0,'missing_numbering_part':False}
    candidates=[]
    for para in paras:
        sid=pkg.sid(para)
        lvl=HEADER_STYLES.get(sid)
        if lvl is None:continue
        report['heading_paragraphs']+=1
        txt=pkg.text(para)
        num_id,ilvl=get_num(para)
        if is_special_heading(txt):
            report['skipped_special']+=1
            # Only add explicit suppression once there is a profile to protect from.
            candidates.append(('suppress',para,lvl))
        elif has_manual_number(txt):
            if convert_manual and num_id in (None,'0'):
                candidates.append(('convert',para,lvl))
            else:
                report['skipped_manual']+=1
                candidates.append(('suppress' if num_id in (None,'0') else 'keep',para,lvl))
        elif num_id not in (None,'0'):
            # Do not hijack existing numbering, including an earlier CPDOCX scheme.
            report['skipped_existing_numbering']+=1
        else:
            candidates.append(('number',para,lvl))
    selected=[(p,lvl) for action,p,lvl in candidates if action in ('number','convert')]
    if selected and 'word/numbering.xml' not in pkg.data:
        raise DocumentError('No numbering.xml in this DOCX; use a compatible Word save-as copy first.')
    if not selected:
        # No new numbering. Preserve full original ZIP byte-for-byte.
        write_new(out,src.read_bytes())
        return {**report,'scheme':scheme,'output':out.name,'changed':False,'changed_parts':[]}
    numbering=parse_xml(pkg.data['word/numbering.xml'])
    num_id=ensure_multilevel_numbering(numbering,scheme)
    for action,para,lvl in candidates:
        if action in ('number','convert'):
            if action=='convert':
                if remove_manual_prefix_if_safe(para):
                    report['converted_manual']+=1
                else:
                    report['skipped_manual']+=1
                    suppress_number(para)
                    continue
            if numbered_paragraph(para,lvl,num_id):report['formatted']+=1
            else:report['already']+=1
        elif action=='suppress':
            suppress_number(para)
    changes={}
    for part,tree in [('word/numbering.xml',numbering),('word/document.xml',pkg.document)]:
        data=xml_bytes(tree)
        if data!=pkg.data[part]:changes[part]=data
    write_new(out,pkg.serialized(changes))
    return {**report,'scheme':scheme,'output':out.name,'changed':bool(changes),
            'changed_parts':sorted(changes)}


def audit_headings(src: Path, *, expected_scheme: str | None = None) -> dict[str,Any]:
    """Read-only structural heading QA; never guesses a school's compliance."""
    pkg=Package(src)
    numbering=parse_xml(pkg.data['word/numbering.xml']) if 'word/numbering.xml' in pkg.data else None
    entries,jumps=analyze_heading_paragraphs(pkg.document,numbering)
    counts=Counter(e['level'] for e in entries)
    schemes=Counter(e['numbering'] for e in entries if e['numbering'])
    active=[e for e in entries if not e['special']]
    warnings=[]
    if jumps:warnings.append({'code':'heading_hierarchy_jump','count':jumps})
    manual=sum(e['manual_number'] for e in active)
    if manual:warnings.append({'code':'manual_heading_numbering','count':manual})
    unnumbered=sum(e['num_id'] in (None,'0') and not e['manual_number'] for e in active)
    if expected_scheme and expected_scheme != 'none':
        mismatched=sum(e['numbering'] != expected_scheme and not e['manual_number'] for e in active)
        if mismatched:warnings.append({'code':'numbering_profile_mismatch','count':mismatched})
    return {'file':Path(src).name,'heading_paragraphs':len(entries),
            'counts':{f'level{i}':counts.get(i,0) for i in (1,2,3)},
            'manual_prefixes':manual,'unnumbered':unnumbered,'hierarchy_jumps':jumps,
            'profile_counts':dict(schemes),'expected_scheme':expected_scheme,
            'warnings':warnings,'review_required':bool(warnings),
            'note':'Unknown custom heading styles and template-specific rules require human review.'}


def format_code(src: Path, out: Path, style_names: list[str] | None = None) -> dict[str,Any]:
    """Conservatively restyle explicitly classified code paragraphs and code styles."""
    src,out=Path(src),Path(out)
    if src.suffix.lower()!=out.suffix.lower():
        raise DocumentError('Code formatting must preserve DOCX/DOTX extension.')
    if src.resolve()==out.resolve() or out.exists():
        raise DocumentError('Use a new output path; source is never overwritten.')
    pkg=Package(src);pkg.assert_editable()
    names=list(CODE_STYLE_NAMES)+(style_names or [])
    matched=_style_ids(pkg,names)
    if style_names and any(not _style_ids(pkg,[name]) for name in style_names):
        raise DocumentError('Some explicitly selected code styles do not exist in the document.')
    for sid in matched: set_code_style(pkg.style_map[sid])
    paras=0
    for p in pkg.document.findall('.//w:p',NS):
        if pkg.sid(p) in matched:
            set_code_paragraph(p)
            paras+=1
    changed={}
    if matched:
        for name,tree in [('word/styles.xml',pkg.styles),('word/document.xml',pkg.document)]:
            b=xml_bytes(tree)
            if b!=pkg.data[name]:changed[name]=b
    write_new(out,pkg.serialized(changed))
    return {'output':out.name,'changed':bool(changed),'code_styles':len(matched),
            'code_paragraphs':paras,'changed_parts':sorted(changed)}


def format_tables(src: Path, out: Path, *, all_tables: bool = False,
                  variant: str = 'three-line') -> dict[str,Any]:
    """Format only explicitly tagged tables unless --all is requested."""
    src,out=Path(src),Path(out)
    if variant not in TABLE_VARIANTS:raise DocumentError('Unknown table variant.')
    if src.suffix.lower()!=out.suffix.lower():
        raise DocumentError('Table formatting must preserve DOCX/DOTX extension.')
    if src.resolve()==out.resolve() or out.exists():
        raise DocumentError('Use a new output path; source is never overwritten.')
    pkg=Package(src);pkg.assert_editable()
    touched=0
    for tbl in pkg.document.findall('.//w:tbl',NS):
        # An outer table and its nested table are formatted independently.
        tag=tbl.find('w:tblPr/w:tblStyle',NS)
        if all_tables or (tag is not None and tag.get(q('val'))==TABLE_STYLE_ID):
            set_academic_table(tbl,variant)
            touched+=1
    changed={}
    if touched:
        create_academic_table_style(pkg.styles)
        for name,tree in [('word/styles.xml',pkg.styles),('word/document.xml',pkg.document)]:
            b=xml_bytes(tree)
            if b!=pkg.data[name]:changed[name]=b
    write_new(out,pkg.serialized(changed))
    return {'output':out.name,'changed':bool(changed),'formatted_tables':touched,
            'selection':'all explicitly authorized' if all_tables else 'tagged academic tables only',
            'style':variant,'changed_parts':sorted(changed)}


def convert(src: Path, out: Path) -> None:
    src,out = Path(src),Path(out)
    if out.suffix.lower() not in ('.docx','.dotx'):
        raise DocumentError('Conversion output must be .docx or .dotx.')
    if src.resolve() == out.resolve() or out.exists():
        raise DocumentError('Conversion requires a new output path.')
    pkg=Package(src); pkg.assert_editable()
    if src.suffix.lower() == out.suffix.lower():
        write_new(out,src.read_bytes()); return
    ct=parse_xml(pkg.data['[Content_Types].xml'])
    for el in ct:
        if el.get('PartName') == '/word/document.xml':
            el.set('ContentType', DOTX_CT if out.suffix.lower()=='.dotx' else DOCX_CT)
    write_new(out,pkg.serialized({'[Content_Types].xml':xml_bytes(ct)}))


def validate_brief(obj: Any) -> dict[str,Any]:
    if not isinstance(obj,dict) or not isinstance(obj.get('title'),str) or not obj['title'].strip():
        raise DocumentError('Brief must be a JSON object with a nonempty string title.')
    if set(obj)-{'title','subtitle','info','blocks','document_type','numbering','bibliography','citation_mode','reference_standard','caption_numbering','caption_separator'}:
        raise DocumentError('Unknown top-level brief keys: ' + ', '.join(sorted(set(obj)-{'title','subtitle','info','blocks','document_type','numbering','bibliography','citation_mode','reference_standard','caption_numbering','caption_separator'})))
    for key in ('subtitle','info'):
        if key in obj and not isinstance(obj[key],str): raise DocumentError(key+' must be text.')
    doc_type=obj.get('document_type','general')
    numbering=obj.get('numbering','auto')
    if doc_type not in DOCUMENT_TYPES: raise DocumentError('Unsupported document_type: '+str(doc_type))
    if numbering not in NUMBERING_CHOICES: raise DocumentError('Unsupported numbering: '+str(numbering))
    citation_mode=obj.get('citation_mode','numeric-superscript')
    if citation_mode not in CITATION_MODES:
        raise DocumentError('citation_mode must be numeric-superscript, numeric-inline or author-year.')
    standard=obj.get('reference_standard',DEFAULT_REFERENCE_STANDARD)
    if standard not in REFERENCE_STANDARDS:
        raise DocumentError('Unsupported reference_standard: '+str(standard))
    caption_mode=obj.get('caption_numbering','auto')
    if caption_mode not in ('auto','global','chapter'):
        raise DocumentError('caption_numbering must be auto/global/chapter.')
    if caption_mode=='chapter' and select_scheme(doc_type,numbering) not in ('academic-chapter','academic-decimal','humanities'):
        raise DocumentError('Chapter captions require Arabic-numbered Heading 1; use academic-decimal/chapter or global captions.')
    if obj.get('caption_separator','-') not in ('.','-'):
        raise DocumentError('caption_separator must be . or -.')
    refs=obj.get('bibliography',[])
    if not isinstance(refs,list) or len(refs)>2000:
        raise DocumentError('bibliography must be a list of at most 2000 sources.')
    biblio_audit=audit_bibliography(refs,standard=standard,citation_mode=citation_mode)
    if biblio_audit['errors']:
        examples=[{'id':x.get('id'),'errors':x['errors']} for x in biblio_audit['entries'] if x['errors']]
        raise DocumentError('Invalid bibliography metadata: '+json.dumps(examples[:5],ensure_ascii=False))
    seen_sources={entry['id'] for entry in refs}
    blocks=obj.get('blocks',[])
    if not isinstance(blocks,list) or len(blocks)>5000: raise DocumentError('blocks must be an array of at most 5000 objects.')
    allowed={'paragraph','heading','bullets','numbered','note','code','reference','table','pagebreak','toc','figure','equation','list-of-figures','list-of-tables'}
    for i,b in enumerate(blocks):
        if not isinstance(b,dict) or b.get('type') not in allowed: raise DocumentError(f'Unsupported block at index {i}.')
        t=b['type']
        keys={'type'}
        if t in {'paragraph','heading','note','code','reference'}:
            keys.add('text')
            if not isinstance(b.get('text'),str): raise DocumentError(f'block {i}: text must be a string.')
            if t=='heading':
                keys.add('level');keys.add('numbered')
                if 'numbered' in b and type(b['numbered']) is not bool: raise DocumentError(f'block {i}: numbered must be boolean.')
                if type(b.get('level')) is not int or b['level'] not in (1,2,3): raise DocumentError(f'block {i}: heading level must be 1, 2 or 3.')
            if t=='paragraph' and any(line[:1] in (' ','\u3000','\t') for line in b['text'].splitlines()):
                raise DocumentError(f'block {i}: remove manual leading padding; the paragraph style handles indentation.')
        elif t in {'bullets','numbered'}:
            keys.add('items')
            if not isinstance(b.get('items'),list) or not b['items'] or not all(isinstance(x,str) for x in b['items']):
                raise DocumentError(f'block {i}: items must be a nonempty array of strings.')
        elif t=='toc':
            pass
        elif t=='table':
            keys|={'headers','rows','caption','style','id','continuation','note'}
            h,rows=b.get('headers'),b.get('rows')
            if not isinstance(h,list) or not h or len(h)>20 or not all(isinstance(x,str) for x in h):
                raise DocumentError(f'block {i}: headers must contain 1-20 strings.')
            if not isinstance(rows,list) or not all(isinstance(r,list) and len(r)==len(h) and all(isinstance(c,str) for c in r) for r in rows):
                raise DocumentError(f'block {i}: every table row must match the header width and contain strings.')
            if 'caption' in b and not isinstance(b['caption'],str): raise DocumentError(f'block {i}: caption must be text.')
            if 'id' in b and (not isinstance(b['id'],str) or not KEY_RE.fullmatch(b['id']) or not b.get('caption')):
                raise DocumentError(f'block {i}: table id must be ASCII and needs a caption.')
            if 'style' in b and b['style'] not in TABLE_VARIANTS: raise DocumentError(f'block {i}: table style must be three-line, grid or none.')
            if 'note' in b and not isinstance(b['note'],str):
                raise DocumentError(f'block {i}: note must be text.')
            cont=b.get('continuation',{'mode':'repeat-header'})
            if not isinstance(cont,dict) or set(cont)-{'mode','rows_per_part','page_break_before'}:
                raise DocumentError(f'block {i}: invalid continuation configuration.')
            mode=cont.get('mode','repeat-header')
            if mode not in ('repeat-header','explicit-chunks'):
                raise DocumentError(f'block {i}: invalid continuation mode.')
            if mode=='repeat-header' and set(cont)-{'mode'}:
                raise DocumentError(f'block {i}: repeat-header mode does not accept chunk options.')
            if mode=='explicit-chunks':
                per=cont.get('rows_per_part')
                if not isinstance(per,int) or isinstance(per,bool) or per<1 or per>=len(rows):
                    raise DocumentError(f'block {i}: rows_per_part must be between 1 and len(rows)-1.')
                if not b.get('caption') or not b.get('id'):
                    raise DocumentError(f'block {i}: continuation needs caption and stable table id.')
                if 'page_break_before' in cont and type(cont['page_break_before']) is not bool:
                    raise DocumentError(f'block {i}: page_break_before must be boolean.')
        elif t=='figure':
            keys|={'id','caption','image_path','placeholder'}
            if (not isinstance(b.get('id'),str) or not KEY_RE.fullmatch(b['id']) or
                not isinstance(b.get('caption'),str) or not b['caption'].strip()):
                raise DocumentError(f'block {i}: figure requires ASCII id and nonblank caption.')
            if ('image_path' in b)==('placeholder' in b):
                raise DocumentError(f'block {i}: figure requires exactly one of image_path/placeholder.')
            if 'image_path' in b and (not isinstance(b['image_path'],str) or not b['image_path'].strip()):
                raise DocumentError(f'block {i}: image_path must be nonblank.')
            if 'placeholder' in b and not isinstance(b['placeholder'],str):
                raise DocumentError(f'block {i}: placeholder must be text.')
        elif t=='equation':
            keys|={'id','expression'}
            if (not isinstance(b.get('id'),str) or not KEY_RE.fullmatch(b['id']) or
                not isinstance(b.get('expression'),str) or not b['expression'].strip()):
                raise DocumentError(f'block {i}: equation requires ASCII id and nonblank plaintext expression.')
        if set(b)-keys: raise DocumentError(f'block {i}: unsupported keys {sorted(set(b)-keys)}.')
    # Reject duplicate keys per target type and references to missing targets.
    ids={'figure':set(),'table':set(),'equation':set()}
    table_seq=0
    for b in blocks:
        typ=b['type']
        if typ=='table' and b.get('caption'):
            table_seq+=1
            key=b.get('id',f'autoTable{table_seq}')
        elif typ in ('figure','equation'):
            key=b['id']
        else:continue
        if key in ids[typ]:raise DocumentError(f'Duplicate {typ} id: {key}')
        ids[typ].add(key)
    for idx,b in enumerate(blocks):
        if b['type'] not in ('paragraph','note'):continue
        for m in ANY_TOKEN_RE.finditer(b['text']):
            match=TOKEN_RE.fullmatch(m.group())
            if not match:raise DocumentError(f'Block {idx}: invalid reference marker {m.group()[:80]}')
            kind,keys=match.groups()
            if kind=='CITE':
                try:
                    parsed_keys,_=_parse_cluster(keys)
                except ReferenceError as exc:
                    raise DocumentError(f'Block {idx}: '+str(exc)) from exc
                for key in parsed_keys:
                    if key not in seen_sources:raise DocumentError(f'Block {idx}: unknown bibliography id {key}')
            elif kind in ('FIG','TABLE','EQ'):
                typ={'FIG':'figure','TABLE':'table','EQ':'equation'}[kind]
                if keys not in ids[typ]:raise DocumentError(f'Block {idx}: unknown {typ} id {keys}')
            elif kind=='PAGE' and sum(keys in seq for seq in ids.values())!=1:
                raise DocumentError(f'Block {idx}: PAGE id must match one target: {keys}')
    return obj


def add_seq(p: Any, label: str, number: int) -> None:
    from docx.oxml import OxmlElement
    p.add_run('表 ')
    run=p.add_run()
    begin=OxmlElement('w:fldChar'); begin.set(q('fldCharType'),'begin');run._r.append(begin)
    instr=OxmlElement('w:instrText');instr.set('{'+XML+'}space','preserve');instr.text=' SEQ '+label+' \\* ARABIC '
    p.add_run()._r.append(instr)
    sep=OxmlElement('w:fldChar');sep.set(q('fldCharType'),'separate');p.add_run()._r.append(sep)
    p.add_run(str(number))
    end=OxmlElement('w:fldChar');end.set(q('fldCharType'),'end');p.add_run()._r.append(end)
    p.add_run('  ')


def new_document(out: Path, scene: str='generic', brief: dict[str,Any] | None=None,
                 *, brief_dir: Path | None = None, manifest_out: Path | None = None) -> dict[str, Any] | None:
    out=Path(out)
    if out.suffix.lower() not in ('.docx','.dotx'): raise DocumentError('Output must be .docx or .dotx.')
    if out.exists(): raise DocumentError('Output already exists: '+str(out))
    if scene not in SCENES: raise DocumentError('Unknown scene: '+scene)
    if brief is None:
        convert(ROOT/'assets'/SCENES[scene],out); return
    if scene!='generic': raise DocumentError('Use either --brief or a non-generic --scene, not both.')
    brief=validate_brief(brief)
    scheme=select_scheme(brief.get('document_type','general'),brief.get('numbering','auto'))
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.shared import Cm, Pt
    doc=Document(ROOT/'assets/templates/base.docx')
    # Distinct caption paragraph styles allow field-backed lists of figures /
    # tables via TOC \t. A TOC \c over the SEQ identifier would also pick up
    # hidden chapter-reset SEQ fields in some Word implementations.
    from docx.enum.style import WD_STYLE_TYPE
    for caption_style in ('CPFigureCaption','CPTableCaption'):
        if caption_style not in doc.styles:
            st=doc.styles.add_style(caption_style,WD_STYLE_TYPE.PARAGRAPH)
            st.base_style=doc.styles['Caption']
    heading_num_id=(ensure_multilevel_numbering(doc.part.numbering_part.element,scheme)
                    if scheme!='none' else None)
    body=doc._element.body
    for child in list(body):
        if child.tag!=q('sectPr'): body.remove(child)
    doc.add_paragraph(brief['title'],style='Title')
    for key,style in [('subtitle','Subtitle'),('info','文档信息')]:
        if brief.get(key): doc.add_paragraph(brief[key],style=style)
    registry=Registry(doc._element)
    standard=brief.get('reference_standard',DEFAULT_REFERENCE_STANDARD)
    citation_mode=brief.get('citation_mode','numeric-superscript')
    bibliography=brief.get('bibliography',[])
    metadata_report=audit_bibliography(bibliography,standard=standard,citation_mode=citation_mode)
    suffix_years=assign_year_suffix(bibliography)
    requested=plan_citations(brief.get('blocks',[]))
    bib_entries={entry['id']:entry for entry in bibliography}
    if citation_mode=='author-year':
        bib_order=sorted(bib_entries,key=lambda k:(
            (bib_entries[k].get('authors') or [''])[0].casefold(),
            suffix_years.get(k,''),bib_entries[k].get('title','').casefold(),k))
    else:
        bib_order=requested+[key for key in bib_entries if key not in requested]
    caption_mode=brief.get('caption_numbering','auto')
    chapter_scope=(scheme in ('academic-chapter','academic-decimal','humanities') if caption_mode=='auto'
                   else caption_mode=='chapter')
    chapter=0
    caption_separator=brief.get('caption_separator','-')
    count={'figure':0,'table':0,'equation':0}
    table_count=0
    previous_kind=None
    for block in brief.get('blocks',[]):
        kind=block['type']
        after_table=(previous_kind=='table')
        if kind=='heading':
            text=block['text']
            automatic=(scheme!='none' and block.get('numbered',True)
                       and not is_special_heading(text))
            if automatic and has_manual_number(text):
                raise DocumentError('Heading has a manually written number; use clean title text, numbered=false or numbering=none.')
            para=doc.add_heading(text,level=block['level'])
            if automatic and block['level']==1:
                chapter+=1
                if chapter_scope:count={'figure':0,'table':0,'equation':0}
                if chapter_scope:
                    # Native Heading1 numbering is used visually. A hidden SEQ
                    # paragraph field supplies a numeric, updateable chapter id
                    # to captions; STYLEREF 1 \s is inconsistent across Word/LO.
                    hidden=OxmlElement('w:rPr')
                    hidden.append(OxmlElement('w:vanish'))
                    from cross_references import field_runs
                    para._p.extend(field_runs(r'SEQ CPChapter \* ARABIC',str(chapter),rpr=hidden))
                    # Reset all chapter-scoped counters on the heading itself.
                    # This survives SEQ reordering and avoids a renderer-specific
                    # failure of SEQ ... \s 1 in some LibreOffice versions.
                    from cross_references import SEQ_NAME
                    for category in ('figure','table','equation'):
                        para._p.extend(field_runs('SEQ '+SEQ_NAME[category]+r' \r 0','0',rpr=hidden))
            if heading_num_id is not None:
                if automatic:
                    numbered_paragraph(para._p,block['level'],heading_num_id)
                else:
                    suppress_number(para._p)
        elif kind in ('paragraph','note','code','reference'):
            style={'paragraph':'Normal','note':'重点提示','code':'代码段','reference':'参考资料'}[kind]
            if kind=='paragraph':
                # Each newline-separated prose chunk becomes a real Word paragraph.
                for text in re.split(r'\r?\n',block['text']):
                    if text.strip(): doc.add_paragraph(text,style=style)
            else:
                p=doc.add_paragraph(block['text'],style=style)
                if kind=='code':set_code_paragraph(p._p)
        elif kind=='toc':
            title=doc.add_paragraph('目  录')
            title.style='Title'
            title.paragraph_format.keep_with_next=True
            field=doc.add_paragraph()
            begin=OxmlElement('w:fldChar');begin.set(q('fldCharType'),'begin')
            field.add_run()._r.append(begin)
            instr=OxmlElement('w:instrText');instr.text=r' TOC \o "1-3" \h \z \u '
            instr.set('{'+XML+'}space','preserve')
            field.add_run()._r.append(instr)
            sep=OxmlElement('w:fldChar');sep.set(q('fldCharType'),'separate')
            field.add_run()._r.append(sep)
            field.add_run('（在 Word 中右键更新目录域）')
            end=OxmlElement('w:fldChar');end.set(q('fldCharType'),'end')
            field.add_run()._r.append(end)
        elif kind in ('list-of-figures','list-of-tables'):
            label=('插图目录' if kind=='list-of-figures' else '附表目录')
            caption_style=('CPFigureCaption' if kind=='list-of-figures' else 'CPTableCaption')
            heading=doc.add_paragraph(label,style='Title')
            heading.paragraph_format.keep_with_next=True
            field=doc.add_paragraph()
            begin=OxmlElement('w:fldChar');begin.set(q('fldCharType'),'begin')
            field.add_run()._r.append(begin)
            instr=OxmlElement('w:instrText')
            instr.text=r' TOC \t "'+caption_style+r',1" \h \z '
            instr.set('{'+XML+'}space','preserve')
            field.add_run()._r.append(instr)
            sep=OxmlElement('w:fldChar');sep.set(q('fldCharType'),'separate')
            field.add_run()._r.append(sep)
            field.add_run('（打开 Word/WPS 后更新该目录域）')
            end=OxmlElement('w:fldChar');end.set(q('fldCharType'),'end')
            field.add_run()._r.append(end)
        elif kind in ('bullets','numbered'):
            style='List Bullet' if kind=='bullets' else 'List Number'
            new_num_id=None
            if kind=='numbered':
                numroot=doc.part.numbering_part.element
                base_id=doc.styles[style]._element.find('w:pPr/w:numPr/w:numId',NS).get(q('val'))
                base_num=numroot.find('w:num[@w:numId="'+base_id+'"]',NS)
                abstract=base_num.find('w:abstractNumId',NS).get(q('val'))
                new_num_id=max([int(e.get(q('numId'))) for e in numroot.findall('w:num',NS)]+[0])+1
                num=OxmlElement('w:num');num.set(q('numId'),str(new_num_id))
                an=OxmlElement('w:abstractNumId');an.set(q('val'),abstract);num.append(an)
                override=OxmlElement('w:lvlOverride');override.set(q('ilvl'),'0')
                start=OxmlElement('w:startOverride');start.set(q('val'),'1');override.append(start);num.append(override);numroot.append(num)
            for text in block['items']:
                p=doc.add_paragraph(text,style=style)
                if new_num_id is not None:
                    np=p._p.get_or_add_pPr().get_or_add_numPr();np.get_or_add_ilvl().val=0;np.get_or_add_numId().val=new_num_id
        elif kind=='table':
            from long_tables import table_segments, build_editable_table, continuation_caption, add_table_note
            table_count+=1
            tab_id=block.get('id',f'autoTable{table_count}')
            segments=table_segments(block['rows'],block.get('continuation'))
            use_chapter=chapter if chapter_scope else None
            if chapter_scope and chapter<=0:
                raise DocumentError('Chapter-scoped table must follow a numbered first-level heading.')
            count['table']+=1
            target=None
            for part_no,rows in enumerate(segments,1):
                if part_no>1:
                    if block.get('continuation',{}).get('page_break_before',True):
                        doc.add_page_break()
                    cap=doc.add_paragraph(style='CPTableCaption')
                    cap.paragraph_format.keep_with_next=True
                    continuation_caption(cap,target,block['caption'])
                elif block.get('caption'):
                    cap=doc.add_paragraph(style='CPTableCaption')
                    cap.paragraph_format.keep_with_next=True
                    cap.add_run('表')
                    target=registry.add_number(cap,'table',tab_id,count['table'],
                         chapter=use_chapter,sep=caption_separator)
                    cap.add_run('  '+block['caption'])
                table=build_editable_table(doc,block['headers'],rows,
                    variant=block.get('style','three-line'),
                    border_fn=set_academic_table,table_id=tab_id,part_number=part_no)
            if block.get('note'):
                add_table_note(doc,block['note'])
        elif kind=='figure':
            if 'image_path' in block:
                image=Path(block['image_path'])
                if not image.is_absolute():image=(brief_dir or Path.cwd())/image
                if image.suffix.lower() not in ('.png','.jpg','.jpeg','.bmp','.tiff') or not image.is_file():
                    raise DocumentError('Figure image_path must reference an existing PNG/JPEG/BMP/TIFF file.')
                if image.stat().st_size>20*1024*1024:
                    raise DocumentError('Figure image exceeds 20 MiB input limit.')
                image_p=doc.add_paragraph()
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                image_p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                image_p.paragraph_format.keep_with_next=True
                printable=doc.sections[0].page_width-doc.sections[0].left_margin-doc.sections[0].right_margin
                image_p.add_run().add_picture(str(image),width=int(printable*0.82))
            else:
                p=doc.add_paragraph('〔图片占位：'+block['placeholder']+'〕',style='图片占位')
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.keep_with_next=True
            cap=doc.add_paragraph(style='CPFigureCaption')
            if chapter_scope and chapter<=0:
                raise DocumentError('Chapter-scoped figure must follow a numbered first-level heading.')
            count['figure']+=1
            cap.add_run('图')
            registry.add_number(cap,'figure',block['id'],count['figure'],
                                chapter=chapter if chapter_scope else None,sep=caption_separator)
            cap.add_run('  '+block['caption'])
        elif kind=='equation':
            if chapter_scope and chapter<=0:
                raise DocumentError('Chapter-scoped equation must follow a numbered first-level heading.')
            from docx.enum.text import WD_TAB_ALIGNMENT
            p=doc.add_paragraph()
            printable=doc.sections[0].page_width-doc.sections[0].left_margin-doc.sections[0].right_margin
            p.paragraph_format.tab_stops.add_tab_stop(int(printable*0.50),WD_TAB_ALIGNMENT.CENTER)
            p.paragraph_format.tab_stops.add_tab_stop(int(printable*0.98),WD_TAB_ALIGNMENT.RIGHT)
            p.add_run('\t'+block['expression']+'\t（')
            count['equation']+=1
            registry.add_number(p,'equation',block['id'],count['equation'],
                                chapter=chapter if chapter_scope else None,sep=caption_separator)
            p.add_run('）')
        else:
            doc.add_page_break()
        previous_kind=kind
    if bib_entries:
        # Existing front-matter headings are not treated as numbered chapters.
        if not any(b['type']=='heading' and is_special_heading(b['text']) and
                   ''.join(b['text'].split())=='参考文献' for b in brief.get('blocks',[])):
            p=doc.add_heading('参考文献',level=1)
            suppress_number(p._p)
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        for index,key in enumerate(bib_order,1):
            p=doc.add_paragraph(style='参考资料')
            entry=bib_entries[key]
            if citation_mode=='author-year':
                registry.add_author_year(p,key,citation_label(entry,suffix_years[key]))
            else:
                registry.add_bibliography(p,key,index)
            p.add_run(format_entry(entry,standard=standard,citation_mode=citation_mode,
                                   disambiguated_year=suffix_years.get(key)))
    if registry.targets:
        # Rewrite only single-run explicit markers in prose; target bookmarks were
        # established first, so forward figure/table citations also resolve.
        try:
            replace_markers(list(doc._element.body.iter(q('p'))),registry.targets,
                            citation_mode=citation_mode)
        except ReferenceError as exc:
            raise DocumentError(str(exc)) from exc
    settings=doc.settings.element
    update=settings.find('w:updateFields',NS)
    if update is None:update=OxmlElement('w:updateFields');settings.append(update)
    update.set(q('val'),'true')
    doc.core_properties.keywords='CPDOCX;standard='+standard+';citation-mode='+citation_mode
    doc.core_properties.title=brief['title']
    doc.core_properties.author=''
    doc.core_properties.last_modified_by=''
    buf=BytesIO();doc.save(buf)
    b=buf.getvalue()
    if out.suffix.lower()=='.dotx':
        with ZipFile(BytesIO(b)) as z:
            buf=BytesIO()
            with ZipFile(buf,'w',ZIP_DEFLATED) as target:
                for info in z.infolist():
                    data=z.read(info)
                    if info.filename=='[Content_Types].xml':
                        ct=parse_xml(data)
                        for el in ct:
                            if el.get('PartName')=='/word/document.xml':el.set('ContentType',DOTX_CT)
                        data=xml_bytes(ct)
                    target.writestr(info,data)
            b=buf.getvalue()
    if manifest_out is not None and (manifest_out.exists() or manifest_out.resolve()==out.resolve()):
        raise DocumentError('Reference manifest output must be a new separate path.')
    write_new(out,b)
    manifest={'schema_version':'1.0','source_docx':out.name,
             'citation_mode':citation_mode,'reference_standard':standard,
             'bibliography_audit':metadata_report,
             'targets':[{'kind':t.kind,'id':t.key,'bookmark':t.name,
                         'cached':t.cached} for t in registry.targets.values()],
             'note':'Cached values are provisional. Update all fields in Word (Ctrl+A,F9) and review.'}
    if manifest_out is not None:
        write_new(manifest_out,(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n').encode('utf-8'))
    return manifest


def audit_references(src: Path) -> dict[str,Any]:
    pkg=Package(src)
    citation_mode='numeric-superscript'
    try:
        from docx import Document
        props=Document(src).core_properties.keywords or ''
        match=re.search(r'(?:^|;)citation-mode=([a-z-]+)',props)
        if match and match.group(1) in CITATION_MODES:
            citation_mode=match.group(1)
    except (OSError,KeyError,ValueError):
        pass
    result=audit_fields(pkg.document,citation_mode=citation_mode)
    return {'file':Path(src).name,'citation_mode':citation_mode,**result,
            'visual':'not evaluated',
            'official_compliance':'not evaluated; compare target school/style instructions'}


def link_references(src: Path, out: Path, manifest: Path,
                    *, citation_mode: str = 'auto') -> dict[str,Any]:
    """Bind explicit tokens in an existing DOCX to already bookmarked targets.

    A manifest is mandatory because bookmark hashes are intentionally
    non-reversible. Never fabricate a target by parsing plain text '图1-1'.
    """
    if citation_mode not in ('auto',*CITATION_MODES):
        raise DocumentError('Unsupported citation_mode.')
    src,out=Path(src),Path(out)
    if src.resolve()==out.resolve() or out.exists() or src.suffix.lower()!=out.suffix.lower():
        raise DocumentError('link-references requires a new, same-type output path.')
    pkg=Package(src);pkg.assert_editable()
    data=json.loads(Path(manifest).read_text(encoding='utf-8'))
    if not isinstance(data,dict) or not isinstance(data.get('targets'),list):
        raise DocumentError('Manifest requires a targets list from a generated document.')
    manifest_mode=data.get('citation_mode','numeric-superscript')
    if manifest_mode not in CITATION_MODES:
        raise DocumentError('Manifest has unsupported citation mode.')
    if citation_mode=='auto':citation_mode=manifest_mode
    elif citation_mode!=manifest_mode:
        raise DocumentError('Explicit citation mode conflicts with the target manifest.')
    present={el.get(q('name')) for el in pkg.document.iter(q('bookmarkStart'))}
    targets={}
    for entry in data['targets']:
        if (not isinstance(entry,dict) or not {'kind','id','bookmark','cached'}<=entry.keys()):
            raise DocumentError('Invalid target manifest entry.')
        kind,key,name,cache=entry['kind'],entry['id'],entry['bookmark'],str(entry['cached'])
        try:
            if name!=bookmark_name(kind,key):
                raise DocumentError('Manifest bookmark does not match declared kind/id.')
        except ReferenceError as exc:raise DocumentError(str(exc)) from exc
        if name not in present:raise DocumentError('Missing target bookmark: '+name)
        if (kind,key) in targets:raise DocumentError('Duplicate target in manifest.')
        targets[(kind,key)]=Target(kind,key,name,cache)
    # Never rewrite every paragraph from scratch. A run-spanning marker gets
    # rejected, allowing a human to reauthor the marker in a single run.
    try:
        report=replace_markers(list(pkg.document.iter(q('p'))),targets,citation_mode=citation_mode)
    except ReferenceError as exc:
        raise DocumentError(str(exc)) from exc
    updates={}
    if report['replaced']:
        body=xml_bytes(pkg.document)
        if body!=pkg.data['word/document.xml']:updates['word/document.xml']=body
        if 'word/settings.xml' in pkg.data:
            settings=parse_xml(pkg.data['word/settings.xml'])
            f=settings.find('w:updateFields',NS)
            if f is None:
                f=etree.SubElement(settings,q('updateFields'))
            f.set(q('val'),'true')
            updated=xml_bytes(settings)
            if updated!=pkg.data['word/settings.xml']:updates['word/settings.xml']=updated
        if 'docProps/core.xml' in pkg.data:
            core=parse_xml(pkg.data['docProps/core.xml'])
            cpns='{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}'
            keywords=core.find(cpns+'keywords')
            if keywords is None:
                keywords=etree.SubElement(core,cpns+'keywords')
            old=keywords.text or ''
            remainder=[x for x in old.split(';') if x and not x.startswith('citation-mode=')]
            remainder.append('citation-mode='+citation_mode)
            keywords.text=';'.join(remainder)
            updates['docProps/core.xml']=xml_bytes(core)
    write_new(out,pkg.serialized(updates))
    return {'file':src.name,'output':out.name,'changed':bool(updates),**report,
            'note':'Refresh all fields in desktop Word, then review their displayed values.'}


def _summary(result: dict) -> dict:
    """One-line-friendly audit conclusion shared by audit --summary and pipeline."""
    codes={}
    for w in result['warnings']: codes[w['code']]=codes.get(w['code'],0)+1
    return {'file':result.get('file'),'status':result['status'],'scope':result.get('scope'),
        'errors':len(result['errors']),'warnings':len(result['warnings']),'warning_codes':codes,
        'checks':result.get('checks'),'review_required':result.get('review_required')}


def main() -> int:
    parser=argparse.ArgumentParser(description=__doc__)
    subs=parser.add_subparsers(dest='cmd',required=True)
    p=subs.add_parser('new',help='Copy a scene template or build from a JSON brief.')
    p.add_argument('--out',type=Path,required=True);p.add_argument('--scene',choices=SCENES,default='generic');p.add_argument('--brief',type=Path)
    p.add_argument('--manifest-out',type=Path,help='Optional reference target manifest for later marker binding.')
    p=subs.add_parser('validate-brief',help='Validate JSON content before writing a document.')
    p.add_argument('input',type=Path)
    p=subs.add_parser('audit',help='Read-only audit; exit 1 for indent errors, optionally 3 for review.')
    p.add_argument('input',type=Path);p.add_argument('--body-style',action='append',default=[]);p.add_argument('--report',type=Path)
    p.add_argument('--scope',choices=['generic','indent'],default='generic')
    p.add_argument('--details',action='store_true',help='Include paragraph roles and style IDs without prose text.')
    p.add_argument('--strict-review',action='store_true',help='Exit 3 when no errors but warnings remain.')
    p.add_argument('--summary',action='store_true',help='Print one-line JSON conclusion instead of the full report (full JSON still goes to --report).')
    p=subs.add_parser('pipeline',help='Audit then optionally render page PNGs in one command.')
    p.add_argument('input',type=Path);p.add_argument('--body-style',action='append',default=[])
    p.add_argument('--scope',choices=['generic','indent'],default='generic');p.add_argument('--details',action='store_true')
    p.add_argument('--strict-review',action='store_true');p.add_argument('--summary',action='store_true')
    p.add_argument('--render-dir',type=Path,help='Optional empty/absent directory; triggers render_preview.py after audit.')
    p.add_argument('--dpi',type=int,default=130);p.add_argument('--backend',choices=['libreoffice','word'],default='libreoffice')
    p=subs.add_parser('fix-indent',help='Repair prose indentation only, preserving non-prose and contents.')
    p.add_argument('input',type=Path);p.add_argument('--out',type=Path,required=True);p.add_argument('--body-style',action='append',default=[])
    p.add_argument('--strip-manual-padding',action='store_true',help='Explicitly remove leading spaces/tabs from recognized plain prose runs.')
    p=subs.add_parser('format-headings',help='Safely number Heading 1/2/3 with native Word multilevel lists.')
    p.add_argument('input',type=Path);p.add_argument('--out',type=Path,required=True)
    p.add_argument('--scheme',choices=tuple(SCHEMES),default='chinese-formal')
    p.add_argument('--convert-manual',action='store_true',help='Opt-in removal of simple, explicitly prefixed heading text, preserving runs.')
    p=subs.add_parser('audit-headings',help='Read-only heading hierarchy and numbering audit.')
    p.add_argument('input',type=Path);p.add_argument('--scheme',choices=('none',*SCHEMES))
    p.add_argument('--strict-review',action='store_true')
    p=subs.add_parser('format-code',help='Apply neutral light-gray shading to explicit code styles only.')
    p.add_argument('input',type=Path);p.add_argument('--out',type=Path,required=True)
    p.add_argument('--style',action='append',default=[],help='Additional existing code paragraph style name/ID; repeatable.')
    p=subs.add_parser('format-tables',help='Repair tagged academic tables; --all is required for unrelated tables.')
    p.add_argument('input',type=Path);p.add_argument('--out',type=Path,required=True)
    p.add_argument('--all',action='store_true',help='Explicitly format ALL tables in the source.')
    p.add_argument('--style',choices=TABLE_VARIANTS,default='three-line')
    p=subs.add_parser('audit-long-tables',help='Read-only structural audit of table headers, continuation captions, and segment metadata.')
    p.add_argument('input',type=Path);p.add_argument('--strict-review',action='store_true')
    p=subs.add_parser('audit-bibliography',help='Read-only structured bibliography metadata audit for a JSON brief.')
    p.add_argument('input',type=Path)
    p.add_argument('--strict-review',action='store_true')
    p=subs.add_parser('audit-references',help='Read-only audit for dangling fields, bookmarks and citation superscripts.')
    p.add_argument('input',type=Path);p.add_argument('--strict-review',action='store_true')
    p=subs.add_parser('link-references',help='Link explicit markers to bookmarks in an existing DOCX, without guessing numbers.')
    p.add_argument('input',type=Path);p.add_argument('--manifest',type=Path,required=True)
    p.add_argument('--out',type=Path,required=True)
    p.add_argument('--citation-mode',choices=('auto',*CITATION_MODES),default='auto')
    p=subs.add_parser('convert',help='Convert DOCX↔DOTX by changing the real package content type.')
    p.add_argument('input',type=Path);p.add_argument('--out',type=Path,required=True)
    args=parser.parse_args()
    try:
        if args.cmd=='validate-brief':
            brief=validate_brief(json.loads(args.input.read_text(encoding='utf-8')))
            print(json.dumps({'status':'pass','blocks':len(brief.get('blocks',[]))}));return 0
        if args.cmd=='audit-long-tables':
            from long_tables import audit_long_tables
            pkg=Package(args.input)
            result=audit_long_tables(pkg.document)
            print(json.dumps(result,ensure_ascii=False,indent=2))
            return 1 if result['errors'] else 3 if args.strict_review and result['warnings'] else 0
        if args.cmd=='audit-bibliography':
            b=json.loads(args.input.read_text(encoding='utf-8'))
            result=audit_bibliography(b.get('bibliography',[]),
                standard=b.get('reference_standard',DEFAULT_REFERENCE_STANDARD),
                citation_mode=b.get('citation_mode','numeric-superscript'))
            print(json.dumps(result,ensure_ascii=False,indent=2))
            return 1 if result['errors'] else 3 if args.strict_review and result['warnings'] else 0
        if args.cmd=='audit-references':
            result=audit_references(args.input)
            print(json.dumps(result,ensure_ascii=False,indent=2))
            return 1 if result['errors'] else 3 if args.strict_review and result['warnings'] else 0
        if args.cmd=='audit-headings':
            result=audit_headings(args.input,expected_scheme=args.scheme)
            print(json.dumps(result,ensure_ascii=False,indent=2))
            return 3 if args.strict_review and result['review_required'] else 0
        if args.cmd=='audit':
            result=audit(args.input,args.body_style,scope=args.scope,details=args.details)
            payload=(json.dumps(result,ensure_ascii=False,indent=2)+'\n').encode('utf-8')
            if args.report: write_new(args.report,payload)
            if args.summary:
                print(json.dumps(_summary(result),ensure_ascii=False))
            else: print(payload.decode('utf-8'),end='')
            return 1 if result['errors'] else 3 if args.strict_review and result['review_required'] else 0
        if args.cmd=='pipeline':
            result=audit(args.input,args.body_style,scope=args.scope,details=args.details)
            render_result=None
            if args.render_dir:
                import os,subprocess
                rp=Path(__file__).resolve().parent/'render_preview.py'
                env=dict(os.environ,PYTHONIOENCODING='utf-8')
                proc=subprocess.run([sys.executable,str(rp),str(args.input),'--out-dir',str(args.render_dir),
                    '--dpi',str(args.dpi),'--backend',args.backend],capture_output=True,text=True,
                    encoding='utf-8',errors='replace',env=env)
                parsed=None
                if proc.returncode==0:
                    try:
                        parsed=json.loads(proc.stdout)
                        if not isinstance(parsed,dict): parsed=None
                    except ValueError: parsed=None
                if proc.returncode==0 and parsed is not None: render_result=parsed
                else: render_result={'exit':proc.returncode or 1,
                    'error':((proc.stderr or proc.stdout) or 'render produced no parseable JSON')[-1000:]}
            summary=_summary(result)
            if render_result is not None: summary['render']=render_result
            print(json.dumps(summary if args.summary else {'audit':result,'render':render_result},
                ensure_ascii=False,indent=None if args.summary else 2))
            if result['errors']: return 1
            if render_result is not None and render_result.get('exit'): return 2
            return 3 if args.strict_review and result['review_required'] else 0
        if args.cmd=='fix-indent':result=fix_indent(args.input,args.out,args.body_style,args.strip_manual_padding)
        elif args.cmd=='format-headings':result=format_headings(args.input,args.out,scheme=args.scheme,convert_manual=args.convert_manual)
        elif args.cmd=='format-code':result=format_code(args.input,args.out,args.style)
        elif args.cmd=='format-tables':result=format_tables(args.input,args.out,all_tables=args.all,variant=args.style)
        elif args.cmd=='link-references':result=link_references(args.input,args.out,args.manifest,citation_mode=args.citation_mode)
        elif args.cmd=='convert':convert(args.input,args.out);result={'output':str(args.out)}
        else:
            brief=json.loads(args.brief.read_text(encoding='utf-8')) if args.brief else None
            m=new_document(args.out,args.scene,brief,brief_dir=args.brief.parent if args.brief else None,
                           manifest_out=args.manifest_out);result={'output':str(args.out),
                           'reference_targets':len(m['targets']) if m else 0}
        print(json.dumps(result,ensure_ascii=False,indent=2));return 0
    except (DocumentError,ReferenceError,BibliographyError,OSError,ValueError,KeyError) as exc:
        print('ERROR: '+str(exc),file=sys.stderr);return 2


if __name__=='__main__':
    raise SystemExit(main())
