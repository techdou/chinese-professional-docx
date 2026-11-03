"""High-signal OOXML contract tests for editable citations and cross references."""
from __future__ import annotations
import json
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile
from lxml import etree
from docx import Document

ROOT=Path(__file__).resolve().parents[1]
if not (ROOT/'scripts').is_dir():
    ROOT=ROOT/'skills/chinese-professional-docx'
sys.path.insert(0,str(ROOT/'scripts'))
from docx_ops import (new_document, validate_brief, DocumentError, audit_references,
                      link_references, Package)
from cross_references import q, NS, BOOKMARK_PREFIX


def sample(*, scheme='thesis-chapter', sep='-'):
    return {'title':'自动引文与图表测试','document_type':scheme,
     'caption_numbering':'chapter','caption_separator':sep,
     'bibliography':[{'id':'srcA','text':'作者. 论文一[J]. 期刊, 2025.'},
                    {'id':'srcB','text':'作者. 论文二[J]. 期刊, 2026.'}],
     'blocks':[
      {'type':'heading','level':1,'text':'绪论'},
      {'type':'paragraph','text':'先引文献[[CITE:srcB;srcA]]，再引[[CITE:srcB]]，如[[FIG:network]]、[[TABLE:t1]]和[[EQ:loss]]，参见[[PAGE:network]]页。'},
      {'type':'figure','id':'network','caption':'结构示意','placeholder':'图像'},
      {'type':'table','id':'t1','caption':'比较','headers':['列名','数值'],'rows':[['甲','0.9']]},
      {'type':'equation','id':'loss','expression':'a+b=c'},
      {'type':'heading','level':1,'text':'实验'},
      {'type':'paragraph','text':'见[[FIG:network2]]与[[TABLE:t2]]、[[EQ:loss2]]。'},
      {'type':'figure','id':'network2','caption':'第二个结构图','placeholder':'图'},
      {'type':'table','id':'t2','caption':'第二组','headers':['甲','乙'],'rows':[['1','2']]},
      {'type':'equation','id':'loss2','expression':'x=1'},
     ]}


def doc_xml(path):
    with ZipFile(path) as z:
        return etree.fromstring(z.read('word/document.xml'))


def instructions(path):
    return [(e.text or '').strip() for e in doc_xml(path).iter(q('instrText'))]


class ReferenceFieldTests(unittest.TestCase):
    def test_two_chapters_native_seq_refs_and_superscript(self):
        with TemporaryDirectory() as td:
            tmp=Path(td);out=tmp/'paper.docx';manifest=tmp/'refs.json'
            doc_data=new_document(out,brief=sample(),manifest_out=manifest)
            self.assertEqual(len(doc_data['targets']),8)
            inst=instructions(out)
            self.assertEqual(sum(x.startswith('SEQ CPChapter \\*') for x in inst),2)
            for label in ('CPFigure','CPTable','CPEquation'):
                self.assertEqual(sum(x==f'SEQ {label} \\r 0' for x in inst),2)
                self.assertEqual(sum(x==f'SEQ {label} \\* ARABIC' for x in inst),2)
                self.assertFalse(any(f'SEQ {label} \\s' in x for x in inst))
            self.assertEqual(sum(x.startswith('REF _CP_BIB_') for x in inst),3)
            self.assertTrue(any(x.startswith('PAGEREF _CP_FIG_') for x in inst))
            refs=audit_references(out)
            self.assertEqual(refs['errors'],[])
            self.assertEqual(refs['warnings'],[])
            body=doc_xml(out)
            for run in body.findall('.//w:r',NS):
                field=run.find(q('instrText'))
                if field is not None and (field.text or '').strip().startswith('REF _CP_BIB_'):
                    supra=run.find('w:rPr/w:vertAlign',NS)
                    self.assertIsNotNone(supra)
                    self.assertEqual(supra.get(q('val')),'superscript')
            self.assertEqual([x['cached'] for x in doc_data['targets'] if x['kind']=='figure'],['1-1','2-1'])
            self.assertEqual([x['cached'] for x in doc_data['targets'] if x['kind']=='table'],['1-1','2-1'])
            self.assertEqual([x['cached'] for x in doc_data['targets'] if x['kind']=='equation'],['1-1','2-1'])
            self.assertEqual([x['id'] for x in doc_data['targets'] if x['kind']=='bibliography'],['srcB','srcA'])
            self.assertTrue(manifest.exists())
            self.assertIn('bookmark',manifest.read_text())

    def test_decimal_or_dash_separator_and_humanities(self):
        for scene in ('thesis-chapter','thesis-humanities'):
            for sep in ('.','-'):
                with self.subTest(scene=scene,sep=sep),TemporaryDirectory() as d:
                    doc=Path(d)/'r.docx';data=new_document(doc,brief=sample(scheme=scene,sep=sep))
                    ref=next(t for t in data['targets'] if t['kind']=='figure')
                    self.assertEqual(ref['cached'],f'1{sep}1')
                    self.assertFalse(audit_references(doc)['errors'])

    def test_manifest_and_link_existing_markers_without_rebuilding(self):
        with TemporaryDirectory() as td:
            root=Path(td);a=root/'a.docx'; m=root/'manifest.json'
            new_document(a,brief=sample(),manifest_out=m)
            d=Document(a);p=d.add_paragraph('追加核验[[FIG:network]]与引文[[CITE:srcB]]。')
            src=root/'source.docx';d.save(src)
            before=src.read_bytes();out=root/'linked.docx'
            result=link_references(src,out,m)
            self.assertEqual(src.read_bytes(),before)
            self.assertTrue(result['changed'])
            self.assertEqual(result['replaced'],2)
            self.assertNotIn('[[FIG',Document(out).paragraphs[-1].text)
            self.assertEqual(audit_references(out)['errors'],[])

    def test_split_markers_fail_closed_and_source_unmodified(self):
        with TemporaryDirectory() as td:
            root=Path(td);doc=root/'doc.docx';m=root/'refs.json';new_document(doc,brief=sample(),manifest_out=m)
            d=Document(doc);p=d.add_paragraph();p.add_run('前文[[FIG:');p.add_run('network]]。')
            src=root/'split.docx';d.save(src)
            before=src.read_bytes()
            with self.assertRaises(DocumentError):link_references(src,root/'should-not-exist.docx',m)
            self.assertFalse((root/'should-not-exist.docx').exists())
            self.assertEqual(before,src.read_bytes())

    def test_dangling_bookmark_and_duplicate_bookmark_are_errors(self):
        with TemporaryDirectory() as td:
            doc=Path(td)/'valid.docx';new_document(doc,brief=sample())
            d=Document(doc);forbid=BOOKMARK_PREFIX['figure']; bad=None
            for e in d._element.iter(q('bookmarkStart')):
                if (e.get(q('name')) or '').startswith(forbid):bad=e;break
            self.assertIsNotNone(bad)
            bad.getparent().remove(bad)
            tampered=Path(td)/'tampered.docx';d.save(tampered)
            report=audit_references(tampered)
            self.assertTrue(any(e['code']=='dangling-reference' for e in report['errors']))

    def test_invalid_reference_ids_and_modes_rejected_preflight(self):
        a=sample();a['blocks'][1]['text']='如[[FIG:missing]]所示。'
        with self.assertRaises(DocumentError):validate_brief(a)
        b=sample();b['bibliography'][0]['id']='bad id'
        with self.assertRaises(DocumentError):validate_brief(b)
        c=sample();c['caption_separator']='/'
        with self.assertRaises(DocumentError):validate_brief(c)
        d=sample();d['citation_mode']='author-year'
        with self.assertRaises(DocumentError):validate_brief(d)

    def test_preserves_special_unumbered_heading_and_docx_validity(self):
        with TemporaryDirectory() as td:
            brief=sample();brief['blocks'].append({'type':'heading','level':1,'text':'致谢','numbered':False})
            p=Path(td)/'d.docx';new_document(p,brief=brief)
            self.assertFalse(audit_references(p)['errors'])
            d=Document(p)
            self.assertTrue(any(x.text=='致谢' for x in d.paragraphs))
            with ZipFile(p) as z:self.assertIsNone(z.testzip())


if __name__=='__main__':unittest.main()
