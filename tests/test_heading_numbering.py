"""Tests for native multilevel Word heading numbers and safe migration."""
from __future__ import annotations
import json
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile
from docx import Document
from docx.oxml import OxmlElement

ROOT = Path(__file__).resolve().parents[1]
if not (ROOT/'scripts').is_dir():
    ROOT = ROOT/'skills/chinese-professional-docx'
SCRIPT = ROOT/'scripts/docx_ops.py'
sys.path.insert(0,str(ROOT/'scripts'))
from heading_numbering import SCHEMES, DOCUMENT_TYPE_SCHEME, q, NS


def cli(*args, allowed=(0,)):
    proc=subprocess.run([sys.executable,str(SCRIPT),*map(str,args)],text=True,capture_output=True)
    if proc.returncode not in allowed:
        raise AssertionError(f'CLI {args} exit {proc.returncode}: {proc.stdout}\n{proc.stderr}')
    return (json.loads(proc.stdout) if proc.stdout and proc.returncode in (0,3) else None), proc


def part(path,name):
    from lxml import etree
    with ZipFile(path) as z:return etree.fromstring(z.read(name))


def body_numbering(path):
    root=part(path,'word/document.xml')
    result=[]
    for p in root.findall('./w:body/w:p',NS):
        style=p.find('w:pPr/w:pStyle',NS)
        if style is None or style.get(q('val')) not in ('Heading1','Heading2','Heading3'):continue
        np=p.find('w:pPr/w:numPr',NS)
        ids=(np.find('w:ilvl',NS),np.find('w:numId',NS)) if np is not None else (None,None)
        result.append((''.join(t.text or '' for t in p.findall('.//w:t',NS)),
                      style.get(q('val')),ids[0].get(q('val')) if ids[0] is not None else None,
                      ids[1].get(q('val')) if ids[1] is not None else None))
    return result


class HeadingNumberingTests(unittest.TestCase):
    def _brief(self,scheme=None,doc_type='general'):
        obj={'title':'标题体系验证','document_type':doc_type,'blocks':[
            {'type':'heading','level':1,'text':'绪论'},
            {'type':'heading','level':2,'text':'研究背景'},
            {'type':'heading','level':3,'text':'研究问题'},
            {'type':'paragraph','text':'正文一。'},
            {'type':'heading','level':2,'text':'研究目标'},
            {'type':'heading','level':1,'text':'方法'},
            {'type':'heading','level':2,'text':'特征编码'},
            {'type':'heading','level':3,'text':'模态融合'},
            {'type':'heading','level':1,'text':'参考文献','numbered':False},
            {'type':'table','headers':['指标','值'],'rows':[['F1','0.91']]},
            {'type':'code','text':'if True:\n    print("ok")'},
        ]}
        if scheme is not None:obj['numbering']=scheme
        return obj

    def _create(self,tmp,brief,name):
        b=tmp/(name+'.json');b.write_text(json.dumps(brief,ensure_ascii=False))
        out=tmp/(name+'.docx');cli('validate-brief',b);cli('new','--brief',b,'--out',out)
        return out

    def test_profiles_are_native_multilevel_numbering_linked_to_headings(self):
        for scheme in SCHEMES:
            with self.subTest(scheme=scheme),TemporaryDirectory() as d:
                tmp=Path(d);out=self._create(tmp,self._brief(scheme),scheme)
                paras=body_numbering(out)
                self.assertEqual([x[1] for x in paras],['Heading1','Heading2','Heading3','Heading2','Heading1','Heading2','Heading3','Heading1'])
                self.assertEqual([x[2] for x in paras[:7]],['0','1','2','1','0','1','2'])
                self.assertTrue(all(x[3]==paras[0][3] for x in paras[:7]))
                self.assertEqual(paras[-1][3],'0')  # references unnumbered
                num=part(out,'word/numbering.xml')
                abstract=next(x for x in num.findall('w:abstractNum',NS)
                              if x.find('w:name',NS) is not None and x.find('w:name',NS).get(q('val'))=='CPDOCX-HEADING-'+scheme)
                levels=abstract.findall('w:lvl',NS)
                self.assertEqual(len(levels),3)
                self.assertEqual([e.find('w:pStyle',NS).get(q('val')) for e in levels],['Heading1','Heading2','Heading3'])
                self.assertEqual([e.find('w:lvlText',NS).get(q('val')) for e in levels], [s[1] for s in SCHEMES[scheme]])
                audit,_=cli('audit-headings',out,'--scheme',scheme,'--strict-review')
                self.assertFalse(audit['review_required'])
                self.assertEqual(audit['profile_counts'],{scheme:7})
                self.assertEqual(len(Document(out).tables),1)

    def test_default_scene_selection_and_explicit_none(self):
        for doc_type,scheme in DOCUMENT_TYPE_SCHEME.items():
            with self.subTest(doc_type=doc_type),TemporaryDirectory() as d:
                p=self._create(Path(d),self._brief(doc_type=doc_type),'doc')
                audit,_=cli('audit-headings',p)
                self.assertEqual(audit['profile_counts'],{scheme:7})
        with TemporaryDirectory() as d:
            p=self._create(Path(d),self._brief(scheme='none'),'no')
            self.assertTrue(all(x[3] is None for x in body_numbering(p)))

    def test_reject_manual_numbers_in_new_brief_and_schema_contract(self):
        with TemporaryDirectory() as d:
            tmp=Path(d);b=self._brief()
            b['blocks'][0]['text']='一、绪论'
            text=tmp/'brief.json';text.write_text(json.dumps(b,ensure_ascii=False))
            dest=tmp/'result.docx'
            r,_=cli('validate-brief',text)
            self.assertEqual(r['status'],'pass')
            _,p=cli('new','--brief',text,'--out',dest,allowed=(2,))
            self.assertEqual(p.returncode,2)
            self.assertIn('manually written number',p.stderr)
            self.assertFalse(dest.exists())
            b['blocks'][0]['numbered']=False
            text.write_text(json.dumps(b,ensure_ascii=False))
            cli('new','--brief',text,'--out',dest)
            self.assertEqual(body_numbering(dest)[0][-1],'0')

    def test_existing_doc_opt_in_conversion_and_idempotent(self):
        with TemporaryDirectory() as d:
            tmp=Path(d)
            src=tmp/'old.docx'
            doc=Document(ROOT/'assets/templates/base.docx')
            body=doc._element.body
            for el in list(body):
                if el.tag != q('sectPr'):body.remove(el)
            doc.add_heading('一、研究背景',level=1)
            p=doc.add_heading('1.1 技术路线',level=2)
            p.runs[-1].bold=True
            doc.add_heading('参考文献',level=1)
            doc.add_heading('后续规划',level=1)
            doc.add_paragraph('正文保持原样。')
            doc.save(src)
            original=src.read_bytes()
            kept=tmp/'skip.docx';res,_=cli('format-headings',src,'--scheme','chinese-formal','--out',kept)
            self.assertEqual(res['formatted'],1)
            self.assertEqual(res['skipped_manual'],2)
            self.assertEqual(body_numbering(kept)[0][-1],'0')
            out=tmp/'fixed.docx'
            res,_=cli('format-headings',src,'--convert-manual','--scheme','chinese-formal','--out',out)
            self.assertEqual(res['converted_manual'],2)
            self.assertEqual(res['formatted'],3)
            self.assertEqual([x[0] for x in body_numbering(out)],['研究背景','技术路线','参考文献','后续规划'])
            self.assertEqual(Document(out).paragraphs[-1].text,'正文保持原样。')
            self.assertEqual(src.read_bytes(),original)
            second=tmp/'second.docx'
            res,_=cli('format-headings',out,'--convert-manual','--scheme','chinese-formal','--out',second)
            self.assertFalse(res['changed'])
            self.assertEqual(out.read_bytes(),second.read_bytes())

    def test_existing_numbered_other_scheme_is_preserved(self):
        with TemporaryDirectory() as d:
            tmp=Path(d)
            src=self._create(tmp,self._brief('academic-decimal'),'base')
            result=tmp/'other.docx';res,_=cli('format-headings',src,'--scheme','humanities','--out',result)
            self.assertFalse(res['changed']);self.assertEqual(src.read_bytes(),result.read_bytes())
            self.assertEqual(res['skipped_existing_numbering'],7)

    def test_toc_field_is_real_word_field_and_does_not_number_front_matter(self):
        with TemporaryDirectory() as d:
            tmp=Path(d);b=self._brief('chinese-formal')
            b['blocks'].insert(0,{'type':'toc'})
            p=self._create(tmp,b,'toc');xml=part(p,'word/document.xml')
            self.assertTrue(any('TOC \\o "1-3"' in (e.text or '') for e in xml.findall('.//w:instrText',NS)))
            self.assertTrue(xml.findall('.//w:fldChar',NS))

if __name__ == '__main__': unittest.main()
