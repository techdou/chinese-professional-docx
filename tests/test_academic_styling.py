"""Regression tests for full v1.2.1 -> v1.4.0 academic style upgrade."""
from __future__ import annotations
import json
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile
from io import BytesIO
from lxml import etree
from docx import Document

ROOT=Path(__file__).resolve().parents[1]
if not (ROOT/'scripts').is_dir():
    ROOT=ROOT/'skills/chinese-professional-docx'
SCRIPTS=ROOT/'scripts'
sys.path.insert(0,str(SCRIPTS))
from academic_styles import W,NS,q,TABLE_STYLE_ID,CODE_GRAY,CODE_STYLE_NAMES
from docx_ops import parse_xml


def doc_part(path:Path,part='word/document.xml'):
    with ZipFile(path) as z:
        return parse_xml(z.read(part))


def signature(path:Path):
    x=doc_part(path)
    return (''.join(el.text or '' for el in x.findall('.//w:t',NS)),
            len(x.findall('.//w:tbl',NS)),
            len(x.findall('.//w:drawing',NS)))


def table_edges(t):
    return {etree.QName(x).localname:x.get(q('val')) for x in t.findall('w:tblPr/w:tblBorders/*',NS)}


def run(*args):
    p=subprocess.run([sys.executable,str(SCRIPTS/'docx_ops.py'),*map(str,args)],
                     capture_output=True,text=True)
    if p.returncode:
        raise AssertionError(f'Code={p.returncode}, stdout={p.stdout}, stderr={p.stderr}')
    return json.loads(p.stdout)


class AcademicStylingTests(unittest.TestCase):
    def test_all_original_word_assets_preserved_and_neutral(self):
        files=[ROOT/'assets/templates/base.docx',ROOT/'assets/templates/base.dotx']+sorted((ROOT/'assets/examples').glob('*.docx'))
        self.assertEqual(len(files),6)
        for path in files:
            with self.subTest(file=path.name):
                s=doc_part(path,'word/styles.xml')
                code=[x for x in s.findall('w:style',NS) if x.find('w:name',NS) is not None
                      and x.find('w:name',NS).get(q('val'))=='代码段']
                self.assertEqual(len(code),1)
                self.assertEqual(code[0].find('w:pPr/w:shd',NS).get(q('fill')),CODE_GRAY)
                self.assertIsNone(code[0].find('w:pPr/w:pBdr',NS))
                self.assertTrue(any(x.get(q('styleId'))==TABLE_STYLE_ID for x in s.findall('w:style',NS)))
                for st in s.findall('w:style',NS):
                    if st.get(q('type'))=='paragraph':
                        name=st.find('w:name',NS)
                        is_code=name is not None and name.get(q('val')) in CODE_STYLE_NAMES
                        if not is_code:self.assertIsNone(st.find('w:pPr/w:shd',NS))
                        self.assertIsNone(st.find('w:pPr/w:pBdr',NS))
                        color=st.find('w:rPr/w:color',NS)
                        if color is not None:self.assertEqual(color.get(q('val')),'000000')
                for t in doc_part(path).findall('.//w:tbl',NS):
                    self.assertEqual(table_edges(t),{'top':'single','left':'nil','bottom':'single',
                                                      'right':'nil','insideH':'nil','insideV':'nil'})
                    self.assertFalse(t.findall('.//w:tcPr/w:shd',NS))
                    self.assertTrue(t.findall('w:tr[1]/w:trPr/w:tblHeader',NS))

    def test_brief_code_and_table_defaults_and_explicit_grid(self):
        brief={'title':'论文排版检查','blocks':[
            {'type':'heading','level':1,'text':'实验设置'},
            {'type':'paragraph','text':'这里是普通正文，实验设置应当表达清楚。'},
            {'type':'code','text':'def add(x, y):\n    return x + y\n\nprint(add(1, 2))'},
            {'type':'table','caption':'方法性能比较',
             'headers':['方法','Macro-F1（%）','耗时（s）'],
             'rows':[['方法 A','80.2','10.1'],['方法 B','82.5','9.8']]},
            {'type':'table','caption':'混淆矩阵（网格特例）','style':'grid',
             'headers':['标签','0','1'],'rows':[['0','4','2'],['1','1','5']]},
            {'type':'table','style':'none','headers':['备注','值'],
             'rows':[['检查','通过']]},
        ]}
        with TemporaryDirectory() as td:
            b=Path(td)/'b.json';b.write_text(json.dumps(brief,ensure_ascii=False),encoding='utf-8')
            out=Path(td)/'new.docx'
            run('validate-brief',b)
            run('new','--brief',b,'--out',out)
            d=Document(out)
            codes=[p for p in d.paragraphs if p.style.name=='代码段']
            self.assertEqual(len(codes),1)
            self.assertEqual(codes[0].text,'def add(x, y):\n    return x + y\n\nprint(add(1, 2))')
            self.assertEqual(len(d.tables),3)
            x=doc_part(out)
            t=x.findall('.//w:tbl',NS)
            self.assertEqual(table_edges(t[0])['insideV'],'nil')
            self.assertEqual(table_edges(t[0])['top'],'single')
            self.assertEqual(t[0].find('w:tblPr/w:tblBorders/w:top',NS).get(q('sz')),'12')
            self.assertEqual(t[0].find('w:tr/w:tc/w:tcPr/w:tcBorders/w:bottom',NS).get(q('sz')),'8')
            self.assertEqual(table_edges(t[1])['insideV'],'single')
            self.assertEqual(table_edges(t[2])['top'],'nil')
            cp=next(p for p in x.findall('.//w:p',NS)
                    if p.find('w:pPr/w:pStyle',NS) is not None
                    and p.find('w:pPr/w:pStyle',NS).get(q('val'))==codes[0].style.style_id)
            self.assertEqual(cp.find('w:pPr/w:shd',NS).get(q('fill')),CODE_GRAY)
            self.assertEqual(cp.find('w:pPr/w:ind',NS),None)
            audit=run('audit',out,'--summary')
            self.assertEqual(audit['errors'],0)

    def test_existing_doc_safe_code_format_and_idempotence(self):
        with TemporaryDirectory() as td:
            src=Path(td)/'src.docx'
            d=Document(ROOT/'assets/templates/base.docx')
            d.add_paragraph('  正文带原有缩进  ',style='Normal')
            code=d.add_paragraph('if True:\n    print("ok")',style='代码段')
            ppr=code._p.get_or_add_pPr()
            for e in ppr.findall(q('shd')):ppr.remove(e)
            d.save(src)
            out=Path(td)/'new.docx'
            result=run('format-code',src,'--out',out)
            self.assertTrue(result['changed'])
            self.assertEqual(result['code_paragraphs'],1)
            self.assertEqual(signature(src),signature(out))
            self.assertTrue(src.exists())
            result2=run('format-code',out,'--out',Path(td)/'second.docx')
            self.assertFalse(result2['changed'])
            self.assertEqual(signature(out),signature(Path(td)/'second.docx'))

    def test_table_repair_is_opt_in_and_preserves_cells(self):
        with TemporaryDirectory() as td:
            src=Path(td)/'src.docx';d=Document(ROOT/'assets/templates/base.docx')
            t=d.add_table(rows=2,cols=2)
            t.style='Table Grid'
            t.cell(0,0).text='甲';t.cell(0,1).text='乙'
            t.cell(1,0).text='A';t.cell(1,1).text='B'
            d.save(src)
            default=Path(td)/'untouched.docx'
            r=run('format-tables',src,'--out',default)
            self.assertFalse(r['changed'])
            self.assertEqual(src.read_bytes(),default.read_bytes())
            dst=Path(td)/'out.docx'
            r=run('format-tables',src,'--out',dst,'--all')
            self.assertEqual(r['formatted_tables'],1)
            self.assertEqual(signature(src),signature(dst))
            self.assertEqual(table_edges(doc_part(dst).find('.//w:tbl',NS))['insideV'],'nil')
            self.assertEqual([c.text for row in Document(dst).tables[0].rows for c in row.cells],['甲','乙','A','B'])
            run('format-tables',dst,'--out',Path(td)/'none.docx','--all','--style','none')
            self.assertEqual(table_edges(doc_part(Path(td)/'none.docx').find('.//w:tbl',NS))['top'],'nil')

    def test_rejects_overwrite_and_unrecognized_code(self):
        with TemporaryDirectory() as td:
            src=Path(td)/'s.docx';d=Document();d.add_paragraph('print("do not guess code")');d.save(src)
            dst=Path(td)/'code.docx'
            r=run('format-code',src,'--out',dst)
            self.assertEqual(r['code_paragraphs'],0)
            self.assertEqual(src.read_bytes(),dst.read_bytes())
            p=subprocess.run([sys.executable,str(SCRIPTS/'docx_ops.py'),'format-code',str(src),'--out',str(src)],capture_output=True,text=True)
            self.assertEqual(p.returncode,2)
            self.assertIn('new output path',p.stderr)

    def test_dotx_content_type_and_headings_unchanged(self):
        with TemporaryDirectory() as td:
            src=ROOT/'assets/templates/base.dotx'
            dst=Path(td)/'new.dotx'
            run('format-code',src,'--out',dst)
            with ZipFile(dst) as z:
                self.assertTrue(z.testzip() is None)
                self.assertIn('word/styles.xml',z.namelist())
            original=doc_part(src,'word/styles.xml')
            revised=doc_part(dst,'word/styles.xml')
            for name in ('Normal','Heading1','Heading2','Heading3'):
                a=original.find(f"w:style[@w:styleId='{name}']/w:rPr/w:sz",NS)
                b=revised.find(f"w:style[@w:styleId='{name}']/w:rPr/w:sz",NS)
                self.assertEqual(a.get(q('val')) if a is not None else None,
                                 b.get(q('val')) if b is not None else None)

if __name__=='__main__':unittest.main()
