"""Long-table continuation and safety regression tests."""
from __future__ import annotations
from pathlib import Path
from tempfile import TemporaryDirectory
import sys
import unittest
from zipfile import ZipFile
from lxml import etree
from docx import Document

ROOT=Path(__file__).resolve().parents[1]
if not (ROOT/'scripts').is_dir():
    ROOT=ROOT/'skills/chinese-professional-docx'
sys.path.insert(0,str(ROOT/'scripts'))
from docx_ops import new_document, validate_brief, DocumentError, Package
from long_tables import audit_long_tables
from cross_references import q, NS


def brief(mode='explicit-chunks',rows_per_part=7,include_note=True):
    tbl={'type':'table','id':'ablation','caption':'消融实验结果','headers':['方法','Macro-F1','Micro-F1'],
         'rows':[[f'模型{i:02d}',f'{i/100:.2f}',f'{i/120:.3f}'] for i in range(20)],
         'continuation':{'mode':mode} if mode=='repeat-header' else
                        {'mode':mode,'rows_per_part':rows_per_part,'page_break_before':True}}
    if include_note:tbl['note']='注：数据仅作排版测试。'
    return {'title':'长表测试','document_type':'thesis-chapter','blocks':[
        {'type':'heading','level':1,'text':'实验'},
        {'type':'paragraph','text':'如[[TABLE:ablation]]所示。'},tbl,
    ]}


class LongTableTests(unittest.TestCase):
    def test_explicit_continuations_repeat_header_only_one_seq_caption(self):
        with TemporaryDirectory() as td:
            path=Path(td)/'test.docx';new_document(path,brief=brief())
            doc=Document(path)
            self.assertEqual(len(doc.tables),3)
            self.assertEqual([len(t.rows) for t in doc.tables],[8,8,7])
            self.assertEqual(sum(p.text.startswith('续表') for p in doc.paragraphs),2)
            self.assertEqual(sum(p.text.startswith('表') for p in doc.paragraphs),1)
            self.assertEqual(sum('注：' in p.text for p in doc.paragraphs),1)
            xml=Package(path).document
            inst=[(x.text or '').strip() for x in xml.iter(q('instrText'))]
            self.assertEqual(sum(s.startswith('SEQ CPTable \\* ARABIC') for s in inst),1)
            self.assertEqual(sum(s.startswith('REF _CP_TAB_') for s in inst),3)  # body + 2 continuations
            for table in xml.iter(q('tbl')):
                self.assertIsNotNone(table.find('w:tr[1]/w:trPr/w:tblHeader',NS))
                self.assertTrue(all(row.find('w:trPr/w:cantSplit',NS) is not None for row in table.findall('w:tr',NS)))
            audit=audit_long_tables(xml)
            self.assertEqual(audit['errors'],[])
            self.assertEqual(audit['warnings'],[])
            self.assertIn('ablation',audit['groups'])

    def test_native_repeat_header_preserves_single_table(self):
        with TemporaryDirectory() as td:
            path=Path(td)/'native.docx';new_document(path,brief=brief(mode='repeat-header'))
            self.assertEqual(len(Document(path).tables),1)
            self.assertFalse(any(p.text.startswith('续表') for p in Document(path).paragraphs))
            audit=audit_long_tables(Package(path).document)
            self.assertEqual(audit['errors'],[])
            self.assertEqual(audit['warnings'],[])

    def test_reject_unsafe_chunk_settings_without_output(self):
        for val in (None,True,0,-1,20,21,'10'):
            with self.subTest(val=val):
                s=brief();s['blocks'][2]['continuation']['rows_per_part']=val
                with self.assertRaises(DocumentError):validate_brief(s)
        s=brief();s['blocks'][2].pop('id')
        with self.assertRaises(DocumentError):validate_brief(s)
        s=brief();s['blocks'][2]['continuation']={'mode':'repeat-header','rows_per_part':7}
        with self.assertRaises(DocumentError):validate_brief(s)

    def test_no_content_loss_or_mutation_in_long_table(self):
        with TemporaryDirectory() as td:
            out=Path(td)/'result.docx'; data=brief(); prior=[r[:] for r in data['blocks'][2]['rows']]
            new_document(out,brief=data)
            self.assertEqual(prior,data['blocks'][2]['rows'])
            vals=[tuple(c.text for c in row.cells) for t in Document(out).tables for row in t.rows[1:]]
            self.assertEqual(vals, list(map(tuple,prior)))
            self.assertTrue(ZipFile(out).testzip() is None)


if __name__=='__main__':unittest.main()
