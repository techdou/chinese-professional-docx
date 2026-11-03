"""Contract tests for GB/T 7714—2025 metadata and editable citations.

All sample records are explicitly fictional; validators do not verify sources.
"""
from __future__ import annotations
from copy import deepcopy
import json
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile
from docx import Document
from lxml import etree

ROOT=Path(__file__).resolve().parents[1]
if not (ROOT/'scripts').is_dir():
    ROOT=ROOT/'skills/chinese-professional-docx'
sys.path.insert(0,str(ROOT/'scripts'))
from bibliography_7714 import (audit_bibliography,audit_entry,format_entry,
                               assign_year_suffix)
from docx_ops import new_document, validate_brief, audit_references, DocumentError
from cross_references import q

EXAMPLE=json.loads((ROOT/'assets/examples/brief-gbt7714-2025.json').read_text())


def code_and_result(path):
    with ZipFile(path) as z:
        doc=etree.fromstring(z.read('word/document.xml'))
    return doc,[(v.text or '').strip() for v in doc.iter(q('instrText'))]


class TypedBibliographyTests(unittest.TestCase):
    def test_nine_core_types_and_2025_standard(self):
        audit=audit_bibliography(EXAMPLE['bibliography'])
        self.assertEqual(audit['errors'],0)
        self.assertEqual(audit['warnings'],0)
        self.assertEqual(len(audit['entries']),9)
        for row in EXAMPLE['bibliography']:
            s=format_entry(row)
            self.assertIn('['+row['type'],s)
            self.assertTrue(s.endswith('.'))
            self.assertNotIn('[[',s)
        s=format_entry(EXAMPLE['bibliography'][1])
        self.assertIn('[C]//',s)
        self.assertNotIn('[C]. //',s)

    def test_dataset_and_preprint_are_not_webpages(self):
        for k,t in [('datasetC','[DS/OL]'),('preprintD','[PP/OL]')]:
            e=next(e for e in EXAMPLE['bibliography'] if e['id']==k)
            f=format_entry(e)
            self.assertIn(t,f)
            self.assertIn('[2026-11-02]',f)
            self.assertIn(e['url'],f)

    def test_source_facts_are_never_invented(self):
        ds=deepcopy(EXAMPLE['bibliography'][2])
        del ds['accessed']
        self.assertIn('missing-accessed',audit_entry(ds)['errors'])
        self.assertIn('2026-11-02',format_entry(EXAMPLE['bibliography'][2]))
        invalid=deepcopy(EXAMPLE['bibliography'][3]);invalid['published_date']='2026-02-30'
        self.assertIn('invalid-published_date-date',audit_entry(invalid)['errors'])
        bad=deepcopy(EXAMPLE['bibliography'][4]);bad['url']='file:///etc/passwd'
        self.assertIn('invalid-online-url',audit_entry(bad)['errors'])
        self.assertRaises(DocumentError,validate_brief,{**EXAMPLE,'bibliography':[bad]})

    def test_doi_not_repeated_when_present_in_access_path(self):
        ds=deepcopy(EXAMPLE['bibliography'][2]);ds['doi']='10.5555/demo';ds['url']='https://doi.org/10.5555/demo'
        s=format_entry(ds)
        self.assertEqual(s.count('10.5555/demo'),1)
        ds['url']='https://example.org/data'
        s=format_entry(ds)
        self.assertIn('DOI:10.5555/demo',s)

    def test_preformatted_legacy_is_supported_but_not_claimed_verified(self):
        x={'id':'classic','text':'甲. 某教材[M]. 北京: 出版社, 2015.'}
        report=audit_bibliography([x],standard='gbt7714-2015-compat')
        self.assertEqual(report['errors'],0)
        self.assertEqual(report['warnings'],1)
        self.assertIn('unverified',report['entries'][0]['warnings'][0])
        self.assertEqual(format_entry(x),x['text'])
        report=audit_bibliography([x],citation_mode='author-year')
        self.assertTrue(report['errors'])

    def test_author_year_a_b_stable_and_generated_in_correct_order(self):
        sample=deepcopy(EXAMPLE)
        one=deepcopy(sample['bibliography'][0]);one['id']='a';one['title']='B title';one['year']='2025'
        two=deepcopy(one);two['id']='b';two['title']='A title'
        mapping=assign_year_suffix([one,two])
        self.assertEqual(mapping,{'a':'2025b','b':'2025a'})
        sample['citation_mode']='author-year'
        sample['bibliography']=[one,two]
        sample['blocks']=[{'type':'heading','level':1,'text':'绪论'},
            {'type':'paragraph','text':'参见[[CITE:a;b]]和[[CITE:b@123-125]]。'}]
        with TemporaryDirectory() as td:
            path=Path(td)/'doc.docx'
            new_document(path,brief=sample)
            xml,instructions=code_and_result(path)
            self.assertEqual(sum(c.startswith('REF _CP_BIB_') for c in instructions),3)
            self.assertFalse(any(c.startswith('SEQ CPBibliography') for c in instructions))
            self.assertFalse(audit_references(path)['errors'])
            self.assertFalse(any(w['code']=='citation-not-superscript'
                                 for w in audit_references(path)['warnings']))
            r=Document(path)
            bib=' '.join(p.text for p in r.paragraphs if p.style.name=='参考资料')
            self.assertIn('2025a',bib)
            self.assertIn('2025b',bib)

    def test_numeric_ranges_locators_real_ref_fields_and_unchanged_body(self):
        with TemporaryDirectory() as td:
            out=Path(td)/'r.docx'
            new_document(out,brief=EXAMPLE)
            xml,inst=code_and_result(out)
            nums=sum(c.startswith('REF _CP_BIB_') for c in inst)
            self.assertEqual(nums,10)
            self.assertFalse(audit_references(out)['errors'])
            rendered=''.join(t.text or '' for t in xml.iter(q('t')))
            self.assertIn('23-25',rendered)
            self.assertIn('-',rendered)
            self.assertNotIn('[[CITE:',rendered)
            self.assertEqual(out.suffix,'.docx')

    def test_invalid_locator_duplicate_cluster_and_reference_keys_fail_closed(self):
        for bad in ('[[CITE:journalA@-]]','[[CITE:journalA;conferenceB@12]]',
                    '[[CITE:journalA;journalA]]','[[CITE:missing]]'):
            with self.subTest(bad=bad):
                b=deepcopy(EXAMPLE)
                b['blocks']=[{'type':'paragraph','text':bad}]
                self.assertRaises(DocumentError,validate_brief,b)

    def test_metadata_schema_preserves_legacy_and_structured(self):
        import jsonschema
        schema=json.loads((ROOT/'assets/schemas/brief.schema.json').read_text())
        jsonschema.validate(EXAMPLE,schema)
        legacy=json.loads((ROOT/'assets/examples/brief-references.json').read_text())
        jsonschema.validate(legacy,schema)
        bad=deepcopy(EXAMPLE);bad['bibliography'][0]['unknown']='yes'
        with self.assertRaises(jsonschema.ValidationError):jsonschema.validate(bad,schema)

    def test_document_preserves_original_template_and_source(self):
        with TemporaryDirectory() as td:
            sample=deepcopy(EXAMPLE)
            sample['bibliography'][0]['doi']='badvalue'
            out=Path(td)/'should-not-exist.docx'
            self.assertRaises(DocumentError,new_document,out,brief=sample)
            self.assertFalse(out.exists())


if __name__=='__main__':unittest.main()
