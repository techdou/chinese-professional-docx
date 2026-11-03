"""Full local ZIP baseline update must preserve existing binary assets and avoid development debris."""
from __future__ import annotations
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import importlib.util
import unittest

ROOT = Path(__file__).resolve().parents[1]
if not (ROOT/'scripts').is_dir():
    ROOT = ROOT/'skills/chinese-professional-docx'
SPEC = importlib.util.spec_from_file_location('local_zip_merger', ROOT/'tools/merge_full_original.py')
module = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(module)


class CompleteZipMergeTests(unittest.TestCase):
    def test_merge_preserves_private_font_bytes_and_legacy_cache_without_new_git(self):
        with TemporaryDirectory() as tmp:
            base = Path(tmp)
            source = base/'old.zip'; out = base/'new.zip'; upgraded = base/'upgraded'
            (upgraded/'scripts').mkdir(parents=True)
            (upgraded/'.git').mkdir()
            (upgraded/'__pycache__').mkdir()
            (upgraded/'SKILL.md').write_text('---\nname: chinese-professional-docx\n---\n',encoding='utf-8')
            (upgraded/'scripts/academic_styles.py').write_bytes(b'# upgraded source\n')
            (upgraded/'.git/config').write_text('never-export')
            (upgraded/'__pycache__/metadata.bin').write_bytes(b'never-export')
            prefix='chinese-professional-docx/'
            font=b'legacy-font-local-only-bytes'
            with ZipFile(source,'w') as old:
                old.writestr(prefix+'SKILL.md','old')
                old.writestr(prefix+'assets/fonts/example.ttf',font)
                old.writestr(prefix+'.venv/Lib/site-packages/foo/__pycache__/a.pyc',b'pre-existing')
                old.writestr(prefix+'scripts/keep.py','intact')
            module.merge(source,out,upgraded)
            with ZipFile(out) as merged:
                self.assertEqual(merged.read(prefix+'assets/fonts/example.ttf'),font)
                self.assertEqual(merged.read(prefix+'scripts/keep.py'),b'intact')
                self.assertEqual(merged.read(prefix+'scripts/academic_styles.py'),b'# upgraded source\n')
                self.assertEqual(merged.read(prefix+'.venv/Lib/site-packages/foo/__pycache__/a.pyc'),b'pre-existing')
                self.assertNotIn(prefix+'.git/config',merged.namelist())
                self.assertNotIn(prefix+'__pycache__/metadata.bin',merged.namelist())
                self.assertIsNone(merged.testzip())
            with ZipFile(source) as original:
                self.assertEqual(original.read(prefix+'SKILL.md'),b'old')

    def test_existing_output_is_not_overwritten(self):
        with TemporaryDirectory() as tmp:
            root=Path(tmp); original=root/'source.zip';out=root/'output.zip'
            (root/'src').mkdir();(root/'src/SKILL.md').write_text('skill')
            with ZipFile(original,'w') as z:z.writestr('chinese-professional-docx/SKILL.md','old')
            out.write_bytes(b'pristine-output')
            with self.assertRaises(ValueError):module.merge(original,out,root/'src')
            self.assertEqual(out.read_bytes(),b'pristine-output')

if __name__=='__main__':unittest.main()
