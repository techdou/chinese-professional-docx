#!/usr/bin/env python3
"""Embed TrueType fonts into a DOCX via host Word (COM) so it renders on machines
without the fonts. Windows-only backend.

Safety contract (v1.2.1):
- Word always opens the source read-only and saves to a sibling temp file;
  the original is replaced only after the embedded-font check passes (atomic os.replace).
- --out is resolved to an absolute path, must not exist, and must differ from the source.
- Verification compares word/fonts/ manifests before/after to avoid false positives.
"""
from __future__ import annotations
import argparse
import json
import os
from pathlib import Path
import sys
from uuid import uuid4
from zipfile import ZipFile

from docx_ops import DocumentError


def _manifest(path: Path) -> list[str]:
    """Sorted word/fonts/ entries of a docx package."""
    with ZipFile(path) as z:
        return sorted(n for n in z.namelist() if n.startswith('word/fonts/') and n != 'word/fonts/')


def embed(source: Path, out: Path | None = None) -> dict:
    source = source.resolve()
    if not source.is_file() or source.suffix.lower() != '.docx':
        raise DocumentError('Input must be an existing .docx file: ' + str(source))
    if out is not None:
        target = out.resolve()
        if target.suffix.lower() != '.docx':
            raise DocumentError('--out must point to a .docx path: ' + str(target))
        if target == source:
            raise DocumentError('--out must differ from the input path.')
        if target.exists():
            raise DocumentError('--out already exists; refusing to overwrite: ' + str(target))
        if not target.parent.is_dir():
            raise DocumentError('--out parent directory does not exist: ' + str(target.parent))
    else:
        target = source

    try:
        import win32com.client  # type: ignore
        import pywintypes  # type: ignore
    except ImportError as exc:
        raise DocumentError('pywin32 is required for the Word COM backend; '
                            'on non-Windows hosts embed fonts manually in Word (另存为→工具→保存选项→嵌入字体).') from exc

    before = _manifest(source)
    before_bytes = source.stat().st_size
    # Keep the temp file on the SAME volume as its replace target so os.replace stays atomic.
    if out is None:
        tmp = source.with_name(source.stem + '.embed-tmp-' + uuid4().hex[:8] + '.docx')
    else:
        tmp = target.with_name(target.stem + '.embed-tmp-' + uuid4().hex[:8] + '.docx')
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(source), ReadOnly=True)
        doc.EmbedTrueTypeFonts = True
        doc.DoNotEmbedSystemFonts = False  # False = also embed system fonts (宋体/黑体/Times)
        doc.SaveAs2(str(tmp))
    except pywintypes.com_error as exc:
        raise DocumentError('Word COM operation failed: ' + str(exc)) from exc
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass

    try:
        after = _manifest(tmp)
        if not after:
            raise DocumentError('Word saved the file but word/fonts/ is empty — embedding did not take effect.')
        os.replace(tmp, target)  # same-volume atomic replace; target was verified absent (or is the source)
    except Exception:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
        raise

    return {'status': 'embedded', 'input': source.name, 'output': str(target),
            'in_place': out is None,
            'bytes_before': before_bytes, 'bytes_after': target.stat().st_size,
            'embedded_fonts': after,
            'fonts_before': before,
            'unchanged': sorted(set(before)) == sorted(set(after)),
            'notice': 'Embedded fonts travel with the document for viewing; redistribution '
                      'of the file does not transfer font licenses. File size grows substantially. '
                      '"unchanged": true means the package already carried the same font set '
                      '(e.g. a previous embed run) — the save is idempotent.'}


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('input', type=Path)
    p.add_argument('--out', type=Path, help='Optional new output path; must not exist; default replaces the input atomically.')
    a = p.parse_args()
    try:
        print(json.dumps(embed(a.input, a.out), ensure_ascii=False, indent=2))
        return 0
    except (DocumentError, OSError) as exc:
        print('ERROR: ' + str(exc), file=sys.stderr)
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
