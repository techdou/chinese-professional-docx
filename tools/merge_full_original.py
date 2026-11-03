#!/usr/bin/env python3
"""Update a local full ZIP baseline in place conceptually, without replacing it.

The upgrade bundle does not carry redistributed fonts or the old Windows .venv.
This utility merges current source/assets with the user's existing original ZIP,
retaining the ORIGINAL local font binaries and environment entries verbatim.
Run locally; use a new --out path. Fonts and the local .venv are never carried. Uses only Python standard library.
"""
from __future__ import annotations
import argparse
from hashlib import sha256
from pathlib import Path, PurePosixPath
from zipfile import ZipFile, ZIP_DEFLATED

PREFIX='chinese-professional-docx/'
IGNORED={'.DS_Store'}

def safe(name):
    parts=PurePosixPath(name).parts
    return bool(name and name.startswith(PREFIX) and not name.startswith('/')
                and '\\' not in name and '..' not in parts)

def merge(original:Path, output:Path, upgraded:Path):
    if original.resolve()==output.resolve() or output.exists():
        raise ValueError('Use a new output path. Original ZIP will never be overwritten.')
    if not (upgraded/'SKILL.md').is_file():
        raise ValueError('Cannot locate upgraded SKILL.md next to this tool.')
    new_files={PREFIX+str(p.relative_to(upgraded)).replace('\\','/'):p
        for p in upgraded.rglob('*') if p.is_file() and not any(
           part in ('.git','.venv','__pycache__','.pytest_cache','.local','qa','dist') for part in p.parts)
        and p.name not in IGNORED and not p.suffix.lower() in
           ('.ttf','.ttc','.otf','.woff','.woff2','.pyc')}
    with ZipFile(original) as archive:
        entries=[i for i in archive.infolist() if not i.is_dir()]
        old={i.filename:i for i in entries}
        if len(old)!=len(entries) or not all(safe(n) for n in old):
            raise ValueError('Source ZIP contains duplicate or unsafe paths.')
        if PREFIX+'SKILL.md' not in old:
            raise ValueError('This is not the supplied chinese-professional-docx full archive.')
        preserved={n:i for n,i in old.items() if n not in new_files}
        output.parent.mkdir(parents=True,exist_ok=True)
        try:
            with ZipFile(output,'w',ZIP_DEFLATED) as z:
                for n,i in old.items():
                    if n in new_files:
                        z.writestr(n,new_files[n].read_bytes())
                    else:
                        z.writestr(i,archive.read(n))
                for n,p in sorted(new_files.items()):
                    if n not in old:z.writestr(n,p.read_bytes())
        except Exception:
            output.unlink(missing_ok=True)
            raise
    # The local baseline ZIP may already contain Python caches, including inside
    # its Windows .venv. Retain those unchanged while rejecting NEW leaked
    # source-control/cache paths from our local development tree.
    with ZipFile(original) as _in, ZipFile(output) as _out:
        def forbidden(n):
            return '/.git/' in n or '/__pycache__/' in n
        original_legacy = {n for n in _in.namelist() if forbidden(n)}
        unexpected = {n for n in _out.namelist() if forbidden(n)} - original_legacy
        if unexpected:
            output.unlink(missing_ok=True)
            raise ValueError('Merged archive introduces source-control/cache files: '
                             + ', '.join(sorted(unexpected)[:3]))
    # Check preservation by content hash without writing any extracted font file.
    preserved_fonts=[]
    with ZipFile(original) as a, ZipFile(output) as b:
        for n in preserved:
            if n.startswith(PREFIX+'assets/fonts/'):
                if sha256(a.read(n)).digest()!=sha256(b.read(n)).digest():
                    output.unlink(missing_ok=True)
                    raise ValueError('Preserved font entry changed: '+n)
                preserved_fonts.append(n)
        if PREFIX+'scripts/academic_styles.py' not in b.namelist():
            output.unlink(missing_ok=True)
            raise ValueError('Upgrade source missing in output ZIP.')
    print('New complete local ZIP:',output)
    print('Updated source files:',len(new_files))
    print('Original font file entries preserved byte-for-byte:',len(preserved_fonts))
    print('Original ZIP unchanged:',original)

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('original',type=Path,help='Your existing full chinese-professional-docx.zip')
    p.add_argument('--out',type=Path,required=True,help='New full merged ZIP')
    a=p.parse_args()
    merge(a.original,a.out,Path(__file__).resolve().parents[1])
