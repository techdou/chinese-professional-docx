#!/usr/bin/env python3
"""Read-only, standard-library preflight. Never installs packages or uses a network."""
from __future__ import annotations
import argparse
import importlib.util
import json
from pathlib import Path
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
REQUIRED = ('SKILL.md', 'assets/style-profile.json', 'assets/templates/base.docx',
            'assets/templates/base.dotx', 'scripts/docx_ops.py')


def available(module: str) -> bool:
    try:
        return importlib.util.find_spec(module) is not None
    except (ImportError, ValueError):
        return False


def probe_word_com() -> dict:
    """Best-effort activation probe of host Word via COM; read-only, always quits."""
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return {'status': 'unavailable', 'reason': 'pywin32 not installed'}
    word = None
    try:
        word = win32com.client.Dispatch('Word.Application')
        version = str(word.Version)
        return {'status': 'ready', 'word_version': version}
    except Exception as exc:  # COM errors are opaque strings; type name is enough
        return {'status': 'unavailable', 'reason': type(exc).__name__}
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def probe_fonts_dir(root: Path) -> dict:
    fonts = root / 'assets' / 'fonts'
    if not fonts.is_dir():
        return {'status': 'absent'}
    files = sorted(p.name for p in fonts.iterdir() if p.is_file())
    return {'status': 'present', 'count': len(files), 'files': files,
            'notice': 'Local-only font assets (licensed Windows fonts): preview registration '
                      'and embedding via scripts; never redistribute.'}


def check(root: Path = ROOT, require_render: bool = False) -> dict:
    missing = [name for name in REQUIRED if not (root / name).is_file()]
    packages = {name: available(module) for module, name in
                [('docx', 'python-docx'), ('lxml', 'lxml'), ('fitz', 'PyMuPDF')]}
    office = shutil.which('soffice') or shutil.which('libreoffice')
    mac = Path('/Applications/LibreOffice.app/Contents/MacOS/soffice')
    if not office and mac.is_file():
        office = str(mac)
    core = sys.version_info >= (3, 10) and not missing and packages['python-docx'] and packages['lxml']
    preview = bool(core and office and packages['PyMuPDF'])
    fonts = {}
    fc = shutil.which('fc-match')
    if fc:
        for name, aliases in [('宋体', ('宋体', 'SimSun')), ('黑体', ('黑体', 'SimHei')),
                              ('Times New Roman', ('Times New Roman',))]:
            try:
                p = subprocess.run([fc, '-f', '%{family}', name], capture_output=True,
                                   text=True, timeout=5, check=False)
                match = p.stdout.strip()
                exact = any(x.strip() in aliases for x in match.split(','))
                fonts[name] = {'matched_family': match, 'exact_family': exact and p.returncode == 0}
            except (OSError, subprocess.SubprocessError) as exc:
                fonts[name] = {'probe_error': type(exc).__name__, 'exact_family': False}
    return {'status': 'ready' if core and (preview or not require_render) else 'blocked',
            'python': {'executable': sys.executable, 'version': sys.version.split()[0],
                       'supported': sys.version_info >= (3, 10)},
            'skill_root': str(root), 'missing_assets': missing, 'packages': packages,
            'document_ops': 'ready' if core else 'blocked',
            'bundled_preview': 'ready' if preview else 'unavailable',
            'libreoffice': office, 'fonts': fonts, 'font_probe': 'fontconfig' if fc else 'not-available',
            'word_com': probe_word_com(),
            'assets_fonts': probe_fonts_dir(root),
            'network_accessed': False, 'files_modified': False, 'host_discovery': 'not-tested',
            'next_action': 'Run the selected document route; inspect font substitutions before visual QA.' if core
                else 'Use a prepared Python environment; install scripts/requirements.txt only with authorization.',
            'notice': 'Font lookup and filesystem checks do not prove actual Word/WPS rendering or client activation.'}


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--require-render', action='store_true', help='Require the bundled preview backend, not a host renderer.')
    a = p.parse_args()
    report = check(require_render=a.require_render)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report['status'] == 'ready' else 2


if __name__ == '__main__':
    raise SystemExit(main())
