#!/usr/bin/env python3
"""Render trusted DOCX/DOTX to page PNGs using LibreOffice and PyMuPDF.

Outputs are previews, not proof of Microsoft Word/WPS fidelity. Use a sandbox
for untrusted Office documents. Embedded content and external links are never
intentionally activated by this script; rendering software still needs isolation.
"""
from __future__ import annotations
import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from docx_ops import Package, DocumentError, convert, write_new, parse_xml, xml_bytes, q
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED


def preview_font_copy(source: Path, out: Path) -> dict:
    """Explicit Noto fallback in a temporary preview COPY, never the source."""
    package=Package(source)
    mapping={'宋体':'Noto Serif CJK SC','SimSun':'Noto Serif CJK SC',
             '黑体':'Noto Sans CJK SC','SimHei':'Noto Sans CJK SC'}
    fc=shutil.which('fc-match')
    if fc:
        for family in set(mapping.values()):
            found=subprocess.run([fc,'-f','%{family}',family],capture_output=True,
                                 text=True,timeout=10,check=False).stdout
            if family not in found:
                raise DocumentError('Requested preview fallback font is not installed: '+family)
    changes=0;buf=BytesIO()
    with ZipFile(buf,'w',ZIP_DEFLATED) as z:
        for info in package.infos:
            data=package.data[info.filename]
            if info.filename.startswith('word/') and info.filename.endswith('.xml'):
                root=parse_xml(data);count=0
                for fonts in root.iter(q('rFonts')):
                    old=fonts.get(q('eastAsia'))
                    if old in mapping:
                        fonts.set(q('eastAsia'),mapping[old]);count+=1
                        fonts.attrib.pop(q('eastAsiaTheme'),None)
                if count:data=xml_bytes(root);changes+=count
            z.writestr(info,data)
    write_new(out,buf.getvalue())
    return {'mode':'explicit-preview-only','mapping':mapping,'font_nodes_changed':changes,
            'source_modified':False,'notice':'Noto preview does not certify original-font Word layout.'}


def _pdf_to_pages(pdf: Path, output: Path, dpi: int) -> list[str]:
    """Shared by both backends: rasterize the PDF into page PNGs and keep preview.pdf."""
    import fitz
    pngs=[]
    with fitz.open(pdf) as doc:
        for i,page in enumerate(doc,1):
            path=output/f'page-{i}.png';pix=page.get_pixmap(matrix=fitz.Matrix(dpi/72,dpi/72),alpha=False)
            write_new(path,pix.tobytes('png'));pngs.append(path.name)
    write_new(output/'preview.pdf',pdf.read_bytes())
    return pngs


def render(source: Path, output: Path, dpi: int = 130, timeout: int = 120,
           fallback_cjk: bool = False) -> dict:
    source=source.resolve();output=output.resolve()
    Package(source)
    if output.exists() and any(output.iterdir()):
        raise DocumentError('Preview directory must be absent or empty; old QA pages are never reused.')
    office=shutil.which('soffice') or shutil.which('libreoffice')
    if not office:
        mac=Path('/Applications/LibreOffice.app/Contents/MacOS/soffice')
        office=str(mac) if mac.exists() else None
    if not office: raise DocumentError('LibreOffice is missing. Install it or render with the host Word/DOCX tool.')
    try:
        import fitz
    except ImportError as exc:
        raise DocumentError('Missing PyMuPDF. Install requirements-render.txt in the working virtual environment.') from exc
    output.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='chinese-docx-render-') as tmp:
        tmp=Path(tmp);working=tmp/'input.docx'
        font_substitution={'mode':'system-default','source_modified':False}
        if fallback_cjk:
            converted=tmp/'original.docx';convert(source,converted)
            font_substitution=preview_font_copy(converted,working)
        else:convert(source,working)
        profile=tmp/'office-profile';home=tmp/'home';home.mkdir()
        env=os.environ.copy();env['HOME']=str(home);env['SAL_USE_VCLPLUGIN']='svp'
        cmd=[office,'-env:UserInstallation='+profile.as_uri(),'--headless','--convert-to','pdf',
             '--outdir',str(tmp),str(working)]
        try:
            proc=subprocess.run(cmd,capture_output=True,text=True,timeout=timeout,env=env,check=False)
        except subprocess.TimeoutExpired as exc:
            raise DocumentError('LibreOffice timed out. No successful visual review may be claimed.') from exc
        pdf=tmp/'input.pdf'
        if proc.returncode!=0 or not pdf.exists() or pdf.stat().st_size==0:
            raise DocumentError('LibreOffice conversion failed: '+(proc.stderr or proc.stdout)[-2000:])
        # Keep preview.pdf local for comparing field layout; it is not a normal final deliverable.
        pngs=_pdf_to_pages(pdf,output,dpi)
        result={'input':source.name,'pages':len(pngs),'png_files':pngs,
                'renderer':'LibreOffice + PyMuPDF','visual_review':'pending-human-or-agent-image-inspection',
                'font_notice':'Original document font names are preserved; previews may use substitutes.',
                'preview_font_substitution':font_substitution}
        fc=shutil.which('fc-match')
        if fc:
            matches={}
            for font in ('宋体','黑体','Times New Roman'):
                probe=subprocess.run([fc,'-f','%{family}',font],capture_output=True,text=True,timeout=10,check=False)
                matches[font]=probe.stdout.strip()
            result['fontconfig_matches']=matches
        write_new(output/'render.json',(json.dumps(result,ensure_ascii=False,indent=2)+'\n').encode())
        return result


def render_word(source: Path, output: Path, dpi: int = 130) -> dict:
    """Render via host Word (COM) then PyMuPDF: highest fidelity on Windows,
    including embedded fonts. The source is opened read-only and never modified.
    Note: the generic --timeout flag applies to the libreoffice backend only;
    the COM call itself has no timeout guard."""
    source=source.resolve();output=output.resolve()
    Package(source)
    if output.exists() and any(output.iterdir()):
        raise DocumentError('Preview directory must be absent or empty; old QA pages are never reused.')
    try: import fitz  # noqa: F401  (probe before COM work)
    except ImportError as exc:
        raise DocumentError('Missing PyMuPDF. Install requirements-render.txt in the working virtual environment.') from exc
    try: import win32com.client, pywintypes
    except ImportError as exc:
        raise DocumentError('Word backend requires pywin32 and a desktop Word install (Windows only).') from exc
    output.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='chinese-docx-render-') as tmp:
        tmp=Path(tmp);pdf=tmp/'render.pdf'
        word=None;doc=None
        try:
            word=win32com.client.Dispatch('Word.Application')
            word.Visible=False;word.DisplayAlerts=0
            doc=word.Documents.Open(str(source),ReadOnly=True)
            doc.SaveAs2(str(pdf),FileFormat=17)
        except pywintypes.com_error as exc:
            raise DocumentError('Word COM operation failed: '+str(exc)) from exc
        finally:
            try:
                if doc is not None: doc.Close(False)
            except Exception: pass
            try:
                if word is not None: word.Quit()
            except Exception: pass
        if not pdf.exists() or pdf.stat().st_size==0:
            raise DocumentError('Word COM PDF export produced no output.')
        pngs=_pdf_to_pages(pdf,output,dpi)
        result={'input':source.name,'pages':len(pngs),'png_files':pngs,
                'renderer':'Word COM + PyMuPDF','visual_review':'pending-human-or-agent-image-inspection',
                'font_notice':'Host Word renders with installed or embedded fonts; closest to real Word fidelity on this machine.',
                'source_modified':False}
        write_new(output/'render.json',(json.dumps(result,ensure_ascii=False,indent=2)+'\n').encode())
        return result


def main() -> int:
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('input',type=Path);p.add_argument('--out-dir',type=Path,required=True)
    p.add_argument('--dpi',type=int,default=130);p.add_argument('--timeout',type=int,default=120)
    p.add_argument('--fallback-cjk',action='store_true',help='Use installed Noto CJK fonts in a temporary preview copy only.')
    p.add_argument('--backend',choices=['libreoffice','word'],default='libreoffice',help='word = host Word via COM (Windows): highest fidelity, honors embedded fonts.')
    a=p.parse_args()
    try:
        if not 72<=a.dpi<=300:raise DocumentError('DPI must be between 72 and 300.')
        if a.fallback_cjk and a.backend=='word':raise DocumentError('--fallback-cjk applies to the libreoffice backend only.')
        if a.backend=='word':result=render_word(a.input,a.out_dir,a.dpi)
        else:result=render(a.input,a.out_dir,a.dpi,a.timeout,a.fallback_cjk)
        print(json.dumps(result,ensure_ascii=False,indent=2));return 0
    except (DocumentError,OSError,ValueError,subprocess.SubprocessError) as exc:
        print('ERROR: '+str(exc),file=sys.stderr);return 2
if __name__=='__main__':raise SystemExit(main())
