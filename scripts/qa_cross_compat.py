#!/usr/bin/env python3
"""Read-only Word/WPS/LibreOffice interop gate for Chinese academic DOCX.

This process NEVER claims native Microsoft Word or WPS coverage without a
verified native host. OOXML checks, optional LibreOffice round-trip checks,
and host-specific manual steps are separately reported.
"""
from __future__ import annotations
import argparse
from collections import Counter
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from lxml import etree

sys.path.insert(0,str(Path(__file__).parent))
from docx_ops import Package, DocumentError
from cross_references import audit_fields, q
from long_tables import audit_long_tables


def field_signature(root):
    """Count field instructions by category, independent of generated IDs."""
    labels=[]
    for e in root.iter(q('instrText')):
        raw=' '.join((e.text or '').strip().split())
        if not raw:continue
        labels.append(raw.split(' ',1)[0].upper())
    return dict(Counter(labels))


def inspect(path:Path) -> dict:
    package=Package(path)
    fields=audit_fields(package.document)
    tables=audit_long_tables(package.document)
    return {'name':path.name,'field_signature':field_signature(package.document),
            'references':fields,'tables':tables,
            'bookmarks':len(list(package.document.iter(q('bookmarkStart')))),
            'text_length':len(''.join(x.text or '' for x in package.document.iter(q('t')))),
            'package_integrity':'pass'}


def lo_roundtrip(path:Path) -> dict:
    soffice=shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice:return {'status':'not-tested','reason':'LibreOffice executable unavailable'}
    with TemporaryDirectory(prefix='cpdocx-roundtrip-') as td:
        root=Path(td);src=root/'in';out=root/'out';home=root/'profile'
        src.mkdir();out.mkdir();
        target=src/'interop.docx';shutil.copyfile(path,target)
        cmd=[soffice,'-env:UserInstallation='+home.as_uri(),'--headless','--norestore',
             '--convert-to','docx:Office Open XML Text','--outdir',str(out),str(target)]
        try:
            p=subprocess.run(cmd,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,timeout=80,
                env={**os.environ,'HOME':td})
        except (OSError,subprocess.TimeoutExpired) as exc:
            return {'status':'fail','error':str(exc)}
        dst=out/'interop.docx'
        if p.returncode!=0 or not dst.exists():
            return {'status':'fail','return_code':p.returncode,
                    'stderr':p.stderr[-500:],'stdout':p.stdout[-500:]}
        try:report=inspect(dst)
        except Exception as exc:return {'status':'fail','error':str(exc)}
        return {'status':'completed','result':report,
                'warning':'LibreOffice roundtrip does not prove Word/WPS equivalence.'}


def qa(path:Path,do_roundtrip:bool) -> dict:
    source=inspect(path)
    roundtrip=lo_roundtrip(path) if do_roundtrip else {'status':'not-tested','reason':'--roundtrip-libreoffice not provided'}
    errs=list(source['references']['errors'])+list(source['tables']['errors'])
    warnings=list(source['references']['warnings'])+list(source['tables']['warnings'])
    if roundtrip.get('status')=='completed':
        converted=roundtrip['result'];errs.extend(converted['references']['errors'])
        errs.extend(converted['tables']['errors'])
        # Roundtrip changed field types/number count can be legitimate host behavior,
        # but must be reviewed instead of silently counted as proof of equivalence.
        if source['field_signature']!=converted['field_signature']:
            warnings.append({'code':'roundtrip-field-type-change','source':source['field_signature'],
                             'roundtrip':converted['field_signature']})
        if source['text_length']>converted['text_length']*1.1:
            warnings.append({'code':'roundtrip-visible-text-shrink',
                             'original_chars':source['text_length'],
                             'roundtrip_chars':converted['text_length']})
    elif roundtrip.get('status')=='fail':errs.append({'code':'roundtrip-failed',**roundtrip})
    return {'source':source,'libreoffice':roundtrip,
            'microsoft_word':{'status':'not-tested','reason':'No Microsoft Word COM/UI session available; use native checklist.'},
            'wps_writer':{'status':'not-tested','reason':'No WPS Writer session available; use native checklist.'},
            'errors':errs,'warnings':warnings,
            'native_host_complete':False,
            'passed_automated_gates':not errs,
            'verification_warning':'All cross-application field results require a host update and human visual check.'}


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('input',type=Path)
    p.add_argument('--report',type=Path)
    p.add_argument('--roundtrip-libreoffice',action='store_true')
    x=p.parse_args()
    try:
        output=qa(x.input,x.roundtrip_libreoffice)
        text=json.dumps(output,ensure_ascii=False,indent=2)+'\n'
        if x.report:
            if x.report.exists():raise DocumentError('Output report already exists; do not overwrite.')
            x.report.parent.mkdir(parents=True,exist_ok=True)
            x.report.write_text(text,encoding='utf-8')
        print(text,end='')
        return 1 if output['errors'] else 0
    except (OSError,DocumentError) as e:
        print('ERROR: '+str(e),file=sys.stderr);return 2

if __name__=='__main__':raise SystemExit(main())
