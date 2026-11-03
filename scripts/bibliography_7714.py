"""Conservative GB/T 7714—2025 structured-reference processing.

The standard has more optional branches than an automated formatter can infer.
The supported core types are deliberately explicit: missing facts are reported,
not fabricated. Legacy preformatted entries remain editable but unverified.
The user, institutional or publisher style may override the default standard.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping
from urllib.parse import urlparse
import re

STANDARD = 'gbt7714-2025'
STANDARDS = (STANDARD, 'gbt7714-2015-compat', 'institution-custom')
CITATION_MODES = ('numeric-superscript','numeric-inline','author-year')
TYPES = ('M','J','N','C','D','R','S','P','EB','A','CM','DS','PP','G','CP','DB','Z')
SUPPORTED_FORMATS = frozenset(('M','J','C','D','R','S','EB','DS','PP'))
FIELDS = frozenset(('id','text','type','title','subtitle','authors','year','container',
                    'volume','issue','pages','publisher','place','edition','version',
                    'platform','published_date','accessed','url','doi','pid',
                    'conference','conference_year','institution','medium','publisher_date',
                    'report_number','standard_number','language','conference_place','citation_author'))
ISO_DATE = re.compile(r'^\d{4}-\d{2}-\d{2}$')
DOI_RE = re.compile(r'^10\.\d{4,9}/\S+$',re.I)
TRAILING_NUM = re.compile(r'^\s*\[\d+\]')


class BibliographyError(ValueError):
    pass


def _nonblank(value: Any) -> bool:
    return isinstance(value,str) and bool(value.strip())


def author_list(entry: Mapping[str,Any]) -> str:
    authors=entry.get('authors',[])
    if not authors:return ''
    language=entry.get('language','zh')
    join='，' if language=='zh' else ', '
    suffix='等' if language=='zh' else 'et al.'
    return join.join(authors[:3]) + (join+suffix if len(authors)>3 else '')


def author_short(entry: Mapping[str,Any]) -> str:
    authors=entry.get('authors',[])
    if not authors:return '佚名' if entry.get('language','zh')=='zh' else 'Anon'
    if entry.get('citation_author'):return entry['citation_author']
    if len(authors)==1:return authors[0]
    if entry.get('language','zh')=='zh':return authors[0]+' 等'
    return authors[0].split(' ',1)[0]+' et al.'


def _date(value: str) -> bool:
    if not ISO_DATE.fullmatch(value):return False
    try: date.fromisoformat(value);return True
    except ValueError:return False


def audit_entry(entry: Any, *, standard: str=STANDARD) -> dict[str,Any]:
    errors=[];warnings=[]
    if not isinstance(entry,dict):
        return {'id':None,'errors':['entry-not-object'],'warnings':[],'format':'invalid'}
    uid=entry.get('id')
    if not isinstance(uid,str) or not re.fullmatch(r'[A-Za-z][A-Za-z0-9_.-]{0,63}',uid):
        errors.append('invalid-id')
    extra=set(entry)-FIELDS
    if extra:errors.append('unknown-fields:'+','.join(sorted(extra)))
    if standard not in STANDARDS:errors.append('unknown-standard')
    if 'text' in entry:
        if not _nonblank(entry.get('text')):errors.append('empty-preformatted-text')
        if TRAILING_NUM.match(entry.get('text','')):errors.append('manual-bibliography-number')
        if set(entry)-{'id','text'}:errors.append('mixed-preformatted-and-structured')
        warnings.append('preformatted-unverified-against-'+standard)
        return {'id':uid,'type':None,'format':'preformatted','errors':errors,'warnings':warnings}
    typ=entry.get('type')
    if typ not in TYPES:errors.append('invalid-or-missing-type')
    if not _nonblank(entry.get('title')):errors.append('missing-title')
    if standard!=STANDARD:
        errors.append('structured-format-requires-gbt7714-2025')
    if typ not in SUPPORTED_FORMATS:
        errors.append('unsupported-structured-type-use-preformatted-after-review')
    authors=entry.get('authors')
    if authors is not None:
        if not isinstance(authors,list) or not authors or not all(_nonblank(x) for x in authors):
            errors.append('invalid-authors')
        elif len(set(authors))!=len(authors):warnings.append('duplicate-author-names')
    elif typ in ('J','M','C','D','R','DS','PP'):
        warnings.append('missing-creators-review-source')
    if entry.get('language','zh') not in ('zh','en'):
        errors.append('invalid-language')
    for field in sorted(FIELDS-{'id','text','type','authors','language'}):
        if field in entry and not _nonblank(entry[field]):errors.append('invalid-'+field)
    year=entry.get('year')
    if year is not None and not (isinstance(year,str) and re.fullmatch(r'(?:1|2)\d{3}',year)):
        errors.append('invalid-year')
    for field in ('published_date','accessed','publisher_date'):
        if field in entry and not _date(entry[field]):errors.append('invalid-'+field+'-date')
    if 'conference_year' in entry and not re.fullmatch(r'(?:1|2)\d{3}',entry['conference_year']):
        errors.append('invalid-conference-year')
    if 'medium' in entry and entry['medium'] not in ('OL','CD','DK','MT'):
        errors.append('invalid-medium')
    online=bool(entry.get('url')) or entry.get('medium')=='OL'
    if entry.get('url'):
        u=urlparse(entry['url'])
        if u.scheme.lower() not in ('http','https') or not u.netloc or u.username or u.password:
            errors.append('invalid-online-url')
    if online and entry.get('medium') not in (None,'OL'):
        errors.append('url-medium-mismatch')
    if online and not _nonblank(entry.get('url')):errors.append('online-access-path-required')
    if entry.get('doi') and not DOI_RE.fullmatch(entry['doi'].removeprefix('doi:').removeprefix('DOI:')):
        errors.append('invalid-doi')
    if entry.get('pid') and entry.get('doi'):
        warnings.append('pid-and-doi-review-duplicate')
    required={
      'J':('container','year'), 'M':('publisher','place','year'),
      'C':('conference','conference_year'), 'D':('institution','place','year'),
      'R':('year',), 'S':('standard_number','year'),
      'EB':('url','accessed'), 'DS':('url','accessed'),
      'PP':('url','accessed')
    }.get(typ,())
    for k in required:
        if not _nonblank(entry.get(k)):errors.append('missing-'+k)
    if online and typ in ('J','M','C','D','R','S','DS','PP','EB') and not _nonblank(entry.get('accessed')):
        # Journal DOI landing pages may omit access dates in the published
        # examples; flag review instead of inventing a date.
        if typ in ('J','M','C','D','R','S'):
            warnings.append('online-access-date-review')
    if typ in ('DS','PP') and not entry.get('version'):
        warnings.append('version-not-provided-review-if-known')
    if typ in ('DS','PP') and not entry.get('platform'):
        warnings.append('platform-not-provided-review-if-known')
    if typ=='C' and entry.get('container'):
        warnings.append('conference-article-vs-proceedings-review')
    if typ=='S' and not entry.get('title') and entry.get('standard_number'):
        warnings.append('standard-title-required')
    return {'id':uid,'type':typ,'format':'structured',
            'errors':list(dict.fromkeys(errors)),'warnings':list(dict.fromkeys(warnings))}


def audit_bibliography(entries: Any, *, standard: str=STANDARD,
                       citation_mode: str='numeric-superscript') -> dict[str,Any]:
    if standard not in STANDARDS:raise BibliographyError('Unsupported standard '+standard)
    if citation_mode not in CITATION_MODES:raise BibliographyError('Unsupported citation_mode '+citation_mode)
    if not isinstance(entries,list) or len(entries)>2000:
        raise BibliographyError('Bibliography must be a list with at most 2000 records.')
    results=[audit_entry(e,standard=standard) for e in entries]
    keys=[r['id'] for r in results if r['id']]
    duplicates=[k for k,n in Counter(keys).items() if n>1]
    if duplicates:results.append({'id':None,'errors':['duplicate-ids:'+','.join(duplicates)],'warnings':[]})
    if citation_mode=='author-year':
        for obj,result in zip(entries,results):
            if result['errors']:continue
            if result['format']=='preformatted':result['errors'].append('author-year-needs-structured-metadata')
            elif not (obj.get('authors') and obj.get('year')):
                result['errors'].append('author-year-needs-authors-and-year')
    return {'standard':standard,'citation_mode':citation_mode,'entries':results,
            'errors':sum(len(r['errors']) for r in results),
            'warnings':sum(len(r['warnings']) for r in results),
            'review_required':any(r['warnings'] or r['errors'] for r in results),
            'claims':'Metadata rules cover supported core types only; source verification and school/journal style review remain necessary.'}


def assign_year_suffix(entries: list[dict[str,Any]]) -> dict[str,str]:
    """Disambiguate authors/year deterministically by title, never array order."""
    groups=defaultdict(list)
    for entry in entries:
        if 'text' not in entry and entry.get('authors') and entry.get('year'):
            groups[(author_short(entry).casefold(),entry['year'])].append(entry)
    years={entry['id']:entry.get('year','') for entry in entries}
    for (_,year),same in groups.items():
        if len(same)<=1:continue
        if len(same)>26:
            raise BibliographyError('More than 26 same-author/year records require manual name-year disambiguation.')
        for i,entry in enumerate(sorted(same,key=lambda e:(e.get('title','').casefold(),e['id']))):
            years[entry['id']]=year+chr(ord('a')+i)
    return years


def format_entry(entry: Mapping[str,Any], *, standard: str=STANDARD,
                 citation_mode: str='numeric-superscript',
                 disambiguated_year: str | None=None) -> str:
    if 'text' in entry:return entry['text'].strip()
    check=audit_entry(entry,standard=standard)
    if check['errors']:raise BibliographyError(f"Invalid {entry.get('id')}: {check['errors']}")
    kind=entry['type'];lang=entry.get('language','zh')
    author=author_list(entry)
    title=entry['title'].strip()
    if entry.get('subtitle'):title+='：'+entry['subtitle'].strip()
    online=bool(entry.get('url')) or entry.get('medium')=='OL'
    marker=f'[{kind}{"/OL" if online else ("/"+entry["medium"] if entry.get("medium") else "")}]'
    year=(disambiguated_year or entry.get('year',''))
    sep='，' if lang=='zh' else ', '
    # Author-year bibliography places year immediately after creators, rather
    # than keeping the numeric-system placement. Exact publisher overrides
    # should be handled by a reviewed preformatted record.
    lead=(author+(sep+year if year and citation_mode=='author-year' else '')+'. ') if author else ''
    head=lead+title+marker+('' if kind=='C' else '. ')
    suffix=''
    if kind=='J':
        suffix=entry['container']+((', '+(entry['year'] if citation_mode!='author-year' else '')) if citation_mode!='author-year' else '')
        suffix+=(', '+entry['volume'] if entry.get('volume') else '')
        suffix+=(' ('+entry['issue']+')' if entry.get('issue') else '')
        suffix+=(': '+entry['pages'] if entry.get('pages') else '')
    elif kind=='M':
        suffix=(entry.get('edition','')+'. ' if entry.get('edition') else '')+entry['place']+': '+entry['publisher']
        if citation_mode!='author-year':suffix+=', '+entry['year']
        if entry.get('pages'):suffix+=': '+entry['pages']
    elif kind=='C':
        suffix='//'+entry['conference']+', '+entry['conference_year']
        if entry.get('pages'):suffix+=': '+entry['pages']
    elif kind=='D':
        suffix=entry['place']+': '+entry['institution']
        if citation_mode!='author-year':suffix+=', '+entry['year']
        if entry.get('pages'):suffix+=': '+entry['pages']
    elif kind=='R':
        suffix=(entry.get('place','')+': ' if entry.get('place') else '')+entry.get('publisher','')
        if citation_mode!='author-year':suffix+=(', ' if suffix else '')+entry['year']
        if entry.get('report_number'):suffix+=(', '+entry['report_number'])
    elif kind=='S':
        suffix=entry['standard_number']
        if citation_mode!='author-year':suffix+=', '+entry['year']
    elif kind in ('DS','PP'):
        if entry.get('version'):suffix+=entry['version']+'. '
        if entry.get('platform'):suffix+=entry['platform']
        if entry.get('published_date'):suffix+=' ('+entry['published_date']+')'
        suffix+=' ['+entry['accessed']+']'
    elif kind=='EB':
        if entry.get('platform'):suffix+=entry['platform']
        if entry.get('published_date'):suffix+=' ('+entry['published_date']+')'
        suffix+=(' ['+entry['accessed']+']')
    suffix=suffix.strip().rstrip('.')
    if online:
        if suffix:suffix+='.'
        suffix+=' '+entry['url']
        doi=entry.get('doi')
        if doi and doi.casefold() not in entry['url'].casefold():suffix+=' DOI:'+doi
        pid=entry.get('pid')
        if pid and pid.casefold() not in entry['url'].casefold():suffix+=' '+pid
    elif entry.get('doi'):
        suffix+=' DOI:'+entry['doi']
    if suffix:head+=suffix.strip().lstrip('.')
    return head.rstrip(' .')+'.'


def citation_label(entry: Mapping[str,Any], year: str) -> str:
    return author_short(entry)+'，'+year
