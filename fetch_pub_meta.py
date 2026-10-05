#!/usr/bin/env python3
"""Look up DOI, keywords and abstract for each paper in build_site.PUBS (OpenAlex + Crossref).
Only confident title matches are kept. Output: pub_meta.json (reviewable; build_site.py reads it)."""
import difflib, json, os, re, time, urllib.error, urllib.parse, urllib.request
from build_site import PUBS, BRAND

MAIL = BRAND['email']
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'pub_meta.json')

def get(url):
    time.sleep(0.6)
    req = urllib.request.Request(url, headers={'User-Agent': f'SYNTERA-site/1.0 (mailto:{MAIL})'})
    for attempt in range(3):
        try:
            return json.load(urllib.request.urlopen(req, timeout=25))
        except urllib.error.HTTPError as e:
            if e.code != 429 or attempt == 2: raise
            time.sleep(5 * (attempt + 1))

def norm(t):
    return re.sub(r'[^a-z0-9 ]', '', re.sub(r'<[^>]+>', '', t).lower().replace('-', ' ')).strip()

def sim(a, b):
    return difflib.SequenceMatcher(None, norm(a), norm(b)).ratio()

def unabstract(idx):
    if not idx: return ''
    words = {}
    for w, pos in idx.items():
        for p in pos: words[p] = w
    return ' '.join(words[i] for i in sorted(words))

def clean(t):
    t = re.sub(r'<[^>]+>', ' ', t or '')
    t = re.sub(r'^\s*abstract\s*[:.]?\s*', '', t, flags=re.I)
    return re.sub(r'\s+', ' ', t).strip()

meta = json.load(open(OUT, encoding='utf-8')) if os.path.exists(OUT) else {}
for pid, typ, year, authors, title, venue, doi, areas, note in PUBS:
    if pid in meta and meta[pid].get('checked'): continue
    rec = dict(doi=doi, keywords=[], abstract='', checked=True, source='')
    try:
        r = get('https://api.openalex.org/works?search=%s&per-page=5&mailto=%s' % (urllib.parse.quote(title), MAIL))
        for w in r['results']:
            if not (w['title'] or '').lower().startswith(('corrigendum','erratum','retraction')) and sim(w['title'] or '', title) >= 0.9 and abs((w.get('publication_year') or year) - year) <= 1:
                rec['doi'] = rec['doi'] or (w.get('doi') or '').replace('https://doi.org/', '')
                rec['keywords'] = [k['display_name'] for k in w.get('keywords', [])][:8]
                rec['abstract'] = clean(unabstract(w.get('abstract_inverted_index')))
                rec['source'] = 'OpenAlex'
                break
    except urllib.error.HTTPError as e:
        if e.code == 429: print('rate-limited, skipping', pid); continue
        print('openalex fail', pid, e)
    except Exception as e:
        print('openalex fail', pid, e)
    try:
        if rec['doi'] and not rec['abstract']:
            c = get('https://api.crossref.org/works/%s' % urllib.parse.quote(rec['doi']))['message']
            rec['abstract'] = clean(c.get('abstract', ''))
        elif not rec['doi']:
            r = get('https://api.crossref.org/works?query.title=%s&rows=3&mailto=%s' % (urllib.parse.quote(title), MAIL))
            for c in r['message']['items']:
                if sim((c.get('title') or [''])[0], title) >= 0.92:
                    rec['doi'] = c['DOI']; rec['abstract'] = clean(c.get('abstract', '')); rec['source'] += '+Crossref'; break
    except Exception as e:
        print('crossref fail', pid, e)
    meta[pid] = rec
    json.dump(meta, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f"{pid:22} doi={'Y' if rec['doi'] else '-'} kw={len(rec['keywords'])} abs={'Y' if rec['abstract'] else '-'}")
json.dump(meta, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
