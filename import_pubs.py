#!/usr/bin/env python3
"""Parse publications.txt (CV LaTeX + four Google Scholar lists) and write pubs_extra.json:
every paper not already in build_site.PUBS, de-duplicated by title. Priority: curated PUBS > CV > Scholar."""
import difflib, json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ['SYNTERA_NO_EXTRA'] = '1'
import build_site as B

RAW = open(os.path.join(HERE, 'publications.txt'), encoding='utf-8').read().replace('\r', '')
RAW = re.sub(r'(?<=\w)�(?=\w)', '’', RAW).replace('�', '')
LINES = RAW.split('\n')

def norm(t):
    return re.sub(r'[^a-z0-9]', '', t.lower().replace('optimisation', 'optimization').replace('recognising', 'recognizing'))

def initials_fmt(ini):
    ini = ini.replace('.', '')
    if ini[:2] == 'Sh': ini = 'S'
    return ' '.join(c + '.' for c in ini if c.isalpha())

def scholar_authors(s):
    s = s.strip()
    out = []
    for tok in re.split(r',\s*', s):
        tok = tok.strip()
        if not tok: continue
        if tok in ('...', '…'): out.append('…'); continue
        parts = tok.split()
        if len(parts) >= 2 and re.fullmatch(r'[A-Z][A-Za-z]{0,2}', parts[0]) and (parts[0].isupper() or parts[0] == 'Sh'):
            out.append('%s, %s' % (' '.join(parts[1:]), initials_fmt(parts[0])))
        elif len(parts) >= 2:
            out.append('%s, %s' % (parts[-1], initials_fmt(''.join(w[0] for w in parts[:-1]))))
        else:
            out.append(tok)
    return ', '.join(out)

# ───────── CV (LaTeX) ─────────
def tex(s):
    s = s.replace('\\Saremi{}', 'Saremi, S.').replace('\\&', '&').replace('\\ldots', '…').replace('--', '–').replace('\\_', '_')
    s = re.sub(r'\\\\', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def parse_cv():
    cv = '\n'.join(LINES[:271])
    out = []
    section = ''
    for blk in re.split(r'\\textbf\{|\\item ', cv):
        pass
    for m in re.finditer(r'\\textbf\{(Doctoral Dissertation|Books|Book Chapters|Journal Articles[^}]*|Conference Papers)\}(.*?)(?=\\textbf\{(?:Doctoral|Books|Book Chapters|Journal|Conference)|\Z)', cv, re.S):
        sec, body = m.group(1), m.group(2)
        typ = {'D': 'thesis', 'B': 'book' if sec == 'Books' else 'chapter', 'J': 'journal', 'C': 'conference'}[sec[0]]
        for it in re.findall(r'\\item (.*?)(?=\\item |\\end\{etaremune\})', body, re.S):
            t = re.search(r'(?<!\()\\textbf\{(.*?)\}\.?', it)
            if not t: continue
            title = tex(t.group(1)).rstrip('.')
            head = it[:t.start()]
            ym = re.search(r'\((\d{4})\)', head)
            year = int(ym.group(1)) if ym else 0
            au = tex(head[:ym.start()] if ym else head).strip().rstrip(',')
            au = re.sub(r'^(.*?),?\s*\(?$', r'\1', au)
            au = re.sub(r',\s*&\s*', ', ', au).replace(' & ', ', ')
            rest = it[t.end():]
            v = re.search(r'\\textit\{(.*?)\}(.*?)(?:\(\\textbf|\\href|$)', rest, re.S)
            if v:
                venue = tex(v.group(1) + v.group(2)).strip(' .,')
            else:
                venue = tex(re.sub(r'\\href.*', '', re.sub(r'\(\\textbf.*', '', rest))).strip(' .,')
                venue = re.sub(r'^In ', '', venue) if typ == 'chapter' else venue
            if sec == 'Doctoral Dissertation': venue = 'Doctoral dissertation, Griffith University, Brisbane, Australia'
            d = re.search(r'doi\.org/([^}\s]+)\}', it)
            doi = d.group(1).replace('\\_', '_') if d else ''
            q = re.search(r'\(\\textbf\{(Q\d[^}]*)\}\)', it)
            out.append(dict(type=typ, year=year, authors=au, title=title, venue=venue, doi=doi, note=q.group(1).split(',')[0] if q else '', src='cv'))
    return out

# ───────── Scholar lists ─────────
TERM = re.compile(r'^(?:(\d+)\*?\s*\t\s*(\d{4})?|(\d{4}))\s*$')
def parse_scholar():
    out = []
    i = 0
    n = len(LINES)
    while i < n:
        if LINES[i].strip() == 'Title' and i + 2 < n and LINES[i + 1].strip() == 'Cited by':
            i += 3; buf = []
            while i < n and not LINES[i].startswith('Articles 1'):
                ln = LINES[i].strip('\t ').rstrip()
                mt = TERM.match(ln.replace(' ', '\t', 0))
                mt2 = re.match(r'^\d+\*?\s+\d{4}$', ln)       # "193 2023" with a space
                if (mt or mt2) and buf:
                    y = re.findall(r'\d{4}', ln)
                    year = int(y[-1]) if y else 0
                    out.append(entry(buf, year)); buf = []
                elif re.fullmatch(r'\d+\*?\s*', ln) and buf:   # "1" / "2 " cited-by only, no year
                    out.append(entry(buf, 0)); buf = []
                elif ln.startswith('Affiliations of all authors'):
                    pass
                elif ln:
                    buf.append(ln)
                i += 1
            for e in tail(buf): out.append(e)
        i += 1
    return out

def entry(buf, year):
    title = buf[0]
    au = buf[1] if len(buf) > 1 else ''
    venue = ' '.join(buf[2:]) if len(buf) > 2 else ''
    return dict(raw_title=title, raw_au=au, year=year, venue=venue, src='scholar')

def tail(buf):
    # unterminated trailing lines: pairs of (title, authors[, venue]); keep titles only, authors are abbreviated lists
    res, i = [], 0
    while i < len(buf):
        title = buf[i]; au = buf[i + 1] if i + 1 < len(buf) else ''
        venue = ''; step = 2
        if i + 2 < len(buf) and not re.search(r'^[A-Z]{1,3} [A-Z]', buf[i + 2]) and len(buf[i + 2]) < 60 and ',' not in buf[i + 2][:12]:
            pass
        res.append(dict(raw_title=title, raw_au=au, year=0, venue=venue, src='scholar'))
        i += step
    return res

def classify(venue, title):
    v = (venue + ' ' + title).lower()
    if re.search(r'preprint|ssrn|under review', v): return 'preprint'
    if re.search(r'initiating and sustaining|evidence-based practice: an integ|gerontological nursing: a holistic|ethical, legal and social issues', v): return 'chapter'
    if re.search(r'universiti teknologi|university of tasmania|^queensland university of technology$', venue.lower()): return 'other'
    if re.search(r'proceedings|conference|congress|symposium|workshop|acis|compendium|uniSC research conference|papers from', v, re.I): return 'conference'
    if re.search(r'springer|nature-inspired|\bchapter\b|academic press|john wiley|optimisation algorithms for hand posture', v): return 'chapter'
    if re.search(r'^(university|queensland university|griffith)|^[^,]*university[^,]*$', venue.lower()) and not re.search(r'journal|proceedings', venue.lower()): return 'other'
    if not venue.strip(): return 'other'
    return 'journal'

AREA_RULES = [
    ('home', r'smart home|ageing in place|aging-in-place|assistive|gerontech|ageing|aged care|older adults in the home|health-assistive'),
    ('health', r'health|medical|clinical|patient|nurs|pregnan|opioid|injur|fracture|sports|concussion|older adults|ageing|aging|disease|physio|obesity|epidemiolog|hospital|wellbeing|dispensing|mobility program'),
    ('agri', r'agricultur|crop|farm'),
    ('edu', r'education|(?<!machine )(?<!deep )learning|student|assessment|teaching|curricul|academic integrity|undergraduate|generative ai'),
    ('connect', r'\biot\b|internet of thing|intrusion|cyber|sensor|blockchain|wireless|adversar|signcryption|signature'),
    ('mobility', r'vehic|v2i|\biov\b|traffic|transport'),
]
def areas_for(title, venue):
    t = (title + ' ' + venue).lower()
    return [a for a, rx in AREA_RULES if re.search(rx, t)]

def main():
    have = [(norm(p[4]), p[4]) for p in B.PUBS]
    cand = []
    for e in parse_cv():
        cand.append(e)
    for e in parse_scholar():
        e = dict(e)
        e['title'] = e.pop('raw_title').strip()
        e['authors'] = scholar_authors(e.pop('raw_au'))
        e['doi'] = ''; e['note'] = ''
        e['type'] = classify(e['venue'], e['title'])
        cand.append(e)
    DROP = {norm(t) for t in [
        'Biogeography-based optimisation with chaos', 'in Designing Photonic Crystal Filters',
        'Exploring the factors affecting sustainable human resource productivity in railway lines. Sustainability, 14 (1), 225',
        'Identifying places of interest for tourists using knowledge discovery [Conference poster].',
        'Knowledge Discovery for Tourism Using Data Mining and Qualitative Analysis',
        'Grasshopper Optimisation Algorithm', 'Grasshopper Optimization for Multi-Objective Problems',
        'Multi-objective Grey Wolf Optimizer', 'Salp Swarm Algorithm']}
    cand = [e for e in cand if not (e['src'] == 'cv' and e['year'] == 0)]
    cand = [e for e in cand if not re.match(r'^[A-Z]{1,3} [A-Z]\w+,', e['title'])]
    cand.append(dict(type='conference', year=0, authors='Dermody, G., Bryant, R. A.', venue='', doi='', note='', src='scholar',
                     title='Participatory approach to build capacity: Nurse-led research to overcome insufficient mobility in hospitalized older adults'))
    kept = []
    for e in cand:
        k = norm(e['title'])
        if len(k) < 12 or k in DROP: continue
        dup = None
        for hk, _ in have:
            if hk == k or difflib.SequenceMatcher(None, hk, k).ratio() >= 0.93 or (len(k) > 25 and (k in hk or hk in k) and min(len(k), len(hk)) > 0.8 * max(len(k), len(hk))):
                dup = hk; break
        if dup:
            continue
        have.append((k, e['title'])); kept.append(e)
    # fill missing years from any duplicate; leave 0 as unknown
    res = []
    used = {p[0] for p in B.PUBS}
    for e in sorted(kept, key=lambda x: (-x['year'], x['title'])):
        pid = re.sub(r'[^a-z0-9]+', '-', e['title'].lower()).strip('-')[:42].strip('-') + ('-%d' % e['year'] if e['year'] else '')
        while pid in used: pid += 'x'
        used.add(pid)
        res.append([pid, e['type'], e['year'], e['authors'], e['title'], e['venue'], e['doi'], areas_for(e['title'], e['venue']), e['note']])
    json.dump(res, open(os.path.join(HERE, 'pubs_extra.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    from collections import Counter
    print(len(cand), 'candidates ->', len(res), 'new;', Counter(r[1] for r in res))

if __name__ == '__main__':
    main()
