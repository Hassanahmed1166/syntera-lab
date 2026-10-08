#!/usr/bin/env python3
"""SYNTERA Research Group static site generator. Run: python build_site.py  ->  site/"""
import hashlib, html, json, os, posixpath, re, shutil, zipfile, xml.etree.ElementTree as ET
from PIL import Image, ImageOps

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, 'docs')
E = html.escape

# ───────────────────────── brand ─────────────────────────
SITE_URL = 'https://syntera.au'
PAGES = []   # (fname, noindex) collected for sitemap

BRAND = dict(
    name='SYNTERA Research Group', short='SYNTERA',
    full='SYNTERA Research Group: Applied AI & Connected Systems',
    tagline='Intelligence in Synergy',
    host='University of the Sunshine Coast', host_short='UniSC',
    host_url='https://www.usc.edu.au/',
    email='join@syntera.au',   # general address (JSON-LD, privacy requests)
    email_student='prospective.student@syntera.au', email_join='join@syntera.au', email_collab='collaborations@syntera.au',
    scholar='https://scholar.google.com/citations?user=oDhSiscAAAAJ',
)

AREAS = [
 dict(id='health', name='AI for Health', short='SYNTERA Health', color='#E83E8C', icon='health',
      topics=['Disease prediction','Medical imaging','Remote monitoring','Wearables'],
      why='Health systems generate more data than clinicians can read. AI can turn images, signals and records into earlier, fairer and more personal care, provided it is trustworthy and fits clinical practice.',
      challenges=['Data quality, bias and missing context in health records','Explaining model output to clinicians and patients','Privacy, consent and ethical use of sensitive data','Moving from prototype to everyday clinical workflow'],
      approach='We combine machine learning, medical image and signal analysis, digital health and bibliometric methods, working with clinicians and health information researchers.'),
 dict(id='home', name='Smart Health Home', short='SYNTERA Home', color='#8B5CF6', icon='home',
      topics=['Elderly care','Fall detection','Ambient assisted living','Home sensors'],
      why='Most people want to age in their own homes. Passive sensors and AI can support independence and safety, if people and carers trust them.',
      challenges=['Interpreting sensor streams as meaningful health events','Technology adoption by older adults, carers and clinicians','Privacy, autonomy and ethical design of monitoring','Integrating home data into clinical decisions'],
      approach='Human-centred design informed by nursing and caregiving needs, linked with IoT sensing, data visualisation for health professionals and stakeholder research.'),
 dict(id='agri', name='AI for Agriculture', short='SYNTERA Agri', color='#16A34A', icon='agri',
      topics=['Crop disease detection','Precision farming','Smart irrigation','Yield prediction'],
      why='Food systems face climate pressure and volatile markets. AI can help farmers and policy makers see problems sooner and plan with better forecasts.',
      challenges=['Limited labelled field data','Robust vision models in changing outdoor conditions','Reliable forecasting under market and weather shocks'],
      approach='Computer vision, robotics and autonomous systems, together with machine learning and time-series forecasting for agricultural data.'),
 dict(id='edu', name='AI for Education', short='SYNTERA Edu', color='#D97706', icon='edu',
      topics=['Personalized learning','Learning analytics','AI tutors','Student success'],
      why='Generative AI is changing how students learn and how institutions assess. Evidence-based frameworks can help education use it responsibly.',
      challenges=['Responsible and fair integration of generative AI','Academic integrity and authentic assessment','Understanding online peer learning and engagement','Supporting career and skill development'],
      approach='Design-based research, learning analytics, language technologies and recommender systems, tested with real students and educators.'),
 dict(id='connect', name='Internet of Things', short='SYNTERA Connect', color='#3A7BFF', icon='connect',
      topics=['Edge AI','Sensor networks','IoT security','Energy-efficient devices'],
      why='Billions of small connected devices sense and act in the real world. They need to be secure, lightweight and intelligent at the edge.',
      challenges=['Securing resource-constrained devices','Intrusion detection on the edge','Data quality across protocol stacks','Trust, privacy and reliability of connected systems'],
      approach='Cryptography and cybersecurity, adversarially robust machine learning, network protocols and reliability engineering for IoT and IoMT.'),
 dict(id='mobility', name='Internet of Vehicles', short='SYNTERA Mobility', color='#0E9F8E', icon='mobility',
      topics=['V2X communication','Traffic prediction','Driver safety','Smart transport'],
      why='Connected vehicles need fast, safe and secure communication with other vehicles and infrastructure to make roads safer and transport smarter.',
      challenges=['Low-latency, safety-critical V2I/V2X messaging','Secure, lightweight protocols for vehicular networks','Consensus and resilience under high mobility'],
      approach='Protocol design and simulation, secure message exchange, blockchain-based consensus and machine-learning-assisted protocol selection.'),
]
AREA = {a['id']: a for a in AREAS}
KNOWS_ABOUT = ['Applied artificial intelligence', 'Machine learning', 'Deep learning', 'Internet of Things', 'Internet of Vehicles', 'AI for health', 'Smart health home', 'AI for agriculture', 'AI for education', 'Cybersecurity', 'Explainable AI', 'Research collaboration', 'Industry research partnerships']

ICONS = {
 'health':'<path d="M3 12h4l2-5 4 10 2-5h6"/><path d="M12 21s-8-5-8-11a4.5 4.5 0 0 1 8-2.5A4.5 4.5 0 0 1 20 10"/>',
 'home':'<path d="M3 11l9-8 9 8"/><path d="M5 10v10h14V10"/><path d="M9.5 15a3.5 3.5 0 0 1 5 0M8 13a6 6 0 0 1 8 0"/>',
 'agri':'<path d="M5 19c0-8 4-13 14-14 0 9-4 14-12 14"/><path d="M5 19c3-5 6-8 10-10"/>',
 'edu':'<path d="M2 9l10-5 10 5-10 5z"/><path d="M6 11.5V16c3 2.5 9 2.5 12 0v-4.5"/><path d="M22 9v6"/>',
 'connect':'<rect x="7" y="7" width="10" height="10" rx="2"/><path d="M10 2v3M14 2v3M10 19v3M14 19v3M2 10h3M2 14h3M19 10h3M19 14h3"/>',
 'mobility':'<path d="M5 15l1.5-5A2 2 0 0 1 8.4 8.5h7.2a2 2 0 0 1 1.9 1.5L19 15"/><path d="M4 15h16v3H4z"/><circle cx="7.5" cy="18" r=".6"/><circle cx="16.5" cy="18" r=".6"/><path d="M8 4.5a6 6 0 0 1 8 0M10 6.5a3 3 0 0 1 4 0"/>',
}
def icon(name, cls='ico'):
    return f'<svg class="{cls}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{ICONS[name]}</svg>'

# ───────────────────────── members ─────────────────────────
def read_xlsx():
    z = zipfile.ZipFile(os.path.join(ROOT, 'members (1).xlsx'))
    ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    q = '{%s}' % ns['m']
    ss = [''.join(t.text or '' for t in si.iter(q + 't')) for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', ns)]
    rows = []
    for row in ET.fromstring(z.read('xl/worksheets/sheet1.xml')).iter(q + 'row'):
        d = {}
        for c in row.findall('m:c', ns):
            col = re.match(r'[A-Z]+', c.get('r')).group()
            v = c.find('m:v', ns)
            d[col] = '' if v is None else (ss[int(v.text)] if c.get('t') == 's' else v.text)
        rows.append(d)
    return rows[1:]

def slug(s):
    s = re.sub(r'[^a-z0-9]+', '-', s.lower()).strip('-')
    return s

FLAGS = {'Australia':'au','Pakistan':'pk','Bangladesh':'bd','Iran':'ir','Argentina':'ar','Norway':'no','Iraq':'iq','China':'cn','United Kingdom':'gb','Algeria':'dz'}
# Photo file per member id
PHOTOS = {
 'shahrzad-saremi':'shahrzad saremi.jpeg','rania-shibl':'rania shibl.jpeg','mostafa-kamalpour':'Mostafa Kamalpour.jpeg',
 'hassan-ahmed':'hassan-ahmed1.png','dana-dermody':'Gordana (Dana) Dermody.jpeg','abdullah-khan':'abdullah-khan.jpg',
 'nazmul-hossain':'Nazmul Hossain.png','arooj-fatima':'arooj fatima.png','samia-mujahid':'samia mujahid.HEIC',
 'nida-ali':'nida ali.jpeg','svetlana-kolos':'Svetlana Kolos.jpg','sadegh-rajaei':'Sadegh Rajaei.png',
 'tobias-romano':'Tobias.jpeg','shiva-jahanaray':'shiva.jpg','aittezaz-ahmad':'Aittezaz Ahmad.jpeg',
 'mohamadali-rezaeimanesh':'Mohamadali Rezaeimanesh.jpeg','amir-h-malekijoo':'Amir Malekijo.jpeg','jie-zhu':'Jie Zhu.jpeg',
 'thiwanka-kaushalya-nagasanga':'Thiwanka Kaushalya Nagasanga.PNG','muhammad-irfan-aslam':'Muhammad Irfan Aslam.jpeg',
 'sepehr-amooeinejad':'Sepehr Amooeinejad.jpeg','damilare-ogunjobi':'Damilare Ogunjobi.jpeg','bisma-ali':'Bisma Ali.png',
 'hilda-jemutai-bitok':'Hilda Jemutai Bitok.jpeg','meerab-fatima':'Meerab Fatima.jpeg','malahat-mardani':'Malahat.png',
 'mounes-mardani':'Mounes.png','manar-makki-shaalan':'manar-makki.jpg','abdul-mateen':'abdul-mateen.jpg',
 'ghalib-nadeem':'ghalib-nadeem.jpg','hina-mehboob':'hina-mehboob.jpg','javeria-iqbal':'javeria-iqbal.jpg','bilal-aslam':'Bilal Aslam.jpeg','hasanga-uyanhewage':'Hasanga Uyanhewage.jpeg',
 'sadegh-rajaei':'Sadegh Rajaei.png',
 'shiva-ilkhani-zadeh':'Shiva Ilkhani zadeh.jpeg','ali-hasnain':'Ali Hasnain.jpeg','ramsha-khan':'ramsha.jpeg','alan-liew':'Professor Alan Wee-Chung Liew.jpg','fawad-zaidi':'Syed Fawad.png','mana-mirzaei':'mana mirzai.jpeg','malak-emziane':'Malak EMZIANE.jpeg',
}
# Research-area tags (first pass from stated interests; director to confirm)
AREAS_OF = {
 'rania-shibl':['health'],'mostafa-kamalpour':['health'],'dana-dermody':['home','health','connect'],
 'nazmul-hossain':['connect'],'samia-mujahid':['health'],'nida-ali':['edu'],'svetlana-kolos':['health'],
 'tobias-romano':['agri'],'shiva-jahanaray':['health'],'aittezaz-ahmad':['connect'],
 'sepehr-amooeinejad':['health'],'damilare-ogunjobi':['connect'],'hilda-jemutai-bitok':['connect','mobility'],
 'meerab-fatima':['edu'],'malahat-mardani':['health'],'mounes-mardani':['health'],
 'shahrzad-saremi':['health','home','agri','edu','connect','mobility'],'hassan-ahmed':['edu','connect'],'abdullah-khan':['connect'],'arooj-fatima':['connect'],
 'bisma-ali':['health'],'bilal-aslam':['health'],
}
GROUP_OF = {}; STUDENT_IDS = set()   # students are listed under Researchers & Academics, after the other researchers
for i in ['shahrzad-saremi','rania-shibl','dana-dermody']: GROUP_OF[i] = 'leadership'
for i in ['alan-liew','mostafa-kamalpour','shiva-ilkhani-zadeh','mana-mirzaei','fawad-zaidi']: GROUP_OF[i] = 'advisors'
GROUP_OF['hassan-ahmed'] = 'lead'
for i in ['svetlana-kolos','mohamadali-rezaeimanesh','amir-h-malekijoo','jie-zhu','thiwanka-kaushalya-nagasanga',
          'meerab-fatima','malahat-mardani','mounes-mardani']: GROUP_OF[i] = 'researchers'; STUDENT_IDS.add(i)
GROUPS = [('leadership','Leadership & Founders'),('advisors','Scientific Advisory Board'),('lead','Team Lead'),
          ('researchers','Students & Researchers')]
SHORT_INTEREST = {'Passive sensor monitoring and interpretation of daily activity patterns':'Passive sensor monitoring',
 'Data visualisation for health professionals':'Health data visualisation','Telehealth and wearables':'Telehealth & wearables',
 'Human�computer interaction, usability and technology adoption':'HCI & technology adoption',
 'Human-centred design informed by clinical and caregiving needs':'Human-centred design',
 'Integration of sensor data into clinical workflows and care decisions':'Sensor data in clinical workflows',
 'Privacy, autonomy and ethical design of monitoring systems':'Ethical monitoring design'}

EXTRA_MEMBERS = [
 ('manar-makki-shaalan','Manar Makki Shaalan','Assistant Lecturer in Mathematics · PhD Scholar',['University of Babylon, Babylon'],'Iraq',
  ['Graph theory','Domination theory','Network reliability'],'','https://scholar.google.com/citations?user=T4QkHMsAAAAJ','https://www.linkedin.com/in/manar-makki-shaalan-937371439/'),
 ('abdul-mateen','Abdul Mateen','Lecturer',['Department of Computer Science, National University of Computer and Emerging Sciences (FAST-NUCES), Chiniot-Faisalabad Campus'],'Pakistan',
  [],'0000-0002-8607-7783','https://scholar.google.com/citations?user=CrlrCQkAAAAJ','https://www.linkedin.com/in/ammateen49/'),
 ('ghalib-nadeem','Ghalib Nadeem','Researcher · PhD Scholar',['Department of Computer Science, Huazhong University of Science and Technology, Wuhan'],'China',
  [],'','https://scholar.google.com/citations?user=9-Y_EvAAAAAJ','https://www.linkedin.com/in/ghalib-nadeem-023678189/'),
 ('hina-mehboob','Hina Mehboob','BE Computer Software Engineering',['National University of Sciences and Technology (NUST), Islamabad'],'Pakistan',
  [],'','','https://www.linkedin.com/in/hina-mehboob-nust/'),
 ('javeria-iqbal','Javeria Iqbal','Researcher',['Department of Computer Science, National University of Computer and Emerging Sciences, Islamabad'],'Pakistan',
  [],'0009-0000-5685-5452','','https://www.linkedin.com/in/jave530'),
 ('bilal-aslam','Bilal Aslam','Independent Researcher',[],'',
  ['Machine learning','Deep learning','Healthcare AI','IoMT','Trustworthy AI'],'','','https://www.linkedin.com/in/bilal-aslam-777b6712b'),
 ('hasanga-uyanhewage','Hasanga Uyanhewage','Customer Service Advisor',['Tesco Mobile, Tesco UK'],'United Kingdom',
  ['Software engineering','Human–AI interaction','AI governance','Privacy & data protection','Trust & decision-making'],'0009-0009-3528-4940',
  'https://scholar.google.com/citations?hl=en&user=n7aYL-MAAAAJ','https://www.linkedin.com/in/hasangauyanhewa'),
 ('shiva-ilkhani-zadeh','Shiva Ilkhani Zadeh','Senior Lecturer in Business and Management',['Business School, Business and Law Department, Bournemouth University'],'United Kingdom',
  ['Leadership','Sustainability','AI in tourism','Nudging'],'0000-0002-9362-663X','https://scholar.google.com/citations?user=Fmh5rhIAAAAJ','https://www.linkedin.com/in/shiva-ilkhanizadeh'),
 ('ali-hasnain','Ali Hasnain','Researcher',['Department of Software Engineering, University of Sahiwal'],'Pakistan',
  ['Computer vision','Medical AI','Large language models','Agentic AI','Deep learning'],'0009-0002-6979-4059','https://scholar.google.com/citations?user=lgIdgBcAAAAJ','https://www.linkedin.com/in/ali-hasnain-aa88252a3/'),
 ('malak-emziane','Malak Emziane','Computer Science Engineer · Part-Time Teacher',['University of Tipaza'],'Algeria',
  ['Quantum cryptography','IoT','AI','Secure 5G/6G networks'],'0009-0001-1485-8169','https://scholar.google.com/citations?user=K8km0VoAAAAJ','https://www.linkedin.com/in/malak-emziane-a407681a9'),
 ('ramsha-khan','Ramsha Khan','Researcher · MPhil Virology & Molecular Pathology',['University of Lahore'],'Pakistan',
  ['AI for health','Biomedical research','Molecular medicine','Virology & infectious diseases','AI in healthcare','Bioinformatics'],'0009-0004-6469-300X','https://scholar.google.com/citations?user=GFuPTmYAAAAJ','https://www.linkedin.com/in/ramshakhan13'),
 ('alan-liew','Alan Wee-Chung Liew','Head of School, School of Information and Communication Technology',['School of Information and Communication Technology, Griffith University, Gold Coast'],'Australia',
  ['Artificial intelligence','AI for health','Medical imaging','Multimodal AI','Trustworthy & explainable AI','Graph learning & foundation models','Machine learning','Computer vision','Pattern recognition','Bioinformatics'],'0000-0001-6718-7584','https://scholar.google.com.au/citations?user=CNgJ3LYAAAAJ','https://www.linkedin.com/in/alan-liew-0214a138/'),
 ('fawad-zaidi','Syed Fawad M. Zaidi','Senior Academic',['Torrens University Australia'],'Australia',
  ['Artificial intelligence & intelligent systems','Human-centred design & design thinking','Serious games & immersive learning technologies','Digital health & health informatics','Learning analytics & educational innovation'],'0000-0002-3027-4139','https://scholar.google.com/citations?user=eZ22LtIAAAAJ','https://www.linkedin.com/in/syedfawadmustafazaidi/'),
 ('mana-mirzaei','Mana Mirzaei','Lecturer, School of Business and Creative Industries',['University of the Sunshine Coast, Sunshine Coast, Queensland'],'Australia',
  ['AI','Business data analysis','Machine learning','Optimisation'],'0000-0002-7380-3985','https://scholar.google.com/citations?user=Jec46jcAAAAJ',''),
]
EXTRA_META = {'shiva-ilkhani-zadeh': dict(prefix='Dr.', role='Senior Researcher'),
              'alan-liew': dict(prefix='Professor', role='Academic Advisor', links=[('Griffith Experts','https://experts.griffith.edu.au/7401-alan-weechung-liew'),('Scopus','https://www.scopus.com/authid/detail.uri?authorId=7005648281'),('ResearchGate','https://www.researchgate.net/profile/Alan_Wee_Chung_Liew')]),
              'fawad-zaidi': dict(prefix='Dr.', role='Academic Advisor'), 'mana-mirzaei': dict(prefix='Dr.', role='Academic')}

def build_members():
    out = []
    for r in read_xlsx():
        raw = r['B'].strip()
        m = re.match(r'^(Dr\.|Prof\.)?\s*(.*?)(?:,\s*(.*))?$', raw)
        prefix, name, suffix = (m.group(1) or ''), m.group(2).strip(), (m.group(3) or '')
        alias = ''
        nm = re.search(r'\((.*?)\)', name)
        if nm:
            alias = nm.group(1); name = re.sub(r'\s*\(.*?\)', '', name)
        mid = slug(name)
        if mid == 'gordana-dermody': mid = 'dana-dermody'
        display = f'{alias} Dermody' if alias else name
        pos = r.get('E', '').strip()
        interests = [SHORT_INTEREST.get(x.strip(), x.strip()) for x in r.get('L', '').replace('�', '–').split(';') if x.strip()]
        orcid = r.get('I', '').replace('https://orcid.org/', '').strip()
        level = r.get('D', '').strip(); gpos = r.get('C', '').strip()
        role = {'Founder - Director':'Founder · Director','Cofounder':'Cofounder'}.get(level, '')
        if mid == 'shahrzad-saremi': role = 'Founder · Director'
        elif mid == 'rania-shibl': role = 'Co-Director · Cofounder'
        elif mid == 'dana-dermody': role = 'Cofounder · Senior Researcher'
        elif mid == 'mostafa-kamalpour': role = 'Senior Researcher · Sessional Academic'
        elif mid == 'hassan-ahmed': role = 'Team Lead · Senior Researcher'
        elif mid == 'abdullah-khan': role = 'Senior Researcher'
        elif level == 'Masters Student': role = "Master's Student"
        elif mid == 'hilda-jemutai-bitok': role = 'PhD Student'
        else: role = 'Researcher'
        out.append(dict(
            id=mid, prefix=prefix, name=display if alias else name, full=name, suffix=suffix, alias=alias,
            group=GROUP_OF.get(mid, 'researchers'), role=role, title=pos,
            inst=[x.strip() for x in r.get('F', '').split(';') if x.strip()], country=r['G'].strip(),
            areas=AREAS_OF.get(mid, []), interests=interests, orcid=orcid,
            scholar=r.get('J', '').strip(), linkedin=r.get('K', '').strip(), email=r.get('H', '').strip()))
    for mid, name, title, inst, country, interests, orcid, scholar, linkedin in EXTRA_MEMBERS:
        out.append(dict(id=mid, prefix='', name=name, full=name, suffix='', alias='', group=GROUP_OF.get(mid, 'researchers'), role='Researcher',
            title=title, inst=inst, country=country, areas=AREAS_OF.get(mid, []), interests=interests, orcid=orcid,
            scholar=scholar, linkedin=linkedin, email=''))
        out[-1].update(EXTRA_META.get(mid, {}))
    for m in out:
        if m['id'] == 'rania-shibl': m['prefix'] = 'Professor'
        if m['id'] == 'dana-dermody': m['prefix'] = 'Assoc. Prof.'
        if m['id'] == 'mostafa-kamalpour': m['prefix'] = 'Dr.'; m['name'] = 'Mostafa Kamalpour, PhD'; m['suffix'] = ''
    first = ['shahrzad-saremi', 'rania-shibl', 'dana-dermody', 'alan-liew', 'mostafa-kamalpour', 'shiva-ilkhani-zadeh', 'mana-mirzaei', 'fawad-zaidi', 'hassan-ahmed']
    order = [g for g, _ in GROUPS]
    out.sort(key=lambda m: (order.index(m['group']), m['id'] in STUDENT_IDS, first.index(m['id']) if m['id'] in first else len(first)))
    return out

# face-centred crops (left, top, size) in source pixels
CROPS = {'shahrzad-saremi': (390, 0, 680), 'rania-shibl': (306, 50, 640), 'dana-dermody': (95, 20, 680),
         'hassan-ahmed': (250, 60, 820), 'mostafa-kamalpour': (10, 0, 490)}

def process_photos(members):
    d = os.path.join(OUT, 'images', 'team'); os.makedirs(d, exist_ok=True)
    src = os.path.join(ROOT, 'images', 'members')
    for m in members:
        f = PHOTOS.get(m['id'])
        if not f or not os.path.exists(os.path.join(src, f)):
            m['photo'] = ''; m['pv'] = ''; continue
        im = ImageOps.exif_transpose(Image.open(os.path.join(src, f))).convert('RGB')
        w, h = im.size; s = min(w, h)
        if m['id'] in CROPS:
            l, t, cs = CROPS[m['id']]
            im = im.crop((l, t, l + cs, t + cs)).resize((360, 360), Image.LANCZOS)
            im.save(os.path.join(d, m['id'] + '.jpg'), quality=82, optimize=True, progressive=True)
            m['photo'] = m['id'] + '.jpg'; m['pv'] = hashlib.md5(open(os.path.join(d, m['photo']), 'rb').read()).hexdigest()[:8]; continue
        left = (w - s) // 2
        top = 0 if h > w else 0            # faces sit in the upper part of portrait shots
        if h > w: top = int((h - s) * 0.12)
        im = im.crop((left, top, left + s, top + s)).resize((360, 360), Image.LANCZOS)
        im.save(os.path.join(d, m['id'] + '.jpg'), quality=82, optimize=True, progressive=True)
        m['photo'] = m['id'] + '.jpg'; m['pv'] = hashlib.md5(open(os.path.join(d, m['photo']), 'rb').read()).hexdigest()[:8]

# ───────────────────────── publications ─────────────────────────
# (id, type, year, authors, title, venue, doi, areas, note)
PUBS = [
 ('egenai-dbr','journal',2026,'Saremi, S., Mirzaei, M., Rasti, A., Nooraei Abadeh, M., Varposhti, M., Shibl, R., Ahmed, H., Aldakheel, S. K. A., Mealy, E., Wang, K., Humphreys, D.','EGenAI-DBR: A design-based framework for responsible generative AI integration in higher education','Education Innovations: Systems and Future Learning, 1(2)','',['edu'],''),
 ('coi-online-peer','journal',2026,'Dokhanchi, M., Saremi, S., Shibl, R., Heidari, M., Ahmed, H., Mansoor, D., … Mirzaei, M.','An extended community of inquiry framework for monitoring and predicting online peer learning participation','Journal of Applied Research in Higher Education, 18(8), 113–141','',['edu'],'Q2'),
 ('selm-ctr','journal',2026,'Ali, Z., Ahmed, H., Khan, A., Saremi, S., Shibl, R., Mirzaei, M., Rastegari, P., Wang, M.','SELM-CTR: A stacking ensemble deep learning model with SHAP-based analysis for large-scale click-through rate prediction','International Journal of Intelligent Computing and Cybernetics, 19(3)','',['connect'],'Q2'),
 ('latent-squeeze','journal',2026,'Mirzaei, M., Saremi, S., Khan, A., Ahmed, H., Rastegari, P., Nooraei Abadeh, M., Shibl, R., Varposhti, M., Nguyen, T. T.','Latent-space feature squeezing for adversarially robust intrusion detection on IoT edge devices','International Journal of Intelligent Computing and Cybernetics, ahead of print, 1–37','10.1108/IJICC-06-2026-0590',['connect'],''),
 ('iot-certless','journal',2026,'Dadkhah, P., Rastegari, P., Dakhilalian, M., Yeoh, P., Wang, M., Saremi, S., … Gupta, B. B.','An IoT-aware certificateless signature scheme for protection against type-I and type-II super adversaries','IoT, 7(2)','',['connect'],'Q1'),
 ('smep-iov','journal',2026,'Ghadim, G. O., Rastegari, P., Dakhilalian, M., Hendessi, F., Saremi, S., Shibl, R., … Nguyen, T. T.','Cryptanalysis and improvement of the SMEP-IoV protocol: A secure and lightweight protocol for message exchange in IoV paradigm','IoT, 7(2), 31','',['mobility','connect'],'Q1'),
 ('hwsn-signcryption','journal',2026,'Dadkhah, P., Rastegari, P., Dakhilalian, M., Yeoh, P., Wang, M., Saremi, S., Gupta, B. B.','Protecting HWSNs from super adversaries with robust certificateless signcryption','Telecom, 7(2), 37','',['connect'],'Q2'),
 ('maps-pbft','journal',2026,'Bitok, H. J., Wang, M., Desmond, D.','MAPS-PBFT: Mobility-aware persistent-survivor PBFT for vehicular networks','Concurrency and Computation: Practice and Experience, 38(16)','',['mobility'],'Q2'),
 ('agri-nigeria','journal',2026,'Ogunjobi, D., Shibl, R., Saremi, S.','Temporal feature engineering for agricultural commodity price forecasting in Nigeria: Evaluating machine learning, deep learning and time-series approaches','Journal of Agribusiness in Developing and Emerging Economies, 1–28','',['agri'],''),
 ('crisis-hybrid','journal',2026,'Ahmed, H., Khan, A., Fatima, A., Mateen, A., Saremi, S., Shibl, R., Rajaei, S., Mirzaei, M.','Crisis-induced hybrid learning, cognitive offloading, and generative AI reliance among Pakistani CS undergraduates','Education Innovations: Systems and Future Learning, 1(1), 568–589','10.1108/EISFL-06-2026-0098',['edu'],''),
 ('health-smart-home','journal',2026,'Dermody, G., Shibl, R., Wang, M., Ward, A., Watson, J., North, K., Blake, J., …','Multi perspective considerations for health smart home: Early phase exploratory study','Journal of Advanced Nursing, 82(1), 849–866','',['home','health'],''),
 ('v2i-fahp','journal',2026,'Suleman, D., Shibl, R., Wang, M., Ansari, K.','A triangular hybrid FAHP-FTOPSIS and contextual bandit dynamic protocol selection for V2I communications','Vehicular Communications, 101081','',['mobility'],''),
 ('v2i-multichannel','journal',2026,'Suleman, D., Shibl, R., Wang, M., Ansari, K.','A multi-channel architecture for concurrent safety and non-safety V2I communications','IEEE Open Journal of Intelligent Transportation Systems','',['mobility'],''),
 ('bridging-divide','journal',2026,'Dermody, G., Wadsworth, D., El Haddad, M., Prichard, R., Benson, A., Benson, T., …','Bridging the digital divide: A multi-method evaluation of nursing readiness for digital health technology','Journal of Advanced Nursing, 82(4), 3752–3766','',['health'],''),
 ('ai-start-ups','journal',2025,'Azizi, N., Akhavan, P., Davison, C., Haass, O., Saremi, S., Zaidi, S. F. M.','AI-driven process innovation: Transforming service start-ups in the digital age','Electronics, 14(16), 3240','10.3390/electronics14163240',[],'Q4'),
 ('quic-sctp','journal',2025,'Suleman, D., Shibl, R., Wang, M., Ansari, K.','A simulated study of IoT ALPs over legacy TCP/UDP versus QUIC and SCTP for V2I communications','IEEE Transactions on Intelligent Transportation Systems','',['mobility','connect'],''),
 ('iot-taxonomy','journal',2025,'Suleman, D., Shibl, R., Ansari, K., Yakoi, P. S.','A taxonomy proposal of information assurance and data quality solutions in smart cities','Franklin Open, 100436','',['connect'],''),
 ('gerotech','journal',2025,'Fritz, R. L., Dermody, G., Masaki, H., Greiner, C., Atakro, C. A., Mohammadi, M., …','A Global Gerontechnology Center for Nursing Science for improving aging-in-place technologies','Nursing Outlook, 73(6), 102590','',['home','health'],''),
 ('smart-home-data','journal',2025,'Dermody, G., Cook, D. J., Fritz, R. L.','Interpretation of health-smart home data and implications for clinical decision-making: Inductive content analysis','JMIR Nursing, 8(1), e75234','',['home','health'],''),
 ('iot-data-quality','journal',2023,'Suleman, D., Shibl, R., Ansari, K.','Investigation of data quality assurance across IoT protocol stack for V2I interactions','Smart Cities, 6(5), 2680–2705','',['mobility','connect'],''),
 ('digital-health-dq','journal',2023,'Syed, R., Eden, R., Makasi, T., Chukwudi, I., Mamudu, A., Kamalpour, M., …','Digital health data quality issues: Systematic review','Journal of Medical Internet Research, 25, e42615','',['health'],''),
 ('smart-home-readiness','journal',2021,'Dermody, G., Fritz, R., Glass, C., Dunham, M., Whitehead, L.','Factors influencing community-dwelling older adults’ readiness to adopt smart home technology: A qualitative exploratory study','Journal of Advanced Nursing, 77(12), 4847–4861','',['home','health'],''),
 ('online-communities','journal',2020,'Kamalpour, M., Watson, J., Buys, L.','How can online communities support resilience factors among older adults','International Journal of Human-Computer Interaction, 36(14), 1342–1353','',['health'],''),
 ('salp','journal',2017,'Mirjalili, S., Gandomi, A. H., Mirjalili, S. Z., Saremi, S., Faris, H., Mirjalili, S. M.','Salp swarm algorithm: A bio-inspired optimizer for engineering design problems','Advances in Engineering Software, 114, 163–191','10.1016/j.advengsoft.2017.07.002',[],'5,900+ citations'),
 ('goa','journal',2017,'Saremi, S., Mirjalili, S., Lewis, A.','Grasshopper optimisation algorithm: Theory and application','Advances in Engineering Software, 105, 30–47','10.1016/j.advengsoft.2017.01.004',[],'3,300+ citations'),
 ('mogwo','journal',2016,'Mirjalili, S., Saremi, S., Mirjalili, S. M., Coelho, L. S.','Multi-objective grey wolf optimizer: A novel algorithm for multi-criterion optimization','Expert Systems with Applications, 47, 106–119','10.1016/j.eswa.2015.10.039',[],'2,100+ citations'),
 ('acis-ai-discussion','conference',2025,'Dokhanchi, M., Saremi, S., Mealy, E., Shibl, R.','Integrating AI in online discussion tools: A critical perspective','Proceedings of the Australasian Conference on Information Systems (ACIS 2025)','',['edu'],''),
 ('acis-research-impact','conference',2025,'Kamalpour, M., Nili, A., Yigitcanlar, T.','Assessing research impact: A framework and implications for management information systems research','Australasian Conference on Information Systems (ACIS 2025)','',[],''),
 ('acis-ai-communities','conference',2025,'Soroush, L., Bekrian, E., Kamalpour, M.','A systematic literature review on how AI shapes communication in online health communities','Australasian Conference on Information Systems (ACIS 2025)','',['health'],''),
 ('aag-smart-home','conference',2025,'Dermody, D., Shibl, R., Wang, M., Ward, A., Watson, J., Mealy, E., Yeoh, P. L., …','What will it take? Triangulating stakeholder perspectives on smart home technology for ageing in place','Proceedings of the 58th Australian Association of Gerontology Conference','',['home'],''),
 ('sports-injury','conference',2026,'Kamalpour, M., Saremi, S., Shibl, R., Mirzaei, M., Bedford, A., Mealy, E.','Mapping artificial intelligence and machine learning research in sports injury prediction: A bibliometric analysis','18th Australasian Conference on Mathematics and Computers in Sport','',['health'],''),
 ('sports-concussion','conference',2026,'Ellethy, H., Kamalpour, M., Shibl, R., Saremi, S., Mirzaei, M.','Three decades of AI in sports-related concussion research: A bibliometric analysis (1996–2026)','18th Australasian Conference on Mathematics and Computers in Sport','',['health'],''),
 ('iomt-ids','preprint',2026,'Ahmed, H., Dermody, G., Mateen, A., Saremi, S., Shibl, R.','Intrusion detection for the Internet of Medical Things: A cost-aware comparative benchmark of machine learning approaches','Under review','',['health','connect'],''),
 ('hand-book','book',2019,'Saremi, S., Mirjalili, S.','Optimization algorithms for hand posture estimation','Algorithms for Intelligent Systems, Springer (book)','',[],''),
 ('nature-inspired','chapter',2020,'Saremi, S., Mirjalili, S., Mirjalili, S. Z., Dong, J.','Grasshopper optimization algorithm: Theory, literature review, and application in hand posture estimation','In Nature-Inspired Optimizers (pp. 107–122), Springer','',[],''),
]
_x = os.path.join(ROOT, 'pubs_extra.json')
if os.path.exists(_x) and not os.environ.get('SYNTERA_NO_EXTRA'):
    PUBS += [tuple(r) for r in json.load(open(_x, encoding='utf-8'))]

MEMBER_KEYS = {}   # (surname, initial) -> member id ; filled in main()
def build_member_keys(members):
    for m in members:
        parts = m['full'].split()
        sur = parts[-1].lower()
        MEMBER_KEYS[(sur, parts[0][0].lower())] = m['id']
        if m['id'] == 'dana-dermody': MEMBER_KEYS[(sur, '*')] = m['id']   # one source lists her as "Dermody, D."

ALL_AREAS = {'shahrzad-saremi', 'rania-shibl', 'alan-liew'}   # work across every research area
CORE_TOO = {'alan-liew'}                                    # also listed under Core AI & Methods
INTEREST_RULES = [
    ('health', r'health|medical|clinical|biomedical|disease|virolog|molecular|bioinformat|iomt|telehealth|wearable|sports injury|nurs|patient|eeg|neuro|genom|concussion'),
    ('home', r'smart home|ageing|aging|assistive|ambient|elder|passive sensor|aged care'),
    ('agri', r'agricultur|crop|farm'),
    ('edu', r'educat|teaching|student|pedagog|learning analytics|serious game|immersive learning|computer-assisted|linguistic|tutor|e-learning|career|skill'),
    ('connect', r'\biot\b|internet of things|iomt|cyber|security|privacy|blockchain|networking|network (?:security|reliab)|sensor network|wireless|5g|6g|cryptograph|intrusion|reliability engineering'),
    ('mobility', r'vehic|\biov\b|traffic|transport|v2x|v2i')]

def derive_areas(members):
    """Research areas of each person, from evidence: (1) their stated interests, (2) the areas of papers they co-authored
    (including paper keywords from pub_meta.json). A paper-based area needs 3+ papers, or 2+ papers making up 40%+ of their
    output, or (for members who listed no interests) half of their papers. Saremi and Shibl work across every area."""
    order = [a['id'] for a in AREAS]
    mp = os.path.join(ROOT, 'pub_meta.json')
    meta = json.load(open(mp, encoding='utf-8')) if os.path.exists(mp) else {}
    ev = {m['id']: {} for m in members}; npub = {m['id']: 0 for m in members}
    for p in PUBS:
        kws = ' '.join(meta.get(p[0], {}).get('keywords', [])).lower()
        areas = set(p[7]) | {a for a, rx in INTEREST_RULES if kws and re.search(rx, kws)}
        for _, mid in parse_authors(p[3]):
            if mid in ev:
                npub[mid] += 1
                for a in areas: ev[mid][a] = ev[mid].get(a, 0) + 1
    for m in members:
        if m['id'] in ALL_AREAS: m['areas'] = list(order); m['core'] = m['id'] in CORE_TOO; continue
        t = ' '.join(m['interests']).lower()
        got = {a for a, rx in INTEREST_RULES if re.search(rx, t)}
        n = npub[m['id']]
        for a, c in ev[m['id']].items():
            if c >= 3 or (c >= 2 and c / n >= .4) or (not m['interests'] and c / n >= .5): got.add(a)
        m['areas'] = [a for a in order if a in got]
        m['core'] = not m['areas']

def parse_authors(s):
    out = []
    for part in re.findall(r'[^,]+,\s*(?:[A-Z]\.\s?)+|…', s):
        t = part.strip()
        mid = None
        if ',' in t:
            sur, ini = [x.strip() for x in t.split(',', 1)]
            mid = MEMBER_KEYS.get((sur.lower(), ini[0].lower())) or MEMBER_KEYS.get((sur.lower(), '*'))
        out.append([t, mid])
    return out

def export_pubs():
    items = []
    mp = os.path.join(ROOT, 'pub_meta.json')
    meta = json.load(open(mp, encoding='utf-8')) if os.path.exists(mp) else {}
    for pid, typ, year, authors, title, venue, doi, areas, note in sorted(PUBS, key=lambda x: (-x[2], x[4])):
        mt = meta.get(pid, {})
        items.append(dict(id=pid, type=typ, year=year, authors=parse_authors(authors), title=title, venue=venue,
                          doi=doi or mt.get('doi', ''), areas=areas, note=note,
                          keywords=mt.get('keywords', []), abstract=mt.get('abstract', '')))
    areas = {a['id']: dict(short=a['name'], color=a['color']) for a in AREAS}
    return 'window.SYNTERA_PUBS = %s;\nwindow.SYNTERA_AREAS = %s;\n' % (json.dumps(items, ensure_ascii=False, indent=1), json.dumps(areas))

# ───────────────────────── page chrome ─────────────────────────
NAV = [('research.html','Research','research'),('publications.html','Publications','publications'),
       ('people.html','People','people'),('collaborate.html','Collaborate','collaborate'),('about.html','About','about'),('contact.html','Contact','contact')]
LOGO = ('<svg viewBox="0 0 40 40" aria-hidden="true"><defs><linearGradient id="lg" x1="0" y1="0" x2="1" y2="1">'
        '<stop offset="0" stop-color="#E83E8C"/><stop offset="1" stop-color="#3A7BFF"/></linearGradient></defs>'
        '<rect width="40" height="40" rx="10" fill="#0F1A47"/>'
        '<path d="M28 12c-2-3-14-3-14 3 0 7 14 3 14 10 0 6-12 6-15 2" fill="none" stroke="url(#lg)" stroke-width="3" stroke-linecap="round"/>'
        '<circle cx="28" cy="12" r="3" fill="#FF8FC3"/><circle cx="13" cy="27" r="3" fill="#8FB4FF"/><circle cx="20" cy="20" r="2" fill="#fff"/></svg>')

import hashlib
VER = ''   # set in main() from asset contents


def clean_links(doc, fname):
    """Point internal <a href> at extensionless URLs (what Cloudflare Pages serves) so crawlers and visitors skip the .html redirect."""
    d = posixpath.dirname(fname)
    def fix(mo):
        h = mo.group(2)
        if re.match(r'^(?:[a-zA-Z][a-zA-Z0-9+.-]*:|/|#|\?)', h): return mo.group(0)
        mm = re.match(r'([^#?]*)([#?].*)?$', h)
        path, rest = mm.group(1), mm.group(2) or ''
        if not path.endswith('.html'): return mo.group(0)
        t = posixpath.normpath(posixpath.join(d, path))
        t = '/' if t == 'index.html' else '/' + t[:-5]
        return f'{mo.group(1)}{t}{rest}"'
    return re.sub(r'(<a\b[^>]*?\bhref=")([^"]+)"', fix, doc)

def ld_tag(obj):
    return '<script type="application/ld+json">' + json.dumps(obj, ensure_ascii=False) + '</script>\n'

def crumbs_ld(fname, title):
    items = [('Home', SITE_URL + '/')]
    if fname.startswith('research/'): items.append(('Research', SITE_URL + '/research'))
    elif fname.startswith('people/'): items.append(('People', SITE_URL + '/people'))
    items.append((title, SITE_URL + '/' + re.sub(r'\.html$', '', fname)))
    return ld_tag({"@context": "https://schema.org", "@type": "BreadcrumbList", "itemListElement": [
        {"@type": "ListItem", "position": i + 1, "name": n, "item": u} for i, (n, u) in enumerate(items)]})

_PUBMETA = None
def pub_doi(p):
    global _PUBMETA
    if _PUBMETA is None:
        mp = os.path.join(ROOT, 'pub_meta.json')
        _PUBMETA = json.load(open(mp, encoding='utf-8')) if os.path.exists(mp) else {}
    return p[6] or _PUBMETA.get(p[0], {}).get('doi', '')

PUB_TYPES = {'journal':'Journal','conference':'Conference','book':'Book','chapter':'Chapter','thesis':'Thesis','preprint':'Preprint','other':'Other'}
def static_pubs(pred=None, limit=None):
    """Server-rendered publication list (JS replaces it with the interactive version) so search engines can read every paper."""
    rows = sorted(PUBS, key=lambda x: (-x[2], x[4]))
    if pred: rows = [p for p in rows if pred(p)]
    if limit: rows = rows[:limit]
    out = []
    for p in rows:
        au = '; '.join(f'<a class="au" href="/people/{mid}"><strong>{E(n)}</strong></a>' if mid else E(n) for n, mid in parse_authors(p[3]))
        doi = pub_doi(p)
        act = f'<div class="pub__act"><a class="pbtn" href="https://doi.org/{E(doi)}" rel="noopener" target="_blank">DOI<span class="doi"> {E(doi)}</span></a></div>' if doi else ''
        out.append(f'<li class="pub" id="pub-{p[0]}"><div class="pub__meta"><span class="pill pill--{p[1]}">{PUB_TYPES[p[1]]}</span><span class="pub__year">{p[2] or "n.d."}</span></div>'
                   f'<h3 class="pub__title">{E(p[4])}</h3><p class="pub__au">{au}</p><p class="pub__venue"><em>{E(p[5])}</em></p>{act}</li>')
    return ''.join(out)

def page(fname, title, desc, body, active='', depth=0, extra_js='', home=False, pubs_on=False, ld_extra='', full_title=None):
    p = '../' * depth
    cur = ' aria-current="page"'
    nav = ''.join(f'<li><a href="{p}{h}"{cur if k == active else ""}>{t}</a></li>' for h, t, k in NAV)
    clean = 'index.html' if fname == '404.html' else fname
    curl = SITE_URL + '/' + ('' if clean == 'index.html' else re.sub(r'\.html$', '', clean))
    robots = '<base href="/">\n' if fname == '404.html' else ''
    ld = ''
    if home:
        ld = ld_tag({"@context": "https://schema.org", "@graph": [
            {"@type": ["ResearchOrganization", "Organization"], "@id": SITE_URL + "/#org", "name": BRAND['name'], "alternateName": "SYNTERA", "url": SITE_URL + '/',
             "logo": {"@type": "ImageObject", "url": SITE_URL + '/images/og-image.png'}, "image": SITE_URL + '/images/og-image.png',
             "description": "Applied AI and connected systems research group open to academic collaboration, funded research projects, industry partnerships and commissioned research.",
             "email": BRAND['email_collab'], "slogan": BRAND['tagline'],
             "address": {"@type": "PostalAddress", "addressRegion": "Queensland", "addressCountry": "AU"},
             "areaServed": "Worldwide",
             "contactPoint": [{"@type": "ContactPoint", "contactType": "research collaboration and partnerships", "email": BRAND['email_collab'], "availableLanguage": "English"},
                              {"@type": "ContactPoint", "contactType": "prospective students", "email": BRAND['email_student'], "availableLanguage": "English"}],
             "parentOrganization": {"@type": "CollegeOrUniversity", "name": "University of the Sunshine Coast", "url": BRAND['host_url']},
             "knowsAbout": KNOWS_ABOUT, "sameAs": [BRAND['scholar']]},
            {"@type": "WebSite", "@id": SITE_URL + "/#site", "url": SITE_URL + '/', "name": BRAND['name'], "inLanguage": "en-AU", "publisher": {"@id": SITE_URL + "/#org"}}]})
    elif fname != '404.html':
        ld = crumbs_ld(fname, title)
    ld += ld_extra
    if fname != '404.html': PAGES.append(fname)
    full_title = full_title or (BRAND['full'] if home else f'{title} · {BRAND["name"]}')
    areas_f = ''.join(f'<li><a href="{p}research/{a["id"]}.html">{a["short"]}</a></li>' for a in AREAS)
    data_js = f'<script src="{p}data/publications.js?v={VER}"></script><script src="{p}js/pubs.js?v={VER}"></script>' if pubs_on else ''
    doc = f'''<!doctype html>
<html lang="en-AU">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{E(full_title)}</title>
<meta name="description" content="{E(desc)}">
<meta name="theme-color" content="#0F1A47">
<meta name="robots" content="{'noindex' if fname == '404.html' else 'index,follow,max-image-preview:large,max-snippet:-1,max-video-preview:-1'}">
<meta name="author" content="{BRAND['name']}">
<link rel="canonical" href="{curl}">
<meta property="og:title" content="{E(full_title)}"><meta property="og:description" content="{E(desc)}"><meta property="og:type" content="website"><meta property="og:locale" content="en_AU">
<meta property="og:url" content="{curl}"><meta property="og:site_name" content="{BRAND['name']}"><meta property="og:image" content="{SITE_URL}/images/og-image.png"><meta property="og:image:width" content="1200"><meta property="og:image:height" content="630"><meta property="og:image:alt" content="SYNTERA Research Group: Applied AI and Connected Systems">
<meta name="twitter:card" content="summary_large_image"><meta name="twitter:title" content="{E(full_title)}"><meta name="twitter:description" content="{E(desc)}"><meta name="twitter:image" content="{SITE_URL}/images/og-image.png">
{robots}{ld}
<link rel="icon" href="{p}images/favicon.svg" type="image/svg+xml">
<link rel="stylesheet" href="{p}css/style.css?v={VER}">
<script>try{{var t=localStorage.getItem('syntera-theme');if(t)document.documentElement.dataset.theme=t}}catch(e){{}}</script>
</head>
<body data-base="{p}">
<a class="skip" href="#main">Skip to content</a>
<header class="hdr"><div class="wrap hdr__in">
  <a class="brand" href="{p}index.html" title="{BRAND['name']}">{LOGO}<span><b>SYNTERA</b><small>Research Group</small></span></a>
  <nav aria-label="Main"><ul id="nav" class="nav">{nav}<li class="nav__cta"><a class="btn btn--pink" href="{p}join.html">Join Us</a></li></ul></nav>
  <div class="hdr__act">
    <button class="icon-btn" id="theme" aria-label="Toggle dark mode" title="Toggle dark mode"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M21 13A9 9 0 1 1 11 3a7 7 0 0 0 10 10z"/></svg></button>
    <button class="icon-btn burger" id="burger" aria-label="Menu" aria-expanded="false" aria-controls="nav"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M4 7h16M4 12h16M4 17h16"/></svg></button>
  </div>
</div></header>
<main id="main">
{body}
</main>
<footer class="ftr"><div class="wrap">
  <div class="ftr__grid">
    <div><a class="brand brand--ftr" href="{p}index.html">{LOGO}<span><b>SYNTERA</b><small>Research Group</small></span></a>
      <p>{BRAND['tagline']}. An applied AI and connected systems research group open to academic collaboration, funded projects and industry partnerships.</p></div>
    <div><h4>Research</h4><ul>{areas_f}</ul></div>
    <div><h4>The group</h4><ul><li><a href="{p}about.html">About</a></li><li><a href="{p}people.html">People</a></li><li><a href="{p}publications.html">Publications</a></li><li><a href="{p}collaborate.html">Collaborate with us</a></li><li><a href="{p}join.html">Join us</a></li><li><a href="{p}privacy.html">Privacy</a></li></ul></div>
    <div><h4>Get in touch</h4><ul><li><a href="mailto:{BRAND['email_student']}">{BRAND['email_student']}</a></li><li><a href="mailto:{BRAND['email_join']}">{BRAND['email_join']}</a></li><li><a href="mailto:{BRAND['email_collab']}">{BRAND['email_collab']}</a></li><li><a href="{BRAND['scholar']}" rel="noopener" target="_blank">Director on Google Scholar</a></li></ul><a class="btn btn--pink btn--sm" href="{p}join.html">Join Us</a></div>
  </div>
  <div class="ftr__bar"><span>© 2026 {BRAND['name']}</span><span><a href="{p}privacy.html">Privacy</a></span></div>
</div></footer>
<script src="{p}js/site.js?v={VER}"></script>
{data_js}
{extra_js}
</body></html>'''
    path = os.path.join(OUT, fname); os.makedirs(os.path.dirname(path), exist_ok=True)
    open(path, 'w', encoding='utf-8').write(clean_links(doc, fname))

def flag(m, p=''):
    f = FLAGS.get(m['country'])
    return f'<img class="flag" src="{p}images/flags/{f}.png" alt="" width="20" height="14">' if f else ''

def avatar(m, p='', cls='av'):
    if m['photo']:
        return f'<img class="{cls}" src="{p}images/team/{m["photo"]}?v={m["pv"]}" alt="{E(m["name"])}" width="120" height="120" loading="lazy">'
    ini = ''.join(w[0] for w in m['full'].split()[:2]).upper()
    return f'<span class="{cls} {cls}--ini" aria-hidden="true">{ini}</span>'


DETAILS = {
 'shahrzad-saremi': dict(
  quals='',
  title='Lecturer, ICT and Computer Science',
  roles=['Program Coordinator: Bachelor of Information and Communications Technology'],
  bio=['Dr Shahrzad Saremi is a researcher and academic with over a decade of experience in computing and information technology. She has published more than 20 high-impact journal articles, attracting over 15,000 citations, reflecting the significant international reach of her research contributions.',
       'Her research spans multiple interdisciplinary domains, including bio-inspired optimisation algorithms, human–computer interaction (HCI), machine learning, Internet of Things (IoT), Internet of Vehicles (IoV), and Smart Health Home systems. She is widely recognised for her co-development of nature-inspired metaheuristic algorithms, including the Grasshopper Optimisation Algorithm and Salp Swarm Algorithm, which are extensively applied to complex engineering and computational challenges. Her work in HCI investigates user experience, gesture recognition, augmented reality, and creative technologies, with a particular focus on design education and knowledge management in organisational settings.',
       'She also explores knowledge-sharing behaviours in organisational and educational settings, examining the interplay between motivation and culture to enhance learning and risk management outcomes. Her research combines technical innovation with human-centred inquiry.',
       'Dr Saremi is deeply committed to fostering engaging, student-centred learning environments and is passionate about mentoring the next generation of researchers.'],
  supervision='Dr Saremi actively supervises Higher Degree by Research (HDR) candidates and welcomes enquiries from motivated researchers whose interests intersect with her areas of expertise. Prospective domestic and international candidates are encouraged to reach out to explore potential research directions.',
  grants=['<b>LAUNCH Partnership Grant (2026)</b>, University of the Sunshine Coast. Lead Investigator, $29,357. Project: “Rapid MRI-Based Knee Segmentation: Leveraging Deep Learning for Patient-Specific 3D Models and Clinical Translation”.'],
  teaching=['Programming','Computer Science','Cybersecurity','Human-Computer Interaction','System Analysis','Math','Requirement Engineering','Computer Organization and Operating Systems','Device and Network Security','Information Systems','Business Intelligence','AI'],
  awards=['Fellow of the Higher Education Academy (FHEA), 2026: Advance HE recognition of professional practice and excellence in higher education',
          'Selected for the 2026 Essence of Research Leadership Program, University of the Sunshine Coast',
          'Vice Chancellor’s Learning and Teaching Award, 2024',
          'Top 2 percent scientist in the Stanford/Elsevier list, 2023',
          'Pro Vice-Chancellor Research Commendation, 2021',
          'PSH (Professor Susan Holland) Casual Academic Performance Award, 2020',
          'Interviewed by Vogue as an active woman in AI (teaching and research), 2020',
          'IIIS Research Impact Awards, Griffith University, 2016',
          'Full scholarship for her PhD, Griffith University, 2014',
          'Her master’s thesis project (Marker Puzzle) was selected for display at the School of Information Technology and Electronic Innovation showcase at the University of Queensland'],
  education=['<b>PhD, Computer Science</b>, Griffith University, Brisbane (2014–2018). Thesis: <i>Evolutionary Hand Posture Estimation for Image-based Gesture Detection Systems</i>',
             '<b>Master of Science, Interaction Design</b>, University of Queensland (2012–2014). Thesis: <i>Marker Puzzle: A novel Augmented Reality framework for learning grammar in primary school</i>',
             '<b>Bachelor of Science, Information Technology</b>, Multimedia University, Cyberjaya, Malaysia (2008–2011). Thesis: <i>Using Augmented Reality in Calendar</i>',
             '<b>Graduate Certificate in Teaching and Learning</b>, Torrens University Australia (2022–2023)'],
  expertise=[('Human-Computer Interaction (HCI)', 'gesture detection, hand posture estimation, augmented reality frameworks for learning, user interface design and human-centred interaction'),
             ('Artificial Intelligence and Machine Learning', 'deep learning, optimization algorithms, multi-objective optimization, neural networks, and AI applications in education and enterprise systems'),
             ('Information Technology Education', 'curriculum design, teaching and learning methodologies, IT professional practice, cybersecurity education and data analytics training'),
             ('Optimization and Metaheuristics', 'evolutionary algorithms, nature-inspired optimization (grasshopper, grey wolf, particle swarm), hand shape optimization and multi-objective problem solving'),
             ('Emerging Technologies', 'blockchain, Internet of Things (IoT), data visualization and augmented reality, with applications in education and enterprise')],
  experience=[('2025–present', 'Program Coordinator, Bachelor of Information and Communication Technology, School of Science, Technology and Engineering, University of the Sunshine Coast'),
              ('2025', 'Lecturer and Course Coordinator, Computer Organization and Operating System, University of the Sunshine Coast'),
              ('2024–2025', 'Lecturer, Holmes Institute'),
              ('2019–2025', 'Learning Facilitator, Torrens University'),
              ('2019–present', 'Adjunct Research Fellow, Griffith University'),
              ('2014–2019', 'Lecturer and Research Assistant, Griffith College, Griffith University')],
  students=['Higher Degree by Research supervision at PhD and Master by Research level, in AI applications, health informatics and information systems, and Internet of Vehicles privacy.',
            '<a href="hilda-jemutai-bitok.html">Hilda Jemutai Bitok</a> (PhD): <i>Preserving Privacy of Sensitive Information in Resource Constrained Internet of Vehicles Environment</i>'],
  skills=[('Programming', 'Python, C++, MATLAB, Java, SQL'),
          ('Machine learning and AI', 'TensorFlow, PyTorch, Keras, Scikit-learn, OpenCV, NLTK'),
          ('Data analytics', 'Tableau, Power BI, Excel, R, SPSS, statistical modelling'),
          ('Web development', 'HTML5, CSS3, JavaScript, PHP, WordPress'),
          ('Design tools', 'Adobe Creative Suite, Canva, Figma, Edraw Max, MS Visio'),
          ('Development environments', 'Visual Studio Code, PyCharm, Jupyter Notebook, GitHub, GitLab, Docker'),
          ('Databases', 'MySQL, PostgreSQL, MongoDB, Oracle'),
          ('Cloud platforms', 'Google Cloud, AWS, Microsoft Azure'),
          ('Augmented reality', 'ARKit, ARCore, Unity, Unreal Engine'),
          ('Academic tools', 'LaTeX, Mendeley, EndNote, Zotero, Covidence')],
  courses=[('University of the Sunshine Coast', ['Computer Organization and Operating System (Course Coordinator, 2025)']),
           ('Holmes Institute', ['IS Governance and Risk', 'Database Design', 'System Analysis and Design', 'Computer Forensics', 'Leveraging IT Advantages for Managers', 'Professional Issues in IS Ethics and Practice']),
           ('Torrens University', ['Creative Enterprises', 'IT Professional Practice', 'Data and Networking', 'Introduction to Programming', 'Cybersecurity', 'Secure by Design', 'Human Centred by Design', 'Big Data and Analytics', 'Deep Learning', 'Mathematical Foundation of AI', 'Cloud Computing', 'Microservices', 'Requirement Engineering', 'Data Modelling and Database Design']),
           ('Griffith College', ['Computer Skills', 'Information System Foundations', 'Human Computer Interaction', 'Essential Mathematics', 'Information Design', 'Introduction to Computing', 'Foundation of Computing Systems', 'Digital Technologies'])],
  languages='English, Persian (Farsi) and Azeri (native or bilingual); Turkish (professional working); Arabic (elementary)',
  service='Reviewer for international conferences and journals in computer science and education. Adjunct Research Fellow, Griffith University (2019–present).',
  metrics='15,573 citations (12,216 since 2021) · h-index 21 (18 since 2021) · i10-index 22 (21 since 2021) · Google Scholar, 6 Oct 2026'),
}

DETAILS['alan-liew'] = dict(
  title='Head of School, School of Information and Communication Technology, Griffith University',
  roles=['Professor, Griffith University · Founder and Lead, AI4Health Lab · Co-founder and Co-lead, TrustAGI Lab'],
  bio=['Professor Alan Liew is Head of the School of Information and Communication Technology at Griffith University and an internationally recognised researcher in artificial intelligence (AI), machine learning, medical imaging, computer vision and bioinformatics. His research focuses on advanced, trustworthy and translational AI methods for complex real-world problems, particularly in health and biomedical applications.',
       'At Griffith he founded the AI4Health Lab, which he leads, and co-founded the TrustAGI Lab, which he co-leads. His current research spans AI for health and medical imaging, multimodal learning, trustworthy and explainable AI, graph learning and foundation models, computer vision, machine learning and bioinformatics. A major emphasis is translating advances in AI into real-world applications through collaboration with clinicians, health researchers, scientists, government and industry, including Gold Coast University Hospital and Queensland Health.',
       'Professor Liew joined Griffith University in 2007. Before that he was an Assistant Professor in the Department of Computer Science and Engineering at The Chinese University of Hong Kong and a Senior Research Fellow in the Department of Electrical Engineering at City University of Hong Kong.',
       'He has published more than 300 journal and conference papers and two books, and holds three international patents.'],
  metrics='h-index 53 on Google Scholar and 42 on Scopus · More than 300 papers · 34 PhD completions at Griffith (20 as Principal Supervisor)',
  grants=["More than $13.46 million in research, industry and project funding as chief investigator or collaborating investigator, from the Australian Research Council (ARC), National Health and Medical Research Council (NHMRC), Medical Research Future Fund (MRFF), Australia's Economic Accelerator (AEA), Office of National Intelligence, CSIRO, Department of Foreign Affairs and Trade, Queensland Health and industry partners."],
  supervision='Since joining Griffith he has achieved 34 PhD completions, 20 as Principal Supervisor. He currently supervises doctoral researchers working on multimodal medical AI, trustworthy and explainable AI, graph foundation models, large language models, privacy-preserving learning, medical image analysis, image and video understanding, and AI applications in health and engineering.',
  education=['<b>PhD, Electrical and Electronic Engineering</b>, University of Tasmania, Hobart (1993–1996)',
             '<b>Bachelor of Engineering (First Class Honours), Electrical and Electronic Engineering</b>, University of Auckland (1989–1992)'],
  experience=[('2022 – present','Professor, School of Information and Communication Technology, Griffith University (Gold Coast)'),
              ('1 Jun 2022 – present','Head of School, School of ICT, Griffith University'),
              ('1 Mar 2019 – 31 Dec 2024','Deputy Director, Institute for Integrated and Intelligent Systems, Griffith University'),
              ('1 Oct 2018 – 31 May 2022','Deputy Head of School (Research), School of ICT, Griffith University'),
              ('2010 – 2021','Associate Professor, School of ICT, Griffith University'),
              ('Sep – Dec 2011','Visiting Professor, Linguistic Lab, Department of Chinese Language and Literature, Peking University'),
              ('2007 – 2009','Senior Lecturer, School of ICT, Griffith University'),
              ('2004 – 2006','Assistant Professor, Department of Computer Science and Engineering, The Chinese University of Hong Kong'),
              ('2002 – 2004','Senior Research Fellow, Department of Electronic Engineering, City University of Hong Kong'),
              ('1997 – 2002','Research Fellow, Department of Electronic Engineering, City University of Hong Kong')],
  service='Associate Editor of IEEE Transactions on Fuzzy Systems, Springer Nature Computer Science, International Journal of Computational Intelligence Systems and Machine Intelligence Research. Member of organising and programme committees of international conferences, assessor for nationally competitive research grants, and reviewer for international journals and conferences.',
  awards=['Fellow, Queensland Academy of Arts and Sciences','Fellow, Australian Computer Society','Senior Member, IEEE (since 2005)',"Stanford University World's Top 2% Scientists (Computer Science: AI and Image Processing), recognised since 2021 for career-long impact"])

def details_html(m):
    d = DETAILS.get(m['id'])
    if not d: return ''
    li = lambda xs: ''.join(f'<li>{x}</li>' for x in xs)
    out = ''
    if d.get('quals'): out += f'<h2>Qualifications</h2><p>{E(d["quals"])}</p>'
    if d.get('bio'): out += '<h2>About</h2>' + ''.join(f'<p>{E(p)}</p>' for p in d['bio'])
    if d.get('metrics'): out += f'<p class="muted">{E(d["metrics"])}</p>'
    if d.get('supervision'): out += f'<h2>Research supervision</h2><p>{E(d["supervision"])}</p>'
    if d.get('grants'): out += f'<h2>Research grants</h2><ul>{li(d["grants"])}</ul>'
    if d.get('expertise'): out += '<h2>Research expertise</h2><ul>' + li(f'<b>{E(a)}</b>: {E(b)}' for a, b in d['expertise']) + '</ul>'
    if d.get('education'): out += f'<h2>Academic background</h2><ul>{li(d["education"])}</ul>'
    if d.get('experience'): out += '<h2>Professional experience</h2><ul>' + li(f'<b>{E(a)}</b>: {E(b)}' for a, b in d['experience']) + '</ul>'
    if d.get('students'): out += f'<h2>Supervised research</h2><ul>{li(d["students"])}</ul>'
    if d.get('teaching'): out += f'<h2>Teaching areas</h2><ul class="tags tags--dark">{li(E(x) for x in d["teaching"])}</ul>'
    if d.get('courses'): out += '<h2>Courses taught</h2>' + ''.join(f'<details class="abs"><summary>{E(a)}</summary><p>{E("; ".join(b))}</p></details>' for a, b in d['courses'])
    if d.get('skills'): out += '<h2>Technical skills</h2><ul>' + li(f'<b>{E(a)}</b>: {E(b)}' for a, b in d['skills']) + '</ul>'
    if d.get('languages'): out += f'<h2>Languages</h2><p>{E(d["languages"])}</p>'
    if d.get('service'): out += f'<h2>Professional service</h2><p>{E(d["service"])}</p>'
    if d.get('awards'): out += f'<h2>Awards and fellowships</h2><ul>{li(E(x) for x in d["awards"])}</ul>'
    return out

def sectioned_profile(m, insts, achips, ints, pubsec):
    """Profile body split into navigable sections (used when DETAILS exist for a member)."""
    d = DETAILS.get(m['id'])
    if not d: return None
    li = lambda xs: ''.join(f'<li>{x}</li>' for x in xs)
    S = []   # (id, nav label, inner html)
    a = ''.join(f'<p>{E(p)}</p>' for p in d.get('bio', []))
    if d.get('metrics'): a += f'<p class="metric-line">{E(d["metrics"])}</p>'
    a += f'<h3>Affiliation</h3><p>{insts}</p>'
    S.append(('about', 'About', a))
    r = ''
    if d.get('expertise'): r += '<h3>Expertise</h3><ul>' + li(f'<b>{E(x)}</b>: {E(y)}' for x, y in d['expertise']) + '</ul>'
    r += f'<h3>Research areas</h3><div class="chips">{achips}</div><h3>Research interests</h3><ul class="tags tags--dark">{ints}</ul>'
    if d.get('grants'): r += f'<h3>Research grants</h3><ul>{li(d["grants"])}</ul>'
    if d.get('supervision') or d.get('students'):
        r += '<h3>Supervision</h3>' + (f'<p>{E(d["supervision"])}</p>' if d.get('supervision') else '') + (f'<ul>{li(d["students"])}</ul>' if d.get('students') else '')
    S.append(('research', 'Research', r))
    b = ''
    if d.get('education'): b += f'<h3>Education</h3><ul>{li(d["education"])}</ul>'
    if d.get('experience'): b += '<h3>Experience</h3><ul class="timeline">' + li(f'<span class="when">{E(x)}</span> {E(y)}' for x, y in d['experience']) + '</ul>'
    S.append(('background', 'Education & experience', b))
    t = ''
    if d.get('teaching'): t += f'<h3>Teaching areas</h3><ul class="tags tags--dark">{li(E(x) for x in d["teaching"])}</ul>'
    if d.get('courses'): t += '<h3>Courses taught</h3>' + ''.join(f'<details class="abs"><summary>{E(x)}</summary><p>{E("; ".join(y))}</p></details>' for x, y in d['courses'])
    S.append(('teaching', 'Teaching', t))
    k = ''
    if d.get('skills'): k += '<h3>Technical skills</h3><ul>' + li(f'<b>{E(x)}</b>: {E(y)}' for x, y in d['skills']) + '</ul>'
    if d.get('languages'): k += f'<h3>Languages</h3><p>{E(d["languages"])}</p>'
    if d.get('service'): k += f'<h3>Professional service</h3><p>{E(d["service"])}</p>'
    S.append(('skills', 'Skills & service', k))
    if d.get('awards'): S.append(('awards', 'Awards', f'<ul>{li(E(x) for x in d["awards"])}</ul>'))
    if pubsec: S.append(('publications', 'Publications', pubsec.replace('<h2>Publications</h2>', '')))
    nav = '<nav class="snav" aria-label="Profile sections"><div class="wrap"><ul>' + ''.join(f'<li><a href="#{i}">{E(n)}</a></li>' for i, n, _ in S) + '</ul></div></nav>'
    secs = ''.join(f'<section class="psec" id="{i}"><h2>{E(n)}</h2>{h}</section>' for i, n, h in S)
    return nav + f'<div class="wrap psecs">{secs}</div>'

def disp_name(m):
    return ' '.join(x for x in [m['prefix'], m['name']] if x)

CARD_LINK_GROUPS = {'leadership', 'advisors'}   # these cards show profile badges (ORCID, Scholar, LinkedIn, website)
def card_links(m):
    out = []
    if m['orcid']: out.append(('ORCID', f'https://orcid.org/{m["orcid"]}'))
    if m['scholar']: out.append(('Scholar', m['scholar']))
    if m['linkedin']: out.append(('LinkedIn', m['linkedin']))
    for t, u in m.get('links', []): out.append(('Website' if 'Experts' in t else t, u))
    return out

def person_card(m, p=''):
    areas = ''.join(f'<span class="dot" style="--c:{AREA[a]["color"]}" title="{AREA[a]["name"]}"></span>' for a in m['areas'])
    inst = E(m['inst'][0]) if m['inst'] else ''
    role = f'<p class="pcard__role">{E(m["role"])}</p>' if m['group'] == 'leadership' else ''   # role tags only for founders; groups are headed on the People page
    links = card_links(m) if m['group'] in CARD_LINK_GROUPS else []
    badges = ''
    if links:
        badges = '<div class="pcard__links">' + ''.join(f'<a href="{E(u)}" rel="noopener" target="_blank" aria-label="{E(m["name"])} on {t}">{t}</a>' for t, u in links) + '</div>'
    cls = 'pcard' + (f' pcard--links pcard--l{1 if len(links) <= 3 else 2 if len(links) <= 5 else 3}' if links else '')
    return (f'<li class="{cls}" data-areas="{" ".join(m["areas"] + (["methods"] if m.get("core") else [])) or "methods"}" data-name="{E((m["name"]+" "+m["alias"]+" "+" ".join(m["interests"])).lower())}">'
            f'<a href="{p}people/{m["id"]}.html">{avatar(m, p)}'
            f'<h3>{E(disp_name(m))}</h3>{role}'
            f'<p class="pcard__inst">{flag(m, p)}<span>{inst}</span></p><div class="dots">{areas}</div></a>{badges}</li>')

# ───────────────────────── pages ─────────────────────────
def build_pages(members):
    by_id = {m['id']: m for m in members}
    countries = sorted({m['country'] for m in members if m['country']})
    n_pubs = len(PUBS)
    partners = [
     ('University of the Sunshine Coast','university-of-the-sunshine-coast','au','https://www.usc.edu.au/'),
     ('Southern Cross University','southern-cross-university','au','https://www.scu.edu.au/'),
     ('University of Birmingham','university-of-birmingham','gb','https://www.birmingham.ac.uk/'),
     ('Simula Research Laboratory','simula-research-laboratory','no','https://www.simula.no/'),
     ('Tsinghua University','tsinghua-university','cn','https://www.tsinghua.edu.cn/en/'),
     ('Huazhong University of Science and Technology','huazhong-university-of-science-and-technology','cn','https://english.hust.edu.cn/'),
     ('Northwestern Polytechnical University','northwestern-polytechnical-university','cn','https://en.nwpu.edu.cn/'),
     ('Jožef Stefan Institute','jozef-stefan-institute','si','https://www.ijs.si/'),
     ('Chulalongkorn University','chulalongkorn-university','th','https://www.chula.ac.th/en/'),
     ('Universitas Atma Jaya Yogyakarta','universitas-atma-jaya-yogyakarta','id','https://www.uajy.ac.id/'),
     ('COMSATS University Islamabad','comsats-university-islamabad','pk','https://www.comsats.edu.pk/'),
     ('University of Babylon','university-of-babylon','iq','https://uobabylon.edu.iq/'),
    ]
    match = {'university-of-the-sunshine-coast': 'Sunshine Coast', 'southern-cross-university': 'Southern Cross',
             'university-of-birmingham': 'Birmingham', 'simula-research-laboratory': 'Simula', 'tsinghua-university': 'Tsinghua',
             'huazhong-university-of-science-and-technology': 'Huazhong', 'northwestern-polytechnical-university': 'Northwestern Polytechnical',
             'jozef-stefan-institute': 'Stefan', 'chulalongkorn-university': 'Chulalongkorn', 'universitas-atma-jaya-yogyakarta': 'Atma Jaya',
             'comsats-university-islamabad': 'COMSATS', 'university-of-babylon': 'Babylon'}
    people_at = {sl: [m for m in members if any(match[sl].lower() in i.lower() for i in m['inst'])] for sl in match}
    partners = [p for p in partners if people_at[p[1]]]
    area_cnt = {a['id']: sum(a['id'] in m['areas'] for m in members) for a in AREAS}

    # ---------- home
    area_cards = ''.join(
        f'<a class="area reveal" style="--c:{a["color"]}" href="research/{a["id"]}.html"><span class="area__ico">{icon(a["icon"])}</span>'
        f'<h3>{a["name"]}</h3><p class="area__short">{a["short"]}{"<span class=pill>Emerging</span>" if area_cnt[a["id"]] < 2 else ""}</p>'
        f'<p>{", ".join(a["topics"][:3])}…</p></a>' for a in AREAS)
    latest = ''.join(f'<li><a href="publications.html#pub-{p[0]}"><span class="pill pill--{p[1]}">{p[2]}</span> {E(p[4])}</a><small>{E(p[5])}</small></li>'
                     for p in sorted([x for x in PUBS if x[1] in ('journal','conference')], key=lambda x: -x[2])[:5])
    faces = ''.join(f'<li><a href="people/{m["id"]}.html" title="{E(disp_name(m))}">{avatar(m, "", "av av--sm")}</a></li>' for m in members)
    stats = [(len(members), 'Members'), (len(countries), 'Countries'), (n_pubs, 'Featured publications')]
    stat_html = ''.join(f'<div class="stat"><b data-count="{n}">{n}</b><span>{l}</span></div>' for n, l in stats)
    def logo_card(n, sl, u, pre=''):
        ppl = ', '.join(f'<a href="{pre}people/{m["id"]}.html">{E(m["name"])}</a>' for m in people_at[sl])
        return (f'<li><a class="logo" href="{u}" rel="noopener" target="_blank"><img src="{pre}images/partners/{sl}.webp" alt="{E(n)}" loading="lazy"></a>'
                f'<p>{E(n)}</p><small>{ppl}</small></li>')
    marquee = ''.join(logo_card(n, sl, u) for n, sl, c, u in partners)
    body = f'''
<section class="hero"><canvas id="net" aria-hidden="true"></canvas><div class="wrap hero__in">
  <p class="badge">Applied AI &amp; Connected Systems</p>
  <h1>Intelligence in <span class="grad">Synergy</span></h1>
  <p class="lead">Applied AI and connected systems research for healthier lives, sustainable food, personal learning and safer mobility. Open to academic collaborations, funded projects and industry partners.</p>
  <div class="cta"><a class="btn btn--pink" href="collaborate.html">Collaborate With Us</a><a class="btn btn--ghost" href="research.html">Explore Research</a></div>
</div></section>
<section class="stats"><div class="wrap stats__in">{stat_html}</div></section>
<section class="sec"><div class="wrap"><p class="eyebrow">Research areas</p><h2>Six areas, one connected vision</h2>
  <p class="sub">From hospitals and homes to farms, classrooms and roads, we design AI and connected technologies (IoT, IoV) that solve real problems.</p>
  <div class="grid grid--areas">{area_cards}</div></div></section>
<section class="sec sec--ice"><div class="wrap split">
  <div><p class="eyebrow">Latest papers</p><h2>Recent research output</h2><ul class="plist">{latest}</ul><p><a class="more" href="publications.html">All publications →</a></p></div>
  <div class="panel"><p class="eyebrow">Our name</p><h2>Syn · Te · Era</h2>
    <p><b>SYNTERA</b> brings together <b>Syn</b>ergy, In<b>te</b>lligence and E<b>ra</b>: a new era where AI works in synergy with health, agriculture, education, homes and mobility.</p>
    <p class="muted">Members and collaborators across {len(countries)} countries.</p>
    <a class="btn btn--blue" href="about.html">About the group</a></div></div></section>
<section class="sec"><div class="wrap"><p class="eyebrow">People</p><h2>The team</h2>
  <ul class="faces">{faces}</ul><p><a class="more" href="people.html">Meet everyone →</a></p></div></section>
<section class="sec sec--ice"><div class="wrap"><p class="eyebrow">Partner with us</p><h2>Academic collaboration, funded projects and industry partnerships</h2>
  <p class="sub">Whether you want a research partner for a grant, a university team to take on a paid project, or an AI pilot in a real setting, we would like to hear from you.</p>
  <div class="grid grid--3"><div class="panel"><h3>Academic collaboration</h3><p>Joint research and papers, shared methods and data, visiting researchers and co-supervision.</p></div>
  <div class="panel"><h3>Funded and paid projects</h3><p>Joint grant applications, industry-funded research and commissioned projects with agreed scope and deliverables.</p></div>
  <div class="panel"><h3>Industry and community pilots</h3><p>Proofs of concept and evaluations in health, care, agriculture, education, IoT and transport.</p></div></div>
  <p><a class="btn btn--pink" href="collaborate.html">See how to work with us</a></p></div></section>
<section class="sec"><div class="wrap"><div class="banner"><div><h2>Open to collaboration</h2>
  <p>Students, researchers, industry and institutions are welcome.</p></div>
  <div class="cta"><a class="btn btn--pink" href="collaborate.html">Collaborate</a><a class="btn btn--ghost" href="join.html">Join the group</a><a class="btn btn--ghost" href="contact.html">Contact us</a></div></div></div></section>'''
    page('index.html', BRAND['name'], 'Applied AI research group open to academic collaboration, funded and commissioned projects and industry partners in health, agriculture, education, IoT and IoV.', body, home=True,
         full_title='SYNTERA Research Group | AI Research Collaboration & Industry Partnerships',
         extra_js='<script src="js/hero.js?v=' + VER + '"></script>')

    # ---------- research overview
    methods = {}
    for m in members:
        for i in m['interests']:
            methods[i] = methods.get(i, 0) + 1
    core = ['Machine Learning','Deep Learning','Artificial Intelligence','Cybersecurity','NLP','Computer Vision','Explainable AI (XAI)','Information Systems']
    core_html = ''.join(f'<span class="chip chip--plain">{E(c)}</span>' for c in core)
    cards = ''.join(
        f'<a class="area area--lg" style="--c:{a["color"]}" href="research/{a["id"]}.html"><span class="area__ico">{icon(a["icon"])}</span>'
        f'<h3>{a["name"]}</h3><p class="area__short">{a["short"]}{"<span class=pill>Emerging</span>" if area_cnt[a["id"]] < 2 else ""}</p>'
        f'<p>{E(a["why"])}</p><ul class="tags">{"".join(f"<li>{t}</li>" for t in a["topics"])}</ul>'
        f'<small>{area_cnt[a["id"]]} team member{"s" if area_cnt[a["id"]] != 1 else ""}</small></a>' for a in AREAS)
    body = f'''<section class="phead"><div class="wrap"><p class="eyebrow">Research</p><h1>Six areas of applied AI</h1>
<p class="lead">AI for health, smart health homes, agriculture, education, the Internet of Things and the Internet of Vehicles.</p></div></section>
<section class="sec"><div class="wrap"><div class="grid grid--2">{cards}</div></div></section>
<section class="sec sec--ice"><div class="wrap"><p class="eyebrow">Core AI &amp; methods</p><h2>The toolkit behind every area</h2>
<p class="sub">Machine learning, deep learning, language and vision models, explainable and trustworthy AI, and cybersecurity are shared across all six areas.</p>
<div class="chips">{core_html}</div></div></section>
<section class="sec"><div class="wrap"><div class="banner"><div><h2>Work on these problems with us</h2><p>We partner on joint research, grant applications and funded or commissioned projects in every area.</p></div>
<div class="cta"><a class="btn btn--pink" href="collaborate.html">Collaborate with SYNTERA</a></div></div></div></section>'''
    page('research.html', 'Applied AI Research Areas', 'Six applied AI research areas open to collaboration and funded projects: AI for health, smart health home, agriculture, education, Internet of Things and Internet of Vehicles.', body, 'research')

    # ---------- area pages
    for a in AREAS:
        team = [m for m in members if a['id'] in m['areas']]
        pubs = [p for p in PUBS if a['id'] in p[7]]
        sec_team = (f'<section class="sec sec--ice"><div class="wrap"><p class="eyebrow">Team</p><h2>Who works on this</h2><ul class="grid grid--people">{"".join(person_card(m, "../") for m in team)}</ul></div></section>' if team else '')
        sec_pubs = (f'<section class="sec"><div class="wrap"><p class="eyebrow">Publications</p><h2>Selected publications</h2><ol class="publist" data-pubs data-area="{a["id"]}" data-limit="6">{static_pubs(lambda p, a=a: a["id"] in p[7], 6)}</ol><p><a class="more" href="../publications.html?area={a["id"]}">All {a["short"]} papers →</a></p></div></section>' if pubs else '')
        emerging = (f'<p class="notice">{a["short"]} is an emerging area for the group. Interested? <a href="../join.html#collaborate">Get in touch</a>.</p>' if len(team) < 2 else '')
        body = f'''<section class="phead phead--area" style="--c:{a["color"]}"><div class="wrap"><p class="eyebrow"><a href="../research.html">Research</a> / {a["short"]}</p>
<div class="phead__row"><span class="area__ico area__ico--xl">{icon(a["icon"])}</span><div><h1>{a["name"]}</h1>
<ul class="tags tags--light">{"".join(f"<li>{t}</li>" for t in a["topics"])}</ul></div></div></div></section>
<section class="sec"><div class="wrap prose"><h2>Why it matters</h2><p>{E(a["why"])}</p>{emerging}
<h2>Key challenges</h2><ul>{"".join(f"<li>{E(c)}</li>" for c in a["challenges"])}</ul>
<h2>Our approach</h2><p>{E(a["approach"])}</p></div></section>{sec_pubs}{sec_team}
<section class="sec"><div class="wrap"><div class="banner"><div><h2>Collaborate on {E(a["name"])}</h2>
<p>We welcome academic partners, funders and organisations who want to co-fund or commission research in {E(a["name"].lower())}: {E(", ".join(t.lower() for t in a["topics"]))}. Joint research, grant partnerships and paid proof-of-concept or evaluation projects are all possible.</p></div>
<div class="cta"><a class="btn btn--pink" href="mailto:{BRAND["email_collab"]}?subject=Collaboration%20on%20{a["name"].replace(" ", "%20")}">Email us</a><a class="btn btn--ghost" href="../collaborate.html">How we work with partners</a></div></div></div></section>'''
        page(f'research/{a["id"]}.html', a['name'] + ' Research', f'{a["name"]} research at SYNTERA: {", ".join(t.lower() for t in a["topics"][:3])}. Open to academic collaboration, funded projects and industry partnerships.', body, 'research', 1, pubs_on=True)

    # ---------- people
    chips = '<button class="fchip is-on" data-area="all">All</button>' + ''.join(
        f'<button class="fchip" data-area="{a["id"]}" style="--c:{a["color"]}">{a["name"]} <small>{area_cnt[a["id"]]}</small></button>' for a in AREAS) + \
        f'<button class="fchip" data-area="methods">Core AI &amp; Methods <small>{sum(bool(m.get("core")) for m in members)}</small></button>'
    secs = ''
    for gid, gt in GROUPS:
        g = [m for m in members if m['group'] == gid]
        if g: secs += f'<section class="pgroup" id="g-{gid}"><h2>{gt} <small>{len(g)}</small></h2><ul class="grid grid--people">{"".join(person_card(m) for m in g)}</ul></section>'
    body = f'''<section class="phead"><div class="wrap"><p class="eyebrow">People</p><h1>The SYNTERA team</h1>
<p class="lead">{len(members)} researchers, academics and students across {len(countries)} countries.</p></div></section>
<section class="sec"><div class="wrap"><div class="filters"><input id="q" type="search" placeholder="Search people or interests" aria-label="Search people"><div class="fchips" id="fchips">{chips}</div></div>
<div id="people">{secs}</div><p id="none" class="notice" hidden>No one matches that filter.</p></div></section>'''
    page('people.html', 'People', 'The director, researchers, academics and students of SYNTERA Research Group.', body, 'people', extra_js='<script src="js/people.js?v=' + VER + '"></script>')

    # ---------- profiles
    for i, m in enumerate(members):
        lk = []
        if m['orcid']: lk.append(('ORCID', f'https://orcid.org/{m["orcid"]}'))
        if m['scholar']: lk.append(('Google Scholar', m['scholar']))
        if m['linkedin']: lk.append(('LinkedIn', m['linkedin']))
        lk += m.get('links', [])
        links = ''.join(f'<a class="btn btn--ghost-d btn--sm" href="{E(u)}" rel="noopener" target="_blank">{t}</a>' for t, u in lk)
        achips = ''.join(f'<a class="chip" style="--c:{AREA[a]["color"]}" href="../research/{a}.html">{AREA[a]["name"]}</a>' for a in m['areas']) + ('<span class="chip chip--plain">Core AI &amp; Methods</span>' if m.get('core') else '')
        ints = ''.join(f'<li>{E(x)}</li>' for x in m['interests'])
        insts = '<br>'.join(E(x) for x in m['inst']) or '<span class="muted">Affiliation to be confirmed</span>'
        pubsec = f'<h2>Publications</h2><ol class="publist" data-pubs data-member="{m["id"]}">{static_pubs(lambda p, m=m: any(mid == m["id"] for _, mid in parse_authors(p[3])))}</ol>' if any(m['id'] == mid for p in PUBS for _, mid in parse_authors(p[3])) else ''
        prv = by_id[members[i-1]['id']] if i else members[-1]; nxt = members[(i+1) % len(members)]
        sp = sectioned_profile(m, insts, achips, ints, pubsec)
        main_html = f'''<section class="sec"><div class="wrap prose"><h2>Affiliation</h2><p>{insts}</p>
<h2>Research areas</h2><div class="chips">{achips}</div>
<h2>Research interests</h2><ul class="tags tags--dark">{ints}</ul>{pubsec}</div></section>'''
        pager = f'<div class="wrap"><nav class="pager"><a href="{prv["id"]}.html">← {E(disp_name(prv))}</a><a href="{nxt["id"]}.html">{E(disp_name(nxt))} →</a></nav></div>'
        body = f'''<section class="phead phead--profile"><div class="wrap profile"><div class="profile__ph">{avatar(m, "../", "av av--lg")}</div>
<div><p class="eyebrow"><a href="../people.html">People</a></p><h1>{E(disp_name(m))}{f"<small>{E(m['suffix'])}</small>" if m["suffix"] else ""}</h1>
<p class="lead">{E(m["role"])}{" · " + E(DETAILS.get(m["id"], {}).get("title", m["title"])) if m["title"] else ""}</p>
{"".join(f'<p class="lead lead--sub">{E(r)}</p>' for r in DETAILS.get(m["id"], {}).get("roles", []))}<p class="where">{flag(m, "../")} {E(m["country"])}</p><div class="cta">{links}</div></div></div></section>
{sp if sp else main_html}{pager}'''
        person = {"@context": "https://schema.org", "@type": "Person", "@id": f'{SITE_URL}/people/{m["id"]}#person', "name": m['name'], "url": f'{SITE_URL}/people/{m["id"]}',
                  "jobTitle": m['role'], "worksFor": {"@id": SITE_URL + "/#org"}, "memberOf": {"@id": SITE_URL + "/#org"},
                  "knowsAbout": m['interests'] or [AREA[a]['name'] for a in m['areas']], "sameAs": [u for _, u in lk]}
        if m['photo']: person["image"] = f'{SITE_URL}/images/team/{m["photo"]}?v={m["pv"]}'
        if m['inst']: person["affiliation"] = [{"@type": "Organization", "name": x} for x in m['inst']]
        person = {k: v for k, v in person.items() if v}
        ints_s = ', '.join(m['interests'][:4])
        page(f'people/{m["id"]}.html', disp_name(m), f'{disp_name(m)}, {m["role"]} at SYNTERA Research Group' + (f'. Research: {ints_s}.' if ints_s else '.') + ' Open to collaboration.', body, 'people', 1, pubs_on=True, ld_extra=ld_tag(person))

    # ---------- publications
    types = [('all','All')] + [(k, t) for k, t in [('journal','Journal'),('conference','Conference'),('book','Book'),('chapter','Chapter'),('thesis','Thesis'),('preprint','Preprint'),('other','Other')] if any(p[1] == k for p in PUBS)]
    years = sorted({p[2] for p in PUBS if p[2]}, reverse=True)
    tchips = ''.join(f'<button class="fchip{" is-on" if k=="all" else ""}" data-type="{k}">{t}</button>' for k, t in types)
    achips = '<button class="fchip is-on" data-area="all">All areas</button>' + ''.join(f'<button class="fchip" data-area="{a["id"]}" style="--c:{a["color"]}">{a["name"]}</button>' for a in AREAS)
    yopts = '<option value="all">All years</option>' + ''.join(f'<option>{y}</option>' for y in years)
    body = f'''<section class="phead"><div class="wrap"><p class="eyebrow">Publications</p><h1>Research output</h1>
<p class="lead">Papers, chapters, books and theses by group members and the director. Group members are shown in bold and link to their profiles.</p>
<p class="muted metrics">Director's Google Scholar profile (6 Oct 2026): 15,573 citations (12,216 since 2021) · h-index 21 (18 since 2021) · i10-index 22 (21 since 2021). <a href="{BRAND['scholar']}" rel="noopener" target="_blank">View profile</a></p></div></section>
<section class="sec"><div class="wrap"><div class="filters"><input id="q" type="search" placeholder="Search title, author, keyword or venue" aria-label="Search publications">
<select id="year" aria-label="Year">{yopts}</select></div>
<div class="fchips" id="tchips">{tchips}</div><div class="fchips" id="achips">{achips}</div>
<p id="count" class="muted" aria-live="polite"></p><ol class="publist" id="pubs" data-pubs>{static_pubs()}</ol><p id="none" class="notice" hidden>No publications match those filters.</p></div></section>'''
    page('publications.html', 'Publications', f'{n_pubs} journal articles, conference papers, books and chapters on applied AI, IoT, IoV, health and education from SYNTERA Research Group. Find research collaborators.', body, 'publications', pubs_on=True)

    # ---------- about
    values = [('Synergy','We achieve more together, across disciplines, institutions and countries.'),('Impact','We pursue research that solves real problems for real people.'),
              ('Responsibility','We build AI that is ethical, explainable and trustworthy.'),('Openness','We share methods, code and findings.'),('Curiosity','We ask bold questions and learn from every result.')]
    vals = ''.join(f'<li><h3>{t}</h3><p>{d}</p></li>' for t, d in values)
    plog = marquee
    director = by_id['shahrzad-saremi']
    body = f'''<section class="phead"><div class="wrap"><p class="eyebrow">About</p><h1>Intelligence in Synergy</h1>
<p class="lead">{BRAND['full']}.</p></div></section>
<section class="sec" id="story"><div class="wrap prose"><h2>Our story</h2>
<p><b>SYNTERA</b> stands for <b>Syn</b>ergy + In<b>te</b>lligence + E<b>ra</b>: a new era in which artificial intelligence works in synergy with health, agriculture, education, homes and mobility.</p>
<p>The group was founded by <a href="people/shahrzad-saremi.html">Dr. Shahrzad Saremi</a>, <a href="people/rania-shibl.html">Professor Rania Shibl</a> and <a href="people/dana-dermody.html">Assoc. Prof. Dana Dermody</a>, and brings together researchers, academics and students from {len(countries)} countries to build AI and connected systems (IoT and IoV) for real-world problems.</p>
<p>We collaborate with universities, research institutes, health and care providers, industry and government. See <a href="collaborate.html">how to partner with us</a> on joint research, funded projects and commissioned work.</p>
<p class="muted">SYNTERA Research Group is not affiliated with any commercial company of a similar name.</p></div></section>
<section class="sec sec--ice" id="mission"><div class="wrap"><div class="grid grid--2"><div class="panel"><h2>Mission</h2>
<p>To design and apply AI and connected technologies (IoT, IoV) that solve real problems in health, agriculture, education and everyday living.</p></div>
<div class="panel"><h2>Vision</h2><p>A future where intelligent, connected systems make life healthier, food more sustainable, learning more personal and mobility safer.</p></div></div>
<h2 class="mt">Core values</h2><ul class="values">{vals}</ul></div></section>
<section class="sec" id="director"><div class="wrap"><div class="dir"><div>{avatar(director, "", "av av--lg")}</div>
<div class="prose"><p class="eyebrow">Founder &amp; Director</p><h2><a href="people/{director["id"]}.html">{E(disp_name(director))}</a></h2><p>{E(DETAILS["shahrzad-saremi"]["title"])}, {E(director["inst"][0])}.</p>
<p>{E(DETAILS["shahrzad-saremi"]["bio"][0])}</p>
<p>{E(DETAILS["shahrzad-saremi"]["bio"][1].split(". She is widely")[0])}.</p><p><a class="btn btn--blue btn--sm" href="people/{director["id"]}.html">Full profile</a></p></div></div></div></section>
'''
    page('about.html', 'About Our Applied AI Research Group', 'The story, mission, vision and values of SYNTERA Research Group, an applied AI and connected systems group open to academic and industry collaboration.', body, 'about')

    # ---------- join
    body = f'''<section class="phead"><div class="wrap"><p class="eyebrow">Join us</p><h1>Open to collaboration</h1>
<p class="lead">PhD and Master's students, postdocs, research assistants, interns, industry and academic partners are all welcome.</p>
<a class="btn btn--pink" href="mailto:{BRAND['email_join']}?subject=Joining%20SYNTERA%20Research%20Group">Reach out</a></div></section>
<section class="sec" id="positions"><div class="wrap prose"><h2>Open positions</h2><p>No positions are open right now. Send us a short note anyway: we like hearing from motivated people.</p></div></section>
<section class="sec sec--ice" id="apply"><div class="wrap prose"><h2>How to apply</h2>
<p>Email us with the following attached:</p><ul><li><b>CV</b></li><li><b>Academic transcript</b></li><li><b>Research statement</b> (one page)</li></ul>
<p>In your message, tell us:</p><ol><li>Who you are</li><li>Which of our <a href="research.html">research areas</a> interests you</li><li>Where we can learn more about your work (Scholar, GitHub, ORCID)</li><li>What you can offer</li></ol>
<p><a class="btn btn--pink" href="mailto:{BRAND['email_student']}?subject=Prospective%20student%20enquiry">Email as a prospective student</a> <a class="btn btn--ghost" href="mailto:{BRAND['email_join']}?subject=Joining%20SYNTERA%20Research%20Group">Postdocs, RAs and interns</a></p>
<p class="muted">Prospective students: <a href="mailto:{BRAND['email_student']}">{BRAND['email_student']}</a>. Everyone else joining the group: <a href="mailto:{BRAND['email_join']}">{BRAND['email_join']}</a>.</p></div></section>
<section class="sec" id="collaborate"><div class="wrap prose"><h2>Collaborate with us</h2>
<p>We work with universities, hospitals, aged-care providers, farms, schools and industry. If you have a real problem where applied AI or connected systems could help, or want a research partner for a funded or paid project, we would like to talk. See <a href="collaborate.html">academic collaboration, funding and industry partnerships</a>.</p>
<p><a class="btn btn--blue" href="mailto:{BRAND['email_collab']}?subject=Collaboration%20with%20SYNTERA%20Research%20Group">Propose a collaboration</a></p>
<p class="muted">Collaboration enquiries: <a href="mailto:{BRAND['email_collab']}">{BRAND['email_collab']}</a></p></div></section>'''
    page('join.html', 'Join Us: PhD, Postdoc & Research Roles', 'PhD, Master\'s, postdoc, research assistant and internship opportunities in applied AI at SYNTERA, and how to apply.', body, 'join')

    # ---------- contact
    soc = f'<a class="btn btn--ghost-d btn--sm" href="{BRAND["scholar"]}" rel="noopener" target="_blank">Director on Google Scholar</a>'
    body = f'''<section class="phead"><div class="wrap"><p class="eyebrow">Contact</p><h1>Get in touch</h1></div></section>
<section class="sec"><div class="wrap"><div class="grid grid--3">
<div class="panel"><h3>Email</h3><p><b>Prospective students</b><br><a href="mailto:{BRAND['email_student']}">{BRAND['email_student']}</a></p><p><b>Joining the group</b><br><a href="mailto:{BRAND['email_join']}">{BRAND['email_join']}</a></p><p><b>Collaborations</b><br><a href="mailto:{BRAND['email_collab']}">{BRAND['email_collab']}</a></p><p class="muted">Messages reach the group Director, <a href="people/shahrzad-saremi.html">Dr. Shahrzad Saremi</a>.</p></div>
<div class="panel"><h3>Director's affiliation</h3><p><a href="{BRAND['host_url']}" rel="noopener" target="_blank">{BRAND['host']}</a></p><p class="muted">School of Science, Technology and Engineering, Queensland, Australia.</p></div>
<div class="panel"><h3>Elsewhere</h3><p>{soc}</p></div></div></div></section>'''
    page('contact.html', 'Contact & Collaboration Enquiries', 'Contact SYNTERA Research Group about research collaboration, funded or commissioned projects, industry partnerships and student enquiries.', body, 'contact')

    # ---------- collaborate (academic collaboration, funding, industry partnerships, paid projects)
    mail = lambda subj: f'mailto:{BRAND["email_collab"]}?subject=' + subj.replace(' ', '%20')
    ways = [
     ('Academic collaboration', 'Joint research and co-authored papers, shared methods and datasets, visiting researchers, co-supervision of PhD and Master\'s students, and joint workshops with universities and research institutes.'),
     ('Funding and joint grant applications', 'We partner on grant proposals as a research partner or lead investigator, from early scoping to submission, including industry-linked and government-funded schemes. Our director leads a 2026 LAUNCH Partnership Grant on AI-based knee MRI segmentation.'),
     ('Commissioned and paid research', 'Organisations can engage the group for a defined project: a proof of concept, data analysis, model development and evaluation, a technical review or an AI feasibility study. Scope, deliverables, IP and fees are agreed in writing before work starts.'),
     ('Industry and community pilots', 'Pilot AI and IoT solutions in real settings such as clinics, aged care, farms, classrooms and transport, with privacy, ethics and explainability built in.'),
     ('Student and talent partnerships', 'Industry-linked student projects, internships and higher degree research (PhD and Master\'s) that you co-fund or co-supervise with us.'),
     ('Knowledge exchange', 'Seminars, workshops, expert advice and co-designed training on responsible and applied AI for your team or community.'),
    ]
    ways_html = ''.join(f'<div class="panel"><h3>{t}</h3><p>{d}</p></div>' for t, d in ways)
    area_html = ''.join(f'<li><a href="research/{a["id"]}.html"><b>{a["name"]}</b></a>: {E(", ".join(t.lower() for t in a["topics"]))}.</li>' for a in AREAS)
    steps = ['<b>Email us</b> at <a href="mailto:%s">%s</a> with your problem or idea, your organisation, timeline and funding situation (if any).' % (BRAND['email_collab'], BRAND['email_collab']),
             '<b>Scoping conversation.</b> We discuss the goals, data, risks and the best way to work together, whether that is a collaboration, a grant or a paid project.',
             '<b>Proposal and agreement.</b> We outline scope, deliverables, timeline, budget, data handling, IP and publication terms in writing.',
             '<b>Deliver and share.</b> We run the project with regular check-ins and share the results, reports, code or papers as agreed.']
    steps_html = ''.join(f'<li>{s}</li>' for s in steps)
    partner_html = ', '.join(f'<a href="{u}" rel="noopener" target="_blank">{E(n)}</a>' for n, sl, c, u in partners)
    faqs = [
     ('Can we commission a paid research project with SYNTERA?', 'Yes. Organisations can engage the group for a defined, funded project such as a proof of concept, data analysis, model development, evaluation or feasibility study. We agree scope, deliverables, timeline, fees, data handling and intellectual property in writing before work starts. Email %s to start.' % BRAND['email_collab']),
     ('Can we apply for a research grant together?', 'Yes. We join grant applications as a research partner or lead investigator and can help scope the research, build the partnership and prepare the proposal, including industry-linked and government-funded schemes in Australia and internationally.'),
     ('Which fields does SYNTERA work in?', 'Applied AI and connected systems: AI for health, smart health homes, AI for agriculture, AI for education, the Internet of Things and the Internet of Vehicles, supported by machine learning, deep learning, computer vision, language models, explainable AI and cybersecurity.'),
     ('Who can collaborate with SYNTERA?', 'Universities and research institutes, hospitals and aged-care providers, farms and agri-businesses, schools and education providers, companies, government agencies and community organisations, in Australia and overseas.'),
     ('Do you work with partners outside Australia?', 'Yes. SYNTERA has members in %d countries and works with partner institutions on several continents. We collaborate remotely and in person.' % len(countries)),
     ('Where is SYNTERA based?', 'SYNTERA is led by Dr. Shahrzad Saremi, a lecturer at the University of the Sunshine Coast in Queensland, Australia, with co-director Professor Rania Shibl.'),
     ('Do you offer student projects, internships or PhD supervision?', 'Yes. We welcome PhD and Master\'s applicants, research assistants and interns, and can run industry-linked student projects. See the Join Us page for how to apply.'),
     ('How do we get started?', 'Send a short email to %s describing the problem, your organisation, timeline and whether funding is available. We will reply to arrange a scoping conversation.' % BRAND['email_collab']),
    ]
    faq_html = ''.join(f'<details class="abs"><summary>{E(q)}</summary><p>{E(a)}</p></details>' for q, a in faqs)
    faq_ld = ld_tag({"@context": "https://schema.org", "@type": "FAQPage", "mainEntity": [
        {"@type": "Question", "name": q, "acceptedAnswer": {"@type": "Answer", "text": a}} for q, a in faqs]})
    svc_ld = ld_tag({"@context": "https://schema.org", "@type": "Service", "name": "Research collaboration, funded and commissioned research projects",
        "serviceType": "Academic research collaboration and industry research partnerships", "provider": {"@id": SITE_URL + "/#org"}, "areaServed": "Worldwide",
        "description": "Joint research, grant partnerships, industry-funded and commissioned research projects, pilots and student projects in applied AI and connected systems.",
        "url": SITE_URL + "/collaborate"})
    body = f'''<section class="phead"><div class="wrap"><p class="eyebrow">Collaborate</p><h1>Research collaboration, funded projects and industry partnerships</h1>
<p class="lead">Partner with SYNTERA on academic collaborations, grant applications, industry-funded research and commissioned (paid) projects in applied AI and connected systems.</p>
<div class="cta"><a class="btn btn--pink" href="{mail('Collaboration with SYNTERA Research Group')}">Start a collaboration</a><a class="btn btn--ghost" href="#process">How it works</a></div></div></section>
<section class="sec"><div class="wrap"><p class="eyebrow">Ways to work with us</p><h2>Six ways to partner</h2>
<p class="sub">SYNTERA is a research group of {len(members)} researchers, academics and students across {len(countries)} countries, led from the University of the Sunshine Coast, Queensland, Australia. We work with universities, health and care providers, industry and government.</p>
<div class="grid grid--3">{ways_html}</div></div></section>
<section class="sec sec--ice" id="areas"><div class="wrap prose"><p class="eyebrow">Where we can help</p><h2>Research areas open to partnership</h2>
<ul>{area_html}</ul>
<p>Our methods include machine learning, deep learning, computer vision, natural language processing, explainable AI and cybersecurity. Browse our <a href="research.html">research areas</a> and <a href="publications.html">{n_pubs} publications</a>.</p></div></section>
<section class="sec" id="process"><div class="wrap prose"><p class="eyebrow">How it works</p><h2>From first email to results</h2><ol>{steps_html}</ol>
<p><a class="btn btn--blue" href="{mail('Collaboration with SYNTERA Research Group')}">Email {BRAND['email_collab']}</a></p></div></section>
<section class="sec sec--ice" id="who"><div class="wrap prose"><p class="eyebrow">Who we work with</p><h2>Universities, institutes, industry and government</h2>
<p>Our members come from and collaborate with {partner_html}, among others. Meet the <a href="people.html">team</a> or read about the <a href="about.html">group</a>.</p></div></section>
<section class="sec" id="faq"><div class="wrap prose"><p class="eyebrow">FAQ</p><h2>Frequently asked questions</h2>{faq_html}</div></section>
<section class="sec"><div class="wrap"><div class="banner"><div><h2>Have a project, a grant or a problem to solve?</h2><p>Tell us what you need. We will reply to arrange a conversation.</p></div>
<div class="cta"><a class="btn btn--pink" href="{mail('Collaboration with SYNTERA Research Group')}">Email collaborations</a><a class="btn btn--ghost" href="join.html">Join the group</a></div></div></div></section>'''
    page('collaborate.html', 'Research Collaboration & Funded Projects', 'Partner with SYNTERA on academic collaboration, joint grants, industry-funded and commissioned (paid) research projects in applied AI, IoT and health.', body, 'collaborate', ld_extra=faq_ld + svc_ld)

    # ---------- privacy
    body = f'''<section class="phead"><div class="wrap"><p class="eyebrow">Privacy</p><h1>Privacy</h1></div></section>
<section class="sec"><div class="wrap prose"><p>This website has no accounts, forms, analytics or advertising cookies, and loads no third-party scripts or fonts.</p>
<h2>What is stored in your browser</h2><ul><li><code>syntera-theme</code>: your light or dark preference.</li></ul>
<h2>Member information</h2><p>Photos and profile details appear with each member's agreement. Member email addresses are not published. To correct or remove information about you, email <a href="mailto:{BRAND['email']}">{BRAND['email']}</a>.</p>
<h2>Outbound links</h2><p>Links to ORCID, Google Scholar, LinkedIn and partner sites open in a new tab and are governed by those sites' own policies.</p>
<p class="muted">Last reviewed: October 2026.</p></div></section>'''
    page('privacy.html', 'Privacy', 'What the SYNTERA Research Group website stores and how to request changes to personal information.', body, '')

# ───────────────────────── assets ─────────────────────────
CSS = r'''
:root{--navy:#1A2A6C;--hero:#0F1A47;--bg:#fff;--alt:#E8F0FF;--alt2:#FFE6F1;--text:#14213D;--muted:#52607F;--blue:#3A7BFF;--blue-ink:#2457D6;--blue-btn:#2D66E6;--pink:#E83E8C;--pink-ink:#BE1A6B;--pink-btn:#D02C7A;--pink-soft:#FF8FC3;--card:#fff;--line:#D6DEF0;--strong:#7886A8;--grad:linear-gradient(135deg,#E83E8C,#3A7BFF);--sh:0 6px 24px rgba(15,26,71,.09)}
:root[data-theme=dark]{--bg:#0A0F1F;--alt:#0F1730;--alt2:#1a1230;--text:#E8EDF7;--muted:#9AA7C4;--blue-ink:#8FB0FF;--blue-btn:#8FB0FF;--pink-ink:#FF8FC3;--pink-btn:#FF8FC3;--card:#121A33;--line:#243056;--strong:#6c7aa0;--sh:0 6px 24px rgba(0,0,0,.4)}
@media(prefers-color-scheme:dark){:root:not([data-theme=light]){--bg:#0A0F1F;--alt:#0F1730;--alt2:#1a1230;--text:#E8EDF7;--muted:#9AA7C4;--blue-ink:#8FB0FF;--blue-btn:#8FB0FF;--pink-ink:#FF8FC3;--pink-btn:#FF8FC3;--card:#121A33;--line:#243056;--strong:#6c7aa0;--sh:0 6px 24px rgba(0,0,0,.4)}}
*{box-sizing:border-box}html{scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--text);font:16px/1.65 Inter,system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;-webkit-font-smoothing:antialiased}
h1,h2,h3,h4{font-family:"Space Grotesk",Inter,system-ui,sans-serif;line-height:1.15;margin:0 0 .5em;letter-spacing:-.01em}
h1{font-size:clamp(2.1rem,5vw,3.6rem)}h2{font-size:clamp(1.55rem,3vw,2.2rem)}h3{font-size:1.15rem}
a{color:var(--blue-ink)}img{max-width:100%}
.wrap{width:min(1160px,100% - 32px);margin-inline:auto}
.skip{position:absolute;left:-999px;top:8px;background:#fff;color:#000;padding:8px 12px;z-index:100}.skip:focus{left:8px}
:focus-visible{outline:3px solid var(--blue);outline-offset:2px;border-radius:4px}
.muted{color:var(--muted)}.mt{margin-top:2rem}
/* header */
.hdr{position:sticky;top:0;z-index:50;background:var(--hero);color:#fff;border-bottom:1px solid rgba(255,255,255,.1)}
.hdr__in{display:flex;align-items:center;gap:20px;min-height:66px}
.brand{display:flex;align-items:center;gap:10px;text-decoration:none;color:#fff}.brand svg{width:38px;height:38px}
.brand span{display:flex;flex-direction:column;line-height:1.1}.brand b{font:700 1.15rem "Space Grotesk",sans-serif;letter-spacing:.06em}.brand small{font-size:.72rem;opacity:.8}
.hdr nav{margin-left:auto}.nav{display:flex;align-items:center;gap:4px;list-style:none;margin:0;padding:0}
.nav a:not(.btn){color:#fff;text-decoration:none;padding:8px 12px;border-radius:8px;font-weight:500;font-size:.95rem}
.nav a:not(.btn):hover,.nav a[aria-current]{background:rgba(255,255,255,.12)}
.nav__cta{margin-left:8px}
.hdr__act{display:flex;gap:6px}.icon-btn{background:none;border:0;color:#fff;width:40px;height:40px;border-radius:10px;cursor:pointer;display:grid;place-items:center}.icon-btn:hover{background:rgba(255,255,255,.12)}.icon-btn svg{width:22px;height:22px}
.burger{display:none}
@media(max-width:900px){.burger{display:grid}.hdr nav{position:absolute;left:0;right:0;top:100%;background:var(--hero);margin:0;display:none;padding:8px 16px 16px}
.hdr nav.open{display:block}.nav{flex-direction:column;align-items:stretch}.nav a{display:block}.nav__cta{margin:8px 0 0}.nav__cta .btn{display:block;text-align:center}.hdr__act{margin-left:auto}}
/* buttons */
.btn{display:inline-block;padding:11px 22px;border-radius:999px;font-weight:600;text-decoration:none;border:2px solid transparent;cursor:pointer;font-size:.95rem;transition:transform .15s,box-shadow .15s}
.btn:hover{transform:translateY(-2px);box-shadow:var(--sh)}.btn--sm{padding:7px 16px;font-size:.85rem}
.btn--pink{background:var(--pink-btn);color:#fff}.btn--blue{background:var(--blue-btn);color:#fff}
:root[data-theme=dark] .btn--pink,:root[data-theme=dark] .btn--blue{color:#0A0F1F}
@media(prefers-color-scheme:dark){:root:not([data-theme=light]) .btn--pink,:root:not([data-theme=light]) .btn--blue{color:#0A0F1F}}
.btn--ghost{border-color:rgba(255,255,255,.6);color:#fff}.btn--ghost-d{border-color:var(--strong);color:var(--text)}
.cta{display:flex;flex-wrap:wrap;gap:10px;margin-top:1.4rem}
/* hero */
.hero{position:relative;color:#fff;background:linear-gradient(135deg,#0F1A47,#1A2A6C 55%,#2c4a9a);overflow:hidden}
.hero::before,.hero::after{content:"";position:absolute;border-radius:50%;filter:blur(90px);opacity:.32}
.hero::before{width:420px;height:420px;background:#E83E8C;left:-120px;top:-120px}.hero::after{width:460px;height:460px;background:#3A7BFF;right:-140px;bottom:-160px}
#net{position:absolute;inset:0;width:100%;height:100%}
.hero__in{position:relative;padding:clamp(70px,12vw,140px) 0}.hero__in>*{max-width:760px}
.badge{display:inline-block;border:1px solid rgba(255,255,255,.35);border-radius:999px;padding:5px 14px;font-size:.85rem;margin:0 0 1.2rem}
.grad{background:linear-gradient(90deg,#FF8FC3,#8FB4FF);-webkit-background-clip:text;background-clip:text;color:transparent}
.lead{font-size:1.2rem;max-width:62ch}.lead--sub{font-size:1.02rem;margin:.2rem 0 .6rem}.hero .lead{color:#E3E9FA}
/* stats */
.stats{background:var(--hero);color:#fff;border-top:1px solid rgba(255,255,255,.1)}.stats__in{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;padding:28px 0;text-align:center}
.stat b{display:block;font:700 2.4rem "Space Grotesk",sans-serif;background:linear-gradient(90deg,#FF8FC3,#8FB4FF);-webkit-background-clip:text;background-clip:text;color:transparent}.stat span{font-size:.9rem;color:#C9D4F2}
/* sections */
.sec{padding:clamp(48px,7vw,88px) 0}.sec--ice{background:var(--alt)}.sec--blush{background:var(--alt2)}
.eyebrow{font:600 .8rem Inter,sans-serif;letter-spacing:.12em;text-transform:uppercase;color:var(--blue-ink);margin:0 0 .6rem}.eyebrow a{color:inherit}
.sub{max-width:68ch;color:var(--muted)}
.phead{background:linear-gradient(135deg,#0F1A47,#1A2A6C);color:#fff;padding:clamp(40px,6vw,72px) 0}.phead .eyebrow{color:var(--pink-soft)}.phead .lead{color:#DCE4FA}.phead .muted{color:#B8C4E6}.phead a{color:#CFE0FF}
.phead h1{font-size:clamp(2rem,4.5vw,3rem)}.metrics{font-size:.9rem;margin-top:1rem}
.phead--area{background:linear-gradient(135deg,#0F1A47,#1A2A6C 60%,color-mix(in srgb,var(--c) 55%,#1A2A6C))}
.phead__row{display:flex;gap:20px;align-items:center}
.grid{display:grid;gap:20px;list-style:none;padding:0;margin:1.6rem 0}
.grid--areas{grid-template-columns:repeat(auto-fit,minmax(260px,1fr))}.grid--2{grid-template-columns:repeat(auto-fit,minmax(min(100%,420px),1fr))}.grid--3{grid-template-columns:repeat(auto-fit,minmax(260px,1fr))}
.grid--people{grid-template-columns:repeat(auto-fill,minmax(190px,1fr))}
.area{display:block;background:var(--card);border:1px solid var(--line);border-top:4px solid var(--c);border-radius:16px;padding:22px;text-decoration:none;color:var(--text);transition:transform .2s,box-shadow .2s}
.area:hover{transform:translateY(-4px);box-shadow:var(--sh)}.area p{margin:.3rem 0;color:var(--muted);font-size:.94rem}.area h3{color:var(--text)}
.area__ico{display:inline-grid;place-items:center;width:46px;height:46px;border-radius:12px;background:color-mix(in srgb,var(--c) 14%,transparent);color:var(--c);margin-bottom:10px}.area__ico .ico{width:26px;height:26px}
.area__ico--xl{width:76px;height:76px;margin:0;background:rgba(255,255,255,.14);color:#fff;flex:none}.area__ico--xl .ico{width:42px;height:42px}
.area__short{font-weight:600;color:var(--text)!important}
.pill{display:inline-block;font-size:.72rem;font-weight:600;padding:2px 9px;border-radius:999px;background:var(--alt);color:var(--blue-ink);margin-left:8px;vertical-align:middle}
.pill--journal{background:#E8F0FF;color:#2457D6}.pill--conference{background:#FFE6F1;color:#BE1A6B}.pill--chapter{background:#E9F7EE;color:#15803D}.pill--book,.pill--thesis,.pill--other{background:#EEE9FB;color:#5B3FB0}.pill--preprint{background:#FFF3DC;color:#92560A}
.tags{display:flex;flex-wrap:wrap;gap:6px;list-style:none;padding:0;margin:.6rem 0}.tags li{font-size:.8rem;padding:3px 10px;border-radius:999px;background:var(--alt);color:var(--text)}
.tags--light li{background:rgba(255,255,255,.16);color:#fff}
.chips{display:flex;flex-wrap:wrap;gap:8px;margin:.6rem 0}.chip{font-size:.8rem;font-weight:600;padding:4px 12px;border-radius:999px;border:1px solid var(--c,var(--strong));color:var(--text);text-decoration:none;background:color-mix(in srgb,var(--c,transparent) 10%,transparent)}.chip--plain{border-color:var(--line);background:var(--card)}
.split{display:grid;grid-template-columns:1.3fr 1fr;gap:36px;align-items:start}@media(max-width:820px){.split{grid-template-columns:1fr}}
.panel{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:26px;box-shadow:var(--sh)}
.plist{list-style:none;padding:0;margin:1rem 0}.plist li{padding:12px 0;border-bottom:1px solid var(--line)}.plist a{font-weight:600;text-decoration:none}.plist small{display:block;color:var(--muted);margin-top:2px}
.more{font-weight:600}
.faces{display:flex;flex-wrap:wrap;gap:10px;list-style:none;padding:0;margin:1.4rem 0}
.av{width:100%;aspect-ratio:1;border-radius:50%;object-fit:cover;background:var(--alt);display:block}.av--sm{width:58px;height:58px;border:2px solid var(--bg);box-shadow:var(--sh)}.av--lg{width:190px;height:190px;box-shadow:var(--sh)}
.av--ini{display:grid;place-items:center;background:var(--grad);color:#fff;font:700 1.2rem "Space Grotesk",sans-serif}
.logos{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:16px;list-style:none;padding:0;margin:1.6rem 0}
.logos{grid-template-columns:repeat(auto-fill,minmax(240px,1fr))}.logos li{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:16px;text-align:center}.logos .logo{display:grid;place-items:center;background:#fff;border-radius:10px;height:84px;padding:8px;margin-bottom:10px}.logos img{max-height:64px;max-width:100%;object-fit:contain}.logos p{margin:0;font-weight:600;font-size:.92rem}.logos small{display:block;color:var(--muted);margin-top:4px;line-height:1.5}
.banner{background:var(--grad);color:#fff;border-radius:22px;padding:clamp(28px,5vw,52px);display:flex;flex-wrap:wrap;justify-content:space-between;gap:20px;align-items:center}.banner h2{margin:0 0 .3rem;color:#fff}.banner p{margin:0;color:#fff}.banner .btn--pink{background:#0F1A47;color:#fff}.banner .btn--ghost{border-color:#fff}
/* people */
.pcard>a{display:block;text-align:center;text-decoration:none;color:var(--text);background:var(--card);border:1px solid var(--line);border-radius:16px;padding:18px 14px;height:100%;transition:transform .2s,box-shadow .2s}.pcard:hover>a{transform:translateY(-4px);box-shadow:var(--sh)}
.pcard .av{width:104px;height:104px;margin:0 auto 12px}.pcard h3{font-size:1rem;margin-bottom:2px}.pcard__role{margin:0;font-size:.84rem;font-weight:600;color:var(--pink-ink)}
.pcard__inst{margin:.4rem 0 0;font-size:.76rem;color:var(--muted);display:flex;gap:6px;justify-content:center;align-items:flex-start;text-align:left;line-height:1.35}.flag{width:20px;height:14px;object-fit:cover;border-radius:2px;margin-top:2px;flex:none}
.pcard__inst span{display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
.pcard{position:relative}.pcard--l1>a{padding-bottom:50px}.pcard--l2>a{padding-bottom:78px}.pcard--l3>a{padding-bottom:104px}.pcard--links>a::after{display:none}
.pcard__links{position:absolute;left:8px;right:8px;bottom:12px;z-index:2;display:flex;flex-wrap:wrap;gap:5px;justify-content:center;transition:transform .3s cubic-bezier(.2,.7,.2,1)}
.pcard:hover .pcard__links{transform:translateY(-8px)}
.pcard__links a{font:600 .7rem Inter,sans-serif;padding:3px 9px;border-radius:999px;border:1px solid var(--line);background:var(--alt);color:var(--text);text-decoration:none;transition:background .2s,color .2s,border-color .2s}
.pcard__links a:hover,.pcard__links a:focus-visible{background:var(--blue);border-color:var(--blue);color:#fff}
.dots{display:flex;gap:5px;justify-content:center;margin-top:8px;min-height:10px}.dot{width:9px;height:9px;border-radius:50%;background:var(--c)}
.pgroup h2 small{font-size:.9rem;color:var(--muted);font-weight:500}.pgroup{margin-bottom:2.2rem}
.filters{display:flex;flex-wrap:wrap;gap:12px;margin-bottom:12px}
input[type=search],select{padding:11px 16px;border-radius:12px;border:1.5px solid var(--strong);background:var(--card);color:var(--text);font:inherit;min-width:260px}
.fchips{display:flex;flex-wrap:wrap;gap:8px;margin:8px 0}.fchip{border:1.5px solid var(--strong);background:var(--card);color:var(--text);padding:6px 14px;border-radius:999px;cursor:pointer;font:600 .85rem Inter,sans-serif}
.fchip small{opacity:.7}.fchip.is-on{background:var(--navy);border-color:var(--navy);color:#fff}
.notice{background:var(--alt2);border-left:4px solid var(--pink);padding:12px 16px;border-radius:8px}
/* profile */
.profile{display:flex;gap:32px;align-items:center}.profile h1 small{display:block;font-size:1rem;font-weight:500;opacity:.8;font-family:Inter,sans-serif;margin-top:6px}
.profile .av--lg{border:4px solid rgba(255,255,255,.3)}.where{display:flex;gap:8px;align-items:center;margin:.2rem 0}
@media(max-width:700px){.profile{flex-direction:column;text-align:center;align-items:center}.profile .cta,.where{justify-content:center}}
.prose{max-width:820px}.prose h2{margin-top:1.8rem;font-size:1.4rem}.prose h2:first-child{margin-top:0}
.tags--dark li{background:var(--alt)}
.pager{display:flex;justify-content:space-between;gap:12px;margin-top:2.5rem;padding-top:1.2rem;border-top:1px solid var(--line)}.pager a{font-weight:600;text-decoration:none}
/* about */
.values{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px;list-style:none;padding:0}.values li{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px}.values h3{color:var(--pink-ink)}.values p{margin:0;color:var(--muted);font-size:.93rem}
.dir{display:flex;gap:32px;align-items:center;flex-wrap:wrap}
/* publications */
.publist{list-style:none;padding:0;margin:1rem 0}
.pub{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px 20px;margin-bottom:14px}.pub[hidden]{display:none}.pub:target{border-color:var(--blue);box-shadow:0 0 0 3px color-mix(in srgb,var(--blue) 25%,transparent)}
.pub__meta{display:flex;align-items:center;gap:6px;font-size:.82rem;color:var(--muted)}.pub__meta .pill{margin-left:0}.pub__year{margin-left:auto;font-weight:700}
.snav{position:sticky;top:66px;z-index:40;background:var(--bg);border-bottom:1px solid var(--line);box-shadow:0 4px 14px rgba(15,26,71,.06)}
.snav ul{display:flex;gap:6px;list-style:none;margin:0;padding:10px 0;overflow-x:auto;scrollbar-width:thin}
.snav a{display:block;white-space:nowrap;padding:7px 16px;border-radius:999px;font-weight:600;font-size:.88rem;text-decoration:none;color:var(--text);border:1.5px solid var(--line)}
.snav a:hover{border-color:var(--blue-ink);color:var(--blue-ink)}.snav a.on{background:var(--navy);border-color:var(--navy);color:#fff}
.psecs{max-width:900px;padding-bottom:30px}.psec{scroll-margin-top:140px;padding:34px 0;border-bottom:1px solid var(--line)}.psec:last-child{border-bottom:0}
.psec>h2{font-size:1.5rem;position:relative;padding-left:16px}.psec>h2::before{content:"";position:absolute;left:0;top:.12em;bottom:.12em;width:5px;border-radius:3px;background:var(--grad)}
.psec h3{font-size:1.02rem;margin:1.4rem 0 .5rem;color:var(--pink-ink)}.psec ul{padding-left:1.2rem}.psec li{margin:.35rem 0}
.psec p{text-align:justify;hyphens:auto}.psec p.metric-line{text-align:left}.metric-line{display:inline-block;background:var(--alt);padding:6px 14px;border-radius:999px;font-weight:600;font-size:.9rem}
.timeline{list-style:none;padding:0!important}.timeline li{padding-left:0}.timeline .when{display:inline-block;min-width:104px;font-weight:700;color:var(--blue-ink)}
@media(max-width:640px){.snav{top:66px}.timeline .when{display:block}}
.kw{display:flex;flex-wrap:wrap;gap:6px;list-style:none;padding:0;margin:.5rem 0}.kw li{font-size:.76rem;padding:2px 10px;border-radius:999px;background:var(--alt);color:var(--muted);border:1px solid var(--line)}
.abs{margin:.5rem 0}.abs summary{cursor:pointer;font-weight:600;font-size:.88rem;color:var(--blue-ink)}.abs p{font-size:.92rem;color:var(--muted);margin:.5rem 0 0;max-width:80ch}
.pub__act{display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin-top:10px}.pub__act .chips{margin:0}.pbtn{font:600 .82rem Inter,sans-serif;padding:5px 14px;border-radius:10px;border:1.5px solid var(--strong);background:var(--card);color:var(--text);text-decoration:none;cursor:pointer}.pbtn:hover{border-color:var(--blue-ink);color:var(--blue-ink)}.pbtn .doi{font-weight:400;opacity:.75}
@media(max-width:640px){.pbtn .doi{display:none}}
.pub__title{margin:.4rem 0;font-size:1.08rem}.pub__au,.pub__venue{margin:.2rem 0;font-size:.92rem;color:var(--muted)}.pub__au strong{color:var(--text)}.pub__au a.au{text-decoration:none}.pub__au a.au:hover strong{text-decoration:underline;color:var(--blue-ink)}
/* footer */
.ftr{background:var(--hero);color:#C9D4F2;padding:56px 0 24px}.ftr a{color:#E3EBFF}.ftr h4{color:#fff;font-size:.95rem}.ftr ul{list-style:none;padding:0;margin:0}.ftr li{margin:6px 0;font-size:.92rem}
.ftr__grid{display:grid;grid-template-columns:1.6fr 1fr 1fr 1.3fr;gap:32px}@media(max-width:860px){.ftr__grid{grid-template-columns:1fr 1fr}}@media(max-width:520px){.ftr__grid{grid-template-columns:1fr}}
.ftr__bar{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;border-top:1px solid rgba(255,255,255,.14);margin-top:32px;padding-top:18px;font-size:.85rem}
.ftr .btn{margin-top:10px;color:#fff}:root[data-theme=dark] .ftr .btn--pink{color:#0A0F1F}
.reveal{opacity:0;transform:translateY(14px);transition:opacity .5s,transform .5s}.reveal.in{opacity:1;transform:none}
@media(prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important;scroll-behavior:auto!important}.reveal{opacity:1;transform:none}}
.js-off .reveal{opacity:1;transform:none}
/* ===== polish: motion, hover, graphics ===== */
@keyframes rise{from{opacity:0;transform:translateY(22px)}to{opacity:1;transform:none}}
@keyframes drift{0%,100%{transform:translate(0,0) scale(1)}50%{transform:translate(40px,28px) scale(1.12)}}
@keyframes drift2{0%,100%{transform:translate(0,0) scale(1)}50%{transform:translate(-46px,-30px) scale(1.1)}}
@keyframes shimmer{to{background-position:200% 0}}
@keyframes floaty{0%,100%{transform:translateY(0)}50%{transform:translateY(-7px)}}
@keyframes pulse{0%{box-shadow:0 0 0 0 rgba(232,62,140,.5)}100%{box-shadow:0 0 0 14px rgba(232,62,140,0)}}
@keyframes fadein{from{opacity:0}to{opacity:1}}
@keyframes bg{to{background-position:100% 100%}}
body{animation:fadein .5s ease both}
#prog{position:fixed;left:0;top:0;height:3px;width:100%;transform:scaleX(0);transform-origin:left;background:var(--grad);z-index:100;pointer-events:none}
.hdr{transition:box-shadow .25s,background .25s}.hdr.is-scrolled{box-shadow:0 8px 28px rgba(5,10,35,.35);background:rgba(15,26,71,.92);-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px)}
.brand svg{transition:transform .5s cubic-bezier(.3,1.4,.5,1)}.brand:hover svg{transform:rotate(-12deg) scale(1.12)}
.nav a:not(.btn){position:relative;transition:background .2s}
.nav a:not(.btn)::after{content:"";position:absolute;left:14px;right:14px;bottom:4px;height:2px;border-radius:2px;background:var(--grad);transform:scaleX(0);transform-origin:left;transition:transform .3s}
.nav a:not(.btn):hover::after,.nav a[aria-current]::after{transform:scaleX(1)}
.btn{position:relative;overflow:hidden;transition:transform .2s,box-shadow .25s,filter .2s}
.btn::after{content:"";position:absolute;top:0;left:-70%;width:50%;height:100%;background:linear-gradient(100deg,transparent,rgba(255,255,255,.35),transparent);transform:skewX(-20deg);transition:left .6s}
.btn:hover::after{left:130%}.btn:active{transform:translateY(0) scale(.97)}
.btn--pink{background:linear-gradient(135deg,var(--pink-btn),#9b2fae)}.btn--blue{background:linear-gradient(135deg,var(--blue-btn),#5a4fe0)}
.btn--pink:hover{box-shadow:0 10px 26px rgba(232,62,140,.4)}.btn--blue:hover{box-shadow:0 10px 26px rgba(58,123,255,.4)}
.btn--ghost:hover{background:rgba(255,255,255,.14)}
:root[data-theme=dark] .btn--pink{background:var(--pink-btn)}:root[data-theme=dark] .btn--blue{background:var(--blue-btn)}
.hero{background-size:200% 200%;animation:bg 16s ease-in-out infinite alternate}
.hero::before{animation:drift 14s ease-in-out infinite}.hero::after{animation:drift2 17s ease-in-out infinite}
.hero__in{background-image:radial-gradient(rgba(255,255,255,.09) 1.2px,transparent 1.4px);background-size:28px 28px}
.hero__in>*{animation:rise .8s cubic-bezier(.2,.7,.2,1) both}.hero__in>*:nth-child(2){animation-delay:.12s}.hero__in>*:nth-child(3){animation-delay:.24s}.hero__in>*:nth-child(4){animation-delay:.36s}.hero__in>*:nth-child(5){animation-delay:.48s}
.badge{background:rgba(255,255,255,.08);-webkit-backdrop-filter:blur(6px);backdrop-filter:blur(6px);position:relative;padding-left:30px}
.badge::before{content:"";position:absolute;left:12px;top:50%;width:8px;height:8px;margin-top:-4px;border-radius:50%;background:var(--pink);animation:pulse 2s infinite}
.grad{background:linear-gradient(90deg,#FF8FC3,#8FB4FF,#FF8FC3);background-size:200% 100%;-webkit-background-clip:text;background-clip:text;animation:shimmer 6s linear infinite}
.phead{position:relative;overflow:hidden}.phead::before{content:"";position:absolute;right:-120px;top:-140px;width:420px;height:420px;border-radius:50%;background:radial-gradient(circle,rgba(232,62,140,.35),transparent 65%);animation:drift 15s ease-in-out infinite;pointer-events:none}
.phead::after{content:"";position:absolute;left:-100px;bottom:-180px;width:380px;height:380px;border-radius:50%;background:radial-gradient(circle,rgba(58,123,255,.35),transparent 65%);animation:drift2 18s ease-in-out infinite;pointer-events:none}
.phead>*{position:relative;z-index:1}
.stat{transition:transform .25s}.stat:hover{transform:translateY(-4px)}.stat+.stat{border-left:1px solid rgba(255,255,255,.1)}
.eyebrow{display:flex;align-items:center;gap:10px}.eyebrow::before{content:"";width:26px;height:3px;border-radius:2px;background:var(--grad);flex:none}
.sec--ice,.sec--blush{position:relative}
.sec--ice::before{content:"";position:absolute;inset:0;background-image:radial-gradient(color-mix(in srgb,var(--blue) 14%,transparent) 1.2px,transparent 1.4px);background-size:26px 26px;opacity:.7;pointer-events:none}
.sec--ice>.wrap{position:relative}
.reveal{opacity:0;transform:none;transition:transform .25s cubic-bezier(.2,.7,.2,1),box-shadow .25s,border-color .25s}
.reveal.in{opacity:1;animation:rise .7s cubic-bezier(.2,.7,.2,1) var(--d,0s) backwards}
.spot{position:relative;overflow:hidden}
.spot::before{content:"";position:absolute;inset:0;background:radial-gradient(280px circle at var(--mx,50%) var(--my,50%),color-mix(in srgb,var(--c,var(--blue)) 18%,transparent),transparent 62%);opacity:0;transition:opacity .3s;pointer-events:none}
.spot:hover::before{opacity:1}
.area::after{content:"\2192";position:absolute;right:20px;bottom:16px;font-weight:700;color:var(--c);opacity:0;transform:translateX(-8px);transition:.3s}
.area:hover{transform:translateY(-6px);box-shadow:0 18px 38px color-mix(in srgb,var(--c) 28%,transparent);border-color:color-mix(in srgb,var(--c) 45%,var(--line))}
.area:hover::after{opacity:1;transform:none}
.area__ico{transition:transform .4s cubic-bezier(.3,1.5,.5,1),background .3s,color .3s}.area:hover .area__ico{transform:scale(1.15) rotate(-6deg);background:var(--c);color:#fff}
.area__ico--xl{animation:floaty 5s ease-in-out infinite}
.pcard>a{position:relative;overflow:hidden;transition:transform .3s cubic-bezier(.2,.7,.2,1),box-shadow .3s,border-color .3s}
.pcard:hover>a{transform:translateY(-8px);box-shadow:0 20px 40px rgba(58,123,255,.22);border-color:var(--blue)}
.pcard>a::after{content:"View profile \2192";position:absolute;left:0;right:0;bottom:0;padding:9px 0;font:600 .8rem Inter,sans-serif;color:#fff;background:var(--grad);transform:translateY(100%);transition:transform .3s cubic-bezier(.2,.7,.2,1)}
.pcard:hover>a::after,.pcard>a:focus-visible::after{transform:none}
.pcard .av{transition:transform .4s cubic-bezier(.3,1.4,.5,1),box-shadow .3s;box-shadow:0 0 0 3px var(--card),0 0 0 5px var(--line)}
.pcard:hover>a .av{transform:scale(1.07);box-shadow:0 0 0 3px var(--card),0 0 0 6px var(--pink),0 12px 26px rgba(232,62,140,.35)}
.pcard h3{transition:color .2s}.pcard:hover>a h3{color:var(--blue-ink)}
.dot{transition:transform .3s}.pcard:hover>a .dot{transform:scale(1.4)}
.faces .av--sm{transition:transform .3s cubic-bezier(.3,1.5,.5,1),box-shadow .3s;position:relative}.faces li:hover{z-index:5;position:relative}
.faces .av--sm:hover{transform:translateY(-8px) scale(1.18);box-shadow:0 0 0 3px var(--pink),0 12px 24px rgba(232,62,140,.4)}
.profile .av--lg{animation:floaty 6s ease-in-out infinite;box-shadow:0 0 0 4px rgba(255,255,255,.18),0 0 0 9px rgba(232,62,140,.28),0 18px 40px rgba(0,0,0,.35);border:0}
.chip{transition:transform .2s,background .2s,box-shadow .2s}.chip:hover{transform:translateY(-2px);background:color-mix(in srgb,var(--c,var(--blue)) 22%,transparent);box-shadow:var(--sh)}
.tags li,.kw li{transition:transform .2s,background .2s}.tags li:hover,.kw li:hover{transform:translateY(-2px);background:color-mix(in srgb,var(--blue) 18%,var(--alt))}
.snav a{transition:all .25s}.snav a:hover{transform:translateY(-2px)}
.timeline li{transition:transform .2s}.timeline li:hover{transform:translateX(4px)}
.panel,.values li,.pub{transition:transform .3s cubic-bezier(.2,.7,.2,1),box-shadow .3s,border-color .3s}
.panel:hover,.values li:hover{transform:translateY(-4px);box-shadow:0 16px 34px rgba(58,123,255,.16)}
.values li{border-top:3px solid var(--pink)}
.pub{border-left:4px solid transparent}.pub:hover{transform:translateX(4px);border-left-color:var(--pink);box-shadow:var(--sh)}
.pbtn{transition:all .2s}.pbtn:hover{transform:translateY(-2px);box-shadow:var(--sh)}
.plist li{transition:padding-left .25s}.plist li:hover{padding-left:10px}
.fchip{transition:all .2s}.fchip:hover{transform:translateY(-2px);border-color:var(--blue-ink)}.fchip.is-on{box-shadow:0 6px 16px rgba(26,42,108,.3)}
input[type=search],select{transition:border-color .2s,box-shadow .2s}input[type=search]:focus,select:focus{outline:0;border-color:var(--blue);box-shadow:0 0 0 4px color-mix(in srgb,var(--blue) 22%,transparent)}
.banner{background-size:200% 200%;animation:bg 8s ease-in-out infinite alternate;position:relative;overflow:hidden}
.banner::before{content:"";position:absolute;right:-60px;top:-80px;width:260px;height:260px;border-radius:50%;background:rgba(255,255,255,.14)}
.banner>*{position:relative}
.more{text-decoration:none;display:inline-block;transition:transform .2s}.more:hover{transform:translateX(5px)}
.ftr a{transition:color .2s,padding-left .2s}.ftr li a:hover{color:var(--pink-soft);padding-left:4px}
.ftr{position:relative;overflow:hidden}.ftr::before{content:"";position:absolute;left:0;right:0;top:0;height:3px;background:var(--grad)}
#top{position:fixed;right:20px;bottom:20px;width:46px;height:46px;border:0;border-radius:50%;background:var(--grad);color:#fff;cursor:pointer;display:grid;place-items:center;box-shadow:0 8px 22px rgba(232,62,140,.4);opacity:0;transform:translateY(16px) scale(.8);pointer-events:none;transition:all .3s;z-index:60}
#top.show{opacity:1;transform:none;pointer-events:auto}#top:hover{transform:translateY(-4px)}#top svg{width:22px;height:22px}
@media(hover:none){.pcard>a::after{display:none}.spot::before{display:none}}
/* butterflies */
#bfly{position:absolute;inset:0;pointer-events:none;z-index:2;overflow:hidden}
.bf{position:absolute;left:0;top:0;width:46px;height:40px;will-change:transform;filter:drop-shadow(0 3px 5px rgba(15,26,71,.28))}
.bf svg{width:100%;height:100%;overflow:visible}
.bf .wl,.bf .wr{transform-box:fill-box;animation:flap var(--f,.32s) ease-in-out infinite alternate}
.bf .wl{transform-origin:100% 50%}.bf .wr{transform-origin:0% 50%}
@keyframes flap{from{transform:scaleX(1)}to{transform:scaleX(.18)}}

.hdr .brand{position:relative}
.bf-logo{position:absolute;left:-9px;top:-5px;width:26px;height:22px;pointer-events:none;transform-origin:50% 80%;animation:perch 4.5s ease-in-out infinite;filter:drop-shadow(0 2px 4px rgba(0,0,0,.35))}
.bf-logo svg{width:100%;height:100%;overflow:visible}
.bf-logo .wl,.bf-logo .wr{transform-box:fill-box;animation:flap .42s ease-in-out infinite alternate}
.bf-logo .wl{transform-origin:100% 50%}.bf-logo .wr{transform-origin:0% 50%}
@keyframes perch{0%,100%{transform:translate(0,0) rotate(-24deg)}30%{transform:translate(2px,-3px) rotate(-14deg)}60%{transform:translate(-1px,-1px) rotate(-30deg)}}
@media print{#bfly{display:none}}
'''

JS_SITE = r'''(function(){
var d=document,r=d.documentElement;
var th=d.getElementById('theme');
if(th)th.addEventListener('click',function(){var dark=r.dataset.theme?r.dataset.theme==='dark':matchMedia('(prefers-color-scheme:dark)').matches;var n=dark?'light':'dark';r.dataset.theme=n;try{localStorage.setItem('syntera-theme',n)}catch(e){}});
var b=d.getElementById('burger'),nav=d.querySelector('.hdr nav');
if(b)b.addEventListener('click',function(){var o=nav.classList.toggle('open');b.setAttribute('aria-expanded',o)});
d.addEventListener('keydown',function(e){if(e.key==='Escape'&&nav){nav.classList.remove('open');b&&b.setAttribute('aria-expanded',false)}});
d.querySelectorAll('.pcard,.values li,.panel,.banner,.phead__row,.dir').forEach(function(x){x.classList.add('reveal')});
d.querySelectorAll('.grid,.values,.faces,.split,.pgroup').forEach(function(g){var i=0;g.querySelectorAll(':scope>.reveal,:scope>li.reveal').forEach(function(c){c.style.setProperty('--d',Math.min(i++,10)*.05+'s')})});
if('IntersectionObserver' in window){var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('in');io.unobserve(e.target)}})},{threshold:.1});d.querySelectorAll('.reveal').forEach(function(x){io.observe(x)})}else d.querySelectorAll('.reveal').forEach(function(x){x.classList.add('in')});
var reduce=matchMedia('(prefers-reduced-motion:reduce)').matches;
d.querySelectorAll('[data-count]').forEach(function(el){var n=+el.dataset.count;if(reduce)return;var t0=null;el.textContent='0';function f(t){t0=t0||t;var p=Math.min((t-t0)/900,1);el.textContent=Math.round(n*p);if(p<1)requestAnimationFrame(f)}requestAnimationFrame(f)});
var sn=d.querySelectorAll('.snav a');
if(sn.length&&'IntersectionObserver' in window){var map={};sn.forEach(function(a){map[a.getAttribute('href').slice(1)]=a});
var so=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){sn.forEach(function(a){a.classList.remove('on')});map[e.target.id].classList.add('on')}})},{rootMargin:'-140px 0px -65% 0px'});
d.querySelectorAll('.psec').forEach(function(x){so.observe(x)})}

var hd=d.querySelector('.hdr'),bar=d.createElement('div');bar.id='prog';d.body.appendChild(bar);
var tp=d.createElement('button');tp.id='top';tp.type='button';tp.setAttribute('aria-label','Back to top');tp.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M12 19V5M5 12l7-7 7 7"/></svg>';d.body.appendChild(tp);
tp.addEventListener('click',function(){scrollTo({top:0,behavior:reduce?'auto':'smooth'})});
var tick=false;function onS(){var h=d.documentElement,m=h.scrollHeight-h.clientHeight,y=h.scrollTop;bar.style.transform='scaleX('+(m>0?y/m:0)+')';hd&&hd.classList.toggle('is-scrolled',y>10);tp.classList.toggle('show',y>600);tick=false}
addEventListener('scroll',function(){if(!tick){tick=true;requestAnimationFrame(onS)}},{passive:true});onS();
d.querySelectorAll('.area,.pcard>a,.pub,.panel,.values li').forEach(function(x){x.classList.add('spot');x.addEventListener('pointermove',function(e){var r=x.getBoundingClientRect();x.style.setProperty('--mx',(e.clientX-r.left)+'px');x.style.setProperty('--my',(e.clientY-r.top)+'px')})});


(function(){var b=d.querySelector('.hdr .brand');if(!b||reduce)return;var e=d.createElement('span');e.className='bf-logo';e.setAttribute('aria-hidden','true');
e.innerHTML='<svg viewBox="-20 -16 40 32"><g class="wl"><path d="M0 0C-6-14-18-15-19-7-20-1-10 2 0 0Z" fill="#FF8FC3"/><path d="M0 1C-8 3-15 9-11 13-7 16-1 9 0 1Z" fill="#E83E8C"/><circle cx="-11" cy="-6" r="2.2" fill="#fff" opacity=".75"/></g><g class="wr"><path d="M0 0C6-14 18-15 19-7 20-1 10 2 0 0Z" fill="#FF8FC3"/><path d="M0 1C8 3 15 9 11 13 7 16 1 9 0 1Z" fill="#E83E8C"/><circle cx="11" cy="-6" r="2.2" fill="#fff" opacity=".75"/></g><rect x="-1" y="-6" width="2" height="14" rx="1" fill="#fff"/></svg>';
b.appendChild(e)})();
var hero=d.querySelector('.hero');
if(!reduce&&hero){(function(){
var cols=[['#FF8FC3','#E83E8C'],['#8FB4FF','#3A7BFF'],['#C4A3FF','#8B5CF6'],['#FFC2DE','#E83E8C'],['#9ED0FF','#2D66E6']];
var n=innerWidth<700?3:5,L=d.createElement('div');L.id='bfly';L.setAttribute('aria-hidden','true');hero.appendChild(L);var W=0,H=0;function sz(){W=L.offsetWidth;H=L.offsetHeight}sz();addEventListener('resize',sz);
var mx=-999,my=-999;hero.addEventListener('pointermove',function(e){var r=L.getBoundingClientRect();mx=e.clientX-r.left;my=e.clientY-r.top},{passive:true});hero.addEventListener('pointerleave',function(){mx=my=-999});
function svg(c){return '<svg viewBox="-20 -16 40 32"><g class="wl"><path d="M0 0C-6-14-18-15-19-7-20-1-10 2 0 0Z" fill="'+c[0]+'"/><path d="M0 1C-8 3-15 9-11 13-7 16-1 9 0 1Z" fill="'+c[1]+'"/><circle cx="-11" cy="-6" r="2.2" fill="#fff" opacity=".75"/></g><g class="wr"><path d="M0 0C6-14 18-15 19-7 20-1 10 2 0 0Z" fill="'+c[0]+'"/><path d="M0 1C8 3 15 9 11 13 7 16 1 9 0 1Z" fill="'+c[1]+'"/><circle cx="11" cy="-6" r="2.2" fill="#fff" opacity=".75"/></g><rect x="-1" y="-6" width="2" height="14" rx="1" fill="#14213D"/></svg>'}
var B=[];
for(var i=0;i<n;i++){var el=d.createElement('div');el.className='bf';el.style.setProperty('--f',(.26+Math.random()*.16)+'s');el.innerHTML=svg(cols[i%cols.length]);L.appendChild(el);
B.push({el:el,x:Math.random()*W,y:Math.random()*H,tx:0,ty:0,a:0,sp:.9+Math.random()*.9,ph:Math.random()*6,sc:.7+Math.random()*.6,rest:0})}
function pick(b){b.tx=Math.random()*W;b.ty=30+Math.random()*Math.max(H-80,10)}
B.forEach(pick);
var t=0,run=true,vis=true,inv=true;
function step(){if(!run)return;t+=.03;
B.forEach(function(b){
if(b.rest>0){b.rest--;b.el.style.setProperty('--f','.9s')}else{b.el.style.setProperty('--f','.3s')
var dx=b.tx-b.x,dy=b.ty-b.y,dist=Math.hypot(dx,dy);
var fx=b.x-mx,fy=b.y-my,fd=Math.hypot(fx,fy);
if(fd<130){b.x+=fx/fd*3.2;b.y+=fy/fd*3.2}
if(dist<30){pick(b);if(Math.random()<.35)b.rest=90+Math.random()*120}
else{var ang=Math.atan2(dy,dx)+Math.sin(t*2+b.ph)*.9;b.x+=Math.cos(ang)*b.sp;b.y+=Math.sin(ang)*b.sp+Math.sin(t*5+b.ph)*.5;b.a=ang}}
var rot=(Math.cos(b.a)*14);
b.el.style.transform='translate('+b.x.toFixed(1)+'px,'+b.y.toFixed(1)+'px) rotate('+rot.toFixed(1)+'deg) scale('+b.sc+')'});
requestAnimationFrame(step)}
function upd(){var r=vis&&inv;if(r&&!run){run=true;step()}else run=r}
d.addEventListener('visibilitychange',function(){vis=!d.hidden;upd()});
if('IntersectionObserver' in window)new IntersectionObserver(function(e){inv=e[0].isIntersecting;upd()}).observe(hero);
step();
})()}
})();'''

JS_HERO = r'''(function(){
var c=document.getElementById('net');if(!c||matchMedia('(prefers-reduced-motion:reduce)').matches||innerWidth<700)return;
var x=c.getContext('2d'),W,H,P=[],mx=-999,my=-999,run=true;
function size(){W=c.width=c.offsetWidth;H=c.height=c.offsetHeight;P=[];var n=Math.round(W*H/16000);for(var i=0;i<n;i++)P.push({x:Math.random()*W,y:Math.random()*H,vx:(Math.random()-.5)*.35,vy:(Math.random()-.5)*.35,k:i%2})}
size();addEventListener('resize',size);
c.parentNode.addEventListener('pointermove',function(e){var r=c.getBoundingClientRect();mx=e.clientX-r.left;my=e.clientY-r.top});
new IntersectionObserver(function(e){run=e[0].isIntersecting;if(run)loop()}).observe(c);
function loop(){if(!run)return;x.clearRect(0,0,W,H);
for(var i=0;i<P.length;i++){var a=P[i];a.x+=a.vx;a.y+=a.vy;if(a.x<0||a.x>W)a.vx*=-1;if(a.y<0||a.y>H)a.vy*=-1;
for(var j=i+1;j<P.length;j++){var b=P[j],dx=a.x-b.x,dy=a.y-b.y,d=dx*dx+dy*dy;if(d<17000){x.strokeStyle='rgba('+(a.k?'143,180,255':'255,143,195')+','+(1-d/17000)*.45+')';x.beginPath();x.moveTo(a.x,a.y);x.lineTo(b.x,b.y);x.stroke()}}
var near=(a.x-mx)*(a.x-mx)+(a.y-my)*(a.y-my)<9000;x.fillStyle=near?'#FF8FC3':a.k?'#8FB4FF':'#FF8FC3';x.beginPath();x.arc(a.x,a.y,near?3.2:2,0,6.3);x.fill()}
requestAnimationFrame(loop)}loop();
})();'''

JS_PEOPLE = r'''(function(){
var q=document.getElementById('q'),chips=document.querySelectorAll('#fchips .fchip'),cards=document.querySelectorAll('.pcard'),area='all';
function run(){var s=q.value.trim().toLowerCase(),shown=0;
cards.forEach(function(c){var ok=(area==='all'||c.dataset.areas.split(' ').indexOf(area)>-1)&&(!s||c.dataset.name.indexOf(s)>-1);c.hidden=!ok;if(ok)shown++});
document.querySelectorAll('.pgroup').forEach(function(g){g.hidden=!g.querySelector('.pcard:not([hidden])')});
document.getElementById('none').hidden=shown>0}
q.addEventListener('input',run);
chips.forEach(function(b){b.addEventListener('click',function(){chips.forEach(function(x){x.classList.remove('is-on')});b.classList.add('is-on');area=b.dataset.area;run()})});
var p=new URLSearchParams(location.search).get('area');if(p){var b=document.querySelector('#fchips [data-area="'+p+'"]');if(b)b.click()}
})();'''

JS_PUBS = r'''(function(){
var D=window.SYNTERA_PUBS||[],A=window.SYNTERA_AREAS||{},B=document.body.dataset.base||'';
function esc(t){return String(t).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})}
var T={journal:'Journal',conference:'Conference',book:'Book',chapter:'Chapter',thesis:'Thesis',preprint:'Preprint',other:'Other'};
function card(p){
 var au=p.authors.map(function(a){return a[1]?'<a class="au" href="/people/'+a[1]+'"><strong>'+esc(a[0])+'</strong></a>':esc(a[0])}).join('; ');
 var chips=p.areas.map(function(a){return '<a class="chip" style="--c:'+A[a].color+'" href="/research/'+a+'">'+A[a].short+'</a>'}).join('');
 var kw=(p.keywords||[]).map(function(k){return '<li>'+esc(k)+'</li>'}).join('');
 var li=document.createElement('li');li.className='pub';li.id='pub-'+p.id;
 li.innerHTML='<div class="pub__meta"><span class="pill pill--'+p.type+'">'+T[p.type]+'</span>'+(p.note?'<span class="pill">'+esc(p.note)+'</span>':'')+'<span class="pub__year">'+(p.year||'n.d.')+'</span></div><h3 class="pub__title">'+esc(p.title)+'</h3><p class="pub__au">'+au+'</p><p class="pub__venue"><em>'+esc(p.venue)+'</em></p>'
 +(kw?'<ul class="kw" aria-label="Keywords">'+kw+'</ul>':'')
 +(p.abstract?'<details class="abs"><summary>Abstract</summary><p>'+esc(p.abstract)+'</p></details>':'')
 +'<div class="pub__act">'+(p.doi?'<a class="pbtn" href="https://doi.org/'+esc(p.doi)+'" rel="noopener" target="_blank">DOI<span class="doi"> '+esc(p.doi)+'</span></a>':'')+'<button class="pbtn" type="button" data-cite="'+p.id+'">Cite (BibTeX)</button>'+chips+'</div>';
 return li}
function bib(p){var au=p.authors.filter(function(a){return a[0]!=='…'}).map(function(a){return a[0]}).join(' and ');
 var t={journal:'article',conference:'inproceedings',book:'book',chapter:'incollection',thesis:'phdthesis',preprint:'unpublished',other:'misc'}[p.type];
 var vf={journal:'journal',conference:'booktitle',book:'publisher',chapter:'booktitle',thesis:'school',preprint:'note',other:'howpublished'}[p.type];
 return '@'+t+'{'+p.id.replace(/-/g,'')+p.year+',\n  author = {'+au+'},\n  title = {'+p.title+'},\n  '+vf+' = {'+p.venue+'},\n  year = {'+p.year+'}'+(p.doi?',\n  doi = {'+p.doi+'}':'')+'\n}'}
document.addEventListener('click',function(e){var b=e.target.closest('[data-cite]');if(!b)return;var p=D.filter(function(x){return x.id===b.dataset.cite})[0],old=b.textContent;
 function done(){b.textContent='Copied';setTimeout(function(){b.textContent=old},1500)}
 if(navigator.clipboard)navigator.clipboard.writeText(bib(p)).then(done,function(){prompt('BibTeX',bib(p))});else prompt('BibTeX',bib(p))});
function fill(box,list){box.innerHTML='';list.forEach(function(p){box.appendChild(card(p))})}
document.querySelectorAll('[data-pubs]').forEach(function(box){
 if(box.id==='pubs')return;
 var l=D.filter(function(p){return (!box.dataset.member||p.authors.some(function(a){return a[1]===box.dataset.member}))&&(!box.dataset.area||p.areas.indexOf(box.dataset.area)>-1)});
 if(box.dataset.limit)l=l.slice(0,+box.dataset.limit);fill(box,l)});
var box=document.getElementById('pubs');if(!box)return;
var q=document.getElementById('q'),yr=document.getElementById('year'),st={type:'all',area:'all'};
var hay=D.map(function(p){return (p.title+' '+p.authors.map(function(a){return a[0]}).join(' ')+' '+p.venue+' '+(p.keywords||[]).join(' ')).toLowerCase()});
var tc=document.querySelectorAll('#tchips .fchip'),ac=document.querySelectorAll('#achips .fchip'),first=true;
function run(){var s=q.value.trim().toLowerCase();
 var l=D.filter(function(p,i){return (st.type==='all'||p.type===st.type)&&(st.area==='all'||p.areas.indexOf(st.area)>-1)&&(yr.value==='all'||String(p.year)===yr.value)&&(!s||hay[i].indexOf(s)>-1)});
 fill(box,l);document.getElementById('count').textContent=l.length+' publication'+(l.length===1?'':'s');document.getElementById('none').hidden=l.length>0;
 if(first&&location.hash){var t=document.getElementById(location.hash.slice(1));if(t)t.scrollIntoView()}first=false}
function bind(list,key,attr){list.forEach(function(b){b.addEventListener('click',function(){list.forEach(function(x){x.classList.remove('is-on')});b.classList.add('is-on');st[key]=b.dataset[attr];run()})})}
bind(tc,'type','type');bind(ac,'area','area');q.addEventListener('input',run);yr.addEventListener('change',run);
var u=new URLSearchParams(location.search);
if(u.get('area')){var b=document.querySelector('#achips [data-area="'+u.get('area')+'"]');if(b){ac.forEach(function(x){x.classList.remove('is-on')});b.classList.add('is-on');st.area=u.get('area')}}
if(u.get('type')){var t=document.querySelector('#tchips [data-type="'+u.get('type')+'"]');if(t){tc.forEach(function(x){x.classList.remove('is-on')});t.classList.add('is-on');st.type=u.get('type')}}
if(u.get('q'))q.value=u.get('q');
if(u.get('year'))yr.value=u.get('year');
run();
})();'''

FAVICON = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 40"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#E83E8C"/><stop offset="1" stop-color="#3A7BFF"/></linearGradient></defs><rect width="40" height="40" rx="10" fill="#0F1A47"/><path d="M28 12c-2-3-14-3-14 3 0 7 14 3 14 10 0 6-12 6-15 2" fill="none" stroke="url(#g)" stroke-width="3" stroke-linecap="round"/><circle cx="28" cy="12" r="3" fill="#FF8FC3"/><circle cx="13" cy="27" r="3" fill="#8FB4FF"/></svg>'

def make_og_image():
    from PIL import Image, ImageDraw, ImageFont
    W, H = 1200, 630
    im = Image.new('RGB', (W, H), '#0F1A47'); d = ImageDraw.Draw(im)
    for i in range(H):   # vertical gradient
        t = i / H; d.line([(0, i), (W, i)], fill=(int(15 + 25 * t), int(26 + 10 * t), int(71 + 60 * t)))
    for x, y, r, c in [(980, 150, 190, (232, 62, 140)), (1080, 470, 140, (58, 123, 255)), (820, 560, 90, (255, 143, 195))]:
        ov = Image.new('RGBA', (W, H), (0, 0, 0, 0)); ImageDraw.Draw(ov).ellipse([x - r, y - r, x + r, y + r], fill=c + (70,))
        im.paste(ov, (0, 0), ov)
    d = ImageDraw.Draw(im)
    def font(sz, bold=True):
        for n in (['arialbd.ttf', 'segoeuib.ttf'] if bold else ['arial.ttf', 'segoeui.ttf']) + ['DejaVuSans-Bold.ttf' if bold else 'DejaVuSans.ttf']:
            try: return ImageFont.truetype(n, sz)
            except OSError: pass
        return ImageFont.load_default()
    d.text((80, 190), 'SYNTERA', font=font(130), fill='white')
    d.text((86, 340), 'Research Group', font=font(54, False), fill=(255, 143, 195))
    d.text((86, 430), 'Applied AI & Connected Systems', font=font(40, False), fill=(200, 210, 240))
    d.text((86, 540), 'syntera.au', font=font(34), fill=(143, 180, 255))
    os.makedirs(os.path.join(OUT, 'images'), exist_ok=True)
    im.save(os.path.join(OUT, 'images', 'og-image.png'), optimize=True)

def write(path, txt):
    p = os.path.join(OUT, path); os.makedirs(os.path.dirname(p), exist_ok=True)
    open(p, 'w', encoding='utf-8').write(txt)

def main():
    if os.path.exists(OUT): shutil.rmtree(OUT)
    os.makedirs(OUT)
    members = build_members()
    process_photos(members)
    build_member_keys(members)
    derive_areas(members)
    write('data/publications.js', export_pubs())
    for sub in ('flags', 'partners'):
        shutil.copytree(os.path.join(ROOT, 'images', sub), os.path.join(OUT, 'images', sub))
    write('css/style.css', CSS); write('js/site.js', JS_SITE); write('js/hero.js', JS_HERO)
    write('js/people.js', JS_PEOPLE); write('js/pubs.js', JS_PUBS); write('images/favicon.svg', FAVICON)
    global VER
    VER = hashlib.md5((CSS + JS_SITE + JS_PUBS + JS_PEOPLE + JS_HERO + open(os.path.join(OUT, 'data', 'publications.js'), encoding='utf-8').read()).encode()).hexdigest()[:8]
    build_pages(members)
    used = set(re.findall(r'images/partners/([\w-]+\.webp)', open(os.path.join(OUT, 'index.html'), encoding='utf-8').read()))
    for f in os.listdir(os.path.join(OUT, 'images', 'partners')):
        if f not in used: os.remove(os.path.join(OUT, 'images', 'partners', f))
    write('.nojekyll', '')
    today = __import__('datetime').date.today().isoformat()
    urls = ''.join('<url><loc>%s/%s</loc><lastmod>%s</lastmod><priority>%s</priority></url>' % (SITE_URL, '' if f == 'index.html' else re.sub(r'\.html$', '', f), today, '1.0' if f == 'index.html' else '0.9' if f in ('collaborate.html', 'research.html') else '0.8' if '/' not in f else '0.6') for f in sorted(PAGES))
    write('sitemap.xml', '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">' + urls + '</urlset>')
    write('robots.txt', 'User-agent: *\nAllow: /\n\nSitemap: %s/sitemap.xml\n' % SITE_URL)
    write('llms.txt', f"""# {BRAND['name']}

> Applied AI and connected systems research group led from the University of the Sunshine Coast, Queensland, Australia. Open to academic collaboration, joint grant applications, industry-funded and commissioned (paid) research projects, industry pilots and student projects.

## Key pages
- [Collaborate with us]({SITE_URL}/collaborate): ways to partner, process and FAQ
- [Research areas]({SITE_URL}/research): AI for health, smart health home, AI for agriculture, AI for education, Internet of Things, Internet of Vehicles
- [Publications]({SITE_URL}/publications): journal articles, conference papers, books and chapters
- [People]({SITE_URL}/people): director, researchers, academics and students
- [About]({SITE_URL}/about)
- [Join us]({SITE_URL}/join): PhD, Master's, postdoc, research assistant and internship opportunities
- [Contact]({SITE_URL}/contact)

## Contact
- Collaborations, funding and paid projects: {BRAND['email_collab']}
- Prospective students: {BRAND['email_student']}
- Joining the group: {BRAND['email_join']}
""")
    write('_headers', """/*
  X-Content-Type-Options: nosniff
  Referrer-Policy: strict-origin-when-cross-origin
  X-Frame-Options: SAMEORIGIN
  Permissions-Policy: camera=(), microphone=(), geolocation=()
  Strict-Transport-Security: max-age=31536000; includeSubDomains

/css/*
  Cache-Control: public, max-age=31536000, immutable
/js/*
  Cache-Control: public, max-age=31536000, immutable
/images/*
  Cache-Control: public, max-age=2592000
/data/*
  Cache-Control: public, max-age=31536000, immutable
""")
    make_og_image()
    page('404.html', 'Page not found', 'This page could not be found.', '<section class="sec"><div class="wrap" style="text-align:center;padding:80px 0"><h1>Page not found</h1><p>The page you are looking for does not exist or has moved.</p><p><a class="btn btn--pink" href="/">Back to the home page</a></p></div></section>')
    n = sum(len(f) for _, _, f in os.walk(OUT))
    print(f'{len(members)} members, {len(PUBS)} publications, {n} files -> {OUT}')

if __name__ == '__main__':
    main()
