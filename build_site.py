#!/usr/bin/env python3
"""SYNTERA Research Group static site generator. Run: python build_site.py  ->  site/"""
import html, json, os, re, shutil, zipfile, xml.etree.ElementTree as ET
from PIL import Image, ImageOps

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, 'docs')
E = html.escape

# ───────────────────────── brand ─────────────────────────
BRAND = dict(
    name='SYNTERA Research Group', short='SYNTERA',
    full='SYNTERA Research Group: Applied AI & Connected Systems',
    tagline='Intelligence in Synergy',
    host='University of the Sunshine Coast', host_short='UniSC',
    host_url='https://www.usc.edu.au/',
    email='ssaremi@usc.edu.au',
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

FLAGS = {'Australia':'au','Pakistan':'pk','Bangladesh':'bd','Iran':'ir','Argentina':'ar','Norway':'no','Iraq':'iq','China':'cn','United Kingdom':'gb'}
# Photo file per member id
PHOTOS = {
 'shahrzad-saremi':'shahrzad saremi.jpeg','rania-shibl':'rania shibl.jpeg','mostafa-kamalpour':'Mostafa Kamalpour.jpeg',
 'hassan-ahmed':'hassan-ahmed.jpg','dana-dermody':'Gordana (Dana) Dermody.jpeg','abdullah-khan':'abdullah-khan.jpg',
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
GROUP_OF = {}
for i in ['shahrzad-saremi','rania-shibl','dana-dermody']: GROUP_OF[i] = 'leadership'
for i in ['mostafa-kamalpour','hassan-ahmed']: GROUP_OF[i] = 'leads'
for i in ['svetlana-kolos','mohamadali-rezaeimanesh','amir-h-malekijoo','jie-zhu','thiwanka-kaushalya-nagasanga',
          'meerab-fatima','malahat-mardani','mounes-mardani']: GROUP_OF[i] = 'students'
GROUPS = [('leadership','Leadership & Founders'),('leads','Senior Researchers'),
          ('researchers','Researchers & Academics'),('students',"Master's & Undergraduate Students")]
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
]

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
        out.append(dict(id=mid, prefix='', name=name, full=name, suffix='', alias='', group='researchers', role='Researcher',
            title=title, inst=inst, country=country, areas=AREAS_OF.get(mid, []), interests=interests, orcid=orcid,
            scholar=scholar, linkedin=linkedin, email=''))
    return out

def process_photos(members):
    d = os.path.join(OUT, 'images', 'team'); os.makedirs(d, exist_ok=True)
    src = os.path.join(ROOT, 'images', 'members')
    for m in members:
        f = PHOTOS.get(m['id'])
        if not f or not os.path.exists(os.path.join(src, f)):
            m['photo'] = ''; continue
        im = ImageOps.exif_transpose(Image.open(os.path.join(src, f))).convert('RGB')
        w, h = im.size; s = min(w, h)
        left = (w - s) // 2
        top = 0 if h > w else 0            # faces sit in the upper part of portrait shots
        if h > w: top = int((h - s) * 0.12)
        im = im.crop((left, top, left + s, top + s)).resize((360, 360), Image.LANCZOS)
        im.save(os.path.join(d, m['id'] + '.jpg'), quality=82, optimize=True, progressive=True)
        m['photo'] = m['id'] + '.jpg'

# ───────────────────────── publications ─────────────────────────
# (id, type, year, authors, title, venue, doi, areas, note)
PUBS = [
 ('egenai-dbr','journal',2026,'Saremi, S., Mirzaei, M., Rasti, A., Nooraei Abadeh, M., Varposhti, M., Shibl, R., Ahmed, H., Aldakheel, S. K. A., Mealy, E., Wang, K., Humphreys, D.','EGenAI-DBR: A design-based framework for responsible generative AI integration in higher education','Education Innovations: Systems and Future Learning, 1(2)','',['edu'],''),
 ('coi-online-peer','journal',2026,'Dokhanchi, M., Saremi, S., Shibl, R., Heidari, M., Ahmed, H., Mansoor, D., … Mirzaei, M.','An extended community of inquiry framework for monitoring and predicting online peer learning participation','Journal of Applied Research in Higher Education, 18(8), 113–141','',['edu'],'Q2'),
 ('selm-ctr','journal',2026,'Ali, Z., Ahmed, H., Khan, A., Saremi, S., Shibl, R., Mirzaei, M., Rastegari, P., Wang, M.','SELM-CTR: A stacking ensemble deep learning model with SHAP-based analysis for large-scale click-through rate prediction','International Journal of Intelligent Computing and Cybernetics, 19(3)','',['connect'],'Q2'),
 ('latent-squeeze','journal',2026,'Mirzaei, M., Saremi, S., Khan, A., Ahmed, H., Rastegari, P., Nooraei Abadeh, M., Shibl, R., …','Latent-space feature squeezing for adversarially robust intrusion detection on IoT edge devices','International Journal of Intelligent Computing and Cybernetics, 19(4)','',['connect'],''),
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
    areas = {a['id']: dict(short=a['short'], color=a['color']) for a in AREAS}
    return 'window.SYNTERA_PUBS = %s;\nwindow.SYNTERA_AREAS = %s;\n' % (json.dumps(items, ensure_ascii=False, indent=1), json.dumps(areas))

# ───────────────────────── page chrome ─────────────────────────
NAV = [('research.html','Research','research'),('publications.html','Publications','publications'),
       ('people.html','People','people'),('about.html','About','about'),('contact.html','Contact','contact')]
LOGO = ('<svg viewBox="0 0 40 40" aria-hidden="true"><defs><linearGradient id="lg" x1="0" y1="0" x2="1" y2="1">'
        '<stop offset="0" stop-color="#E83E8C"/><stop offset="1" stop-color="#3A7BFF"/></linearGradient></defs>'
        '<rect width="40" height="40" rx="10" fill="#0F1A47"/>'
        '<path d="M28 12c-2-3-14-3-14 3 0 7 14 3 14 10 0 6-12 6-15 2" fill="none" stroke="url(#lg)" stroke-width="3" stroke-linecap="round"/>'
        '<circle cx="28" cy="12" r="3" fill="#FF8FC3"/><circle cx="13" cy="27" r="3" fill="#8FB4FF"/><circle cx="20" cy="20" r="2" fill="#fff"/></svg>')

import hashlib
VER = ''   # set in main() from asset contents

def page(fname, title, desc, body, active='', depth=0, extra_js='', home=False, pubs_on=False):
    p = '../' * depth
    cur = ' aria-current="page"'
    nav = ''.join(f'<li><a href="{p}{h}"{cur if k == active else ""}>{t}</a></li>' for h, t, k in NAV)
    full_title = BRAND['full'] if home else f'{title} · {BRAND["name"]}'
    areas_f = ''.join(f'<li><a href="{p}research/{a["id"]}.html">{a["short"]}</a></li>' for a in AREAS)
    data_js = f'<script src="{p}data/publications.js?v={VER}"></script><script src="{p}js/pubs.js?v={VER}"></script>' if pubs_on else ''
    doc = f'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{E(full_title)}</title>
<meta name="description" content="{E(desc)}">
<meta name="theme-color" content="#0F1A47">
<meta property="og:title" content="{E(full_title)}"><meta property="og:description" content="{E(desc)}"><meta property="og:type" content="website">
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
      <p>{BRAND['tagline']}. An applied AI and connected systems research group.</p></div>
    <div><h4>Research</h4><ul>{areas_f}</ul></div>
    <div><h4>The group</h4><ul><li><a href="{p}about.html">About</a></li><li><a href="{p}people.html">People</a></li><li><a href="{p}publications.html">Publications</a></li><li><a href="{p}join.html">Join us</a></li><li><a href="{p}privacy.html">Privacy</a></li></ul></div>
    <div><h4>Get in touch</h4><ul><li><a href="mailto:{BRAND['email']}">{BRAND['email']}</a></li><li><a href="{BRAND['scholar']}" rel="noopener" target="_blank">Director on Google Scholar</a></li></ul><a class="btn btn--pink btn--sm" href="{p}join.html">Join Us</a></div>
  </div>
  <div class="ftr__bar"><span>© 2026 {BRAND['name']}</span><span><a href="{p}privacy.html">Privacy</a></span></div>
</div></footer>
<script src="{p}js/site.js?v={VER}"></script>
{data_js}
{extra_js}
</body></html>'''
    path = os.path.join(OUT, fname); os.makedirs(os.path.dirname(path), exist_ok=True)
    open(path, 'w', encoding='utf-8').write(doc)

def flag(m, p=''):
    f = FLAGS.get(m['country'])
    return f'<img class="flag" src="{p}images/flags/{f}.png" alt="" width="20" height="14">' if f else ''

def avatar(m, p='', cls='av'):
    if m['photo']:
        return f'<img class="{cls}" src="{p}images/team/{m["photo"]}" alt="{E(m["name"])}" width="120" height="120" loading="lazy">'
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

def person_card(m, p=''):
    areas = ''.join(f'<span class="dot" style="--c:{AREA[a]["color"]}" title="{AREA[a]["short"]}"></span>' for a in m['areas'])
    inst = E(m['inst'][0]) if m['inst'] else ''
    return (f'<li class="pcard" data-areas="{" ".join(m["areas"]) or "methods"}" data-name="{E((m["name"]+" "+m["alias"]+" "+" ".join(m["interests"])).lower())}">'
            f'<a href="{p}people/{m["id"]}.html">{avatar(m, p)}'
            f'<h3>{E(disp_name(m))}</h3><p class="pcard__role">{E(m["role"])}</p>'
            f'<p class="pcard__inst">{flag(m, p)}<span>{inst}</span></p><div class="dots">{areas}</div></a></li>')

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
    stats = [(len(members), 'Members'), (len(countries), 'Countries'), (len(partners), 'Partner institutions'), (n_pubs, 'Featured publications')]
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
  <p class="lead">Applied AI and connected systems for healthier lives, sustainable food, personal learning and safer mobility.</p>
  <div class="cta"><a class="btn btn--pink" href="research.html">Explore Research</a><a class="btn btn--ghost" href="join.html">Join Us</a></div>
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
<section class="sec sec--blush"><div class="wrap"><p class="eyebrow">Collaborators</p><h2>Institutions we work with</h2>
  <p class="sub">Universities where our members are based.</p>
  <ul class="logos">{marquee}</ul></div></section>
<section class="sec"><div class="wrap"><div class="banner"><div><h2>Open to collaboration</h2>
  <p>Students, researchers, industry and institutions are welcome.</p></div>
  <div class="cta"><a class="btn btn--pink" href="join.html">Ways to join</a><a class="btn btn--ghost" href="contact.html">Contact us</a></div></div></div></section>'''
    page('index.html', BRAND['name'], 'SYNTERA Research Group is an applied AI and connected systems research group working on health, smart homes, agriculture, education, IoT and IoV.', body, home=True,
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
<div class="chips">{core_html}</div></div></section>'''
    page('research.html', 'Research', 'Six research areas: AI for health, smart health home, agriculture, education, Internet of Things and Internet of Vehicles.', body, 'research')

    # ---------- area pages
    for a in AREAS:
        team = [m for m in members if a['id'] in m['areas']]
        pubs = [p for p in PUBS if a['id'] in p[7]]
        sec_team = (f'<section class="sec sec--ice"><div class="wrap"><p class="eyebrow">Team</p><h2>Who works on this</h2><ul class="grid grid--people">{"".join(person_card(m, "../") for m in team)}</ul></div></section>' if team else '')
        sec_pubs = (f'<section class="sec"><div class="wrap"><p class="eyebrow">Publications</p><h2>Selected publications</h2><ol class="publist" data-pubs data-area="{a["id"]}" data-limit="6"></ol><p><a class="more" href="../publications.html?area={a["id"]}">All {a["short"]} papers →</a></p></div></section>' if pubs else '')
        emerging = (f'<p class="notice">{a["short"]} is an emerging area for the group. Interested? <a href="../join.html#collaborate">Get in touch</a>.</p>' if len(team) < 2 else '')
        body = f'''<section class="phead phead--area" style="--c:{a["color"]}"><div class="wrap"><p class="eyebrow"><a href="../research.html">Research</a> / {a["short"]}</p>
<div class="phead__row"><span class="area__ico area__ico--xl">{icon(a["icon"])}</span><div><h1>{a["name"]}</h1>
<ul class="tags tags--light">{"".join(f"<li>{t}</li>" for t in a["topics"])}</ul></div></div></div></section>
<section class="sec"><div class="wrap prose"><h2>Why it matters</h2><p>{E(a["why"])}</p>{emerging}
<h2>Key challenges</h2><ul>{"".join(f"<li>{E(c)}</li>" for c in a["challenges"])}</ul>
<h2>Our approach</h2><p>{E(a["approach"])}</p></div></section>{sec_pubs}{sec_team}'''
        page(f'research/{a["id"]}.html', a['name'], f'{a["name"]} at SYNTERA Research Group: why it matters, our approach, publications and team.', body, 'research', 1, pubs_on=True)

    # ---------- people
    chips = '<button class="fchip is-on" data-area="all">All</button>' + ''.join(
        f'<button class="fchip" data-area="{a["id"]}" style="--c:{a["color"]}">{a["short"].replace("SYNTERA ","")} <small>{area_cnt[a["id"]]}</small></button>' for a in AREAS) + \
        f'<button class="fchip" data-area="methods">Core AI &amp; Methods <small>{sum(not m["areas"] for m in members)}</small></button>'
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
        links = ''.join(f'<a class="btn btn--ghost-d btn--sm" href="{E(u)}" rel="noopener" target="_blank">{t}</a>' for t, u in lk)
        achips = ''.join(f'<a class="chip" style="--c:{AREA[a]["color"]}" href="../research/{a}.html">{AREA[a]["short"]}</a>' for a in m['areas']) or '<span class="chip chip--plain">Core AI &amp; Methods</span>'
        ints = ''.join(f'<li>{E(x)}</li>' for x in m['interests'])
        insts = '<br>'.join(E(x) for x in m['inst']) or '<span class="muted">Affiliation to be confirmed</span>'
        pubsec = f'<h2>Publications</h2><ol class="publist" data-pubs data-member="{m["id"]}"></ol>' if any(m['id'] == mid for p in PUBS for _, mid in parse_authors(p[3])) else ''
        prv = by_id[members[i-1]['id']] if i else members[-1]; nxt = members[(i+1) % len(members)]
        sp = sectioned_profile(m, insts, achips, ints, pubsec)
        main_html = f'''<section class="sec"><div class="wrap prose"><h2>Affiliation</h2><p>{insts}</p>
<h2>Research areas</h2><div class="chips">{achips}</div>
<h2>Research interests</h2><ul class="tags tags--dark">{ints}</ul>{pubsec}</div></section>'''
        pager = f'<div class="wrap"><nav class="pager"><a href="{prv["id"]}.html">← {E(prv["name"])}</a><a href="{nxt["id"]}.html">{E(nxt["name"])} →</a></nav></div>'
        body = f'''<section class="phead phead--profile"><div class="wrap profile"><div class="profile__ph">{avatar(m, "../", "av av--lg")}</div>
<div><p class="eyebrow"><a href="../people.html">People</a></p><h1>{E(disp_name(m))}{f"<small>{E(m['suffix'])}</small>" if m["suffix"] else ""}</h1>
<p class="lead">{E(m["role"])}{" · " + E(DETAILS.get(m["id"], {}).get("title", m["title"])) if m["title"] else ""}</p>
{"".join(f'<p class="lead lead--sub">{E(r)}</p>' for r in DETAILS.get(m["id"], {}).get("roles", []))}<p class="where">{flag(m, "../")} {E(m["country"])}</p><div class="cta">{links}</div></div></div></section>
{sp if sp else main_html}{pager}'''
        page(f'people/{m["id"]}.html', m['name'], f'{m["name"]}, {m["role"]} at SYNTERA Research Group. Research interests and links.', body, 'people', 1, pubs_on=True)

    # ---------- publications
    types = [('all','All')] + [(k, t) for k, t in [('journal','Journal'),('conference','Conference'),('book','Book'),('chapter','Chapter'),('thesis','Thesis'),('preprint','Preprint'),('other','Other')] if any(p[1] == k for p in PUBS)]
    years = sorted({p[2] for p in PUBS if p[2]}, reverse=True)
    tchips = ''.join(f'<button class="fchip{" is-on" if k=="all" else ""}" data-type="{k}">{t}</button>' for k, t in types)
    achips = '<button class="fchip is-on" data-area="all">All areas</button>' + ''.join(f'<button class="fchip" data-area="{a["id"]}" style="--c:{a["color"]}">{a["short"].replace("SYNTERA ","")}</button>' for a in AREAS)
    yopts = '<option value="all">All years</option>' + ''.join(f'<option>{y}</option>' for y in years)
    body = f'''<section class="phead"><div class="wrap"><p class="eyebrow">Publications</p><h1>Research output</h1>
<p class="lead">Papers, chapters, books and theses by group members and the director. Group members are shown in bold and link to their profiles.</p>
<p class="muted metrics">Director's Google Scholar profile (6 Oct 2026): 15,573 citations (12,216 since 2021) · h-index 21 (18 since 2021) · i10-index 22 (21 since 2021). <a href="{BRAND['scholar']}" rel="noopener" target="_blank">View profile</a></p></div></section>
<section class="sec"><div class="wrap"><div class="filters"><input id="q" type="search" placeholder="Search title, author, keyword or venue" aria-label="Search publications">
<select id="year" aria-label="Year">{yopts}</select></div>
<div class="fchips" id="tchips">{tchips}</div><div class="fchips" id="achips">{achips}</div>
<p id="count" class="muted" aria-live="polite"></p><ol class="publist" id="pubs" data-pubs></ol><p id="none" class="notice" hidden>No publications match those filters.</p></div></section>'''
    page('publications.html', 'Publications', 'Journal articles, conference papers and book chapters from SYNTERA Research Group, filterable by area, type and year.', body, 'publications', pubs_on=True)

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
<p>The group was founded by <a href="people/shahrzad-saremi.html">Dr. Shahrzad Saremi</a> and <a href="people/rania-shibl.html">Dr. Rania Shibl</a>, and brings together researchers, academics and students from {len(countries)} countries to build AI and connected systems (IoT and IoV) for real-world problems.</p>
<p class="muted">SYNTERA Research Group is not affiliated with any commercial company of a similar name.</p></div></section>
<section class="sec sec--ice" id="mission"><div class="wrap"><div class="grid grid--2"><div class="panel"><h2>Mission</h2>
<p>To design and apply AI and connected technologies (IoT, IoV) that solve real problems in health, agriculture, education and everyday living.</p></div>
<div class="panel"><h2>Vision</h2><p>A future where intelligent, connected systems make life healthier, food more sustainable, learning more personal and mobility safer.</p></div></div>
<h2 class="mt">Core values</h2><ul class="values">{vals}</ul></div></section>
<section class="sec" id="director"><div class="wrap"><div class="dir"><div>{avatar(director, "", "av av--lg")}</div>
<div class="prose"><p class="eyebrow">Founder &amp; Director</p><h2><a href="people/{director["id"]}.html">{E(disp_name(director))}</a></h2><p>{E(DETAILS["shahrzad-saremi"]["title"])}, {E(director["inst"][0])}.</p>
<p>{E(DETAILS["shahrzad-saremi"]["bio"][0])}</p>
<p>{E(DETAILS["shahrzad-saremi"]["bio"][1].split(". She is widely")[0])}.</p><p><a class="btn btn--blue btn--sm" href="people/{director["id"]}.html">Full profile</a></p></div></div></div></section>
<section class="sec sec--blush" id="partners"><div class="wrap"><p class="eyebrow">Collaborators</p><h2>Partner institutions</h2><ul class="logos">{plog}</ul></div></section>'''
    page('about.html', 'About', 'The story, mission, vision and values of SYNTERA Research Group.', body, 'about')

    # ---------- join
    body = f'''<section class="phead"><div class="wrap"><p class="eyebrow">Join us</p><h1>Open to collaboration</h1>
<p class="lead">PhD and Master's students, postdocs, research assistants, interns, industry and academic partners are all welcome.</p>
<a class="btn btn--pink" href="mailto:{BRAND['email']}?subject=Joining%20SYNTERA%20Research%20Group">Reach out</a></div></section>
<section class="sec" id="positions"><div class="wrap prose"><h2>Open positions</h2><p>No positions are open right now. Send us a short note anyway: we like hearing from motivated people.</p></div></section>
<section class="sec sec--ice" id="apply"><div class="wrap prose"><h2>How to apply</h2>
<p>Email us with the following attached:</p><ul><li><b>CV</b></li><li><b>Academic transcript</b></li><li><b>Research statement</b> (one page)</li></ul>
<p>In your message, tell us:</p><ol><li>Who you are</li><li>Which of our <a href="research.html">research areas</a> interests you</li><li>Where we can learn more about your work (Scholar, GitHub, ORCID)</li><li>What you can offer</li></ol>
<p><a class="btn btn--pink" href="mailto:{BRAND['email']}?subject=Joining%20SYNTERA%20Research%20Group">Email the group</a></p></div></section>
<section class="sec" id="collaborate"><div class="wrap prose"><h2>Collaborate with us</h2>
<p>We work with universities, hospitals, aged-care providers, farms, schools and industry. If you have a real problem where applied AI or connected systems could help, we would like to talk.</p>
<p><a class="btn btn--blue" href="mailto:{BRAND['email']}?subject=Collaboration%20with%20SYNTERA%20Research%20Group">Propose a collaboration</a></p></div></section>'''
    page('join.html', 'Join Us', 'Open PhD, Master\'s, postdoc, research assistant and internship opportunities, and how to apply.', body, 'join')

    # ---------- contact
    soc = f'<a class="btn btn--ghost-d btn--sm" href="{BRAND["scholar"]}" rel="noopener" target="_blank">Director on Google Scholar</a>'
    body = f'''<section class="phead"><div class="wrap"><p class="eyebrow">Contact</p><h1>Get in touch</h1></div></section>
<section class="sec"><div class="wrap"><div class="grid grid--3">
<div class="panel"><h3>Email</h3><p><a href="mailto:{BRAND['email']}">{BRAND['email']}</a></p><p class="muted">Contact for the group (Director, <a href="people/shahrzad-saremi.html">Dr. Shahrzad Saremi</a>).</p></div>
<div class="panel"><h3>Director's affiliation</h3><p><a href="{BRAND['host_url']}" rel="noopener" target="_blank">{BRAND['host']}</a></p><p class="muted">School of Science, Technology and Engineering, Queensland, Australia.</p></div>
<div class="panel"><h3>Elsewhere</h3><p>{soc}</p></div></div></div></section>'''
    page('contact.html', 'Contact', 'Contact SYNTERA Research Group: email and links.', body, 'contact')

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
.pcard a{display:block;text-align:center;text-decoration:none;color:var(--text);background:var(--card);border:1px solid var(--line);border-radius:16px;padding:18px 14px;height:100%;transition:transform .2s,box-shadow .2s}.pcard a:hover{transform:translateY(-4px);box-shadow:var(--sh)}
.pcard .av{width:104px;height:104px;margin:0 auto 12px}.pcard h3{font-size:1rem;margin-bottom:2px}.pcard__role{margin:0;font-size:.84rem;font-weight:600;color:var(--pink-ink)}
.pcard__inst{margin:.4rem 0 0;font-size:.76rem;color:var(--muted);display:flex;gap:6px;justify-content:center;align-items:flex-start;text-align:left;line-height:1.35}.flag{width:20px;height:14px;object-fit:cover;border-radius:2px;margin-top:2px;flex:none}
.pcard__inst span{display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
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
'''

JS_SITE = r'''(function(){
var d=document,r=d.documentElement;
var th=d.getElementById('theme');
if(th)th.addEventListener('click',function(){var dark=r.dataset.theme?r.dataset.theme==='dark':matchMedia('(prefers-color-scheme:dark)').matches;var n=dark?'light':'dark';r.dataset.theme=n;try{localStorage.setItem('syntera-theme',n)}catch(e){}});
var b=d.getElementById('burger'),nav=d.querySelector('.hdr nav');
if(b)b.addEventListener('click',function(){var o=nav.classList.toggle('open');b.setAttribute('aria-expanded',o)});
d.addEventListener('keydown',function(e){if(e.key==='Escape'&&nav){nav.classList.remove('open');b&&b.setAttribute('aria-expanded',false)}});
if('IntersectionObserver' in window){var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('in');io.unobserve(e.target)}})},{threshold:.1});d.querySelectorAll('.reveal').forEach(function(x){io.observe(x)})}else d.querySelectorAll('.reveal').forEach(function(x){x.classList.add('in')});
var reduce=matchMedia('(prefers-reduced-motion:reduce)').matches;
d.querySelectorAll('[data-count]').forEach(function(el){var n=+el.dataset.count;if(reduce)return;var t0=null;el.textContent='0';function f(t){t0=t0||t;var p=Math.min((t-t0)/900,1);el.textContent=Math.round(n*p);if(p<1)requestAnimationFrame(f)}requestAnimationFrame(f)});
var sn=d.querySelectorAll('.snav a');
if(sn.length&&'IntersectionObserver' in window){var map={};sn.forEach(function(a){map[a.getAttribute('href').slice(1)]=a});
var so=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){sn.forEach(function(a){a.classList.remove('on')});map[e.target.id].classList.add('on')}})},{rootMargin:'-140px 0px -65% 0px'});
d.querySelectorAll('.psec').forEach(function(x){so.observe(x)})}
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
 var au=p.authors.map(function(a){return a[1]?'<a class="au" href="'+B+'people/'+a[1]+'.html"><strong>'+esc(a[0])+'</strong></a>':esc(a[0])}).join('; ');
 var chips=p.areas.map(function(a){return '<a class="chip" style="--c:'+A[a].color+'" href="'+B+'research/'+a+'.html">'+A[a].short+'</a>'}).join('');
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

def write(path, txt):
    p = os.path.join(OUT, path); os.makedirs(os.path.dirname(p), exist_ok=True)
    open(p, 'w', encoding='utf-8').write(txt)

def main():
    if os.path.exists(OUT): shutil.rmtree(OUT)
    os.makedirs(OUT)
    members = build_members()
    process_photos(members)
    build_member_keys(members)
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
    n = sum(len(f) for _, _, f in os.walk(OUT))
    print(f'{len(members)} members, {len(PUBS)} publications, {n} files -> {OUT}')

if __name__ == '__main__':
    main()
