"""
Rapport Cleaner — Loading Systems
Nettoie automatiquement les rapports d'intervention PDF de techniciens.
"""

import os, re, sys, json, threading, tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pdfplumber
from PIL import Image as PILImage, ImageTk
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, Image as RLImage)
from reportlab.lib.units import mm

def resource_path(filename):
    """Retourne le chemin absolu vers une ressource (compatible PyInstaller)."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, filename)

# ── Thèmes UI ────────────────────────────────────────────────────────────────
THEMES = {
    'clair': {
        'C_BG':       '#f5f5f5',
        'C_PANEL':    '#ffffff',
        'C_CARD':     '#f0f0f0',
        'C_BORDER':   '#d0d0d0',
        'C_TEXT':     '#1a1a1a',
        'C_TEXT2':    '#666666',
        'C_ENTRY_BG': '#ffffff',
    },
    'sombre': {
        'C_BG':       '#1e1e1e',
        'C_PANEL':    '#252526',
        'C_CARD':     '#2d2d2d',
        'C_BORDER':   '#3e3e3e',
        'C_TEXT':     '#e8eaf0',
        'C_TEXT2':    '#aaaaaa',
        'C_ENTRY_BG': '#3c3c3c',
    },
}

# Couleurs fixes (indépendantes du thème)
C_ACCENT  = '#e10033'
C_ACCENT2 = '#b80029'
C_SUCCESS = '#2ea043'

# Couleurs actives (initialisées au chargement, mises à jour selon le thème)
C_BG = C_PANEL = C_CARD = C_BORDER = C_TEXT = C_TEXT2 = C_ENTRY_BG = ''

def apply_theme(theme_name):
    global C_BG, C_PANEL, C_CARD, C_BORDER, C_TEXT, C_TEXT2, C_ENTRY_BG
    t = THEMES.get(theme_name, THEMES['clair'])
    C_BG       = t['C_BG']
    C_PANEL    = t['C_PANEL']
    C_CARD     = t['C_CARD']
    C_BORDER   = t['C_BORDER']
    C_TEXT     = t['C_TEXT']
    C_TEXT2    = t['C_TEXT2']
    C_ENTRY_BG = t['C_ENTRY_BG']

apply_theme('clair')  # thème par défaut

# ── Config ────────────────────────────────────────────────────────────────────
CONFIG_PATH = os.path.join(os.path.expanduser('~'), '.rapport_cleaner_config.json')

DEFAULT_BLACKLIST = [
    'ras', 'x', 'graissage resserrage', 'graissage', 'gressage resserrage',
    'sécurité ok', 'securite ok', 'condamné', 'condamne',
    'occupé par camion en permanence', 'vétuste', 'vetuste',
    'bavette supérieur vétuste', 'bavette superieur vetuste',
]

DEFAULT_CORRECTIONS = {
    'poignet': 'poignée', 'chassepied': 'chasse-pied',
    'poignetlsf': 'poignée LSF', 'plusieursfois': 'plusieurs fois',
    'déforméqui': 'déformé qui', 'ferryqui': 'ferry qui',
    'bâchecôté': 'bâche côté', 'automanu': 'auto/manu',
    'biquette': 'béquille', 'devisencours': 'Devis en cours',
    'dequai': 'de quai', 'spotà': 'spot à',
    'choquer': 'choquée', 'choqué': 'choquée',
}

def load_config():
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
    except:
        cfg = {}
    if 'blacklist'    not in cfg: cfg['blacklist']    = list(DEFAULT_BLACKLIST)
    if 'corrections'  not in cfg: cfg['corrections']  = dict(DEFAULT_CORRECTIONS)
    if 'known_ok'     not in cfg: cfg['known_ok']     = []
    if 'theme'        not in cfg: cfg['theme']        = 'clair'
    return cfg

def save_config(cfg):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except:
        pass

# ── Traitement texte ──────────────────────────────────────────────────────────
NEVER_SUFFIX = {'bas','les','des','sur','par','pas','ou','et','en','un','une',
                'du','de','la','le','hs','ras','nord','sud','est','ral','plus',
                'avec','pour','dans'}

def fix_word_breaks(text):
    if not text: return text
    text = re.sub(r'\r\n|\r', '\n', text)
    lines = text.split('\n')
    result = []; i = 0
    while i < len(lines):
        line = lines[i]
        if i+1 < len(lines):
            nxt = lines[i+1].strip()
            if re.fullmatch(r'[a-zA-ZÀ-ÿ]{1,4}', nxt) and nxt.lower() not in NEVER_SUFFIX:
                m = re.search(r'[a-zA-ZÀ-ÿ]+$', line.rstrip())
                if m:
                    lw = m.group()
                    ends_c = bool(re.search(r'[^aeiouàâéèêëîïôùûüyAEIOUÀÂÉÈÊËÎÏÔÙÛÜY]$', lw))
                    if (ends_c or len(lw)<=11) and len(lw+nxt)>=5:
                        result.append(line.rstrip()[:m.start()]+lw+nxt); i+=2; continue
        result.append(line); i+=1
    return re.sub(r' {2,}', ' ', ' '.join(result)).strip()

def strip_choc(text):
    if not text: return text
    t = text.strip()
    if not re.search(r'\bchoc\b', t, re.IGNORECASE): return t
    has_hs = bool(re.search(r'\bHS\b', t, re.IGNORECASE))
    choc_re = re.compile(
        r'(?:léger\s+|leger\s+)?choc\s+'
        r'(?:(?:panneau\s+)?(?:bas\s+)?(?:et\s+)?(?:inter\s+)?)?'
        r'(?:sur\s+(?:hublot\s+)?)?(?:de\s+l\'extérieur\s+)?'
        r'(?:\d+x\d+(?:x\d+)?\s*)?(?:(?:nordsud|nord|sud|ral)\s*)?(?:\d{3,4}\s*)?'
        r'(?:poignée\s+[àa]\s+\w+\s*)?'
        r'(?:(?:pb|pi|ph|panneau\s+(?:bas|intermédiaire|intermediaire|haut))\s*)?'
        r'(?:\+\s*(?:pb|pi|ph|panneau\s+(?:bas|intermédiaire|intermediaire|haut))\s*)*'
        r'(?:\+?\s*hublot\s+\w+(?:\s+\w+)?\s+[\d×x]+(?:x\d+)?\s*)?', re.IGNORECASE)
    s = choc_re.sub('', t)
    for pat in [r'\b\d{3,4}x\d{3,4}(?:x\d+)?\b', r'\bx\d+\b(?!\s*cm)',
                r'\b(?:nordsud|nord|sud|ral\s*\d*)\b',
                r'\b(?:intérieur|extérieur|interieur|exterieur)\s+(?:\w+\s+)?\d{3,4}\b',
                r'\bet\s+(?:intérieur|extérieur)\s+\w+\b',
                r'\b\d+\s*x\b(?!\s*\d)',
                r'\b(?:extérieur|intérieur|exterieur|interieur)\s*(?:et\s*)?$',
                r'\b(?:inter|panneau|bas|léger|leger|droite|gauche)\b',
                r'\bpoutre\s+avant\b']:
        s = re.sub(pat, '', s, flags=re.IGNORECASE)
    s = re.sub(r'\bpanneau(bas|haut|intermédiaire|intermediaire)\b', r'Panneau \1', s, flags=re.IGNORECASE)
    s = re.sub(r'^[\s\+\-,;]+','',s); s = re.sub(r'[\s\+\-,;]+$','',s)
    s = re.sub(r' {2,}',' ',s).strip()
    if has_hs:
        garbled = bool(re.match(r'^[a-zA-Z]\s*[,\.]',s)) or \
                  len(re.findall(r'\b[a-zA-ZÀ-ÿ]\b',s))>2 or (s and len(s)<8)
        if garbled or not s: return t
    return s

NOISE_WORDS = {'rien','a','à','signaler','ras','condamné','condamne','choc',
               'leger','léger','panneau','bas','extérieur','exterieur','inter',
               'et','de','le','la','les','du','sur','hublot','x'}

def is_blacklisted_full(text, blacklist):
    if not text or not text.strip(): return True
    t = text.strip()
    if t.upper() == 'X': return True
    tl = t.lower()
    for term in blacklist:
        if re.fullmatch(re.escape(term.strip()), tl, re.IGNORECASE): return True
        try:
            if re.fullmatch(term.strip(), tl, re.IGNORECASE): return True
        except: pass
    words = re.findall(r'[a-zA-ZÀ-ÿ]+', tl)
    return len([w for w in words if w not in NOISE_WORDS]) == 0

def strip_blacklisted_parts(text, blacklist):
    """Supprime les parties blacklistées dans une cellule mixte."""
    if not text: return text
    segments = re.split(r'\s*\+\s*', text.strip())
    kept = [s.strip() for s in segments
            if s.strip() and not is_blacklisted_full(s.strip(), blacklist)]
    return ' + '.join(kept).strip()

def apply_corrections(text, corrections):
    for wrong, right in corrections.items():
        text = re.sub(r'\b' + re.escape(wrong) + r'\b', right, text, flags=re.IGNORECASE)
    return text

def clean_cell(text, corrections=None, blacklist=None):
    if not text: return ''
    t = fix_word_breaks(text)
    if corrections: t = apply_corrections(t, corrections)
    t = strip_choc(t)
    t = re.sub(r'\s*/?\s*(?:fuite\s+)?remplacement\s+effectué\b', '', t, flags=re.IGNORECASE)
    t = re.sub(r'^[\s/\-,;]+','',t); t = re.sub(r'[\s/\-,;]+$','',t)
    t = re.sub(r' {2,}',' ',t).strip()
    bl = blacklist or []
    if is_blacklisted_full(t, bl): return ''
    t = strip_blacklisted_parts(t, bl)
    if is_blacklisted_full(t, bl): return ''
    return t

# ── Détection structure PDF ───────────────────────────────────────────────────
def detect_structure(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        tables = page.extract_tables()
        if not tables or not tables[0]: return None
        header = [fix_word_breaks(c).strip().lower() if c else '' for c in tables[0][0]]
    if any('nom' in h for h in header) and any('commentaire' in h for h in header):
        return {'style': 'nom_commentaire', 'headers': header}
    data_cols = {}; n_col = None
    for i, h in enumerate(header):
        if re.search(r'^#$|^n°?$|^n\s*°?\s*de\s*série|^numéros?$|^numero$|^num$', h): n_col = i
        elif re.search(r'\bn\b|numéro', h) and n_col is None: n_col = i
        elif re.search(r'porte', h): data_cols['porte'] = i
        elif re.search(r'niveleur|quai', h): data_cols['niv'] = i
        elif re.search(r'sas', h): data_cols['sas'] = i
        elif re.search(r'butoir|butoire', h): data_cols['but'] = i
        elif re.search(r'rideau', h): data_cols['rideau'] = i
        elif re.search(r'cale', h): data_cols['cale'] = i
        elif re.search(r'chandelle', h): data_cols['chandelle'] = i
    if n_col is None: n_col = 0
    return {'style': 'standard', 'headers': header, 'n_col': n_col, 'data_cols': data_cols}

def detect_unknown_words(pdf_path, corrections, blacklist):
    KNOWN = {
        'hs','ras','choc','panneau','bas','inter','intermédiaire','haut','niveleur',
        'porte','sas','butoire','rideau','cale','tendeur','long','court','extensible',
        'crochet','joint','bavette','hublot','câble','cable','verrou','butée','butee',
        'flexible','verin','hydraulique','vidange','soudure','fixation','roulette',
        'parachute','moteur','carte','relais','électronique','spot','led','luminaire',
        'graissage','resserrage','vétuste','condamné','occupé','camion','permanence',
        'plaque','nacelle','toucan','traverse','devis','cours','poignée','poignet',
        'chasse','pied','béquille','biquette','charnière','spirale','raccordement',
        'boîte','rampe','benne','locale','souple','sécurité','suspente','rail',
        'montant','barre','écartement','corde','tirage','orange','mètres','câbles',
        'traction','diamètre','paire','galets','support','emboîtement','crawford',
        'tablier','usure','avancé','lame','finale','asservissement','absence',
        'cellule','fuite','caisson','arrière','cuve','groupe','choquée','poutre',
        'tordu','profil','alu','horizontal','ferry','prévoir','passage','commercial',
        'arrêt','urgence','complet','contact','manque','mfz','lsf','pi','pb','ph',
        'auto','manu','béquille','charnière','supérieur','inférieur',
    }
    unknowns = set()
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables: continue
            for row in tables[0][1:]:
                for cell in row:
                    if not cell: continue
                    t = fix_word_breaks(cell).lower()
                    t = apply_corrections(t, corrections)
                    for word in re.findall(r'[a-zA-ZÀ-ÿ]{4,}', t):
                        if word not in KNOWN and not is_blacklisted_full(word, blacklist):
                            if len(word) <= 15:
                                unknowns.add(word)
    return unknowns

# ── Génération PDF ────────────────────────────────────────────────────────────
TXT = colors.HexColor('#222222')
HEADER_BG = colors.HexColor('#404040')

def make_cell(text, bold=False, size=8, color=None):
    if color is None: color = TXT
    return Paragraph(str(text), ParagraphStyle('c', fontSize=size,
        fontName='Helvetica-Bold' if bold else 'Helvetica',
        textColor=color, leading=size*1.3))

def make_img(name, img_dir):
    path = os.path.join(img_dir, f"{name}.jpg")
    if not os.path.exists(path): return ''
    try:
        with PILImage.open(path) as im: w, h = im.size
        mh, mw = 20*mm, 23*mm; r = w/h
        rw = min(mw, mh*r) if r>=1 else (min(mh, mw/r)*r)
        rh = rw/r if r>=1 else min(mh, mw/r)
        return RLImage(path, width=rw, height=rh)
    except: return ''

def extract_and_map_images(pdf_path, img_dir, n_col=1):
    os.makedirs(img_dir, exist_ok=True)
    img_map = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            imgs = page.images
            if not imgs: continue
            ph = page.height
            tset = page.find_tables()
            if not tset: continue
            tables = page.extract_tables()
            tbl = tset[0]
            for img in imgs:
                name = img['name']
                path = os.path.join(img_dir, f"{name}.jpg")
                if not os.path.exists(path):
                    try:
                        raw = img['stream'].get_rawdata()
                        if raw[:2] == b'\xff\xd8':
                            open(path,'wb').write(raw)
                        else:
                            data = img['stream'].get_data()
                            w, h = img['srcsize']
                            if len(data)==w*h*3: pil=PILImage.frombytes('RGB',(w,h),data)
                            elif len(data)==w*h*4: pil=PILImage.frombytes('RGBA',(w,h),data)
                            elif len(data)==w*h: pil=PILImage.frombytes('L',(w,h),data)
                            else: continue
                            pil.save(path, quality=85)
                    except: continue
            row_imgs = {}
            for img in imgs:
                mid = (img['y0']+img['y1'])/2
                best, bdist = None, float('inf')
                for ri, row in enumerate(tbl.rows):
                    ry0=ph-row.bbox[3]; ry1=ph-row.bbox[1]
                    if ry0<=mid<=ry1: best=ri; break
                    d=abs(mid-(ry0+ry1)/2)
                    if d<bdist: bdist=d; best=ri
                if best is not None and best<len(tables[0]):
                    n_val = tables[0][best][n_col]
                    if n_val: n_val = fix_word_breaks(n_val).strip()
                    if n_val and n_val not in ('N°','N','Numéros','#','N°\nde\nsérie',''):
                        row_imgs.setdefault(n_val,[]).append((img['x0'],img['name']))
            for nv, il in row_imgs.items():
                il.sort(key=lambda x:x[0])
                names = [i[1] for i in il]
                if nv not in img_map: img_map[nv]=names
                else: img_map[nv].extend(names)
    return img_map

def generate_pdf(pdf_path, output_path, structure, corrections, blacklist, log_fn=None):
    def log(msg):
        if log_fn: log_fn(msg)
    style = structure.get('style', 'standard')
    img_dir = os.path.join(tempfile.gettempdir(), 'rapport_cleaner_imgs')
    img_n_col = structure.get('n_col', 1)
    if style == 'nom_commentaire': img_n_col = 0
    elif img_n_col == 0: img_n_col = 1

    log("Extraction des images...")
    img_map = extract_and_map_images(pdf_path, img_dir, n_col=img_n_col)
    log(f"  → {sum(len(v) for v in img_map.values())} image(s) pour {len(img_map)} ligne(s)")

    log("Lecture du tableau...")
    if style == 'nom_commentaire':
        rows_data, quais = _read_nom_commentaire(pdf_path, corrections, blacklist)
    else:
        rows_data = _read_standard(pdf_path, structure, corrections, blacklist)
        quais = None

    log("Génération du PDF...")
    _build_pdf(output_path, rows_data, img_map, img_dir, structure, quais, log)
    log(f"✓ PDF généré : {output_path}")

def _read_standard(pdf_path, structure, corrections, blacklist):
    dc = structure.get('data_cols', {}); n_col = structure.get('n_col', 0)
    col_order = ['porte','niv','sas','but','rideau','cale','chandelle']
    rows_data = []; seen = set()
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables: continue
            for row in tables[0]:
                if not row or not row[n_col]: continue
                n = fix_word_breaks(row[n_col]).strip()
                if not n or n in ('N','N°','#','Numéros','∑') or n in seen: continue
                if re.search(r'[a-zA-Z]{3,}', n) and not re.search(r'\d', n):
                    if n.lower() in ('n','n°','numéros','image','photo','porte','niveleur','sas'): continue
                seen.add(n)
                fields = []
                for role in col_order:
                    idx = dc.get(role)
                    raw = row[idx] if idx is not None and idx<len(row) and row[idx] else ''
                    fields.append(clean_cell(raw, corrections, blacklist))
                rows_data.append((len(rows_data), n) + tuple(fields))
    return rows_data

def _read_nom_commentaire(pdf_path, corrections, blacklist):
    quais = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables: continue
            for row in tables[0]:
                if not row or not row[0] or 'série' in (row[0] or '').lower(): continue
                serie = row[0].strip()
                nom = fix_word_breaks(row[1]) if row[1] else ''
                com = clean_cell(fix_word_breaks(row[2]) if row[2] else '', corrections, blacklist)
                if re.search(r'porte sectionnelle', nom, re.IGNORECASE):
                    m = re.search(r'(?:abloy|assa)\s+(\d+)', nom, re.IGNORECASE) or re.search(r'(\d+)\s*$', nom)
                    if m: quais.setdefault(int(m.group(1)),{})['porte']=(serie,com)
                elif re.search(r'niveleur', nom, re.IGNORECASE):
                    m = re.search(r':\s*(\d+)', nom) or re.search(r'(\d+)\s*$', nom)
                    if m: quais.setdefault(int(m.group(1)),{})['niv']=(serie,com)
                elif re.search(r'\bsas\b', nom, re.IGNORECASE):
                    m = re.search(r'(?:abloy)\s+(\d+)', nom, re.IGNORECASE) or re.search(r'(\d+)\s*$', nom)
                    if m: quais.setdefault(int(m.group(1)),{})['sas']=(serie,com)
    rows_data = []
    for qn in sorted(quais.keys()):
        d = quais[qn]
        rows_data.append((qn, str(qn),
                          d.get('porte',('',''))[1],
                          d.get('niv',  ('',''))[1],
                          d.get('sas',  ('',''))[1], '', ''))
    return rows_data, quais

def _build_pdf(output_path, rows_data, img_map, img_dir, structure, quais, log):
    style = structure.get('style','standard')
    dc = structure.get('data_cols', {})
    col_order = ['porte','niv','sas','but','rideau','cale','chandelle']
    active_cols = [r for r in col_order if r in dc] if style=='standard' else ['porte','niv','sas']
    COL_LABELS = {'porte':'Porte sectionnelle','niv':'Niveleur / Quai','sas':'SAS',
                  'but':'Butoire','rideau':'Rideau','cale':'Cale','chandelle':'Chandelle'}
    N_PHOTOS=4; N_IMG_W=22*mm; N_COL_W=14*mm
    page_w = landscape(A4)[0]-20*mm
    data_col_w = (page_w - N_PHOTOS*N_IMG_W - N_COL_W) / max(len(active_cols),1)
    col_widths = [N_COL_W]+[data_col_w]*len(active_cols)+[N_IMG_W]*N_PHOTOS
    headers = ['N°']+[COL_LABELS.get(r,r) for r in active_cols]+[f'Photo {i+1}' for i in range(N_PHOTOS)]
    header_row = [make_cell(h,bold=True,size=8,color=colors.white) for h in headers]
    table_data = [header_row]
    style_cmds = [
        ('BACKGROUND',(0,0),(-1,0),HEADER_BG),('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('ALIGN',(0,0),(-1,0),'CENTER'),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,0),8),
        ('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#cccccc')),
        ('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),
        ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
        ('ROWHEIGHT',(0,0),(-1,0),10*mm),
    ]
    alt = 0
    for row in rows_data:
        n = row[1]; fields = list(row[2:])
        active_fields = fields[:len(active_cols)]
        if not any(f.strip() for f in active_fields): continue
        if style=='nom_commentaire' and quais:
            qn=int(n); d=quais.get(qn,{})
            imgs = (img_map.get(d.get('porte',('',''))[0],[]) +
                    img_map.get(d.get('niv',  ('',''))[0],[]) +
                    img_map.get(d.get('sas',  ('',''))[0],[]))
        else:
            imgs = img_map.get(n,[])
        img_cells = [make_img(imgs[i],img_dir) if i<len(imgs) else '' for i in range(N_PHOTOS)]
        data_cells = [make_cell(f,bold=bool(f),size=7.5) for f in active_fields]
        table_data.append([make_cell(n,bold=True,size=7.5)]+data_cells+img_cells)
        idx=len(table_data)-1
        style_cmds.append(('BACKGROUND',(0,idx),(-1,idx),colors.HexColor('#f2f2f2') if alt%2==0 else colors.white))
        style_cmds.append(('ROWHEIGHT',(0,idx),(-1,idx),22*mm if imgs else 7*mm))
        alt+=1

    # Summary
    summary_rows=[]
    for row in rows_data:
        n=row[1]; fields=list(row[2:])
        if style=='nom_commentaire':
            summary_rows.append((0,n,fields[0] if len(fields)>0 else '',
                                   fields[1] if len(fields)>1 else '',
                                   fields[2] if len(fields)>2 else '','',''))
        else:
            cm={r:i for i,r in enumerate(active_cols)}
            summary_rows.append((0,n,
                fields[cm['porte']] if 'porte' in cm else '',
                fields[cm['niv']]   if 'niv'   in cm else '',
                fields[cm['sas']]   if 'sas'   in cm else '','',''))

    fname = os.path.splitext(os.path.basename(output_path))[0]
    # Nettoyer le nom : supprimer uniquement les suffixes qu'on rajoute (clean, v1, v2, v3...)
    raw_title = fname.replace('_',' ').replace('-',' ')
    société = re.sub(r'\s*[\-_]?\s*(?:clean\s*(?:v\d+)?|v\d+)\s*$', '', raw_title, flags=re.IGNORECASE).strip()
    société = re.sub(r'\s{2,}', ' ', société).strip()
    story = _build_summary(summary_rows, société)
    main_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    main_table.setStyle(TableStyle(style_cmds))
    story.append(main_table)
    doc = SimpleDocTemplate(output_path, pagesize=landscape(A4),
        leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    doc.build(story)

def _build_summary(rows_data, title):
    ts=ParagraphStyle('ts',fontSize=13,fontName='Helvetica-Bold',spaceAfter=8,alignment=1)
    ns=ParagraphStyle('n',fontSize=8.5,fontName='Helvetica',textColor=colors.HexColor('#333333'),spaceAfter=4,leading=13)
    def eq(seg):
        m=re.match(r'^\s*(\d+)\s*(?:x\s*)?',seg.strip()); return int(m.group(1)) if m else 1
    cats={}; tcats={}; vns=[]
    def add(c,n,q=1): cats.setdefault(c,{}); cats[c][n]=cats[c].get(n,0)+q
    def addt(l,n,q=1): tcats.setdefault(l,{}); tcats[l][n]=tcats[l].get(n,0)+q
    for row in rows_data:
        n=row[1]
        for ci,f in enumerate(row[2:]):
            if not f: continue
            fl=f.lower(); is_sas=(ci==2)
            if 'vidange' in fl or 'hydraulique' in fl:
                if n not in vns: vns.append(n)
            elif is_sas and ('tendeur' in fl or 'crochet' in fl or
                    re.search(r'\b\d+\s*(?:courts?|longs?|extensibles?)\b',fl) or
                    re.search(r'\b\d+\s*[cls]\b',fl)):
                for seg in re.split(r'[+,]',fl):
                    seg=seg.strip(); q=eq(seg)
                    if re.search(r'\bcrochets?\b',seg) or re.search(r'\b\d+\s*s\b',seg):
                        add('Crochet S',n,q); continue
                    ht='tendeur' in seg; he='extensible' in seg
                    sc=bool(re.search(r'\b\d+\s*courts?\b|\b\d+\s*c\b',seg))
                    sl=bool(re.search(r'\b\d+\s*longs?\b|\b\d+\s*l\b',seg))
                    dc2=bool(re.search(r'\btendeurs?\s+courts?\b',seg))
                    dl=bool(re.search(r'\btendeurs?\s+longs?\b',seg))
                    de=bool(re.search(r'\btendeurs?\s+\w*extensibles?\b|\bextensible\b',seg))
                    if not(ht or he or sc or sl or dc2 or dl or de): continue
                    if he or de: typ='E'
                    elif dc2 or sc: typ='C'
                    elif dl or sl: typ='L'
                    else:
                        m2=re.search(r'tendeurs?\s+([a-zA-Z])',seg); typ=m2.group(1).upper() if m2 else '?'
                    addt({'E':'Tendeur E (extensible)','L':'Tendeur L (long)','C':'Tendeur C (court)'}.get(typ,f'Tendeur {typ}'),n,q)
            else:
                hp=bool(re.search(r'\bpanneau\b|\binterm[eé]diaire\b|\bpi\b|\bpb\b|\bph\b',fl))
                hj=bool(re.search(r'\bjoint\b',fl))
                if hp or hj:
                    for seg in re.split(r'[,+]|\bet\b',fl):
                        seg=seg.strip()
                        if not seg: continue
                        q=eq(seg)
                        if re.search(r'\bjoint\b',seg): add('Joint',n,q); continue
                        if re.search(r'\bpanneau\s+bas\b|\bpb\b',seg): add('Panneau bas',n,q)
                        if re.search(r'\bpanneau\s+haut\b|\bph\b',seg): add('Panneau haut',n,q)
                        if re.search(r'\bpanneau\s+(?:inter\w*|interm[eé]diaire)\b|\bpi\b|\binterm[eé]diaire\b|\binter\s*(?:hublot|hs|et|\b)',seg): add('Panneau intermédiaire',n,q)
                elif 'suspente' in fl: add('Suspente à refaire',n)
                elif 'flexible' in fl: add('Flexible HS',n,eq(fl))
                elif 'verrou' in fl: add('Verrou HS',n)
                elif 'roulette' in fl: add('Roulette manquante',n,eq(fl))
                elif 'câble' in fl or 'cable' in fl: add('Câble acier',n,eq(fl))
                elif 'butée' in fl or 'butee' in fl: add('Butée HS',n,eq(fl))
                elif 'bavette' in fl: add('Bavette HS',n)
                elif 'hublot' in fl: add('Hublot HS',n)
                elif 'parachute' in fl: add('Parachute HS',n)
                elif 'moteur' in fl: add('Moteur HS',n)
                elif 'relais' in fl or 'carte' in fl or 'électronique' in fl: add('Défaut électronique',n)
                elif 'spot' in fl or 'luminaire' in fl or 'led' in fl: add('Éclairage HS',n)
                elif 'poignée' in fl or 'poignet' in fl: add('Poignée HS',n)
                elif 'chasse' in fl: add('Chasse-pied HS',n)
                elif 'béquille' in fl or 'bequille' in fl: add('Béquille sécurité absente',n)
                elif 'charnière' in fl: add('Charnière HS',n)
                elif 'soudure' in fl: add('Soudure à refaire',n)
                elif 'traverse' in fl: add('Traverse déformée',n)
                elif 'devis' in fl: add('Devis en cours',n)
                elif 'cellule' in fl or 'asservissement' in fl: add('Absence cellule asservissement',n)
                elif 'tendeur' in fl: addt('Tendeur L (long)',n,eq(fl))
                else: add('Autre',n)
    def fmt(lbl,d):
        tot=sum(d.values()); parts=[f"{k} ({q})" if q>1 else k for k,q in d.items()]
        return f"<b>{lbl}</b> ({tot}) : {', '.join(parts)}"
    # Titre centré : "Nom société — Rapport d'intervention"
    titre_final = f"{title} — Rapport d'intervention"

    # Logo en haut à droite
    story = []
    logo_path = resource_path('1631305813263.jpg')
    if os.path.exists(logo_path):
        try:
            logo_h = 14*mm
            with PILImage.open(logo_path) as im:
                lw, lh = im.size
            logo_w = logo_h * lw / lh
            logo = RLImage(logo_path, width=logo_w, height=logo_h)
            # Tableau 1 ligne : titre centré + logo à droite
            header_table = Table(
                [[Paragraph(titre_final, ts), logo]],
                colWidths=[247*mm, logo_w + 4*mm]
            )
            header_table.setStyle(TableStyle([
                ('ALIGN',    (0,0),(0,0), 'CENTER'),
                ('ALIGN',    (1,0),(1,0), 'RIGHT'),
                ('VALIGN',   (0,0),(-1,-1), 'MIDDLE'),
                ('LEFTPADDING',  (0,0),(-1,-1), 0),
                ('RIGHTPADDING', (0,0),(-1,-1), 0),
                ('TOPPADDING',   (0,0),(-1,-1), 0),
                ('BOTTOMPADDING',(0,0),(-1,-1), 4),
            ]))
            story.append(header_table)
        except:
            story.append(Paragraph(titre_final, ts))
    else:
        story.append(Paragraph(titre_final, ts))
    if vns: story.append(Paragraph(f"<b>Vidange groupe hydraulique recommandée</b> ({len(vns)}) : {', '.join(vns)}",ns))
    for lbl in sorted(tcats.keys()): story.append(Paragraph(fmt(lbl,tcats[lbl]),ns))
    for c,d in cats.items(): story.append(Paragraph(fmt(c,d),ns))
    story.append(Spacer(1,6))
    story.append(Table([['']],colWidths=[257*mm],style=TableStyle([('LINEABOVE',(0,0),(-1,-1),0.5,colors.HexColor('#cccccc'))])))
    story.append(Spacer(1,6))
    return story

# ── Fenêtre Paramètres ────────────────────────────────────────────────────────
class SettingsWindow(tk.Toplevel):
    def __init__(self, parent, cfg, on_save):
        super().__init__(parent)
        self.title("Paramètres — Rapport Cleaner")
        self.configure(bg=C_BG)
        self.resizable(True, True)
        self.geometry("700x520")
        self.grab_set()
        self.cfg = cfg
        self.on_save = on_save
        self._build()
        self._center(parent)

    def _center(self, parent):
        self.update_idletasks()
        pw=parent.winfo_x(); py=parent.winfo_y()
        pw2=parent.winfo_width(); ph2=parent.winfo_height()
        w,h=self.winfo_width(),self.winfo_height()
        self.geometry(f"+{pw+(pw2-w)//2}+{py+(ph2-h)//2}")

    def _build(self):
        s=ttk.Style(); s.theme_use('default')
        s.configure('S.TNotebook',background=C_BG,borderwidth=0)
        s.configure('S.TNotebook.Tab',background=C_PANEL,foreground=C_TEXT2,padding=[14,6],font=('Helvetica',9))
        s.map('S.TNotebook.Tab',background=[('selected',C_CARD)],foreground=[('selected',C_TEXT)])
        nb=ttk.Notebook(self,style='S.TNotebook')
        nb.pack(fill='both',expand=True,padx=12,pady=12)

        # ── Onglet 0 — Options générales ─────────────────────────────────────
        tab0=tk.Frame(nb,bg=C_CARD); nb.add(tab0,text='  ⚙  Options générales  ')
        tk.Label(tab0,text="Apparence",font=('Helvetica',10,'bold'),bg=C_CARD,fg=C_TEXT).pack(anchor='w',padx=16,pady=(16,6))
        tk.Frame(tab0,bg=C_BORDER,height=1).pack(fill='x',padx=16,pady=(0,12))
        theme_f=tk.Frame(tab0,bg=C_CARD); theme_f.pack(fill='x',padx=16,pady=(0,8))
        tk.Label(theme_f,text="Mode d'affichage :",bg=C_CARD,fg=C_TEXT,font=('Helvetica',9),width=20,anchor='w').pack(side='left')
        self.theme_var=tk.StringVar(value=self.cfg.get('theme','clair'))
        for val,lbl in [('clair','☀  Mode clair'),('sombre','🌙  Mode sombre')]:
            tk.Radiobutton(theme_f,text=lbl,variable=self.theme_var,value=val,
                           bg=C_CARD,fg=C_TEXT,selectcolor=C_CARD,
                           activebackground=C_CARD,font=('Helvetica',9),padx=12).pack(side='left',padx=4)
        tk.Label(tab0,text="Le changement de thème sera appliqué au prochain démarrage de l'application.",
                 bg=C_CARD,fg=C_TEXT2,font=('Helvetica',8),wraplength=580).pack(anchor='w',padx=16,pady=(8,0))

        # ── Onglet 1 — Blacklist ──────────────────────────────────────────────
        tab1=tk.Frame(nb,bg=C_CARD); nb.add(tab1,text='  🚫  Blacklist  ')
        tk.Label(tab1,text="Termes à ignorer.\n• Cellule contenant uniquement ce terme → supprimée.\n• Cellule contenant ce terme parmi d'autres → seul ce terme est retiré.",
                 bg=C_CARD,fg=C_TEXT2,font=('Helvetica',8),justify='left',wraplength=640).pack(anchor='w',padx=12,pady=(10,4))
        lf1=tk.Frame(tab1,bg=C_CARD); lf1.pack(fill='both',expand=True,padx=12,pady=(0,4))
        self.bl_list=tk.Listbox(lf1,font=('Courier',9),bg=C_ENTRY_BG,fg=C_TEXT,
            selectbackground=C_ACCENT,relief='flat',borderwidth=0,activestyle='none')
        sb1=tk.Scrollbar(lf1,command=self.bl_list.yview,bg=C_PANEL)
        self.bl_list.config(yscrollcommand=sb1.set)
        sb1.pack(side='right',fill='y'); self.bl_list.pack(side='left',fill='both',expand=True)
        self._refresh_bl()
        add_f1=tk.Frame(tab1,bg=C_CARD); add_f1.pack(fill='x',padx=12,pady=(0,8))
        tk.Label(add_f1,text="Nouveau terme :",bg=C_CARD,fg=C_TEXT2,font=('Helvetica',8)).pack(side='left')
        self.bl_entry=tk.Entry(add_f1,width=28,bg=C_ENTRY_BG,fg=C_TEXT,
            insertbackground=C_TEXT,relief='flat',font=('Helvetica',9))
        self.bl_entry.pack(side='left',padx=(6,8))
        self.bl_entry.bind('<Return>',lambda e: self._add_bl())
        tk.Button(add_f1,text="+ Ajouter",command=self._add_bl,bg=C_ACCENT,fg='white',
                  relief='flat',padx=10,pady=4,cursor='hand2',font=('Helvetica',8,'bold')).pack(side='left',padx=(0,8))
        tk.Button(add_f1,text="Supprimer la sélection",command=self._del_bl,
                  bg=C_PANEL,fg=C_TEXT2,relief='flat',padx=10,pady=4,cursor='hand2').pack(side='left',padx=(0,8))
        tk.Button(add_f1,text="Réinitialiser",command=self._reset_bl,
                  bg=C_PANEL,fg=C_TEXT2,relief='flat',padx=10,pady=4,cursor='hand2').pack(side='left')

        # ── Onglet 2 — Corrections ────────────────────────────────────────────
        tab2=tk.Frame(nb,bg=C_CARD); nb.add(tab2,text='  ✏️  Corrections  ')
        tk.Label(tab2,text='Corrections automatiques de mots (mot erroné → correction)',
                 bg=C_CARD,fg=C_TEXT2,font=('Helvetica',8)).pack(anchor='w',padx=12,pady=(10,4))
        lf2=tk.Frame(tab2,bg=C_CARD); lf2.pack(fill='both',expand=True,padx=12,pady=(0,4))
        self.corr_list=tk.Listbox(lf2,font=('Courier',9),bg=C_ENTRY_BG,fg=C_TEXT,
            selectbackground=C_ACCENT,relief='flat',borderwidth=0,activestyle='none')
        sb2=tk.Scrollbar(lf2,command=self.corr_list.yview,bg=C_PANEL)
        self.corr_list.config(yscrollcommand=sb2.set)
        sb2.pack(side='right',fill='y'); self.corr_list.pack(side='left',fill='both',expand=True)
        self._refresh_corr()
        af=tk.Frame(tab2,bg=C_CARD); af.pack(fill='x',padx=12,pady=(0,8))
        tk.Label(af,text="Erroné :",bg=C_CARD,fg=C_TEXT2,font=('Helvetica',8)).pack(side='left')
        self.cw=tk.Entry(af,width=16,bg=C_ENTRY_BG,fg=C_TEXT,insertbackground=C_TEXT,relief='flat',font=('Helvetica',9))
        self.cw.pack(side='left',padx=(4,8))
        tk.Label(af,text="→",bg=C_CARD,fg=C_TEXT2,font=('Helvetica',10)).pack(side='left')
        self.cr=tk.Entry(af,width=16,bg=C_ENTRY_BG,fg=C_TEXT,insertbackground=C_TEXT,relief='flat',font=('Helvetica',9))
        self.cr.pack(side='left',padx=(4,8))
        tk.Button(af,text="+ Ajouter",command=self._add_corr,bg=C_ACCENT,fg='white',
                  relief='flat',padx=8,pady=3,cursor='hand2',font=('Helvetica',8,'bold')).pack(side='left',padx=4)
        tk.Button(af,text="Supprimer",command=self._del_corr,bg=C_PANEL,fg=C_TEXT2,
                  relief='flat',padx=8,pady=3,cursor='hand2').pack(side='left')

        # ── Onglet 3 — Mots acceptés ──────────────────────────────────────────
        tab3=tk.Frame(nb,bg=C_CARD); nb.add(tab3,text='  ✅  Mots acceptés  ')
        tk.Label(tab3,text="Mots que l'outil ne signalera plus comme inconnus.",
                 bg=C_CARD,fg=C_TEXT2,font=('Helvetica',8)).pack(anchor='w',padx=12,pady=(10,4))
        of=tk.Frame(tab3,bg=C_CARD); of.pack(fill='both',expand=True,padx=12,pady=(0,4))
        self.ok_list=tk.Listbox(of,font=('Courier',9),bg=C_ENTRY_BG,fg=C_TEXT,
            selectbackground=C_ACCENT,relief='flat',borderwidth=0,activestyle='none')
        sb3=tk.Scrollbar(of,command=self.ok_list.yview,bg=C_PANEL)
        self.ok_list.config(yscrollcommand=sb3.set)
        sb3.pack(side='right',fill='y'); self.ok_list.pack(side='left',fill='both',expand=True)
        for w in sorted(self.cfg.get('known_ok',[])): self.ok_list.insert('end',f'  {w}')
        add_f3=tk.Frame(tab3,bg=C_CARD); add_f3.pack(fill='x',padx=12,pady=(0,8))
        tk.Label(add_f3,text="Nouveau mot :",bg=C_CARD,fg=C_TEXT2,font=('Helvetica',8)).pack(side='left')
        self.ok_entry=tk.Entry(add_f3,width=20,bg=C_ENTRY_BG,fg=C_TEXT,
            insertbackground=C_TEXT,relief='flat',font=('Helvetica',9))
        self.ok_entry.pack(side='left',padx=(6,8))
        self.ok_entry.bind('<Return>',lambda e: self._add_ok())
        tk.Button(add_f3,text="+ Ajouter",command=self._add_ok,bg=C_ACCENT,fg='white',
                  relief='flat',padx=10,pady=4,cursor='hand2',font=('Helvetica',8,'bold')).pack(side='left',padx=(0,8))
        tk.Button(add_f3,text="Supprimer la sélection",command=self._del_ok,
                  bg=C_PANEL,fg=C_TEXT2,relief='flat',padx=10,pady=4,cursor='hand2').pack(side='left',padx=(0,8))
        tk.Button(add_f3,text="Tout effacer",command=self._clear_ok,
                  bg=C_PANEL,fg=C_TEXT2,relief='flat',padx=10,pady=4,cursor='hand2').pack(side='left')

        # ── Boutons bas ───────────────────────────────────────────────────────
        bot=tk.Frame(self,bg=C_BG); bot.pack(fill='x',padx=12,pady=(0,12))
        tk.Button(bot,text="✓ Enregistrer",command=self._save,bg=C_SUCCESS,fg='white',
                  font=('Helvetica',10,'bold'),padx=20,pady=6,relief='flat',cursor='hand2').pack(side='right')
        tk.Button(bot,text="Annuler",command=self.destroy,bg=C_PANEL,fg=C_TEXT2,
                  font=('Helvetica',10),padx=14,pady=6,relief='flat',cursor='hand2').pack(side='right',padx=8)

    def _refresh_bl(self):
        self.bl_list.delete(0,'end')
        for term in self.cfg.get('blacklist',[]):
            self.bl_list.insert('end',f'  {term}')

    def _add_bl(self):
        term=self.bl_entry.get().strip()
        if term:
            bl=self.cfg.setdefault('blacklist',[])
            if term not in bl:
                bl.append(term)
                self._refresh_bl()
            self.bl_entry.delete(0,'end')

    def _del_bl(self):
        sel=self.bl_list.curselection()
        if sel:
            term=self.bl_list.get(sel[0]).strip()
            bl=self.cfg.get('blacklist',[])
            if term in bl: bl.remove(term)
            self._refresh_bl()

    def _reset_bl(self):
        self.cfg['blacklist']=list(DEFAULT_BLACKLIST)
        self._refresh_bl()

    def _add_ok(self):
        word=self.ok_entry.get().strip()
        if word:
            ok=self.cfg.setdefault('known_ok',[])
            if word not in ok:
                ok.append(word)
                self.ok_list.insert('end',f'  {word}')
            self.ok_entry.delete(0,'end')

    def _refresh_corr(self):
        self.corr_list.delete(0,'end')
        for w,r in sorted(self.cfg.get('corrections',{}).items()):
            self.corr_list.insert('end',f'  {w}  →  {r}')

    def _add_corr(self):
        w=self.cw.get().strip(); r=self.cr.get().strip()
        if w and r:
            self.cfg.setdefault('corrections',{})[w]=r
            self._refresh_corr()
            self.cw.delete(0,'end'); self.cr.delete(0,'end')

    def _del_corr(self):
        sel=self.corr_list.curselection()
        if sel:
            line=self.corr_list.get(sel[0]).strip()
            wrong=line.split('→')[0].strip()
            self.cfg.get('corrections',{}).pop(wrong,None)
            self._refresh_corr()

    def _del_ok(self):
        sel=self.ok_list.curselection()
        if sel:
            word=self.ok_list.get(sel[0]).strip()
            ok_l=self.cfg.get('known_ok',[])
            if word in ok_l: ok_l.remove(word)
            self.ok_list.delete(sel[0])

    def _clear_ok(self):
        self.cfg['known_ok']=[]; self.ok_list.delete(0,'end')

    def _save(self):
        self.cfg['theme']=self.theme_var.get()
        save_config(self.cfg)
        self.on_save()
        self.destroy()

# ── Application principale ────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Rapport Cleaner — Loading Systems")
        self.resizable(False, False)
        self.cfg = load_config()
        apply_theme(self.cfg.get('theme', 'clair'))
        self.configure(bg=C_BG)
        self.pdf_path = tk.StringVar()
        self.out_path = tk.StringVar()

        # Icône barre des tâches
        try:
            _ico_img = ImageTk.PhotoImage(file=resource_path('1631305813263.jpg'))
            self.iconphoto(True, _ico_img)
            self._ico_ref = _ico_img
        except Exception as e:
            print(f"Icône barre des tâches : {e}")

        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w,h=self.winfo_width(),self.winfo_height()
        sw,sh=self.winfo_screenwidth(),self.winfo_screenheight()
        self.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

    def _build_ui(self):
        # Barre accent en haut
        tk.Frame(self,bg=C_ACCENT,height=4).pack(fill='x')

        # Titre + logo + bouton paramètres
        title_f=tk.Frame(self,bg=C_BG); title_f.pack(fill='x',padx=20,pady=(14,8))

        # Logo Loading Systems (page d'accueil)
        try:
            _logo_pil = PILImage.open(resource_path('logoloadingsystemspng_5c2f7debbf555.png'))
            _lh = 38
            _lw = int(_logo_pil.width * _lh / _logo_pil.height)
            _logo_pil = _logo_pil.resize((_lw, _lh), PILImage.LANCZOS)
            _logo_tk = ImageTk.PhotoImage(_logo_pil)
            lbl_logo = tk.Label(title_f, image=_logo_tk, bg=C_BG)
            lbl_logo.image = _logo_tk
            lbl_logo.pack(side='left', padx=(0, 14))
        except Exception as e:
            print(f"Logo accueil : {e}")
            tk.Label(title_f,text="Rapport Cleaner",font=('Helvetica',18,'bold'),bg=C_BG,fg=C_TEXT).pack(side='left')
            tk.Label(title_f,text="  Loading Systems",font=('Helvetica',11),bg=C_BG,fg=C_TEXT2).pack(side='left',pady=(4,0))

        tk.Button(title_f,text="⚙  Paramètres",command=self._open_settings,
                  bg=C_PANEL,fg=C_TEXT2,relief='flat',padx=12,pady=5,
                  font=('Helvetica',9),cursor='hand2',
                  activebackground=C_CARD,activeforeground=C_TEXT).pack(side='right')

        tk.Frame(self,bg=C_BORDER,height=1).pack(fill='x',padx=20)

        # Fichier source
        self._section("Fichier source")
        sf=tk.Frame(self,bg=C_PANEL); sf.pack(fill='x',padx=20,pady=(0,12))
        si=tk.Frame(sf,bg=C_PANEL); si.pack(fill='x',padx=12,pady=10)
        self.lbl_src=tk.Label(si,text="Aucun fichier sélectionné",bg=C_ENTRY_BG,fg=C_TEXT2,
                              font=('Helvetica',9),anchor='w',padx=10,pady=8,width=55)
        self.lbl_src.pack(side='left',fill='x',expand=True)
        tk.Button(si,text="Parcourir...",command=self._pick_pdf,bg=C_ACCENT,fg='white',
                  relief='flat',padx=14,pady=8,font=('Helvetica',9,'bold'),cursor='hand2',
                  activebackground=C_ACCENT2,activeforeground='white').pack(side='left',padx=(8,0))

        # Fichier de sortie
        self._section("Fichier de sortie")
        of=tk.Frame(self,bg=C_PANEL); of.pack(fill='x',padx=20,pady=(0,12))
        oi=tk.Frame(of,bg=C_PANEL); oi.pack(fill='x',padx=12,pady=10)
        self.lbl_out=tk.Label(oi,text="Aucun emplacement choisi",bg=C_ENTRY_BG,fg=C_TEXT2,
                              font=('Helvetica',9),anchor='w',padx=10,pady=8,width=55)
        self.lbl_out.pack(side='left',fill='x',expand=True)
        tk.Button(oi,text="Choisir...",command=self._pick_out,bg=C_PANEL,fg=C_TEXT,
                  relief='flat',padx=14,pady=8,font=('Helvetica',9),cursor='hand2',
                  activebackground=C_CARD).pack(side='left',padx=(8,0))

        # Progress
        pf=tk.Frame(self,bg=C_BG); pf.pack(fill='x',padx=20,pady=(0,4))
        s=ttk.Style(); s.theme_use('default')
        s.configure('Blue.Horizontal.TProgressbar',troughcolor=C_ENTRY_BG,
                    background=C_ACCENT,borderwidth=0,lightcolor=C_ACCENT,darkcolor=C_ACCENT)
        self.progress=ttk.Progressbar(pf,mode='indeterminate',length=560,
                                      style='Blue.Horizontal.TProgressbar')
        self.progress.pack(fill='x')

        # Journal
        self._section("Journal")
        lf=tk.Frame(self,bg=C_PANEL); lf.pack(fill='both',expand=True,padx=20,pady=(0,12))
        self.log_box=scrolledtext.ScrolledText(lf,height=8,width=70,font=('Courier',8),
            state='disabled',bg=C_ENTRY_BG,fg='#a8c0a0',insertbackground=C_TEXT,
            relief='flat',borderwidth=0,padx=10,pady=8)
        self.log_box.pack(fill='both',expand=True,padx=1,pady=1)

        # Boutons bas
        bf=tk.Frame(self,bg=C_BG); bf.pack(fill='x',padx=20,pady=(0,18))
        self.btn_run=tk.Button(bf,text="▶   Générer le PDF propre",
            font=('Helvetica',11,'bold'),bg=C_ACCENT,fg='white',
            padx=24,pady=10,relief='flat',cursor='hand2',
            activebackground=C_ACCENT2,activeforeground='white',
            state='disabled',command=self._run)
        self.btn_run.pack(side='left')
        tk.Button(bf,text="Effacer le journal",command=self._clear_log,
                  bg=C_PANEL,fg=C_TEXT2,relief='flat',padx=12,pady=10,
                  font=('Helvetica',9),cursor='hand2').pack(side='left',padx=10)

    def _section(self, label):
        f=tk.Frame(self,bg=C_BG); f.pack(fill='x',padx=20,pady=(8,4))
        tk.Label(f,text=label.upper(),font=('Helvetica',7,'bold'),bg=C_BG,fg=C_TEXT2).pack(side='left')
        tk.Frame(f,bg=C_BORDER,height=1).pack(side='left',fill='x',expand=True,padx=(8,0),pady=4)

    def _pick_pdf(self):
        path=filedialog.askopenfilename(title="Choisir le rapport PDF",
            filetypes=[("Fichiers PDF","*.pdf"),("Tous","*.*")])
        if path:
            self.pdf_path.set(path)
            name=os.path.basename(path)
            self.lbl_src.config(text=f"  {name}",fg=C_TEXT)
            out=os.path.splitext(path)[0]+'_clean.pdf'
            self.out_path.set(out)
            self.lbl_out.config(text=f"  {os.path.basename(out)}",fg=C_TEXT)
            self._check_ready()
            self._log(f"📂 {name}")

    def _pick_out(self):
        path=filedialog.asksaveasfilename(title="Enregistrer le PDF nettoyé",
            defaultextension=".pdf",filetypes=[("Fichiers PDF","*.pdf")])
        if path:
            self.out_path.set(path)
            self.lbl_out.config(text=f"  {os.path.basename(path)}",fg=C_TEXT)
            self._check_ready()

    def _check_ready(self):
        if self.pdf_path.get() and self.out_path.get():
            self.btn_run.config(state='normal')

    def _log(self, msg):
        self.log_box.config(state='normal')
        self.log_box.insert('end', msg+'\n')
        self.log_box.see('end')
        self.log_box.config(state='disabled')

    def _clear_log(self):
        self.log_box.config(state='normal')
        self.log_box.delete('1.0','end')
        self.log_box.config(state='disabled')

    def _open_settings(self):
        SettingsWindow(self, self.cfg,
                       on_save=lambda: self._log("✓ Paramètres sauvegardés"))

    def _run(self):
        pdf=self.pdf_path.get(); out=self.out_path.get()
        if not pdf or not out: return
        self.btn_run.config(state='disabled')
        self.progress.start(10)
        self._log("\n─── Démarrage ───")
        def worker():
            try:
                self._log("Analyse de la structure...")
                structure=detect_structure(pdf)
                if not structure:
                    self.after(0,lambda: self._log("❌ Impossible de lire le PDF.")); return
                self._log(f"Structure : {structure.get('style')} — {list(structure.get('data_cols',{}).keys()) or 'nom/commentaire'}")
                self._log("Analyse du vocabulaire...")
                unknowns=detect_unknown_words(pdf,self.cfg.get('corrections',{}),self.cfg.get('blacklist',[]))
                known_already=set(self.cfg.get('corrections',{}).keys())|set(self.cfg.get('known_ok',[]))
                new_unknowns=unknowns-known_already
                if new_unknowns:
                    self.after(0,lambda u=new_unknowns: self._ask_unknowns(u,pdf,out,structure)); return
                self._do_generate(pdf,out,structure)
            except Exception as e:
                self.after(0,lambda: self._log(f"❌ Erreur : {e}"))
            finally:
                self.after(0,self._stop_progress)
        threading.Thread(target=worker,daemon=True).start()

    def _ask_unknowns(self, unknowns, pdf, out, structure):
        win=tk.Toplevel(self); win.title("Mots inhabituels détectés")
        win.configure(bg=C_BG); win.resizable(False,True); win.grab_set()
        tk.Label(win,text=f"  {len(unknowns)} mot(s) inhabituel(s) trouvé(s). Que faire ?",
                 font=('Helvetica',10,'bold'),bg=C_BG,fg=C_TEXT,pady=10).pack(anchor='w',padx=12)
        tk.Label(win,text="  ✓ Garder = conserver tel quel   |   ✏ Corriger = remplacer   |   ✗ Ignorer = supprimer",
                 font=('Helvetica',8),bg=C_BG,fg=C_TEXT2).pack(anchor='w',padx=12,pady=(0,8))
        frame=tk.Frame(win,bg=C_BG); frame.pack(fill='both',expand=True,padx=12)
        canvas=tk.Canvas(frame,bg=C_BG,highlightthickness=0,height=min(320,len(unknowns)*46+10))
        sb=tk.Scrollbar(frame,orient='vertical',command=canvas.yview,bg=C_PANEL)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right',fill='y'); canvas.pack(side='left',fill='both',expand=True)
        inner=tk.Frame(canvas,bg=C_BG); canvas.create_window((0,0),window=inner,anchor='nw')
        decisions={}
        for word in sorted(unknowns):
            rf=tk.Frame(inner,bg=C_CARD,pady=6); rf.pack(fill='x',pady=2,padx=4)
            tk.Label(rf,text=f"  «{word}»",font=('Courier',9,'bold'),bg=C_CARD,fg=C_TEXT,width=22,anchor='w').pack(side='left')
            action=tk.StringVar(value='ok'); correction=tk.StringVar(value=word)
            for val,lbl in [('ok','✓ Garder'),('correction','✏ Corriger :'),('blacklist','✗ Ignorer')]:
                tk.Radiobutton(rf,text=lbl,variable=action,value=val,bg=C_CARD,fg=C_TEXT,
                               selectcolor=C_CARD,activebackground=C_CARD,font=('Helvetica',8)).pack(side='left',padx=4)
            tk.Entry(rf,textvariable=correction,width=14,bg=C_ENTRY_BG,fg=C_TEXT,
                     insertbackground=C_TEXT,relief='flat',font=('Helvetica',9)).pack(side='left',padx=4)
            decisions[word]=(action,correction)
        inner.update_idletasks(); canvas.config(scrollregion=canvas.bbox('all'))
        def confirm():
            for word,(av,cv) in decisions.items():
                act=av.get()
                if act=='correction': self.cfg.setdefault('corrections',{})[word]=cv.get()
                elif act=='blacklist':
                    bl=self.cfg.setdefault('blacklist',[])
                    if word not in bl: bl.append(word)
                else:
                    ok=self.cfg.setdefault('known_ok',[])
                    if word not in ok: ok.append(word)
            save_config(self.cfg)
            self._log(f"✓ {len(decisions)} mot(s) traité(s)")
            win.destroy()
            self._do_generate(pdf,out,structure)
        tk.Button(win,text="✓ Confirmer et générer",command=confirm,bg=C_SUCCESS,fg='white',
                  font=('Helvetica',10,'bold'),padx=20,pady=8,relief='flat',cursor='hand2').pack(pady=12)

    def _do_generate(self, pdf, out, structure):
        def worker():
            try:
                generate_pdf(pdf,out,structure,
                    self.cfg.get('corrections',{}),
                    self.cfg.get('blacklist',[]),
                    log_fn=lambda m: self.after(0,lambda msg=m: self._log(msg)))
                self.after(0,lambda: messagebox.showinfo("Terminé ✓",f"PDF généré avec succès !\n\n{out}"))
            except Exception as e:
                self.after(0,lambda: self._log(f"❌ Erreur : {e}"))
                self.after(0,lambda: messagebox.showerror("Erreur",str(e)))
            finally:
                self.after(0,self._stop_progress)
                self.after(0,lambda: self.btn_run.config(state='normal'))
        threading.Thread(target=worker,daemon=True).start()

    def _stop_progress(self):
        self.progress.stop()

if __name__ == '__main__':
    app = App()
    app.mainloop()
