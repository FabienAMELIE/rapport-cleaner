"""
Rapport Cleaner — Loading Systems
Nettoie automatiquement les rapports d'intervention PDF de techniciens.
"""

import os, re, sys, json, threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pdfplumber
from PIL import Image as PILImage
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, Image as RLImage)
from reportlab.lib.units import mm

# ── Config persistence ────────────────────────────────────────────────────────
CONFIG_PATH = os.path.join(os.path.expanduser('~'), '.rapport_cleaner_config.json')

def load_config():
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {'blacklist_extra': [], 'corrections': {}, 'known_structures': {}}

def save_config(cfg):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except:
        pass

# ── Core processing logic ─────────────────────────────────────────────────────
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
                    if (ends_c or len(lw) <= 11) and len(lw+nxt) >= 5:
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

BASE_BLACKLIST = [
    r'^r[ae]s$',
    r'^x$',
    r'^g?r[ae]iss[ae]ge\s*r?e?ss?err[ae]ge?$',
    r'^g?r[ae]iss[ae]ge$',
    r'^gressage\s*resserrage$',
    r'^s[ée]curit[ée]\s*ok$',
    r'^condamné$',
    r'^occupé\s+par\s+camion\s+en\s+permanence$',
    r'^vétuste$', r'^vetuste$',
    r'^\s*$',
]

def is_blacklisted(text, extra_patterns=None):
    if not text or not text.strip(): return True
    t = text.strip()
    all_patterns = BASE_BLACKLIST + (extra_patterns or [])
    for pat in all_patterns:
        if re.fullmatch(pat, t, re.IGNORECASE): return True
    if t.upper() == 'X': return True
    # Check if all meaningful words are noise
    words = re.findall(r'[a-zA-ZÀ-ÿ]+', t.lower())
    return len([w for w in words if w not in NOISE_WORDS]) == 0

def apply_corrections(text, corrections):
    """Apply learned text corrections."""
    for wrong, right in corrections.items():
        text = re.sub(r'\b' + re.escape(wrong) + r'\b', right, text, flags=re.IGNORECASE)
    return text

def clean_cell(text, corrections=None, extra_blacklist=None):
    if not text: return ''
    t = fix_word_breaks(text)
    if corrections: t = apply_corrections(t, corrections)
    t = strip_choc(t)
    # Filter "remplacement effectué"
    t = re.sub(r'\s*/?\s*(?:fuite\s+)?remplacement\s+effectué\b', '', t, flags=re.IGNORECASE)
    t = re.sub(r'^[\s/\-,;]+','',t); t = re.sub(r'[\s/\-,;]+$','',t)
    t = re.sub(r' {2,}',' ',t).strip()
    return '' if is_blacklisted(t, extra_blacklist) else t

# ── PDF structure detection ───────────────────────────────────────────────────
def detect_structure(pdf_path):
    """
    Auto-detect PDF table structure.
    Returns dict with:
      - headers: list of column headers
      - data_cols: dict mapping role -> col_index
      - n_col: index of N/ID column
      - style: 'standard' (N + cols) or 'nom_commentaire' (N° série + Nom + Commentaire)
    """
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        tables = page.extract_tables()
        if not tables or not tables[0]: return None
        header = [fix_word_breaks(c).strip().lower() if c else '' for c in tables[0][0]]

    # Detect nom/commentaire style (KN Soissons)
    if any('nom' in h for h in header) and any('commentaire' in h for h in header):
        return {'style': 'nom_commentaire', 'headers': header}

    # Standard style — find key columns
    data_cols = {}
    n_col = None
    for i, h in enumerate(header):
        if re.search(r'^#$|^n°?$|^n\s*°?\s*de\s*série', h): n_col = i
        elif re.search(r'\bn\b|numéro', h) and n_col is None: n_col = i
        elif re.search(r'porte', h): data_cols['porte'] = i
        elif re.search(r'niveleur|quai', h): data_cols['niv'] = i
        elif re.search(r'sas', h): data_cols['sas'] = i
        elif re.search(r'butoir|butoire', h): data_cols['but'] = i
        elif re.search(r'rideau', h): data_cols['rideau'] = i
        elif re.search(r'cale', h): data_cols['cale'] = i
        elif re.search(r'chandelle', h): data_cols['chandelle'] = i

    if n_col is None: n_col = 0

    return {
        'style': 'standard',
        'headers': header,
        'n_col': n_col,
        'data_cols': data_cols,
    }

def detect_unknown_words(pdf_path, structure, corrections, extra_blacklist):
    """Scan all cells for words that look unusual/unknown — possible typos or new terms."""
    KNOWN_WORDS = {
        'hs','ras','choc','panneau','bas','inter','intermédiaire','haut','niveleur',
        'porte','sas','butoire','rideau','cale','tendeur','long','court','extensible',
        'crochet','joint','bavette','hublot','câble','cable','verrou','butée','butee',
        'flexible','verin','hydraulique','vidange','soudure','fixation','roulette',
        'parachute','moteur','carte','relais','électronique','spot','led','luminaire',
        'graissage','resserrage','vétuste','condamné','occupé','camion','permanence',
        'plaque','nacelle','toucan','traversee','traverse','devis','cours',
        'poignée','poignet','chasse','pied','béquille','biquette','charnière',
        'spirale','raccordement','boîte','rampe','benne','locale','souple','sécurité',
        'suspension','suspente','rail','montant','barre','écartement','corde','tirage',
        'orange','mètres','câbles','traction','diamètre','paire','galets','support',
        'roulette','emboîtement','crawford','poignée','tablier','usure','avancé',
        'lame','finale','nacelle','asservissement','absence','cellule','fuite',
        'bavette','caisson','arrière','cuve','groupe','choquer','choquée','poutre',
        'tordu','profil','alu','horizontal','fermer','ferry','niveleur','prévoir',
        'passage','commercial','arrêt','urgence','complet','contact','manque','ras',
        'rien','signaler','x','ok','mfz','lsf','hs','pi','pb','ph','auto','manu',
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
                        if word not in KNOWN_WORDS and not is_blacklisted(word, extra_blacklist):
                            # Only flag words that look like they could be typos
                            # (not too long, not obviously a real French word)
                            if len(word) <= 15:
                                unknowns.add(word)
    return unknowns

# ── PDF generation ────────────────────────────────────────────────────────────
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
        rw = min(mw, mh*r) if r >= 1 else (min(mh, mw/r)*r)
        rh = rw/r if r >= 1 else min(mh, mw/r)
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
                if best is not None and best < len(tables[0]):
                    n_val = tables[0][best][n_col]
                    if n_val:
                        n_val = fix_word_breaks(n_val).strip()
                    if n_val and n_val not in ('N°','N','Numéros','#','N°\nde\nsérie',''):
                        row_imgs.setdefault(n_val,[]).append((img['x0'],img['name']))
            for nv, il in row_imgs.items():
                il.sort(key=lambda x:x[0])
                names = [i[1] for i in il]
                if nv not in img_map: img_map[nv]=names
                else: img_map[nv].extend(names)
    return img_map

def generate_pdf(pdf_path, output_path, structure, corrections, extra_blacklist, log_fn=None):
    import tempfile
    def log(msg):
        if log_fn: log_fn(msg)

    style = structure.get('style', 'standard')
    # Use system temp dir to avoid permission issues on Windows
    img_dir = os.path.join(tempfile.gettempdir(), 'rapport_cleaner_imgs')
    n_col = structure.get('n_col', 1)

    log("Extraction des images...")
    # For image mapping, use the N column (quai number), not the # column
    img_n_col = structure.get('n_col', 1)
    if img_n_col == 0 and structure.get('style') == 'standard':
        img_n_col = 1
    # For nom_commentaire style, images are mapped by N° série (col 0)
    if structure.get('style') == 'nom_commentaire':
        img_n_col = 0
    img_map = extract_and_map_images(pdf_path, img_dir, n_col=img_n_col)
    log(f"  → {sum(len(v) for v in img_map.values())} image(s) extraite(s) pour {len(img_map)} ligne(s)")

    log("Lecture du tableau...")

    if style == 'nom_commentaire':
        rows_data, quais = _read_nom_commentaire(pdf_path, corrections, extra_blacklist)
    else:
        rows_data = _read_standard(pdf_path, structure, corrections, extra_blacklist)
        quais = None

    log("Génération du PDF...")
    _build_pdf(output_path, rows_data, img_map, img_dir, structure, quais, log)
    log(f"✓ PDF généré : {output_path}")
def _read_standard(pdf_path, structure, corrections, extra_blacklist):
    dc = structure.get('data_cols', {})
    n_col = structure.get('n_col', 0)
    # Determine col order: porte, niv, sas, but, rideau, cale, chandelle
    col_order = ['porte','niv','sas','but','rideau','cale','chandelle']

    rows_data = []
    seen = set()
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables: continue
            for row in tables[0]:
                if not row or not row[n_col]: continue
                n = fix_word_breaks(row[n_col]).strip()
                if not n or n in ('N','N°','#','Numéros','∑') or n in seen: continue
                if re.search(r'[a-zA-Z]{3,}', n) and not re.search(r'\d', n):
                    # Probably a header repeat or label row — skip
                    if n.lower() in ('n','n°','numéros','image','photo','porte','niveleur','sas'): continue
                seen.add(n)
                fields = []
                for role in col_order:
                    idx = dc.get(role)
                    raw = row[idx] if idx is not None and idx < len(row) and row[idx] else ''
                    fields.append(clean_cell(raw, corrections, extra_blacklist))
                rows_data.append((len(rows_data), n) + tuple(fields))
    return rows_data

def _read_nom_commentaire(pdf_path, corrections, extra_blacklist):
    """Handle KN Soissons style: N°série / Nom / Commentaire."""
    quais = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables: continue
            for row in tables[0]:
                if not row or not row[0] or 'série' in (row[0] or '').lower(): continue
                serie = row[0].strip()
                nom = fix_word_breaks(row[1]) if row[1] else ''
                com = clean_cell(fix_word_breaks(row[2]) if row[2] else '', corrections, extra_blacklist)
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
        cp = d.get('porte',('',''))[1]
        cn = d.get('niv',('',''))[1]
        cs = d.get('sas',('',''))[1]
        rows_data.append((qn, str(qn), cp, cn, cs, '', ''))
    return rows_data, quais

def _build_pdf(output_path, rows_data, img_map, img_dir, structure, quais, log):
    style = structure.get('style','standard')
    dc = structure.get('data_cols', {})
    col_order = ['porte','niv','sas','but','rideau','cale','chandelle']
    active_cols = [r for r in col_order if r in dc] if style=='standard' else ['porte','niv','sas']

    COL_LABELS = {
        'porte':'Porte sectionnelle','niv':'Niveleur / Quai',
        'sas':'SAS','but':'Butoire','rideau':'Rideau',
        'cale':'Cale','chandelle':'Chandelle'
    }
    N_PHOTOS = 4
    N_IMG_WIDTH = 22*mm
    N_COL_WIDTH = 14*mm

    # Distribute remaining width among data cols
    page_w = landscape(A4)[0] - 20*mm  # margins
    img_total = N_PHOTOS * N_IMG_WIDTH
    n_total = N_COL_WIDTH
    data_total = page_w - img_total - n_total
    data_col_w = data_total / max(len(active_cols), 1)
    col_widths = [N_COL_WIDTH] + [data_col_w]*len(active_cols) + [N_IMG_WIDTH]*N_PHOTOS

    headers = ['N°'] + [COL_LABELS.get(r,r) for r in active_cols] + [f'Photo {i+1}' for i in range(N_PHOTOS)]
    header_row = [make_cell(h, bold=True, size=8, color=colors.white) for h in headers]

    table_data = [header_row]
    style_cmds = [
        ('BACKGROUND',(0,0),(-1,0),HEADER_BG),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('ALIGN',(0,0),(-1,0),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,0),8),
        ('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#cccccc')),
        ('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),
        ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
        ('ROWHEIGHT',(0,0),(-1,0),10*mm),
    ]

    alt = 0
    for row in rows_data:
        n = row[1]
        fields = list(row[2:])
        # Only take fields for active cols
        active_fields = fields[:len(active_cols)]
        if not any(f.strip() for f in active_fields): continue

        # Get images
        if style == 'nom_commentaire' and quais:
            qn = int(n)
            d = quais.get(qn, {})
            ps = d.get('porte',('',''))[0]
            ns_s = d.get('niv',('',''))[0]
            ss_s = d.get('sas',('',''))[0]
            imgs = (img_map.get(ps,[]) + img_map.get(ns_s,[]) +
                    img_map.get(ss_s,[]))
        else:
            imgs = img_map.get(n, [])

        img_cells = [make_img(imgs[i], img_dir) if i<len(imgs) else '' for i in range(N_PHOTOS)]
        data_cells = [make_cell(f, bold=bool(f), size=7.5) for f in active_fields]
        table_data.append([make_cell(n, bold=True, size=7.5)] + data_cells + img_cells)

        idx = len(table_data)-1
        bg = colors.HexColor('#f2f2f2') if alt%2==0 else colors.white
        style_cmds.append(('BACKGROUND',(0,idx),(-1,idx),bg))
        style_cmds.append(('ROWHEIGHT',(0,idx),(-1,idx),22*mm if imgs else 7*mm))
        alt += 1

    # Build summary
    summary_rows = []
    for row in rows_data:
        n = row[1]; fields = list(row[2:])
        # Map to (row_num, n, sas_field, other_fields...)
        # For summary, col_idx 2 = sas position
        if style == 'nom_commentaire':
            # fields = [porte, niv, sas, ...]
            summary_rows.append((0, n, fields[0] if len(fields)>0 else '',
                                  fields[1] if len(fields)>1 else '',
                                  fields[2] if len(fields)>2 else '', '', ''))
        else:
            # Map active cols to standard positions
            col_map = {r:i for i,r in enumerate(active_cols)}
            p = fields[col_map['porte']] if 'porte' in col_map else ''
            niv = fields[col_map['niv']] if 'niv' in col_map else ''
            sas = fields[col_map['sas']] if 'sas' in col_map else ''
            summary_rows.append((0, n, p, niv, sas, '', ''))

    # Extract title from filename
    fname = os.path.splitext(os.path.basename(output_path))[0]
    title = fname.replace('_',' ').replace('-',' ')

    story = _build_summary(summary_rows, title)
    main_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    main_table.setStyle(TableStyle(style_cmds))
    story.append(main_table)

    doc = SimpleDocTemplate(output_path, pagesize=landscape(A4),
        leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    doc.build(story)

def _build_summary(rows_data, title):
    ts=ParagraphStyle('ts',fontSize=13,fontName='Helvetica-Bold',spaceAfter=2)
    ss=ParagraphStyle('ss',fontSize=9,fontName='Helvetica',textColor=colors.HexColor('#666666'),spaceAfter=8)
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
                    lbl={'E':'Tendeur E (extensible)','L':'Tendeur L (long)','C':'Tendeur C (court)'}.get(typ,f'Tendeur {typ}')
                    addt(lbl,n,q)
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

    story=[Paragraph(title,ts),Paragraph("Rapport d'intervention nettoyé automatiquement",ss)]
    if vns: story.append(Paragraph(f"<b>Vidange groupe hydraulique recommandée</b> ({len(vns)}) : {', '.join(vns)}",ns))
    for lbl in sorted(tcats.keys()): story.append(Paragraph(fmt(lbl,tcats[lbl]),ns))
    for c,d in cats.items(): story.append(Paragraph(fmt(c,d),ns))
    story.append(Spacer(1,6))
    story.append(Table([['']],colWidths=[257*mm],style=TableStyle([('LINEABOVE',(0,0),(-1,-1),0.5,colors.HexColor('#cccccc'))])))
    story.append(Spacer(1,6))
    return story

# ── GUI ───────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Rapport Cleaner — Loading Systems")
        self.resizable(False, False)
        self.configure(bg='#f0f0f0')

        self.cfg = load_config()
        self.pdf_path = tk.StringVar()
        self.out_path = tk.StringVar()

        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

    def _build_ui(self):
        pad = {'padx':14,'pady':8}

        # Title
        tk.Label(self, text="Rapport Cleaner", font=('Helvetica',16,'bold'),
                 bg='#f0f0f0', fg='#222').pack(pady=(18,2))
        tk.Label(self, text="Loading Systems — nettoyage automatique des rapports d'intervention",
                 font=('Helvetica',9), bg='#f0f0f0', fg='#666').pack(pady=(0,14))

        # PDF input
        frm1 = tk.LabelFrame(self, text=" Fichier source ", bg='#f0f0f0', font=('Helvetica',9))
        frm1.pack(fill='x', **pad)
        tk.Entry(frm1, textvariable=self.pdf_path, width=52,
                 state='readonly').pack(side='left', padx=8, pady=8)
        tk.Button(frm1, text="Parcourir...", command=self._pick_pdf,
                  width=12).pack(side='left', padx=(0,8))

        # Output
        frm2 = tk.LabelFrame(self, text=" Fichier de sortie ", bg='#f0f0f0', font=('Helvetica',9))
        frm2.pack(fill='x', **pad)
        tk.Entry(frm2, textvariable=self.out_path, width=52,
                 state='readonly').pack(side='left', padx=8, pady=8)
        tk.Button(frm2, text="Choisir...", command=self._pick_out,
                  width=12).pack(side='left', padx=(0,8))

        # Progress
        self.progress = ttk.Progressbar(self, mode='indeterminate', length=400)
        self.progress.pack(pady=(6,0))

        # Log
        frm3 = tk.LabelFrame(self, text=" Journal ", bg='#f0f0f0', font=('Helvetica',9))
        frm3.pack(fill='both', expand=True, **pad)
        self.log_box = scrolledtext.ScrolledText(frm3, height=8, width=62,
            font=('Courier',8), state='disabled', bg='#1e1e1e', fg='#d4d4d4')
        self.log_box.pack(padx=8, pady=8)

        # Buttons
        btn_frm = tk.Frame(self, bg='#f0f0f0')
        btn_frm.pack(pady=(4,16))
        self.btn_run = tk.Button(btn_frm, text="▶  Générer le PDF propre",
            font=('Helvetica',11,'bold'), bg='#0066cc', fg='white',
            padx=20, pady=8, command=self._run, state='disabled')
        self.btn_run.pack(side='left', padx=8)
        tk.Button(btn_frm, text="Effacer le journal", command=self._clear_log,
                  padx=10, pady=8).pack(side='left', padx=8)

    def _pick_pdf(self):
        path = filedialog.askopenfilename(
            title="Choisir le rapport PDF",
            filetypes=[("Fichiers PDF","*.pdf"),("Tous","*.*")])
        if path:
            self.pdf_path.set(path)
            # Auto-suggest output path
            base = os.path.splitext(path)[0]
            self.out_path.set(base + '_clean.pdf')
            self._check_ready()
            self._log(f"📂 Fichier chargé : {os.path.basename(path)}")

    def _pick_out(self):
        path = filedialog.asksaveasfilename(
            title="Enregistrer le PDF nettoyé",
            defaultextension=".pdf",
            filetypes=[("Fichiers PDF","*.pdf")])
        if path:
            self.out_path.set(path)
            self._check_ready()

    def _check_ready(self):
        if self.pdf_path.get() and self.out_path.get():
            self.btn_run.config(state='normal')

    def _log(self, msg):
        self.log_box.config(state='normal')
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.log_box.config(state='disabled')

    def _clear_log(self):
        self.log_box.config(state='normal')
        self.log_box.delete('1.0','end')
        self.log_box.config(state='disabled')

    def _run(self):
        pdf = self.pdf_path.get()
        out = self.out_path.get()
        if not pdf or not out: return

        self.btn_run.config(state='disabled')
        self.progress.start(10)
        self._log("\n─── Démarrage du traitement ───")

        def worker():
            try:
                self._log("Analyse de la structure du rapport...")
                structure = detect_structure(pdf)
                if not structure:
                    self._log("❌ Impossible de lire la structure du PDF.")
                    return

                self._log(f"Structure détectée : {structure.get('style','?')} — colonnes : {list(structure.get('data_cols',{}).keys()) or 'nom/commentaire'}")

                # Check for unknown words and ask user
                self._log("Analyse du vocabulaire...")
                unknowns = detect_unknown_words(
                    pdf, structure,
                    self.cfg.get('corrections',{}),
                    self.cfg.get('blacklist_extra',[]))

                if unknowns:
                    # Filter out already-known words from config
                    known_already = set(self.cfg.get('corrections',{}).keys()) | \
                                    set(self.cfg.get('known_ok',[]))
                    new_unknowns = unknowns - known_already
                    if new_unknowns:
                        self.after(0, lambda u=new_unknowns: self._ask_unknowns(u, pdf, out, structure))
                        return

                self._do_generate(pdf, out, structure)

            except Exception as e:
                self.after(0, lambda: self._log(f"❌ Erreur : {e}"))
            finally:
                self.after(0, self._stop_progress)

        threading.Thread(target=worker, daemon=True).start()

    def _ask_unknowns(self, unknowns, pdf, out, structure):
        """Show dialog for unknown words — ask user what to do with each."""
        win = tk.Toplevel(self)
        win.title("Mots inconnus détectés")
        win.resizable(False, False)
        win.grab_set()

        tk.Label(win, text=f"Le rapport contient {len(unknowns)} mot(s) inhabituel(s).\nPour chacun, indiquez si c'est une faute de frappe ou un terme à ignorer.",
                 font=('Helvetica',9), justify='left', padx=12, pady=8).pack()

        frame = tk.Frame(win); frame.pack(padx=12, pady=4, fill='both', expand=True)
        canvas = tk.Canvas(frame, height=min(300, len(unknowns)*42+20))
        sb = tk.Scrollbar(frame, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y'); canvas.pack(side='left', fill='both', expand=True)
        inner = tk.Frame(canvas); canvas.create_window((0,0), window=inner, anchor='nw')

        decisions = {}  # word -> {'action': 'ok'/'correction'/'blacklist', 'value': str}

        for word in sorted(unknowns):
            row_f = tk.Frame(inner); row_f.pack(fill='x', pady=3, padx=4)
            tk.Label(row_f, text=f"  «{word}»", font=('Courier',9,'bold'), width=20, anchor='w').pack(side='left')
            action = tk.StringVar(value='ok')
            correction = tk.StringVar(value=word)

            tk.Radiobutton(row_f, text="OK, garder", variable=action, value='ok').pack(side='left')
            tk.Radiobutton(row_f, text="Corriger en :", variable=action, value='correction').pack(side='left')
            tk.Entry(row_f, textvariable=correction, width=16).pack(side='left')
            tk.Radiobutton(row_f, text="Ignorer/supprimer", variable=action, value='blacklist').pack(side='left')
            decisions[word] = (action, correction)

        inner.update_idletasks()
        canvas.config(scrollregion=canvas.bbox('all'))

        def confirm():
            for word, (action_var, corr_var) in decisions.items():
                act = action_var.get()
                if act == 'correction':
                    self.cfg.setdefault('corrections',{})[word] = corr_var.get()
                elif act == 'blacklist':
                    if word not in self.cfg.setdefault('blacklist_extra',[]):
                        self.cfg['blacklist_extra'].append(rf'\b{re.escape(word)}\b')
                else:
                    self.cfg.setdefault('known_ok',[])
                    if word not in self.cfg['known_ok']:
                        self.cfg['known_ok'].append(word)
            save_config(self.cfg)
            win.destroy()
            self._log(f"✓ {len(decisions)} mot(s) traité(s), configuration sauvegardée")
            self._do_generate(pdf, out, structure)

        tk.Button(win, text="✓ Confirmer et générer", command=confirm,
                  bg='#0066cc', fg='white', font=('Helvetica',10,'bold'),
                  padx=16, pady=6).pack(pady=10)

    def _do_generate(self, pdf, out, structure):
        def worker():
            try:
                generate_pdf(
                    pdf, out, structure,
                    self.cfg.get('corrections',{}),
                    self.cfg.get('blacklist_extra',[]),
                    log_fn=lambda m: self.after(0, lambda msg=m: self._log(msg))
                )
                self.after(0, lambda: messagebox.showinfo(
                    "Terminé",
                    f"PDF généré avec succès !\n\n{out}"))
            except Exception as e:
                self.after(0, lambda: self._log(f"❌ Erreur : {e}"))
                self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
            finally:
                self.after(0, self._stop_progress)
                self.after(0, lambda: self.btn_run.config(state='normal'))
        threading.Thread(target=worker, daemon=True).start()

    def _stop_progress(self):
        self.progress.stop()

if __name__ == '__main__':
    app = App()
    app.mainloop()
