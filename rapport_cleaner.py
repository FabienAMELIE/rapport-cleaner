"""
Rapport Cleaner — Loading Systems
Nettoie automatiquement les rapports d'intervention PDF de techniciens.
"""

import os, re, sys, json, threading, tempfile, urllib.request, urllib.error, subprocess, random
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
# Drag & drop : tkinterdnd2 pour supporter le glisser-déposer de fichiers
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
import pdfplumber
from PIL import Image as PILImage, ImageTk
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, Image as RLImage, Flowable)
from reportlab.lib.units import mm
from reportlab.pdfgen.canvas import Canvas as _BaseCanvas

VERSION = "V0.3.6.4.4"
GITHUB_REPO = "FabienAMELIE/rapport-cleaner"

# ── Phrases rigolotes du journal ──────────────────────────────────────────────
_FUNNY_LINES = [
    "Aucun panneau bas n'a été choqué lors de la génération de ce rapport...",
    "Obtention du numéro de Romana, à des fins professionnelles... évidemment...",
    "Versement du café d'Abdarrahime...",
    "Vérification que Laetitia ne soit pas énervée...",
    "Pause obligatoire pour goûter les gâteaux de Nicha...",
    "Installation des Bethouart de quais...",
    "Mise en relation avec la police municipale...",
    "Recalibration des girafes en mode paysage...",
    "Tentative de déchiffrage des éléments du rapport...",
    "Écoute de la 467ème fois où Gurvan vante les qualités de la Bretagne...",
    "Vérification que l'usine ne s'est pas encore trompée dans les commandes...",
    "Organisation d'une réunion commerciale interminable...",
    "Pourquoi tu lis ça ??...",
    "Demande à Cyril pour avoir les références...",
]

def _get_funny_lines():
    """Retourne 1 ou 2 phrases rigolotes tirées aléatoirement."""
    k = random.randint(1, 2)
    return random.sample(_FUNNY_LINES, k)
GITHUB_EXE_NAME = "RapportCleaner.exe"
UPDATER_FLAG = "--updated"

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
    'compacteur', 'benne',
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
    # Normalisation des chiffres exposants (ex: ⁸ → 8, causent sinon des ■ dans pdfplumber)
    _SUPERSCRIPTS = str.maketrans('⁰¹²³⁴⁵⁶⁷⁸⁹', '0123456789')
    text = text.translate(_SUPERSCRIPTS)
    text = re.sub(r'\r\n|\r', '\n', text)
    # Préserver les sauts de ligne intentionnels (double \n = séparateur d'anomalies)
    text = re.sub(r'\n\s*\n', ' §§ ', text)
    lines = text.split('\n')
    result = []; i = 0
    while i < len(lines):
        line = lines[i]
        if i+1 < len(lines):
            nxt = lines[i+1].strip()
            # Cas 1 : ligne suivante entiere = suffixe court alpha (comportement original)
            if re.fullmatch(r'[a-zA-ZÀ-ÿ]{1,4}', nxt) and nxt.lower() not in NEVER_SUFFIX:
                m = re.search(r'[a-zA-ZÀ-ÿ]+$', line.rstrip())
                if m:
                    lw = m.group()
                    ends_c = bool(re.search(r'[^aeiouàâéèêëîïôùûüyAEIOUÀÂÉÈÊËÎÏÔÙÛÜY]$', lw))
                    if len(nxt) == 1 and nxt.lower() not in {'s','x','e','é'}:
                        result.append(line); i+=1; continue
                    if (ends_c or len(lw)<=11) and len(lw+nxt)>=5:
                        result.append(line.rstrip()[:m.start()]+lw+nxt); i+=2; continue
            else:
                # Cas 2 : prefixe alpha court suivi d'autre chose sur la ligne suivante
                # Ex: "maintenan\nce/ cellule..." -> "maintenance / cellule..."
                # Ex: "asservissem\nent hs" -> "asservissement hs"
                # Ex: "Sectionneur\nr hs" -> "Sectionneur hs" (doublon de lettre)
                pfx_m = re.match(r'^([a-zA-ZÀ-ÿ]{1,4})([^a-zA-ZÀ-ÿ].*)?$', nxt)
                if pfx_m:
                    pfx = pfx_m.group(1)
                    rest = (pfx_m.group(2) or '').lstrip()
                    m = re.search(r'[a-zA-ZÀ-ÿ]+$', line.rstrip())
                    if m:
                        lw = m.group()
                        ends_c = bool(re.search(r'[^aeiouàâéèêëîïôùûüyAEIOUÀÂÉÈÊËÎÏÔÙÛÜY]$', lw))
                        # Doublon de lettre : "Sectionneur\nr hs" -> supprimer le 'r' residuel
                        if len(pfx) == 1 and lw.lower().endswith(pfx.lower()):
                            lines[i+1] = rest
                            result.append(line.rstrip()); i+=1; continue
                        # Suffixe non-isole : fusionner le prefixe, reinjecter le reste
                        # Ne pas fusionner si pfx commence par une majuscule (≥2 chars) :
                        # c'est le début d'un nouvel élément, pas une coupure de mot
                        # Ex: "HS\nChoc PB + PI" → ne pas produire "HSChoc PB + PI"
                        if len(pfx) >= 2 and pfx[0].isupper():
                            result.append(line); i+=1; continue
                        if pfx.lower() not in NEVER_SUFFIX:
                            if len(pfx) == 1 and pfx.lower() not in {'s','x','e','é'}:
                                # Exception : "r" après voyelle finale reconstruit -eur/-ier/-oir
                                # Ex: "supérieu" + "r" → "supérieur"
                                if pfx.lower() == 'r' and re.search(r'[aeouiéèêàùûî]$', lw, re.IGNORECASE):
                                    pass  # autorise la fusion
                                else:
                                    result.append(line); i+=1; continue
                            if (ends_c or len(lw)<=11) and len(lw+pfx)>=5:
                                lines[i+1] = rest
                                result.append(line.rstrip()[:m.start()]+lw+pfx); i+=1; continue
        result.append(line); i+=1
    joined = re.sub(r' {2,}', ' ', ' '.join(result)).strip()
    # Restaurer les séparateurs intentionnels
    joined = joined.replace('§§', ' / ')
    # Post-processing : supprimer les fragments residuels (artefacts PDF 3 lignes)
    # Ex: 'asservissement ement hs' -> 'asservissement hs'  ('ement' = fin de 'asservissement')
    # Ex: 'remplacement ent /'     -> 'remplacement /'
    # Ex: 'sectionneur r hs'       -> 'sectionneur hs'      ('r' = doublon)
    # Exclusions : mots courts legitimes du francais (NEVER_SUFFIX) + 'ce','se','me','te','ne'
    _RESIDUAL_EXCL = NEVER_SUFFIX | {'ce', 'se', 'me', 'te', 'ne', 'que', 'elle', 'ment'}
    _ALWAYS_DROP_FRAGMENTS = {'ement', 'issement'}
    def _drop_residual(m):
        word, frag = m.group(1), m.group(2)
        if frag.lower() in _ALWAYS_DROP_FRAGMENTS:
            return word
        if word.lower().endswith(frag.lower()) and frag.lower() not in _RESIDUAL_EXCL:
            return word
        return m.group(0)
    joined = re.sub(r'([a-zA-ZÀ-ÿ]{5,}) ([a-zA-ZÀ-ÿ]{1,8})\b', _drop_residual, joined)
    return joined

def strip_choc(text):
    if not text: return text
    t = text.strip()
    if not re.search(r'\bchoc(?:ué?e?)?\b', t, re.IGNORECASE): return t
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
    # Étape 1 : filtrer les segments commençant par "choc" sans HS
    # Tagguer chaque segment : True = issu d'un segment choc, False = segment non-choc à préserver
    segs = re.split(r'\s*\+\s*', t)
    kept_tagged = []
    prev_was_choc = False  # tracker indépendant pour détecter segment choc même si résidu vide
    for seg in segs:
        seg_s = seg.strip()
        if re.match(r'^(?:léger\s+)?choc\b', seg_s, re.IGNORECASE):
            prev_was_choc = True
            if re.search(r'\bHS\b', seg_s, re.IGNORECASE):
                kept_tagged.append((True, seg_s))
            else:
                residual = choc_re.sub('', seg_s).strip()
                residual = re.sub(r'^[\s\+\-,;.]+', '', residual).strip()
                if residual:
                    kept_tagged.append((True, residual))
        else:
            seg_s_clean = seg_s
            # Si le segment précédent était un segment-choc, supprimer le préfixe pb/pi/ph
            # orphelin en tête (ex: "Choc PB + PI Suspente..." → après split, "PI Suspente...")
            if prev_was_choc:
                seg_s_clean = re.sub(r'^(?:pb|pi|ph)\s+', '', seg_s_clean, flags=re.IGNORECASE).strip()
            # Si CE segment contient "choc" (milieu), le suivant peut aussi avoir un préfixe orphelin
            prev_was_choc = bool(re.search(r'\bchoc\b', seg_s, re.IGNORECASE))
            kept_tagged.append((False, seg_s_clean))
    # Étape 2 : appliquer nettoyage UNIQUEMENT aux segments choc-dérivés
    _CLEANUP_PATS = [
        r'\b\d{3,4}x\d{3,4}(?:x\d+)?\b', r'\bx\d+\b(?!\s*cm)',
        r'\b(?:nordsud|nord|sud|ral\s*\d*)\b',
        r'\b(?:intérieur|extérieur|interieur|exterieur)\s+(?:\w+\s+)?\d{3,4}\b',
        r'\bet\s+(?:intérieur|extérieur)\s+\w+\b',
        r'\b\d+\s*x\b(?!\s*\d)',
        r'\b(?:extérieur|intérieur|exterieur|interieur)\s*(?:et\s*)?$',
        r'\b(?:inter|panneau|bas|haut|léger|leger|droite|gauche)\b',  # + haut (point 3)
        r'\bpoutre\s+avant\b',
    ]
    result_parts = []
    for is_choc, part in kept_tagged:
        if not part: continue
        if is_choc:
            # Segment choc : nettoyage complet
            p = choc_re.sub('', part)
            p = re.sub(r'\bchoquée?\b', '', p, flags=re.IGNORECASE)
            for pat in _CLEANUP_PATS:
                p = re.sub(pat, '', p, flags=re.IGNORECASE)
            p = re.sub(r'\bpanneau(bas|haut|intermédiaire|intermediaire)\b', r'Panneau \1', p, flags=re.IGNORECASE)
            p = re.sub(r'^[\s\+\-,;]+', '', p)
            p = re.sub(r'[\s\+\-,;]+$', '', p)
            p = re.sub(r' {2,}', ' ', p).strip()
            if p: result_parts.append(p)
        elif re.search(r'\bchoc\b', part, re.IGNORECASE):
            prev_was_choc = True  # le segment suivant peut avoir un préfixe pb/pi/ph orphelin
            # Choc au milieu du segment : supprimer la description choc + résidus
            # Après split sur +, ce segment ne contient jamais profil/joint → _CLEANUP_PATS safe
            p = choc_re.sub('', part)
            p = re.sub(r'\bchoquée?\b', '', p, flags=re.IGNORECASE)
            for pat in _CLEANUP_PATS:
                p = re.sub(pat, '', p, flags=re.IGNORECASE)
            p = re.sub(r'^[\s\+\-,;.]+', '', p)
            p = re.sub(r'[\s\+\-,;]+$', '', p)
            p = re.sub(r' {2,}', ' ', p).strip()
            if p: result_parts.append(p)
        else:
            # Segment sans choc : préserver intégralement
            result_parts.append(part)
    s = ' + '.join(result_parts)
    if has_hs:
        garbled = bool(re.match(r'^[a-zA-Z]\s*[,\.]',s)) or \
                  len(re.findall(r'\b[a-zA-ZÀ-ÿ]\b',s))>2 or (s and len(s)<8)
        if garbled or not s: return t
    # Si le résultat est un seul mot isolé (résidu du choc), on considère que tout est du choc
    if s and len(s.split()) <= 1 and not has_hs:
        return ''
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

# Paires de mots du domaine qui peuvent être fusionnés sans espace
_FUSED_PATTERNS = [
    # tendeur(s) + qualificatif
    (r'\btendeurs?(longs?|courts?|extensibles?)\b', lambda m: 'tendeurs ' + m.group(1)),
    # panneau + position
    (r'\bpanneau(bas|haut|intermédiaire|intermediaire)\b', lambda m: 'panneau ' + m.group(1)),
    # flexible + type
    (r'\bflexible(verin|vérin)\b', lambda m: 'flexible ' + m.group(1)),
    # verin + type
    (r'\b(verin|vérin)(bavette|principal|lèvre|levre)\b', lambda m: m.group(1) + ' ' + m.group(2)),
    # choc + panneau colles sans espace : ChocPB -> Choc PB
    (r'\bchoc(pb|pi|ph)\b', lambda m: 'Choc ' + m.group(1).upper()),
    # chiffre collé à tendeur(s) (ex: "2tendeurlong" → "2 tendeur long")
    (r'(\d)(tendeurs?)', lambda m: m.group(1) + ' ' + m.group(2)),
    # tendeur(s) collé à un mot suivant (ex: "tendeurcourt" → "tendeur court")
    (r'(tendeurs?)([a-zA-ZÀ-ÿ])', lambda m: m.group(1) + ' ' + m.group(2)),
    # bâche + supérieur/e fusionnés (ex: "Bachesupérieur hs")
    (r'\bbaches?(sup[eé]rieure?)\b', lambda m: 'Bâche ' + m.group(1)),
    # absence + de
    (r'\babsencede\b', 'absence de'),
    # cellule + d
    (r'\bcellulede\b', "cellule d'"),
]

def fix_fused_words(text):
    if not text: return text
    for pat, repl in _FUSED_PATTERNS:
        if callable(repl):
            text = re.sub(pat, repl, text, flags=re.IGNORECASE)
        else:
            text = re.sub(pat, repl, text, flags=re.IGNORECASE)
    return text

def _propagate_hs(text):
    """Si le dernier segment contient HS, le propage aux segments précédents qui n'en ont pas.
    Ex: "Panneau bas + inter bas HS" → "Panneau bas HS + inter bas HS"
    Ne propage pas à travers un point (artefact ". HS" attaché à un segment non-HS).
    Ne propage pas aux segments "rampe/non accessible" (constats sans anomalie).
    """
    if not text or not re.search(r'\bHS\b', text, re.IGNORECASE): return text
    # Supprimer les HS orphelins après un point : "en nacelle. HS +" → "en nacelle. +"
    text = re.sub(r'\.\s+HS\b\s*(?=\+)', '.', text, flags=re.IGNORECASE)
    segs = re.split(r'\s*\+\s*', text.strip())
    if len(segs) <= 1: return text
    if not re.search(r'\bHS\b', segs[-1], re.IGNORECASE): return text
    _NO_PROPAGATE = re.compile(r'non\s+accessible|rampe\s+ressort|rampe\s+non', re.IGNORECASE)
    result = []
    for seg in segs:
        seg_s = seg.strip()
        if seg_s and not re.search(r'\bHS\b', seg_s, re.IGNORECASE) and not _NO_PROPAGATE.search(seg_s):
            seg_s = seg_s + ' HS'
        result.append(seg_s)
    return ' + '.join(result)

def clean_cell(text, corrections=None, blacklist=None):
    if not text: return ''
    t = fix_word_breaks(text)
    # Normaliser les '+' sans espaces
    t = re.sub(r'(?<=[^\s])\+(?=[^\s])', ' + ', t)
    # Normaliser les '/' sans espaces → ' / ' (ex: "hs/1long" → "hs / 1long")
    t = re.sub(r'\s*/\s*', ' / ', t)
    t = fix_fused_words(t)
    t = strip_choc(t)                              # passe 1 : avant corrections
    if corrections: t = apply_corrections(t, corrections)
    t = strip_choc(t)                              # passe 2 : après corrections
    t = _propagate_hs(t)                           # propagation HS au dernier segment

    # ── Filtre "effectué" ──────────────────────────────────────────────────────
    # Regex tolérant aux fautes : effectué/effectue/éffectué/efectue...
    _EFF_RE = re.compile(r'[eé]ff?ectu[eé]e?', re.IGNORECASE)
    _SEP_RE = re.compile(r'\s*/\s*|\s*\+\s*')
    if _EFF_RE.search(t):
        segs = _SEP_RE.split(t)
        if len(segs) > 1:
            # Avec séparateur : drop les segments contenant "effectué",
            # garder les autres (soumis aux règles normales)
            kept = [s.strip() for s in segs if not _EFF_RE.search(s)]
            t = ' / '.join(kept)
        else:
            # Sans séparateur : conservatif
            # effectué sans HS → drop ; effectué + HS → garde tout
            if not re.search(r'\bHS\b', t, re.IGNORECASE):
                t = ''
            # sinon : garde tout tel quel
    # Pas de cleanup résiduel "remplacement effectué" ici — la règle conservatrice
    # (sans séparateur + HS) doit garder tout le texte intact.
    t = re.sub(r'^[\s/\-,;.]+','',t); t = re.sub(r'[\s/\-,;.]+$','',t)
    t = re.sub(r' {2,}',' ',t).strip()
    # Strip word-level "vétuste/vetuste" (ex: "Sas vétuste" → "Sas" → vide → drop)
    _VET_RE = re.compile(r'\bv[eé]tuste\b|\busure\b', re.IGNORECASE)
    if _VET_RE.search(t):
        t_stripped = _VET_RE.sub('', t).strip()
        t_stripped = re.sub(r'^[\s/\-,;.]+','',t_stripped).strip()
        # Si après strip il ne reste que du bruit ou des mots-conteneurs sans valeur → drop
        _CONTAINER_WORDS = {'sas', 'niveleur', 'porte', 'portesectionnelle'}
        residual = re.sub(r'\s+', '', t_stripped.lower())
        if not t_stripped or residual in _CONTAINER_WORDS:
            return ''
        t = t_stripped
    bl = blacklist or []
    if is_blacklisted_full(t, bl): return ''
    t = strip_blacklisted_parts(t, bl)
    if is_blacklisted_full(t, bl): return ''
    return t

# ── Détection structure PDF ───────────────────────────────────────────────────
def _looks_like_numero_header(h):
    """Retourne True si le header ressemble à une colonne 'N°' ou 'Numéro'."""
    if not h: return False
    h = h.strip().lower()
    return bool(re.search(r'^#$|^n°?$|^n$|^n\s*°?\s*de\s*série|^numéros?$|^numero$|^num$|^n\s*°?$', h)) or \
           bool(re.search(r'^\s*numéros?\b|^\s*numero\b|^\s*n°\s', h))

def _looks_like_photo_header(h):
    """Retourne True si le header est une colonne Photo/Image."""
    if not h: return False
    h = h.strip().lower()
    # Matche : "Photo", "Photos", "Image", "Images", "Photo Porte", "Photo Quai", "Photo PS", "Photo +", etc.
    return bool(re.search(r'^photos?\b|^images?\b', h))

def _column_values(tables, col_idx):
    """Retourne la liste des valeurs (non-header) d'une colonne, nettoyées."""
    vals = []
    for ri, row in enumerate(tables[0]):
        if ri == 0: continue  # skip header
        if col_idx < len(row) and row[col_idx]:
            v = fix_word_breaks(row[col_idx]).strip()
            vals.append(v)
        else:
            vals.append('')
    return vals

def _column_looks_numeric(vals):
    """Retourne True si la colonne contient majoritairement des numéros/codes courts avec des chiffres."""
    if not vals: return False
    non_empty = [v for v in vals if v.strip()]
    if not non_empty: return False
    numeric_count = 0
    for v in non_empty:
        v = v.strip()
        # Une valeur "numérique" : doit contenir au moins un chiffre, ET être court (≤ 15 chars),
        # ET pas trop de mots (≤ 2 espaces). Ex: "201", "A12", "50B", "1 devant quai"
        # "Ras", "ressort droit et gauche déformé" ne passent PAS (pas de chiffre ou trop long)
        has_digit = bool(re.search(r'\d', v))
        if has_digit and len(v) <= 15 and v.count(' ') <= 2:
            numeric_count += 1
    return numeric_count / len(non_empty) >= 0.5

def _column_follows_row_index(vals):
    """Retourne True si les valeurs suivent exactement l'index de ligne (1, 2, 3, 4...)."""
    if not vals: return False
    for i, v in enumerate(vals, start=1):
        if v.strip() and v.strip() != str(i):
            return False
    return True

def detect_structure(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        tables = page.extract_tables()
        if not tables or not tables[0]: return None
        header_raw = tables[0][0]
        header = [re.sub(r'\s+', ' ', c).strip().lower() if c else '' for c in header_raw]
    # Robuste aux coupures pdfplumber ("Commenta\nire") et fautes de saisie ("Comentaire", "Commentair")
    # → supprime les espaces avant comparaison, cherche le préfixe "comment" (≥7 chars, sans ambiguïté)
    # → "nom"/"noms" vérifié par fullmatch pour éviter les faux positifs ("Dénomination", etc.)
    _h_nospace = [''.join(h.split()) for h in header]
    _has_nom = any(re.fullmatch(r'noms?', h) for h in _h_nospace)
    _has_commentaire = any('comment' in h for h in _h_nospace)
    if _has_nom and _has_commentaire:
        return {'style': 'nom_commentaire', 'headers': header}

    photo_cols = [i for i, h in enumerate(header) if _looks_like_photo_header(h)]

    n_col = None
    if len(header) >= 1:
        col0_header_numeric = _looks_like_numero_header(header[0])
        col1_header_numeric = len(header) >= 2 and _looks_like_numero_header(header[1])
        col0_vals = _column_values(tables, 0)
        col1_vals = _column_values(tables, 1) if len(header) >= 2 else []
        col0_numeric = _column_looks_numeric(col0_vals)
        col1_numeric = _column_looks_numeric(col1_vals)
        col0_follows = _column_follows_row_index(col0_vals)
        col1_follows = _column_follows_row_index(col1_vals)

        if col0_header_numeric or col0_numeric:
            if col1_header_numeric or col1_numeric:
                if not col1_follows or col1_header_numeric:
                    n_col = 1
                else:
                    n_col = 0
            else:
                n_col = 0
        elif col1_header_numeric or col1_numeric:
            n_col = 1

    if n_col is None: n_col = 0

    # Candidats : colonnes non photo, non n_col, non header numéro, non header vide
    candidate_cols = []
    for i in range(len(header)):
        if i == n_col: continue
        if i in photo_cols: continue
        if _looks_like_numero_header(header[i]): continue
        if not header[i].strip(): continue
        candidate_cols.append(i)

    # Scanner TOUTES les pages pour ne garder que les colonnes avec du contenu utile
    cols_with_content = set()
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_tables = page.extract_tables()
            if not page_tables: continue
            for ri, row in enumerate(page_tables[0]):
                if ri == 0: continue
                for ci in candidate_cols:
                    if ci >= len(row) or not row[ci]: continue
                    val = re.sub(r'\s+', ' ', row[ci]).strip()
                    if not val: continue
                    vl = val.lower()
                    if vl in ('ras', 'x', '/', '', 'condamné', 'condamne'): continue
                    cols_with_content.add(ci)

    data_col_indices = [i for i in candidate_cols if i in cols_with_content]

    data_col_labels = {}
    for i in data_col_indices:
        label = header_raw[i] if i < len(header_raw) and header_raw[i] else ''
        label = re.sub(r'\s+', ' ', label).strip()
        if label:
            label = label[0].upper() + label[1:]
        data_col_labels[i] = label or f"Colonne {i+1}"

    n_photos = max(2, min(6, len(photo_cols))) if photo_cols else 3

    # Stocker le header normalisé pour comparaison pages 2+
    header_normalized = [re.sub(r'\s+', ' ', c).strip().lower() if c else '' for c in header_raw]

    return {'style': 'standard', 'headers': header, 'n_col': n_col,
            'data_col_indices': data_col_indices,
            'data_col_labels': data_col_labels,
            'photo_cols': photo_cols,
            'n_photos': n_photos,
            'header_normalized': header_normalized}

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
    unknowns = {}  # mot → set de N° d'équipement où il apparaît
    header_normalized = None
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables: continue
            # Mémoriser le header de la page 1 pour détecter les répétitions pages 2+
            if header_normalized is None:
                header_normalized = [re.sub(r'\s+', ' ', c).strip().lower() if c else '' for c in tables[0][0]]
            for ri, row in enumerate(tables[0]):
                if ri == 0: continue  # toujours skip la ligne 0
                # Skip si c'est un header répété (pages 2+)
                row_norm = [re.sub(r'\s+', ' ', c).strip().lower() if c else '' for c in row]
                if row_norm == header_normalized:
                    continue
                # Trouver le N° de cette ligne (col 0 ou 1)
                n_val = ''
                for ci in range(min(2, len(row))):
                    if row[ci] and re.search(r'\d', row[ci]):
                        n_val = fix_word_breaks(row[ci]).strip()
                        break
                for cell in row:
                    if not cell: continue
                    t = fix_word_breaks(cell).lower()
                    t = apply_corrections(t, corrections)
                    for word in re.findall(r'[a-zA-ZÀ-ÿ]{4,}', t):
                        if word not in KNOWN and not is_blacklisted_full(word, blacklist):
                            if len(word) <= 15:
                                unknowns.setdefault(word, set())
                                if n_val: unknowns[word].add(n_val)
    return unknowns

# ── Génération PDF ────────────────────────────────────────────────────────────
TXT = colors.HexColor('#222222')
HEADER_BG = colors.HexColor('#404040')

def make_cell(text, bold=False, size=8, color=None):
    if color is None: color = TXT
    # Remplacer ' / ' par <br/> pour segmenter visuellement les anomalies multiples
    text = str(text).replace(' / ', '<br/>')
    return Paragraph(text, ParagraphStyle('c', fontSize=size,
        fontName='Helvetica-Bold' if bold else 'Helvetica',
        textColor=color, leading=size*1.3))

def make_img(name, img_dir):
    path = os.path.join(img_dir, f"{name}.jpg")
    if not os.path.exists(path): return ''
    try:
        with PILImage.open(path) as im: w, h = im.size
        # Max dimensions légèrement inférieures à la cellule pour éviter tout débordement
        mh, mw = 19*mm, 21*mm
        r = w/h
        if r >= 1:
            rw = min(mw, mh*r); rh = rw/r
        else:
            rh = min(mh, mw/r); rw = rh*r
        # Sécurité : jamais plus grand que la cellule
        if rh > mh: rw = rw * mh/rh; rh = mh
        if rw > mw: rh = rh * mw/rw; rw = mw
        return RLImage(path, width=rw, height=rh)
    except: return ''

def extract_and_map_images(pdf_path, img_dir, n_col=1):
    # Nettoyer les anciennes images pour éviter le cache périmé
    if os.path.exists(img_dir):
        for f in os.listdir(img_dir):
            if f.endswith(('.jpg', '.png', '.jpeg')):
                try: os.remove(os.path.join(img_dir, f))
                except: pass
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

def generate_pdf(pdf_path, output_path, structure, corrections, blacklist, log_fn=None, progress_fn=None):
    def log(msg):
        if log_fn: log_fn(msg)
    def progress(val):
        if progress_fn: progress_fn(val)
    style = structure.get('style', 'standard')
    img_dir = os.path.join(tempfile.gettempdir(), 'rapport_cleaner_imgs')
    img_n_col = structure.get('n_col', 0)
    if style == 'nom_commentaire': img_n_col = 0

    progress(55)
    log("Extraction des images...")
    img_map = extract_and_map_images(pdf_path, img_dir, n_col=img_n_col)
    log(f"  → {sum(len(v) for v in img_map.values())} image(s) pour {len(img_map)} ligne(s)")

    progress(70)
    log("Lecture du tableau...")
    if style == 'nom_commentaire':
        rows_data, quais = _read_nom_commentaire(pdf_path, corrections, blacklist)
    else:
        rows_data = _read_standard(pdf_path, structure, corrections, blacklist)
        quais = None

    progress(85)
    log("Génération du PDF...")
    # Phrases rigolotes glissées dans le journal
    funny = _get_funny_lines()
    for f in funny:
        log(f)
    _build_pdf(output_path, rows_data, img_map, img_dir, structure, quais, log)
    log(f"✓ PDF généré : {output_path}")

def _read_standard(pdf_path, structure, corrections, blacklist):
    n_col = structure.get('n_col', 0)
    data_col_indices = structure.get('data_col_indices', [])
    header_normalized = structure.get('header_normalized', [])
    rows_data = []; seen = set()
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables: continue
            for ri, row in enumerate(tables[0]):
                if not row: continue
                # Skip si c'est un header répété (comparer avec le header page 1)
                if ri == 0:
                    row_norm = [re.sub(r'\s+', ' ', c).strip().lower() if c else '' for c in row]
                    if row_norm == header_normalized:
                        continue  # header répété, on skip
                n_raw = row[n_col] if n_col < len(row) else ''
                # Ligne sans numéro : note globale technicien
                if not n_raw or not n_raw.strip():
                    note_parts = []
                    for idx in data_col_indices:
                        raw = row[idx] if idx < len(row) and row[idx] else ''
                        if raw.strip():
                            note_parts.append(fix_word_breaks(raw).strip())
                    if note_parts:
                        note_text = ' / '.join(note_parts)
                        padding = tuple([''] * max(0, len(data_col_indices) - 1))
                        rows_data.append((len(rows_data), '__NOTE__', note_text) + padding)
                    continue
                n = fix_word_breaks(n_raw).strip()
                if not n or n in ('N','N°','#','Numéros','Numéro','Numero','∑') or n in seen: continue
                if re.search(r'[a-zA-Z]{3,}', n) and not re.search(r'\d', n):
                    if n.lower() in ('n','n°','numéros','numéro','numero','image','photo','porte','niveleur','sas'): continue
                # Si le N° est un texte long (>15 chars avec espaces), c'est du contenu, pas un numéro
                # Ex: "Barrière manque 1 morceau la lisse normalement 6060/110 + support lisse"
                # Ex: "63 porte local maintenance" → n_display = "63", suffixe ignoré
                # Ex: "7 rampe" → déjà court, ne passe pas ici
                if len(n) > 15 and n.count(' ') >= 2:
                    # Cas : texte long commençant par un chiffre → extraire le chiffre + suffixe contextuel
                    m_num = re.match(r'^(\d+)\s*(.*)', n.strip())
                    if m_num:
                        num_part = m_num.group(1)
                        suffix = m_num.group(2).strip()
                        # Conserver le suffixe contextuel court (rampe, local...) mais ignorer les longs descriptifs
                        _GENERIC_LABELS = {'porte','entrée','entree','local','rampe','chaufferie',
                                           'bureau','accueil','couloir','escalier','sortie','accès','acces',
                                           'fenwick','maintenance','studio','photo'}
                        suffix_words = suffix.split()
                        # Garder uniquement si le suffixe est court ET contient un mot contextuel connu
                        if suffix_words and suffix_words[0].lower() in _GENERIC_LABELS and len(suffix_words) <= 3:
                            n_display = f"{num_part} {suffix_words[0]}"
                        else:
                            n_display = num_part
                    else:
                        # Texte long commençant par un mot (ex: "Barrière manque 1 morceau...")
                        n_display = re.sub(r'\d+', '', n.split()[0]).strip() or n.split()[0]
                        n_display = n_display.capitalize()
                    # Ignorer les labels génériques sans anomalie
                    _GENERIC_LABELS = {'porte','entrée','entree','local','rampe','chaufferie',
                                       'bureau','accueil','couloir','escalier','sortie','accès','acces'}
                    is_generic = n_display.lower() in _GENERIC_LABELS or \
                                 n.split()[0].lower() in _GENERIC_LABELS
                    all_data_empty = not any(
                        row[idx] and clean_cell(fix_word_breaks(row[idx]), corrections, blacklist)
                        for idx in data_col_indices if idx < len(row))
                    if is_generic and all_data_empty:
                        continue  # label de porte sans anomalie → skip
                    cleaned_n = clean_cell(n, corrections, blacklist) if not m_num else ''
                    if not cleaned_n and all_data_empty:
                        continue  # tout est vide, on skip
                    if n_display in seen: continue
                    seen.add(n_display)
                    fields = [cleaned_n] + [''] * max(0, len(data_col_indices) - 1)
                    # Remplir les autres colonnes normalement
                    for fi, idx in enumerate(data_col_indices):
                        raw = row[idx] if idx < len(row) and row[idx] else ''
                        val = clean_cell(raw, corrections, blacklist)
                        if val and fi < len(fields):
                            fields[fi] = fields[fi] + (' / ' + val if fields[fi] else val)
                    rows_data.append((len(rows_data), n_display) + tuple(fields))
                    continue
                seen.add(n)
                fields = []
                for idx in data_col_indices:
                    raw = row[idx] if idx < len(row) and row[idx] else ''
                    fields.append(clean_cell(raw, corrections, blacklist))
                rows_data.append((len(rows_data), n) + tuple(fields))
    return rows_data

def _extract_sublabel(nom):
    """Extrait le sous-label client/site depuis le champ Nom.
    Stratégie : après le premier nombre standalone dans le nom, tout ce qui reste
    (nettoyé des mots-clés techniques) devient le sous-label.
    Ex: "Porte sectionnelle motorisé horman 2 exertis connect" → "exertis connect"
    Ex: "Porte local poubelle" (sans numéro) → "local poubelle"
    """
    # Mots techniques à ignorer dans le sous-label
    _TECH = re.compile(
        r'\b(?:porte\s+section\w*|porte|niveleur|sas|motoris[eé]|horman|'
        r'bavette\s+pivotant\w*|sectionnell\w*|ramp\w*)\b',
        re.IGNORECASE
    )
    nom_clean = _TECH.sub('', nom).strip()
    # Trouver le premier nombre standalone
    m = re.search(r'\b(\d+)\b', nom_clean)
    if m:
        after = nom_clean[m.end():].strip(' .,;-')
        if after:
            return after.strip()
        return ''
    # Pas de numéro → tout le nom nettoyé est le sous-label
    return nom_clean.strip(' .,;-') if nom_clean.strip() else ''

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
                sublabel = _extract_sublabel(nom)
                if re.search(r'porte\s+section', nom, re.IGNORECASE):
                    m = re.search(r'(?:abloy|assa)\s+(\d+)', nom, re.IGNORECASE) or \
                        re.search(r'(\d+)\s*$', nom) or \
                        re.search(r'^(\d+)\b', nom) or \
                        re.search(r'\b(\d+)\b', nom)
                    if m:
                        q = quais.setdefault(int(m.group(1)),{})
                        q['porte'] = (serie, com)
                        if sublabel and 'sublabel' not in q:
                            q['sublabel'] = sublabel
                elif re.search(r'niveleur', nom, re.IGNORECASE):
                    m = re.search(r':\s*(\d+)', nom) or \
                        re.search(r'^(\d+)\b', nom) or \
                        re.search(r'(\d+)\s*$', nom) or \
                        re.search(r'\b(\d+)\b', nom)
                    if m:
                        q = quais.setdefault(int(m.group(1)),{})
                        q['niv'] = (serie, com)
                        if sublabel and 'sublabel' not in q:
                            q['sublabel'] = sublabel
                elif re.search(r'\bsas\b', nom, re.IGNORECASE):
                    m = re.search(r'(?:abloy)\s+(\d+)', nom, re.IGNORECASE) or \
                        re.search(r'^(\d+)\b', nom) or \
                        re.search(r'(\d+)\s*$', nom) or \
                        re.search(r'\b(\d+)\b', nom)
                    if m:
                        q = quais.setdefault(int(m.group(1)),{})
                        q['sas'] = (serie, com)
                        if sublabel and 'sublabel' not in q:
                            q['sublabel'] = sublabel
    rows_data = []
    for qn in sorted(quais.keys()):
        d = quais[qn]
        sublabel = d.get('sublabel', '')
        rows_data.append((qn, str(qn),
                          d.get('porte',('',''))[1],
                          '',
                          d.get('niv',  ('',''))[1],
                          d.get('sas',  ('',''))[1],
                          '', '', '', '',
                          sublabel))  # sous-label en position 10
    return rows_data, quais

def _build_pdf(output_path, rows_data, img_map, img_dir, structure, quais, log):
    style = structure.get('style','standard')
    
    if style == 'nom_commentaire':
        # Mode nom_commentaire : colonnes fixes Porte/Niveleur/SAS (logique existante)
        active_col_labels = ['Porte sectionnelle', 'Niveleur / Quai', 'SAS']
        N_PHOTOS = 3
    else:
        # Mode standard : on prend les labels tels quels depuis le PDF
        data_col_indices = structure.get('data_col_indices', [])
        data_col_labels = structure.get('data_col_labels', {})
        active_col_labels = [data_col_labels.get(i, f"Col {i+1}") for i in data_col_indices]
        N_PHOTOS = structure.get('n_photos', 3)
    
    n_data_cols = len(active_col_labels)
    N_IMG_W = 23*mm
    N_COL_W = 14*mm
    page_w = landscape(A4)[0] - 20*mm
    data_col_w = (page_w - N_PHOTOS*N_IMG_W - N_COL_W) / max(n_data_cols, 1)
    col_widths = [N_COL_W] + [data_col_w]*n_data_cols + [N_IMG_W]*N_PHOTOS
    headers = ['N°'] + active_col_labels + [f'Photo {i+1}' for i in range(N_PHOTOS)]
    header_row = [make_cell(h, bold=True, size=8, color=colors.white) for h in headers]
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
        n = row[1]
        if n == '__NOTE__': continue  # notes globales → récap uniquement
        fields = list(row[2:])
        # En mode nom_commentaire, le tuple est (porte, rapide_vide, niv, sas, ...)
        # On doit extraire porte/niv/sas
        if style == 'nom_commentaire':
            active_fields = [
                fields[0] if len(fields) > 0 else '',  # porte
                fields[2] if len(fields) > 2 else '',  # niv
                fields[3] if len(fields) > 3 else '',  # sas
            ]
        else:
            # En mode standard, les fields sont déjà dans l'ordre des data_col_indices
            active_fields = [fields[i] if i < len(fields) else '' for i in range(n_data_cols)]
        
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

        # Sous-label client/site (mode nom_commentaire uniquement)
        sublabel = row[10] if style == 'nom_commentaire' and len(row) > 10 else ''
        if sublabel:
            n_cell = make_cell(f'{n}\n{sublabel}', bold=True, size=7.5)
            # Remettre le sous-label en plus petit et italique via Paragraph
            from reportlab.platypus import Paragraph
            from reportlab.lib.styles import ParagraphStyle
            _n_style = ParagraphStyle('ncell', fontName='Helvetica-Bold', fontSize=7.5,
                                      alignment=1, leading=10)
            _sub_style = ParagraphStyle('subcell', fontName='Helvetica-Oblique', fontSize=6,
                                        alignment=1, leading=8, textColor=colors.HexColor('#666666'))
            n_cell = [Paragraph(n, _n_style), Paragraph(sublabel, _sub_style)]
        else:
            n_cell = make_cell(n, bold=True, size=7.5)

        table_data.append([n_cell]+data_cells+img_cells)
        idx=len(table_data)-1
        style_cmds.append(('BACKGROUND',(0,idx),(-1,idx),colors.HexColor('#f2f2f2') if alt%2==0 else colors.white))
        style_cmds.append(('ROWHEIGHT',(0,idx),(-1,idx),21*mm if imgs else 7*mm))
        alt+=1

    # Summary : on passe tous les fields + les labels pour que le summary puisse chercher partout
    summary_rows=[]; tech_notes=[]
    for row in rows_data:
        n=row[1]
        if n == '__NOTE__':
            note_text = row[2] if len(row) > 2 else ''
            if note_text: tech_notes.append(note_text)
            continue
        fields=list(row[2:])
        if style == 'nom_commentaire':
            summary_rows.append((0, n,
                                 fields[0] if len(fields)>0 else '',  # porte
                                 fields[2] if len(fields)>2 else '',  # niv
                                 fields[3] if len(fields)>3 else ''))  # sas
        else:
            summary_rows.append((0, n) + tuple(fields))

    fname = os.path.splitext(os.path.basename(output_path))[0]
    raw_title = fname.replace('_',' ').replace('-',' ')
    société = re.sub(r'\s*[\-_]?\s*(?:clean\s*(?:v\d+)?|v\d+)\s*$', '', raw_title, flags=re.IGNORECASE).strip()
    société = re.sub(r'\s{2,}', ' ', société).strip()
    story = _build_summary(summary_rows, société, active_col_labels, tech_notes, style)
    main_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    main_table.setStyle(TableStyle(style_cmds))
    story.append(main_table)

    class _LastPageCanvas(_BaseCanvas):
        """Affiche la version uniquement sur la derniere page, sans creer de page supplementaire."""
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._saved_page_states = []
        def showPage(self):
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()
        def save(self):
            n_pages = len(self._saved_page_states)
            for i, state in enumerate(self._saved_page_states):
                self.__dict__.update(state)
                if i == n_pages - 1:
                    self.saveState()
                    self.setFont('Helvetica', 7)
                    self.setFillColor(colors.HexColor('#aaaaaa'))
                    self.drawRightString(landscape(A4)[0] - 10*mm, 5*mm, VERSION)
                    self.restoreState()
                super().showPage()
            super().save()

    doc = SimpleDocTemplate(output_path, pagesize=landscape(A4),
        leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    doc.build(story, canvasmaker=_LastPageCanvas)

def _condense_summary_label(text):
    """Condense les textes avec dimensions pour le résumé.
    Ex: '3 raidisseurs diamètre 40mm longueur 4250mm 2 raidisseurs diamètre 30mm longueur 4250mm'
    → '5 raidisseurs'
    Ex: '3 sangles de longueur 7000mm chacune largeur 50mm'
    → '3 sangles'
    """
    if not text: return text
    # Détecter si le texte contient des dimensions (mm, cm, m, diamètre, longueur, largeur, épaisseur)
    if not re.search(r'\b(?:diamètre|diametre|longueur|largeur|épaisseur|epaisseur|\d+\s*(?:mm|cm|m)\b)', text, re.IGNORECASE):
        return text
    # Extraire les groupes "N items" et additionner par type d'item
    item_counts = {}
    for m in re.finditer(r'(\d+)\s+([a-zA-ZÀ-ÿ]+)', text):
        qty = int(m.group(1))
        item = m.group(2).strip().lower()
        # Ignorer les items qui sont des dimensions ou des prépositions
        if item in ('mm','cm','m','de','du','des','la','le','les','diamètre','diametre',
                     'longueur','largeur','épaisseur','epaisseur','x','chacune','chacun'):
            continue
        item_counts[item] = item_counts.get(item, 0) + qty
    if item_counts:
        parts = [f"{q} {name}" for name, q in item_counts.items()]
        return ' + '.join(parts)
    return text

class CheckboxField(Flowable):
    """Case à cocher AcroForm — cochable nativement dans Acrobat/Adobe Reader."""
    def __init__(self, name, size=4):
        Flowable.__init__(self)
        self.name = name
        sz = size * mm
        self.width = sz
        self.height = sz

    def draw(self):
        # Coordonnées absolues via la matrice de transformation courante
        m = self.canv._currentMatrix
        abs_x, abs_y = m[4], m[5]
        self.canv.acroForm.checkbox(
            name=self.name,
            tooltip=self.name,
            x=abs_x, y=abs_y,
            size=self.height,
            checked=False,
            buttonStyle='check',
            borderStyle='solid',
            borderColor=colors.black,
            fillColor=colors.white,
            textColor=colors.black,
            forceBorder=True,
        )


class CheckboxWithTextField(Flowable):
    """Case à cocher + champ texte fournisseur côte à côte (AcroForm)."""
    def __init__(self, cb_name, tf_name, total_width, size=4):
        Flowable.__init__(self)
        self.cb_name = cb_name
        self.tf_name = tf_name
        sz = size * mm
        self.cb_size = sz
        self.gap = 2 * mm
        self.width = total_width
        self.height = sz
        self.tf_width = max(total_width - sz - self.gap, 10*mm)

    def draw(self):
        m = self.canv._currentMatrix
        abs_x, abs_y = m[4], m[5]
        # Case à cocher à gauche
        self.canv.acroForm.checkbox(
            name=self.cb_name,
            tooltip=self.cb_name,
            x=abs_x, y=abs_y,
            size=self.cb_size,
            checked=False,
            buttonStyle='check',
            borderStyle='solid',
            borderColor=colors.black,
            fillColor=colors.white,
            textColor=colors.black,
            forceBorder=True,
        )
        # Champ texte fournisseur à droite
        self.canv.acroForm.textfield(
            name=self.tf_name,
            tooltip='Nom du fournisseur',
            x=abs_x + self.cb_size + self.gap,
            y=abs_y,
            width=self.tf_width,
            height=self.cb_size,
            borderColor=colors.HexColor('#aaaaaa'),
            fillColor=colors.HexColor('#fafafa'),
            textColor=colors.black,
            fontSize=7,
            fontName='Helvetica',
            borderWidth=0.5,
            borderStyle='solid',
            forceBorder=True,
        )

def _build_summary(rows_data, title, active_col_labels=None, tech_notes=None, style='standard'):
    if active_col_labels is None: active_col_labels = ['Porte sectionnelle','Niveleur / Quai','SAS']
    if tech_notes is None: tech_notes = []
    ts=ParagraphStyle('ts',fontSize=13,fontName='Helvetica-Bold',spaceAfter=8,alignment=1)
    ns=ParagraphStyle('n',fontSize=8.5,fontName='Helvetica',textColor=colors.HexColor('#333333'),spaceAfter=4,leading=13)
    note_style=ParagraphStyle('note',fontSize=8.5,fontName='Helvetica-Oblique',
        textColor=colors.HexColor('#555555'),spaceAfter=4,leading=13)
    def eq(seg):
        s = seg.strip()
        # Priorité : chiffre en début "3 tendeurs", "2 x ..."
        m = re.match(r'^\s*(\d+)\s*(?:[xX*×]\s*)?', s)
        if m and m.group(1) and int(m.group(1)) > 0:
            return int(m.group(1))
        # Chiffre en fin précédé de x, *, × : "Tendeur long x3", "Tendeur long *2", "Tendeur long ×1"
        m2 = re.search(r'[xX*×]\s*(\d+)\s*$', s)
        return int(m2.group(1)) if m2 else 1
    cats={}; tcats={}; vns=[]
    _label_canon = {}  # normalized_key -> canonical display label
    def _norm_label(lbl):
        import unicodedata
        # Normalise cassé/cassée → hs pour fusionner les labels
        lbl = re.sub(r'\bcass[eé]e?\b', 'hs', lbl, flags=re.IGNORECASE)
        nfd = unicodedata.normalize('NFD', lbl.lower())
        return ''.join(c for c in nfd if not unicodedata.combining(c))
    def add(c,n,q=1):
        # Normaliser cassé/cassée → hs dans le label affiché aussi
        c = re.sub(r'\bcass[eé]e?\b', 'hs', c, flags=re.IGNORECASE)
        nk = _norm_label(c)
        if nk not in _label_canon: _label_canon[nk] = c
        can = _label_canon[nk]
        cats.setdefault(can,{}); cats[can][n]=cats[can].get(n,0)+q
    def addt(l,n,q=1): tcats.setdefault(l,{}); tcats[l][n]=tcats[l].get(n,0)+q
    # Détection colonne SAS par son label (robuste aux noms variés : "SAS", "Sas d'étanchéité", etc.)
    sas_indices = set()
    for ci, label in enumerate(active_col_labels):
        if label and re.search(r'\bsas\b', label, re.IGNORECASE):
            sas_indices.add(ci)
    for row in rows_data:
        n=row[1]
        for ci,f in enumerate(row[2:]):
            if not f: continue
            fl=f.lower(); is_sas=(ci in sas_indices)
            # Vidange : detection stricte (vidange explicite OU quantite d'huile mentionnee)
            # 'hydraulique' seul (ex: reference materiel) ne declenche plus la regle
            if 'vidange' in fl or re.search(r"\d+\s*l\s*d'huile", fl, re.IGNORECASE):
                if n not in vns: vns.append(n)
            if is_sas and ('tendeur' in fl or 'crochet' in fl or
                    re.search(r'\b\d+\s*(?:courts?|longs?|extensibles?)\b',fl) or
                    re.search(r'\b\d+\s*[cls]\b',fl) or
                    re.search(r'(?:courts?|longs?|extensibles?)\s*[xX*×]\s*\d+',fl)):
                # Split sur + , et aussi sur transitions chiffre→type pour capturer "3 long 2 court 2 s"
                raw_segs = re.split(r'[+,&/\\]', fl)
                segs_sas = []
                for rs in raw_segs:
                    rs = rs.strip()
                    # Redécouper sur transitions \d+ (long|court|extensible|crochet|s)
                    parts = re.split(r'(?<=\s)(?=\d+\s*(?:longs?|courts?|extensibles?|crochets?|[lces]\b))', rs)
                    segs_sas.extend([p.strip() for p in parts if p.strip()])
                for seg in segs_sas:
                    seg=seg.strip(); q=eq(seg)
                    if re.search(r'\bcrochets?\b',seg) or re.search(r'\b\d+\s*s\b',seg):
                        addt('Crochet S',n,q); continue
                    ht='tendeur' in seg; he='extensible' in seg
                    sc=bool(re.search(r'\b\d+\s*courts?\b|\b\d+\s*c\b|courts?\s*[xX*×]\s*\d+',seg))
                    sl=bool(re.search(r'\b\d+\s*longs?\b|\b\d+\s*l\b|longs?\s*[xX*×]\s*\d+',seg))
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
                not_vetuste = not bool(re.search(r'\bvétuste\b|\bvetuste\b|\busure\b', fl))
                # Traitement segment par segment : chaque segment est testé contre les règles connues
                # Les segments non reconnus tombent dans le bloc dynamique
                # Cela évite que matched=True global fasse perdre des segments non reconnus
                # Pre-decoupage : split sur separateurs classiques (+, /, &, ,)
                _raw_segs = [s.strip() for s in re.split(r'\s*[+/&,\\]\s*', f.strip()) if s.strip()]

                # Pre-decoupage intelligent (V0.3.3) : sur segments longs (>60 chars) sans separateur,
                # detecter les frontieres entre anomalies via mots-cles connus.
                # Deux passes : (1) apres ponctuation / apres HS / avant N+kw (case-insensitive)
                #               (2) avant mot-cle capitalise mid-phrase (case-sensitive)
                _SKW = [
                    'verrou','poignée','poignee','suspente','câble','cable',
                    'flexible','roulette','hublot','bavette','parachute',
                    'butée','butee','rail','montant','support','galet',
                    'corde','soudure','charnière','charniere','cellule',
                    'carte','moteur','spot','luminaire','barre','panneau',
                ]
                _skw = '|'.join(_SKW)
                _PAT1 = (r'(?<=[.!?])\s+'
                         r'|(?<=hs)\s+(?=(?:' + _skw + r'))'
                         r'|\s+(?=\d+\s+(?:' + _skw + r'))')
                _PAT2 = r'(?<=[a-zà-ÿ\d])\s+(?=(?:' + '|'.join(w[0].upper()+w[1:] for w in _SKW) + r'))'
                segments_full = []
                for raw_seg in _raw_segs:
                    if len(raw_seg) > 60 and not re.search(r'[+/&,\\]', raw_seg):
                        p1 = [s.strip() for s in re.split(_PAT1, raw_seg, flags=re.IGNORECASE) if s.strip()]
                        sub = []
                        for seg1 in p1:
                            if len(seg1) > 40:
                                sub.extend([s.strip() for s in re.split(_PAT2, seg1) if s.strip()])
                            else:
                                sub.append(seg1)
                        segments_full.extend(sub)
                    else:
                        segments_full.append(raw_seg)

                has_bavette_flexible = False  # non utilisé, conservé pour compatibilité

                _EFF_SUM = re.compile(r'[eé]ff?ectu[eé]e?', re.IGNORECASE)
                _SEP_SUM = re.compile(r'\s*/\s*|\s*\+\s*')
                for seg in segments_full:
                    sl = seg.lower()
                    seg_matched = False
                    q = eq(seg)
                    # Cas sans séparateur + HS + effectué → catch-all brut (trop ambigu à catégoriser)
                    if _EFF_SUM.search(seg) and re.search(r'\bHS\b', seg, re.IGNORECASE):
                        if len(_SEP_SUM.split(seg)) == 1:
                            label = seg.strip().rstrip('.')
                            label = re.sub(r'\s*\(([^)]+)\)', r' \1', label).strip()
                            add(label, n)
                            seg_matched = True

                    # Panneaux, profils et joints
                    hp = bool(re.search(r'\bpanneau\b|\binterm[eé]diaire\b|\bpi\b|\bpb\b|\bph\b', sl))
                    hj = bool(re.search(r'\bjoint\b', sl))
                    hprofil = bool(re.search(r'\bprofil\b', sl))
                    if hp or hj or hprofil:
                        seg_matched = True
                        has_profil = hprofil
                        has_joint  = hj
                        has_panneau = hp
                        is_bas     = bool(re.search(r'\bbas\b|\bpb\b', sl))
                        is_haut    = bool(re.search(r'\bhaut\b|\bph\b', sl))
                        is_lat     = bool(re.search(r'\blat[eé]ral\b', sl))
                        is_inter   = bool(re.search(r'\bpanneau\s+(?:inter\w*|interm[eé]diaire)\b|\bpi\b|\binterm[eé]diaire\b', sl))
                        # Quantité panneau : ignorer les cotes dimensionnelles (ex: 2850x610)
                        m_pq = re.match(r'^\s*(\d+)\s+(?:panneau|profil|joint)', seg, re.IGNORECASE)
                        pq = int(m_pq.group(1)) if m_pq else 1
                        # Cas profil/joint : prioritaire si présent (pas un remplacement de panneau)
                        if has_profil or has_joint:
                            if has_profil and has_joint:
                                if is_bas:  add('Profil & joint bas HS', n, pq)
                                elif is_lat: add('Profil & joint latéral HS', n, pq)
                                elif is_haut: add('Profil & joint haut HS', n, pq)
                                else: add('Profil & joint HS', n, pq)
                            elif has_profil:
                                if is_bas:  add('Profil panneau bas HS', n, pq)
                                elif is_lat: add('Profil latéral HS', n, pq)
                                elif is_haut: add('Profil haut HS', n, pq)
                                else: add('Profil HS', n, pq)
                            else:  # joint seul
                                if is_bas:  add('Joint bas HS', n, pq)
                                elif is_lat: add('Joint latéral HS', n, pq)
                                elif is_haut: add('Joint haut HS', n, pq)
                                else: add('Joint HS', n, pq)
                        # Panneau sans profil/joint → remplacement panneau
                        if has_panneau and not has_profil and not has_joint:
                            if re.search(r'\bpanneau\s+bas\b|\bpb\b', sl): add('Panneau bas', n, pq)
                            if re.search(r'\bpanneau\s+haut\b|\bph\b', sl): add('Panneau haut', n, pq)
                            if is_inter: add('Panneau intermédiaire', n, pq)

                    if 'suspente' in sl: add('Suspente à refaire', n); seg_matched = True
                    if 'flexible' in sl and not_vetuste:
                        seg_matched = True
                        has_bav = bool(re.search(r'\b(?:bavette|lèvre|levre)\b', sl))
                        has_pri = 'principal' in sl
                        if has_bav and has_pri:
                            add('Flexible vérin bavette HS', n); add('Flexible vérin principal HS', n)
                        elif has_bav:
                            add('Flexible vérin bavette HS', n)
                        elif has_pri:
                            add('Flexible vérin principal HS', n)
                        else:
                            add('Flexible HS', n, eq(sl))
                    if 'verrou' in sl and not_vetuste: add('Verrou HS', n); seg_matched = True
                    if 'roulette' in sl: add('Roulette manquante', n, eq(sl)); seg_matched = True
                    # Regle omega : prioritaire sur 'Cable HS' generique
                    if re.search(r'\boméga\b|\bomega\b|passage\s+de\s+c[âa]ble\s+ext', sl, re.IGNORECASE):
                        add('Oméga passage de câble HS', n); seg_matched = True
                    elif ('câble' in sl or 'cable' in sl) and not_vetuste:
                        has_repl = 'remplacement' in sl or 'remplacer' in sl
                        has_hs_seg = 'hs' in sl
                        is_traction_levage = bool(re.search(r'c[âa]bles?\s+de\s+(traction|levage)', sl, re.IGNORECASE))
                        if is_traction_levage and not has_hs_seg and not has_repl:
                            add('Câble de traction/levage à remplacer', n); seg_matched = True
                        elif has_hs_seg or has_repl:
                            seg_matched = True
                            if re.search(r'spiral', sl): add('Câble spiralé HS', n)
                            elif re.search(r'contact\s*mou?e?', sl): add('Contact mou de câble HS', n)
                            elif re.search(r'tirage|levage|traction', sl): add('Câble de traction HS', n)
                            else: add('Câble HS', n)
                    if ('butée' in sl or 'butee' in sl) and not_vetuste:
                        has_equerre = bool(re.search(r'\béquerre\b|\bequerre\b', sl))
                        has_butee = bool(re.search(r'\bbutée\b|\bbutee\b', sl))
                        if has_equerre and has_butee: add('Équerre et butée HS', n); seg_matched = True
                        elif has_equerre: add('Équerre de butée HS', n); seg_matched = True
                        else:
                            has_support = bool(re.search(r'\bsupport\b', sl))
                            if has_support: add('Support butée HS', n, eq(sl))
                            else: add('Butée HS', n, eq(sl))
                            seg_matched = True
                    if 'bavette' in sl and not_vetuste and 'flexible' not in sl:
                        if re.search(r'\btordu[e]?\b|d[eé]form[eé][e]?', sl):
                            add('Bavette tordue', n)
                        else:
                            add('Bavette HS', n)
                        seg_matched = True
                    if 'hublot' in sl and not_vetuste: add('Hublot HS', n); seg_matched = True
                    if 'parachute' in sl and not_vetuste: add('Parachute HS', n); seg_matched = True
                    if 'moteur' in sl and not_vetuste: add('Moteur HS', n); seg_matched = True
                    if ('relais' in sl or 'carte' in sl or 'électronique' in sl) and not_vetuste: add('Défaut électronique', n); seg_matched = True
                    if ('spot' in sl or 'luminaire' in sl or 'led' in sl) and not_vetuste: add('Éclairage HS', n); seg_matched = True
                    if ('poignée' in sl or 'poignet' in sl) and not_vetuste: add('Poignée HS', n); seg_matched = True
                    if 'chasse' in sl and not_vetuste: add('Chasse-pied HS', n); seg_matched = True
                    if ('béquille' in sl or 'bequille' in sl): add('Béquille sécurité absente', n); seg_matched = True
                    if 'charnière' in sl and not_vetuste: add('Charnière HS', n); seg_matched = True
                    if 'soudure' in sl: add('Soudure à refaire', n); seg_matched = True
                    if 'traverse' in sl: add('Traverse déformée', n); seg_matched = True
                    if 'cellule' in sl or 'asservissement' in sl: add('Absence cellule asservissement', n); seg_matched = True
                    # Ampoules : toutes variantes fusionnées par position int/ext
                    if re.search(r'\bampoule\b', sl) and not_vetuste:
                        if re.search(r'\bint[eé]r', sl): add('Ampoule feu rouge intérieur HS', n)
                        elif re.search(r'\bext[eé]r|\bext\b', sl): add('Ampoule feu rouge extérieur HS', n)
                        else: add('Ampoule HS', n)
                        seg_matched = True
                    # Bagues de serrage : agrégation avec quantité
                    if re.search(r'\bbague\b', sl) and re.search(r'\bserrage\b', sl):
                        m_bq = re.search(r'(\d+)\s+bague', sl)
                        bq = int(m_bq.group(1)) if m_bq else 1
                        addt('Ajout bague de serrage', n, bq)
                        seg_matched = True
                    # Exclusion résumé : constats d'accès sans anomalie à chiffrer
                    if re.search(r'rampe\s+(?:ressort\s+)?non\s+accessible', sl, re.IGNORECASE):
                        seg_matched = True  # supprimé du résumé
                    # Exclusion résumé : constats/observations sans action de chiffrage
                    _EXCLUDE_SUMMARY = [
                        r'fonctionnement\s+en\s+homme\s+mort',
                        r'niveleur\s+fonctionne\s+m[eê]me\s+porte\s+ferm',
                        r'cable\s+dans\s+l.armoire\s+.*non\s+branch',
                        r'cable\s+non\s+branch',
                        r'aimant\s+pour\s+retirer\s+les\s+cales',
                        r'pr[eé]voir\s+une?\s+p[aâ]te',
                        r'p[aâ]te\s+sur\s+la\s+porte',
                        r'calematic',
                        r'atalian\s+va\s+faire',
                    ]
                    if not seg_matched and any(re.search(p, sl, re.IGNORECASE) for p in _EXCLUDE_SUMMARY):
                        seg_matched = True
                    if ('corde' in sl or 'câble' in sl and 'traction' in sl) and not_vetuste:
                        no_action = any(x in sl for x in ['a fixer','a refixer','stock client','remplacement effectué'])
                        if not no_action: add('Corde HS', n); seg_matched = True

                    # Segments vétuste : jamais dans le résumé
                    if not seg_matched and ('vétuste' in sl or 'vetuste' in sl):
                        seg_matched = True  # supprime du résumé
                    # Segment non reconnu → bloc dynamique
                    if not seg_matched:
                        # Point 5 : axe + dimension -> "Axe a remplacer"
                        if re.search(r'\baxe\b', sl) and re.search(r'\d+[,.]?\d*\s*(?:mm|creux|plein)', sl, re.IGNORECASE):
                            add('Axe à remplacer', n); seg_matched = True
                    if not seg_matched:
                        label = _condense_summary_label(seg.strip().rstrip('.'))
                        # Normaliser les parentheses pour eviter les doublons de label
                        label = re.sub(r'\s*\(([^)]+)\)', r' \1', label).strip()
                        # Point 6 : bibliotheque de patterns courts pour segments verbeux
                        _SHORT_PATTERNS = [
                            (r'rail\b.{0,30}\bcomplet\b|prévoir\s+rail\s+complet|prevoir\s+rail\s+complet', 'Remplacement rail complet'),
                            (r'rail\b.{0,30}\b(tordu|détaché|detache|déformé|deforme)', 'Rail HS'),
                            (r'montant\b.{0,30}\b(tordu|détaché|detache|déformé|deforme)', 'Montant tordu'),
                            (r'nacelle\s+(requise|nécessaire|necessaire|toucan|toucan|3b|prévoir|prevoir)', 'Nacelle requise'),
                            (r'pas\s+d.accès\s+nacelle|pas\s+d.acces\s+nacelle|prévoir\s+une\s+nacelle|prevoir\s+une\s+nacelle', 'Nacelle requise'),
                            (r'besoin\s+d.un\s+chariot|chariot\b.{0,20}\bquai', 'Chariot nécessaire'),
                            (r'\bgalet[s]?\b.{0,20}\b(long[s]?|remplacer|remplacé|hs)', 'Galets à remplacer'),
                            (r'coffret\s+de\s+commande', 'Coffret de commande à remplacer'),  # normalise tous les variants
                            (r'b[aâ]che?\s+p[eé]riph[eé]r', 'Bâche périphérique HS'),
                            (r'b[aâ]che?\s+lat[eé]r', 'Bâche latérale HS'),
                        ]
                        _PROBLEM_INDICATOR = re.compile(
                            r'\bhs\b|cass[eé]e?|tordu[e]?|manque|manquant[e]?|absent[e]?|'
                            r'fuite|d[eé]faut|urgent|urgence|'
                            r'[àa]\s+remplacer|[àa]\s+refaire|[àa]\s+replacer|[àa]\s+changer|'
                            r'd[eé]tach|d[eé]form|bris[eé]e?|rompu[e]?|bloqu[eé]e?|'
                            r'manquante|soudure|remplacement',
                            re.IGNORECASE
                        )
                        if label:
                            for pat, short_lbl in _SHORT_PATTERNS:
                                if re.search(pat, sl, re.IGNORECASE):
                                    label = short_lbl
                                    break
                        if label and _PROBLEM_INDICATOR.search(label): add(label, n)
    def fmt(lbl,d):
        tot=sum(d.values()); parts=[f"{k} ({q})" if q>1 else k for k,q in d.items()]
        return f"<b>{lbl}</b> ({tot}) : {', '.join(parts)}"
    # Titre centré : "Nom société — Rapport d'intervention"
    titre_final = f"{title} — Rapport d'intervention"

    # Logo en haut à droite
    story = []
    story.append(Spacer(1, 6*mm))  # marge supérieure : centre visuellement titre/logo
    logo_path = resource_path('LS_LOGO_HOR_RGB_TRANSPARANT.png')
    if os.path.exists(logo_path):
        try:
            logo_h = 14*mm
            with PILImage.open(logo_path) as im:
                lw, lh = im.size
            logo_w = logo_h * lw / lh
            logo = RLImage(logo_path, width=logo_w, height=logo_h)
            # Tableau 1 ligne : titre centré + logo à droite (largeur utile ≈ 277mm)
            page_content_w = landscape(A4)[0] - 20*mm
            logo_col_w = logo_w + 4*mm
            title_col_w = page_content_w - logo_col_w
            # Table 3 colonnes : spacer | titre centré | logo
            # Le spacer gauche = logo_col_w => titre parfaitement centré sur la page entière
            center_col_w = page_content_w - 2 * logo_col_w
            header_table = Table(
                [['', Paragraph(titre_final, ts), logo]],
                colWidths=[logo_col_w, center_col_w, logo_col_w]
            )
            header_table.setStyle(TableStyle([
                ('ALIGN',    (1,0),(1,0), 'CENTER'),
                ('ALIGN',    (2,0),(2,0), 'RIGHT'),
                ('VALIGN',   (0,0),(-1,-1), 'MIDDLE'),
                ('LEFTPADDING',  (0,0),(-1,-1), 0),
                ('RIGHTPADDING', (0,0),(-1,-1), 0),
                ('TOPPADDING',   (0,0),(-1,-1), 0),
                ('BOTTOMPADDING',(0,0),(-1,-1), 4),
            ]))
            story.append(header_table)
        except Exception as e:
            story.append(Paragraph(titre_final, ts))
    else:
        story.append(Paragraph(titre_final, ts))
    if vns: story.append(Paragraph(f"<b>Vidange groupe hydraulique recommandée</b> ({len(vns)}) : {', '.join(vns)}",ns))
    # ── Tableau résumé avec cases à cocher ───────────────────────────────────
    page_content_w = landscape(A4)[0] - 20*mm
    TEXT_W  = 175*mm   # colonne anomalie
    FOUR_W  = 48*mm    # "Devis fournisseur demandé" (case + champ texte)
    CB_W    = page_content_w - TEXT_W - FOUR_W  # "Intégré dans le devis"
    FOUR_INNER = FOUR_W - 8*mm
    hdr_cb = ParagraphStyle('hcb', fontSize=7.5, fontName='Helvetica-Bold',
                             alignment=1, textColor=colors.white, leading=10)
    summary_items = []
    # Ordre d'affichage des tendeurs/crochets : E > L > Crochet S > C
    _TENDEUR_ORDER = ['Tendeur E (extensible)', 'Tendeur L (long)', 'Tendeur C (court)', 'Crochet S']
    for lbl in _TENDEUR_ORDER:
        if lbl in tcats: summary_items.append(fmt(lbl, tcats[lbl]))
    for lbl in sorted(tcats.keys()):
        if lbl not in _TENDEUR_ORDER: summary_items.append(fmt(lbl, tcats[lbl]))
    # Panneaux consécutifs (PB -> PI -> PH) puis reste des cats
    def _group_key(lbl):
        ll = lbl.lower()
        if 'bague' in ll: return (1, 'a', lbl)
        if 'panneau bas' in ll or 'profil & joint bas' in ll or 'joint bas' in ll or 'profil panneau bas' in ll: return (2, 'a', lbl)
        if 'panneau inter' in ll: return (2, 'b', lbl)
        if 'panneau haut' in ll: return (2, 'c', lbl)
        if 'joint' in ll and 'profil' in ll: return (3, 'a', lbl)
        if 'profil' in ll: return (3, 'b', lbl)
        if 'joint' in ll: return (3, 'c', lbl)
        if 'ampoule' in ll: return (4, 'a', lbl)
        if 'butée' in ll or 'butee' in ll: return (5, 'a', lbl)
        if 'bavette' in ll: return (6, 'a', lbl)
        if 'soudure' in ll: return (7, 'a', lbl)
        return (10, 'z', lbl)
    for c, d in sorted(cats.items(), key=lambda x: _group_key(x[0])):
        summary_items.append(fmt(c, d))
    if summary_items:
        tbl_data = [[
            Paragraph('', ns),
            Paragraph('Devis fournisseur demandé', hdr_cb),
            Paragraph('Intégré dans<br/>le devis', hdr_cb),
        ]]
        for i, text in enumerate(summary_items):
            tbl_data.append([
                Paragraph(text, ns),
                CheckboxWithTextField(f'fournisseur_{i}', f'fournisseur_nom_{i}', FOUR_INNER),
                CheckboxField(f'devis_{i}'),
            ])
        sum_tbl = Table(tbl_data, colWidths=[TEXT_W, FOUR_W, CB_W])
        sum_tbl.setStyle(TableStyle([
            ('BACKGROUND',    (1, 0), (2, 0),  HEADER_BG),
            ('ALIGN',         (1, 0), (2, -1), 'CENTER'),
            ('VALIGN',        (0, 0), (-1, -1),'MIDDLE'),
            ('TOPPADDING',    (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING',   (0, 0), (-1, -1), 4),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
            ('LINEBELOW',     (0, 0), (-1, -1), 0.3, colors.HexColor('#dddddd')),
            ('ROWHEIGHT',     (0, 0), (-1,  0), 8*mm),
        ]))
        story.append(sum_tbl)
    for note in tech_notes:
        story.append(Paragraph(f"<b>Note technicien :</b> {note}", note_style))
    story.append(Spacer(1, 8))
    return story

# ── Fenêtre Paramètres ────────────────────────────────────────────────────────
# ── Système de mise à jour ────────────────────────────────────────────────────

def check_latest_version():
    """Interroge l'API GitHub et retourne (tag_latest, download_url) ou (None, None) si erreur."""
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        req = urllib.request.Request(url, headers={"User-Agent": "RapportCleaner"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode())
        tag = data.get("tag_name", "")
        for asset in data.get("assets", []):
            if asset.get("name") == GITHUB_EXE_NAME:
                return tag, asset["browser_download_url"]
        return tag, None
    except Exception:
        return None, None

def do_update(download_url, progress_fn=None, log_fn=None):
    """MAJ intégrée dans la barre de progression :
    0→90%  : téléchargement (progression réelle)
    90→95% : renommage ancien exe
    95→100%: placement nouvel exe
    100%   : auto-relaunch + fermeture
    Au prochain démarrage, _old.exe est supprimé.
    ⚠️ auto-relaunch à tester SentinelOne (V0.3.6.4.2)
    """
    def log(msg):
        if log_fn: log_fn(msg)
    def prog(v):
        if progress_fn: progress_fn(v)
    try:
        import time
        exe_path = os.path.abspath(sys.executable)
        exe_dir  = os.path.dirname(exe_path)
        exe_name = os.path.basename(exe_path)
        old_path = os.path.join(exe_dir, exe_name.replace('.exe', '_old.exe'))
        new_tmp  = os.path.join(exe_dir, exe_name.replace('.exe', '_update.exe'))

        # ── Téléchargement 0→90% ──────────────────────────────────────────────
        def _reporthook(block, block_size, total):
            if total > 0:
                pct = min(int(block * block_size / total * 90), 90)
                prog(pct)
        urllib.request.urlretrieve(download_url, new_tmp, reporthook=_reporthook)
        prog(90)

        # ── Renommage 90→95% ─────────────────────────────────────────────────
        if os.path.exists(old_path):
            try: os.remove(old_path)
            except: pass
        os.rename(exe_path, old_path)
        prog(95)

        # ── Placement 95→100% ────────────────────────────────────────────────
        os.rename(new_tmp, exe_path)
        prog(100)

        # ── Auto-relaunch + fermeture ─────────────────────────────────────────
        time.sleep(0.8)  # laisser voir la barre verte
        try:
            subprocess.Popen([exe_path, UPDATER_FLAG])
        except Exception:
            pass
        os._exit(0)  # sortie forcée : court-circuite le nettoyage _MEI de PyInstaller → pas de warning

    except Exception as e:
        log(f"Erreur mise à jour : {e}")


class SettingsWindow(tk.Toplevel):
    def __init__(self, parent, cfg, on_save, open_tab=0):
        super().__init__(parent)
        self.title("Paramètres — Rapport Cleaner")
        self.configure(bg=C_BG)
        self.resizable(True, True)
        self.geometry("700x520")
        self.grab_set()
        self.cfg = cfg
        self.on_save = on_save
        self._open_tab = open_tab
        self._build()
        if self._open_tab: self._nb.select(self._open_tab)
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
        s.configure('S.TNotebook.Tab',background=C_PANEL,foreground=C_TEXT2,padding=[14,7],font=('Helvetica',10))
        s.map('S.TNotebook.Tab',background=[('selected',C_CARD)],foreground=[('selected',C_TEXT)])
        nb=ttk.Notebook(self,style='S.TNotebook')
        nb.pack(fill='both',expand=True,padx=12,pady=12)
        self._nb = nb

        # ── Onglet 0 — Options générales ─────────────────────────────────────
        tab0=tk.Frame(nb,bg=C_CARD); nb.add(tab0,text='  Options générales  ')
        tk.Label(tab0,text="Apparence",font=('Helvetica',10,'bold'),bg=C_CARD,fg=C_TEXT).pack(anchor='w',padx=16,pady=(16,6))
        tk.Frame(tab0,bg=C_BORDER,height=1).pack(fill='x',padx=16,pady=(0,12))
        theme_f=tk.Frame(tab0,bg=C_CARD); theme_f.pack(fill='x',padx=16,pady=(0,8))
        tk.Label(theme_f,text="Mode d'affichage :",bg=C_CARD,fg=C_TEXT,font=('Helvetica',9),width=20,anchor='w').pack(side='left')
        self.theme_var=tk.StringVar(value=self.cfg.get('theme','clair'))
        # ☀ et ☾ sont des symboles Unicode — s'alignent nativement avec le texte Tkinter
        for val,lbl in [('clair','☀  Mode clair'),('sombre','☾  Mode sombre')]:
            tk.Radiobutton(theme_f,text=lbl,variable=self.theme_var,value=val,
                           bg=C_CARD,fg=C_TEXT,selectcolor=C_CARD,
                           activebackground=C_CARD,font=('Helvetica',9),padx=12).pack(side='left',padx=4)
        tk.Label(tab0,text="Le changement de thème sera appliqué au prochain démarrage de l'application.",
                 bg=C_CARD,fg=C_TEXT2,font=('Helvetica',8),wraplength=580).pack(anchor='w',padx=16,pady=(8,0))

        # ── Onglet 1 — Blacklist ──────────────────────────────────────────────
        tab1=tk.Frame(nb,bg=C_CARD); nb.add(tab1,text='  Blacklist  ')
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
        tab2=tk.Frame(nb,bg=C_CARD); nb.add(tab2,text='  Corrections  ')
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
        tab3=tk.Frame(nb,bg=C_CARD); nb.add(tab3,text='  Mots acceptés  ')
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

        # ── Onglet 4 — Mises à jour ──────────────────────────────────────────────
        tab4=tk.Frame(nb,bg=C_CARD); nb.add(tab4,text='  Mises à jour  ')
        tk.Label(tab4,text="Vérifier si une nouvelle version de Rapport Cleaner est disponible.",
                 bg=C_CARD,fg=C_TEXT2,font=('Helvetica',8)).pack(anchor='w',padx=16,pady=(16,4))
        tk.Label(tab4,text=f"Version actuelle : {VERSION}",
                 bg=C_CARD,fg=C_TEXT,font=('Helvetica',9,'bold')).pack(anchor='w',padx=16,pady=(0,12))
        tk.Frame(tab4,bg=C_BORDER,height=1).pack(fill='x',padx=16,pady=(0,14))
        self._upd_status=tk.StringVar(value="")
        self._upd_lbl=tk.Label(tab4,textvariable=self._upd_status,bg=C_CARD,fg=C_TEXT2,
                               font=('Helvetica',9),wraplength=540,justify='left')
        self._upd_lbl.pack(anchor='w',padx=16,pady=(0,10))
        # Bouton téléchargement masqué par défaut
        self._upd_download_url=None
        upd_btn_f=tk.Frame(tab4,bg=C_CARD); upd_btn_f.pack(anchor='w',padx=16)
        tk.Button(upd_btn_f,text="Vérifier les mises à jour",command=self._check_update,
                  bg=C_PANEL,fg=C_TEXT,relief='flat',padx=18,pady=8,cursor='hand2',
                  font=('Helvetica',10,'bold')).pack(side='left',padx=(0,10))
        # Bouton créé comme enfant de upd_btn_f (pas tab4) pour que pack() fonctionne correctement
        self._upd_btn_dl=tk.Button(upd_btn_f,text="Télécharger et installer",
                                   command=self._do_update,bg=C_ACCENT,fg='white',
                                   relief='flat',padx=18,pady=8,cursor='hand2',
                                   font=('Helvetica',10,'bold'))
        # Masqué par défaut, révélé quand une MAJ est détectée

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

    def _check_update(self):
        self._upd_status.set("Vérification en cours...")
        self._upd_lbl.config(fg=C_TEXT2)
        self._upd_btn_dl.pack_forget()
        def worker():
            tag, url = check_latest_version()
            if tag is None:
                self.after(0, lambda: (
                    self._upd_status.set("Impossible de contacter GitHub. Vérifiez votre connexion."),
                    self._upd_lbl.config(fg='#e05050')))
            elif tag == VERSION:
                self.after(0, lambda: (
                    self._upd_status.set(f"✓ Vous avez déjà la dernière version ({VERSION})."),
                    self._upd_lbl.config(fg=C_SUCCESS)))
            else:
                self._upd_download_url = url
                def show_dl(t=tag, u=url):
                    self._upd_status.set(f"Nouvelle version disponible : {t}\nCliquez sur le bouton pour mettre à jour.")
                    self._upd_lbl.config(fg=C_ACCENT)
                    if u:
                        self._upd_btn_dl.pack(side='left', padx=(10,0))
                    else:
                        self._upd_status.set(f"Nouvelle version disponible : {t}\nMais le fichier .exe est introuvable dans la release.")
                self.after(0, show_dl)
        threading.Thread(target=worker, daemon=True).start()

    def _do_update(self):
        if not self._upd_download_url: return
        self._upd_btn_dl.config(state='disabled')

        # ── Fenêtre de progression modale ────────────────────────────────────
        prog_win = tk.Toplevel(self)
        prog_win.title("Mise à jour")
        prog_win.configure(bg=C_BG)
        prog_win.resizable(False, False)
        prog_win.grab_set()
        prog_win.attributes('-topmost', True)
        prog_win.protocol("WM_DELETE_WINDOW", lambda: None)
        pw, ph = 380, 100
        self.update_idletasks()
        sx = self.winfo_x() + (self.winfo_width()  - pw) // 2
        sy = self.winfo_y() + (self.winfo_height() - ph) // 2
        prog_win.geometry(f"{pw}x{ph}+{sx}+{sy}")

        prog_bar_var = tk.IntVar(value=0)
        prog_bar = ttk.Progressbar(prog_win, mode='determinate',
                                   variable=prog_bar_var, maximum=100, length=340)
        prog_bar.pack(pady=(22, 8), padx=20)
        prog_lbl = tk.Label(prog_win, text="Téléchargement...",
                            bg=C_BG, fg=C_TEXT2, font=('Helvetica', 8))
        prog_lbl.pack()

        def _set_progress(v):
            prog_bar_var.set(v)
            if v < 90:
                prog_lbl.config(text=f"Téléchargement... {v}%")
            elif v < 95:
                prog_lbl.config(text="Renommage...")
            elif v < 100:
                prog_lbl.config(text="Installation...")
            else:
                prog_bar.config(style='green.Horizontal.TProgressbar')
                prog_lbl.config(text="Mise à jour installée ✓")
            prog_win.update_idletasks()

        # Style barre verte pour la complétion
        _s = ttk.Style()
        _s.configure('green.Horizontal.TProgressbar',
                     troughcolor=C_PANEL, background='#2ecc71')

        def worker():
            do_update(self._upd_download_url,
                      progress_fn=lambda v: self.after(0, lambda val=v: _set_progress(val)),
                      log_fn=None)
        threading.Thread(target=worker, daemon=True).start()

    def _save(self):
        self.cfg['theme']=self.theme_var.get()
        save_config(self.cfg)
        self.on_save()
        self.destroy()

# ── Application principale ────────────────────────────────────────────────────
# Classe de base : TkinterDnD.Tk si drag & drop disponible, sinon tk.Tk
_AppBase = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk

class App(_AppBase):
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
            _ico_img = ImageTk.PhotoImage(file=resource_path('LS_LOGO_VERT_RGB_TRANSPARANT.png'))
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

        # Titre + logo centré + bouton paramètres
        title_f=tk.Frame(self,bg=C_BG); title_f.pack(fill='x',padx=20,pady=(14,8))
        # Grid à 3 colonnes : espacement gauche (poids 1) | logo centré | bouton paramètres (poids 1)
        title_f.columnconfigure(0, weight=1, uniform='titlecol')
        title_f.columnconfigure(1, weight=0)
        title_f.columnconfigure(2, weight=1, uniform='titlecol')

        # Logo Loading Systems (horizontal) centré
        try:
            _logo_pil = PILImage.open(resource_path('LS_LOGO_HOR_RGB_TRANSPARANT.png'))
            _lh = 56  # plus grand qu'avant (était 38)
            _lw = int(_logo_pil.width * _lh / _logo_pil.height)
            _logo_pil = _logo_pil.resize((_lw, _lh), PILImage.LANCZOS)
            _logo_tk = ImageTk.PhotoImage(_logo_pil)
            lbl_logo = tk.Label(title_f, image=_logo_tk, bg=C_BG)
            lbl_logo.image = _logo_tk
            lbl_logo.grid(row=0, column=1, sticky='')
        except Exception as e:
            print(f"Logo accueil : {e}")
            tk.Label(title_f,text="Rapport Cleaner",font=('Helvetica',18,'bold'),bg=C_BG,fg=C_TEXT).grid(row=0, column=1)

        tk.Button(title_f,text="≡  Paramètres",command=self._open_settings,
                  bg=C_PANEL,fg=C_TEXT2,relief='flat',padx=12,pady=5,
                  font=('Helvetica',9),cursor='hand2',
                  activebackground=C_CARD,activeforeground=C_TEXT).grid(row=0, column=2, sticky='e')

        tk.Frame(self,bg=C_BORDER,height=1).pack(fill='x',padx=20)

        # Fichier source (avec drag & drop si disponible)
        self._section("Fichier source")
        sf=tk.Frame(self,bg=C_PANEL); sf.pack(fill='x',padx=20,pady=(0,12))
        si=tk.Frame(sf,bg=C_PANEL); si.pack(fill='x',padx=12,pady=10)
        drop_hint = "  Glisser-déposer un PDF ici ou cliquer sur Parcourir" if DND_AVAILABLE else "Aucun fichier sélectionné"
        self.lbl_src=tk.Label(si,text=drop_hint,bg=C_ENTRY_BG,fg=C_TEXT2,
                              font=('Helvetica',9),anchor='w',padx=10,pady=8,width=55)
        self.lbl_src.pack(side='left',fill='x',expand=True)
        tk.Button(si,text="Parcourir...",command=self._pick_pdf,bg=C_ACCENT,fg='white',
                  relief='flat',padx=14,pady=8,font=('Helvetica',9,'bold'),cursor='hand2',
                  activebackground=C_ACCENT2,activeforeground='white').pack(side='left',padx=(8,0))

        # Activer le drag & drop sur le champ source
        if DND_AVAILABLE:
            for widget in (self.lbl_src, si, sf):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind('<<Drop>>', self._on_drop)
                widget.dnd_bind('<<DropEnter>>', lambda e: self.lbl_src.config(bg=C_CARD))
                widget.dnd_bind('<<DropLeave>>', lambda e: self.lbl_src.config(bg=C_ENTRY_BG))

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

        # Progress bar (mode déterministe, cachée au repos)
        pf=tk.Frame(self,bg=C_BG); pf.pack(fill='x',padx=20,pady=(0,4))
        s=ttk.Style(); s.theme_use('default')
        s.configure('Red.Horizontal.TProgressbar',troughcolor=C_ENTRY_BG,
                    background=C_ACCENT,borderwidth=0,lightcolor=C_ACCENT,darkcolor=C_ACCENT)
        s.configure('Green.Horizontal.TProgressbar',troughcolor=C_ENTRY_BG,
                    background=C_SUCCESS,borderwidth=0,lightcolor=C_SUCCESS,darkcolor=C_SUCCESS)
        self.progress=ttk.Progressbar(pf,mode='determinate',maximum=100,value=0,length=560,
                                      style='Red.Horizontal.TProgressbar')
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
        # Version + bannière MAJ disponible (en bas à droite)
        self._version_lbl = tk.Label(bf,text=VERSION,font=('Helvetica',8),bg=C_BG,fg=C_TEXT2,cursor='arrow')
        self._version_lbl.pack(side='right',pady=(6,0))
        # Vérification silencieuse au démarrage (après 2s pour laisser l'UI se charger)
        self.after(2000, self._startup_update_check)

    def _startup_update_check(self):
        """Vérification silencieuse de mise à jour au démarrage. Si nouvelle version → bannière cliquable."""
        def worker():
            try:
                url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
                req = urllib.request.Request(url, headers={'User-Agent': 'RapportCleaner'})
                with urllib.request.urlopen(req, timeout=5) as r:
                    data = json.loads(r.read())
                tag = data.get('tag_name', '')
                assets = data.get('assets', [])
                dl_url = next((a['browser_download_url'] for a in assets
                               if a['name'].lower().endswith('.exe')), None)
                if tag and tag != VERSION and dl_url:
                    self.after(0, lambda t=tag, u=dl_url: self._show_update_banner(t, u))
            except: pass
        threading.Thread(target=worker, daemon=True).start()

    def _show_update_banner(self, new_tag, dl_url):
        """Affiche une bannière discrète non-modale en bas à droite si une MAJ est disponible."""
        self._update_banner_url = dl_url
        self._version_lbl.config(
            text=f"↑ {new_tag} disponible",
            fg=C_ACCENT, cursor='hand2',
            font=('Helvetica', 8, 'underline'))
        self._version_lbl.bind('<Button-1>', lambda e: self._open_settings_update_tab())

    def _open_settings_update_tab(self):
        """Ouvre la fenêtre Paramètres directement sur l'onglet Mises à jour."""
        win = SettingsWindow(self, self.cfg,
                             on_save=lambda: self._log("✓ Paramètres sauvegardés"),
                             open_tab=4)
        win.focus_force()

    def _section(self, label):
        f=tk.Frame(self,bg=C_BG); f.pack(fill='x',padx=20,pady=(8,4))
        tk.Label(f,text=label.upper(),font=('Helvetica',7,'bold'),bg=C_BG,fg=C_TEXT2).pack(side='left')
        tk.Frame(f,bg=C_BORDER,height=1).pack(side='left',fill='x',expand=True,padx=(8,0),pady=4)

    def _pick_pdf(self):
        path=filedialog.askopenfilename(title="Choisir le rapport PDF",
            filetypes=[("Fichiers PDF","*.pdf"),("Tous","*.*")])
        if path:
            self._set_pdf_path(path)

    def _on_drop(self, event):
        """Callback quand un fichier est déposé sur la zone source."""
        # Restaurer le fond normal
        if DND_AVAILABLE:
            self.lbl_src.config(bg=C_ENTRY_BG)
        # event.data contient les chemins séparés par espaces, entourés de { } si nom contient espace
        raw = event.data.strip()
        # Nettoyer le format tkdnd : "{chemin 1} {chemin 2}" ou "chemin1 chemin2"
        paths = []
        if raw.startswith('{'):
            # Plusieurs fichiers avec espaces dans le nom
            import re as _re
            paths = _re.findall(r'\{([^}]*)\}', raw)
            # Ajouter aussi les chemins hors accolades
            remaining = _re.sub(r'\{[^}]*\}', '', raw).strip()
            if remaining:
                paths.extend(remaining.split())
        else:
            paths = raw.split()
        # Ne traiter que le premier PDF trouvé
        for p in paths:
            p = p.strip('"').strip("'")
            if p.lower().endswith('.pdf') and os.path.isfile(p):
                self._set_pdf_path(p)
                return
        # Aucun PDF valide
        self._log("⚠ Aucun fichier PDF valide dans le drop")

    def _set_pdf_path(self, path):
        """Applique un chemin PDF au champ source et prépare la sortie par défaut."""
        self.pdf_path.set(path)
        name = os.path.basename(path)
        self.lbl_src.config(text=f"  {name}", fg=C_TEXT)
        out = os.path.splitext(path)[0] + '_clean.pdf'
        self.out_path.set(out)
        self.lbl_out.config(text=f"  {os.path.basename(out)}", fg=C_TEXT)
        self._check_ready()
        self._log(f"📂 {name}")

    def _pick_out(self):
        # Pré-remplir avec le nom et le dossier du fichier de sortie déjà proposé
        current_out = self.out_path.get()
        init_dir = os.path.dirname(current_out) if current_out else ''
        init_file = os.path.basename(current_out) if current_out else ''
        path=filedialog.asksaveasfilename(title="Enregistrer le PDF nettoyé",
            defaultextension=".pdf",filetypes=[("Fichiers PDF","*.pdf")],
            initialdir=init_dir, initialfile=init_file)
        if path:
            self.out_path.set(path)
            self.lbl_out.config(text=f"  {os.path.basename(path)}",fg=C_TEXT)
            self._check_ready()

    def _check_ready(self):
        if self.pdf_path.get() and self.out_path.get():
            self.btn_run.config(state='normal')

    def _open_pdf(self, path):
        """Ouvre le PDF généré avec le lecteur par défaut."""
        try:
            os.startfile(path)
        except Exception:
            pass  # si l'ouverture échoue, on ne bloque pas

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
        self._set_progress(0)
        self._log("\n─── Démarrage ───")
        def worker():
            try:
                self._set_progress(10)
                self._log("Analyse de la structure...")
                structure=detect_structure(pdf)
                if not structure:
                    self.after(0,lambda: self._log("❌ Impossible de lire le PDF.")); return
                cols_info = list(structure.get('data_col_labels',{}).values()) if structure.get('style')=='standard' else ['nom/commentaire']
                self._log(f"Structure : {structure.get('style')} — {cols_info}")
                self._set_progress(30)
                self._log("Analyse du vocabulaire...")
                unknowns=detect_unknown_words(pdf,self.cfg.get('corrections',{}),self.cfg.get('blacklist',[]))
                known_already=set(self.cfg.get('corrections',{}).keys())|set(self.cfg.get('known_ok',[]))
                new_unknowns={w: locs for w, locs in unknowns.items() if w not in known_already}
                if new_unknowns:
                    self.after(0,lambda u=new_unknowns: self._ask_unknowns(u,pdf,out,structure)); return
                self._do_generate(pdf,out,structure)
            except Exception as e:
                self.after(0,lambda: self._log(f"❌ Erreur : {e}"))
                self.after(0,lambda: self._set_progress(0))
            finally:
                self.after(0,lambda: self.btn_run.config(state='normal'))
        threading.Thread(target=worker,daemon=True).start()

    def _ask_unknowns(self, unknowns, pdf, out, structure):
        win=tk.Toplevel(self); win.title("Mots inhabituels détectés")
        win.configure(bg=C_BG); win.resizable(True,True)
        win.geometry("680x500"); win.grab_set()

        tk.Label(win,text=f"  {len(unknowns)} mot(s) inhabituel(s) trouvé(s). Que faire ?",
                 font=('Helvetica',10,'bold'),bg=C_BG,fg=C_TEXT,pady=10).pack(anchor='w',padx=12)
        tk.Label(win,text="  ✓ Garder = conserver tel quel   |   → Corriger = remplacer par un autre texte   |   ✗ Supprimer = retirer du PDF",
                 font=('Helvetica',8),bg=C_BG,fg=C_TEXT2).pack(anchor='w',padx=12,pady=(0,8))

        frame=tk.Frame(win,bg=C_BG); frame.pack(fill='both',expand=True,padx=12)
        canvas=tk.Canvas(frame,bg=C_BG,highlightthickness=0)
        sb=tk.Scrollbar(frame,orient='vertical',command=canvas.yview,bg=C_PANEL)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right',fill='y'); canvas.pack(side='left',fill='both',expand=True)
        inner=tk.Frame(canvas,bg=C_BG); canvas.create_window((0,0),window=inner,anchor='nw')

        decisions={}
        for word in sorted(unknowns):
            locs = unknowns[word]  # set de N° d'équipement
            # Tri numérique des N° (évite '10' < '2' en tri string)
            def _n_sort_key(x):
                m = re.match(r'^(\d+)', x.strip())
                return (int(m.group(1)), x) if m else (99999, x)
            locs_sorted = sorted(locs, key=_n_sort_key)
            # Troncature à 6 N° pour éviter les labels démesurés
            MAX_LOCS = 6
            if len(locs_sorted) > MAX_LOCS:
                loc_text = f" (N° {', '.join(locs_sorted[:MAX_LOCS])}…)"
            else:
                loc_text = f" (N° {', '.join(locs_sorted)})" if locs_sorted else ""
            rf=tk.Frame(inner,bg=C_CARD,pady=8); rf.pack(fill='x',pady=2,padx=4)

            # Mot + localisation (pas de width fixe pour ne pas tronquer)
            tk.Label(rf,text=f"  «{word}»{loc_text}",font=('Courier',9,'bold'),
                     bg=C_CARD,fg=C_TEXT,anchor='w').pack(side='left')

            action=tk.StringVar(value='ok')
            correction=tk.StringVar(value='')

            # Champ correction (caché par défaut)
            corr_frame=tk.Frame(rf,bg=C_CARD)
            corr_entry=tk.Entry(corr_frame,textvariable=correction,width=18,
                bg=C_ENTRY_BG,fg=C_TEXT,insertbackground=C_TEXT,
                relief='flat',font=('Helvetica',9))
            corr_entry.pack(side='left',padx=4)
            tk.Label(corr_frame,text="(saisir le remplacement)",
                     bg=C_CARD,fg=C_TEXT2,font=('Helvetica',7)).pack(side='left')

            def on_action_change(var=action, cf=corr_frame):
                if var.get()=='correction':
                    cf.pack(side='left',padx=(0,4))
                else:
                    cf.pack_forget()

            for val,lbl in [('ok','✓ Garder'),('correction','→ Corriger'),('blacklist','✗ Supprimer')]:
                tk.Radiobutton(rf,text=lbl,variable=action,value=val,
                               bg=C_CARD,fg=C_TEXT,selectcolor=C_CARD,
                               activebackground=C_CARD,font=('Helvetica',8),
                               command=lambda v=action,cf=corr_frame: on_action_change(v,cf)
                               ).pack(side='left',padx=6)

            decisions[word]=(action,correction)

        inner.update_idletasks()
        canvas.config(scrollregion=canvas.bbox('all'))
        inner.bind('<Configure>', lambda e: canvas.config(scrollregion=canvas.bbox('all')))

        def confirm():
            for word,(av,cv) in decisions.items():
                act=av.get()
                if act=='correction':
                    val=cv.get().strip()
                    if val:
                        self.cfg.setdefault('corrections',{})[word]=val
                    else:
                        # Si champ vide → garder
                        ok=self.cfg.setdefault('known_ok',[])
                        if word not in ok: ok.append(word)
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

        tk.Button(win,text="✓ Confirmer et générer",command=confirm,
                  bg=C_SUCCESS,fg='white',font=('Helvetica',10,'bold'),
                  padx=20,pady=8,relief='flat',cursor='hand2').pack(pady=12)

    def _do_generate(self, pdf, out, structure):
        def worker():
            try:
                self._set_progress(50)
                generate_pdf(pdf,out,structure,
                    self.cfg.get('corrections',{}),
                    self.cfg.get('blacklist',[]),
                    log_fn=lambda m: self.after(0,lambda msg=m: self._log(msg)),
                    progress_fn=lambda v: self.after(0,lambda val=v: self._set_progress(val)))
                self.after(0,lambda: self._set_progress(100, done=True))
                self.after(0,lambda o=out: self._open_pdf(o))
            except PermissionError:
                def _ask_retry(p=pdf, o=out, s=structure):
                    if messagebox.askretrycancel(
                        "Fichier inaccessible",
                        f"Le fichier de sortie est déjà ouvert dans une autre application.\n\n"
                        f"Fermez-le puis cliquez sur Réessayer.\n\n{os.path.basename(o)}"
                    ):
                        self._do_generate(p, o, s)
                    else:
                        self._set_progress(0)
                self.after(0, _ask_retry)
            except Exception as e:
                self.after(0,lambda: self._log(f"❌ Erreur : {e}"))
                self.after(0,lambda: messagebox.showerror("Erreur",str(e)))
                self.after(0,lambda: self._set_progress(0))
            finally:
                self.after(0,lambda: self.btn_run.config(state='normal'))
        threading.Thread(target=worker,daemon=True).start()

    def _set_progress(self, value, done=False):
        """Met à jour la barre de progression (0-100). Si done=True, passe en vert puis reset."""
        if done:
            self.progress.configure(style='Green.Horizontal.TProgressbar')
            self.progress['value'] = 100
            # Reset après 2 secondes
            self.after(2000, lambda: (self.progress.configure(style='Red.Horizontal.TProgressbar'),
                                      self.progress.__setitem__('value', 0)))
        else:
            self.progress.configure(style='Red.Horizontal.TProgressbar')
            self.progress['value'] = value

if __name__ == '__main__':
    # Nettoyage de l'ancien exe _old laissé par une mise à jour précédente.
    # Lancé en thread avec retries car l'ancien exe peut encore être en cours de fermeture.
    try:
        _exe = os.path.abspath(sys.executable)
        _old = _exe.replace('.exe', '_old.exe')
        if os.path.exists(_old):
            def _cleanup_old(_p=_old):
                import time as _t
                for _ in range(10):
                    _t.sleep(1)
                    try:
                        if os.path.exists(_p):
                            os.remove(_p)
                            break
                    except: pass
            threading.Thread(target=_cleanup_old, daemon=True).start()
    except: pass
    app = App()
    if UPDATER_FLAG in sys.argv:
        # Petite bannière de notification mise à jour
        def _show_updated_banner():
            banner = tk.Toplevel(app)
            banner.overrideredirect(True)
            banner.configure(bg=C_SUCCESS)
            banner.attributes('-topmost', True)
            app.update_idletasks()
            bw, bh = 320, 44
            sx = app.winfo_x() + (app.winfo_width() - bw) // 2
            sy = app.winfo_y() + 60
            banner.geometry(f"{bw}x{bh}+{sx}+{sy}")
            tk.Label(banner, text=f"✓  Logiciel mis à jour vers {VERSION} !",
                     bg=C_SUCCESS, fg='white', font=('Helvetica', 10, 'bold')).pack(expand=True)
            banner.after(3500, banner.destroy)
        app.after(600, _show_updated_banner)
    app.mainloop()
