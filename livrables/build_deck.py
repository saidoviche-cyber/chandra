"""Construction du dossier de présentation Al Horra x Marjane (16:9)."""
import os
from PIL import Image, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
TMP = "/tmp/deck_assets"
os.makedirs(TMP, exist_ok=True)

# ---------------------------------------------------------------- palette
PURPLE_D = RGBColor(0x35, 0x12, 0x4C)
PURPLE = RGBColor(0x5B, 0x2A, 0x86)
PURPLE_L = RGBColor(0x8E, 0x6B, 0xB0)
GOLD = RGBColor(0xC7, 0x9A, 0x2E)
OLIVE = RGBColor(0x6E, 0x7F, 0x2A)
CREAM = RGBColor(0xFA, 0xF7, 0xEF)
CREAM2 = RGBColor(0xF2, 0xEC, 0xDF)
INK = RGBColor(0x22, 0x1E, 0x26)
GREY = RGBColor(0x6B, 0x66, 0x71)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SERIF = "Georgia"
SANS = "Arial"

W, H = 13.333, 7.5
ML = 0.95          # marge gauche
CW = 11.43         # largeur utile

prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)

FONT = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 40)
WARN = []


# ---------------------------------------------------------------- helpers
def measure(text, pt, width_in, spacing=1.15):
    """Hauteur estimée (pouces) d'un texte wrappé, métrique DejaVu (large)."""
    scale = pt / 40.0
    max_w = (width_in - 0.06) * 72.0 / scale
    lines = 0
    for para in text.split("\n"):
        if not para.strip():
            lines += 1
            continue
        cur = ""
        n = 1
        for word in para.split(" "):
            trial = (cur + " " + word).strip()
            if FONT.getlength(trial) <= max_w or not cur:
                cur = trial
            else:
                n += 1
                cur = word
        lines += n
    return lines * pt * spacing * 1.2 / 72.0


def tb(slide, x, y, w, h, *, anchor=MSO_ANCHOR.TOP, wrap=True):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    return box


def para(tf, text, *, size=13, font=SANS, color=INK, bold=False, italic=False,
         align=PP_ALIGN.LEFT, space_before=0, space_after=0, spacing=1.15,
         first=False, rtl=False):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.line_spacing = spacing
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    if rtl:
        pPr = p._pPr if p._pPr is not None else p._p.get_or_add_pPr()
        pPr.set("rtl", "1")
    run = p.add_run()
    run.text = text
    f = run.font
    f.name = font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    return p


def rect(slide, x, y, w, h, fill, line=None, lw=0.75, shape=MSO_SHAPE.RECTANGLE,
         adj=None, shadow=False):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(lw)
    s.shadow.inherit = shadow
    if adj is not None:
        s.adjustments[0] = adj
    s.text_frame.margin_left = s.text_frame.margin_right = 0
    s.text_frame.margin_top = s.text_frame.margin_bottom = 0
    return s


def fit_image(src, w_in, h_in, name, method="cover", focus=(0.5, 0.5)):
    """Recadre l'image au bon ratio sans déformation."""
    out = os.path.join(TMP, f"{name}.png")
    im = Image.open(src).convert("RGB")
    target = w_in / h_in
    iw, ih = im.size
    ratio = iw / ih
    if method == "cover":
        if ratio > target:
            nw = int(ih * target)
            cx = int((iw - nw) * focus[0])
            im = im.crop((cx, 0, cx + nw, ih))
        else:
            nh = int(iw / target)
            cy = int((ih - nh) * focus[1])
            im = im.crop((0, cy, iw, cy + nh))
    im = im.resize((int(w_in * 150), int(h_in * 150)), Image.LANCZOS)
    im.save(out, quality=92)
    return out


def pic(slide, src, x, y, w, h, name, **kw):
    p = fit_image(src, w, h, name, **kw)
    return slide.shapes.add_picture(p, Inches(x), Inches(y), Inches(w), Inches(h))


def base(color=WHITE):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = color
    return s


def header(slide, kicker, title, sub=None):
    t = tb(slide, ML, 0.62, 8.0, 0.25)
    para(t.text_frame, kicker.upper(), size=9.5, bold=True, color=GOLD, first=True)
    th = 0.62 if len(title) < 46 else 1.0
    t2 = tb(slide, ML, 0.92, CW, th)
    para(t2.text_frame, title, size=29, font=SERIF, bold=True, color=PURPLE_D,
         first=True, spacing=1.05)
    if measure(title, 29, CW, 1.05) > th:
        WARN.append(f"TITRE trop long: {title}")
    y = 0.92 + max(th, measure(title, 29, CW, 1.05)) + 0.12
    rect(slide, ML + 0.02, y, 0.85, 0.05, GOLD)
    if sub:
        ts = tb(slide, ML, y + 0.2, CW * 0.86, 0.6)
        para(ts.text_frame, sub, size=12.5, color=GREY, first=True, spacing=1.25)
        if measure(sub, 12.5, CW * 0.86) > 0.6:
            WARN.append(f"SOUS-TITRE trop long: {sub[:50]}")
        return y + 0.2 + measure(sub, 12.5, CW * 0.86) + 0.28
    return y + 0.35


def footer(slide, n, dark=False):
    col = RGBColor(0xD9, 0xCF, 0xE6) if dark else RGBColor(0xDD, 0xD8, 0xE0)
    ln = rect(slide, ML, 6.95, CW, 0.008, col)
    f = tb(slide, ML, 7.05, 9.0, 0.25)
    para(f.text_frame, "Al Horra  ×  Marjane   ·   Activation 2026",
         size=8.5, color=RGBColor(0xB9, 0xB2, 0xC0) if dark else GREY, first=True)
    page = tb(slide, W - ML - 0.6, 7.05, 0.6, 0.25)
    para(page.text_frame, f"{n:02d}", size=8.5, bold=True,
         color=RGBColor(0xB9, 0xB2, 0xC0) if dark else PURPLE,
         align=PP_ALIGN.RIGHT, first=True)


def card(slide, x, y, w, h, *, fill=CREAM, line=CREAM2, accent=None):
    c = rect(slide, x, y, w, h, fill, line, 0.75, MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.06)
    if accent:
        rect(slide, x + 0.30, y - 0.02, 0.62, 0.05, accent)
    return c


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ================================================================ 01 cover
s = base()
pic(s, os.path.join(ASSETS, "olives.png"), 0, 0, W, 2.55, "cover_band", focus=(0.5, 0.55))
rect(s, 0, 2.55, W, 0.06, GOLD)

t = tb(s, ML, 3.05, 10.5, 0.3)
para(t.text_frame, "DOSSIER D'ACTIVATION  ·  CAMPAGNE 2026", size=10, bold=True,
     color=GOLD, first=True)

t = tb(s, ML, 3.42, 11.0, 1.4)
para(t.text_frame, "Al Horra × Marjane", size=52, font=SERIF, bold=True,
     color=PURPLE_D, first=True, spacing=0.95)
if measure("Al Horra × Marjane", 52, 11.0, 0.95) > 1.4:
    WARN.append("titre cover")

rect(s, ML + 0.02, 4.92, 1.1, 0.055, GOLD)

t = tb(s, ML, 5.18, 8.9, 1.0)
para(t.text_frame, "Activation dégustation & immersion digitale", size=20,
     color=PURPLE, bold=True, first=True, spacing=1.1)
para(t.text_frame, "Huile d'olive 100 % marocaine, naturelle et pure — Top 15 hypermarchés.",
     size=13, color=GREY, space_before=8, spacing=1.2)

t = tb(s, ML, 6.72, 9.5, 0.3)
para(t.text_frame, "Access Développement Service  ·  Département Logistique  ·  Novembre 2026",
     size=9.5, color=GREY, first=True)
notes(s, "Dossier de présentation de l'activation Al Horra en hypermarchés Marjane : "
         "contexte, objectifs, dispositif, déploiement et suivi.")

# ================================================================ 02 en bref
s = base(CREAM)
y = header(s, "Synthèse", "L'activation en un regard",
           "Faire passer le consommateur du vrac à la marque par l'expérience : goûter, voir, acheter.")

kpis = [("15", "hypermarchés Marjane", "Top 15 du réseau, déploiement national"),
        ("505", "dégustations + ventes", "Objectif cumulé, par jour, sur 15 magasins"),
        ("10", "magasins en immersion", "Casque de réalité virtuelle sur stand"),
        ("80 %", "du marché en vrac", "De la consommation d'huile d'olive au Maroc")]
cw, gap = 2.66, 0.19
for i, (num, lab, sub) in enumerate(kpis):
    x = ML + i * (cw + gap)
    card(s, x, y + 0.05, cw, 2.15, fill=WHITE, line=CREAM2, accent=GOLD)
    t = tb(s, x + 0.32, y + 0.52, cw - 0.6, 0.75)
    para(t.text_frame, num, size=38, font=SERIF, bold=True, color=PURPLE, first=True)
    lab_y = y + 1.26
    t2 = tb(s, x + 0.32, lab_y, cw - 0.6, 0.4)
    para(t2.text_frame, lab, size=11, bold=True, color=INK, first=True, spacing=1.1)
    lh = measure(lab, 11, cw - 0.6)
    if lh > 0.4:
        WARN.append(f"kpi label {i}")
    t3 = tb(s, x + 0.32, lab_y + lh + 0.06, cw - 0.6, 0.5)
    para(t3.text_frame, sub, size=9.5, color=GREY, first=True, spacing=1.15)
    if measure(sub, 9.5, cw - 0.6) > 0.5:
        WARN.append("kpi sub")

band = rect(s, ML, y + 2.55, CW, 1.15, PURPLE_D, None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.08)
t = tb(s, ML + 0.45, y + 2.75, CW - 0.9, 0.8)
para(t.text_frame, "Le principe", size=9.5, bold=True, color=GOLD, first=True)
para(t.text_frame, "La dégustation fait juger la qualité par le consommateur lui-même ; "
                   "le casque immersif montre le processus de fabrication. "
                   "Ensemble, ils transforment un essai en confiance, puis en achat.",
     size=13, color=WHITE, space_before=6, spacing=1.25)
if measure("La dégustation fait juger la qualité par le consommateur lui-même ; le casque immersif montre le processus de fabrication. Ensemble, ils transforment un essai en confiance, puis en achat.", 13, CW - 0.9) > 0.8:
    WARN.append("principe band")
footer(s, 2)

# ================================================================ 03 contexte
s = base()
y = header(s, "Contexte", "Un marché dominé par le vrac")

lw = 6.35
items = [("80 %", "de la consommation marocaine d'huile d'olive se fait encore en vrac."),
         ("Zéro repère", "Pas de marque, pas de garantie de qualité : le consommateur achète à la mesure."),
         ("Un produit qui gagne à être goûté", "Al Horra : une huile 100 % marocaine, naturelle et pure.")]
yy = y + 0.02
for num, txt in items:
    t = tb(s, ML, yy, lw, 0.3)
    para(t.text_frame, num, size=12.5, bold=True, color=PURPLE, first=True)
    t2 = tb(s, ML, yy + 0.29, lw, 0.7)
    para(t2.text_frame, txt, size=12.5, color=INK, first=True, spacing=1.25)
    hh = measure(txt, 12.5, lw)
    rect(s, ML, yy + 0.32 + hh, lw - 1.2, 0.012, CREAM2)
    yy += 0.31 + hh + 0.26

card(s, ML, yy + 0.02, lw, 1.05, fill=CREAM, line=CREAM2, accent=GOLD)
t = tb(s, ML + 0.34, yy + 0.24, lw - 0.7, 0.8)
para(t.text_frame, "La dégustation permet au consommateur de juger par lui-même "
                   "la qualité gustative de l'huile Al Horra.",
     size=11.5, italic=True, color=PURPLE_D, first=True, spacing=1.25)
if measure("La dégustation permet au consommateur de juger par lui-même la qualité gustative de l'huile Al Horra.", 11.5, lw - 0.7) > 0.8:
    WARN.append("contexte card")

pic(s, os.path.join(ASSETS, "photo2.jpg"), ML + lw + 0.5, y + 0.02, 4.55, 4.0,
    "ctx_shelf", focus=(0.5, 0.45))
footer(s, 3)

# ================================================================ 04 enjeu
s = base(CREAM)
y = header(s, "Enjeu", "Trois verrous à lever pour recruter")

blocks = [("01", "La confiance", "Le vrac est perçu comme plus authentique. "
                                 "La qualité doit se prouver par le goût, pas par le discours."),
          ("02", "La preuve", "Montrer le processus de fabrication pour crédibiliser "
                              "le 100 % marocain, naturel et pur."),
          ("03", "L'essai", "Créer le contact produit en magasin, au moment exact "
                            "où le consommateur décide.")]
cw = 3.62
for i, (num, title, txt) in enumerate(blocks):
    x = ML + i * (cw + 0.28)
    card(s, x, y + 0.05, cw, 2.05, fill=WHITE, line=CREAM2)
    t = tb(s, x + 0.34, y + 0.3, cw - 0.68, 0.4)
    para(t.text_frame, num, size=11, bold=True, color=GOLD, first=True)
    t = tb(s, x + 0.34, y + 0.62, cw - 0.68, 0.4)
    para(t.text_frame, title, size=16, font=SERIF, bold=True, color=PURPLE_D, first=True)
    t = tb(s, x + 0.34, y + 1.05, cw - 0.68, 0.9)
    para(t.text_frame, txt, size=11.5, color=GREY, first=True, spacing=1.3)
    if measure(txt, 11.5, cw - 0.68) > 0.9:
        WARN.append(f"enjeu card {i}")

rect(s, ML, y + 2.5, CW, 0.9, None, None)
t = tb(s, ML, y + 2.55, CW, 0.3)
para(t.text_frame, "NOTRE RÉPONSE", size=9.5, bold=True, color=GOLD, first=True)
t = tb(s, ML, y + 2.85, CW, 0.6)
para(t.text_frame, "Une activation qui associe dégustation et vidéo visionnée "
                   "à travers un casque immersif.", size=19, font=SERIF,
     color=PURPLE, first=True, spacing=1.15)
if measure("Une activation qui associe dégustation et vidéo visionnée à travers un casque immersif.", 19, CW, 1.15) > 0.6:
    WARN.append("reponse")

pill = [(0, "Dégustation", "Faire goûter, comparer, apprécier"),
        (1, "Immersion", "Montrer la récolte et la fabrication"),
        (2, "Conversion", "Orienter vers le format adapté au foyer")]
yy = y + 3.85
for i, (_, a, b) in enumerate(pill):
    x = ML + i * 3.9
    rect(s, x, yy, 0.34, 0.05, OLIVE if i == 0 else (PURPLE_L if i == 1 else GOLD))
    t = tb(s, x, yy + 0.16, 3.4, 0.3)
    para(t.text_frame, a, size=11.5, bold=True, color=INK, first=True)
    t = tb(s, x, yy + 0.46, 3.4, 0.3)
    para(t.text_frame, b, size=10.5, color=GREY, first=True)
footer(s, 4)

# ================================================================ 05 objectifs qualitatifs
s = base()
y = header(s, "Objectifs", "Objectifs qualitatifs",
           "Trois intentions qui structurent toute l'activation en magasin.")

objs = [("Renforcer", "le positionnement d'Al Horra comme une huile d'olive "
                      "100 % marocaine, naturelle et pure."),
        ("Valoriser", "notre savoir-faire et notre processus de fabrication "
                      "pour renforcer la confiance des consommateurs."),
        ("Favoriser", "l'essai et le recrutement en faisant découvrir "
                      "la qualité gustative d'Al Horra à travers la dégustation.")]
cw = 3.62
for i, (verb, txt) in enumerate(objs):
    x = ML + i * (cw + 0.28)
    card(s, x, y + 0.05, cw, 2.55, fill=CREAM, line=CREAM2)
    t = tb(s, x + 0.34, y + 0.32, cw - 0.68, 0.5)
    para(t.text_frame, f"0{i+1}", size=10.5, bold=True, color=GOLD, first=True)
    t = tb(s, x + 0.34, y + 0.62, cw - 0.68, 0.45)
    para(t.text_frame, verb, size=19, font=SERIF, bold=True, color=PURPLE_D, first=True)
    t = tb(s, x + 0.34, y + 1.2, cw - 0.68, 1.2)
    para(t.text_frame, txt, size=11.5, color=GREY, first=True, spacing=1.3)
    if measure(txt, 11.5, cw - 0.68) > 1.2:
        WARN.append(f"obj q {i}")

t = tb(s, ML, y + 2.95, CW, 0.35)
para(t.text_frame, "Résultat attendu : un consommateur qui a goûté, compris et reconnu "
                   "la qualité — et qui rachète.", size=14, italic=True, color=PURPLE, first=True)
if measure("Résultat attendu : un consommateur qui a goûté, compris et reconnu la qualité — et qui rachète.", 14, CW) > 0.35:
    WARN.append("res q")
footer(s, 5)

# ================================================================ 06 objectifs quantitatifs
s = base(CREAM)
y = header(s, "Objectifs", "Objectifs par magasin",
           "Objectif = dégustations + ventes réalisées par jour, par point de vente.")

stores = ["Marjane Ain Sbaa", "Marjane El Jadida", "Marjane Salé", "Marjane Bouregreg",
          "Marjane Fès 2", "Marjane Menara", "Marjane Massira", "Marjane Agadir Founty",
          "Marjane Kenitra", "Marjane Tanger Medina", "Marjane Mohammedia",
          "Marjane Derb Sultan", "Marjane Hay Hassani", "Marjane Sidi Othmane",
          "Marjane Oujda"]
quota = [40, 25, 40, 30, 30, 40, 40, 40, 50, 25, 25, 30, 40, 25, 25]
total = sum(quota)

rows, cols = len(stores) + 2, 3
tw = 7.5
tbl_shape = s.shapes.add_table(rows, cols, Inches(ML), Inches(y + 0.02), Inches(tw),
                               Inches(0.33 * rows))
table = tbl_shape.table
table.columns[0].width = Inches(0.55)
table.columns[1].width = Inches(4.85)
table.columns[2].width = Inches(2.1)
for r in range(rows):
    table.rows[r].height = Inches(0.255)

data = [("#", "Hypermarché Marjane", "Objectif / jour")] + \
       [(str(i + 1), st, str(q)) for i, (st, q) in enumerate(zip(stores, quota))] + \
       [("", "TOTAL — 15 magasins", str(total))]
for r, row in enumerate(data):
    for c, val in enumerate(row):
        cell = table.cell(r, c)
        cell.margin_left = Inches(0.08)
        cell.margin_right = Inches(0.08)
        cell.margin_top = Inches(0.02)
        cell.margin_bottom = Inches(0.02)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        if r == 0:
            cell.fill.fore_color.rgb = PURPLE_D
        elif r == rows - 1:
            cell.fill.fore_color.rgb = GOLD
        else:
            cell.fill.fore_color.rgb = WHITE if r % 2 else CREAM2
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER if c != 1 else PP_ALIGN.LEFT
        run = p.add_run()
        run.text = val
        f = run.font
        f.name = SANS
        f.size = Pt(9.5)
        f.bold = (r == 0 or r == rows - 1)
        f.color.rgb = WHITE if r == 0 else INK

# panneau de lecture
x = ML + tw + 0.45
pw = W - ML - x
card(s, x, y + 0.02, pw, 1.45, fill=PURPLE_D, line=None)
t = tb(s, x + 0.3, y + 0.16, pw - 0.6, 0.25)
para(t.text_frame, "Objectif cumulé", size=9, bold=True, color=GOLD, first=True)
t = tb(s, x + 0.3, y + 0.4, pw - 0.6, 0.6)
para(t.text_frame, f"{total}", size=32, font=SERIF, bold=True, color=WHITE, first=True)
t = tb(s, x + 0.3, y + 1.06, pw - 0.6, 0.25)
para(t.text_frame, "contacts + ventes par jour", size=9.5, color=RGBColor(0xD9, 0xCF, 0xE6), first=True)

lines = [("≈ 34", "contacts par magasin et par jour"),
         ("50", "ambition la plus haute — Kenitra"),
         ("25", "plancher — 5 magasins")]
yy = y + 1.68
for num, lab in lines:
    t = tb(s, x + 0.32, yy, pw - 0.6, 0.3)
    para(t.text_frame, num, size=15, font="Georgia", bold=True, color=PURPLE, first=True)
    t = tb(s, x + 1.05, yy + 0.04, pw - 1.4, 0.5)
    para(t.text_frame, lab, size=9.5, color=GREY, first=True, spacing=1.15)
    if measure(lab, 9.5, pw - 1.4) > 0.5:
        WARN.append("kpi line")
    yy += 0.72

footer(s, 6)

# ================================================================ 07 identité visuelle
s = base()
y = header(s, "Offre", "Nouvelle identité visuelle Al Horra",
           "Une gamme lisible en linéaire, du format d'essai au format famille.")
ih = 3.91
pic(s, os.path.join(ASSETS, "gammes.png"), ML, y + 0.02, 8.1, ih, "gammes")
t = tb(s, ML, y + ih + 0.1, 8.1, 0.3)
para(t.text_frame, "La gamme Al Horra 2026 déclinée en quatre formats.",
     size=9, italic=True, color=GREY, first=True)

x = ML + 8.35
pw = W - ML - x
facts = [("0,5 L", "format d'entrée, essai et découverte"),
         ("1 L", "cœur de gamme, achat mensuel"),
         ("2 L", "format familial intermédiaire"),
         ("5 L", "format famille, achat de fond")]
yy = y + 0.05
for fmt, desc in facts:
    rect(s, x, yy + 0.08, 0.05, 0.42, OLIVE)
    t = tb(s, x + 0.2, yy, pw - 0.2, 0.3)
    para(t.text_frame, fmt, size=13, bold=True, color=PURPLE_D, first=True)
    t = tb(s, x + 0.2, yy + 0.26, pw - 0.2, 0.3)
    para(t.text_frame, desc, size=9.5, color=GREY, first=True)
    yy += 0.72

card(s, x, yy + 0.12, pw, 1.15, fill=CREAM, line=CREAM2)
t = tb(s, x + 0.25, yy + 0.3, pw - 0.5, 0.85)
para(t.text_frame, "Le violet et le vert olive portent la signature « 100 % marocaine » "
                   "sur l'ensemble des formats.", size=9.5, color=INK, first=True, spacing=1.25)
if measure("Le violet et le vert olive portent la signature « 100 % marocaine » sur l'ensemble des formats.", 9.5, pw - 0.5) > 0.85:
    WARN.append("identite note")
footer(s, 7)

# ================================================================ 08 film
s = base(CREAM)
y = header(s, "Contenu", "Le film 2026, au cœur de l'expérience")
pic(s, os.path.join(ASSETS, "olives.png"), ML, y + 0.02, CW, 1.55, "film_band", focus=(0.5, 0.5))

lw = 5.6
t = tb(s, ML, y + 1.78, lw, 0.9)
para(t.text_frame, "Nouveau spot Al Horra 2026", size=15, font=SERIF, bold=True,
     color=PURPLE_D, first=True)
para(t.text_frame, "La récolte, la transformation et la pureté de l'huile : "
                   "un contenu qui montre ce que le vrac ne peut pas montrer.",
     size=11.5, color=GREY, space_before=6, spacing=1.3)
if measure("La récolte, la transformation et la pureté de l'huile : un contenu qui montre ce que le vrac ne peut pas montrer.", 11.5, lw) > 0.95:
    WARN.append("film texte")

# carte script arabe
cx = ML + lw + 0.5
cw = W - ML - cx
card(s, cx, y + 1.72, cw, 2.3, fill=WHITE, line=CREAM2, accent=GOLD)
t = tb(s, cx + 0.3, y + 1.86, cw - 0.6, 0.25)
para(t.text_frame, "SCRIPT DU FILM — VERSION DARIJA", size=8.5, bold=True, color=GOLD, first=True)
t = tb(s, cx + 0.3, y + 2.14, cw - 0.6, 1.1)
para(t.text_frame, "السلام", size=13, color=PURPLE, first=True, rtl=True, align=PP_ALIGN.RIGHT)
para(t.text_frame, "أجي نعرّفك على زيت الزيتون الحرّة، زيت طبيعية 100%، مغربية 100% وحرّة 100%",
     size=14, color=PURPLE_D, rtl=True,
     align=PP_ALIGN.RIGHT, space_before=8, spacing=1.5)
t = tb(s, cx + 0.3, y + 3.28, cw - 0.6, 0.6)
para(t.text_frame, "« Venez découvrir l'huile d'olive Al Horra : 100 % naturelle, "
                   "100 % marocaine, 100 % pure. »", size=9.5, italic=True,
     color=GREY, first=True, spacing=1.25)
if measure("« Venez découvrir l'huile d'olive Al Horra : 100 % naturelle, 100 % marocaine, 100 % pure. »", 9.5, cw - 0.6) > 0.6:
    WARN.append("script fr")
# trois preuves
for i, (a, b) in enumerate([("Récolte", "des olives, à maturité"),
                            ("Transformation", "sans additif, sans mélange"),
                            ("Pureté", "100 % marocaine, 100 % naturelle")]):
    x = ML + i * 1.95
    rect(s, x, y + 3.62, 0.05, 0.42, OLIVE)
    t = tb(s, x + 0.18, y + 3.6, 1.7, 0.3)
    para(t.text_frame, a, size=10, bold=True, color=PURPLE_D, first=True)
    t = tb(s, x + 0.18, y + 3.86, 1.75, 0.5)
    para(t.text_frame, b, size=8.5, color=GREY, first=True, spacing=1.15)
footer(s, 8)

# ================================================================ 09 dispositif
s = base()
y = header(s, "Dispositif", "Deux mécaniques, une seule expérience",
           "Le stand de dégustation installe le contact produit ; le casque immersif apporte la preuve.")

cw = 4.55
card(s, ML, y + 0.05, cw, 3.0, fill=CREAM, line=CREAM2, accent=OLIVE)
t = tb(s, ML + 0.34, y + 0.55, cw - 0.68, 0.35)
para(t.text_frame, "1", size=11, bold=True, color=OLIVE, first=True)
t = tb(s, ML + 0.34, y + 0.85, cw - 0.68, 0.45)
para(t.text_frame, "Stand de dégustation", size=16, font=SERIF, bold=True, color=PURPLE_D, first=True)
for i, b in enumerate(["Pain et huile servis par une animatrice formée",
                       "Argumentaire centré sur le 100 % naturel et pur",
                       "Mise en avant des formats 1 L et 5 L",
                       "Support du dépliant anniversaire"]):
    t = tb(s, ML + 0.34, y + 1.32 + i * 0.4, cw - 0.68, 0.4)
    para(t.text_frame, "—  " + b, size=10.5, color=INK, first=True, spacing=1.2)
    if measure(b, 10.5, cw - 0.9) > 0.45:
        WARN.append(f"dispo A {i}")

# plus
cx = ML + cw + 0.42
t = tb(s, cx, y + 1.6, 0.6, 0.6)
para(t.text_frame, "+", size=34, font=SERIF, bold=True, color=GOLD, first=True,
     align=PP_ALIGN.CENTER)

x2 = cx + 0.75
card(s, x2, y + 0.05, cw, 3.0, fill=CREAM, line=CREAM2, accent=PURPLE_L)
t = tb(s, x2 + 0.34, y + 0.55, cw - 0.68, 0.35)
para(t.text_frame, "2", size=11, bold=True, color=PURPLE_L, first=True)
t = tb(s, x2 + 0.34, y + 0.85, cw - 0.68, 0.45)
para(t.text_frame, "Immersion digitale", size=16, font=SERIF, bold=True, color=PURPLE_D, first=True)
for i, b in enumerate(["Casque de réalité virtuelle sur stand",
                       "Film du processus de fabrication",
                       "10 magasins équipés en priorité",
                       "Un temps d'immersion de quelques minutes"]):
    t = tb(s, x2 + 0.34, y + 1.32 + i * 0.4, cw - 0.68, 0.4)
    para(t.text_frame, "—  " + b, size=10.5, color=INK, first=True, spacing=1.2)
    if measure(b, 10.5, cw - 0.9) > 0.45:
        WARN.append(f"dispo B {i}")

pic(s, os.path.join(ASSETS, "vr.png"), 10.35, y + 3.18, 1.95, 1.18, "vr", focus=(0.5, 0.5))
t = tb(s, ML, y + 3.35, 7.4, 0.9)
para(t.text_frame, "Le casque immersif, installé sur le stand de dégustation.",
     size=12, italic=True, color=PURPLE, first=True)
para(t.text_frame, "Dix magasins en priorité, sur les quinze du TOP 15.",
     size=9.5, color=GREY, space_before=4)
footer(s, 9)

# ================================================================ 10 parcours
s = base(CREAM)
y = header(s, "Mécanique", "Le parcours consommateur en cinq temps")

steps = [("Attirer", "Stand visible, signalétique 100 % marocaine", "contacts stand"),
         ("Déguster", "Pain + huile, comparer avec le vrac", "dégustations"),
         ("S'immerger", "Casque : récolte et fabrication", "temps d'immersion"),
         ("Convertir", "Dépliant anniversaire, format adapté", "unités vendues"),
         ("Fidéliser", "Réachat, relevé des ventes", "réachat constaté")]
n = len(steps)
cw = (CW - 0.15 * (n - 1)) / n
for i, (title, txt, micro) in enumerate(steps):
    x = ML + i * (cw + 0.15)
    box = rect(s, x, y + 0.1, cw, 2.45, WHITE, CREAM2, 0.75,
               MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.07)
    rect(s, x + 0.3, y + 0.38, 0.45, 0.45, PURPLE if i < 4 else GOLD,
         shape=MSO_SHAPE.OVAL)
    t = tb(s, x + 0.3, y + 0.46, 0.45, 0.3)
    para(t.text_frame, str(i + 1), size=12, bold=True, color=WHITE, first=True,
         align=PP_ALIGN.CENTER)
    t = tb(s, x + 0.3, y + 1.02, cw - 0.6, 0.4)
    para(t.text_frame, title, size=14, font=SERIF, bold=True, color=PURPLE_D, first=True)
    t = tb(s, x + 0.3, y + 1.45, cw - 0.6, 0.85)
    para(t.text_frame, txt, size=10, color=GREY, first=True, spacing=1.25)
    if measure(txt, 10, cw - 0.6) > 0.85:
        WARN.append(f"parcours {i}")
    rect(s, x + 0.3, y + 2.12, cw - 0.6, 0.012, CREAM2)
    t = tb(s, x + 0.3, y + 2.2, cw - 0.6, 0.3)
    para(t.text_frame, micro.upper(), size=7.5, bold=True, color=GOLD, first=True)
    if i < n - 1:
        ar = rect(s, x + cw + 0.02, y + 1.25, 0.11, 0.16, GOLD, shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
        ar.rotation = 90

card(s, ML, y + 2.88, CW, 1.25, fill=CREAM2, line=CREAM2, accent=GOLD)
t = tb(s, ML + 0.42, y + 3.12, CW - 0.84, 0.9)
para(t.text_frame, "Chaque étape est mesurée : nombre de contacts, temps d'immersion, "
                   "unités vendues et réachat constaté.", size=12.5, color=PURPLE_D, first=True,
     spacing=1.2)
para(t.text_frame, "Support : la fiche de relevé quotidienne remise à chaque animatrice avec le dépliant.",
     size=10, color=GREY, space_before=6)
if measure("Chaque étape est mesurée : nombre de contacts, temps d'immersion, unités vendues et réachat constaté.", 12.5, CW - 0.84) > 0.6:
    WARN.append("parcours note")
footer(s, 10)

# ================================================================ 11 top 15
s = base()
y = header(s, "Déploiement", "Le TOP 15 des hypermarchés Marjane",
           "Priorité au potentiel : les 15 magasins les plus contributeurs du réseau.")
cw, ch, gap = 2.16, 0.94, 0.155
for i, st in enumerate(stores):
    col, row = i % 5, i // 5
    x = ML + col * (cw + gap)
    yy = y + 0.05 + row * (ch + gap)
    rect(s, x, yy, cw, ch, CREAM if row % 2 == 0 else WHITE, CREAM2, 0.75,
         MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.08)
    t = tb(s, x + 0.18, yy + 0.13, 0.5, 0.25)
    para(t.text_frame, f"{i+1:02d}", size=9.5, bold=True, color=GOLD, first=True)
    name = st.replace("Marjane ", "")
    t = tb(s, x + 0.18, yy + 0.36, cw - 0.36, 0.45)
    para(t.text_frame, name, size=10.5, bold=True, color=PURPLE_D, first=True, spacing=1.15)
    t = tb(s, x + 0.18, yy + 0.64, cw - 0.36, 0.25)
    para(t.text_frame, f"objectif {quota[i]} / jour", size=8.5, color=GREY, first=True)
    if measure(name, 10.5, cw - 0.36) > 0.45:
        WARN.append(f"store {name}")

t = tb(s, ML, y + 3.55, CW, 0.3)
para(t.text_frame, "Hypermarchés Marjane · classement interne du réseau",
     size=8.5, italic=True, color=GREY, first=True)
footer(s, 11)

# ================================================================ 12 top 10 immersion
s = base(CREAM)
y = header(s, "Focus", "Les 10 magasins en immersion digitale",
           "Le casque est déployé là où le trafic et le potentiel de recrutement sont les plus forts.")
top10 = ["Marjane Hay Hassani", "Marjane Marrakech Menara", "Marjane Ain Sebaa",
         "Marjane Tanger", "Marjane Agadir Founty", "Marjane Derb Sultan",
         "Marjane Mohammedia", "Marjane Salé", "Marjane Bouregreg",
         "Marjane Marrakech Massira"]
for i, st in enumerate(top10):
    col, row = i // 5, i % 5
    x = ML + col * 5.75
    yy = y + 0.05 + row * 0.79
    rect(s, x, yy, 5.45, 0.63, WHITE, CREAM2, 0.75, MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.12)
    rect(s, x + 0.22, yy + 0.16, 0.36, 0.36, PURPLE_L, shape=MSO_SHAPE.OVAL)
    t = tb(s, x + 0.22, yy + 0.22, 0.36, 0.25)
    para(t.text_frame, str(i + 1), size=10, bold=True, color=WHITE, first=True,
         align=PP_ALIGN.CENTER)
    t = tb(s, x + 0.72, yy + 0.2, 4.5, 0.3)
    para(t.text_frame, st.replace("Marjane ", "Marjane "), size=11.5, bold=True,
         color=PURPLE_D, first=True)

x = ML + 11.5
t = tb(s, ML, y + 4.05, CW, 0.3)
para(t.text_frame, "10 magasins sur les 15 du TOP : l'immersion concentre l'effort là où l'essai se transforme le mieux.",
     size=11, italic=True, color=PURPLE, first=True)
if measure("10 magasins sur les 15 du TOP : l'immersion concentre l'effort là où l'essai se transforme le mieux.", 11, CW) > 0.3:
    WARN.append("top10 note")
footer(s, 12)

# ================================================================ 13 dépliants anniversaire
s = base()
y = header(s, "Opération anniversaire", "Objectifs du dépliant anniversaire",
           "Le dépliant est le support de conversion : il fixe l'objectif de chaque magasin.")
rows, cols = len(stores) + 2, 2
tw = 7.5
shp = s.shapes.add_table(rows, cols, Inches(ML), Inches(y + 0.02), Inches(tw), Inches(0.33 * rows))
table = shp.table
table.columns[0].width = Inches(5.4)
table.columns[1].width = Inches(2.1)
for r in range(rows):
    table.rows[r].height = Inches(0.255)
data = [("Hypermarché", "Obj. dégus + vente / jour")] + \
       [(st, str(q)) for st, q in zip(stores, quota)] + \
       [("TOTAL — 15 magasins", str(total))]
for r, row in enumerate(data):
    for c, val in enumerate(row):
        cell = table.cell(r, c)
        cell.margin_left = Inches(0.1)
        cell.margin_right = Inches(0.1)
        cell.margin_top = Inches(0.02)
        cell.margin_bottom = Inches(0.02)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        cell.fill.fore_color.rgb = PURPLE_D if r == 0 else (GOLD if r == rows - 1 else (WHITE if r % 2 else CREAM2))
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
        run = p.add_run()
        run.text = val
        run.font.name = SANS
        run.font.size = Pt(9.5)
        run.font.bold = r == 0 or r == rows - 1
        run.font.color.rgb = WHITE if r == 0 else INK

x = ML + tw + 0.45
pw = W - ML - x
card(s, x, y + 0.02, pw, 1.55, fill=CREAM, line=CREAM2, accent=GOLD)
t = tb(s, x + 0.3, y + 0.28, pw - 0.6, 0.6)
para(t.text_frame, "505", size=40, font=SERIF, bold=True, color=PURPLE, first=True)
t = tb(s, x + 0.3, y + 0.92, pw - 0.6, 0.5)
para(t.text_frame, "contacts + ventes visés par jour, tous magasins", size=10, color=GREY, first=True, spacing=1.2)
if measure("contacts + ventes visés par jour, tous magasins", 10, pw - 0.6) > 0.5:
    WARN.append("dep total")

tips = ["Le dépliant rappelle l'offre anniversaire au moment du passage en caisse.",
        "Chaque animatrice reporte son réalisé par rapport à son objectif quotidien.",
        "Écart objectif / réalisé analysé à mi-parcours pour réallouer l'effort."]
yy = y + 1.78
for i, txt in enumerate(tips):
    rect(s, x + 0.3, yy + 0.06, 0.05, 0.3, PURPLE_L)
    t = tb(s, x + 0.48, yy, pw - 0.78, 0.65)
    para(t.text_frame, txt, size=9.5, color=INK, first=True, spacing=1.25)
    hh = measure(txt, 9.5, pw - 0.78)
    if hh > 0.65:
        WARN.append(f"tip {i}")
    yy += hh + 0.3
footer(s, 13)

# ================================================================ 14 script
s = base(CREAM)
y = header(s, "Terrain", "Script animatrice",
           "Un script court, en darija, qui déclenche la dégustation puis l'immersion.")

card(s, ML, y + 0.05, 6.1, 2.55, fill=WHITE, line=CREAM2, accent=GOLD)
t = tb(s, ML + 0.34, y + 0.3, 5.4, 0.25)
para(t.text_frame, "VERSION DARIJA", size=8.5, bold=True, color=GOLD, first=True)
t = tb(s, ML + 0.34, y + 0.62, 5.4, 1.85)
para(t.text_frame, "السلام", size=15, color=PURPLE, first=True, rtl=True, align=PP_ALIGN.RIGHT)
para(t.text_frame, "أجي نعرّفك على زيت الزيتون الحرّة، زيت طبيعية 100%\nمغربية 100% وحرّة 100%",
     size=16, color=PURPLE_D, rtl=True, align=PP_ALIGN.RIGHT, space_before=10, spacing=1.6)

x2 = ML + 6.45
cw = W - ML - x2
card(s, x2, y + 0.05, cw, 2.55, fill=WHITE, line=CREAM2, accent=PURPLE_L)
t = tb(s, x2 + 0.34, y + 0.3, cw - 0.68, 0.25)
para(t.text_frame, "TRADUCTION ET INTENTION", size=8.5, bold=True, color=GOLD, first=True)
t = tb(s, x2 + 0.34, y + 0.62, cw - 0.68, 0.8)
para(t.text_frame, "« Venez découvrir l'huile d'olive Al Horra : 100 % naturelle, "
                   "100 % marocaine, 100 % pure. »", size=13, italic=True,
     color=PURPLE_D, first=True, spacing=1.3)
t = tb(s, x2 + 0.34, y + 1.55, cw - 0.68, 0.85)
para(t.text_frame, "Objectif du script : engager la conversation, annoncer les trois preuves "
                   "(naturelle, marocaine, pure) et déclencher la dégustation.", size=10.5,
     color=GREY, first=True, spacing=1.3)
if measure("Objectif du script : engager la conversation, annoncer les trois preuves (naturelle, marocaine, pure) et déclencher la dégustation.", 10.5, cw - 0.68) > 0.85:
    WARN.append("script intention")

# déroulé 3 temps
steps = [("1. Accueillir", "Saluer, proposer le pain et l'huile Al Horra."),
         ("2. Prouver", "Énoncer les trois preuves, laisser goûter sans commenter."),
         ("3. Prolonger", "Proposer le casque immersif, puis le format adapté au foyer.")]
cw = 3.62
for i, (a, b) in enumerate(steps):
    x = ML + i * (cw + 0.28)
    rect(s, x, y + 2.85, cw, 0.03, CREAM2)
    t = tb(s, x, y + 3.0, cw, 0.3)
    para(t.text_frame, a, size=11.5, bold=True, color=PURPLE, first=True)
    t = tb(s, x, y + 3.32, cw, 0.7)
    para(t.text_frame, b, size=10, color=GREY, first=True, spacing=1.25)
    if measure(b, 10, cw) > 0.7:
        WARN.append(f"deroule {i}")
footer(s, 14)

# ================================================================ 15 planning
s = base()
y = header(s, "Calendrier", "Planning de déploiement",
           "Séquence de mise en œuvre, à caler avec Marjane et l'équipe commerciale.")
phases = [("J-30", "Validation & production",
           ["Validation du plan et des 15 magasins", "Production PLV, dépliants, stands",
            "Sélection et formation des animatrices"]),
          ("J-15", "Logistique & implantation",
           ["Livraison des stands et des casques", "Installation et contrôle qualité en magasin",
            "Briefing des chefs de rayon"]),
          ("J0 → J+10", "Activation en magasin",
           ["Dégustation et immersion digitale", "Relevé quotidien des contacts et des ventes",
            "Ajustement des objectifs à mi-parcours"]),
          ("J+11", "Bilan & recommandations",
           ["Consolidation des résultats par magasin", "Comparaison réalisé / objectif",
            "Recommandations pour la vague suivante"])]
cw = (CW - 0.2 * 3) / 4
for i, (when, title, bullets) in enumerate(phases):
    x = ML + i * (cw + 0.2)
    rect(s, x, y + 0.08, cw, 0.42, PURPLE_D if i < 3 else GOLD,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.2)
    t = tb(s, x, y + 0.16, cw, 0.3)
    para(t.text_frame, when, size=10.5, bold=True, color=WHITE, first=True, align=PP_ALIGN.CENTER)
    t = tb(s, x, y + 0.62, cw, 0.5)
    para(t.text_frame, title, size=12.5, font=SERIF, bold=True, color=PURPLE_D, first=True, spacing=1.1)
    yy = y + 1.12
    for b in bullets:
        t = tb(s, x, yy, cw, 0.7)
        para(t.text_frame, "—  " + b, size=9.5, color=GREY, first=True, spacing=1.25)
        yy += measure(b, 9.5, cw - 0.2) + 0.14
    if yy > y + 3.7:
        WARN.append(f"planning col {i}")

t = tb(s, ML, y + 3.05, CW, 0.3)
para(t.text_frame, "CONDITIONS DE RÉUSSITE", size=9, bold=True, color=GOLD, first=True)
cond = ["Emplacements validés par Marjane dans les 15 magasins",
        "Casques disponibles et testés avant le lancement",
        "Stock tampon 1 L et 5 L pour absorber la demande"]
for i, c in enumerate(cond):
    x = ML + i * 3.9
    rect(s, x, y + 3.38, 0.05, 0.5, OLIVE if i == 0 else (PURPLE_L if i == 1 else GOLD))
    t = tb(s, x + 0.18, y + 3.34, 3.4, 0.7)
    para(t.text_frame, c, size=9.5, color=INK, first=True, spacing=1.25)
    if measure(c, 9.5, 3.35) > 0.7:
        WARN.append(f"cond {i}")

t = tb(s, ML, y + 4.25, CW, 0.3)
para(t.text_frame, "Calendrier indicatif : les dates fermes seront fixées après validation du plan.",
     size=8.5, italic=True, color=GREY, first=True)
footer(s, 15)

# ================================================================ 16 KPI
s = base(CREAM)
y = header(s, "Pilotage", "Indicateurs de suivi",
           "Quatre indicateurs suffisent à piloter l'activation en cours de route.")

kpi_rows = [("Dégustations + ventes / jour", "505 cumulés", "Relevé quotidien animatrice"),
            ("Moyenne par magasin", "≈ 34 / jour", "Relevé quotidien animatrice"),
            ("Taux de conversion essai → achat", "à fixer avec Marjane", "Relevé + tickets de caisse"),
            ("Unités vendues par format", "suivi 1 L / 5 L", "Ventes magasin"),
            ("Magasins équipés en immersion", "10 / 15", "Comptage équipement")]
rows, cols = len(kpi_rows) + 1, 3
shp = s.shapes.add_table(rows, cols, Inches(ML), Inches(y + 0.05), Inches(CW), Inches(0.5 + 0.55 * (rows - 1)))
table = shp.table
table.columns[0].width = Inches(4.9)
table.columns[1].width = Inches(3.0)
table.columns[2].width = Inches(3.53)
for r in range(rows):
    table.rows[r].height = Inches(0.52)
data = [("Indicateur", "Cible", "Source")] + kpi_rows
for r, row in enumerate(data):
    for c, val in enumerate(row):
        cell = table.cell(r, c)
        cell.margin_left = Inches(0.14)
        cell.margin_top = Inches(0.04)
        cell.margin_bottom = Inches(0.04)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        cell.fill.fore_color.rgb = PURPLE_D if r == 0 else (WHITE if r % 2 else CREAM2)
        p = cell.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = val
        run.font.name = SANS
        run.font.size = Pt(10.5)
        run.font.bold = r == 0 or c == 1
        run.font.color.rgb = WHITE if r == 0 else (PURPLE_D if c == 1 else INK)

t = tb(s, ML, y + 3.5, CW, 0.3)
para(t.text_frame, "Les cibles marquées « à fixer » seront arbitrées avec Marjane avant lancement.",
     size=8.5, italic=True, color=GREY, first=True)
footer(s, 16)

# ================================================================ 17 preuves terrain
s = base()
y = header(s, "Preuves", "L'activation en images",
           "Un dispositif déjà éprouvé en magasin : stand, dégustation et mise en avant linéaire.")
imgs = [(os.path.join(ASSETS, "photo1.jpg"), "Dégustation : pain et huile, conseil par l'animatrice"),
        (os.path.join(ASSETS, "photo2.jpg"), "Mise en avant linéaire, du 1 L au 5 L"),
        (os.path.join(ASSETS, "photo3.jpg"), "Animation en magasin, univers food et famille")]
cw, gap = 3.62, 0.28
for i, (src, cap) in enumerate(imgs):
    x = ML + i * (cw + gap)
    pic(s, src, x, y + 0.05, cw, 3.35, f"proof{i}", focus=(0.5, 0.35))
    t = tb(s, x, y + 3.5, cw, 0.6)
    para(t.text_frame, cap, size=9.5, italic=True, color=GREY, first=True, spacing=1.2)
    if measure(cap, 9.5, cw) > 0.6:
        WARN.append(f"proof cap {i}")
footer(s, 17)

# ================================================================ 18 prochaines étapes
s = base()
y = header(s, "Suite", "Prochaines étapes")

steps = [("Valider", "le calendrier et la liste définitive des 15 magasins."),
         ("Confirmer", "les objectifs quotidiens par magasin (dépliants anniversaire)."),
         ("Produire", "la PLV, les dépliants et le kit d'animation."),
         ("Former", "les animatrices au script darija et au discours produit."),
         ("Cadrer", "le reporting quotidien et le bilan de fin d'opération.")]
yy = y + 0.05
for i, (verb, txt) in enumerate(steps):
    rect(s, ML, yy, 0.05, 0.55, GOLD if i % 2 == 0 else OLIVE)
    t = tb(s, ML + 0.3, yy - 0.02, 3.0, 0.3)
    para(t.text_frame, verb, size=14, font=SERIF, bold=True, color=PURPLE_D, first=True)
    t = tb(s, ML + 3.2, yy, 7.6, 0.55)
    para(t.text_frame, txt, size=12, color=INK, first=True, spacing=1.2)
    if measure(txt, 12, 7.6) > 0.55:
        WARN.append(f"next {i}")
    yy += 0.72

card(s, ML, yy + 0.12, CW, 1.0, fill=PURPLE_D, line=None)
t = tb(s, ML + 0.45, yy + 0.32, CW - 0.9, 0.65)
para(t.text_frame, "Décision attendue", size=9.5, bold=True, color=GOLD, first=True)
para(t.text_frame, "Validation du plan d'activation et du budget de production "
                   "pour lancer la phase de déploiement.",
     size=13, color=WHITE, space_before=6)
if measure("Validation du plan d'activation et du budget de production pour lancer la phase de déploiement.", 13, CW - 0.9) > 0.45:
    WARN.append("decision")
footer(s, 18)

# ================================================================ 19 fin
s = base(PURPLE_D)
pic(s, os.path.join(ASSETS, "olives.png"), 0, 5.55, W, 1.95, "end_band", focus=(0.5, 0.5))
t = tb(s, ML, 1.45, 9.0, 0.3)
para(t.text_frame, "MERCI", size=10, bold=True, color=GOLD, first=True)
t = tb(s, ML, 1.8, 10.5, 1.3)
para(t.text_frame, "Al Horra × Marjane", size=42, font=SERIF, bold=True, color=WHITE, first=True,
     spacing=1.0)
rect(s, ML + 0.02, 3.15, 1.1, 0.05, GOLD)
t = tb(s, ML, 3.42, 9.5, 1.2)
para(t.text_frame, "Activation dégustation & immersion digitale — 2026", size=17, color=WHITE, first=True)
para(t.text_frame, "Access Développement Service · Département Logistique", size=12,
     color=RGBColor(0xD9, 0xCF, 0xE6), space_before=10)
t = tb(s, ML, 4.45, 9.5, 0.4)
para(t.text_frame, "Said Taga  ·  +212 661 648 142  ·  saidoviche@gmail.com", size=11,
     color=GOLD, first=True)

NOTES = {
 1: "Dossier de présentation de l'activation Al Horra en hypermarchés Marjane : contexte, objectifs, dispositif, déploiement et suivi.",
 2: "Quatre chiffres d'ouverture : 15 magasins, 505 contacts/jour, 10 magasins en immersion, un marché encore à 80 % en vrac.",
 3: "Le vrac domine : c'est à la fois la menace et l'opportunité. Le consommateur n'a aucun repère de marque ni de qualité.",
 4: "Trois verrous — la confiance, la preuve, l'essai. La réponse : dégustation + casque immersif.",
 5: "Reprendre les trois objectifs du brief, sans les paraphraser.",
 6: "505 contacts/jour cumulés, moyenne 34 par magasin, Kenitra à 50, cinq magasins à 25.",
 7: "Nouvelle identité : quatre formats, un bloc-marque unique lisible en linéaire.",
 8: "Le film montre le processus de fabrication : c'est la preuve que le vrac ne peut pas apporter.",
 9: "Deux mécaniques complémentaires : le stand installe le contact produit, le casque apporte la preuve.",
 10: "Cinq temps, chacun mesuré sur la fiche de relevé quotidienne.",
 11: "TOP 15 Marjane avec l'objectif quotidien par magasin.",
 12: "Dix magasins équipés du casque : le cœur du dispositif d'immersion.",
 13: "Le dépliant anniversaire fixe l'objectif de chaque magasin : 505 au total.",
 14: "Script darija court : les trois preuves, puis la dégustation.",
 15: "Planning indicatif : les dates fermes restent à valider avec Marjane et l'équipe commerciale.",
 16: "Cinq indicateurs suffisent ; le taux de conversion essai/achat reste à arbitrer avec Marjane.",
 17: "Photos du dispositif déjà déployé en magasin.",
 18: "Les cinq décisions à prendre pour lancer la phase de déploiement.",
 19: "Merci — place aux questions et aux arbitrages.",
}
for i, sl in enumerate(prs.slides, 1):
    txt = NOTES.get(i)
    if txt and not sl.notes_slide.notes_text_frame.text.strip():
        sl.notes_slide.notes_text_frame.text = txt

cp = prs.core_properties
cp.title = "Al Horra x Marjane — Activation dégustation & immersion digitale 2026"
cp.author = "Access Développement Service — Département Logistique"
cp.subject = "Dossier d'activation 2026"
cp.keywords = "Al Horra, Marjane, dégustation, immersion digitale, huile d'olive"
cp.category = "Présentation commerciale"

prs.save(os.path.join(HERE, "Al_Horra_Marjane_Activation_2026.pptx"))
print("slides:", len(prs.slides.__iter__.__self__._sldIdLst))
print("WARNINGS:", len(WARN))
for w in WARN:
    print(" -", w)
