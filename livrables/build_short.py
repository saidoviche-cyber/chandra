"""Déclinaison courte (8 slides) — réunion de validation. Même charte que le dossier complet."""
import os
from PIL import Image, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
TMP = "/tmp/short_assets"
os.makedirs(TMP, exist_ok=True)

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
ML, CW = 0.95, 11.43

prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)

FONT = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 40)
WARN = []


def measure(text, pt, width_in, spacing=1.15):
    scale = pt / 40.0
    max_w = (width_in - 0.06) * 72.0 / scale
    lines = 0
    for para_txt in text.split("\n"):
        if not para_txt.strip():
            lines += 1
            continue
        cur, n = "", 1
        for word in para_txt.split(" "):
            trial = (cur + " " + word).strip()
            if FONT.getlength(trial) <= max_w or not cur:
                cur = trial
            else:
                n += 1
                cur = word
        lines += n
    return lines * pt * spacing * 1.2 / 72.0


def tb(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return box


def para(tf, text, *, size=13, font=SANS, color=INK, bold=False, italic=False,
         align=PP_ALIGN.LEFT, space_before=0, spacing=1.15, first=False):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.line_spacing = spacing
    p.space_before = Pt(space_before)
    run = p.add_run()
    run.text = text
    f = run.font
    f.name = font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    return p


def rect(slide, x, y, w, h, fill, line=None, lw=0.75, shape=MSO_SHAPE.RECTANGLE, adj=None):
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
    s.shadow.inherit = False
    if adj is not None:
        s.adjustments[0] = adj
    s.text_frame.margin_left = s.text_frame.margin_right = 0
    s.text_frame.margin_top = s.text_frame.margin_bottom = 0
    return s


def card(slide, x, y, w, h, fill=CREAM, line=CREAM2, accent=None):
    c = rect(slide, x, y, w, h, fill, line, 0.75, MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.06)
    if accent:
        rect(slide, x + 0.30, y - 0.02, 0.62, 0.05, accent)
    return c


def fit_image(src, w_in, h_in, name, focus=(0.5, 0.5)):
    out = os.path.join(TMP, f"{name}.png")
    im = Image.open(src).convert("RGB")
    target = w_in / h_in
    iw, ih = im.size
    if iw / ih > target:
        nw = int(ih * target)
        cx = int((iw - nw) * focus[0])
        im = im.crop((cx, 0, cx + nw, ih))
    else:
        nh = int(iw / target)
        cy = int((ih - nh) * focus[1])
        im = im.crop((0, cy, iw, cy + nh))
    im.resize((int(w_in * 150), int(h_in * 150)), Image.LANCZOS).save(out, quality=92)
    return out


def pic(slide, src, x, y, w, h, name, **kw):
    return slide.shapes.add_picture(fit_image(src, w, h, name, **kw),
                                    Inches(x), Inches(y), Inches(w), Inches(h))


def base(color=WHITE):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = color
    return s


def header(slide, kicker, title, sub=None):
    t = tb(slide, ML, 0.6, 8.0, 0.25)
    para(t.text_frame, kicker.upper(), size=9.5, bold=True, color=GOLD, first=True)
    th = 0.62
    t2 = tb(slide, ML, 0.9, CW, th)
    para(t2.text_frame, title, size=29, font=SERIF, bold=True, color=PURPLE_D, first=True, spacing=1.05)
    if measure(title, 29, CW, 1.05) > th:
        WARN.append(f"titre long : {title}")
    y = 0.9 + th + 0.12
    rect(slide, ML + 0.02, y, 0.85, 0.05, GOLD)
    if sub:
        ts = tb(slide, ML, y + 0.2, CW * 0.9, 0.5)
        para(ts.text_frame, sub, size=12, color=GREY, first=True, spacing=1.2)
        hsub = measure(sub, 12, CW * 0.9)
        if hsub > 0.5:
            WARN.append(f"sur-titre long : {sub[:40]}")
        return y + 0.2 + hsub + 0.22
    return y + 0.33


def footer(slide, n):
    rect(slide, ML, 6.95, CW, 0.008, RGBColor(0xDD, 0xD8, 0xE0))
    f = tb(slide, ML, 7.05, 9.0, 0.25)
    para(f.text_frame, "Al Horra  ×  Marjane   ·   Activation 2026", size=8.5, color=GREY, first=True)
    pg = tb(slide, W - ML - 0.6, 7.05, 0.6, 0.25)
    para(pg.text_frame, f"{n:02d}", size=8.5, bold=True, color=PURPLE, align=PP_ALIGN.RIGHT, first=True)


def notes(slide, txt):
    slide.notes_slide.notes_text_frame.text = txt


# ---------------------------------------------------------------- 01
s = base()
pic(s, os.path.join(ASSETS, "olives.png"), 0, 0, W, 2.35, "s_cover", focus=(0.5, 0.55))
rect(s, 0, 2.35, W, 0.06, GOLD)

t = tb(s, ML, 2.85, 10.5, 0.3)
para(t.text_frame, "RÉUNION DE VALIDATION  ·  NOVEMBRE 2026", size=10, bold=True, color=GOLD, first=True)
t = tb(s, ML, 3.2, 11.0, 1.3)
para(t.text_frame, "Al Horra × Marjane", size=52, font=SERIF, bold=True, color=PURPLE_D, first=True, spacing=0.95)
rect(s, ML + 0.02, 4.7, 1.1, 0.055, GOLD)
t = tb(s, ML, 4.95, 9.5, 1.1)
para(t.text_frame, "Activation dégustation & immersion digitale", size=20, bold=True, color=PURPLE, first=True)
para(t.text_frame, "L'essentiel du dispositif 2026, en huit slides.", size=13, color=GREY, space_before=8)
t = tb(s, ML, 6.55, 9.5, 0.3)
para(t.text_frame, "Access Développement Service  ·  Département Logistique", size=9.5, color=GREY, first=True)
notes(s, "Version courte du dossier : contexte, objectifs, dispositif, déploiement, pilotage et décisions attendues.")

# ---------------------------------------------------------------- 02
s = base(CREAM)
y = header(s, "Contexte", "Un marché encore en vrac", "Huit litres sur dix se vendent sans marque, sans repère de qualité.")

card(s, ML, y + 0.05, 3.55, 2.15, fill=WHITE, line=CREAM2, accent=GOLD)
t = tb(s, ML + 0.35, y + 0.32, 2.9, 1.0)
para(t.text_frame, "80 %", size=52, font=SERIF, bold=True, color=PURPLE, first=True)
t = tb(s, ML + 0.35, y + 1.38, 2.9, 1.05)
para(t.text_frame, "de la consommation marocaine d'huile d'olive se fait en vrac.", size=12, color=INK, first=True, spacing=1.3)
if measure("de la consommation marocaine d'huile d'olive se fait en vrac.", 12, 2.9) > 1.0:
    WARN.append("80 % texte")

locks = [("La confiance", "Le vrac paraît plus authentique : la qualité doit se prouver par le goût."),
         ("La preuve", "Montrer la fabrication pour crédibiliser le 100 % marocain et pur."),
         ("L'essai", "Créer le contact produit au moment exact de la décision d'achat.")]
x0 = ML + 3.85
for i, (title, txt) in enumerate(locks):
    yy = y + 0.05 + i * 0.88
    rect(s, x0, yy + 0.08, 0.05, 0.5, [OLIVE, PURPLE_L, GOLD][i])
    t = tb(s, x0 + 0.22, yy, 7.3, 0.3)
    para(t.text_frame, title, size=12.5, bold=True, color=PURPLE_D, first=True)
    t = tb(s, x0 + 0.22, yy + 0.3, 7.3, 0.5)
    para(t.text_frame, txt, size=10.5, color=GREY, first=True, spacing=1.2)
    if measure(txt, 10.5, 7.3) > 0.5:
        WARN.append(f"verrou {i}")

card(s, ML, y + 2.95, CW, 0.95, fill=PURPLE_D, line=None)
t = tb(s, ML + 0.45, y + 3.18, CW - 0.9, 0.6)
para(t.text_frame, "Notre réponse : dégustation + vidéo immersive sur casque — goûter, voir, acheter.",
     size=14, color=WHITE, first=True, spacing=1.15)
if measure("Notre réponse : dégustation + vidéo immersive sur casque — goûter, voir, acheter.", 14, CW - 0.9) > 0.6:
    WARN.append("reponse band")
footer(s, 2)

# ---------------------------------------------------------------- 03
s = base()
y = header(s, "Objectifs", "Ce que l'on veut obtenir")

objs = [("Renforcer", "le positionnement d'Al Horra comme huile 100 % marocaine, naturelle et pure."),
        ("Valoriser", "le savoir-faire et le processus de fabrication pour gagner la confiance."),
        ("Favoriser", "l'essai et le recrutement par la dégustation.")]
cw = (CW - 0.56) / 3
for i, (verb, txt) in enumerate(objs):
    x = ML + i * (cw + 0.28)
    card(s, x, y + 0.05, cw, 1.95, fill=CREAM, line=CREAM2, accent=[GOLD, PURPLE_L, OLIVE][i])
    t = tb(s, x + 0.34, y + 0.28, cw - 0.68, 0.35)
    para(t.text_frame, f"0{i+1}", size=10, bold=True, color=GOLD, first=True)
    t = tb(s, x + 0.34, y + 0.55, cw - 0.68, 0.4)
    para(t.text_frame, verb, size=18, font=SERIF, bold=True, color=PURPLE_D, first=True)
    t = tb(s, x + 0.34, y + 1.0, cw - 0.68, 0.8)
    para(t.text_frame, txt, size=11, color=GREY, first=True, spacing=1.25)
    if measure(txt, 11, cw - 0.68) > 0.8:
        WARN.append(f"objectif {i}")

stats = [("15", "hypermarchés Marjane"), ("505", "contacts + ventes / jour"), ("10", "magasins en immersion")]
sy = y + 2.35
for i, (num, lab) in enumerate(stats):
    x = ML + i * 3.9
    t = tb(s, x, sy, 3.4, 0.7)
    para(t.text_frame, num, size=42, font=SERIF, bold=True, color=PURPLE, first=True)
    t = tb(s, x, sy + 0.72, 3.4, 0.35)
    para(t.text_frame, lab, size=11.5, color=GREY, first=True)
    if i < 2:
        rect(s, x + 3.72, sy + 0.05, 0.01, 0.95, CREAM2)
footer(s, 3)

# ---------------------------------------------------------------- 04
s = base(CREAM)
y = header(s, "Dispositif", "Deux mécaniques, une seule expérience",
           "Le stand installe le contact produit, le casque apporte la preuve.")
cw = 4.6
card(s, ML, y + 0.05, cw, 2.95, fill=WHITE, line=CREAM2, accent=OLIVE)
t = tb(s, ML + 0.34, y + 0.3, cw - 0.68, 0.3)
para(t.text_frame, "1", size=11, bold=True, color=OLIVE, first=True)
t = tb(s, ML + 0.34, y + 0.55, cw - 0.68, 0.4)
para(t.text_frame, "Stand de dégustation", size=16, font=SERIF, bold=True, color=PURPLE_D, first=True)
for i, b in enumerate(["Pain et huile servis par une animatrice formée",
                       "Argumentaire 100 % naturel et pur",
                       "Mise en avant des formats 1 L et 5 L",
                       "Dépliant anniversaire remis au client"]):
    t = tb(s, ML + 0.34, y + 1.15 + i * 0.45, cw - 0.68, 0.4)
    para(t.text_frame, "—  " + b, size=10.5, color=INK, first=True, spacing=1.2)

cx = ML + cw + 0.45
t = tb(s, cx, y + 1.3, 0.55, 0.5)
para(t.text_frame, "+", size=32, font=SERIF, bold=True, color=GOLD, first=True, align=PP_ALIGN.CENTER)

x2 = cx + 0.65
card(s, x2, y + 0.05, cw, 2.95, fill=WHITE, line=CREAM2, accent=PURPLE_L)
t = tb(s, x2 + 0.34, y + 0.3, cw - 0.68, 0.3)
para(t.text_frame, "2", size=11, bold=True, color=PURPLE_L, first=True)
t = tb(s, x2 + 0.34, y + 0.55, cw - 0.68, 0.4)
para(t.text_frame, "Immersion digitale", size=16, font=SERIF, bold=True, color=PURPLE_D, first=True)
for i, b in enumerate(["Casque de réalité virtuelle sur stand",
                       "Film du processus de fabrication",
                       "10 magasins équipés en priorité",
                       "Quelques minutes par visiteur"]):
    t = tb(s, x2 + 0.34, y + 1.15 + i * 0.45, cw - 0.68, 0.4)
    para(t.text_frame, "—  " + b, size=10.5, color=INK, first=True, spacing=1.2)

t = tb(s, ML, y + 3.25, 8.2, 0.9)
para(t.text_frame, "Le casque est installé sur le stand, dans le parcours de dégustation.",
     size=12, italic=True, color=PURPLE, first=True)
para(t.text_frame, "Objectif : que le consommateur voie ce que le vrac ne peut pas montrer.",
     size=10, color=GREY, space_before=4)
footer(s, 4)

# ---------------------------------------------------------------- 05
s = base()
y = header(s, "Mécanique", "Le parcours en cinq temps")
steps = [("Attirer", "Stand visible, signalétique 100 % marocaine", "contacts stand"),
         ("Déguster", "Pain + huile, comparer avec le vrac", "dégustations"),
         ("S'immerger", "Casque : récolte et fabrication", "temps d'immersion"),
         ("Convertir", "Dépliant anniversaire, format adapté", "unités vendues"),
         ("Fidéliser", "Réachat, relevé des ventes", "réachat")]
n = len(steps)
cw = (CW - 0.15 * (n - 1)) / n
for i, (title, txt, micro) in enumerate(steps):
    x = ML + i * (cw + 0.15)
    rect(s, x, y + 0.1, cw, 2.72, WHITE, CREAM2, 0.75, MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.07)
    rect(s, x + 0.3, y + 0.36, 0.42, 0.42, PURPLE if i < 4 else GOLD, shape=MSO_SHAPE.OVAL)
    t = tb(s, x + 0.3, y + 0.43, 0.42, 0.25)
    para(t.text_frame, str(i + 1), size=11.5, bold=True, color=WHITE, first=True, align=PP_ALIGN.CENTER)
    t = tb(s, x + 0.3, y + 0.95, cw - 0.6, 0.35)
    para(t.text_frame, title, size=13.5, font=SERIF, bold=True, color=PURPLE_D, first=True)
    t = tb(s, x + 0.3, y + 1.35, cw - 0.6, 1.0)
    para(t.text_frame, txt, size=9.5, color=GREY, first=True, spacing=1.25)
    rect(s, x + 0.3, y + 2.36, cw - 0.6, 0.012, CREAM2)
    t = tb(s, x + 0.3, y + 2.44, cw - 0.6, 0.25)
    para(t.text_frame, micro.upper(), size=7.5, bold=True, color=GOLD, first=True)
    if i < n - 1:
        ar = rect(s, x + cw + 0.02, y + 1.2, 0.11, 0.15, GOLD, shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
        ar.rotation = 90

t = tb(s, ML, y + 3.12, CW, 1.0)
para(t.text_frame, "Chaque temps est documenté par l'animatrice sur sa fiche de relevé : "
                   "contacts, dégustations, temps d'immersion, unités vendues.", size=12, color=INK, first=True, spacing=1.3)
para(t.text_frame, "Le même support alimente le reporting quotidien et le bilan de fin d'opération.",
     size=10.5, color=GREY, space_before=6)
if measure("Chaque temps est documenté par l'animatrice sur sa fiche de relevé : contacts, dégustations, temps d'immersion, unités vendues.", 12, CW) > 0.6:
    WARN.append("parcours note")
footer(s, 5)

# ---------------------------------------------------------------- 06
s = base(CREAM)
y = header(s, "Déploiement", "Le TOP 15 Marjane, avec ses objectifs",
           "Objectif de dégustations + ventes par jour, par magasin.")
stores = ["Ain Sbaa", "El Jadida", "Salé", "Bouregreg", "Fès 2", "Menara", "Massira",
          "Agadir Founty", "Kenitra", "Tanger Medina", "Mohammedia", "Derb Sultan",
          "Hay Hassani", "Sidi Othmane", "Oujda"]
quota = [40, 25, 40, 30, 30, 40, 40, 40, 50, 25, 25, 30, 40, 25, 25]
total = sum(quota)
cw, ch, gap = 2.16, 0.86, 0.155
for i, (st, q) in enumerate(zip(stores, quota)):
    col, row = i % 5, i // 5
    x = ML + col * (cw + gap)
    yy = y + 0.05 + row * (ch + gap)
    rect(s, x, yy, cw, ch, WHITE if row % 2 else CREAM2, CREAM2, 0.75, MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.08)
    t = tb(s, x + 0.18, yy + 0.11, 0.5, 0.25)
    para(t.text_frame, f"{i+1:02d}", size=9, bold=True, color=GOLD, first=True)
    t = tb(s, x + 0.18, yy + 0.33, cw - 0.36, 0.3)
    para(t.text_frame, st, size=10.5, bold=True, color=PURPLE_D, first=True)
    t = tb(s, x + 0.18, yy + 0.58, cw - 0.36, 0.25)
    para(t.text_frame, f"{q} / jour", size=8.5, color=GREY, first=True)
    if measure(st, 10.5, cw - 0.36) > 0.3:
        WARN.append(f"magasin {st}")

by = y + 3.15
card(s, ML, by, CW, 0.85, fill=PURPLE_D, line=None)
t = tb(s, ML + 0.42, by + 0.16, 5.4, 0.5)
para(t.text_frame, f"{total}", size=24, font=SERIF, bold=True, color=WHITE, first=True)
t = tb(s, ML + 1.75, by + 0.4, 4.4, 0.4)
para(t.text_frame, "contacts + ventes / jour", size=10, color=RGBColor(0xD9, 0xCF, 0xE6), first=True)
t = tb(s, ML + 6.5, by + 0.33, CW - 6.9, 0.4)
para(t.text_frame, "≈ 34 par magasin  ·  pic 50 (Kenitra)  ·  10 magasins avec casque",
     size=9.5, color=RGBColor(0xD9, 0xCF, 0xE6), first=True)
if measure("≈ 34 par magasin  ·  pic 50 (Kenitra)  ·  10 magasins avec casque", 9.5, CW - 6.9) > 0.4:
    WARN.append("bandeau totaux")
footer(s, 6)

# ---------------------------------------------------------------- 07
s = base()
y = header(s, "Calendrier & pilotage", "Quatre phases, cinq indicateurs")
lw = 6.6
phases = [("J-30", "Valider et produire", "Plan, magasins, PLV, formation des animatrices"),
          ("J-15", "Implanter", "Livraison stands et casques, briefing des chefs de rayon"),
          ("J0 → J+10", "Activer", "Dégustation, immersion, relevé quotidien des ventes"),
          ("J+11", "Tirer le bilan", "Réalisé / objectif par magasin, recommandations")]
for i, (when, title, txt) in enumerate(phases):
    yy = y + 0.02 + i * 1.02
    rect(s, ML, yy + 0.06, 0.05, 0.72, [PURPLE_D, PURPLE, PURPLE_L, GOLD][i])
    t = tb(s, ML + 0.22, yy, 1.35, 0.3)
    para(t.text_frame, when, size=9.5, bold=True, color=GOLD, first=True)
    t = tb(s, ML + 1.55, yy - 0.03, lw - 1.6, 0.3)
    para(t.text_frame, title, size=12.5, font=SERIF, bold=True, color=PURPLE_D, first=True)
    t = tb(s, ML + 1.55, yy + 0.26, lw - 1.6, 0.45)
    para(t.text_frame, txt, size=9.5, color=GREY, first=True, spacing=1.2)
    if measure(txt, 9.5, lw - 1.6) > 0.45:
        WARN.append(f"phase {i}")

x = ML + lw + 0.5
pw = W - ML - x
kpis = [("505", "contacts + ventes / jour"), ("≈ 34", "moyenne par magasin"),
        ("10 / 15", "magasins équipés"), ("à fixer", "taux de conversion essai → achat")]
for i, (num, lab) in enumerate(kpis):
    yy = y + 0.02 + i * 0.92
    t = tb(s, x, yy, pw, 0.45)
    para(t.text_frame, num, size=19, font=SERIF, bold=True, color=PURPLE, first=True)
    t = tb(s, x, yy + 0.42, pw, 0.4)
    para(t.text_frame, lab, size=9.5, color=GREY, first=True, spacing=1.15)
    if measure(lab, 9.5, pw) > 0.4:
        WARN.append(f"kpi court {i}")

t = tb(s, ML, y + 4.2, CW, 0.3)
para(t.text_frame, "Calendrier indicatif : les dates fermes seront fixées après validation du plan.",
     size=8.5, italic=True, color=GREY, first=True)
footer(s, 7)

# ---------------------------------------------------------------- 08
s = base(PURPLE_D)
t = tb(s, ML, 0.85, 9.0, 0.3)
para(t.text_frame, "DÉCISION", size=10, bold=True, color=GOLD, first=True)
t = tb(s, ML, 1.15, 10.5, 0.8)
para(t.text_frame, "Ce qu'il nous faut valider", size=34, font=SERIF, bold=True, color=WHITE, first=True)
rect(s, ML + 0.02, 2.0, 1.1, 0.05, GOLD)

items = [("Le plan et les 15 magasins", "périmètre et priorités du déploiement"),
         ("Les objectifs par point de vente", "base des dépliants anniversaire"),
         ("Le budget de production", "PLV, stands, dépliants, casques"),
         ("Le calendrier de lancement", "rétroplanning J-30 → J+11")]
for i, (a, b) in enumerate(items):
    x = ML + (i % 2) * 5.6
    yy = 2.42 + (i // 2) * 0.95
    rect(s, x, yy + 0.08, 0.05, 0.5, GOLD if i % 2 == 0 else PURPLE_L)
    t = tb(s, x + 0.22, yy, 4.9, 0.3)
    para(t.text_frame, a, size=13, bold=True, color=WHITE, first=True)
    t = tb(s, x + 0.22, yy + 0.3, 4.9, 0.35)
    para(t.text_frame, b, size=10, color=RGBColor(0xD9, 0xCF, 0xE6), first=True)
    if measure(a, 13, 4.9) > 0.3:
        WARN.append(f"decision {i}")

rect(s, ML, 4.55, CW, 0.008, RGBColor(0x5B, 0x3A, 0x70))
t = tb(s, ML, 4.75, CW, 0.9)
para(t.text_frame, "Dès validation, la production et la logistique démarrent sous quinze jours.",
     size=13.5, italic=True, color=WHITE, first=True, spacing=1.2)

t = tb(s, ML, 5.85, 9.5, 0.3)
para(t.text_frame, "Said Taga  ·  Access Développement Service  ·  Département Logistique", size=10.5, color=GOLD, first=True)
t = tb(s, ML, 6.15, 9.5, 0.3)
para(t.text_frame, "+212 661 648 142  ·  saidoviche@gmail.com", size=9.5, color=RGBColor(0xD9, 0xCF, 0xE6), first=True)
t = tb(s, ML, 6.75, CW, 0.3)
para(t.text_frame, "Merci — place aux arbitrages.", size=11, italic=True, color=RGBColor(0xD9, 0xCF, 0xE6), first=True)
notes(s, "Quatre décisions à obtenir : périmètre, objectifs, budget, calendrier. Sans elles, la production ne peut pas démarrer.")

SHORT_NOTES = {
 2: "Le vrac reste majoritaire : c'est le gisement. Trois verrous à lever — confiance, preuve, essai.",
 3: "Trois objectifs du brief, puis les trois chiffres du dispositif : 15 magasins, 505 contacts/jour, 10 magasins en immersion.",
 4: "Le stand installe le contact produit ; le casque apporte la preuve par l'image du process.",
 5: "Cinq temps, tous mesurés sur la même fiche de relevé.",
 6: "TOP 15 avec objectif quotidien par magasin : 505 au total, moyenne 34, pic à Kenitra.",
 7: "Quatre phases indicatives et cinq indicateurs. Le taux de conversion reste à arbitrer avec Marjane.",
}
for i, sl in enumerate(prs.slides, 1):
    txt = SHORT_NOTES.get(i)
    if txt and not sl.notes_slide.notes_text_frame.text.strip():
        sl.notes_slide.notes_text_frame.text = txt

cp = prs.core_properties
cp.title = "Al Horra x Marjane — Activation 2026 (version courte)"
cp.author = "Access Développement Service — Département Logistique"
cp.subject = "Réunion de validation — 8 slides"

prs.save(os.path.join(HERE, "Al_Horra_Marjane_Activation_2026_version_courte.pptx"))
print("slides :", len(prs.slides._sldIdLst))
print("alertes :", WARN)
