"""Aperçu PNG approximatif du PPTX, pour contrôle visuel (pas un rendu PowerPoint)."""
import io, os, sys
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

DPI = 110
FONTS = {
    (False, False): "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    (True, False): "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    (False, True): "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
    (True, True): "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf",
}
_cache = {}


def font_for(family, bold, px):
    serif = family and "georgia" in family.lower()
    key = (bold, serif, max(6, int(px)))
    if key not in _cache:
        _cache[key] = ImageFont.truetype(FONTS[(bold, serif)], max(6, int(px)))
    return _cache[key]


def emu_in(v):
    return None if v is None else v / 914400.0


def rgb_of(colorobj, default=None):
    try:
        if colorobj is None or colorobj.type is None:
            return default
        return tuple(colorobj.rgb)
    except Exception:
        return default


def shape_fill(sh):
    try:
        f = sh.fill
        if f.type == 1:  # solid
            return tuple(f.fore_color.rgb)
    except Exception:
        pass
    return None


def shape_line(sh):
    try:
        ln = sh.line
        c = rgb_of(ln.color, None)
        if c:
            return c, max(1, int((ln.width or 9525) / 914400 * DPI))
    except Exception:
        pass
    return None, 0


def wrap(draw, text, font, maxw):
    lines = []
    for segment in text.split("\n"):
        words = segment.split(" ")
        cur = ""
        for w in words:
            trial = (cur + " " + w).strip()
            if draw.textlength(trial, font=font) <= maxw or not cur:
                cur = trial
            else:
                lines.append(cur)
                cur = w
        lines.append(cur)
    return lines


def draw_text_frame(draw, tf, x, y, w, h, anchor_top=True):
    ml = emu_in(tf.margin_left) or 0
    mr = emu_in(tf.margin_right) or 0
    mt = emu_in(tf.margin_top) or 0
    mb = emu_in(tf.margin_bottom) or 0
    inner_w = w - ml - mr
    blocks = []
    for p in tf.paragraphs:
        runs = [r for r in p.runs if r.text]
        if not runs:
            blocks.append(([], p.alignment, 0, 0, 1.15, False))
            continue
        r0 = runs[0]
        pt = r0.font.size.pt if r0.font.size else 18
        bold = bool(r0.font.bold)
        fam = r0.font.name or "Arial"
        col = rgb_of(r0.font.color, (0, 0, 0))
        spacing = p.line_spacing or 1.15
        sb = (p.space_before.pt if p.space_before else 0)
        text = "".join(r.text for r in runs)
        blocks.append(([(text, pt, bold, fam, col)], p.alignment, spacing, sb, 0, False))

    # hauteur totale
    rendered = []
    for runs, align, spacing, sb, sa, _ in blocks:
        if not runs:
            rendered.append(([], align, spacing, sb, 0))
            continue
        text, pt, bold, fam, col = runs[0]
        font = font_for(fam, bold, pt * DPI / 72.0)
        lines = wrap(draw, text, font, inner_w * DPI)
        rendered.append(([(l, font, col) for l in lines], align, spacing, sb, pt))
    total_h = 0
    for lines, align, spacing, sb, pt in rendered:
        for _ in lines:
            total_h += (pt or 10) * spacing * 1.2 / 72.0
        total_h += sb / 72.0
    cy = y + mt
    if not anchor_top:
        cy = y + (h - total_h) / 2.0
    for lines, align, spacing, sb, pt in rendered:
        cy += sb / 72.0
        for line, font, col in lines:
            lh = (pt or 10) * spacing * 1.2 / 72.0
            tw = draw.textlength(line, font=font) / DPI
            if align == PP_ALIGN.CENTER:
                tx = x + ml / 2 + (inner_w - tw) / 2
            elif align == PP_ALIGN.RIGHT:
                tx = x + w - mr - tw
            else:
                tx = x + ml
            draw.text((tx * DPI, cy * DPI), line, font=font, fill=col)
            cy += lh


def render(path, outdir):
    prs = Presentation(path)
    Wt = prs.slide_width / 914400.0
    Ht = prs.slide_height / 914400.0
    SW, SH = int(Wt * DPI), int(Ht * DPI)
    os.makedirs(outdir, exist_ok=True)
    outs = []
    for idx, slide in enumerate(prs.slides, 1):
        img = Image.new("RGB", (SW, SH), "white")
        d = ImageDraw.Draw(img)
        try:
            bg = slide.background.fill
            if bg.type == 1:
                img.paste(Image.new("RGB", (SW, SH), tuple(bg.fore_color.rgb)), (0, 0))
                d = ImageDraw.Draw(img)
        except Exception:
            pass
        for sh in slide.shapes:
            x, y = emu_in(sh.left), emu_in(sh.top)
            w, h = emu_in(sh.width), emu_in(sh.height)
            if w is None or h is None:
                continue
            box = [x * DPI, y * DPI, (x + w) * DPI, (y + h) * DPI]
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                im = Image.open(io.BytesIO(sh.image.blob)).convert("RGB")
                im = im.resize((max(1, int(box[2] - box[0])), max(1, int(box[3] - box[1]))), Image.LANCZOS)
                img.paste(im, (int(box[0]), int(box[1])))
                continue
            if getattr(sh, "has_table", False) and sh.has_table:
                tbl = sh.table
                cols = [emu_in(c.width) for c in tbl.columns]
                ys = y
                for r in tbl.rows:
                    rh = emu_in(r.height)
                    xs = x
                    for ci, cell in enumerate(r.cells):
                        cw = cols[ci]
                        f = None
                        try:
                            if cell.fill.type == 1:
                                f = tuple(cell.fill.fore_color.rgb)
                        except Exception:
                            pass
                        d.rectangle([xs * DPI, ys * DPI, (xs + cw) * DPI, (ys + rh) * DPI],
                                    fill=f, outline=(220, 216, 224))
                        draw_text_frame(d, cell.text_frame, xs, ys, cw, rh, True)
                        xs += cw
                    ys += rh
                continue
            fill = shape_fill(sh)
            lc, lw = shape_line(sh)
            st = str(sh.shape_type)
            name = ""
            try:
                name = sh.adjustments._prstGeom if False else ""
            except Exception:
                pass
            prst = sh._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom")
            geom = prst.get("prst") if prst is not None else "rect"
            if geom in ("ellipse",):
                d.ellipse(box, fill=fill, outline=lc, width=max(1, lw))
            elif geom in ("triangle", "rtTriangle") or geom == "isoTriangle":
                d.polygon([(box[2], (box[1] + box[3]) / 2), (box[0], box[1]), (box[0], box[3])],
                          fill=fill, outline=lc)
            elif geom in ("roundRect",):
                d.rounded_rectangle(box, radius=min(box[2] - box[0], box[3] - box[1]) * 0.14,
                                    fill=fill, outline=lc, width=max(1, lw))
            else:
                if fill or lc:
                    d.rectangle(box, fill=fill, outline=lc, width=max(1, lw))
            if sh.has_text_frame and sh.text_frame.text.strip():
                draw_text_frame(d, sh.text_frame, x, y, w, h,
                                sh.text_frame.vertical_anchor != MSO_ANCHOR.MIDDLE)
        out = os.path.join(outdir, f"slide_{idx:02d}.png")
        img.save(out)
        outs.append(out)
    return outs


if __name__ == "__main__":
    outs = render(sys.argv[1], sys.argv[2])
    print(len(outs), "aperçus")
