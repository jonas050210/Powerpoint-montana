#!/usr/bin/env python3
"""Baut die Montana-Präsentation (Morph-Show).

Ergebnis: praesentation/Montana.pptx

Auftakt (Folien 1-3):
  1. Weißer Block mit schwarzen Balken außen herum
  2. Der weiße Block wird zur USA-Karte, die Balken werden die Staatsgrenzen,
     danach fliegen alle 48 Staatsnamen schnell einzeln ein
  3. Zoom auf Montana: Montana wird orange, darunter "Präsentation von Jonas"

Danach: Inhalt mit Morph-Übergängen auf jeder Folie, inkl. Zwischenfolie.

Geodaten: assets/geo/us-states.json (PublicaMundi/MappingAPI, MIT-Lizenz).
Benötigt: python-pptx, pillow, lxml, shapely
Ausführen: python scripts/build_presentation.py [--preview]
"""

import json
import math
import random
import sys
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.util import Inches, Pt
from shapely.geometry import LineString, Point, Polygon, shape
from shapely.ops import unary_union

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets" / "images"
GEO = ROOT / "assets" / "geo" / "us-states.json"
OUT = ROOT / "praesentation" / "Montana.pptx"

EMU = 914400
SW, SH = 13.333, 7.5

# Farben
WHITE = "FFFFFF"
INK = "111827"
BG_DARK = "0B1120"
CARD = "141E33"
GOLD = "E9B44C"
ORANGE = "F28C28"
TEAL = "5EC4B6"
MUTED = "9AA8BF"
TXT = "F5F7FA"
FONT = "Calibri"

EXCLUDE = {"Alaska", "Hawaii", "District of Columbia", "Puerto Rico"}
ANIM_DELAY_DEFAULT = 1400  # ms, Start der Einflüge nach Morph-Übergang


# ---------------------------------------------------------------------------
# Geodaten: Albers-Projektion, Normalisierung auf Breite 1.0 (y nach unten)
# ---------------------------------------------------------------------------
def albers(lon, lat):
    p1, p2, p0 = map(math.radians, (29.5, 45.5, 37.5))
    l0 = math.radians(-96)
    n = (math.sin(p1) + math.sin(p2)) / 2
    C = math.cos(p1) ** 2 + 2 * n * math.sin(p1)
    r0 = math.sqrt(C - 2 * n * math.sin(p0)) / n
    p, l = math.radians(lat), math.radians(lon)
    r = math.sqrt(C - 2 * n * math.sin(p)) / n
    t = n * (l - l0)
    return r * math.sin(t), r0 - r * math.cos(t)


def load_geo():
    data = json.loads(GEO.read_text(encoding="utf-8"))
    raw = {}
    for f in data["features"]:
        name = f["properties"]["name"]
        if name in EXCLUDE:
            continue
        g = shape(f["geometry"])
        parts = [g] if g.geom_type == "Polygon" else list(g.geoms)
        raw[name] = [[albers(*c) for c in p.exterior.coords] for p in parts]

    allp = [Polygon(r) for rings in raw.values() for r in rings]
    union = unary_union(allp)
    minx, miny, maxx, maxy = union.bounds
    W = maxx - minx  # Projektionsbreite

    def norm(x, y):
        return (x - minx) / W, (maxy - y) / W  # u in [0,1], v nach unten

    h = (maxy - miny) / W
    states = {}
    for name, rings in raw.items():
        states[name] = [[norm(*c) for c in ring] for ring in rings]
    # Vereinfachen (gleiche Topologie für den Morph beibehalten)
    tol = 0.0025
    simple = {}
    for name, rings in states.items():
        out = []
        for ring in rings:
            poly = Polygon(ring).simplify(tol, preserve_topology=True)
            out.append(list(poly.exterior.coords)[:-1])
        simple[name] = out
    polys = {n: [Polygon(r) for r in rings] for n, rings in simple.items()}

    # USA-Umriss: Vereinigung, mit Teilen und Löchern
    uni = unary_union([Polygon(r) for rings in simple.values() for r in rings]).simplify(tol, preserve_topology=True)
    uni_parts = [uni] if uni.geom_type == "Polygon" else list(uni.geoms)
    outline = []  # Liste von Konturen (Ring-Koordinaten)
    for part in uni_parts:
        outline.append(list(part.exterior.coords)[:-1])
        for hole in part.interiors:
            outline.append(list(hole.coords)[:-1])

    # Grenz-Segmente: alle Kanten, Duplikate entfernt
    edges = {}
    for rings in simple.values():
        for ring in rings:
            n = len(ring)
            for i in range(n):
                a, b = ring[i], ring[(i + 1) % n]
                key = tuple(sorted([(round(a[0], 5), round(a[1], 5)), (round(b[0], 5), round(b[1], 5))]))
                edges[key] = (a, b)
    segs = list(edges.values())

    # Helena (Hauptstadt)
    hx, hy = albers(-112.0391, 46.5891)
    helena = norm(hx, hy)

    return {
        "states": simple,
        "polys": polys,
        "outline": outline,
        "uni_parts": uni_parts,
        "segs": segs,
        "h": h,
        "helena": helena,
    }


GEO_DATA = load_geo()
H = GEO_DATA["h"]
MT_POLY = GEO_DATA["polys"]["Montana"][0] if len(GEO_DATA["polys"]["Montana"]) == 1 else unary_union(GEO_DATA["polys"]["Montana"])
MT_BOUNDS = MT_POLY.bounds  # (minx, miny, maxx, maxy) in Basis-Koordinaten
MT_CENTER = ((MT_BOUNDS[0] + MT_BOUNDS[2]) / 2, (MT_BOUNDS[1] + MT_BOUNDS[3]) / 2)
MT_W = MT_BOUNDS[2] - MT_BOUNDS[0]
MT_H = MT_BOUNDS[3] - MT_BOUNDS[1]


# ---------------------------------------------------------------------------
# Transformationen (Basis -> Zoll auf der Folie)
# ---------------------------------------------------------------------------
W2 = 8.6  # Breite der USA-Karte auf Folie 2 (Zoll)
X0_2 = SW / 2 - W2 / 2
Y0_2 = SH / 2 - H * W2 / 2 - 0.05


def T2(u, v):
    return X0_2 + u * W2, Y0_2 + v * W2


def zoom_transform(cx, cy, s, push=None):
    """Skaliert um Montanas Mittelpunkt und setzt ihn auf (cx, cy).
    push: Funktion (u,v) -> (dx,dy) in Zoll, zusätzlich verschoben."""
    def f(u, v):
        x = cx + (u - MT_CENTER[0]) * s
        y = cy + (v - MT_CENTER[1]) * s
        if push:
            dx, dy = push(u, v)
            x += dx
            y += dy
        return x, y
    return f


def push_away(u, v, dist=14.0):
    dx, dy = u - MT_CENTER[0], v - MT_CENTER[1]
    d = math.hypot(dx, dy) or 1
    return dx / d * dist, dy / d * dist


# ---------------------------------------------------------------------------
# Hilfsfunktionen für Formen und Text
# ---------------------------------------------------------------------------
def rgb(h):
    return RGBColor.from_string(h)


def emu(v):
    return int(round(v * EMU))


def set_solid_bg(slide, color):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = rgb(color)


def set_gradient_bg(slide, c1, c2, angle=315):
    f = slide.background.fill
    f.gradient()
    f.gradient_angle = angle
    st = f.gradient_stops
    st[0].color.rgb = rgb(c1)
    st[0].position = 0
    st[1].color.rgb = rgb(c2)
    st[1].position = 1.0


def style_shape(shp, fill=None, line=None, line_w=None):
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(fill)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = rgb(line)
        shp.line.width = Pt(line_w or 1)
    return shp


def add_rect(slide, x, y, w, h, fill, name, shape_type=MSO_SHAPE.RECTANGLE, alpha=None, rot=0):
    shp = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.name = name
    style_shape(shp, fill)
    shp.shadow.inherit = False
    if shape_type == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = 0.08
    if alpha is not None:
        srgb = shp._element.spPr.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
        a = etree.SubElement(srgb, "{http://schemas.openxmlformats.org/drawingml/2006/main}alpha")
        a.set("val", str(int(alpha * 100000)))
    if rot:
        shp.rotation = rot
    return shp


def add_text(slide, x, y, w, h, paras, name, anchor=MSO_ANCHOR.TOP, shape=None, fill=None,
             margin=0.0, wrap=True):
    if shape is None:
        shp = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    else:
        shp = add_rect(slide, x, y, w, h, fill or CARD, name, shape_type=shape)
    shp.name = name
    tf = shp.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(margin)
    tf.margin_top = tf.margin_bottom = Inches(margin * 0.6)
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get("align", PP_ALIGN.LEFT)
        if "space_after" in p:
            para.space_after = Pt(p["space_after"])
        r = para.add_run()
        r.text = p["text"]
        f = r.font
        f.name = FONT
        f.size = Pt(p.get("size", 18))
        f.bold = p.get("bold", False)
        f.color.rgb = rgb(p.get("color", TXT))
    return shp


def add_picture_fill(slide, path, x, y, w, h, name):
    iw, ih = Image.open(path).size
    pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))
    pic.name = name
    target, actual = w / h, iw / ih
    if actual > target:
        c = (1 - target / actual) / 2
        pic.crop_left = pic.crop_right = c
    else:
        c = (1 - actual / target) / 2
        pic.crop_top = pic.crop_bottom = c
    return pic


def freeform(slide, contours, name, fill, line=None, line_w=None, shadow=False):
    """contours: Liste von Konturen; jede Kontur Liste von (x,y) in Zoll."""
    first = contours[0]
    b = slide.shapes.build_freeform(emu(first[0][0]), emu(first[0][1]), scale=1.0)
    b.add_line_segments([(emu(x), emu(y)) for x, y in first[1:]], close=True)
    for c in contours[1:]:
        b.move_to(emu(c[0][0]), emu(c[0][1]))
        b.add_line_segments([(emu(x), emu(y)) for x, y in c[1:]], close=True)
    shp = b.convert_to_shape(0, 0)
    shp.name = name
    style_shape(shp, fill, line, line_w)
    if shadow:
        eff = parse_xml(
            '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:outerShdw blurRad="304800" dist="76200" dir="2700000" algn="tl" rotWithShape="0">'
            '<a:srgbClr val="000000"><a:alpha val="55000"/></a:srgbClr></a:outerShdw></a:effectLst>')
        shp._element.spPr.append(eff)
    return shp


def xform_contours(contours, fn):
    return [[fn(u, v) for (u, v) in ring] for ring in contours]


def chrome(slide, label, title, page):
    add_text(slide, 0.7, 0.45, 9, 0.4,
             [{"text": label.upper(), "size": 12, "bold": True, "color": GOLD}], "Label")
    add_text(slide, 0.7, 0.8, 10.4, 0.95,
             [{"text": title, "size": 36, "bold": True, "color": TXT}], "Titel")
    add_rect(slide, 0.7, 1.78, 1.1, 0.05, GOLD, "Linie")
    add_text(slide, 0.7, 6.95, 6, 0.3,
             [{"text": "Montana · Präsentation", "size": 10, "color": MUTED}], "Fusszeile")
    add_text(slide, 11.9, 6.95, 0.75, 0.3,
             [{"text": str(page), "size": 10, "color": MUTED, "align": PP_ALIGN.RIGHT}], "Seitenzahl")


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ---------------------------------------------------------------------------
# Morph-Übergang und Animationen
# ---------------------------------------------------------------------------
def add_morph(slide, dur_ms):
    xml = (
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" '
        'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p159">'
        f'<p:transition spd="slow" p14:dur="{dur_ms}"><p159:morph option="byObject"/></p:transition>'
        '</mc:Choice>'
        '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>'
        '</mc:AlternateContent>'
    )
    el = parse_xml(xml)
    sld = slide._element
    anchor = sld.find("{http://schemas.openxmlformats.org/presentationml/2006/main}clrMapOvr")
    anchor.addnext(el)


DIRS = {
    "left": (8, "ppt_x", "0-#ppt_w/2"),
    "right": (2, "ppt_x", "1+#ppt_w/2"),
    "top": (1, "ppt_y", "0-#ppt_h/2"),
    "bottom": (4, "ppt_y", "1+#ppt_h/2"),
}


class Timing:
    def __init__(self):
        self.nid = 3
        self.groups = []
        self.bld = []

    def _id(self):
        self.nid += 1
        return self.nid

    def _effect(self, spid, direction, node, dur):
        sub, axis, start = DIRS[direction]
        c, s_, x = self._id(), self._id(), self._id()
        return (
            f'<p:par><p:cTn id="{c}" presetID="2" presetClass="entr" presetSubtype="{sub}" fill="hold" '
            f'grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            f'<p:set><p:cBhvr><p:cTn id="{s_}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/>'
            f'</p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst>'
            f'<p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:to><p:strVal val="visible"/></p:to></p:set>'
            f'<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            f'<p:cTn id="{x}" dur="{dur}" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            f'<p:attrNameLst><p:attrName>{axis}</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:tavLst><p:tav tm="0"><p:val><p:strVal val="{start}"/></p:val></p:tav>'
            f'<p:tav tm="100000"><p:val><p:strVal val="#{axis}"/></p:val></p:tav></p:tavLst></p:anim>'
            f'</p:childTnLst></p:cTn></p:par>'
        )

    def auto(self, items, stagger=160, start=ANIM_DELAY_DEFAULT, dur=700):
        """items: (spid, richtung). Starten nach Morph (start ms), gestaffelt."""
        inner = []
        for i, (spid, d) in enumerate(items):
            eff = self._effect(spid, d, "afterEffect", dur)
            pid = self._id()
            inner.append(
                f'<p:par><p:cTn id="{pid}" fill="hold"><p:stCondLst><p:cond delay="{start + i * stagger}"/>'
                f'</p:stCondLst><p:childTnLst>{eff}</p:childTnLst></p:cTn></p:par>')
        g, ig = self._id(), self._id()
        self.groups.append(
            f'<p:par><p:cTn id="{g}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
            f'<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
            f'<p:par><p:cTn id="{ig}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst>{"".join(inner)}</p:childTnLst></p:cTn></p:par>'
            f'</p:childTnLst></p:cTn></p:par>')
        for spid, _ in items:
            self.bld.append(spid)

    def on_click(self, spid, d, dur=500):
        eff = self._effect(spid, d, "clickEffect", dur)
        g, ig = self._id(), self._id()
        self.groups.append(
            f'<p:par><p:cTn id="{g}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
            f'</p:stCondLst><p:childTnLst><p:par><p:cTn id="{ig}" fill="hold"><p:stCondLst>'
            f'<p:cond delay="0"/></p:stCondLst><p:childTnLst>{eff}</p:childTnLst></p:cTn></p:par>'
            f'</p:childTnLst></p:cTn></p:par>')
        self.bld.append(spid)

    def attach(self, slide):
        if not self.groups:
            return
        bld = "".join(f'<p:bldP spid="{s}" grpId="0" animBg="1"/>' for s in dict.fromkeys(self.bld))
        xml = (
            '<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tnLst><p:par>'
            '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
            f'<p:childTnLst>{"".join(self.groups)}</p:childTnLst></p:cTn>'
            '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
            '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
            f'<p:bldLst>{bld}</p:bldLst></p:timing>'
        )
        el = parse_xml(xml)
        sld = slide._element
        alt = sld.xpath("./*[local-name()='AlternateContent']")
        (alt[-1] if alt else sld.find("{http://schemas.openxmlformats.org/presentationml/2006/main}clrMapOvr")).addnext(el)


def ids_of(slide):
    return {s.name: s.shape_id for s in slide.shapes}


# ---------------------------------------------------------------------------
# Folien
# ---------------------------------------------------------------------------
def build():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    blank = prs.slide_layouts[6]
    rng = random.Random(2026)

    def pdir():
        return rng.choice(["left", "right", "top", "bottom"])

    plan = []  # (slide, timing, morph_ms)

    # ======================= FOLIE 1: weißer Block mit schwarzen Balken =======================
    s1 = prs.slides.add_slide(blank)
    set_solid_bg(s1, WHITE)
    uni = GEO_DATA["uni_parts"]
    outline_base = GEO_DATA["outline"]
    # Konturen des USA-Umrisses (Zoll auf Folie 2)
    outline2 = [[T2(u, v) for (u, v) in ring] for ring in outline_base]
    npts = [len(r) for r in outline2]
    # Block-Kontur 1: Rechteck am Folienrand, auf dieselbe Punktzahl verteilt
    def rect_points(n):
        pts = [(0, 0), (SW, 0), (SW, SH), (0, SH)]
        per = 2 * (SW + SH)
        out = []
        for i in range(n):
            d = per * i / n
            if d < SW:
                out.append((d, 0))
            elif d < SW + SH:
                out.append((SW, d - SW))
            elif d < 2 * SW + SH:
                out.append((SW - (d - SW - SH), SH))
            else:
                out.append((0, SH - (d - 2 * SW - SH)))
        return out

    def tiny_circle(cx, cy, n, r=0.004):
        return [(cx + r * math.cos(2 * math.pi * i / n), cy + r * math.sin(2 * math.pi * i / n)) for i in range(n)]

    block_contours = []
    for k, ring in enumerate(outline2):
        if k == 0:
            block_contours.append(rect_points(len(ring)))
        else:
            cu = sum(p[0] for p in ring) / len(ring)
            cv = sum(p[1] for p in ring) / len(ring)
            block_contours.append(tiny_circle(cu, cv, len(ring)))
    umriss1 = freeform(s1, block_contours, "Umriss", WHITE)

    # Balken: Segmente der Staatsgrenzen, am Rand verteilt (lückenlos)
    segs = GEO_DATA["segs"]
    N = len(segs)
    T_FRAME = 0.42
    per = 2 * (SW + SH)
    L = per / N
    frame_pos = []
    for i in range(N):
        d = (i + 0.5) * per / N
        if d < SW:
            frame_pos.append(("top", d, 0.0))
        elif d < SW + SH:
            frame_pos.append(("right", SW, d - SW))
        elif d < 2 * SW + SH:
            frame_pos.append(("bottom", SW - (d - SW - SH), SH))
        else:
            frame_pos.append(("left", 0.0, SH - (d - 2 * SW - SH)))
    cx1, cy1 = SW / 2, SH / 2
    frame_pos.sort(key=lambda p: math.atan2(p[2] - cy1, p[1] - cx1))

    seg_mid = [((a[0] + b[0]) / 2, (a[1] + b[1]) / 2) for a, b in segs]
    seg_ang_center = [math.atan2(T2(*m)[1] - SH / 2, T2(*m)[0] - SW / 2) for m in seg_mid]
    order = sorted(range(N), key=lambda i: seg_ang_center[i])  # Reihenfolge der Segmente

    seg_shape_names = {}
    for rank, i in enumerate(order):
        side, x, y = frame_pos[rank]
        name = f"Seg{i:04d}"
        seg_shape_names[i] = name
        if side in ("top", "bottom"):
            add_rect(s1, x - L / 2, y - (T_FRAME if side == "bottom" else 0), L, T_FRAME, INK, name)
        else:
            add_rect(s1, x - (T_FRAME if side == "right" else 0), y - L / 2, T_FRAME, L, INK, name)
    # Balken im Bild sind jetzt Segmente mit Namen Seg0000... (Reihenfolge gleich wie Grenzen)

    tm1 = Timing()
    # Folie 1 ohne Einflug, die Balken stehen da
    plan.append((s1, tm1, 0))

    # ======================= FOLIE 2: USA-Karte, Grenzen, Namen =======================
    s2 = prs.slides.add_slide(blank)
    set_solid_bg(s2, WHITE)
    umriss2 = freeform(s2, outline2, "Umriss", WHITE)
    # Staatsflächen (weiß, unsichtbar auf weißem Grund) inkl. Montana
    state_shapes = {}
    for name, rings in GEO_DATA["states"].items():
        contours = [[T2(u, v) for (u, v) in ring] for ring in rings]
        state_shapes[name] = freeform(s2, contours, f"Staat_{name}", WHITE)
    # Grenzen (schwarz), selbe Segmente wie auf Folie 1
    for i, (a, b) in enumerate(segs):
        ax, ay = T2(*a)
        bx, by = T2(*b)
        length = math.hypot(bx - ax, by - ay)
        mx, my = (ax + bx) / 2, (ay + by) / 2
        ang = math.degrees(math.atan2(by - ay, bx - ax))
        add_rect(s2, mx - length / 2, my - 0.012, length, 0.024, INK, seg_shape_names[i], rot=ang)

    # Staatsnamen (ausgeschrieben), links 24, rechts 24
    names = sorted(GEO_DATA["states"].keys())
    left_names, right_names = names[:24], names[24:]
    tm2 = Timing()
    name_ids = []
    for col, group, side in ((0, left_names, "left"), (1, right_names, "right")):
        for i, nm in enumerate(group):
            y = 1.15 + i * 0.235
            if col == 0:
                shp = add_text(s2, 0.3, y, 2.0, 0.26, [{"text": nm, "size": 11, "bold": True, "color": INK,
                                                       "align": PP_ALIGN.RIGHT}], f"Name_{nm}",
                               margin=0.02)
            else:
                shp = add_text(s2, 11.15, y, 2.0, 0.26, [{"text": nm, "size": 11, "bold": True, "color": INK}],
                               f"Name_{nm}", margin=0.02)
            name_ids.append((shp.shape_id, side))
    ids2 = ids_of(s2)
    # Nach dem Morph schnell einzeln einfliegen (links und rechts), alternierend von der Seite
    tm2.auto([(sid, side) for sid, side in name_ids], stagger=75, start=2200, dur=320)
    plan.append((s2, tm2, 2200))

    # ======================= FOLIE 3: Zoom auf Montana =======================
    s3 = prs.slides.add_slide(blank)
    set_solid_bg(s3, BG_DARK)
    MT_W3 = 8.2
    s3_scale = MT_W3 / MT_W
    cx3, cy3 = SW / 2, 3.45
    fn3 = zoom_transform(cx3, cy3, s3_scale, push=None)

    def zoom_with_push(u, v, is_mt):
        if is_mt:
            return fn3(u, v)
        dx, dy = push_away(u, v, dist=16.0)
        x, y = fn3(u, v)
        return x + dx, y + dy

    # Umriss (dunkel, verschwindet) - gleiche Punktzahl wie Folie 2
    umriss_out = []
    for ring in outline_base:
        umriss_out.append([zoom_with_push(u, v, False) for (u, v) in ring])
    freeform(s3, umriss_out, "Umriss", BG_DARK)
    # Staaten
    for name, rings in GEO_DATA["states"].items():
        is_mt = name == "Montana"
        contours = [[zoom_with_push(u, v, is_mt) for (u, v) in ring] for ring in rings]
        if is_mt:
            freeform(s3, contours, "Staat_Montana", ORANGE, shadow=True)
        else:
            freeform(s3, contours, f"Staat_{name}", BG_DARK)
    # Grenzen (dunkel), Montana-Kanten bleiben, andere wegschieben
    mt_boundary = MT_POLY.boundary
    for i, (a, b) in enumerate(segs):
        mid = ((a[0] + b[0]) / 2, (a[1] + b[1]) / 2)
        is_mt = mt_boundary.distance(Point(mid)) < 1e-6
        ax, ay = fn3(*a) if is_mt else zoom_with_push(a[0], a[1], False)
        bx, by = fn3(*b) if is_mt else zoom_with_push(b[0], b[1], False)
        length = math.hypot(bx - ax, by - ay)
        mx, my = (ax + bx) / 2, (ay + by) / 2
        ang = math.degrees(math.atan2(by - ay, bx - ax))
        t = 0.028
        add_rect(s3, mx - length / 2, my - t / 2, length, t, BG_DARK, seg_shape_names[i], rot=ang)

    tm3 = Timing()
    lbl3 = add_text(s3, 0.7, 0.5, 4.0, 0.4,
                    [{"text": "MONTANA", "size": 14, "bold": True, "color": GOLD}], "Label")
    praes = add_text(s3, 0.0, 6.35, SW, 0.6,
                     [{"text": "Präsentation von Jonas", "size": 24, "color": TXT,
                       "align": PP_ALIGN.CENTER}], "Praesentation")
    tm3.auto([(praes.shape_id, "bottom"), (lbl3.shape_id, "left")], stagger=150, start=2000, dur=800)
    plan.append((s3, tm3, 2000))

    # ======================= FOLIE 4: Agenda, Montana als Logo oben rechts =======================
    s4 = prs.slides.add_slide(blank)
    set_gradient_bg(s4, "0B1120", "27324A")
    chrome(s4, "01 · Übersicht", "Worum es heute geht", 4)
    s4_scale = 1.35 / MT_W
    fn4 = zoom_transform(12.2, 0.9, s4_scale)
    logo = [[fn4(u, v) for (u, v) in ring] for ring in GEO_DATA["states"]["Montana"]]
    freeform(s4, logo, "Staat_Montana", ORANGE, shadow=True)
    cards = [
        ("01", "Lage", "Wo liegt Montana, und wie groß ist es?"),
        ("02", "Geschichte", "Vom Goldrausch zum 41. Staat"),
        ("03", "Natur", "Glacier National Park und seine Gletscher"),
        ("04", "Wirtschaft", "Landwirtschaft, Bergbau und Tourismus"),
        ("05", "Kultur", "Menschen, Traditionen und Künstler"),
        ("06", "Quiz", "Vier Fragen zum Mitraten"),
    ]
    tm4 = Timing()
    cnames = []
    for i, (nr, t, d) in enumerate(cards):
        col, row = i % 3, i // 3
        x, y = 0.7 + col * 4.1, 2.3 + row * 2.3
        n = f"Karte{i + 1}"
        add_text(s4, x, y, 3.8, 2.0,
                 [{"text": nr, "size": 26, "bold": True, "color": GOLD, "space_after": 4},
                  {"text": t, "size": 22, "bold": True, "color": TXT, "space_after": 4},
                  {"text": d, "size": 14, "color": MUTED}],
                 n, shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.25)
        cnames.append(n)
    ids4 = ids_of(s4)
    tm4.auto([(ids4[n], pdir()) for n in cnames], stagger=150, start=1400)
    plan.append((s4, tm4, 1400))
    set_notes(s4, "Hier siehst du den Ablauf: Lage, Geschichte, Natur, Wirtschaft, Kultur, Quiz.")

    # ======================= FOLIE 5: Lage, Montana links, Helena-Punkt =======================
    s5 = prs.slides.add_slide(blank)
    set_gradient_bg(s5, "0B1120", "27324A")
    chrome(s5, "02 · Lage", "Montana im Überblick", 5)
    s5_scale = 5.0 / MT_W
    fn5 = zoom_transform(3.5, 3.75, s5_scale)
    mt5 = [[fn5(u, v) for (u, v) in ring] for ring in GEO_DATA["states"]["Montana"]]
    freeform(s5, mt5, "Staat_Montana", ORANGE, shadow=True)
    hx, hy = fn5(*GEO_DATA["helena"])
    add_rect(s5, hx - 0.09, hy - 0.09, 0.18, 0.18, WHITE, "Helena_Punkt", shape_type=MSO_SHAPE.OVAL)
    add_text(s5, hx + 0.18, hy - 0.2, 1.8, 0.4,
             [{"text": "Helena", "size": 13, "bold": True, "color": TXT}], "Helena_Label", margin=0.02)
    rows = [("Hauptstadt", "Helena"), ("Größte Stadt", "Billings"),
            ("Fläche", "rund 147.000 Quadratmeilen"), ("Dichte", "ca. 3 Einwohner pro km²")]
    tm5 = Timing()
    rn = []
    for i, (k, v) in enumerate(rows):
        n = f"Zeile{i + 1}"
        add_text(s5, 7.0, 2.2 + i * 1.1, 5.9, 0.9,
                 [{"text": k.upper(), "size": 11, "bold": True, "color": TEAL, "space_after": 2},
                  {"text": v, "size": 20, "bold": True, "color": TXT}],
                 n, shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.25, anchor=MSO_ANCHOR.MIDDLE)
        rn.append(n)
    ids5 = ids_of(s5)
    tm5.auto([(ids5[n], "right") for n in rn] + [(ids5["Helena_Punkt"], "bottom"),
                                              (ids5["Helena_Label"], "left")], stagger=170, start=1400)
    plan.append((s5, tm5, 1400))
    set_notes(s5, "Montana ist der viertgrößte Bundesstaat der USA. Die Hauptstadt ist Helena, die größte Stadt "
                  "Billings. Es leben nur etwa drei Menschen pro Quadratkilometer. Quelle: Wikipedia Montana.")

    # ======================= FOLIE 6: Zwischenfolie Zeitreise =======================
    s6 = prs.slides.add_slide(blank)
    set_solid_bg(s6, BG_DARK)
    add_picture_fill(s6, ASSETS / "titel-hintergrund.jpg", 0, 0, SW, SH, "Hintergrund6")
    add_rect(s6, 0, 0, SW, SH, BG_DARK, "Overlay6", alpha=0.7)
    tm6 = Timing()
    yrs = [("1803", "Louisiana Purchase"), ("1864", "Montana-Territorium"), ("1889", "41. Bundesstaat")]
    yn = []
    for i, (y, t) in enumerate(yrs):
        x = 0.9 + i * 4.1
        add_text(s6, x, 2.6, 3.8, 1.2,
                 [{"text": y, "size": 72, "bold": True, "color": GOLD, "align": PP_ALIGN.CENTER}], f"Jahr{i + 1}")
        add_text(s6, x, 3.9, 3.8, 0.6,
                 [{"text": t, "size": 20, "color": TXT, "align": PP_ALIGN.CENTER}], f"Ereignis{i + 1}")
        yn += [f"Jahr{i + 1}", f"Ereignis{i + 1}"]
    add_text(s6, 0.9, 5.2, 11.5, 0.6,
             [{"text": "Zeitreise", "size": 14, "bold": True, "color": TEAL, "align": PP_ALIGN.CENTER}],
             "ZwischenLabel")
    ids6 = ids_of(s6)
    tm6.auto([(ids6[n], "bottom" if n.startswith("Jahr") else "top") for n in yn] +
             [(ids6["ZwischenLabel"], "top")], stagger=260, start=1400)
    plan.append((s6, tm6, 1400))
    set_notes(s6, "Eine kurze Zeitreise: 1803 kommt das Gebiet zu den USA, 1864 entsteht das Territorium, "
                  "1889 wird Montana der 41. Bundesstaat.")

    # ======================= FOLIE 7: Geschichte =======================
    s7 = prs.slides.add_slide(blank)
    set_gradient_bg(s7, "0B1120", "27324A")
    chrome(s7, "03 · Geschichte", "Vom Goldrausch zum 41. Staat", 7)
    add_rect(s7, 1.0, 3.85, 11.3, 0.05, MUTED, "Zeitachse")
    events = [
        ("1803", "Louisiana Purchase: das Gebiet kommt zu den USA"),
        ("1858", "Goldfund: Zuwanderung und Städte entstehen"),
        ("1864", "Montana-Territorium entsteht (26. Mai)"),
        ("1889", "Aufnahme als 41. Staat (8. November)"),
    ]
    tm7 = Timing()
    au7 = []
    for i, (yr, tx) in enumerate(events):
        cxe = 1.6 + i * 3.3
        add_rect(s7, cxe - 0.16, 3.7, 0.32, 0.32, GOLD, f"Punkt{i + 1}", shape_type=MSO_SHAPE.OVAL)
        add_text(s7, cxe - 1.3, 2.45, 2.6, 0.9,
                 [{"text": yr, "size": 36, "bold": True, "color": GOLD, "align": PP_ALIGN.CENTER}], f"Jahr{i + 1}")
        add_text(s7, cxe - 1.45, 4.35, 2.9, 1.5,
                 [{"text": tx, "size": 16, "color": TXT, "align": PP_ALIGN.CENTER}], f"Text{i + 1}")
        au7 += [f"Jahr{i + 1}", f"Punkt{i + 1}", f"Text{i + 1}"]
    add_text(s7, 0.7, 6.15, 11.9, 0.6,
             [{"text": "Lange vor 1803 lebten hier viele indigene Völker. Ihre Geschichte beginnt nicht mit der Landnahme.",
               "size": 14, "color": MUTED}], "Hinweis")
    au7.append("Hinweis")
    ids7 = ids_of(s7)
    tm7.auto([(ids7[n], pdir()) for n in au7], stagger=170, start=1400)
    plan.append((s7, tm7, 1400))
    set_notes(s7, "1803 kaufen die USA das Gebiet im Louisiana Purchase. 1858 wird Gold gefunden. 1864 entsteht "
                  "das Montana-Territorium. Am 8. November 1889 wird Montana als 41. Staat aufgenommen. "
                  "Quelle: Britannica, History.com.")

    # ======================= FOLIE 8: Natur =======================
    s8 = prs.slides.add_slide(blank)
    set_gradient_bg(s8, "0B1120", "27324A")
    chrome(s8, "04 · Natur", "Glacier National Park", 8)
    stats = [("1910", "Gründung des Nationalparks"),
             ("26", "benannte Gletscher (um 1850: ca. 150)"),
             ("762", "Seen im Park"),
             ("3.190 m", "höchster Gipfel: Mount Cleveland")]
    tm8 = Timing()
    au8 = []
    for i, (num, cap) in enumerate(stats):
        add_text(s8, 0.7, 2.2 + i * 1.12, 5.2, 1.0,
                 [{"text": num, "size": 30, "bold": True, "color": GOLD},
                  {"text": cap, "size": 15, "color": TXT}],
                 f"Stat{i + 1}", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.22, anchor=MSO_ANCHOR.MIDDLE)
        au8.append(f"Stat{i + 1}")
    add_picture_fill(s8, ASSETS / "natur-hintergrund.jpg", 6.5, 0, 6.833, SH, "Bild_Natur")
    ids8 = ids_of(s8)
    tm8.auto([(ids8[n], pdir()) for n in au8], stagger=170, start=1400)
    plan.append((s8, tm8, 1400))
    set_notes(s8, "Der Glacier National Park liegt im Nordwesten Montanas an der kanadischen Grenze, gegründet 1910. "
                  "Heute gibt es 26 benannte Gletscher, um 1850 waren es rund 150. Das Bild ist KI-generiert. "
                  "Quelle: glacier-nationalpark.com, Britannica Kids.")

    # ======================= FOLIE 9: Wirtschaft =======================
    s9 = prs.slides.add_slide(blank)
    set_gradient_bg(s9, "0B1120", "27324A")
    chrome(s9, "05 · Wirtschaft", "Mehr als Rinder und Weizen", 9)
    add_picture_fill(s9, ASSETS / "farm-usgs.jpg", 0.7, 2.2, 5.6, 4.3, "Bild_Farm")
    econ = [("Landwirtschaft", "Rinder- und Getreideanbau, rund 28.000 Farmen"),
            ("Bergbau & Energie", "Öl, Gas, Kohle und Metalle"),
            ("Tourismus", "13,7 Mio. Besucher 2024, so viele wie nie")]
    tm9 = Timing()
    au9 = []
    for i, (t, d) in enumerate(econ):
        add_text(s9, 6.8, 2.2 + i * 1.5, 5.8, 1.3,
                 [{"text": t, "size": 20, "bold": True, "color": TXT, "space_after": 4},
                  {"text": d, "size": 15, "color": MUTED}],
                 f"Karte{i + 1}", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.25, anchor=MSO_ANCHOR.MIDDLE)
        au9.append(f"Karte{i + 1}")
    add_text(s9, 6.8, 6.6, 5.8, 0.35,
             [{"text": "BIP 2025: rund 82,3 Mrd. USD (BEA)", "size": 13, "bold": True, "color": TEAL}], "Kennzahl")
    au9.append("Kennzahl")
    ids9 = ids_of(s9)
    tm9.auto([(ids9[n], pdir()) for n in au9], stagger=170, start=1400)
    plan.append((s9, tm9, 1400))
    set_notes(s9, "Die Wirtschaft ist stark von Landwirtschaft geprägt. Dazu kommen Bergbau und Tourismus. "
                  "Das Farmbild stammt vom USGS und ist gemeinfrei. Quelle: Montana Department of Labor & Industry.")

    # ======================= FOLIE 10: Kultur =======================
    s10 = prs.slides.add_slide(blank)
    set_gradient_bg(s10, "0B1120", "27324A")
    chrome(s10, "06 · Kultur", "Viele Kulturen, eine Landschaft", 10)
    kult = [("11 Stammesnationen", "leben auf sieben Reservaten in Montana. Die Little Shell Band of Chippewa hat kein eigenes Land.", GOLD),
            ("Crow Fair seit 1909", "eines der größten indigenen Treffen Nordamerikas, mit rund 45.000 Besuchern und Teilnehmern.", TEAL),
            ("Charles M. Russell", "„the cowboy artist“ schuf über 2.000 Gemälde über Cowboys, Indigene und Landschaften.", GOLD)]
    tm10 = Timing()
    au10 = []
    for i, (t, d, c) in enumerate(kult):
        y = 2.2 + i * 1.55
        add_text(s10, 0.7, y, 11.9, 1.3,
                 [{"text": t, "size": 22, "bold": True, "color": TXT, "space_after": 4},
                  {"text": d, "size": 16, "color": MUTED}],
                 f"Karte{i + 1}", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.35, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s10, 0.7, y + 0.25, 0.08, 0.8, c, f"Akzent{i + 1}")
        au10 += [f"Karte{i + 1}", f"Akzent{i + 1}"]
    ids10 = ids_of(s10)
    tm10.auto([(ids10[n], pdir()) for n in au10], stagger=170, start=1400)
    plan.append((s10, tm10, 1400))
    set_notes(s10, "Montana ist Heimat von elf Stammesnationen auf sieben Reservaten. Die Crow Fair seit 1909 ist "
                   "eines der größten indigenen Treffen Nordamerikas. Charles M. Russell hat über 2.000 Gemälde "
                   "geschaffen. Quelle: History.com, Montana Beyond, Wikipedia.")

    # ======================= FOLIE 11: Zahlen =======================
    s11 = prs.slides.add_slide(blank)
    set_gradient_bg(s11, "0B1120", "27324A")
    chrome(s11, "07 · Zahlen", "Montana in Zahlen", 11)
    nums = [("1.144.694", "Einwohner (Schätzung 2025)"),
            ("1,01 Mio.", "Acres Glacier National Park"),
            ("13,7 Mio.", "Besucher im Jahr 2024"),
            ("41.", "Bundesstaat seit 8. November 1889")]
    tm11 = Timing()
    au11 = []
    for i, (num, cap) in enumerate(nums):
        add_text(s11, 0.7 + i * 3.05, 2.6, 2.8, 3.3,
                 [{"text": num, "size": 34, "bold": True, "color": GOLD, "space_after": 10},
                  {"text": cap, "size": 16, "color": TXT}],
                 f"Zahl{i + 1}", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.3, anchor=MSO_ANCHOR.MIDDLE)
        au11.append(f"Zahl{i + 1}")
    ids11 = ids_of(s11)
    tm11.auto([(ids11[n], pdir()) for n in au11], stagger=170, start=1400)
    plan.append((s11, tm11, 1400))
    set_notes(s11, "Ein paar Zahlen: rund 1,14 Millionen Einwohner laut Schätzung 2025, rund 1,01 Millionen Acres "
                   "Nationalpark, 13,7 Millionen Besucher 2024 und seit 1889 der 41. Bundesstaat. "
                   "Quelle: Wikipedia, glacier-nationalpark.com.")

    # ======================= FOLIE 12: Fazit =======================
    s12 = prs.slides.add_slide(blank)
    set_gradient_bg(s12, "0B1120", "27324A")
    chrome(s12, "08 · Fazit", "Drei Gedanken zum Mitnehmen", 12)
    pts = ["Montana ist riesig und dünn besiedelt. Weite prägt Leben und Wirtschaft.",
           "Geschichte und Gegenwart hängen zusammen: Gold, Land und indigene Kultur.",
           "Die Natur ist ein Schatz, und die schrumpfenden Gletscher zeigen, wie wichtig Schutz ist."]
    tm12 = Timing()
    au12 = []
    for i, p in enumerate(pts):
        y = 2.2 + i * 1.2
        add_text(s12, 0.7, y, 0.9, 0.9,
                 [{"text": str(i + 1), "size": 30, "bold": True, "color": BG_DARK, "align": PP_ALIGN.CENTER}],
                 f"Nummer{i + 1}", shape=MSO_SHAPE.OVAL, fill=GOLD, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s12, 1.9, y, 10.7, 0.9, [{"text": p, "size": 22, "color": TXT}], f"Punkt{i + 1}",
                 anchor=MSO_ANCHOR.MIDDLE)
        au12 += [f"Nummer{i + 1}", f"Punkt{i + 1}"]
    add_rect(s12, 0.7, 5.85, 11.9, 0.02, MUTED, "Trennlinie")
    add_text(s12, 0.7, 6.0, 11.9, 0.7,
             [{"text": "Big Sky Country: ein Schatz, den es zu bewahren gilt.", "size": 22, "bold": True, "color": GOLD}],
             "Schlussatz")
    au12.append("Schlussatz")
    ids12 = ids_of(s12)
    tm12.auto([(ids12[n], pdir()) for n in au12], stagger=170, start=1400)
    plan.append((s12, tm12, 1400))
    set_notes(s12, "Drei Gedanken zum Schluss: Montana ist weit und dünn besiedelt. Geschichte und Gegenwart "
                   "hängen eng zusammen. Die Natur ist ein Schatz, den man schützen muss. Big Sky Country eben.")

    # ======================= FOLIE 13: Quellen =======================
    s13 = prs.slides.add_slide(blank)
    set_gradient_bg(s13, "0B1120", "27324A")
    chrome(s13, "09 · Quellen", "Quellen & Bildnachweis", 13)
    left = ["Wikipedia: Montana – en.wikipedia.org/wiki/Montana",
            "Wikipedia: History of Montana – en.wikipedia.org/wiki/History_of_Montana",
            "Wikipedia: Native Americans in Montana – en.wikipedia.org/wiki/Native_Americans_in_Montana",
            "Britannica: Montana – britannica.com/place/Montana-state",
            "History.com: Montana – history.com/articles/montana"]
    right = ["Glacier National Park – glacier-nationalpark.com/about",
             "Britannica Kids: Glacier National Park – kids.britannica.com/kids/article/Glacier-National-Park/489368",
             "Montana Dept. of Labor & Industry (2023) – lmi.mt.gov",
             "Montana Beyond: Native American tribes – mtbeyond.com",
             "USGS Image Gallery – usgs.gov/media/images/agricultural-field-great-plains-region-montana",
             "Grenzdaten US-Bundesstaaten – github.com/PublicaMundi/MappingAPI (MIT)"]
    add_text(s13, 0.7, 2.2, 5.8, 3.6, [{"text": t, "size": 12, "color": TXT, "space_after": 9} for t in left],
             "Quellen_Links", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.3)
    add_text(s13, 6.8, 2.2, 5.8, 3.6, [{"text": t, "size": 12, "color": TXT, "space_after": 9} for t in right],
             "Quellen_Links2", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.3)
    add_text(s13, 0.7, 6.0, 11.9, 0.8,
             [{"text": "Bilder: Farmfoto USGS (gemeinfrei). Titel- und Naturbild KI-generiert mit Arena.ai.",
               "size": 13, "color": MUTED, "space_after": 2},
              {"text": "Erstellt mit Arena.ai Agent · Oktober 2026", "size": 13, "color": MUTED}], "Bildnachweis")
    ids13 = ids_of(s13)
    tm13 = Timing()
    tm13.auto([(ids13[n], pdir()) for n in ["Quellen_Links", "Quellen_Links2", "Bildnachweis"]], stagger=200, start=1400)
    plan.append((s13, tm13, 1400))
    set_notes(s13, "Hier sind die Quellen. Hauptquellen: Wikipedia, Britannica, History.com und die offiziellen "
                   "Seiten. Das Farmfoto stammt vom USGS, Titel- und Naturbild sind KI-generiert.")

    # ======================= FOLIE 14: Quiz =======================
    s14 = prs.slides.add_slide(blank)
    set_gradient_bg(s14, "0B1120", "27324A")
    chrome(s14, "10 · Quiz", "Zeit für eure Fragen", 14)
    quiz = [("Wann wurde Montana der 41. US-Bundesstaat?", "8. November 1889"),
            ("Wie heißt die Hauptstadt von Montana?", "Helena"),
            ("Wie viele benannte Gletscher gibt es heute im Glacier National Park?", "26"),
            ("Welchen Rang belegt Montana unter den US-Bundesstaaten nach Fläche?", "Platz 4")]
    tm14 = Timing()
    qn, an = [], []
    for i, (q, a) in enumerate(quiz):
        y = 2.1 + i * 1.18
        add_text(s14, 0.7, y, 7.2, 1.0,
                 [{"text": f"FRAGE {i + 1}", "size": 11, "bold": True, "color": TEAL, "space_after": 2},
                  {"text": q, "size": 15, "bold": True, "color": TXT}],
                 f"Frage{i + 1}", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.22, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s14, 8.2, y, 4.45, 1.0,
                 [{"text": "Lösung", "size": 11, "bold": True, "color": BG_DARK, "space_after": 2},
                  {"text": a, "size": 20, "bold": True, "color": BG_DARK}],
                 f"Antwort{i + 1}", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=GOLD, margin=0.22, anchor=MSO_ANCHOR.MIDDLE)
        qn.append(f"Frage{i + 1}")
        an.append(f"Antwort{i + 1}")
    ids14 = ids_of(s14)
    tm14.auto([(ids14[n], pdir()) for n in qn], stagger=150, start=1400)
    for i, n in enumerate(an):
        tm14.on_click(ids14[n], "right" if i % 2 == 0 else "left")
    plan.append((s14, tm14, 1400))
    set_notes(s14, "Zum Schluss ein Quiz. Frage 1: Wann wurde Montana aufgenommen? Klick: 8. November 1889. "
                   "Frage 2: Die Hauptstadt? Helena. Frage 3: Wie viele benannte Gletscher gibt es heute? 26. "
                   "Frage 4: Rang nach Fläche? Platz 4. Jeder Klick blendet die nächste Lösung ein. Danke!")

    set_notes(s1, "Start: ein weißer Block, außen herum schwarze Balken. Der Block wird gleich zur USA-Karte.")
    set_notes(s2, "Die Balken werden die Staatsgrenzen, der Block wird der Umriss der USA. Danach fliegen alle 48 "
                  "Staatsnamen ein.")
    set_notes(s3, "Zoom auf Montana: Montana wird orange. Präsentation von Jonas.")

    # ---- Übergänge und Animationen anhängen --------------------------------
    # Morph-Dauer je Übergang (ms): der Übergang auf Folie N
    morph_ms = {1: 1400, 2: 2200, 3: 2000}
    for idx, (slide, tm, _) in enumerate(plan, start=1):
        add_morph(slide, morph_ms.get(idx, 1400))
        tm.attach(slide)
    # Notizen für Folien ohne eigene Notizen
    for idx, (slide, _, _) in enumerate(plan, start=1):
        if not slide.has_notes_slide or not slide.notes_slide.notes_text_frame.text:
            set_notes(slide, "")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    print(f"Gespeichert: {OUT}  ({len(prs.slides)} Folien, {N} Grenzsegmente)")


if __name__ == "__main__":
    build()
