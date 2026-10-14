#!/usr/bin/env python3
"""Baut die Montana-Präsentation.

Erzeugt praesentation/Montana.pptx mit:
- Morph-Übergang auf allen Folien (PowerPoint 2019 / Microsoft 365)
- Einflug-Animationen (Richtung pro Element festgelegt, teils zufällig, deterministisch per Seed)
- Quiz-Folie: Antworten fliegen per Klick einzeln ein
- Sprechernotizen auf jeder Folie

Ausführen:  python scripts/build_presentation.py
Benötigt:   python-pptx, pillow
"""

import random
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml import parse_xml
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets" / "images"
OUT = ROOT / "praesentation" / "Montana.pptx"

# Farben und Schrift
BG = "0B1120"
CARD = "141E33"
CARD2 = "1C2A47"
GOLD = "E9B44C"
TEAL = "5EC4B6"
WHITE = "F5F7FA"
MUTED = "9AA8BF"
DARK = "0B1120"
FONT = "Calibri"

SLIDE_W, SLIDE_H = 13.333, 7.5

PRES_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
NSMAP_P = {"p": PRES_NS}

# Richtung -> (presetSubtype, Achse, Start-Ausdruck)
DIRS = {
    "left": (8, "ppt_x", "0-#ppt_w/2"),
    "right": (2, "ppt_x", "1+#ppt_w/2"),
    "top": (1, "ppt_y", "0-#ppt_h/2"),
    "bottom": (4, "ppt_y", "1+#ppt_h/2"),
}


def rgb(hex_):
    return RGBColor.from_string(hex_)


def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)


def add_rect(slide, x, y, w, h, fill, name, shape=MSO_SHAPE.RECTANGLE, alpha=None, line=None):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.name = name
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb(fill)
    if line:
        shp.line.color.rgb = rgb(line)
        shp.line.width = Pt(1)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = 0.08
    if alpha is not None:
        srgb = shp.fill._xPr.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
        a = etree.SubElement(srgb, "{http://schemas.openxmlformats.org/drawingml/2006/main}alpha")
        a.set("val", str(int(alpha * 100000)))
    return shp


def add_text(slide, x, y, w, h, paragraphs, name, anchor=MSO_ANCHOR.TOP, shape=None, fill=None,
             margin=0.0, line=None):
    """paragraphs: Liste von dicts mit text, size, color, bold, align, space_after."""
    if shape is None:
        shp = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    else:
        shp = add_rect(slide, x, y, w, h, fill or CARD, name, shape=shape, line=line)
    shp.name = name
    tf = shp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    m = Inches(margin)
    tf.margin_left = tf.margin_right = m
    tf.margin_top = tf.margin_bottom = Inches(margin * 0.6)
    for i, p in enumerate(paragraphs):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get("align", PP_ALIGN.LEFT)
        if "space_after" in p:
            para.space_after = Pt(p["space_after"])
        run = para.add_run()
        run.text = p["text"]
        f = run.font
        f.name = FONT
        f.size = Pt(p.get("size", 18))
        f.bold = p.get("bold", False)
        f.color.rgb = rgb(p.get("color", WHITE))
    return shp


def add_picture_fill(slide, path, x, y, w, h, name):
    """Bild so zuschneiden, dass es den Rahmen ausfüllt (Mitte zentriert)."""
    img_w, img_h = Image.open(path).size
    pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))
    pic.name = name
    target = w / h
    actual = img_w / img_h
    if actual > target:  # zu breit -> links/rechts beschneiden
        cut = (1 - target / actual) / 2
        pic.crop_left = cut
        pic.crop_right = cut
    else:  # zu hoch -> oben/unten beschneiden
        cut = (1 - actual / target) / 2
        pic.crop_top = cut
        pic.crop_bottom = cut
    return pic


def chrome(slide, label, title, page):
    """Wiederkehrende Elemente: Kicker, Titel, Akzentlinie, Fußzeile, Seitenzahl."""
    add_text(slide, 0.7, 0.45, 9, 0.4,
             [{"text": label.upper(), "size": 12, "bold": True, "color": GOLD}], "Label")
    add_text(slide, 0.7, 0.8, 11.9, 0.95,
             [{"text": title, "size": 36, "bold": True, "color": WHITE}], "Titel")
    add_rect(slide, 0.7, 1.78, 1.1, 0.05, GOLD, "Linie")
    add_text(slide, 0.7, 6.95, 6, 0.3,
             [{"text": "Montana · Präsentation", "size": 10, "color": MUTED}], "Fusszeile")
    add_text(slide, 11.9, 6.95, 0.75, 0.3,
             [{"text": str(page), "size": 10, "color": MUTED, "align": PP_ALIGN.RIGHT}], "Seitenzahl")


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ---------------------------------------------------------------------------
# Morph-Übergang
# ---------------------------------------------------------------------------
def add_morph_transition(slide):
    xml = (
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" '
        'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p159">'
        '<p:transition spd="slow" p14:dur="1100"><p159:morph option="byObject"/></p:transition>'
        '</mc:Choice>'
        '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>'
        '</mc:AlternateContent>'
    )
    el = parse_xml(xml)
    sld = slide._element
    clr = sld.find("{%s}clrMapOvr" % PRES_NS)
    if clr is not None:
        clr.addnext(el)
    else:
        sld.find("{%s}cSld" % PRES_NS).addnext(el)


# ---------------------------------------------------------------------------
# Einflug-Animationen (Fly In), Klick oder automatisch nach Folienstart
# ---------------------------------------------------------------------------
class TimingBuilder:
    def __init__(self):
        self.next_id = 3
        self.groups = []  # XML-Strings der Klick-/Auto-Gruppen
        self.bld = []

    def _id(self):
        i = self.next_id
        self.next_id += 1
        return i

    def _effect(self, spid, direction, node_type, duration=700):
        subtype, axis, start = DIRS[direction]
        cid = self._id()
        sid = self._id()
        xid = self._id()
        return (
            f'<p:par><p:cTn id="{cid}" presetID="2" presetClass="entr" presetSubtype="{subtype}" '
            f'fill="hold" grpId="0" nodeType="{node_type}">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            f'<p:set><p:cBhvr><p:cTn id="{sid}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/>'
            f'</p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst>'
            f'<p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:to><p:strVal val="visible"/></p:to></p:set>'
            f'<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            f'<p:cTn id="{xid}" dur="{duration}" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            f'<p:attrNameLst><p:attrName>{axis}</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:tavLst><p:tav tm="0"><p:val><p:strVal val="{start}"/></p:val></p:tav>'
            f'<p:tav tm="100000"><p:val><p:strVal val="#{axis}"/></p:val></p:tav></p:tavLst></p:anim>'
            f'</p:childTnLst></p:cTn></p:par>'
        )

    def auto_group(self, items, stagger_ms=180):
        """items: Liste (spid, richtung). Starten automatisch beim Folienstart, gestaffelt."""
        inner = []
        for i, (spid, direction) in enumerate(items):
            node = "afterEffect"
            eff = self._effect(spid, direction, node)
            pid = self._id()
            inner.append(
                f'<p:par><p:cTn id="{pid}" fill="hold"><p:stCondLst><p:cond delay="{i * stagger_ms}"/>'
                f'</p:stCondLst><p:childTnLst>{eff}</p:childTnLst></p:cTn></p:par>'
            )
        gid = self._id()
        inner_id = self._id()
        self.groups.append(
            f'<p:par><p:cTn id="{gid}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
            f'<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
            f'<p:par><p:cTn id="{inner_id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst>{"".join(inner)}</p:childTnLst></p:cTn></p:par>'
            f'</p:childTnLst></p:cTn></p:par>'
        )

    def click_group(self, spid, direction):
        """Ein Klick = ein Element fliegt ein."""
        eff = self._effect(spid, direction, "clickEffect")
        gid = self._id()
        inner_id = self._id()
        self.groups.append(
            f'<p:par><p:cTn id="{gid}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
            f'</p:stCondLst><p:childTnLst>'
            f'<p:par><p:cTn id="{inner_id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst>{eff}</p:childTnLst></p:cTn></p:par>'
            f'</p:childTnLst></p:cTn></p:par>'
        )

    def mark_built(self, spid):
        self.bld.append(spid)

    def attach(self, slide):
        if not self.groups:
            return
        bld = "".join(f'<p:bldP spid="{s}" grpId="0" animBg="1"/>' for s in self.bld)
        xml = (
            f'<p:timing xmlns:p="{PRES_NS}"><p:tnLst><p:par>'
            f'<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
            f'<p:childTnLst>{"".join(self.groups)}</p:childTnLst></p:cTn>'
            f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
            f'</p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/>'
            f'</p:tgtEl></p:cond></p:nextCondLst></p:seq>'
            f'</p:childTnLst></p:cTn></p:par></p:tnLst>'
            + (f"<p:bldLst>{bld}</p:bldLst>" if bld else "")
            + "</p:timing>"
        )
        el = parse_xml(xml)
        sld = slide._element
        anchor = sld.find("{%s}transition" % PRES_NS)
        if anchor is None:
            # Morph liegt in mc:AlternateContent direkt nach clrMapOvr
            anchor = sld.xpath("./*[local-name()='AlternateContent']")
            anchor = anchor[0] if anchor else sld.find("{%s}clrMapOvr" % PRES_NS)
        anchor.addnext(el)


def pick_dir(rng):
    return rng.choice(["left", "right", "top", "bottom"])


# ---------------------------------------------------------------------------
# Folien
# ---------------------------------------------------------------------------
def build():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    blank = prs.slide_layouts[6]
    rng = random.Random(2026)  # deterministisch: gleiche Richtungen bei jedem Bau
    slides_timing = []

    # 1 Titel ---------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_bg(s, BG)
    add_picture_fill(s, ASSETS / "titel-hintergrund.jpg", 0, 0, SLIDE_W, SLIDE_H, "Hintergrund")
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, DARK, "Overlay", alpha=0.6)
    add_text(s, 0.9, 2.2, 11, 0.4,
             [{"text": "PRÄSENTATION · 5 MINUTEN", "size": 14, "bold": True, "color": GOLD}], "Label")
    add_text(s, 0.9, 2.65, 11.5, 1.6,
             [{"text": "MONTANA", "size": 96, "bold": True, "color": WHITE}], "Titel")
    add_text(s, 0.9, 4.3, 11, 0.6,
             [{"text": "Big Sky Country – der Schatz des Nordwestens", "size": 24, "color": "D7DEEA"}],
             "Untertitel")
    add_rect(s, 0.9, 5.0, 1.6, 0.06, GOLD, "Linie")
    add_text(s, 0.9, 6.7, 11, 0.4,
             [{"text": "Erstellt mit Arena.ai Agent · Oktober 2026", "size": 14, "color": MUTED}], "Credit")
    tb = TimingBuilder()
    items = [("Label", "left"), ("Titel", "bottom"), ("Untertitel", "right"), ("Linie", "left"), ("Credit", "bottom")]
    ids = {shp.name: shp.shape_id for shp in s.shapes}
    tb.auto_group([(ids[n], d) for n, d in items])
    for n, _ in items:
        tb.mark_built(ids[n])
    slides_timing.append((s, tb))
    set_notes(s, "Willkommen zu unserer Präsentation über Montana. Montana heißt auch Big Sky Country "
                 "und Treasure State. In den nächsten fünf Minuten schauen wir auf Lage, Geschichte, "
                 "Natur, Wirtschaft und Kultur. Zum Schluss gibt es ein kleines Quiz. "
                 "Erstellt wurde die Präsentation mit Arena.ai Agent.")

    # 2 Agenda --------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_bg(s, BG)
    chrome(s, "01 · Übersicht", "Worum es heute geht", 2)
    cards = [
        ("01", "Lage", "Wo liegt Montana, und wie groß ist es?"),
        ("02", "Geschichte", "Vom Goldrausch zum 41. Staat"),
        ("03", "Natur", "Glacier National Park und seine Gletscher"),
        ("04", "Wirtschaft", "Landwirtschaft, Bergbau und Tourismus"),
        ("05", "Kultur", "Menschen, Traditionen und Künstler"),
        ("06", "Quiz", "Vier Fragen zum Mitraten"),
    ]
    tb = TimingBuilder()
    names = []
    for i, (nr, title, desc) in enumerate(cards):
        col, row = i % 3, i // 3
        x = 0.7 + col * 4.1
        y = 2.3 + row * 2.3
        n = f"Karte{i + 1}"
        add_text(s, x, y, 3.8, 2.0,
                 [{"text": nr, "size": 26, "bold": True, "color": GOLD, "space_after": 4},
                  {"text": title, "size": 22, "bold": True, "color": WHITE, "space_after": 4},
                  {"text": desc, "size": 14, "color": MUTED}],
                 n, shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.25)
        names.append(n)
    ids = {shp.name: shp.shape_id for shp in s.shapes}
    tb.auto_group([(ids[n], pick_dir(rng)) for n in names], stagger_ms=150)
    for n in names:
        tb.mark_built(ids[n])
    slides_timing.append((s, tb))
    set_notes(s, "Hier siehst du den Ablauf. Wir starten mit der Lage, gehen dann durch Geschichte, "
                 "Natur und Wirtschaft und schauen auf die Kultur. Danach kommen Zahlen, ein Fazit, "
                 "die Quellen und zum Schluss das Quiz.")

    # 3 Lage ----------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_bg(s, BG)
    chrome(s, "02 · Lage", "Montana im Überblick", 3)
    tb = TimingBuilder()
    anim = []
    big = add_text(s, 0.7, 2.4, 4.8, 2.0,
                   [{"text": "#4", "size": 130, "bold": True, "color": GOLD}], "Grosszahl")
    cap = add_text(s, 0.7, 4.6, 4.8, 1.2,
                   [{"text": "viertgrößter US-Bundesstaat nach Fläche", "size": 22, "color": WHITE},
                    {"text": "Montana · rund 147.000 Quadratmeilen", "size": 16, "color": MUTED}],
                   "Grosszahl_Text")
    rows = [("Hauptstadt", "Helena"), ("Größte Stadt", "Billings"),
            ("Dichte", "nur ca. 3 Einwohner pro km²"), ("Nachbarn", "Kanada, ND, SD, WY, ID")]
    for i, (k, v) in enumerate(rows):
        y = 2.4 + i * 1.1
        n = f"Zeile{i + 1}"
        add_text(s, 6.2, y, 6.4, 0.9,
                 [{"text": k.upper(), "size": 11, "bold": True, "color": TEAL, "space_after": 2},
                  {"text": v, "size": 20, "bold": True, "color": WHITE}],
                 n, shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.25, anchor=MSO_ANCHOR.MIDDLE)
        anim.append(n)
    ids = {shp.name: shp.shape_id for shp in s.shapes}
    tb.auto_group([(ids["Grosszahl"], "left"), (ids["Grosszahl_Text"], "bottom")] +
                  [(ids[n], "right") for n in anim], stagger_ms=160)
    for n in ["Grosszahl", "Grosszahl_Text"] + anim:
        tb.mark_built(ids[n])
    slides_timing.append((s, tb))
    set_notes(s, "Montana ist der viertgrößte Bundesstaat der USA. Trotzdem leben dort nur etwa drei "
                 "Menschen pro Quadratkilometer. Die Hauptstadt ist Helena, die größte Stadt Billings. "
                 "Im Norden grenzt Montana an Kanada. Quelle: Wikipedia Montana, Britannica.")

    # 4 Geschichte ----------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_bg(s, BG)
    chrome(s, "03 · Geschichte", "Vom Goldrausch zum 41. Staat", 4)
    add_rect(s, 1.0, 3.85, 11.3, 0.05, MUTED, "Zeitachse")
    events = [
        ("1803", "Louisiana Purchase: das Gebiet kommt zu den USA"),
        ("1858", "Goldfund: Zuwanderung und Städte entstehen"),
        ("1864", "Montana-Territorium entsteht (26. Mai)"),
        ("1889", "Aufnahme als 41. Staat (8. November)"),
    ]
    tb = TimingBuilder()
    auto = []
    for i, (year, text) in enumerate(events):
        cx = 1.6 + i * 3.3
        dot = add_rect(s, cx - 0.16, 3.7, 0.32, 0.32, GOLD, f"Punkt{i + 1}", shape=MSO_SHAPE.OVAL)
        yr = add_text(s, cx - 1.3, 2.45, 2.6, 0.9,
                      [{"text": year, "size": 36, "bold": True, "color": GOLD, "align": PP_ALIGN.CENTER}],
                      f"Jahr{i + 1}")
        tx = add_text(s, cx - 1.45, 4.35, 2.9, 1.5,
                      [{"text": text, "size": 16, "color": WHITE, "align": PP_ALIGN.CENTER}],
                      f"Text{i + 1}")
        auto += [f"Jahr{i + 1}", f"Punkt{i + 1}", f"Text{i + 1}"]
    add_text(s, 0.7, 6.15, 11.9, 0.6,
             [{"text": "Lange vor 1803 lebten hier viele indigene Völker. Ihre Geschichte beginnt nicht mit der Landnahme.",
               "size": 14, "color": MUTED}], "Hinweis")
    auto.append("Hinweis")
    ids = {shp.name: shp.shape_id for shp in s.shapes}
    tb.auto_group([(ids[n], pick_dir(rng)) for n in auto], stagger_ms=170)
    for n in auto:
        tb.mark_built(ids[n])
    slides_timing.append((s, tb))
    set_notes(s, "1803 kaufen die USA das Gebiet im Louisiana Purchase. 1858 wird Gold gefunden, "
                 "und viele Menschen kommen. 1864 entsteht das Montana-Territorium. Am 8. November 1889 "
                 "wird Montana als 41. Staat aufgenommen. Wichtig: Vorher lebten dort lange indigene "
                 "Völker. Quelle: Britannica, History.com.")

    # 5 Natur ---------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_bg(s, BG)
    chrome(s, "04 · Natur", "Glacier National Park", 5)
    stats = [
        ("1910", "Gründung des Nationalparks"),
        ("26", "benannte Gletscher (um 1850 waren es rund 150)"),
        ("762", "Seen im Park"),
        ("3.190 m", "höchster Gipfel: Mount Cleveland"),
    ]
    tb = TimingBuilder()
    auto = []
    for i, (num, cap) in enumerate(stats):
        y = 2.2 + i * 1.12
        add_text(s, 0.7, y, 5.2, 1.0,
                 [{"text": num, "size": 30, "bold": True, "color": GOLD},
                  {"text": cap, "size": 15, "color": WHITE}],
                 f"Stat{i + 1}", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.22,
                 anchor=MSO_ANCHOR.MIDDLE)
        auto.append(f"Stat{i + 1}")
    add_picture_fill(s, ASSETS / "natur-hintergrund.jpg", 6.5, 0, 6.833, SLIDE_H, "Bild_Natur")
    auto.append("Bild_Natur")
    ids = {shp.name: shp.shape_id for shp in s.shapes}
    tb.auto_group([(ids[n], pick_dir(rng)) for n in auto], stagger_ms=170)
    for n in auto:
        if n != "Bild_Natur":
            tb.mark_built(ids[n])
    slides_timing.append((s, tb))
    set_notes(s, "Der Glacier National Park liegt im Nordwesten Montanas an der kanadischen Grenze. "
                 "Er wurde 1910 gegründet. Heute gibt es dort 26 benannte Gletscher, um 1850 waren es "
                 "noch rund 150. Die Gletscher schrumpfen durch den Klimawandel. Das Bild ist KI-generiert. "
                 "Quelle: glacier-nationalpark.com, Britannica Kids.")

    # 6 Wirtschaft ----------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_bg(s, BG)
    chrome(s, "05 · Wirtschaft", "Mehr als Rinder und Weizen", 6)
    add_picture_fill(s, ASSETS / "farm-usgs.jpg", 0.7, 2.2, 5.6, 4.3, "Bild_Farm")
    econ = [
        ("Landwirtschaft", "Rinder- und Getreideanbau, rund 28.000 Farmen"),
        ("Bergbau & Energie", "Öl, Gas, Kohle und Metalle"),
        ("Tourismus", "13,7 Mio. Besucher 2024, so viele wie nie"),
    ]
    auto = ["Bild_Farm"]
    for i, (t, d) in enumerate(econ):
        y = 2.2 + i * 1.5
        add_text(s, 6.8, y, 5.8, 1.3,
                 [{"text": t, "size": 20, "bold": True, "color": WHITE, "space_after": 4},
                  {"text": d, "size": 15, "color": MUTED}],
                 f"Karte{i + 1}", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.25,
                 anchor=MSO_ANCHOR.MIDDLE)
        auto.append(f"Karte{i + 1}")
    add_text(s, 6.8, 6.7 - 0.1, 5.8, 0.35,
             [{"text": "BIP 2025: rund 82,3 Mrd. USD (BEA)", "size": 13, "bold": True, "color": TEAL}],
             "Kennzahl")
    auto.append("Kennzahl")
    ids = {shp.name: shp.shape_id for shp in s.shapes}
    tb = TimingBuilder()
    tb.auto_group([(ids[n], pick_dir(rng)) for n in auto], stagger_ms=170)
    for n in auto:
        if n != "Bild_Farm":
            tb.mark_built(ids[n])
    slides_timing.append((s, tb))
    set_notes(s, "Die Wirtschaft ist stark von Landwirtschaft geprägt: Rinder und Getreide. Dazu kommen "
                 "Bergbau mit Öl, Gas und Kohle sowie der Tourismus. 2024 kamen 13,7 Millionen Besucher. "
                 "Das BIP lag 2025 bei rund 82,3 Milliarden Dollar. Das Farmbild stammt vom USGS und ist gemeinfrei. "
                 "Quelle: Montana Department of Labor & Industry, Wikipedia Montana.")

    # 7 Kultur --------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_bg(s, BG)
    chrome(s, "06 · Kultur", "Viele Kulturen, eine Landschaft", 7)
    kult = [
        ("11 Stammesnationen", "leben auf sieben Reservaten in Montana. Die Little Shell Band of Chippewa hat kein eigenes Land.", GOLD),
        ("Crow Fair seit 1909", "eines der größten indigenen Treffen Nordamerikas, mit rund 45.000 Besuchern und Teilnehmern.", TEAL),
        ("Charles M. Russell", "„the cowboy artist“ schuf über 2.000 Gemälde über Cowboys, Indigene und Landschaften.", GOLD),
    ]
    auto = []
    for i, (t, d, c) in enumerate(kult):
        y = 2.2 + i * 1.55
        add_text(s, 0.7, y, 11.9, 1.3,
                 [{"text": t, "size": 22, "bold": True, "color": WHITE, "space_after": 4},
                  {"text": d, "size": 16, "color": MUTED}],
                 f"Karte{i + 1}", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.35,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, 0.7, y + 0.25, 0.08, 0.8, c, f"Akzent{i + 1}")
        auto += [f"Karte{i + 1}", f"Akzent{i + 1}"]
    ids = {shp.name: shp.shape_id for shp in s.shapes}
    tb = TimingBuilder()
    tb.auto_group([(ids[n], pick_dir(rng)) for n in auto], stagger_ms=170)
    for n in auto:
        tb.mark_built(ids[n])
    slides_timing.append((s, tb))
    set_notes(s, "Montana ist Heimat von elf Stammesnationen auf sieben Reservaten. Die Little Shell "
                 "Band of Chippewa lebt im Land, hat aber kein eigenes Reservat. Die Crow Fair seit 1909 "
                 "ist eines der größten indigenen Treffen Nordamerikas. Charles M. Russell hat über 2.000 "
                 "Gemälde geschaffen. Quelle: History.com, Montana Beyond, Wikipedia.")

    # 8 Zahlen --------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_bg(s, BG)
    chrome(s, "07 · Zahlen", "Montana in Zahlen", 8)
    nums = [
        ("1.144.694", "Einwohner (Schätzung 2025)"),
        ("1,01 Mio.", "Acres Glacier National Park"),
        ("13,7 Mio.", "Besucher im Jahr 2024"),
        ("41.", "Bundesstaat seit 8. November 1889"),
    ]
    auto = []
    for i, (num, cap) in enumerate(nums):
        x = 0.7 + i * 3.05
        add_text(s, x, 2.6, 2.8, 3.3,
                 [{"text": num, "size": 34, "bold": True, "color": GOLD, "space_after": 10},
                  {"text": cap, "size": 16, "color": WHITE}],
                 f"Zahl{i + 1}", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.3,
                 anchor=MSO_ANCHOR.MIDDLE)
        auto.append(f"Zahl{i + 1}")
    ids = {shp.name: shp.shape_id for shp in s.shapes}
    tb = TimingBuilder()
    tb.auto_group([(ids[n], pick_dir(rng)) for n in auto], stagger_ms=170)
    for n in auto:
        tb.mark_built(ids[n])
    slides_timing.append((s, tb))
    set_notes(s, "Ein paar Zahlen zum Mitnehmen: rund 1,14 Millionen Einwohner laut Schätzung 2025, "
                 "rund 1,01 Millionen Acres Nationalpark, 13,7 Millionen Besucher 2024 und seit 1889 "
                 "der 41. Bundesstaat. Quelle: Wikipedia, glacier-nationalpark.com.")

    # 9 Fazit ---------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_bg(s, BG)
    chrome(s, "08 · Fazit", "Drei Gedanken zum Mitnehmen", 9)
    pts = [
        "Montana ist riesig und dünn besiedelt. Weite prägt Leben und Wirtschaft.",
        "Geschichte und Gegenwart hängen zusammen: Gold, Land und indigene Kultur.",
        "Die Natur ist ein Schatz, und die schrumpfenden Gletscher zeigen, wie wichtig Schutz ist.",
    ]
    auto = []
    for i, p in enumerate(pts):
        y = 2.2 + i * 1.2
        add_text(s, 0.7, y, 0.9, 0.9,
                 [{"text": str(i + 1), "size": 30, "bold": True, "color": DARK, "align": PP_ALIGN.CENTER}],
                 f"Nummer{i + 1}", shape=MSO_SHAPE.OVAL, fill=GOLD, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 1.9, y, 10.7, 0.9,
                 [{"text": p, "size": 22, "color": WHITE}], f"Punkt{i + 1}", anchor=MSO_ANCHOR.MIDDLE)
        auto += [f"Nummer{i + 1}", f"Punkt{i + 1}"]
    add_rect(s, 0.7, 5.85, 11.9, 0.02, MUTED, "Trennlinie")
    add_text(s, 0.7, 6.0, 11.9, 0.7,
             [{"text": "Big Sky Country: ein Schatz, den es zu bewahren gilt.", "size": 22, "bold": True,
               "color": GOLD}], "Schlussatz")
    auto.append("Schlussatz")
    ids = {shp.name: shp.shape_id for shp in s.shapes}
    tb = TimingBuilder()
    tb.auto_group([(ids[n], pick_dir(rng)) for n in auto], stagger_ms=170)
    for n in auto:
        tb.mark_built(ids[n])
    slides_timing.append((s, tb))
    set_notes(s, "Drei Gedanken zum Schluss: Montana ist weit und dünn besiedelt. Geschichte und "
                 "Gegenwart hängen eng zusammen. Die Natur ist ein Schatz, den man schützen muss. "
                 "Big Sky Country eben.")

    # 10 Quellen ------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_bg(s, BG)
    chrome(s, "09 · Quellen", "Quellen & Bildnachweis", 10)
    left = [
        "Wikipedia: Montana – en.wikipedia.org/wiki/Montana",
        "Wikipedia: History of Montana – en.wikipedia.org/wiki/History_of_Montana",
        "Wikipedia: Native Americans in Montana – en.wikipedia.org/wiki/Native_Americans_in_Montana",
        "Britannica: Montana – britannica.com/place/Montana-state",
        "History.com: Montana – history.com/articles/montana",
    ]
    right = [
        "Glacier National Park – glacier-nationalpark.com/about",
        "Britannica Kids: Glacier National Park – kids.britannica.com/kids/article/Glacier-National-Park/489368",
        "Montana Dept. of Labor & Industry (2023) – lmi.mt.gov",
        "Montana Beyond: Native American tribes – mtbeyond.com",
        "USGS Image Gallery – usgs.gov/media/images/agricultural-field-great-plains-region-montana",
    ]
    auto = []
    add_text(s, 0.7, 2.2, 5.8, 3.6,
             [{"text": t, "size": 12, "color": WHITE, "space_after": 9} for t in left],
             "Quellen_Links", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.3)
    add_text(s, 6.8, 2.2, 5.8, 3.6,
             [{"text": t, "size": 12, "color": WHITE, "space_after": 9} for t in right],
             "Quellen_Links2", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.3)
    add_text(s, 0.7, 6.0, 11.9, 0.8,
             [{"text": "Bilder: Farmfoto USGS (gemeinfrei). Titel- und Naturbild KI-generiert mit Arena.ai.",
               "size": 13, "color": MUTED, "space_after": 2},
              {"text": "Erstellt mit Arena.ai Agent · Stand Oktober 2026", "size": 13, "color": MUTED}],
             "Bildnachweis")
    auto = ["Quellen_Links", "Quellen_Links2", "Bildnachweis"]
    ids = {shp.name: shp.shape_id for shp in s.shapes}
    tb = TimingBuilder()
    tb.auto_group([(ids[n], pick_dir(rng)) for n in auto], stagger_ms=200)
    for n in auto:
        tb.mark_built(ids[n])
    slides_timing.append((s, tb))
    set_notes(s, "Hier sind alle Quellen. Die Hauptquellen sind Wikipedia, Britannica, History.com und die "
                 "offiziellen Seiten des Nationalparks sowie der Montana-Behörden. Bilder: Das Farmfoto "
                 "stammt vom USGS. Titel- und Naturbild sind KI-generiert.")

    # 11 Quiz ---------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    set_bg(s, BG)
    chrome(s, "10 · Quiz", "Zeit für eure Fragen", 11)
    quiz = [
        ("Wann wurde Montana der 41. US-Bundesstaat?", "8. November 1889"),
        ("Wie heißt die Hauptstadt von Montana?", "Helena"),
        ("Wie viele benannte Gletscher gibt es heute im Glacier National Park?", "26"),
        ("Welchen Rang belegt Montana unter den US-Bundesstaaten nach Fläche?", "Platz 4"),
    ]
    tb = TimingBuilder()
    q_names = []
    a_names = []
    for i, (q, a) in enumerate(quiz):
        y = 2.1 + i * 1.18
        add_text(s, 0.7, y, 7.2, 1.0,
                 [{"text": f"FRAGE {i + 1}", "size": 11, "bold": True, "color": TEAL, "space_after": 2},
                  {"text": q, "size": 17, "bold": True, "color": WHITE}],
                 f"Frage{i + 1}", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=CARD, margin=0.22,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 8.2, y, 4.45, 1.0,
                 [{"text": "Lösung", "size": 11, "bold": True, "color": DARK, "space_after": 2},
                  {"text": a, "size": 20, "bold": True, "color": DARK}],
                 f"Antwort{i + 1}", shape=MSO_SHAPE.ROUNDED_RECTANGLE, fill=GOLD, margin=0.22,
                 anchor=MSO_ANCHOR.MIDDLE)
        q_names.append(f"Frage{i + 1}")
        a_names.append(f"Antwort{i + 1}")
    ids = {shp.name: shp.shape_id for shp in s.shapes}
    tb.auto_group([(ids[n], pick_dir(rng)) for n in q_names], stagger_ms=150)
    for i, n in enumerate(a_names):
        tb.click_group(ids[n], "right" if i % 2 == 0 else "left")
    for n in q_names + a_names:
        tb.mark_built(ids[n])
    slides_timing.append((s, tb))
    set_notes(s, "Zum Schluss ein kleines Quiz. Frage 1: Wann wurde Montana aufgenommen? Klick: Lösung 8. "
                 "November 1889. Frage 2: Die Hauptstadt? Helena. Frage 3: Wie viele benannte Gletscher gibt "
                 "es heute? 26. Frage 4: Welchen Rang belegt Montana nach Fläche? Platz 4. "
                 "Jeder Klick blendet die nächste Lösung ein. Danke fürs Zuhören!")

    # Übergänge und Animationen anhängen
    for slide, tb in slides_timing:
        add_morph_transition(slide)
        tb.attach(slide)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    print(f"Gespeichert: {OUT}")


if __name__ == "__main__":
    build()
