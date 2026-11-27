"""OHAWE-Strategiedeck als PPTX auf dem HTWD-Master.

Inhalte, Reihenfolge und Formulierungen 1:1 aus ohawe-strategie.qmd.
Gestaltung wie die HTML-Fassung (Karten, Kennzahlen, Zitatbalken, Bildreihen),
Typografie und Farben aus dem Master. Fußzeile sitzt auf jeder Folie gleich tief.
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
from pathlib import Path
import math

SRC = "master oshop.pptx"
OUT = "ohawe-strategie.pptx"
IMG = Path("img")
LOGOS = Path("logos")

ACCENT = RGBColor(0x23, 0x7C, 0x8F)
BLACK = RGBColor(0x00, 0x00, 0x00)
GREY = RGBColor(0x55, 0x55, 0x55)
LIGHT = RGBColor(0xF6, 0xF6, 0xF6)
PAPER = RGBColor(0xEA, 0xF5, 0xF7)
BORDER = RGBColor(0xA8, 0xD2, 0xDB)
HAIR = RGBColor(0xEC, 0xEC, 0xEC)
LINE = 9525  # 1 px
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

FONT = "LM Mono 10"
FONT_TITLE = "LM Mono Caps 10"

W, H = 12192000, 6858000
ML, CW = 538163, 11117262
TITLE_Y, BODY_Y = 376211, 1796400
FOOT_Y = 6539706
NUM_X = 11116800
PT = 12700

prs = Presentation(SRC)
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
for s in ids:
    sldIdLst.remove(s)
for s in ids:
    prs.part.drop_rel(s.rId)

LAYOUT = {l.name: i for i, l in enumerate(prs.slide_masters[0].slide_layouts)}


def slide():
    return prs.slides.add_slide(prs.slide_masters[0].slide_layouts[LAYOUT["Leer"]])


def lines_for(txt, width, size):
    per = max(12, int(width / (0.62 * size * PT)))
    return max(1, math.ceil(len(txt) / per))


def text(x, y, w, h, paras, size=15, color=BLACK, bold=False, font=FONT, space=6,
         align=PP_ALIGN.LEFT, bullet=False, italic=False, spacing=None):
    """paras: Text, (Text, Groesse, fett, Farbe[, Aufzaehlung]) oder Dict."""
    box = prs.slides[-1].shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    first = True
    for item in paras:
        if isinstance(item, dict):
            runs, psize, pbold = item["runs"], item.get("size", size), item.get("bold", bold)
            pcolor, pbullet, pitalic = item.get("color", color), item.get("bullet", bullet), item.get("italic", italic)
            pspace = item.get("space", space)
        else:
            if isinstance(item, tuple):
                txt = item[0]
                psize = item[1] if len(item) > 1 and item[1] else size
                pbold = item[2] if len(item) > 2 and item[2] is not None else bold
                pcolor = item[3] if len(item) > 3 and item[3] else color
                pbullet = item[4] if len(item) > 4 else bullet
            else:
                txt, psize, pbold, pcolor, pbullet = item, size, bold, color, bullet
            runs, pitalic, pspace = [(txt, pbold, False)], italic, space
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(pspace)
        if not pbullet:
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", "0")
            pPr.set("indent", "0")
            etree.SubElement(pPr, qn("a:buNone"))
        for (txt, rbold, ritalic) in runs:
            r = p.add_run()
            r.text = txt
            r.font.size = Pt(psize)
            r.font.bold = rbold
            r.font.italic = ritalic or pitalic
            r.font.color.rgb = pcolor
            r.font.name = font
            if spacing:
                r.font._element.set("spc", str(spacing))
    return box


def title(txt, size=None):
    text(ML, TITLE_Y, CW - 300000, 950000, [txt], size=size or (30 if len(txt) > 40 else 34),
         bold=True, font=FONT_TITLE)


def eyebrow(txt):
    text(ML, BODY_Y - 780000, CW - 300000, 260000, [txt], size=12, bold=True,
         color=ACCENT, font=FONT_TITLE, spacing=160)


def bar(x, y, w, h=52000, color=ACCENT):
    shp = prs.slides[-1].shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def boxed(x, y, w, h, fill=LIGHT, edge=None, line=12000):
    shp = prs.slides[-1].shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if edge:
        shp.line.color.rgb = edge
        shp.line.width = Emu(line)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def card(x, y, w, h, fill=LIGHT):
    boxed(x, y, w, h, fill)
    bar(x, y, w)
    return None


def quote(paras, y=4820000, x=ML, w=CW, size=15):
    h = sum(lines_for(p if isinstance(p, str) else p["runs"][0][0], w - 320000, size)
            for p in paras) * 215000 + 300000
    boxed(x, y, w, h, PAPER)
    bar(x, y, 46000, h)
    text(x + 180000, y + 110000, w - 320000, h - 180000, paras, size=size, bold=True, space=2)


def chip(text_, x, y, w):
    h = lines_for(text_, w - 160000, 13) * 250000 + 210000
    boxed(x, y, w, h, PAPER, BORDER)
    text(x + 80000, y + 70000, w - 160000, h - 140000, [text_], size=13, space=0)
    return h


def footer(number, date="November 2026", line="OHAWE · oSHOP-Strategiemeeting · Stefan Helmert"):
    text(1281600, FOOT_Y, 756000, 180000, [date], size=10, color=GREY)
    text(2126370, FOOT_Y, 8600000, 180000, [line], size=10, color=GREY)
    text(NUM_X, FOOT_Y, 539750, 180000, [str(number)], size=10, color=GREY)


n = 0
# 1 — Titel (wie die HTML-Vorlage: Logos, Kicker, Titel, Unterzeile, Regel, Metazeile)
s = slide()
n += 1
PIC_X = 5850240            # 48 % der Folienbreite wie .ts-photo
s.shapes.add_picture(str(IMG / "hero.png"), Emu(PIC_X), Emu(0), Emu(W - PIC_X), Emu(H))
ff = s.shapes.build_freeform(Emu(PIC_X), 0)
ff.add_line_segments([(Emu(PIC_X + 2408729), 0), (Emu(PIC_X), Emu(H)), (Emu(PIC_X), 0)], close=True)
wedge = ff.convert_to_shape()
wedge.fill.solid(); wedge.fill.fore_color.rgb = WHITE
wedge.line.fill.background(); wedge.shadow.inherit = False

s.shapes.add_picture(str(LOGOS / "htwd_word.png"), Emu(ML), Emu(376211), height=Emu(320000))
s.shapes.add_picture(str(LOGOS / "oshop-logo.png"), Emu(ML + 1250000), Emu(376211), height=Emu(320000))
text(ML, 1450000, 3700000, 300000, ["Werkstattprojekt im T!Raum OSHOP"], size=12, bold=True,
     color=ACCENT, font=FONT_TITLE, spacing=160)
text(ML, 1750000, 2950000, 900000, ["OHAWE"], size=54, bold=True, font=FONT_TITLE)
text(ML, 2700000, 3900000, 900000, ["Open Source Hardware im Handwerk —", "Selbstsouveränität der deutschen Wirtschaft"], size=14)
bar(ML, 3580000, 609600, 47625)
text(ML, 3900000, 3600000, 900000, [
    ("Stefan Helmert", 15, True, BLACK),
    ("Fakultät Informatik / Mathematik, HTW Dresden", 13, False, GREY),
], size=15, space=10)
text(1281600, FOOT_Y, 3000000, 180000, ["OSHOP-Strategiemeeting · November 2026"], size=10, color=GREY)

# 2 — Hightech im Handwerk
n += 1
s = slide()
title("Hightech im Handwerk")
eyebrow("Worum es heute geht")
text(ML, BODY_Y - 460000, CW, 400000, ["Nicht nur Hammer und Feile: Handwerksbetriebe setzen moderne Fertigung ein — und zwar längst."], size=16)
cases = [
    ("Zahntechnik", "Ein Dentallabor mit zwölf Mitarbeitenden fertigt 80 % digital und druckt rund 300 Modelle im Monat.",
     "Quelle: dental.formlabs.com, Anwenderbericht 2024"),
    ("Tischlerei", "Tischlerei Müller (Duderstadt) rüstet 2025 auf 5-Achs-CNC und Fenster-CNC um — für Einzelteil-Komplettbearbeitung.",
     "Quelle: bm-online.de, Reportage 2025"),
    ("Metallbau", "Ein Leipziger Handwerksbetrieb investiert 2022 in ein 5-Achs-Bearbeitungszentrum — für Einzel- und Kleinserien.",
     "Quelle: aischmann.de, Firmenchronik 2022"),
]
cw, gap = 3500000, 300000
karten_top = BODY_Y + 200000
karten = []
for i, (kopf, txt, src) in enumerate(cases):
    karten.append((lines_for(txt, cw - 300000, 15) * 240000
                   + lines_for(src, cw - 300000, 10) * 160000 + 620000))
kh = max(karten)
for i, (kopf, txt, src) in enumerate(cases):
    x = ML + i * (cw + gap)
    card(x, karten_top, cw, kh)
    text(x + 150000, karten_top + 160000, cw - 300000, kh - 220000, [
        (kopf.upper(), 11, True, ACCENT, False), (txt, 15, False, BLACK, True), (src, 10, False, GREY, False),
    ], size=15, space=8)
y = karten_top + kh + 320000
text(ML, y, CW, 300000, [{"runs": [("Für kleinere Handwerksbetriebe oft unattraktiv:", True, False)]}], size=15)
y2 = y + 340000
cx, cw2, cgap = ML, 1990000, 140000
chips = ["hohe Komplexität", "hohe Investitionskosten", "mangelnde Flexibilität",
         "fehlendes Fachwissen", "geringe Stückzahlen"]
hoehe = 0
for i, chip_txt in enumerate(chips):
    hoehe = max(hoehe, chip(chip_txt, cx + i * (cw2 + cgap), y2, cw2))
text(ML, y2 + hoehe + 180000, CW, 260000, ["Problem › zwei Betriebe › offene Baupläne › Maschinen › Ihre Sicht"], size=11, color=GREY)
footer(n)

# 3 — Automatisierungskrise
n += 1
s = slide()
title("Das Handwerk steckt in der Automatisierungskrise")
eyebrow("Ausgangslage")
stats = [("56 %", "der Betriebe nennen Fachkräftemangel als größte Herausforderung"),
         ("−9,1 %", "Tischlereibetriebe im Kammerbezirk Dresden seit 2020"),
         ("76 %", "investieren weniger als 5.000 € pro Jahr in Hard- und Software"),
         ("15 %", "sehen steigenden Automatisierungs- und Digitalisierungsdruck")]
sw, sgap = 2560000, 290000
for i, (zahl, txt) in enumerate(stats):
    x = ML + i * (sw + sgap)
    bar(x, BODY_Y + 100000, 46000, 1250000)
    text(x + 160000, BODY_Y + 90000, sw, 700000, [zahl], size=38, bold=True, color=ACCENT, font=FONT_TITLE)
    text(x + 160000, BODY_Y + 780000, sw, 900000, [txt], size=15)
quote(["Das Handwerk braucht unsere Unterstützung."], y=4820000)
footer(n)

# 4 — Die Not machte erfinderisch
n += 1
s = slide()
title("Die Not machte erfinderisch")
eyebrow("Erfunden in der Werkstatt, für die Werkstatt")
rows = [
    ("1886", "Rohrzange, 1892 der verstellbare Schraubenschlüssel", " — der Mechaniker Johan Petter Johansson baut beides in seiner Werkstatt, weil jede Mutter ein eigenes Werkzeug verlangte.", "später Weltmarktartikel unter dem Namen Bahco"),
    ("1958", "Spreizdübel", " — der Schlosser Artur Fischer feilt ihn in seiner Werkstatt, weil Schrauben im Mauerwerk nicht halten.", "Patent DE 1 097 117"),
    ("1958", "Kabelbinder", " — aus der Elektromontage, statt Binden mit Band und Draht.", "heute in jedem Werkzeugkoffer"),
]
y = BODY_Y + 60000
for idx, (jahr, fett, txt, kursiv) in enumerate(rows):
    text(ML, y, 1400000, 620000, [jahr], size=16, bold=True, color=ACCENT, font=FONT_TITLE)
    text(ML + 1500000, y, CW - 1500000, 620000, [{"runs": [(fett, True, False), (txt, False, False), (kursiv, False, True)]}], size=15)
    if idx < len(rows) - 1:
        bar(ML, y + 700000, CW, LINE, HAIR)
    y += 850000
quote([{"runs": [("Selbst bauen und selbst entwickeln war kein reiner Notbehelf — es war ", True, False), ("Selbstverständnis", True, True), (".", True, False)]}], y=4820000)
footer(n)

# 5 — Werkzeuge bauen, Vorlagen nachbauen
n += 1
s = slide()
title("Werkzeuge bauen, Vorlagen nachbauen")
eyebrow("Das ist keine neue Idee — nur eine vergessene")
colw = 5200000
hist = [("~1240", "Bauhüttenbuch", " des Villard de Honnecourt — Risse und Konstruktionen, zwischen Bauhütten weitergegeben.", "Skizzenbuch, kein Bauplan"),
        ("1588", "Ramelli", " veröffentlicht 195 Maschinenentwürfe — zum Nachbau, für Bauherren und Handwerker.", "Kupferstiche ohne Maße"),
        ("1731", "Schübler", " schreibt Zimmermannskunst „denen Zimmer-Leuten zum besten“ — mit Tafeln für Dächer und Hängewerke.", "Lehrbuch mit Zielgruppe Handwerk")]
y = BODY_Y + 60000
for idx, (jahr, fett, txt, kursiv) in enumerate(hist):
    text(ML, y, 950000, 560000, [jahr], size=15, bold=True, color=ACCENT, font=FONT_TITLE)
    text(ML + 1000000, y, colw - 1000000, 560000, [{"runs": [(fett, True, False), (txt, False, False), (kursiv, False, True)]}], size=14)
    if idx < len(hist) - 1:
        bar(ML, y + 610000, colw, LINE, HAIR)
    y += 760000
quote([{"runs": [("Open Source Hardware hat eine lange Tradition. Was fehlte: ", True, False), ("Maße, Rückfluss, Lizenz.", True, True)]}], y=y + 40000, w=colw)
px = ML + colw + 400000
s.shapes.add_picture(str(IMG / "tools_top.jpg"), Emu(px), Emu(BODY_Y + 60000), Emu(W - px - ML), Emu(3800000))
text(px, BODY_Y + 3950000, W - px - ML, 400000, ["Handwerker bei der Arbeit mit eigenem Werkzeug. Tafel aus „Arts et Métiers“, 19. Jahrhundert (NYPL, gemeinfrei)."], size=10, color=GREY)
footer(n)

# 6 — Ziele
n += 1
s = slide()
title("Ziele")
eyebrow("Was OHAWE erreichen will")
text(ML, BODY_Y - 460000, CW, 460000, [{"runs": [("Ziel ist ", False, False), ("OSH-Transferkompetenz im Handwerk", True, False), (": Betriebe sollen selbst finden, prüfen, nachbauen, anpassen und veröffentlichen können.", False, False)]}], size=16)
ziele = ["Lock-In-Effekte erkennen und vermeiden",
         "Anwendung von OSH-Entwurfsmustern und -Paradigmen, um Barrieren zu senken",
         "eigenständige Recherche geeigneter OSH-Teillösungen",
         "Nachbau und Anpassung an die eigene Werkstatt",
         "Veröffentlichung der eigenen Lösungen unter offener Lizenz",
         "Kooperation mit anderen Handwerksbetrieben, Hochschulen und OSH-Community"]
y = BODY_Y + 300000
for z in ziele:
    text(ML, y, CW, 480000, [z], size=16, bullet=True)
    y += 460000
quote([{"runs": [("Nicht drei Jahre Sondermaschine — sondern ein ", True, False), ("nachbaubarer Baukasten", True, True), (" und die Fähigkeit, ihn zu nutzen.", True, False)]}], y=5060000)
footer(n)

# 7 — Zwei Betriebe
n += 1
s = slide()
title("Zwei Betriebe, Synergieeffekte nutzen")
eyebrow("Und was wir dafür bauen")
colw2 = 5350000
left = ("Fensterbau König",
        ["Silikonfuge an Holzfenstern wird von Hand gezogen",
         "monoton, fehleranfällig, jede Form braucht eine neue Handbewegung",
         "Marktlösungen sind zu unflexibel und zu teuer"],
        [("Unsere Antwort:", True), (" ein Werkzeug, das Dichtstoff dosiert aufträgt und abzieht — auf Basis der vorhandenen Maschine.", False)])
right = ("Restauration Pohl",
         ["alte Lackschichten werden von Hand abgetragen",
          "das Grundmaterial darf dabei nicht verletzt werden",
          "Staub und Dämpfe belasten, Erfahrungswissen geht in Rente"],
         [("Unsere Antwort:", True), (" eine CNC-Maschine, die das Werkstück automatisch vermisst und den Lack mit Laser entfernt.", False)])
for k, (kopf, punkte, antwort) in enumerate((left, right)):
    x = ML + k * (colw2 + 400000)
    y = BODY_Y + 60000
    bar(x, y - 120000, colw2, 38100)
    text(x, y, colw2, 480000, [kopf], size=17, bold=True, color=ACCENT)
    y += 560000
    for p_ in punkte:
        text(x, y, colw2, 420000, [p_], size=15, bullet=True)
        y += 480000
    y += 80000
    ah = lines_for(" ".join(t for t, _ in antwort), colw2 - 320000, 14) * 200000 + 300000
    boxed(x, y, colw2, ah, PAPER)
    bar(x, y, 46000, ah)
    text(x + 180000, y + 130000, colw2 - 320000, ah - 200000,
         [{"runs": [(t, b, False) for t, b in antwort]}], size=14, bold=True, space=2)
text(ML, 5560000, CW, 400000, ["Beide Betriebe sind Projektpartner, nicht Anwender. Gebaut und erprobt wird im Versuchsfeld der HTW Dresden, danach gehen die Maschinen in die Betriebe."], size=10, color=GREY)
footer(n)

# 8 — Zwischenstand Software
n += 1
s = slide()
title("Zwischenstand — Fokus auf Software")
eyebrow("KI findet den Fensterrahmen")
text(ML, BODY_Y - 460000, CW, 380000, ["Die Maschine muss den Rahmen finden, bevor sie fährt: Wir erproben offene Segmentierungsmodelle."], size=16)
shots = [("win6.png", "Ausgangsaufnahme"), ("win6_gdinosam2.png", "Grounding DINO + SAM 2"), ("seggpt_win6_overlay.png", "SegGPT")]
bw, bgap = 3500000, 300000
for i, (f, cap) in enumerate(shots):
    x = ML + i * (bw + bgap)
    s.shapes.add_picture(str(IMG / f), Emu(x), Emu(BODY_Y + 240000), Emu(bw), Emu(2650000))
    text(x, BODY_Y + 2960000, bw, 280000, [cap], size=11, color=GREY)
quote(["Noch Laborstand — aber der Schritt von der Werkzeugidee zur Maschine."], y=5040000)
footer(n)

# 9 — Endergebnis
n += 1
s = slide()
title("Endergebnis (Bsp. Pohl, Laserlackentfernung)")
eyebrow("So soll der Arbeitsplatz aussehen")
colw3 = 5000000
punkte = ["geschlossene Sicherheitszelle mit getöntem Laserschutzglas",
          "direkte Absaugung am Bearbeitungskopf, mehrstufige Filterung",
          "Lackschicht abtragen, ohne das Holz zu verletzen",
          "Bedienung über ein Panel, Werkstück wird vorher vermessen"]
y = BODY_Y + 60000
bar(ML, y - 120000, colw3, 38100)
for p_ in punkte:
    text(ML, y, colw3, 480000, [p_], size=15, bullet=True)
    y += 440000
px = ML + colw3 + 380000
s.shapes.add_picture(str(IMG / "Lasergestützte Fensterrahmen-Aufbereitung.png"), Emu(px), Emu(BODY_Y + 60000), Emu(W - px - ML), Emu(3500000))
text(px, BODY_Y + 3650000, W - px - ML, 400000, ["Konzeptdarstellung der Anlage — Stand vor dem Aufbau."], size=10, color=GREY)
quote(["Kein Schleifstaub, keine Handarbeit über Stunden — und ein Arbeitsplatz, den auch Ungelernte bedienen können."], y=4800000, w=colw3)
footer(n)

# 10 — Danke
n += 1
s = slide()
title("Danke — Diskussion")
colw4 = 5350000
text(ML, BODY_Y + 200000, colw4, 2200000, [
    ("Stefan Helmert", 15, True, BLACK),
    ("Professur Informationsmanagement", 14, False, BLACK),
    ("Fakultät Informatik / Mathematik", 14, False, BLACK),
    ("HTW Dresden", 14, False, BLACK),
    ("stefan.helmert@htw-dresden.de", 14, False, BLACK),
], size=14, space=4)
x2 = ML + colw4 + 400000
text(x2, BODY_Y + 200000, colw4, 2200000, [
    ("Projekt OHAWE", 14, True, BLACK),
    ("Open Source Hardware im Handwerk", 14, False, BLACK),
    ("T!Raum OSHOP · Laufzeit 03/2026–02/2029", 14, False, BLACK),
    ("Verbund: HTW Dresden, Tischlerei Pohl", 14, False, BLACK),
    ("Assoziiert: Tischlerei König, Handwerkskammer Dresden", 14, False, BLACK),
], size=14, space=4)
s.shapes.add_picture(str(LOGOS / "htwd_word.png"), Emu(ML), Emu(4700000), height=Emu(300000))
s.shapes.add_picture(str(LOGOS / "oshop-logo.png"), Emu(ML + 1150000), Emu(4700000), height=Emu(240000))
text(ML, 5250000, CW, 500000, ["Ein Werkstattprojekt im T!Raum OSHOP. Gefördert im Rahmen der Fördermaßnahme „T!Raum — TransferRäume für die Zukunft von Regionen“ des BMBF im Programm „Innovation & Strukturwandel“. Ergebnisse werden unter offener Lizenz veröffentlicht."], size=10, color=GREY)
footer(n)

# 11 — Backup Tabelle
n += 1
s = slide()
title("Backup: Partner und Rollen")
eyebrow("Auf einen Blick")
rows = [("Partner", "Rolle im Projekt"),
        ("HTW Dresden, Professur Informationsmanagement", "Koordination, FuE, Aufbau und Erprobung der Demonstratoren, Dokumentation, Zertifizierung"),
        ("Tischlerei Pohl, Ebersbach-Neugersdorf", "Anforderungen an Lackentfernung, Versuchsergebnisse/Parameter Laser-Lackentfernung, Gefährdungsbeurteilung"),
        ("Tischlerei König, Rammenau", "Fensterbau, stellt die vorhandene Maschine, Anforderungsmanagement, Erprobung (assoziiert)"),
        ("Handwerkskammer Dresden", "Multiplikator in ~21.000 Betriebe, gemeinsame Veranstaltungen mit OSHOP")]
shape = s.shapes.add_table(len(rows), 2, Emu(ML), Emu(BODY_Y + 60000), Emu(CW), Emu(2400000))
tbl = shape.table
tbl.columns[0].width = Emu(4300000)
tbl.columns[1].width = Emu(6800000)
for ri, (a, b) in enumerate(rows):
    for ci, val in enumerate((a, b)):
        cell = tbl.cell(ri, ci)
        cell.text = val
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x11, 0x11, 0x11) if ri == 0 else WHITE
        for p_ in cell.text_frame.paragraphs:
            p_.font.size = Pt(12)
            p_.font.bold = ri == 0 or ci == 0 and ri > 0
            p_.font.name = FONT
            p_.font.color.rgb = WHITE if ri == 0 else BLACK
footer(n)

assert len(prs.slides._sldIdLst) == len(prs.slides), "Folienliste inkonsistent"
prs.save(OUT)


def normal_view(path):
    """lastView im Paket auf die normale Ansicht setzen (PowerPoint oeffnet sonst im Master)."""
    import re, shutil, zipfile
    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "ppt/viewProps.xml":
                if b"lastView=" in data:
                    data = re.sub(rb'lastView="[^"]*"', b'lastView="sldView"', data)
                else:
                    data = data.replace(b"<p:viewPr ", b'<p:viewPr lastView="sldView" ', 1)
                # ohne preferSingleView oeffnet PowerPoint die Datei in der Masteransicht
                data = data.replace(b'<p:normalViewPr horzBarState="maximized"',
                                    b'<p:normalViewPr horzBarState="maximized" preferSingleView="1"')
            zout.writestr(item, data)
    shutil.move(tmp, str(path))


normal_view(OUT)
print("gespeichert:", OUT, len(prs.slides), "Folien")
