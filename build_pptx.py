"""OHAWE-Strategiedeck als PPTX auf dem HTWD-Master.

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

SRC = "master oshop.pptx"
OUT = "ohawe-strategie.pptx"
IMG = Path("img")

ORANGE = RGBColor(0xEC, 0x66, 0x08)
BLACK = RGBColor(0x00, 0x00, 0x00)
GREY = RGBColor(0x55, 0x55, 0x55)
LIGHT = RGBColor(0xF6, 0xF6, 0xF6)
PAPER = RGBColor(0xFF, 0xF6, 0xEE)
HAIR = RGBColor(0xE6, 0xE6, 0xE6)

FONT = "LM Mono 10"
FONT_TITLE = "LM Mono Caps 10"

W, H = 12192000, 6858000
ML, CW = 538163, 11117262
TITLE_Y, BODY_Y = 376211, 1796400
FOOT_Y = 6539706
NUM_X = 11116800

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


def text(x, y, w, h, paras, size=15, color=BLACK, bold=False, font=FONT, space=6, align=PP_ALIGN.LEFT, bullet=False):
    box = prs.slides[-1].shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    first = True
    for item in paras:
        if isinstance(item, tuple):
            txt = item[0]
            psize = item[1] if len(item) > 1 else None
            pbold = item[2] if len(item) > 2 else None
            pcolor = item[3] if len(item) > 3 else None
            pbullet = item[4] if len(item) > 4 else bullet
        else:
            txt, psize, pbold, pcolor = item, None, None, None
            pbullet = bullet
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(space)
        if not pbullet:                      # Titel, Fußzeilen, Aussagen: kein Aufzählungsstrich
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", "0")
            pPr.set("indent", "0")
            etree.SubElement(pPr, qn("a:buNone"))
        r = p.add_run()
        r.text = txt
        r.font.size = Pt(psize or size)
        r.font.bold = bold if pbold is None else pbold
        r.font.color.rgb = pcolor or color
        r.font.name = font
    return box


def title(txt, size=None):
    if size is None:
        size = 30 if len(txt) > 40 else 34
    text(ML, TITLE_Y, CW - 300000, 950000, [txt], size=size, bold=True, font=FONT_TITLE)


def bar(x, y, w, h=52000, color=ORANGE):
    shp = prs.slides[-1].shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def card(x, y, w, h, fill=LIGHT):
    shp = prs.slides[-1].shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    bar(x, y, w)
    return shp


def quote(paras, y=4820000, x=ML, w=CW, size=15):
    h = 120000 * len(paras) + 260000
    box = prs.slides[-1].shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(h))
    box.fill.solid()
    box.fill.fore_color.rgb = PAPER
    box.line.fill.background()
    box.shadow.inherit = False
    bar(x, y, 46000, h)
    text(x + 180000, y + 110000, w - 320000, h - 180000, paras, size=size, bold=True, space=2)


def footer(number, date="November 2026", line="OHAWE · oSHOP-Strategiemeeting · Stefan Helmert"):
    text(1281600, FOOT_Y, 756000, 180000, [date], size=10, color=GREY)
    text(2126370, FOOT_Y, 8600000, 180000, [line], size=10, color=GREY)
    text(NUM_X, FOOT_Y, 539750, 180000, [str(number)], size=10, color=GREY)


n = 0
# 1 — Titel
s = slide()
n += 1
s.shapes.add_picture(str(IMG / "hero.png"), Emu(3240012), Emu(0), Emu(8951989), Emu(6858000))
text(ML, 2100000, 2700000, 1200000, ["OHAWE"], size=52, bold=True, font=FONT_TITLE)
text(ML, 3450000, 2600000, 1500000, ["Open Source Hardware im Handwerk:", "offene Baupläne statt geschlossener Maschinen"], size=15)
text(ML, 4900000, 2600000, 900000, [
    ("Werkstattprojekt OHAWE im T!Raum OSHOP", 13, False, GREY),
    ("Stefan Helmert · HTW Dresden", 13, False, GREY),
], size=13)
text(1281600, FOOT_Y, 3000000, 180000, ["oSHOP-Strategiemeeting · November 2026"], size=10, color=GREY)

# 2 — Hightech im Handwerk
n += 1
s = slide()
title("Hightech im Handwerk")
text(ML, BODY_Y - 420000, CW, 380000, ["Nicht nur Hammer und Feile: Betriebe setzen moderne Fertigung ein — und zwar längst."], size=16)
cases = [
    ("Zahntechnik", "Ein Dentallabor fertigt 80 % digital und druckt rund 300 Modelle im Monat.", "Quelle: dental.formlabs.com, 2024"),
    ("Tischlerei", "Tischlerei Müller rüstet 2025 auf 5-Achs-CNC um, für Einzelteil-Komplettbearbeitung.", "Quelle: bm-online.de, 2025"),
    ("Metallbau", "Leipziger Handwerksbetrieb investiert 2022 in ein 5-Achs-Bearbeitungszentrum.", "Quelle: aischmann.de"),
]
cw, gap = 3500000, 300000
for i, (kopf, txt, src) in enumerate(cases):
    x = ML + i * (cw + gap)
    card(x, BODY_Y + 240000, cw, 2100000)
    text(x + 150000, BODY_Y + 400000, cw - 300000, 1700000, [
        (kopf.upper(), 11, True, ORANGE, False), (txt, 15, False, BLACK, True), (src, 10, False, GREY, False),
    ], size=14, space=8)
quote(["Für kleinere Betriebe oft unattraktiv: hohe Komplexität, hohe Investitionskosten, mangelnde Flexibilität, fehlendes Fachwissen, geringe Stückzahlen."], y=4820000)
footer(n)

# 3 — Automatisierungskrise
n += 1
s = slide()
title("Das Handwerk steckt in der Automatisierungskrise")
stats = [("56 %", "der Betriebe nennen Fachkräftemangel als größte Herausforderung"),
         ("9,1 %", "weniger Tischlereibetriebe im Kammerbezirk Dresden seit 2020"),
         ("76 %", "investieren weniger als 5.000 € pro Jahr in Hard- und Software"),
         ("15 %", "sehen steigenden Automatisierungs- und Digitalisierungsdruck")]
sw, sgap = 2560000, 290000
for i, (zahl, txt) in enumerate(stats):
    x = ML + i * (sw + sgap)
    bar(x, BODY_Y + 100000, 46000, 1250000)
    text(x + 160000, BODY_Y + 90000, sw, 700000, [zahl], size=38, bold=True, color=ORANGE, font=FONT_TITLE)
    text(x + 160000, BODY_Y + 780000, sw, 900000, [txt], size=15)
quote(["Das Handwerk braucht unsere Unterstützung."], y=4820000)
footer(n)

# 4 — Die Not machte erfinderisch
n += 1
s = slide()
title("Die Not machte erfinderisch")
rows = [("1886 / 1892", "Rohrzange und verstellbarer Schraubenschlüssel: Der Mechaniker Johan Petter Johansson baut beides in seiner Werkstatt, weil jede Mutter ein eigenes Werkzeug verlangte."),
        ("1958", "Spreizdübel: Der Schlosser Artur Fischer feilt ihn in seiner Werkstatt, weil Schrauben im Mauerwerk nicht halten (Patent DE 1 097 117)."),
        ("1958", "Kabelbinder: entsteht in der Elektromontage, statt Binden mit Band und Draht.")]
y = BODY_Y + 40000
for jahr, txt in rows:
    text(ML, y, 1400000, 620000, [jahr], size=16, bold=True, color=ORANGE, font=FONT_TITLE)
    text(ML + 1500000, y, CW - 1500000, 620000, [txt], size=15, bullet=True)
    bar(ML, y + 700000, CW, 12000, HAIR)
    y += 850000
quote(["Selbst bauen und selbst entwickeln war kein reiner Notbehelf — es war Selbstverständnis."], y=4820000)
footer(n)

# 5 — Werkzeuge bauen, Vorlagen nachbauen
n += 1
s = slide()
title("Werkzeuge bauen, Vorlagen nachbauen")
colw = 5200000
hist = [("~1240", "Bauhüttenbuch des Villard de Honnecourt, zwischen Bauhütten weitergegeben. Skizzenbuch, kein Bauplan."),
        ("1588", "Ramelli veröffentlicht 195 Maschinenentwürfe zum Nachbau. Kupferstiche ohne Maße."),
        ("1731", "Schübler schreibt Zimmermannskunst „denen Zimmer-Leuten zum besten“.")]
y = BODY_Y + 40000
for jahr, txt in hist:
    text(ML, y, 950000, 560000, [jahr], size=15, bold=True, color=ORANGE, font=FONT_TITLE)
    text(ML + 1000000, y, colw - 1000000, 560000, [txt], size=14, bullet=True)
    bar(ML, y + 610000, colw, 12000, HAIR)
    y += 760000
text(ML, y + 60000, colw, 700000, ["Open Source Hardware hat eine lange Tradition. Was fehlte: Maße, Rückfluss, Lizenz."], size=15, bold=True, color=ORANGE)
px = ML + colw + 400000
s.shapes.add_picture(str(IMG / "tools_top.jpg"), Emu(px), Emu(BODY_Y + 40000), Emu(W - px - ML), Emu(3800000))
text(px, BODY_Y + 3950000, W - px - ML, 400000, ["Handwerker bei der Arbeit mit eigenem Werkzeug · Tafel aus „Arts et Métiers“, 19. Jh. · NYPL, gemeinfrei"], size=10, color=GREY)
footer(n)

# 6 — Ziele
n += 1
s = slide()
title("Ziele")
text(ML, BODY_Y - 420000, CW, 460000, ["Ziel ist OSH-Transferkompetenz im Handwerk: Betriebe sollen selbst finden, prüfen, nachbauen, anpassen und veröffentlichen können."], size=16, bold=True)
ziele = ["Lock-In-Effekte erkennen und vermeiden",
         "Anwendung von OSH-Entwurfsmustern und -Paradigmen, um Barrieren zu senken",
         "eigenständige Recherche geeigneter OSH-Teillösungen",
         "Nachbau und Anpassung an die eigene Werkstatt",
         "Veröffentlichung der eigenen Lösungen unter offener Lizenz",
         "Kooperation mit anderen Handwerksbetrieben, Hochschulen und OSH-Community"]
y = BODY_Y + 320000
for z in ziele:
    text(ML, y, CW, 480000, [z], size=16, bullet=True)
    y += 460000
quote(["Nicht drei Jahre Sondermaschine — sondern ein nachbaubarer Baukasten und die Fähigkeit, ihn zu nutzen."], y=4980000)
footer(n)

# 7 — Zwei Betriebe
n += 1
s = slide()
title("Zwei Betriebe, Synergieeffekte nutzen")
colw2 = 5350000
left = [("Fensterbau König", True),
        ("Silikonfuge an Holzfenstern wird von Hand gezogen", False),
        ("monoton, fehleranfällig, jede Form braucht eine neue Handbewegung", False),
        ("Marktlösungen sind zu unflexibel und zu teuer", False)]
right = [("Restauration Pohl", True),
         ("alte Lackschichten werden von Hand abgetragen", False),
         ("das Grundmaterial darf dabei nicht verletzt werden", False),
         ("Staub und Dämpfe belasten, Erfahrungswissen geht in Rente", False)]
for k, col in enumerate((left, right)):
    x = ML + k * (colw2 + 400000)
    y = BODY_Y + 40000
    for txt, bold in col:
        text(x, y, colw2, 560000, [txt], size=17 if bold else 15, bold=bold,
             color=ORANGE if bold else BLACK, bullet=not bold)
        y += 560000 if bold else 440000
    box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y + 140000), Emu(colw2), Emu(1150000))
    box.fill.solid(); box.fill.fore_color.rgb = PAPER
    box.line.fill.background(); box.shadow.inherit = False
    bar(x, y + 140000, 46000, 1150000)
    antwort = ("Unsere Antwort: ein Werkzeug, das Dichtstoff dosiert aufträgt und abzieht." if k == 0
               else "Unsere Antwort: eine CNC-Maschine, die automatisch vermisst und den Lack mit Laser entfernt.")
    text(x + 180000, y + 260000, colw2 - 320000, 880000, [antwort], size=14, bold=True, space=2)
footer(n)

# 8 — Zwischenstand Software
n += 1
s = slide()
title("Zwischenstand — Fokus auf Software")
text(ML, BODY_Y - 420000, CW, 380000, ["Die Maschine muss den Rahmen finden, bevor sie fährt: Wir erproben offene Segmentierungsmodelle."], size=16)
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
colw3 = 5000000
punkte = ["geschlossene Sicherheitszelle mit getöntem Laserschutzglas",
          "direkte Absaugung am Bearbeitungskopf, mehrstufige Filterung",
          "Lackschicht abtragen, ohne das Holz zu verletzen",
          "Bedienung über ein Panel, Werkstück wird vorher vermessen"]
y = BODY_Y + 40000
for p_ in punkte:
    text(ML, y, colw3, 480000, [p_], size=15, bullet=True)
    y += 400000
px = ML + colw3 + 380000
s.shapes.add_picture(str(IMG / "Lasergestützte Fensterrahmen-Aufbereitung.png"), Emu(px), Emu(BODY_Y + 40000), Emu(W - px - ML), Emu(3800000))
text(px, BODY_Y + 3950000, W - px - ML, 400000, ["Konzeptdarstellung der Anlage — Stand vor dem Aufbau"], size=10, color=GREY)
quote(["Kein Schleifstaub, keine Handarbeit über Stunden."], y=5060000, w=colw3)
footer(n)

# 10 — Danke
n += 1
s = slide()
title("Danke — Diskussion")
text(ML, BODY_Y + 600000, 5200000, 2100000, [
    ("Stefan Helmert · Professur Informationsmanagement", 17, True, ORANGE),
    ("Fakultät Informatik / Mathematik, HTW Dresden", 15, False, BLACK),
    ("stefan.helmert@htw-dresden.de", 15, False, BLACK),
    ("Projekt OHAWE · T!Raum OSHOP · Laufzeit 03/2026–02/2029", 13, False, GREY),
], size=15, space=10)
text(ML, 4700000, 8000000, 500000, ["Ein Werkstattprojekt im T!Raum OSHOP. Gefördert im Rahmen der Fördermaßnahme „T!Raum — TransferRäume für die Zukunft von Regionen“ des BMBF im Programm „Innovation & Strukturwandel“. Ergebnisse werden unter offener Lizenz veröffentlicht."], size=10, color=GREY)
footer(n)

# 11 — Backup Tabelle
n += 1
s = slide()
title("Backup: Partner und Rollen")
rows = [("Partner", "Rolle im Projekt"),
        ("HTW Dresden, Professur Informationsmanagement", "Koordination, FuE, Aufbau und Erprobung der Demonstratoren, Dokumentation, Zertifizierung"),
        ("Tischlerei Pohl, Ebersbach-Neugersdorf", "Anforderungen an Lackentfernung, Versuchsergebnisse/Parameter Laser-Lackentfernung, Gefährdungsbeurteilung"),
        ("Tischlerei König, Rammenau", "Fensterbau, stellt die vorhandene Maschine, Anforderungsmanagement, Erprobung (assoziiert)"),
        ("Handwerkskammer Dresden", "Multiplikator in ~21.000 Betriebe, gemeinsame Veranstaltungen mit OSHOP")]
shape = s.shapes.add_table(len(rows), 2, Emu(ML), Emu(BODY_Y), Emu(CW), Emu(2500000))
tbl = shape.table
tbl.columns[0].width = Emu(4300000)
tbl.columns[1].width = Emu(6800000)
for ri, (a, b) in enumerate(rows):
    for ci, val in enumerate((a, b)):
        cell = tbl.cell(ri, ci)
        cell.text = val
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x11, 0x11, 0x11) if ri == 0 else RGBColor(0xFF, 0xFF, 0xFF)
        for p_ in cell.text_frame.paragraphs:
            p_.font.size = Pt(12)
            p_.font.bold = ri == 0
            p_.font.name = FONT
            p_.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF) if ri == 0 else BLACK
footer(n)

assert len(prs.slides._sldIdLst) == len(prs.slides), "Folienliste inkonsistent"
prs.save(OUT)
print("gespeichert:", OUT, len(prs.slides), "Folien")
