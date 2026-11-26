"""Builds the OHAWE strategy deck as PPTX on the HTWD master (master oshop.pptx)."""
import copy
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pathlib import Path

SRC = "master oshop.pptx"
OUT = "ohawe-strategie.pptx"
IMG = Path("img")
ORANGE = RGBColor(0xEC, 0x66, 0x08)
GREY = RGBColor(0x55, 0x55, 0x55)

prs = Presentation(SRC)
# --- drop the 34 example slides, keep master + layouts ---
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
for sld in ids:            # erst alle Verweise aus der Liste lösen ...
    sldIdLst.remove(sld)
for sld in ids:            # ... dann die Beziehungen (entfernt die Teile)
    prs.part.drop_rel(sld.rId)

L = {n: i for i, l in enumerate(prs.slide_masters[0].slide_layouts) for n in [l.name]}


def add(layout_name):
    s = prs.slides.add_slide(prs.slide_masters[0].slide_layouts[L[layout_name]])
    return s


def ph(slide, idx):
    for p in slide.placeholders:
        if p.placeholder_format.idx == idx:
            return p
    return None


def set_text(placeholder, lines, size=16, bold_first=False):
    """lines: list of (text, level, bold) tuples or plain strings."""
    tf = placeholder.text_frame
    tf.word_wrap = True
    first = True
    for item in lines:
        text, level, bold = (item, 0, False) if isinstance(item, str) else item
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.level = level
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold or (bold_first and level == 0)
        run.font.name = "Open Sans"
        run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)


def chrome(slide, number, footer="OHAWE · oSHOP-Strategiemeeting · Stefan Helmert", date="November 2026"):
    """Date, footer and slide number as text boxes at the master's coordinates."""
    for left, width, txt in ((1281600, 756000, date), (2126370, 8640000, footer), (11116800, 539750, str(number))):
        p = ph(slide, 10 if left == 1281600 and slide.slide_layout.name.startswith("Titel") else None)
        box = slide.shapes.add_textbox(Emu(left), Emu(6539706), Emu(width), Emu(180000))
        tf = box.text_frame
        tf.word_wrap = False
        r = tf.paragraphs[0].add_run()
        r.text = txt
        r.font.size = Pt(10)
        r.font.color.rgb = GREY


def accent_last(placeholder):
    """Last paragraph as orange statement line."""
    para = placeholder.text_frame.paragraphs[-1]
    for r in para.runs:
        r.font.bold = True
        r.font.color.rgb = ORANGE


n = 0
# 1 — Titelfolie mit Bild
s = add("Titelfolie mit Bild")
n += 1
pic = ph(s, 12)
if pic:
    pic.insert_picture(str(IMG / "hero.png"))
set_text(ph(s, 0), [("OHAWE", 0, True)], size=44)
set_text(ph(s, 10), [
    ("Hightech im Handwerk — offene Baupläne statt geschlossener Maschinen", 0, False),
    ("Werkstattprojekt OHAWE im T!Raum OSHOP", 1, False),
    ("Stefan Helmert · HTW Dresden", 1, False),
], size=13)
body_ph = ph(s, 10)
if body_ph:
    body_ph.left, body_ph.top = Emu(538163), Emu(4150000)
    body_ph.width, body_ph.height = Emu(2900000), Emu(1600000)
dt = s.shapes.add_textbox(Emu(538163), Emu(5894863), Emu(2200000), Emu(288000))
dr = dt.text_frame.paragraphs[0].add_run()
dr.text = "oSHOP-Strategiemeeting · November 2026"
dr.font.size = Pt(11)
dr.font.name = "Open Sans"

# 2 — Hightech im Handwerk
s = add("Titel und Inhalt")
n += 1
set_text(ph(s, 0), [("Hightech im Handwerk", 0, True)], size=30)
set_text(ph(s, 13), [
    ("Nicht nur Hammer und Feile: Handwerksbetriebe setzen moderne Fertigung ein — und zwar längst.", 0, False),
    ("Zahntechnik: Ein Dentallabor fertigt 80 % digital, 300 gedruckte Modelle im Monat (Formlabs, 2024)", 0, False),
    ("Tischlerei: Müller rüstet 2025 auf 5-Achs-CNC um, für Einzelteil-Komplettbearbeitung (bm-online.de)", 0, False),
    ("Metallbau: Leipziger Handwerksbetrieb investiert 2022 in ein 5-Achs-Bearbeitungszentrum (aischmann.de)", 0, False),
    ("Für kleine Betriebe oft unattraktiv: Komplexität, Kosten, mangelnde Flexibilität, fehlendes Wissen, geringe Stückzahlen.", 0, False),
], size=15)
accent_last(ph(s, 13))
chrome(s, n)

# 3 — Automatisierungskrise
s = add("Titel und Inhalt")
n += 1
set_text(ph(s, 0), [("Das Handwerk steckt in der Automatisierungskrise", 0, True)], size=28)
set_text(ph(s, 13), [
    ("56 % — der Betriebe nennen Fachkräftemangel als größte Herausforderung", 0, True),
    ("9,1 % weniger Tischlereibetriebe im Kammerbezirk Dresden seit 2020", 0, True),
    ("76 % — investieren weniger als 5.000 € pro Jahr in Hard- und Software", 0, True),
    ("15 % — sehen steigenden Automatisierungs- und Digitalisierungsdruck", 0, True),
    ("Das Handwerk braucht unsere Unterstützung.", 0, False),
], size=18)
accent_last(ph(s, 13))
chrome(s, n)

# 4 — Die Not machte erfinderisch
s = add("Titel und Inhalt")
n += 1
set_text(ph(s, 0), [("Die Not machte erfinderisch", 0, True)], size=30)
set_text(ph(s, 13), [
    ("1886/1892 — Rohrzange und verstellbarer Schraubenschlüssel: Mechaniker Johansson baut beides in seiner Werkstatt.", 0, False),
    ("1958 — Spreizdübel: Schlosser Artur Fischer feilt ihn in seiner Werkstatt (Patent DE 1 097 117).", 0, False),
    ("1958 — Kabelbinder: aus der Elektromontage, statt Binden mit Band und Draht.", 0, False),
    ("Selbst bauen und selbst entwickeln war kein reiner Notbehelf — es war Selbstverständnis.", 0, False),
], size=16)
accent_last(ph(s, 13))
chrome(s, n)

# 5 — Werkzeuge bauen, Vorlagen nachbauen (Text links, Bild rechts)
s = add("Titel, Text, Bild rechts")
n += 1
set_text(ph(s, 0), [("Werkzeuge bauen, Vorlagen nachbauen", 0, True)], size=28)
set_text(ph(s, 13), [
    ("~1240 — Bauhüttenbuch des Villard de Honnecourt, zwischen Bauhütten weitergegeben. Skizzenbuch, kein Bauplan.", 0, False),
    ("1588 — Ramelli veröffentlicht 195 Maschinenentwürfe zum Nachbau. Kupferstiche ohne Maße.", 0, False),
    ("1731 — Schübler schreibt Zimmermannskunst „denen Zimmer-Leuten zum besten“.", 0, False),
    ("Open Source Hardware hat eine lange Tradition. Was fehlte: Maße, Rückfluss, Lizenz.", 0, False),
], size=15)
accent_last(ph(s, 13))
p = ph(s, 15)
if p:
    p.insert_picture(str(IMG / "tools_top.jpg"))
c = ph(s, 16)
if c:
    c.text_frame.text = "Handwerker bei der Arbeit mit eigenem Werkzeug · Tafel aus „Arts et Métiers“, 19. Jh. · NYPL, gemeinfrei"

# 6 — Ziele
s = add("Titel und Inhalt")
n += 1
set_text(ph(s, 0), [("Ziele", 0, True)], size=30)
set_text(ph(s, 13), [
    ("Ziel ist OSH-Transferkompetenz im Handwerk: Betriebe sollen selbst finden, prüfen, nachbauen, anpassen und veröffentlichen können.", 0, False),
    ("Lock-In-Effekte erkennen und vermeiden", 0, False),
    ("Anwendung von OSH-Entwurfsmustern und -Paradigmen, um Barrieren zu senken", 0, False),
    ("eigenständige Recherche geeigneter OSH-Teillösungen", 0, False),
    ("Nachbau und Anpassung an die eigene Werkstatt", 0, False),
    ("Veröffentlichung der eigenen Lösungen unter offener Lizenz", 0, False),
    ("Kooperation mit anderen Handwerksbetrieben, Hochschulen und OSH-Community", 0, False),
], size=16)
chrome(s, n)

# 7 — Zwei Betriebe (zweispaltig)
s = add("Titel und Inhalt zweispaltig")
n += 1
set_text(ph(s, 0), [("Zwei Betriebe, Synergieeffekte nutzen", 0, True)], size=28)
left = [p for p in s.placeholders if p.placeholder_format.idx == 13][0]
right = [p for p in s.placeholders if p.placeholder_format.idx == 14][0]
set_text(left, [
    ("Fensterbau König", 0, True),
    ("Silikonfuge an Holzfenstern wird von Hand gezogen", 0, False),
    ("monoton, fehleranfällig, jede Form braucht eine neue Handbewegung", 0, False),
    ("Marktlösungen sind zu unflexibel und zu teuer", 0, False),
    ("Unsere Antwort: Werkzeug, das Dichtstoff dosiert aufträgt und abzieht.", 0, False),
], size=14)
set_text(right, [
    ("Restauration Pohl", 0, True),
    ("alte Lackschichten werden von Hand abgetragen", 0, False),
    ("das Grundmaterial darf dabei nicht verletzt werden", 0, False),
    ("Staub und Dämpfe belasten, Erfahrungswissen geht in Rente", 0, False),
    ("Unsere Antwort: CNC-Maschine, die vermisst und den Lack mit Laser entfernt.", 0, False),
], size=14)
chrome(s, n)

# 8 — Zwischenstand Software (3 Bilder)
s = add("Titel und Inhalt")
n += 1
set_text(ph(s, 0), [("Zwischenstand — Fokus auf Software", 0, True)], size=28)
set_text(ph(s, 13), [
    ("Die Maschine muss den Rahmen finden, bevor sie fährt: Wir erproben offene Segmentierungsmodelle.", 0, False),
], size=16)
x0, wdt, gap = 538163, 3400000, 340000
top, hgt = 2300000, 2650000
for i, (f, cap) in enumerate([("win6.png", "Ausgangsaufnahme"),
                              ("win6_gdinosam2.png", "Grounding DINO + SAM 2"),
                              ("seggpt_win6_overlay.png", "SegGPT")]):
    s.shapes.add_picture(str(IMG / f), Emu(x0 + i * (wdt + gap)), Emu(top), width=Emu(wdt))
    tb = s.shapes.add_textbox(Emu(x0 + i * (wdt + gap)), Emu(top + hgt + 40000), Emu(wdt), Emu(300000))
    tf = tb.text_frame
    tf.word_wrap = True
    r = tf.paragraphs[0].add_run()
    r.text = cap
    r.font.size = Pt(11)
    r.font.name = "Open Sans"
    r.font.color.rgb = GREY
tb = s.shapes.add_textbox(Emu(538163), Emu(6000000), Emu(11117262), Emu(400000))
r = tb.text_frame.paragraphs[0].add_run()
r.text = "Noch Laborstand — aber der Schritt von der Werkzeugidee zur Maschine."
r.font.size = Pt(14)
r.font.name = "Open Sans"
r.font.bold = True
r.font.name = "Open Sans"
r.font.color.rgb = ORANGE
chrome(s, n)

# 9 — Endergebnis
s = add("Titel, Text, Bild rechts")
n += 1
set_text(ph(s, 0), [("Endergebnis (Bsp. Pohl, Laserlackentfernung)", 0, True)], size=26)
set_text(ph(s, 13), [
    ("geschlossene Sicherheitszelle mit getöntem Laserschutzglas", 0, False),
    ("direkte Absaugung am Bearbeitungskopf, mehrstufige Filterung", 0, False),
    ("Lackschicht abtragen, ohne das Holz zu verletzen", 0, False),
    ("Bedienung über ein Panel, Werkstück wird vorher vermessen", 0, False),
    ("Kein Schleifstaub, keine Handarbeit über Stunden.", 0, False),
], size=15)
accent_last(ph(s, 13))
p = ph(s, 15)
if p:
    p.insert_picture(str(IMG / "Lasergestützte Fensterrahmen-Aufbereitung.png"))
c = ph(s, 16)
if c:
    c.text_frame.text = "Konzeptdarstellung der Anlage — Stand vor dem Aufbau"

# 10 — Schlussfolie
s = add("Schlussfolie 1")
n += 1
set_text(ph(s, 0), [("Danke — Diskussion", 0, True)], size=30)
set_text(ph(s, 13), [
    ("Stefan Helmert · Professur Informationsmanagement", 0, True),
    ("Fakultät Informatik / Mathematik, HTW Dresden", 0, False),
    ("stefan.helmert@htw-dresden.de", 0, False),
    ("Projekt OHAWE · T!Raum OSHOP · Laufzeit 03/2026–02/2029", 0, False),
], size=14)

# 11 — Backup Tabelle
s = add("Tabelle")
n += 1
set_text(ph(s, 0), [("Backup: Partner und Rollen", 0, True)], size=26)
rows = [("Partner", "Rolle im Projekt"),
        ("HTW Dresden, Professur Informationsmanagement", "Koordination, FuE, Aufbau und Erprobung der Demonstratoren, Dokumentation, Zertifizierung"),
        ("Tischlerei Pohl, Ebersbach-Neugersdorf", "Anforderungen an Lackentfernung, Versuchsergebnisse/Parameter Laser-Lackentfernung, Gefährdungsbeurteilung"),
        ("Tischlerei König, Rammenau", "Fensterbau, stellt die vorhandene Maschine, Anforderungsmanagement, Erprobung (assoziiert)"),
        ("Handwerkskammer Dresden", "Multiplikator in ~21.000 Betriebe, gemeinsame Veranstaltungen mit OSHOP")]
tbl_ph = ph(s, 14)
gt = tbl_ph.insert_table(rows=len(rows), cols=2).table if tbl_ph else None
if gt:
    gt.columns[0].width = Emu(4300000)
    gt.columns[1].width = Emu(6600000)
if gt:
    for r_i, (a, b) in enumerate(rows):
        for c_i, val in enumerate((a, b)):
            cell = gt.cell(r_i, c_i)
            cell.text = val
            for para in cell.text_frame.paragraphs:
                para.font.size = Pt(11 if r_i else 11)
                para.font.name = "Open Sans"
                para.font.bold = r_i == 0
                para.font.name = "Open Sans"
chrome(s, n)

n_ids = len(prs.slides._sldIdLst)
assert n_ids == len(prs.slides), f"sldIdLst ({n_ids}) != Folien ({len(prs.slides)})"
prs.save(OUT)
print("gespeichert:", OUT, len(prs.slides.__iter__.__self__._sldIdLst), "Folien")
