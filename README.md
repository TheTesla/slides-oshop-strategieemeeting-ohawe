# OHAWE — Überblick für das oSHOP-Strategiemeeting

Quarto-/reveal.js-Folien (Deutsch) zum Werkstattprojekt **OHAWE** (Open Source Hardware im
Handwerk, T!Raum OSHOP). Zielgruppe: Wissenschaftler:innen der oSHOP-Lenkungsbereiche
Infrastruktur und Transfer. Kein Technikvortrag — Problem, Idee, Anknüpfungspunkte, Stand.

- Folien: `ohawe-strategie.qmd` (Sprechnotizen in `::: {.notes}`)
- Layout: `htwd-theme.css` (HTWD-Corporate-Design aus dem PPT-Master) + `ohawe.css`
- Titelgrafik: `img/hero.svg` (eigene Zeichnung, keine Fotorechte zu klären)
- Logos: HTWD aus dem PPT-Master, OSHOP von <https://www.oshop-network.de/> (Titelfolie,
  Footer, Abschlussfolie)

## Rendern

```bash
quarto render ohawe-strategie.qmd --to revealjs
```

Ergebnis: `ohawe-strategie.html` (einzelne Datei, Ressourcen eingebettet).

PDF-Export (reveal-Druckmodus):

```bash
quarto render ohawe-strategie.qmd --to revealjs
chrome --headless=new --no-pdf-header-footer --window-size=1280,720 \
  --print-to-pdf=ohawe-strategie.pdf "file:///<pfad>/ohawe-strategie.html?print-pdf"
```

## Inhalt

13 Folien, ca. 10 Minuten + Diskussion. Aufbau: Ausgangslage im Handwerk → zwei reale
Probleme der Partnerbetriebe → OSH als Antwort → Demonstratoren → Methode →
Anknüpfungspunkte für Infrastruktur und Transfer → Projektstand → Fragen an das Plenum.
Backup-Folie mit Partnern und Rollen.

## Qualitätssicherung

`_probe.html` misst im Browser je Folie Überlauf (Höhe/Breite). Bei Layoutänderungen
`ohawe-strategie.html` rendern, `_probe.html` über einen lokalen Server öffnen und die
JSON-Ausgabe prüfen (`overY`, `overX`, `overBottom` müssen 0 sein).

## Förderung

T!Raum OSHOP, Fördermaßnahme „T!Raum — TransferRäume für die Zukunft von Regionen“ des
BMBF im Programm „Innovation & Strukturwandel“. Laufzeit 03/2026–02/2029.
