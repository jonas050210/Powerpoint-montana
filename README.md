# Powerpoint-montana

Präsentation über Montana (ca. 5 Minuten, 11 Folien, Deutsch).

## Inhalt
- `praesentation/Montana.pptx` – die fertige Präsentation
- `scripts/build_presentation.py` – erzeugt die Präsentation neu (Morph-Übergänge, Einflug-Animationen, Quiz-Klick-Animationen, Sprechernotizen)
- `docs/recherche-montana.md` – Fakten mit Quellen
- `docs/bildquellen-montana.md` – Bildquellen und Lizenzstatus
- `assets/images/` – verwendete Bilder (Farmfoto USGS, gemeinfrei; Titel- und Naturbild KI-generiert)

## Neu bauen
```bash
pip install python-pptx pillow lxml
python scripts/build_presentation.py
```

## Hinweise
- Der Morph-Übergang benötigt PowerPoint 2019 oder Microsoft 365. LibreOffice zeigt stattdessen den Fallback (Überblenden).
- Quiz-Folie: Die vier Lösungen fliegen per Klick einzeln ein.
- Quellen stehen auf Folie 10, Sprechernotizen auf jeder Folie.
- Erstellt mit Arena.ai Agent.

## Auftakt (Folien 1–3)
- Folie 1: weißer Block mit schwarzen Balken am Rand
- Folie 2: Block wird USA-Karte, Balken werden Staatsgrenzen, dann fliegen 48 Staatsnamen ein
- Folie 3: Zoom auf Montana, orange, darunter „Präsentation von Jonas“
- Geodaten: `assets/geo/us-states.json` (PublicaMundi/MappingAPI, MIT)
