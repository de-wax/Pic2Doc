# Pic2Doc - Bild-zu-Dokument-Generator

Verwandeln Sie Bildsammlungen in professionell formatierte Word-Dokumente mit anpassbaren Bildunterschriften aus Excel-Daten.

**Version:** 0.5.0
**Autor:** René

## Features

### Kernfunktionalität
- ✅ Stapelverarbeitung von tausenden Bildern
- ✅ Lesen von Dateinamen und Bildunterschriften aus Excel-Dateien
- ✅ Intelligentes rasterbasiertes Layout (automatische optimale Anordnung)
- ✅ Mehrspaltige Bildunterschriften mit konfigurierbarem Trennzeichen
- ✅ Automatische Bildgrößenanpassung basierend auf Rasterlayout
- ✅ Kein Umbruch der Bildunterschriften - alles passt auf eine Seite

### Benutzeroberfläche
- ✅ Moderne GUI mit CustomTkinter
- ✅ Theme-Auswahl (System/Hell/Dunkel)
- ✅ Dateiauswahl-Dialoge
- ✅ Echtzeit-Fortschrittsanzeige
- ✅ Abbrechen-Funktion
- ✅ Fehleranzeige-Panel
- ✅ Warnung vor Dateiüberschreibung
- ✅ Versionsnummer in der Titelleiste
- ✅ Automatisches Speichern der Einstellungen bei jeder Änderung
- ✅ Optimiertes kompaktes Layout (700x750)

### Technisch
- ✅ Anpassbare Schriftformatierung (Familie, Größe, Fett, Kursiv, Unterstrichen)
- ✅ Konfigurierbare Seitenränder
- ✅ Testmodus für schnelle Iteration
- ✅ Dauerhafte Konfiguration
- ✅ Eigenständige macOS-Anwendung (kein Python erforderlich)
- ✅ Deutsche Benutzeroberfläche mit englischem Code

## Schnellstart

### Verwendung der Anwendung (Empfohlen)

1. **Doppelklick** auf `Pic2Doc.app` (macOS)
2. **Dateien auswählen** über die GUI:
   - Excel-Datei mit Bildbeschreibungen
   - Bildordner
   - Speicherort für Word-Dokument
3. **Einstellungen konfigurieren** (optional)
4. **Klicken Sie auf "Dokument erstellen"**

### Ausführung aus dem Quellcode

```bash
# Voraussetzungen: Python 3.8 oder höher

# 1. Virtuelle Umgebung einrichten
python3 -m venv venv
source venv/bin/activate  # Auf Mac/Linux

# 2. Abhängigkeiten installieren
pip install -r requirements.txt

# 3. GUI-Anwendung ausführen
python src/gui_main.py

# Oder CLI-Version ausführen
python src/main.py
```

## Datenformat

### Excel-Dateistruktur

Siehe [EXAMPLE_DATA.md](EXAMPLE_DATA.md) für detaillierte Beispiele.

Ihre Excel-Datei sollte enthalten:
- **Spalte A** (oder Ihre gewählte Spalte): Bilddateinamen MIT Erweiterung
- **Spalte I** (oder Ihre gewählte(n) Spalte(n)): Bildunterschrift-Text

Beispiel:

| A              | B        | I      |
|----------------|----------|--------|
| foto001.jpg    | Natur    | N-001  |
| foto002.jpg    | Urban    | U-002  |
| foto003.jpg    | Wildnis  | W-003  |

### Mehrspaltige Bildunterschriften

Kombinieren Sie mehrere Spalten für umfangreichere Bildunterschriften:
- Einzelne Spalte: `I` → "N-001"
- Mehrere Spalten: `B,I` → "Natur - N-001"
- Trennzeichen ist konfigurierbar (Standard: " - ")

### Bildordner

Erstellen Sie einen Ordner mit allen referenzierten Bildern:
```
bilder/
├── foto001.jpg
├── foto002.jpg
└── foto003.jpg
```

Unterstützte Formate: JPEG, PNG und andere Pillow-kompatible Formate.

## Projektstruktur

```
Pic2Doc/
├── src/
│   ├── gui_main.py                # GUI-Einstiegspunkt
│   ├── main.py                    # CLI-Einstiegspunkt
│   ├── gui/
│   │   └── main_window.py         # GUI-Anwendung
│   ├── core/
│   │   ├── excel_reader.py        # Excel-Dateiverarbeitung
│   │   ├── document_generator.py  # Word-Dokumenterstellung
│   │   ├── image_handler.py       # Bildoperationen
│   │   └── config_manager.py      # Konfigurationsverwaltung
│   └── utils/
│       └── constants.py            # Anwendungskonstanten
├── assets/                         # Anwendungssymbole
├── dist/                           # Gebaute Programme
│   └── Pic2Doc.app                # macOS-Anwendung
├── build.py                        # Build-Skript
├── requirements.txt                # Python-Abhängigkeiten
├── VERSION                         # Versionsnummer
├── CHANGELOG.md                    # Versionshistorie
├── EXAMPLE_DATA.md                 # Datenformat-Beispiele
└── README.md                       # Englische Dokumentation
```

## Konfiguration

Die Anwendung speichert Einstellungen in `pic2doc_config.json`:

```json
{
  "excel_file": "daten.xlsx",
  "image_folder": "bilder",
  "output_file": "ausgabe_dokument.docx",
  "filename_column": "A",
  "caption_columns": ["I"],
  "caption_separator": " - ",
  "images_per_page": 3,
  "font_name": "Arial",
  "font_size": 10,
  "test_mode": false,
  "test_image_limit": 10
}
```

Die Konfiguration wird automatisch zwischen Sitzungen gespeichert und geladen.

## Verwendung

### GUI-Modus (Standard)

1. Anwendung starten
2. **Eingabedateien**:
   - Excel-Datei mit Beschreibungen auswählen
   - Ordner mit Bildern auswählen
   - Ausgabeort wählen
3. **Konfiguration**:
   - Dateinamen-Spalte (z.B. "A")
   - Bildunterschrift-Spalten (z.B. "I" oder "B,I")
   - Bilder pro Seite (1-10)
   - Schrifteinstellungen
4. **Optional**:
   - Testmodus aktivieren (nur N Bilder verarbeiten)
   - Theme ändern (System/Hell/Dunkel)
5. **Generieren**: Auf "Dokument erstellen" klicken

### CLI-Modus

```bash
python src/main.py
# Interaktiven Eingabeaufforderungen folgen
```

## Features im Detail

### Intelligentes Raster-Layout

Die Anwendung berechnet automatisch das optimale Raster für Ihre Bilder:
- 3 Bilder → 2×2 Raster (eine leere Zelle)
- 4 Bilder → 2×2 Raster
- 5 Bilder → 3×2 Raster (eine leere Zelle)
- 6 Bilder → 3×2 Raster
- 8 Bilder → 3×3 Raster (eine leere Zelle)

Bilder werden automatisch angepasst, um perfekt in das Raster mit Bildunterschriften zu passen.

### Bildunterschriftenverwaltung

- **Kein Umbruch**: Bildunterschriften brechen nie zur nächsten Seite um
- **Konservativer Abstand**: Automatische Berechnung stellt sicher, dass alles passt
- **Mehrspaltig**: Daten aus mehreren Excel-Spalten kombinieren
- **Anpassbares Trennzeichen**: Standard " - " oder Ihre Wahl

### Fehlerbehandlung

- Warnungen vor Dateiüberschreibung
- Erkennung fehlender Bilder
- Fehler-Panel mit detaillierten Nachrichten
- Bis zu 10 Fehler mit spezifischen Gründen angezeigt

## Fehlerbehebung

### Anwendung startet nicht
- **macOS**: Rechtsklick → Öffnen (nur beim ersten Mal, wegen Sicherheit)
- **Zugriff verweigert**: `chmod +x dist/Pic2Doc.app/Contents/MacOS/Pic2Doc`

### "Excel-Datei nicht gefunden"
- Stellen Sie sicher, dass der Dateipfad korrekt ist
- Datei muss im `.xlsx`-Format sein

### "Bilder nicht gefunden"
- Überprüfen Sie, ob Dateinamen in Excel genau mit tatsächlichen Dateien übereinstimmen (Groß-/Kleinschreibung beachten)
- Dateiendungen in Excel einschließen (z.B. "foto.jpg", nicht "foto")
- Prüfen Sie, ob Bilder im ausgewählten Ordner existieren

### Bildunterschriften erscheinen nicht
- Überprüfen Sie, ob der Bildunterschrift-Spaltenbuchstabe korrekt ist
- Prüfen Sie, ob Zellen Daten enthalten
- Zeile 1 wird als Überschrift behandelt und übersprungen

### Bilder passen nicht richtig
- Versuchen Sie, die Anzahl der Bilder pro Seite zu reduzieren
- Verwenden Sie den Testmodus, um mit dem Layout zu experimentieren
- Prüfen Sie die Bildauflösung (sehr große Bilder können Probleme verursachen)

## Entwicklung

### Anforderungen

- Python 3.8 oder höher
- Siehe `requirements.txt` für Abhängigkeiten

### Build aus Quellcode

```bash
# 1. Virtuelle Umgebung aktivieren
source venv/bin/activate

# 2. Build-Abhängigkeiten installieren
pip install -r requirements.txt

# 3. Build-Skript ausführen
python build.py

# 4. Anwendung erstellt unter dist/Pic2Doc.app
open dist/Pic2Doc.app
```

## Versionshistorie

Siehe [CHANGELOG.md](CHANGELOG.md) für detaillierte Versionshistorie.

### Aktuelle Releases

- **v0.5.0** - Optimiertes GUI-Layout, Version in Titelleiste, kompaktes Design
- **v0.4.9** - Einstellungen-Persistierung in gebündelter App, verbesserte Versions-Sichtbarkeit
- **v0.4.8** - Verbesserungen bei der Versionsnummer-Positionierung
- **v0.4.7** - VERSION-Datei-Bündelung für .app-Paket
- **v0.4.6** - Versionsnummer-Anzeige in der Fußzeile
- **v0.4.5** - Behobene Probleme mit Einstellungen-Persistierung
- **v0.4.0** - GUI mit CustomTkinter, Theme-Unterstützung, erweiterte Fehlerbehandlung
- **v0.3.0** - Rasterbasiertes Layout, mehrspaltige Bildunterschriften, kein Umbruch
- **v0.2.0** - Testmodus, intelligentes Layout, konfigurierbare Ränder
- **v0.1.0** - Erste CLI-Version

## Lizenz

Frei für private und kommerzielle Nutzung.

Siehe [LICENSE](LICENSE) für Details.

## Kontakt

Erstellt von René

Für englische Dokumentation siehe [README.md](README.md)

## Danksagungen

Erstellt mit:
- [python-docx](https://python-docx.readthedocs.io/) - Word-Dokumentgenerierung
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel-Dateilesen
- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - Modernes GUI-Framework
- [Pillow](https://pillow.readthedocs.io/) - Bildverarbeitung
