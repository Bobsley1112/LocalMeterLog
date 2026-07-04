# Verbrauchs-Tracker – Strom / Gas / Wasser

Eine plattformübergreifende Desktop-Anwendung (Windows & Linux) zur Verwaltung
von Zählerständen mit grafischer Auswertung und Excel-Export.

## Funktionen

- Vergangene Werte anzeigen – Tabelle aller Ablesungen inkl. berechnetem
  Verbrauch je Zeitraum (Strom, Gas, Wasser) und Anzahl Tage.
- Neue Werte eingeben – einfaches Formular; Datum im Format TT.MM.JJJJ,
  Zahlen mit Komma oder Punkt.
- Grafische Darstellung – Liniendiagramm je Verbrauchsart, umschaltbar.
- Moderne Oberfläche – CustomTkinter, mit Hell-/Dunkel-/System-Modus.
- Excel-Export – .xlsx mit zwei Blättern und eingebetteten Diagrammen.
- Zaehlerwechsel-Erkennung – neuer Zaehler wird als 'Wechsel' markiert.

## Installation

Voraussetzung: Python 3.9 oder neuer.

    cd verbrauchstracker
    python -m venv .venv
    # Windows:
    .venv\Scripts\activate
    # Linux/macOS:
    source .venv/bin/activate
    pip install -r requirements.txt

### Linux-Hinweis
tkinter ist bei manchen Distributionen nicht vorinstalliert:
    # Debian/Ubuntu
    sudo apt install python3-tk
    # Fedora
    sudo dnf install python3-tkinter

## Starten

    python verbrauchstracker.py

## Datenspeicherung

Alle Eingaben werden automatisch in verbrauchsdaten.json im Programmordner
gespeichert.

## Als eigenständige .exe (optional)

    pip install pyinstaller
    pyinstaller --onefile --windowed --add-data "verbrauchsdaten.json:." verbrauchstracker.py

Unter Windows den Doppelpunkt durch ein Semikolon ersetzen:
    --add-data "verbrauchsdaten.json;."

Das fertige Programm liegt anschließend im Ordner dist/.
