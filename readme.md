# 📈 Verbrauchs-Tracker

Ein minimalistischer, blitzschneller und komplett lokaler Tracker für Strom, Gas und Wasser. Entwickelt für die einfache Erfassung und Auswertung der heimischen Zählerstände – ganz ohne Cloud-Zwang, externe Datenbanken oder unnötigen Overhead.

## ✨ Features

* **3-in-1 Erfassung:** Behalte Strom (kWh), Gas (m³) und Wasser (m³) an einem Ort im Blick.
* **100% Lokal (Privacy First):** Alle Daten werden ausschließlich im `localStorage` deines Browsers gespeichert. Es werden keine Daten an fremde Server gesendet.
* **Interaktive Auswertung:** Filterung der Daten nach Zeitraum (Gesamt, Letzte 365 Tage, Letzte 30 Tage) inklusive dynamischer SVG-Diagramme und Errechnung des Tagesdurchschnitts.
* **Dark/Light Mode:** Passt sich automatisch den Systemeinstellungen an oder kann manuell per Klick umgeschaltet werden.
* **Import & Export:** Sichere deine Daten als lokales JSON-Backup oder exportiere sie für tiefergehende Analysen bequem als CSV- oder Excel-Datei.
* **PWA-Ready:** Dank integriertem Manifest lässt sich der Tracker wie eine native Web-App auf dem Smartphone-Homescreen installieren.

## 🚀 Nutzung / Installation

Da das gesamte Tool aus einer einzigen, autarken HTML-Datei besteht, ist keine komplexe Installation nötig:

1. Lade die Datei `index.html` herunter.
2. Öffne sie in einem beliebigen modernen Webbrowser.
3. Beginne damit, deine Zählerstände einzutragen!

*(Tipp für Self-Hoster: Du kannst die `index.html` natürlich auch einfach auf einen eigenen Webserver ablegen, z. B. über einen leichtgewichtigen Nginx in einem LXC-Container, um sie im gesamten Heimnetzwerk aufzurufen).*

## 🛠️ Technologien

* **HTML5, CSS3 & Vanilla JavaScript**
* **Zero Dependencies:** Es werden keine großen Frameworks wie React oder Vue benötigt. (Der Excel-Export lädt bei Bedarf lediglich eine kleine `xlsx`-Bibliothek nach).

## 🤝 Mitwirken

Feedback, Bug-Reports oder Pull Requests zur Erweiterung der Funktionen sind jederzeit willkommen!

## 📄 Lizenz

Dieses Projekt ist lizenziert unter der [MIT Lizenz](LICENSE).

---
**Autor:** Eugen Himmelsbach
