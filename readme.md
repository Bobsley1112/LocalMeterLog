# 📈 LocalMeterLog

Ein minimalistischer, blitzschneller und datenschutzfreundlicher Tracker für heimische Energiezähler. Entwickelt für die einfache Erfassung, Visualisierung und Auswertung von Verbrauchsdaten – komplett offline und ohne Cloud-Zwang.

## ✨ Features

* **Universell & Dynamisch:** Unterstützt standardmäßig Strom (kWh), Gas (m³), Wasser (m³), Heizöl (l), Pellets (kg) und Fernwärme (kWh). Nicht benötigte Zähler lassen sich per Klick einfach ausblenden.
* **Privacy First (100% Lokal):** Alle Daten werden ausschließlich im `localStorage` deines Webbrowsers gespeichert. Es fließen keine Daten an externe Server ab.
* **Integrierter Kostenrechner:** Hinterlege einfach den aktuellen Arbeitspreis (z. B. €/kWh) für deine aktiven Energieträger, um die laufenden Kosten für jeden ausgewählten Zeitraum sofort im Blick zu haben.
* **Interaktive Auswertung:** Filterung der Daten nach individuellen Zeiträumen (Gesamt, Letzte 365 Tage, Letzte 30 Tage) inklusive dynamischer SVG-Diagramme und Errechnung des Tagesdurchschnitts.
* **Zählerwechsel-Intelligenz:** Das Tool erkennt negative Differenzen automatisch als Zählerwechsel und rechnet nahtlos und fehlerfrei weiter.
* **Export & Backups:** Sichere deine Daten als lokales JSON-Backup oder exportiere sie für tiefergehende Analysen bequem als CSV- oder Excel-Datei (.xlsx).
* **Single-File-Architektur:** Das gesamte Tool (inklusive minimalistischem SVG-Favicon und Styling) besteht aus einer einzigen, autarken `index.html`-Datei. 

## 🚀 Nutzung & Installation

Da LocalMeterLog keine Datenbank und kein aufwendiges Backend benötigt, gibt es zwei simple Wege, es zu nutzen:

### Variante A: Desktop / Smartphone (Lokal)
1. Lade die Datei `index.html` herunter.
2. Öffne sie in einem beliebigen modernen Webbrowser.
3. Lege über die Checkboxen deine aktiven Zähler fest und trage deine Stände ein.

### Variante B: Hosting im Heimnetzwerk (Self-Hosting)
LocalMeterLog lässt sich hervorragend in einer bestehenden Heimserver-Infrastruktur betreiben. Ideal ist beispielsweise die Bereitstellung in einem schlanken Proxmox LXC-Container mit einem einfachen Webserver (wie Nginx oder Lighttpd). 

In Kombination mit einem lokalen DNS-Server (z. B. Pi-hole) und lokalen SSL-Zertifikaten lässt sich das Tool so wunderbar zentral unter einer eigenen internen Domain wie `https://tracker.home.lan` im gesamten Heimnetzwerk aufrufen.

## 🛠️ Technologien

* **HTML5, CSS3 & Vanilla JavaScript**
* **Zero Dependencies:** Es werden keine Frameworks benötigt. Lediglich für den Excel-Export wird bei Bedarf (und vorhandener Internetverbindung) eine kleine `xlsx`-Bibliothek geladen. Die Offline-Alternative (CSV) funktioniert jederzeit.

## 🤝 Mitwirken

Feedback, Bug-Reports oder Pull Requests zur Erweiterung der Funktionen (z. B. weitere Energieträger im Master-Katalog) sind jederzeit willkommen!

## 📄 Lizenz

Dieses Projekt ist lizenziert unter der [MIT Lizenz](LICENSE).

---
**Autor:** Herr Himmelsbach
