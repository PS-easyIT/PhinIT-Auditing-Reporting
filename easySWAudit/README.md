# 🔍 Client Audit Tool

[![Version](https://img.shields.io/badge/version-0.0.1-blue.svg)](https://github.com/PS-easyIT)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

> **Umfassendes PowerShell Tool zur Analyse installierter Programme, laufender Prozesse und Programmnutzung auf Windows-Clients**

---

## 🎯 Übersicht

Das **Client Audit Tool** ist ein leistungsstarkes PowerShell-basiertes Analysetool mit moderner WPF GUI, das entwickelt wurde, um umfassende Informationen über installierte Programme, laufende Prozesse und Programmnutzung auf Windows-Clients zu erfassen.

### Hauptmerkmale

- 🖥️ **Moderne WPF GUI** - Benutzerfreundliche grafische Oberfläche
- 📦 **Umfassende Analyse** - Erfassung von installierten Programmen, Store Apps, Prozessen und mehr
- 💾 **Flexible Exports** - CSV und HTML Export-Optionen
- 🔍 **Intelligente Filter** - Ausschluss von Windows-Systemprogrammen
- 📊 **Detaillierte Berichte** - Übersichtliche Darstellung aller erfassten Daten
- ⚡ **Schnelle Ausführung** - Effiziente Datenerfassung und -verarbeitung

---

## ✨ Features

### Audit-Kategorien

- ✅ **Installierte Programme** - Erfassung aller installierten Anwendungen (Registry-basiert)
- ✅ **Windows Store Apps** - Analyse aller UWP/AppX-Pakete
- ✅ **Laufende Prozesse** - Aktuelle Prozesse mit Speichernutzung
- ✅ **Prefetch-Analyse** - Nutzungshäufigkeit basierend auf Prefetch-Dateien
- ✅ **Programm-Inventar** - Vollständige Liste aller installierten Software
- ✅ **Event-Logs** - Anwendungs-Event-Log-Analyse
- ✅ **Desktop-Verknüpfungen** - Erfassung aller Desktop-Shortcuts

### Export-Funktionen

- 📊 **CSV Export** - Strukturierte Datenexports für Excel/Analyse
- 📄 **HTML Report** - Professioneller, druckfähiger Report mit Styling
- 🎨 **Formatierte Ausgabe** - Übersichtliche Tabellen und Zusammenfassungen

### Filter & Optionen

- ❌ **Windows-Programme ausschließen** - Filtert Microsoft/Windows System-Software
- 🔧 **Selektive Audits** - Wählen Sie nur die benötigten Kategorien
- 📁 **Automatische Organisation** - Strukturierte Ablage der Exports

---

## 📦 Voraussetzungen

### System-Anforderungen

- **Betriebssystem**: Windows 10/11 oder Windows Server 2016+
- **PowerShell**: Version 5.1 oder höher
- **Framework**: .NET Framework 4.5+
- **Berechtigungen**: 
  - Normale Benutzerrechte für grundlegende Audits
  - Administrator-Rechte empfohlen für:
    - Prefetch-Analyse
    - Vollständige Event-Log-Auswertung
    - System-weite Store Apps

### PowerShell-Module

Das Tool verwendet nur eingebaute Windows-Cmdlets - keine zusätzlichen Module erforderlich:
- `Get-ItemProperty` (Registry-Zugriff)
- `Get-AppxPackage` (Store Apps)
- `Get-Process` (Prozesse)
- `Get-WinEvent` (Event-Logs)

---

## 🚀 Installation

### Download & Ausführung

1. **Herunterladen**
   ```powershell
   # Klonen des Repositories
   git clone https://github.com/PS-easyIT/easyAuditing.git
   
   # Oder direkt die PS1-Datei herunterladen
   ```

2. **Ausführungsrichtlinie anpassen** (falls erforderlich)
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **Script starten**
   ```powershell
   cd easyAuditing\ClientAudit
   .\clAudit_V0.0.1.ps1
   ```

### Alternative: Direkte Ausführung

Rechtsklick auf `clAudit_V0.0.1.ps1` → **"Mit PowerShell ausführen"**

---

## 💻 Verwendung

### Grundlegende Bedienung

1. **Tool starten** - Führen Sie die PS1-Datei aus
2. **Optionen wählen** - Aktivieren/Deaktivieren Sie die gewünschten Audit-Kategorien
3. **Filter setzen** (optional) - Schließen Sie Windows-Programme aus
4. **Audit starten** - Klicken Sie auf "▶ Audit starten"
5. **Ergebnisse prüfen** - Wechseln Sie zwischen den Kategorien
6. **Export** - Exportieren Sie die Daten als CSV oder HTML

### GUI-Elemente

```
┌─────────────────────────────────────────────────────────────┐
│  🔍 Client Audit Tool                                       │
├─────────────────────┬───────────────────────────────────────┤
│ System-Info         │ Datenansicht                          │
│ 💻 Computer         │ ┌─────────────────────────────────┐  │
│ 👤 Benutzer         │ │ Dropdown: Kategorie auswählen   │  │
│ 🖥️ OS               │ └─────────────────────────────────┘  │
│                     │                                       │
│ Audit-Optionen:     │ ┌─────────────────────────────────┐  │
│ ☑ Programme         │ │                                 │  │
│ ☑ Store Apps        │ │     DataGrid mit Ergebnissen    │  │
│ ☑ Prozesse          │ │                                 │  │
│ ☐ Prefetch          │ └─────────────────────────────────┘  │
│ ☐ Inventar          │                                       │
│ ☐ Event-Logs        │ Zusammenfassung: X Einträge           │
│ ☑ Verknüpfungen     │                                       │
│                     │                                       │
│ Filter:             │                                       │
│ ☐ Windows exclud    │                                       │
│                     │                                       │
│ ▶ Audit starten     │                                       │
│ 💾 Export           │                                       │
│ 🗑️ Ergebnisse lösch │                                       │
│                     │                                       │
│ Status: Bereit      │                                       │
└─────────────────────┴───────────────────────────────────────┘
```

---

## 🗂️ Audit-Kategorien

### 📦 Installierte Programme

Erfasst alle über die Windows Registry registrierten Programme.

**Datenquellen:**
- `HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*`
- `HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*`
- `HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*`

**Erfasste Informationen:**
- Programmname
- Version
- Publisher/Hersteller
- Installationsdatum

---

### 🏪 Windows Store Apps

Listet alle installierten UWP/AppX-Pakete auf.

**Erfasste Informationen:**
- App-Name
- Version
- Publisher
- Installationspfad

---

### ⚡ Laufende Prozesse

Zeigt aktuell laufende Prozesse mit GUI-Fenstern.

**Erfasste Informationen:**
- Prozessname
- Prozess-ID
- Fenstertitel
- Speichernutzung (MB)

---

### 📊 Prefetch-Analyse

Analysiert Prefetch-Dateien zur Ermittlung der Programmnutzung.

**Hinweis:** Erfordert Administrator-Rechte

**Erfasste Informationen:**
- Programmname (aus Prefetch-Datei)
- Letzter Zugriff
- Erstellungsdatum
- Dateigröße (KB)

---

### 📋 Programm-Inventar

Erstellt eine bereinigte Liste aller installierten Programme (ohne Duplikate).

**Erfasste Informationen:**
- Programmname
- Publisher
- Installationspfad

---

### 📝 Event-Logs

Analysiert Application Event-Logs (letzte 100 Events, IDs: 1000, 1001, 1002).

**Erfasste Informationen:**
- Zeitstempel
- Event-ID
- Level (Information/Warnung/Fehler)
- Quelle
- Nachricht (erste 100 Zeichen)

---

### 🔗 Desktop-Verknüpfungen

Erfasst alle Desktop-Shortcuts (Public + User Desktop).

**Erfasste Informationen:**
- Verknüpfungsname
- Vollständiger Pfad
- Speicherort (Public/User Desktop)
- Letzte Änderung

---

## 💾 Export-Optionen

### CSV Export

```
Desktop\ClientAudit_CSV_20260119_143052\
├── Installierte_Programme.csv
├── Windows_Store_Apps.csv
├── Laufende_Prozesse.csv
├── Prefetch_Analyse.csv
├── Programm_Inventar.csv
├── Event_Logs.csv
└── Programmverknüpfungen.csv
```

**Vorteile:**
- Excel/Power BI kompatibel
- Einfache Weiterverarbeitung
- Große Datenmengen unterstützt

---

### HTML Report

```
Desktop\ClientAudit_Report_20260119_143052.html
```

**Vorteile:**
- Professionelles Layout
- Sofort druckbar
- In jedem Browser öffenbar
- Vollständiger Report in einer Datei

**Report-Struktur:**
- Header mit System-Informationen
- Inhaltsverzeichnis
- Kategorien mit formatierten Tabellen
- Zusammenfassung

---

## 🔧 Filter-Optionen

### Windows-Programme ausschließen

Filtert Microsoft/Windows-bezogene Software heraus:

**Gefiltert werden:**
- Publisher enthält "Microsoft" oder "Windows"
- Programmname enthält "Microsoft", "Windows", "Update"
- Installations-Hotfixes (KB-Nummern)
- System32-basierte Software

**Anwendungsfälle:**
- Fokus auf Drittanbieter-Software
- Reduzierte Datenmenge
- Übersichtlichere Reports
- Lizenzaudits

---

## 📚 Beispiele

### Beispiel 1: Schnelles Software-Audit

```powershell
# 1. Script starten
.\clAudit_V0.0.1.ps1

# 2. Nur "Installierte Programme" aktivieren
# 3. "Windows-Programme ausschließen" aktivieren
# 4. Audit starten
# 5. Als CSV exportieren
```

**Ergebnis:** Liste aller Drittanbieter-Software

---

### Beispiel 2: Vollständiges System-Audit

```powershell
# 1. Als Administrator starten (für Prefetch)
# 2. Alle Kategorien aktivieren
# 3. Keine Filter
# 4. Audit starten
# 5. Als HTML exportieren
```

**Ergebnis:** Vollständiger System-Report mit allen Informationen

---

### Beispiel 3: Prozess-Überwachung

```powershell
# 1. Nur "Laufende Prozesse" aktivieren
# 2. Audit starten
# 3. Nach Speichernutzung sortieren
```

**Ergebnis:** Übersicht der aktiven Anwendungen mit Ressourcennutzung

---

## 📝 Hinweise

### Performance

- **Schnelle Kategorien** (< 5 Sek.): Programme, Store Apps, Prozesse, Verknüpfungen
- **Mittlere Dauer** (5-15 Sek.): Prefetch, Inventar
- **Längere Dauer** (15-30 Sek.): Event-Logs bei großen Log-Dateien

### Berechtigungen

| Kategorie | Benutzer | Admin |
|-----------|----------|-------|
| Programme | ✅ | ✅ |
| Store Apps | ✅ | ✅ |
| Prozesse | ✅ | ✅ |
| Prefetch | ❌ | ✅ |
| Inventar | ✅ | ✅ |
| Event-Logs | ⚠️ (eingeschränkt) | ✅ |
| Verknüpfungen | ✅ | ✅ |

### Best Practices

1. **Regelmäßige Audits** - Monatlich für Software-Compliance
2. **Admin-Rechte** - Für vollständige Daten
3. **Filter verwenden** - Reduziert Datenmenge bei großen Umgebungen
4. **CSV für Analyse** - HTML für Reports/Dokumentation
5. **Vergleiche** - Speichern Sie Reports für Verlaufsanalysen

---

## 🔄 Changelog

### Version 0.0.1 (19.01.2026)
- ✨ Initiales Release
- 🎨 Moderne WPF GUI
- 📦 7 Audit-Kategorien implementiert
- 💾 CSV & HTML Export
- 🔍 Filter-Optionen
- 📊 Interaktive Datenansicht
- ⚡ Performance-Optimierungen

---

## 📜 Lizenz

Dieses Projekt ist unter der MIT-Lizenz lizenziert - siehe [LICENSE](LICENSE) für Details.

---

## 👤 Autor

**Andreas Hepp**  
📧 Email: [Kontakt über Website](https://www.phinit.de)  
🌐 Website: [www.phinit.de](https://www.phinit.de)  
💼 GitHub: [PS-easyIT](https://github.com/PS-easyIT)

---

## 🤝 Beitragen

Contributions, Issues und Feature Requests sind willkommen!  
Fühlen Sie sich frei, das [Issues Page](https://github.com/PS-easyIT/easyAuditing/issues) zu besuchen.

---

## ⭐ Support

Wenn Ihnen dieses Projekt gefällt, geben Sie ihm bitte einen ⭐ auf GitHub!

---

## 🔗 Verwandte Projekte

- [easyADGroups](https://github.com/PS-easyIT/easyADGroups) - Active Directory Gruppenverwaltung
- [easyLAPS](https://github.com/PS-easyIT/easyLAPS) - LAPS Management Tool
- [easyONBOARDING](https://github.com/PS-easyIT/easyONBOARDING) - Mitarbeiter-Onboarding Automation

---

**Entwickelt mit ❤️ in PowerShell**

*Letzte Aktualisierung: 19.01.2026*
