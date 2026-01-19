Das ist perfekt! Die Ordnernamen im Bild (`phinit.de_easy...`) sind sehr sprechend und lassen sich hervorragend als **Modul-Übersicht** in die README integrieren.

Ich habe die Struktur angepasst, um genau diese 5 Module aufzulisten und deren Funktion zu beschreiben.

Hier ist der aktualisierte Markdown-Code für deine `README.md`:

```markdown
<div align="center">

# 🛡️ PhinIT Auditing & Reporting Suite

[![Language: German](https://img.shields.io/badge/Language-Deutsch-black?style=flat-square)](#-deutsche-version)
[![Language: English](https://img.shields.io/badge/Language-English-blue?style=flat-square)](#-english-version)

</div>

---

<a name="-deutsche-version"></a>
## 🇩🇪 Deutsche Version

**Ein zentrales Toolkit für Systemadministratoren zur Analyse, Überwachung und Dokumentation von IT-Umgebungen.**

Dieses Repository bündelt spezialisierte Module ("easy"-Serie), um Transparenz in Server-Landschaften, Active Directory Strukturen und Client-Konfigurationen zu bringen.

### 📂 Enthaltene Module & Funktionen

Das Repository ist in eigenständige Module unterteilt, die spezifische Audit-Aufgaben übernehmen:

#### 1. `phinit.de_easyADReport` (Active Directory)
* **User & Group Audit:** Identifiziert inaktive Nutzer, verwaiste Accounts und verschachtelte Gruppen.
* **Security:** Überprüfung von Admin-Rechten und sensitiven Gruppenmitgliedschaften.

#### 2. `phinit.de_easySRVAudit` (Server Health)
* **Server-Check:** Umfassende Gesundheitsprüfung für Windows Server.
* **Dienste & Logs:** Analyse kritischer Windows-Dienste und Event-Logs auf Fehler.
* **Ressourcen:** Reporting zu Disk-Space, CPU-Auslastung und RAM.

#### 3. `phinit.de_easySWAudit` (Software)
* **Inventur:** Detaillierte Auflistung aller installierten Programme.
* **Versionierung:** Abgleich von Software-Versionen (Patch-Level-Analyse).

#### 4. `phinit.de_easyHWAudit` (Hardware)
* **System-Info:** Auslesen von Modell, Seriennummern, BIOS-Versionen und Garantie-relevanten Daten.
* **Client-Audit:** Schnelle Bestandsaufnahme für Workstations und Laptops.

#### 5. `phinit.de_easyConnections` (Netzwerk)
* **Verbindungs-Audit:** Prüfung von Netzwerkpfaden und Erreichbarkeit.
* **Port-Scan:** Überprüfung offener Ports zu Zielsystemen.

### 🛠️ Voraussetzungen
* **PowerShell Version:** Die benötigte Version (PS 5.1 oder PS Core 7+) ist in der jeweiligen **README des Moduls** vermerkt.
* **RSAT Tools:** Für das Modul `easyADReport` zwingend erforderlich.
* **Berechtigungen:** Ausführungsrechte (ExecutionPolicy) müssen gesetzt sein; Leserechte (bzw. Admin-Rechte) auf den Zielsystemen sind für `SRVAudit` und `SWAudit` notwendig.

### 🤝 Contributing
Verbesserungsvorschläge und Pull Requests sind willkommen!

---
---

<a name="-english-version"></a>
## 🇬🇧 English Version

**A central toolkit for system administrators to analyze, monitor, and document IT environments.**

This repository bundles specialized modules (the "easy" series) designed to bring transparency to Server landscapes, Active Directory structures, and Client configurations.

### 📂 Included Modules & Features

The repository is divided into standalone modules handling specific audit tasks:

#### 1. `phinit.de_easyADReport` (Active Directory)
* **User & Group Audit:** Identify inactive users, orphaned accounts, and nested groups.
* **Security:** Review of Admin rights and sensitive group memberships.

#### 2. `phinit.de_easySRVAudit` (Server Health)
* **Server Check:** Comprehensive health check for Windows Servers.
* **Services & Logs:** Analysis of critical Windows services and Event Logs for errors.
* **Resources:** Reporting on Disk space, CPU usage, and RAM.

#### 3. `phinit.de_easySWAudit` (Software)
* **Inventory:** Detailed listing of all installed applications.
* **Versioning:** Verification of software versions (Patch-level analysis).

#### 4. `phinit.de_easyHWAudit` (Hardware)
* **System Info:** Retrieval of model, serial numbers, BIOS versions, and warranty-related data.
* **Client Audit:** Quick inventory for workstations and laptops.

#### 5. `phinit.de_easyConnections` (Network)
* **Connection Audit:** Verification of network paths and availability.
* **Port Scan:** Check for open ports on target systems.

### 🛠️ Prerequisites
* **PowerShell Version:** The required version (PS 5.1 or PS Core 7+) is noted in the **specific README of each module**.
* **RSAT Tools:** Mandatory for the `easyADReport` module.
* **Permissions:** Execution rights (ExecutionPolicy) must be set; Read permissions (or Admin rights) on target systems are required for `SRVAudit` and `SWAudit`.

### 🤝 Contributing
Suggestions for improvements and Pull Requests are welcome!

```
