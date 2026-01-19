<div align="center">

# 🛡️ PhinIT Auditing & Reporting Suite

[![Language: German](https://img.shields.io/badge/Language-Deutsch-black?style=flat-square)](#-deutsche-version)
[![Language: English](https://img.shields.io/badge/Language-English-blue?style=flat-square)](#-english-version)

</div>

---

<a name="-deutsche-version"></a>
## 🇩🇪 Deutsche Version

**Ein zentrales Toolkit für Systemadministratoren zur Analyse, Überwachung und Dokumentation von IT-Umgebungen.**

Dieses Repository bündelt Skripte und Tools, um Transparenz in Server-Landschaften, Active Directory Strukturen und Client-Konfigurationen zu bringen. Fokus liegt auf Sicherheit, Berechtigungsmanagement (ACLs) und Health-Checks.

### 🚀 Features & Umfang

#### 🖥️ Windows Server Audit & Clients
* **Inventory & Health:** Automatisierte Bestandsaufnahme von Hardware, OS-Versionen und Uptime.
* **Software-Inventur:** Detaillierte Auflistung aller **installierten Programme** und Versionen.
* **Verbindungs-Audit:** Prüfung von Netzwerkverbindungen, offenen Ports und Erreichbarkeit wichtiger Dienste.
* **Performance Reporting:** Schnelle Analyse von Ressourcenengpässen (CPU, RAM, Disk).
* **Patch-Level Analyse:** Übersicht über fehlende Updates.

#### 👥 Active Directory (AD) Audit & Reporting
* **User & Group Reporting:** Identifizieren von inaktiven Nutzern, verschachtelten Gruppen und verwaisten Accounts.
* **Security Audits:** Überprüfung von Admin-Rechten, sensitiven Gruppen und Kennwort-Richtlinien.

#### 🔐 Permissions & Security
* **ACL Scanner:** Rekursive Analyse von NTFS-Berechtigungen auf Fileservern.
* **Freigabe-Berichte:** Übersicht offener Netzwerkfreigaben und deren Zugriffsrechte.

### 🛠️ Voraussetzungen
* **PowerShell Version:** Die benötigte Version (PS 5.1 oder PS Core 7+) ist in der jeweiligen **README des Tools/Skripts** vermerkt.
* **RSAT Tools:** Für AD-Module und Abfragen erforderlich.
* **Berechtigungen:** Ausführungsrechte (ExecutionPolicy) müssen entsprechend gesetzt sein und der User benötigt Leserechte auf den Zielsystemen.

### 📦 Installation & Nutzung

1. **Repository klonen:**
   ```powershell
   git clone [https://github.com/DEIN-USER/PhinIT-Audit-Suite.git](https://github.com/DEIN-USER/PhinIT-Audit-Suite.git)

```

2. **Modul/Skript ausführen:**
Wechseln Sie in das entsprechende Verzeichnis und beachten Sie die dortigen Anweisungen.
```powershell
.\Start-Audit.ps1 -Scope All

```



### 🤝 Contributing

Verbesserungsvorschläge und Pull Requests sind willkommen!


---


<a name="-english-version"></a>

## 🇬🇧 English Version

**A central toolkit for system administrators to analyze, monitor, and document IT environments.**

This repository bundles scripts and tools designed to bring transparency to Server landscapes, Active Directory structures, and Client configurations. The focus is on security, permission management (ACLs), and system health checks.

### 🚀 Features & Scope

#### 🖥️ Windows Server Audit & Clients

* **Inventory & Health:** Automated inventory of hardware, OS versions, and uptime.
* **Software Inventory:** Detailed listing of all **installed programs** and versions.
* **Connection Audit:** Verification of network connections, open ports, and service availability.
* **Performance Reporting:** Rapid analysis of resource bottlenecks (CPU, RAM, Disk).
* **Patch-Level Analysis:** Overview of missing updates.

#### 👥 Active Directory (AD) Audit & Reporting

* **User & Group Reporting:** Identify inactive users, nested groups, and orphaned accounts.
* **Security Audits:** Review of Admin rights, sensitive groups, and password policies.

#### 🔐 Permissions & Security

* **ACL Scanner:** Recursive analysis of NTFS permissions on file servers.
* **Share Reports:** Overview of open network shares and their access rights.

### 🛠️ Prerequisites

* **PowerShell Version:** The required version (PS 5.1 or PS Core 7+) is noted in the **specific README of each tool/script**.
* **RSAT Tools:** Required for Active Directory modules.
* **Permissions:** Execution rights (ExecutionPolicy) must be set accordingly, and the user needs read permissions on the target systems.

### 📦 Installation & Usage

1. **Clone Repository:**
```powershell
git clone [https://github.com/YOUR-USER/PhinIT-Audit-Suite.git](https://github.com/YOUR-USER/PhinIT-Audit-Suite.git)

```


2. **Run Module/Script:**
Navigate to the specific directory and follow the instructions provided there.
```powershell
.\Start-Audit.ps1 -Scope All

```



### 🤝 Contributing

Suggestions for improvements and Pull Requests are welcome!

```

```
