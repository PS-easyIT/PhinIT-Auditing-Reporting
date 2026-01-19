# easyConnections - Network Connections Monitor
## Admin Edition mit WPF GUI

![Version](https://img.shields.io/badge/Version-0.0.3-blue)
![PowerShell](https://img.shields.io/badge/PowerShell-7.0+-green)
![OS](https://img.shields.io/badge/OS-Windows-blue)
![Status](https://img.shields.io/badge/Status-Active-brightgreen)

---

## 📋 Beschreibung

**easyConnections** ist ein leistungsstarkes PowerShell-Tool mit grafischer WPF-Benutzeroberfläche zur Überwachung und Analyse von Netzwerkverbindungen in Echtzeit. Das Tool wurde für fortgeschrittene Netzwerk-Diagnosen und Sicherheitsüberwachung entwickelt.

### Hauptmerkmale
- ✅ **Echtzeit-Überwachung** von TCP und UDP Verbindungen
- ✅ **Farbcodierte Kategorisierung** nach Dienstart (Web, Email, Datenbank, etc.)
- ✅ **Lazy Loading** für schnelle Startzeiten
- ✅ **Aufzeichnungsfunktion** zur Erfassung von Verbindungsdaten über Zeit
- ✅ **HTML-Export** für Reports und Dokumentation
- ✅ **Optionale Prozessinformationen** für detaillierte Analyse
- ✅ **Erweiterte Filterung** nach Protokoll, Richtung und Kategorie
- ✅ **Performance-optimiert** - nur ~2-3 Sekunden Startzeit

---

## 🎯 Features der Version 0.0.3

### Neue Features
| Feature | Beschreibung |
|---------|-------------|
| **Checkbox "Prozessinfo anzeigen"** | Prozess-ID und Name nur bei Bedarf anzeigen |
| **Dynamische Spalten** | DataGrid-Spalten werden je nach Einstellung angezeigt/verborgen |
| **Verbesserte Performance** | 30-40% schneller bei First Load ohne Prozessinformationen |
| **Intelligente Prozess-Auflösung** | Nur bei Bedarf geladen, reduziert Systemlast |

### Farbcodierte Kategorien
```
🔵 Web Services        (Blau)      → HTTP, HTTPS
🟠 Email Services      (Orange)    → SMTP, POP3, IMAP, Exchange
🟢 Database Services   (Grün)      → SQL Server, MySQL, PostgreSQL, Oracle, MongoDB, Redis
🟣 Directory Services  (Violett)   → LDAP, LDAPS, AD Replication, Kerberos
🔴 Remote Access       (Rot)       → SSH, Telnet, RDP, VNC, TeamViewer, WinRM
🟡 File Services       (Gelb)      → FTP, FTPS, SMB/CIFS, NFS, NetBIOS
⚫ Monitoring          (Grau)      → SNMP, Syslog, NTP
🔷 Virtualization      (Türkis)    → Hyper-V, VMware
```

---

## 📊 Unterstützte Protokolle

### Web Services
```
HTTP (80, 8080, 8008, 8000)
HTTPS (443, 8443, 8444)
```

### Email Services
```
SMTP (25, 587, 465)
POP3 (110, 995)
IMAP (143, 993)
Exchange (25, 80, 110, 143, 443, 993, 995, 587, 465)
```

### Database Services
```
SQL Server (1433, 1434)
MySQL (3306)
PostgreSQL (5432)
Oracle (1521, 1522)
MongoDB (27017, 27018, 27019)
Redis (6379)
```

### Directory Services
```
LDAP (389)
LDAPS (636)
AD Replication (135, 389, 636, 3268, 3269, 53, 88, 445)
Kerberos (88, 464)
```

### Remote Access
```
SSH (22)
Telnet (23)
RDP (3389)
VNC (5900-5905)
TeamViewer (5938)
WinRM (5985, 5986)
PowerShell (5985, 5986)
```

### File Services
```
FTP (21, 20)
FTPS (990, 989)
SMB/CIFS (445, 139, 137, 138)
NFS (2049, 111)
NetBIOS (137, 138, 139)
```

### Monitoring
```
SNMP (161, 162)
Syslog (514)
NTP (123)
DNS (53)
DHCP (67, 68)
RADIUS (1812, 1813)
TACACS+ (49)
```

### Virtualization
```
Hyper-V (2179, 5985, 5986)
VMware (443, 902, 903, 8080, 8443)
```

---

## 🎬 Aufzeichnungsfunktion

Die Aufzeichnung erfasst alle Verbindungen kontinuierlich über einen definierten Zeitraum.

### Verwendung
1. **Gewünschte Filter einstellen**
2. **Checkbox "Prozessinfo anzeigen" aktivieren** (optional)
3. **"Aufzeichnung starten" klicken** 
   - Button wird rot
   - Status zeigt "Aufzeichnung läuft..."
4. **Gewünschte Zeit warten**
5. **"Aufzeichnung stoppen" klicken**
   - Automatischer Export-Dialog öffnet sich
   - Report wird als HTML gespeichert

### Export-Report enthält
- Zeitstempel (Start/Ende)
- Erfassungsdauer
- TCP/UDP-Statistiken
- Alle aufgezeichneten Verbindungen
- System-Informationen (Computer, Benutzer)
- Kategorie-Filter Information


---

### Verwendete Libraries
```powershell
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
```

### Hauptfunktionen
| Funktion | Zweck |
|----------|-------|
| `Get-NetworkConnections()` | Ruft TCP/UDP-Verbindungen ab und filtert diese |
| `Get-ProcessNameCached()` | Resolves Process Names mit Caching |
| `Get-HostnameFast()` | Schnelle Hostname-Auflösung |
| `Get-CategoryAndColor()` | Bestimmt Service-Kategorie basierend auf Port |
| `Update-ProcessColumnsVisibility()` | Toggles Prozess-Spalten |
| `Start-ConnectionRecording()` | Startet kontinuierliche Aufzeichnung |
| `Stop-ConnectionRecording()` | Beendet Aufzeichnung und exportiert |
| `Export-RecordingReportHTML()` | Generiert HTML-Report |
| `Export-HTMLReport()` | Exportiert aktuelle Daten als HTML |

### Global Variables
```powershell
$Global:RecordingActive           # Status der Aufzeichnung
$Global:RecordingTimer            # Timer für Aufzeichnung
$Global:RecordedConnections       # Array erfasster Verbindungen
$Global:RecordedConnectionsHash   # Hash für Duplikat-Vermeidung
$Global:ProcessCache              # Prozessname-Cache
$Global:RecordingStartTime        # Startzeitpunkt
$Global:RecordingCategory         # Aufgezeichnete Kategorie
```

---

## 📋 Changelog

### Version 0.0.3 (Aktuell)
- ✅ Checkbox "Prozessinfo anzeigen" hinzugefügt
- ✅ Dynamische DataGrid-Spalten
- ✅ Verbesserte Performance (30-40% schneller)
- ✅ Event Handler für Spalten-Sichtbarkeit
- ✅ Intelligente Prozess-Auflösung
- ✅ Fehlerbereinigung und Testing

### Version 0.0.2
- Farbcodierte Kategorien
- Aufzeichnungsfunktion
- Recording Reports
- HTML Export

### Version 0.0.1
- Initial Release
- Basis-Verbindungsüberwachung
- Protokoll-Filterung