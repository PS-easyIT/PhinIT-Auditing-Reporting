#region [0.0 | EXE-Kompabilität]
# Stellt sicher, dass das Skript als kompilierte EXE korrekt ausgeführt wird.

# HINWEIS: Dieses Skript erfordert für viele Funktionen Administratorrechte.
# Kompilieren Sie die EXE mit dem Parameter '-requireAdmin', um sicherzustellen, dass sie immer mit den erforderlichen Berechtigungen gestartet wird.

# GUI-spezifische Optimierungen für die EXE-Kompilierung
if ($PSVersionTable.PSEdition -ne 'Core' -and -not $ProgressPreference) {
    # Unterdrückt Fortschrittsanzeigen, die als störende Popup-Fenster erscheinen, wenn sie mit -noConsole kompiliert werden.
    $ProgressPreference = 'SilentlyContinue'
}

# Aktiviert visuelle Stile für WinForms/WPF-Steuerelemente, um ein modernes Aussehen zu gewährleisten.
# Dies muss vor der Erstellung von GUI-Objekten erfolgen.
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Universelle Pfadermittlung für Skript- und EXE-Modus
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    $script:ScriptPath = $PSScriptRoot
} else {
    $script:ScriptPath = Split-Path -Parent -Path ([System.Environment]::GetCommandLineArgs()[0])
    if (-not $script:ScriptPath) { $script:ScriptPath = "." }
}
#endregion

# easyWSAudit - Windows Server Audit Tool
# Version 0.3.3 (Debug Edition)
# Autor: Andreas Hepp
# Datum: 13.07.2025
# Beschreibung: Dieses Skript fuehrt eine Reihe von Tests und Abfragen auf einem Windows Server durch, um die Systemkonfiguration und -sicherheit zu pruefen.

# Debug-Modus - Auf $true setzen fuer ausfuehrliche Logging-Informationen
$DEBUG = $true  # TEMPORÄR AUF TRUE FÜR DIAGNOSE
$DebugLogPath = "$env:TEMP\easyWSAudit_Debug.log"

# Debug-Funktion
function Write-DebugLog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$Source = "General",
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO"
    )
    
    if ($DEBUG) {
        try {
            # Create standardized log entry with ISO 8601 timestamp format
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
            $logMessage = "[$timestamp] [$Level] [$Source] $Message"
            
            # Determine log directory and filename
            $logDirectory = Join-Path -Path (Split-Path -Parent $PSCommandPath) -ChildPath "Logs"
            $logFileName = "easyWSAudit-DEBUG_$(Get-Date -Format 'yyyyMMdd').log"
            $logFilePath = Join-Path -Path $logDirectory -ChildPath $logFileName
            
            # Ensure log directory exists with error handling
            if (-not (Test-Path -Path $logDirectory -PathType Container)) {
                try {
                    New-Item -Path $logDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
                }
                catch {
                    # Fall back to temp directory if we can't create the log directory
                    $logDirectory = $env:TEMP
                    $logFilePath = Join-Path -Path $logDirectory -ChildPath $logFileName
                }
            }
            
            # Check if log file exists and implement basic rotation if needed
            if ((Test-Path -Path $logFilePath) -and 
                ((Get-Item -Path $logFilePath).Length -gt 10MB)) {
                $backupName = "easyWSAudit-DEBUG_$(Get-Date -Format 'yyyyMMdd_HHmmss').log.bak"
                $backupPath = Join-Path -Path $logDirectory -ChildPath $backupName
                
                try {
                    Move-Item -Path $logFilePath -Destination $backupPath -Force -ErrorAction Stop
                }
                catch {
                    # If rotation fails, continue with existing file
                    # Just append a note in the log
                    $rotationError = "[$timestamp] [WARNING] [Logger] Failed to rotate log file: $($_.Exception.Message)"
                }
            }
            
            # Write to log file with retry mechanism for locked files
            $maxRetries = 3
            $retryCount = 0
            $success = $false
            
            while (-not $success -and $retryCount -lt $maxRetries) {
                try {
                    # Use mutex-like approach with Add-Content to handle concurrent writes
                    $logMessage | Out-File -FilePath $logFilePath -Append -Encoding UTF8 -ErrorAction Stop
                    $success = $true
                }
                catch [System.IO.IOException] {
                    # File might be locked, wait and retry
                    $retryCount++
                    Start-Sleep -Milliseconds 100
                }
                catch {
                    # For other exceptions, try alternative method
                    try {
                        Add-Content -Path $logFilePath -Value $logMessage -Encoding UTF8 -ErrorAction Stop
                        $success = $true
                    }
                    catch {
                        # If still failing, write to alternate location
                        $fallbackPath = Join-Path -Path $env:TEMP -ChildPath "easyWSAudit_fallback.log"
                        try {
                            Add-Content -Path $fallbackPath -Value "FALLBACK LOG: $logMessage" -Encoding UTF8 -ErrorAction Stop
                        }
                        catch {
                            # Last resort - we tried our best but can't log to disk
                            # Just avoid crashing the application
                        }
                        $retryCount = $maxRetries  # Exit loop
                    }
                }
            }
            
            # If rotation error occurred and we can now write to the file, log it
            if ($rotationError -and $success) {
                $rotationError | Out-File -FilePath $logFilePath -Append -Encoding UTF8 -ErrorAction SilentlyContinue
            }
            
            # Update UI if available - use safe dispatcher pattern
            if ($script:txtDebugOutput -and $window) {
                try {
                    # Invoke with timeout to prevent UI blocking
                    $dispatcher = $script:txtDebugOutput.Dispatcher
                    if ($dispatcher -and $dispatcher.CheckAccess()) {
                        # We're on the UI thread, update directly
                        $script:txtDebugOutput.Text += "$logMessage`r`n"
                        $script:txtDebugOutput.ScrollToEnd()
                    }
                    else {
                        # We're not on the UI thread, invoke via dispatcher
                        $dispatcher.Invoke([Action]{
                            $script:txtDebugOutput.Text += "$logMessage`r`n"
                            $script:txtDebugOutput.ScrollToEnd()
                        }, "Normal", [TimeSpan]::FromMilliseconds(300))
                    }
                }
                catch {
                    # Ignore UI update failures - logging to file is more important
                }
            }
            
            # Update global path for other functions
            $script:DebugLogPath = $logFilePath
        }
        catch {
            # Ultimate fallback - if the logging system itself fails
            # Just ensure the application continues running
        }
    }
}

# Initialize the Debug Log with proper startup information
if ($DEBUG) {
    # Create startup header with system information
    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $computerInfo = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    
    $startupInfo = @(
        "=== easyWSAudit Debug Log Started $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
        "System: $($osInfo.Caption) ($($osInfo.Version))"
        "Computer: $($computerInfo.Name)"
        "User: $($env:USERNAME)"
        "PowerShell: $($PSVersionTable.PSVersion.ToString())"
        "Script Path: $($MyInvocation.MyCommand.Path)"
    )
    
    foreach ($line in $startupInfo) {
        Write-DebugLog $line "Init" "INFO"
    }
}

# Sichere DoEvents-Implementierung für UI-Updates (ersetzt durch WPF Dispatcher)
function Invoke-SafeDoEvents {
    param(
        [int]$SleepMilliseconds = 50
    )
    
    try {
        Write-DebugLog "Invoke-SafeDoEvents: Führe UI-Update über WPF Dispatcher aus." "DoEvents"
        
        # Prüfe, ob eine Window-Instanz und ein Dispatcher vorhanden sind
        if ($null -eq $window -or $null -eq $window.Dispatcher) {
            Write-DebugLog "WARNUNG: Window oder Dispatcher nicht verfügbar. Fallback auf Start-Sleep." "DoEvents"
            Start-Sleep -Milliseconds $SleepMilliseconds
            return
        }

        # Führe eine leere Aktion mit niedriger Priorität im Dispatcher aus.
        # Dies erlaubt der UI, anstehende Nachrichten (Zeichnen, Eingaben) zu verarbeiten.
        # Es ist die empfohlene Alternative zu DoEvents() in WPF.
        $window.Dispatcher.Invoke([Action]{}, "Background")
        
        # Eine kurze Pause kann immer noch nützlich sein, um die CPU-Last zu reduzieren.
        if ($SleepMilliseconds -gt 0) {
            Start-Sleep -Milliseconds $SleepMilliseconds
        }
        
        Write-DebugLog "Invoke-SafeDoEvents: UI-Update über Dispatcher abgeschlossen." "DoEvents"
        
    } catch [System.TimeoutException] {
        Write-DebugLog "WARNUNG: Dispatcher-Timeout in Invoke-SafeDoEvents." "DoEvents"
    } catch {
        Write-DebugLog "KRITISCHER FEHLER in Invoke-SafeDoEvents: $($_.Exception.Message)" "DoEvents"
        # Fallback: Einfache Pause im Fehlerfall
        Start-Sleep -Milliseconds $SleepMilliseconds
    }
}

# Sichere Dispatcher-Invoke Funktion für WPF UI-Updates
function Invoke-SafeDispatcher {
    param(
        [System.Windows.Threading.DispatcherPriority]$Priority = "Background"
    )
    
    try {
        Write-DebugLog "Invoke-SafeDispatcher: Führe WPF Dispatcher-Update aus" "Dispatcher"
        
        # Prüfe ob wir eine Window-Instanz haben
        if ($null -eq $window) {
            Write-DebugLog "WARNUNG: Window-Instanz nicht verfügbar - Dispatcher übersprungen" "Dispatcher"
            return
        }
        
        # Prüfe ob wir auf dem UI-Thread sind
        if ($window.Dispatcher.CheckAccess()) {
            # Wir sind bereits auf dem UI-Thread - führe leeren Dispatch aus um UI zu aktualisieren
            $window.Dispatcher.Invoke([Action]{}, $Priority)
            Write-DebugLog "Dispatcher-Update auf UI-Thread ausgeführt" "Dispatcher"
        } else {
            # Wir sind nicht auf dem UI-Thread - sichere Invoke mit Timeout
            $window.Dispatcher.Invoke([Action]{}, $Priority, [TimeSpan]::FromSeconds(2))
            Write-DebugLog "Dispatcher-Update von Background-Thread ausgeführt" "Dispatcher"
        }
        
    } catch [System.TimeoutException] {
        Write-DebugLog "WARNUNG: Dispatcher-Timeout erreicht" "Dispatcher"
    } catch {
        Write-DebugLog "FEHLER in Invoke-SafeDispatcher: $($_.Exception.Message)" "Dispatcher"
    }
}

# Importiere notwendige Module
Write-DebugLog "Importiere notwendige Module..." "Init"
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Write-DebugLog "Module importiert" "Init"

# Server-Rollen und Systeminformationen Audit
$commands = @(
    # === SYSTEM INFORMATIONEN ===
    @{Name="Systeminformationen"; Command="Get-ComputerInfo"; Type="PowerShell"; Category="System"},
    @{Name="Betriebssystem Details"; Command="Get-CimInstance Win32_OperatingSystem"; Type="PowerShell"; Category="System"},
    @{Name="Hardware Informationen"; Command="Get-CimInstance Win32_ComputerSystem"; Type="PowerShell"; Category="Hardware"},
    @{Name="CPU Informationen"; Command="Get-CimInstance Win32_Processor"; Type="PowerShell"; Category="Hardware"},
    @{Name="Arbeitsspeicher Details"; Command="Get-CimInstance Win32_PhysicalMemory"; Type="PowerShell"; Category="Hardware"},
    @{Name="Festplatten Informationen"; Command="Get-CimInstance Win32_LogicalDisk"; Type="PowerShell"; Category="Storage"},
    @{Name="Volume Informationen"; Command="Get-Volume"; Type="PowerShell"; Category="Storage"},
    @{Name="Installierte Features und Rollen"; Command="Get-WindowsFeature | Where-Object { `$_.Installed -eq `$true }"; Type="PowerShell"; Category="Features"},
    @{Name="Installierte Programme"; Command="Get-CimInstance Win32_Product | Select-Object Name, Version, Vendor | Sort-Object Name"; Type="PowerShell"; Category="Software"},
    @{Name="Windows Updates"; Command="Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 20"; Type="PowerShell"; Category="Updates"},
    @{Name="Netzwerkkonfiguration"; Command="Get-NetIPConfiguration"; Type="PowerShell"; Category="Network"},
    @{Name="Netzwerkadapter"; Command="Get-NetAdapter"; Type="PowerShell"; Category="Network"},
    @{Name="Aktive Netzwerkverbindungen"; Command="Get-NetTCPConnection | Where-Object State -eq 'Listen' | Select-Object LocalAddress, LocalPort, OwningProcess"; Type="PowerShell"; Category="Network"},
    @{Name="Firewall Regeln"; Command="Get-NetFirewallRule | Where-Object Enabled -eq 'True' | Select-Object DisplayName, Direction, Action | Sort-Object DisplayName"; Type="PowerShell"; Category="Security"},
    @{Name="Services (Automatisch)"; Command="Get-Service | Where-Object StartType -eq 'Automatic' | Sort-Object Status, Name"; Type="PowerShell"; Category="Services"},
    @{Name="Services (Laufend)"; Command="Get-Service | Where-Object Status -eq 'Running' | Sort-Object Name"; Type="PowerShell"; Category="Services"},
    @{Name="Geplante Aufgaben"; Command="Get-ScheduledTask | Where-Object State -eq 'Ready' | Select-Object TaskName, TaskPath, State"; Type="PowerShell"; Category="Tasks"},
    @{Name="Event-Log System (Letzte 24h)"; Command="try { Get-WinEvent -FilterHashtable @{LogName='System'; StartTime=(Get-Date).AddDays(-1)} -MaxEvents 50 -ErrorAction Stop | Select-Object TimeCreated, Id, LevelDisplayName, Message } catch { 'System Event Log nicht verfügbar oder keine Events in den letzten 24h' }"; Type="PowerShell"; Category="Events"},
    @{Name="Event-Log Application (Letzte 24h)"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddDays(-1)} -MaxEvents 50 -ErrorAction Stop | Select-Object TimeCreated, Id, LevelDisplayName, Message } catch { 'Application Event Log nicht verfügbar oder keine Events in den letzten 24h' }"; Type="PowerShell"; Category="Events"},
    @{Name="Lokale Benutzer"; Command="Get-LocalUser | Select-Object Name, Enabled, LastLogon, PasswordRequired"; Type="PowerShell"; Category="Security"},
    @{Name="Lokale Gruppen"; Command="Get-LocalGroup | Select-Object Name, Description"; Type="PowerShell"; Category="Security"},
    
    # === ACTIVE DIRECTORY UMFASSENDE AUDITS ===
    @{Name="AD Domain Controller Status"; Command="Get-ADDomainController -Filter * | Select-Object Name, Site, IPv4Address, OperatingSystem, IsGlobalCatalog, IsReadOnly"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Domain Informationen"; Command="Get-ADDomain | Select-Object Name, NetBIOSName, DomainMode, PDCEmulator, RIDMaster, InfrastructureMaster"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Forest Informationen"; Command="Get-ADForest | Select-Object Name, ForestMode, DomainNamingMaster, SchemaMaster, Sites, Domains"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Organizational Units"; Command="Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName, Description | Sort-Object Name"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Domaenen Administratoren"; Command="Get-ADGroupMember -Identity 'Domain Admins' | Get-ADUser -Properties LastLogonDate, PasswordLastSet, Enabled | Select-Object Name, SamAccountName, Enabled, LastLogonDate, PasswordLastSet"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Enterprise Administratoren"; Command="Get-ADGroupMember -Identity 'Enterprise Admins' | Get-ADUser -Properties LastLogonDate, PasswordLastSet, Enabled | Select-Object Name, SamAccountName, Enabled, LastLogonDate, PasswordLastSet"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Schema Administratoren"; Command="Get-ADGroupMember -Identity 'Schema Admins' | Get-ADUser -Properties LastLogonDate, PasswordLastSet, Enabled | Select-Object Name, SamAccountName, Enabled, LastLogonDate, PasswordLastSet"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Privilegierte Gruppen"; Command="@('Domain Admins','Enterprise Admins','Schema Admins','Administrators','Account Operators','Backup Operators','Print Operators','Server Operators') | ForEach-Object { Get-ADGroup -Identity `$_ -Properties Members | Select-Object Name, @{Name='MemberCount';Expression={`$_.Members.Count}} }"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Benutzer ohne Passwort-Ablauf"; Command="Get-ADUser -Filter {PasswordNeverExpires -eq `$true -and Enabled -eq `$true} -Properties PasswordNeverExpires, LastLogonDate | Select-Object Name, SamAccountName, LastLogonDate"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Deaktivierte Benutzer"; Command="Get-ADUser -Filter {Enabled -eq `$false} | Select-Object Name, SamAccountName, DistinguishedName | Sort-Object Name"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Computer Accounts"; Command="Get-ADComputer -Filter * -Properties OperatingSystem, LastLogonDate | Select-Object Name, OperatingSystem, LastLogonDate, Enabled | Sort-Object LastLogonDate -Descending"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Replikations-Status"; Command="repadmin /replsummary"; Type="CMD"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD FSMO Rollen"; Command="Get-ADDomain | Select-Object PDCEmulator, RIDMaster, InfrastructureMaster; Get-ADForest | Select-Object DomainNamingMaster, SchemaMaster"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Sites und Subnets"; Command="Get-ADReplicationSite | Select-Object Name, Description; Get-ADReplicationSubnet -Filter * | Select-Object Name, Site, Location"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Trust Relationships"; Command="Get-ADTrust -Filter * | Select-Object Name, Direction, TrustType, DisallowTransivity"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    
    # === DNS UMFASSENDE DIAGNOSEN ===
    @{Name="DNS Server Konfiguration"; Command="try { Get-DnsServer -ErrorAction Stop | Select-Object ComputerName, ZoneScavenging, EnableDnsSec -WarningAction SilentlyContinue } catch { 'DNS Server Konfiguration: DNS-Server nicht verfügbar oder DNS-Rolle nicht installiert' }"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Server Zonen"; Command="Get-DnsServerZone | Select-Object ZoneName, ZoneType, IsAutoCreated, IsDsIntegrated, IsReverseLookupZone"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Forwarders"; Command="Get-DnsServerForwarder"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Cache-Inhalt"; Command="Get-DnsServerCache | Select-Object -First 20"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Scavenging Einstellungen"; Command="Get-DnsServerScavenging"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Event-Log Errors"; Command="try { Get-WinEvent -FilterHashtable @{LogName='DNS Server'; Level=2,3} -MaxEvents 50 -ErrorAction Stop | Select-Object TimeCreated, Id, LevelDisplayName, Message } catch { 'DNS Server Event Log nicht verfügbar oder keine Error-Events vorhanden' }"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Root Hints"; Command="Get-DnsServerRootHint"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Service Dependencies"; Command="`$Services='DNS','Netlogon','KDC'; ForEach (`$Service in `$Services) {Get-Service `$Service -ErrorAction SilentlyContinue | Where-Object {`$_ -ne `$null} | Select-Object Name, Status, StartType}"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    
    # === DHCP UMFASSENDE AUDITS ===
    @{Name="DHCP Server Konfiguration"; Command="Get-DhcpServerInDC"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Server Settings"; Command="Get-DhcpServerSetting"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP IPv4 Bereiche"; Command="Get-DhcpServerv4Scope | Select-Object ScopeId, Name, StartRange, EndRange, SubnetMask, State, LeaseDuration"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP IPv6 Bereiche"; Command="Get-DhcpServerv6Scope | Select-Object Prefix, Name, State, PreferredLifetime"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Reservierungen"; Command="try { `$scopes = Get-DhcpServerv4Scope -ErrorAction Stop; if (`$scopes) { `$scopes | ForEach-Object { Get-DhcpServerv4Reservation -ScopeId `$_.ScopeId -ErrorAction SilentlyContinue } | Select-Object ScopeId, IPAddress, ClientId, Name, Description } else { 'Keine DHCP-Bereiche konfiguriert' } } catch { 'DHCP-Reservierungen: DHCP-Server nicht verfügbar oder keine Bereiche konfiguriert' }"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Lease-Statistiken"; Command="try { Get-DhcpServerv4ScopeStatistics -ErrorAction Stop | Select-Object ScopeId, AddressesFree, AddressesInUse, PercentageInUse } catch { 'DHCP-Lease-Statistiken: Keine Bereiche konfiguriert oder DHCP-Server nicht verfügbar' }"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Optionen (Server)"; Command="Get-DhcpServerv4OptionValue | Select-Object OptionId, Name, Value"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Failover Konfiguration"; Command="Get-DhcpServerv4Failover | Select-Object Name, PartnerServer, Mode, State"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Event-Log"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Dhcp-Server/Operational'} -MaxEvents 50 -ErrorAction Stop | Select-Object TimeCreated, Id, LevelDisplayName, Message } catch { 'DHCP Server Event Log nicht verfügbar oder keine Events vorhanden' }"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    
    # === IIS UMFASSENDE AUDITS ===
    @{Name="IIS Server Informationen"; Command="Get-IISServerManager | Select-Object *"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Websites"; Command="Get-IISSite | Select-Object Name, Id, State, PhysicalPath, @{Name='Bindings';Expression={`$_.Bindings | ForEach-Object {`$_.Protocol + '://' + `$_.BindingInformation}}}"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Application Pools"; Command="Get-IISAppPool | Select-Object Name, State, ProcessModel, Recycling"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Anwendungen"; Command="Get-WebApplication | Select-Object Site, Path, PhysicalPath, ApplicationPool"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Virtuelle Verzeichnisse"; Command="Get-WebVirtualDirectory | Select-Object Site, Application, Path, PhysicalPath"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS SSL Zertifikate"; Command="Get-ChildItem IIS:SslBindings | Select-Object IPAddress, Port, Host, Thumbprint, Subject"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Modules"; Command="Get-WebGlobalModule | Select-Object Name, Image"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Handler Mappings"; Command="Get-WebHandler | Select-Object Name, Path, Verb, Modules"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Event-Log"; Command="try { Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName='Microsoft-Windows-IIS*'} -MaxEvents 50 -ErrorAction Stop | Select-Object TimeCreated, Id, LevelDisplayName, Message } catch { 'IIS Event Log nicht verfügbar oder keine IIS-Events vorhanden' }"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    
    # === WDS (WINDOWS DEPLOYMENT SERVICES) ===
    @{Name="WDS Server Konfiguration"; Command="wdsutil /get-server /show:config"; Type="CMD"; FeatureName="WDS"; Category="WDS"},
    @{Name="WDS Boot Images"; Command="wdsutil /get-allimages /show:install"; Type="CMD"; FeatureName="WDS"; Category="WDS"},
    @{Name="WDS Install Images"; Command="wdsutil /get-allimages /show:boot"; Type="CMD"; FeatureName="WDS"; Category="WDS"},
    @{Name="WDS Transport Server"; Command="wdsutil /get-transportserver /show:config"; Type="CMD"; FeatureName="WDS"; Category="WDS"},
    @{Name="WDS Multicast"; Command="wdsutil /get-allmulticasttransmissions"; Type="CMD"; FeatureName="WDS"; Category="WDS"},
    @{Name="WDS Client Einstellungen"; Command="try { Get-WdsClient -ErrorAction Stop | Select-Object DeviceID, DeviceName, BootImagePath, ReferralServer } catch { 'WDS Client Einstellungen: Windows-Bereitstellungsdienste nicht konfiguriert oder WDS-Rolle nicht installiert' }"; Type="PowerShell"; FeatureName="WDS"; Category="WDS"},
    
    # === HYPER-V AUDITS ===
    @{Name="Hyper-V Host Informationen"; Command="Get-VMHost | Select-Object ComputerName, LogicalProcessorCount, MemoryCapacity, VirtualMachinePath, VirtualHardDiskPath"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    @{Name="Hyper-V Virtuelle Maschinen"; Command="Get-VM | Select-Object Name, State, CPUUsage, MemoryAssigned, Uptime, Version, Generation"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    @{Name="Hyper-V Switches"; Command="Get-VMSwitch | Select-Object Name, SwitchType, NetAdapterInterfaceDescription, AllowManagementOS"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    @{Name="Hyper-V Snapshots"; Command="Get-VMSnapshot * | Select-Object VMName, Name, SnapshotType, CreationTime, ParentSnapshotName"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    @{Name="Hyper-V Integration Services"; Command="Get-VM | Get-VMIntegrationService | Select-Object VMName, Name, Enabled, PrimaryStatusDescription"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    @{Name="Hyper-V Replikation"; Command="Get-VMReplication | Select-Object VMName, State, Mode, FrequencySec, PrimaryServer, ReplicaServer"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    
    # === FAILOVER CLUSTER AUDITS ===
    @{Name="Cluster Informationen"; Command="Get-Cluster | Select-Object Name, Domain, AddEvictDelay, BackupInProgress, BlockCacheSize"; Type="PowerShell"; FeatureName="Failover-Clustering"; Category="Cluster"},
    @{Name="Cluster Nodes"; Command="Get-ClusterNode | Select-Object Name, State, Type, ID"; Type="PowerShell"; FeatureName="Failover-Clustering"; Category="Cluster"},
    @{Name="Cluster Resources"; Command="Get-ClusterResource | Select-Object Name, ResourceType, State, OwnerNode, OwnerGroup"; Type="PowerShell"; FeatureName="Failover-Clustering"; Category="Cluster"},
    @{Name="Cluster Shared Volumes"; Command="Get-ClusterSharedVolume | Select-Object Name, State, Node, SharedVolumeInfo"; Type="PowerShell"; FeatureName="Failover-Clustering"; Category="Cluster"},
    @{Name="Cluster Networks"; Command="Get-ClusterNetwork | Select-Object Name, Role, State, Address, AddressMask"; Type="PowerShell"; FeatureName="Failover-Clustering"; Category="Cluster"},
    @{Name="Cluster Validation Report"; Command="Test-Cluster -ReportName 'C:\\temp\\ClusterValidation.htm' -Include 'Storage','Network','System Configuration','Inventory'"; Type="PowerShell"; FeatureName="Failover-Clustering"; Category="Cluster"},
    
    # === WINDOWS SERVER UPDATE SERVICES (WSUS) ===
    @{Name="WSUS Server Konfiguration"; Command="Get-WsusServer | Select-Object Name, PortNumber, ServerProtocolVersion, DatabasePath"; Type="PowerShell"; FeatureName="UpdateServices"; Category="WSUS"},
    @{Name="WSUS Update Kategorien"; Command="Get-WsusServer | Get-WsusClassification | Select-Object Classification, ID"; Type="PowerShell"; FeatureName="UpdateServices"; Category="WSUS"},
    @{Name="WSUS Computer Gruppen"; Command="Get-WsusServer | Get-WsusComputerTargetGroup | Select-Object Name, ID, ComputerTargets"; Type="PowerShell"; FeatureName="UpdateServices"; Category="WSUS"},
    @{Name="WSUS Synchronisation Status"; Command="Get-WsusServer | Get-WsusSubscription | Select-Object LastSynchronizationTime, SynchronizeAutomatically, NumberOfSynchronizations"; Type="PowerShell"; FeatureName="UpdateServices"; Category="WSUS"},
    
    # === FILE SERVICES ===
    @{Name="File Shares"; Command="Get-SmbShare | Select-Object Name, Path, Description, ShareType, ShareState"; Type="PowerShell"; FeatureName="FS-FileServer"; Category="FileServices"},
    @{Name="DFS Namespaces"; Command="Get-DfsnRoot | Select-Object Path, Type, State, Description"; Type="PowerShell"; FeatureName="FS-DFS-Namespace"; Category="FileServices"},
    @{Name="DFS Replication Groups"; Command="try { Get-DfsReplicationGroup -ErrorAction Stop | Select-Object GroupName, State, Description } catch { 'DFS Replication Groups: DFS-Replikation nicht konfiguriert oder FS-DFS-Replication-Rolle nicht installiert' }"; Type="PowerShell"; FeatureName="FS-DFS-Replication"; Category="FileServices"},
    @{Name="File Server Resource Manager Quotas"; Command="Get-FsrmQuota | Select-Object Path, Size, SoftLimit, Usage, Description"; Type="PowerShell"; FeatureName="FS-Resource-Manager"; Category="FileServices"},
    @{Name="Shadow Copies"; Command="Get-CimInstance -ClassName Win32_ShadowCopy | Select-Object VolumeName, InstallDate, Count"; Type="PowerShell"; FeatureName="FS-FileServer"; Category="FileServices"},
    
    # === PRINT SERVICES ===
    @{Name="Print Server Drucker"; Command="Get-Printer | Select-Object Name, DriverName, PortName, Shared, Published, PrinterStatus"; Type="PowerShell"; FeatureName="Print-Services"; Category="PrintServices"},
    @{Name="Print Server Treiber"; Command="Get-PrinterDriver | Select-Object Name, Manufacturer, DriverVersion, PrinterEnvironment"; Type="PowerShell"; FeatureName="Print-Services"; Category="PrintServices"},
    @{Name="Print Server Ports"; Command="Get-PrinterPort | Select-Object Name, Description, PortMonitor, PortType"; Type="PowerShell"; FeatureName="Print-Services"; Category="PrintServices"},
    
    # === REMOTE DESKTOP SERVICES ===
    @{Name="RDS Server Informationen"; Command="Get-RDServer | Select-Object Server, Roles"; Type="PowerShell"; FeatureName="RDS-RD-Server"; Category="RDS"},
    @{Name="RDS Session Collections"; Command="Get-RDSessionCollection | Select-Object CollectionName, CollectionDescription, Size"; Type="PowerShell"; FeatureName="RDS-RD-Server"; Category="RDS"},
    @{Name="RDS User Sessions"; Command="Get-RDUserSession | Select-Object CollectionName, DomainName, UserName, SessionState"; Type="PowerShell"; FeatureName="RDS-RD-Server"; Category="RDS"},
    @{Name="RDS Licensing"; Command="Get-RDLicenseConfiguration | Select-Object Mode, LicenseServer"; Type="PowerShell"; FeatureName="RDS-Licensing"; Category="RDS"},
    
    # === CERTIFICATE SERVICES ===
    @{Name="Certificate Authority Info"; Command="certutil -getconfig"; Type="CMD"; FeatureName="ADCS-Cert-Authority"; Category="PKI"},
    @{Name="CA Certificate Templates"; Command="certutil -template"; Type="CMD"; FeatureName="ADCS-Cert-Authority"; Category="PKI"},
    @{Name="Certificate Store - Personal"; Command="Get-ChildItem Cert:\\LocalMachine\\My | Select-Object Subject, Issuer, NotAfter, Thumbprint"; Type="PowerShell"; Category="PKI"},
    @{Name="Certificate Store - Root"; Command="Get-ChildItem Cert:\\LocalMachine\\Root | Select-Object Subject, Issuer, NotAfter, Thumbprint"; Type="PowerShell"; Category="PKI"},
    
    # === ERWEITERTE SICHERHEITS-AUDITS ===
    @{Name="Security Event Log (Letzte 100)"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Security'} -MaxEvents 100 -ErrorAction Stop | Select-Object TimeCreated, Id, LevelDisplayName, UserId, Message } catch { 'Security Event Log nicht verfügbar oder deaktiviert' }"; Type="PowerShell"; Category="Security"},
    @{Name="Failed Logon Attempts"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4625} -MaxEvents 50 -ErrorAction Stop | Select-Object TimeCreated, Message } catch { 'Security Log Event ID 4625 nicht verfügbar oder keine Failed Logon Events vorhanden' }"; Type="PowerShell"; Category="Security"},
    @{Name="Account Lockouts"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4740} -MaxEvents 50 -ErrorAction Stop | Select-Object TimeCreated, Message } catch { 'Security Log Event ID 4740 nicht verfügbar oder keine Account Lockout Events vorhanden' }"; Type="PowerShell"; Category="Security"},
    @{Name="Privilege Use Auditing"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4672,4673,4674} -MaxEvents 50 -ErrorAction Stop | Select-Object TimeCreated, Id, Message } catch { 'Security Log Event IDs 4672/4673/4674 nicht verfügbar oder keine Privilege Use Events vorhanden' }"; Type="PowerShell"; Category="Security"},
    @{Name="Unsecure LDAP Binds"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Directory Service'; ID=2889} -MaxEvents 20 -ErrorAction Stop | Select-Object TimeCreated, Message } catch { 'Directory Service Log Event ID 2889 nicht verfügbar oder keine Unsecure LDAP Bind Events vorhanden' }"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Security"},
    
    # === ACTIVE DIRECTORY FEDERATION SERVICES (ADFS) ===
    @{Name="ADFS Server Konfiguration"; Command="Get-AdfsProperties | Select-Object DisplayName, HostName, HttpPort, HttpsPort, TlsClientPort"; Type="PowerShell"; FeatureName="ADFS-Federation"; Category="ADFS"},
    @{Name="ADFS Relying Party Trusts"; Command="Get-AdfsRelyingPartyTrust | Select-Object Name, Enabled, Identifier, IssuanceAuthorizationRules"; Type="PowerShell"; FeatureName="ADFS-Federation"; Category="ADFS"},
    @{Name="ADFS Claims Provider Trusts"; Command="Get-AdfsClaimsProviderTrust | Select-Object Name, Enabled, Identifier, AcceptanceTransformRules"; Type="PowerShell"; FeatureName="ADFS-Federation"; Category="ADFS"},
    @{Name="ADFS Certificates"; Command="Get-AdfsCertificate | Select-Object CertificateType, Thumbprint, Subject, NotAfter"; Type="PowerShell"; FeatureName="ADFS-Federation"; Category="ADFS"},
    @{Name="ADFS Endpoints"; Command="Get-AdfsEndpoint | Select-Object AddressPath, Enabled, Protocol, SecurityMode"; Type="PowerShell"; FeatureName="ADFS-Federation"; Category="ADFS"},
    
    # === ACTIVE DIRECTORY LIGHTWEIGHT DIRECTORY SERVICES (ADLDS) ===
    @{Name="ADLDS Instances"; Command="Get-CimInstance -ClassName Win32_Service | Where-Object {`$_.Name -like 'ADAM_*'} | Select-Object Name, State, StartMode, PathName"; Type="PowerShell"; FeatureName="ADLDS"; Category="ADLDS"},
    @{Name="ADLDS Configuration"; Command="dsdbutil -c 'activate instance `$instancename' quit | Out-String"; Type="CMD"; FeatureName="ADLDS"; Category="ADLDS"},
    
    # === ACTIVE DIRECTORY RIGHTS MANAGEMENT SERVICES (ADRMS) ===
    @{Name="ADRMS Cluster Info"; Command="Get-RmsCluster | Select-Object ClusterName, ClusterUrl, Version"; Type="PowerShell"; FeatureName="ADRMS"; Category="ADRMS"},
    @{Name="ADRMS Server Info"; Command="Get-RmsServer | Select-Object Name, ClusterName, Version, IsConnected"; Type="PowerShell"; FeatureName="ADRMS"; Category="ADRMS"},
    @{Name="ADRMS Templates"; Command="Get-RmsTemplate | Select-Object Name, Description, Validity, Created"; Type="PowerShell"; FeatureName="ADRMS"; Category="ADRMS"},
    
    # === DEVICE HEALTH ATTESTATION SERVICE ===
    @{Name="Device Health Attestation Service"; Command="Get-DHASActiveEncryptionCertificate; Get-DHASActiveSigningCertificate"; Type="PowerShell"; FeatureName="DeviceHealthAttestationService"; Category="DeviceAttestation"},
    
    # === VOLUME ACTIVATION SERVICES ===
    @{Name="KMS Server Konfiguration"; Command="slmgr /dlv; cscript C:\\Windows\\System32\\slmgr.vbs /dli"; Type="CMD"; FeatureName="VolumeActivation"; Category="VolumeActivation"},
    @{Name="KMS Client Status"; Command="Get-CimInstance -ClassName SoftwareLicensingProduct | Where-Object {`$_.LicenseStatus -eq 1} | Select-Object Name, Description, LicenseStatus"; Type="PowerShell"; FeatureName="VolumeActivation"; Category="VolumeActivation"},
    
    # === WINDOWS SERVER BACKUP ===
    @{Name="Windows Server Backup Policies"; Command="Get-WBPolicy | Select-Object PolicyState, BackupTargets, FilesSpecsToBackup"; Type="PowerShell"; FeatureName="Windows-Server-Backup"; Category="Backup"},
    @{Name="Windows Server Backup Jobs"; Command="Get-WBJob -Previous 10 | Select-Object JobType, JobState, StartTime, EndTime, HResult"; Type="PowerShell"; FeatureName="Windows-Server-Backup"; Category="Backup"},
    @{Name="Windows Server Backup Disks"; Command="Get-WBDisk | Select-Object DiskNumber, Label, InternalDiskNumber"; Type="PowerShell"; FeatureName="Windows-Server-Backup"; Category="Backup"},
    
    # === NETWORK POLICY AND ACCESS SERVICES (NPAS) ===
    @{Name="NPS Server Konfiguration"; Command="netsh nps show config"; Type="CMD"; FeatureName="NPAS"; Category="NPAS"},
    @{Name="NPS Network Policies"; Command="Get-NpsNetworkPolicy | Select-Object PolicyName, Enabled, ProcessingOrder, ConditionText"; Type="PowerShell"; FeatureName="NPAS"; Category="NPAS"},
    @{Name="NPS Connection Request Policies"; Command="Get-NpsConnectionRequestPolicy | Select-Object Name, Enabled, ProcessingOrder"; Type="PowerShell"; FeatureName="NPAS"; Category="NPAS"},
    @{Name="NPS RADIUS Clients"; Command="Get-NpsRadiusClient | Select-Object Name, Address, SharedSecret, VendorName"; Type="PowerShell"; FeatureName="NPAS"; Category="NPAS"},
    
    # === HOST GUARDIAN SERVICE ===
    @{Name="HGS Service Info"; Command="Get-HgsServer | Select-Object Name, State, Version"; Type="PowerShell"; FeatureName="HostGuardianServiceRole"; Category="HGS"},
    @{Name="HGS Attestation Policies"; Command="Get-HgsAttestationPolicy | Select-Object Name, PolicyVersion, Stage"; Type="PowerShell"; FeatureName="HostGuardianServiceRole"; Category="HGS"},
    
    # === REMOTE ACCESS SERVICES ===
    @{Name="DirectAccess Konfiguration"; Command="Get-DAServer | Select-Object ConnectToAddress, TunnelType, AuthenticationMethod"; Type="PowerShell"; FeatureName="RemoteAccess"; Category="RemoteAccess"},
    @{Name="VPN Server Konfiguration"; Command="Get-VpnServerConfiguration | Select-Object TunnelType, EncryptionLevel, IdleDisconnectSeconds"; Type="PowerShell"; FeatureName="RemoteAccess"; Category="RemoteAccess"},
    @{Name="Routing Table"; Command="Get-NetRoute | Select-Object DestinationPrefix, NextHop, RouteMetric, Protocol"; Type="PowerShell"; FeatureName="RemoteAccess"; Category="RemoteAccess"},
    
    # === WINDOWS INTERNAL DATABASE ===
    @{Name="Windows Internal Database Instanzen"; Command="Get-CimInstance -ClassName Win32_Service | Where-Object {`$_.Name -like '*MSSQL*MICROSOFT*'} | Select-Object Name, State, StartMode"; Type="PowerShell"; FeatureName="Windows-Internal-Database"; Category="InternalDB"},
    @{Name="SQL Server Express Instanzen"; Command="Get-ItemProperty 'HKLM:\\SOFTWARE\\Microsoft\\Microsoft SQL Server\\Instance Names\\SQL' -ErrorAction SilentlyContinue"; Type="PowerShell"; FeatureName="Windows-Internal-Database"; Category="InternalDB"},
    
    # === WINDOWS DEFENDER FEATURES ===
    @{Name="Windows Defender Status"; Command="Get-MpComputerStatus | Select-Object AntivirusEnabled, AMServiceEnabled, RealTimeProtectionEnabled, IoavProtectionEnabled"; Type="PowerShell"; FeatureName="Windows-Defender-Features"; Category="WindowsDefender"},
    @{Name="Windows Defender Preferences"; Command="Get-MpPreference | Select-Object DisableRealtimeMonitoring, DisableIntrusionPreventionSystem, DisableIOAVProtection"; Type="PowerShell"; FeatureName="Windows-Defender-Features"; Category="WindowsDefender"},
    @{Name="Windows Defender Threats"; Command="Get-MpThreatDetection | Select-Object -First 20 | Select-Object ThreatID, ActionSuccess, DetectionTime, ThreatName"; Type="PowerShell"; FeatureName="Windows-Defender-Features"; Category="WindowsDefender"},
    
    # === WINDOWS PROCESS ACTIVATION SERVICE (WAS) ===
    @{Name="WAS Service Status"; Command="Get-Service WAS | Select-Object Name, Status, StartType, ServiceType"; Type="PowerShell"; FeatureName="Windows-Process-Activation-Service"; Category="WAS"},
    @{Name="Application Pool WAS"; Command="Get-IISAppPool | Select-Object Name, State, ProcessModel, Enable32BitAppOnWin64"; Type="PowerShell"; FeatureName="Windows-Process-Activation-Service"; Category="WAS"},
    
    # === WINDOWS SEARCH SERVICE ===
    @{Name="Windows Search Service"; Command="Get-Service WSearch | Select-Object Name, Status, StartType"; Type="PowerShell"; FeatureName="Windows-Search-Service"; Category="SearchService"},
    @{Name="Search Indexer Status"; Command="Get-CimInstance -ClassName Win32_Service | Where-Object Name -eq 'WSearch' | Select-Object Name, State, ProcessId"; Type="PowerShell"; FeatureName="Windows-Search-Service"; Category="SearchService"},
    
    # === ERWEITERTE SYSTEM-AUDITS (basierend auf GitHub Best Practices) ===
    @{Name="Lokale Administratoren"; Command="net localgroup administrators"; Type="CMD"; Category="Security"},
    @{Name="Guest Account Status"; Command="net user guest"; Type="CMD"; Category="Security"},
    @{Name="Shared Folders"; Command="net share"; Type="CMD"; Category="FileSharing"},
    @{Name="Benutzer Profile"; Command="Get-ChildItem C:\\Users | Select-Object Name, CreationTime, LastWriteTime"; Type="PowerShell"; Category="UserProfiles"},
    @{Name="Windows Firewall Profile"; Command="netsh advfirewall show allprofiles"; Type="CMD"; Category="Firewall"},
    @{Name="Power Management"; Command="powercfg /a"; Type="CMD"; Category="PowerManagement"},
    @{Name="Credential Manager"; Command="vaultcmd /listschema; vaultcmd /list"; Type="CMD"; Category="CredentialManager"},
    @{Name="Audit Policy Settings"; Command="auditpol.exe /get /category:*"; Type="CMD"; Category="AuditPolicy"},
    @{Name="Group Policy Results"; Command="gpresult /r"; Type="CMD"; Category="GroupPolicy"},
    @{Name="Installed Software (Registry)"; Command="Get-ItemProperty HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate"; Type="PowerShell"; Category="InstalledSoftware"},
    @{Name="Environment Variables"; Command="Get-ChildItem Env: | Sort-Object Name"; Type="PowerShell"; Category="Environment"},
    
    # === AD HEALTH CHECK (basierend auf Microsoft DevBlogs) ===
    @{Name="AD DHCP Server in AD"; Command="Get-ADObject -SearchBase (`"cn=configuration,`" + (Get-ADDomain).DistinguishedName) -Filter `"objectclass -eq 'dhcpclass' -AND Name -ne 'dhcproot'`" | Select-Object Name"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Service Dependencies Health"; Command="`$Services='DNS','DFSR','IsmServ','KDC','NetLogon','NTDS'; ForEach (`$Service in `$Services) {Get-Service `$Service -ErrorAction SilentlyContinue | Where-Object {`$_ -ne `$null} | Select-Object Name, Status, StartType}"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD DC Diagnostics"; Command="dcdiag /test:dns /e /v"; Type="CMD"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Time Sync Status"; Command="w32tm /query /status"; Type="CMD"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Sysvol Replication"; Command="dfsrdiag replicationstate /member:*"; Type="CMD"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    
    # === WINDOWS SERVER ESSENTIALS ===
    @{Name="Server Essentials Dashboard"; Command="Get-WssUser | Select-Object UserName, FullName, Description, Enabled"; Type="PowerShell"; FeatureName="ServerEssentialsRole"; Category="ServerEssentials"},
    @{Name="Server Essentials Backup"; Command="Get-WssClientBackup | Select-Object ComputerName, LastBackupTime, LastSuccessfulBackupTime"; Type="PowerShell"; FeatureName="ServerEssentialsRole"; Category="ServerEssentials"},
    
    # === STORAGE MANAGEMENT ===
    @{Name="Storage Pools"; Command="Get-StoragePool | Select-Object FriendlyName, HealthStatus, OperationalStatus, TotalPhysicalCapacity"; Type="PowerShell"; FeatureName="Windows-Storage-Management"; Category="Storage"},
    @{Name="Virtual Disks"; Command="Get-VirtualDisk | Select-Object FriendlyName, HealthStatus, OperationalStatus, Size, AllocatedSize"; Type="PowerShell"; FeatureName="Windows-Storage-Management"; Category="Storage"},
    @{Name="Storage Spaces"; Command="Get-StorageSpace | Select-Object FriendlyName, HealthStatus, ProvisioningType, ResiliencySettingName"; Type="PowerShell"; FeatureName="Windows-Storage-Management"; Category="Storage"},
    @{Name="Physical Disks"; Command="Get-PhysicalDisk | Select-Object FriendlyName, HealthStatus, OperationalStatus, Size, MediaType"; Type="PowerShell"; FeatureName="Windows-Storage-Management"; Category="Storage"},
    
    # === MIGRATION SERVICES ===
    @{Name="Windows Server Migration Tools"; Command="Get-SmigServerFeature | Select-Object FeatureName, Status"; Type="PowerShell"; FeatureName="Windows-Server-Migration"; Category="Migration"},
    
    # === WINDOWS IDENTITY FOUNDATION ===
    @{Name="Windows Identity Foundation"; Command="Get-WindowsFeature Windows-Identity-Foundation | Select-Object Name, InstallState, FeatureType"; Type="PowerShell"; FeatureName="Windows-Identity-Foundation"; Category="Identity"}
)

# Sicherheitsaudit-Befehle (neu hinzugefügt)
$securityAuditCommands = @(
    # === LOKALE SICHERHEIT ===
    @{Name="Lokale Administratoren"; Command="try { Get-LocalGroupMember -Group 'Administrators' | Select-Object Name, PrincipalSource, ObjectClass } catch { net localgroup administrators }"; Type="PowerShell"; Category="Lokale Sicherheit"},
    @{Name="Lokale Passwortrichtlinie"; Command="try { secedit /export /cfg `$env:TEMP\secpol.cfg /quiet; `$content = Get-Content `$env:TEMP\secpol.cfg -ErrorAction SilentlyContinue; if (`$content) { `$content | Select-String -Pattern 'PasswordComplexity|MinimumPasswordLength|LockoutBadCount' } else { 'secpol.cfg konnte nicht gelesen werden.' } } catch { 'secedit.exe fehlgeschlagen.' }"; Type="PowerShell"; Category="Lokale Sicherheit"},
    @{Name="Audit-Policy Einstellungen"; Command="auditpol.exe /get /category:*"; Type="CMD"; Category="Lokale Sicherheit"},
    @{Name="RDP-Status"; Command="try { Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -Name 'fDenyTSConnections' -ErrorAction Stop } catch { 'RDP-Einstellung nicht gefunden.' }"; Type="PowerShell"; Category="Lokale Sicherheit"},
    @{Name="Angemeldete Benutzer"; Command="try { quser } catch { 'quser.exe nicht verfügbar.' }"; Type="CMD"; Category="Lokale Sicherheit"},

    # === ACTIVE DIRECTORY SICHERHEIT ===
    @{Name="Domänen-Admins Mitglieder (rekursiv)"; Command="try { Get-ADGroupMember -Identity 'Domain Admins' -Recursive | Get-ADUser -Properties SamAccountName,Enabled,LastLogonDate | Select-Object Name,SamAccountName,Enabled,LastLogonDate } catch { 'AD-Modul nicht verfügbar oder Gruppe nicht gefunden.' }"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active Directory Sicherheit"},
    @{Name="Kerberoastable Accounts"; Command="try { Get-ADUser -Filter {ServicePrincipalName -ne `$null -and Enabled -eq `$true} -Properties ServicePrincipalName,PasswordLastSet | Select-Object Name,ServicePrincipalName,PasswordLastSet } catch { 'AD-Modul nicht verfügbar.' }"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active Directory Sicherheit"},
    @{Name="Inaktive AD-Konten (>90 Tage)"; Command="try { Search-ADAccount -UsersOnly -AccountInactive -TimeSpan 90.00:00:00 | Select-Object Name,SamAccountName,LastLogonDate } catch { 'AD-Modul nicht verfügbar.' }"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active Directory Sicherheit"},
    @{Name="Domänen-Trusts"; Command="try { Get-ADTrust -Filter * | Select-Object Name,TrustType,Direction,IsTransitive } catch { 'AD-Modul nicht verfügbar.' }"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active Directory Sicherheit"},
    @{Name="LAPS Passwortablauf"; Command="try { Get-ADComputer -Filter * -Properties ms-Mcs-AdmPwdExpirationTime | Where-Object { $_.'ms-Mcs-AdmPwdExpirationTime' } | Select-Object Name,@{n='LAPSExpiry';e={[datetime]::FromFileTime($_.'ms-Mcs-AdmPwdExpirationTime')}} } catch { 'AD-Modul nicht verfügbar oder LAPS-Attribut nicht gefunden.' }"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active Directory Sicherheit"},
    @{Name="SPN-Enumeration (Top 20)"; Command="try { Get-ADObject -Filter {servicePrincipalName -ne `$null} -Properties servicePrincipalName | Select-Object Name,servicePrincipalName -First 20 } catch { 'AD-Modul nicht verfügbar.' }"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active Directory Sicherheit"},

    # === NETZWERK & FIREWALL SICHERHEIT ===
    @{Name="Offene Netzwerkports (Listen)"; Command="Get-NetTCPConnection -State Listen | Select-Object LocalAddress,LocalPort,OwningProcess | Sort-Object LocalPort"; Type="PowerShell"; Category="Netzwerk Sicherheit"},
    @{Name="SMBv1 Konfiguration"; Command="try { Get-SmbServerConfiguration | Select-Object EnableSMB1Protocol,RequireSecuritySignature,EncryptData } catch { 'SMB-Konfiguration nicht abrufbar.' }"; Type="PowerShell"; Category="Netzwerk Sicherheit"},
    @{Name="Netzwerk-Freigaben mit Berechtigungen"; Command="try { Get-SmbShare | Select-Object Name,Path,Description,ScopeName,EncryptData, @{n='Permissions';e={(Get-SmbShareAccess -Name $_.Name | Select-Object AccountName,AccessControlType,AccessRight) -join '; '}} } catch { 'SMB-Freigaben nicht abrufbar.' }"; Type="PowerShell"; Category="Netzwerk Sicherheit"},
    @{Name="Potenziell unsichere Inbound-Firewall-Regeln"; Command="Get-NetFirewallRule -Direction Inbound -Enabled True | Where-Object { ($_.RemoteAddress -eq 'Any') -and ($_.Action -eq 'Allow') } | Select-Object DisplayName,Action,RemoteAddress,LocalPort,Profile | Sort-Object LocalPort"; Type="PowerShell"; Category="Netzwerk Sicherheit"},

    # === SYSTEM & POLICY ===
    @{Name="Resultant Set of Policy (RSoP)"; Command="try { `$filePath = Join-Path -Path `$env:TEMP -ChildPath 'RSoP.html'; gpresult /h `$filePath /f | Out-Null; if (Test-Path `$filePath) { 'Bericht erstellt unter: ' + `$filePath } else { 'gpresult fehlgeschlagen.' } } catch { 'gpresult fehlgeschlagen.' }"; Type="PowerShell"; Category="System & Policy"}
)

# Variable für die Sicherheitsaudit-Ergebnisse
$global:securityAuditResults = @{}

#  Verbindungsaudit-Befehle
$connectionAuditCommands = @(
    # === NETZWERK-VERBINDUNGSBAUM ===
    @{Name="Verbindungsbaum (Aktive TCP)"; Command="Get-NetTCPConnection | Where-Object { `$_.State -eq 'Established' } | ForEach-Object { `$proc = Get-Process -Id `$_.OwningProcess -ErrorAction SilentlyContinue; [PSCustomObject]@{ LocalIP=`$_.LocalAddress; LocalPort=`$_.LocalPort; RemoteIP=`$_.RemoteAddress; RemotePort=`$_.RemotePort; Process=if(`$proc){`$proc.ProcessName}else{'N/A'}; PID=`$_.OwningProcess; User=if(`$proc){try{`$proc.StartInfo.UserName}catch{'System'}}else{'N/A'} } } | Sort-Object Process, RemoteIP"; Type="PowerShell"; Category="Verbindungsbaum"},
    @{Name="Netzwerk-Topologie-Map"; Command="`$adapters = Get-NetAdapter | Where-Object Status -eq 'Up'; `$routes = Get-NetRoute | Where-Object RouteMetric -lt 500; `$topology = @(); foreach(`$adapter in `$adapters) { `$ip = Get-NetIPAddress -InterfaceIndex `$adapter.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1; if(`$ip) { `$gateway = `$routes | Where-Object { `$_.InterfaceIndex -eq `$adapter.InterfaceIndex -and `$_.DestinationPrefix -eq '0.0.0.0/0' } | Select-Object -First 1; `$topology += [PSCustomObject]@{ Interface=`$adapter.Name; MAC=`$adapter.MacAddress; IP=`$ip.IPAddress; Subnet=`$ip.PrefixLength; Gateway=if(`$gateway){`$gateway.NextHop}else{'N/A'}; Speed=`$adapter.LinkSpeed } } } `$topology | Format-Table -AutoSize"; Type="PowerShell"; Category="Verbindungsbaum"},
    @{Name="Prozess-Netzwerk-Zuordnung (Erweitert)"; Command="Get-NetTCPConnection | Group-Object OwningProcess | ForEach-Object { `$proc = Get-Process -Id `$_.Name -ErrorAction SilentlyContinue; `$connections = `$_.Group; [PSCustomObject]@{ PID=`$_.Name; ProcessName=if(`$proc){`$proc.ProcessName}else{'Unknown'}; ProcessPath=if(`$proc){try{`$proc.MainModule.FileName}catch{'N/A'}}else{'N/A'}; AnzahlVerbindungen=`$connections.Count; AktiveVerbindungen=(`$connections | Where-Object State -eq 'Established').Count; LauschtPorts=(`$connections | Where-Object State -eq 'Listen').Count; ExterneIPs=(`$connections | Where-Object { `$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::' } | Select-Object -ExpandProperty RemoteAddress -Unique | Measure-Object).Count; StartZeit=if(`$proc){`$proc.StartTime}else{'N/A'} } } | Sort-Object AnzahlVerbindungen -Descending"; Type="PowerShell"; Category="Verbindungsbaum"},

    # === AKTIVE NETZWERKVERBINDUNGEN (Erweitert) ===
    @{Name="Alle TCP-Verbindungen (Performance)"; Command="Get-NetTCPConnection | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess, CreationTime | Sort-Object State, LocalPort"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="Etablierte TCP-Verbindungen"; Command="Get-NetTCPConnection -State Established | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, OwningProcess, CreationTime"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="Lauschende Ports (Listen)"; Command="Get-NetTCPConnection -State Listen | Select-Object LocalAddress, LocalPort, OwningProcess | Sort-Object LocalPort"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="UDP-Endpunkte"; Command="Get-NetUDPEndpoint | Select-Object LocalAddress, LocalPort, OwningProcess | Sort-Object LocalPort"; Type="PowerShell"; Category="UDP-Connections"},
    @{Name="Externe Verbindungen (Internet)"; Command="Get-NetTCPConnection | Where-Object {`$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::'} | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess"; Type="PowerShell"; Category="External-Connections"},
    
    # === ERWEITERTE PROZESS-NETZWERK ANALYSE ===
    @{Name="Top-Prozesse nach Verbindungen"; Command="Get-NetTCPConnection | Group-Object OwningProcess | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object { `$proc = Get-Process -Id `$_.Name -ErrorAction SilentlyContinue; [PSCustomObject]@{ PID=`$_.Name; Process=if(`$proc){`$proc.ProcessName}else{'Unknown'}; Verbindungen=`$_.Count; Established=(`$_.Group | Where-Object State -eq 'Established').Count; Listen=(`$_.Group | Where-Object State -eq 'Listen').Count } }"; Type="PowerShell"; Category="Process-Network"},
    @{Name="Verdächtige Prozess-Verbindungen"; Command="Get-NetTCPConnection | Where-Object { `$_.State -eq 'Established' -and `$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.' } | ForEach-Object { `$proc = Get-Process -Id `$_.OwningProcess -ErrorAction SilentlyContinue; if(`$proc -and `$proc.Path -notmatch 'Windows|Program Files') { [PSCustomObject]@{ Process=`$proc.ProcessName; PID=`$_.OwningProcess; Path=try{`$proc.MainModule.FileName}catch{'N/A'}; RemoteIP=`$_.RemoteAddress; RemotePort=`$_.RemotePort; Company=try{`$proc.Company}catch{'N/A'} } } } | Sort-Object Process"; Type="PowerShell"; Category="Process-Network"},
    @{Name="Netzwerk-Statistiken pro Prozess"; Command="Get-Process | Where-Object { `$_.Id -ne 0 } | ForEach-Object { `$pid = `$_.Id; `$tcpConns = @(Get-NetTCPConnection | Where-Object OwningProcess -eq `$pid); `$udpConns = @(Get-NetUDPEndpoint | Where-Object OwningProcess -eq `$pid); if(`$tcpConns.Count -gt 0 -or `$udpConns.Count -gt 0) { [PSCustomObject]@{ ProcessName=`$_.ProcessName; PID=`$pid; TCP_Verbindungen=`$tcpConns.Count; UDP_Verbindungen=`$udpConns.Count; CPU_Percent=0; RAM_MB=[math]::Round(`$_.WorkingSet64/1MB,2) } } } | Sort-Object TCP_Verbindungen -Descending"; Type="PowerShell"; Category="Process-Network"},

    # === LOKALE GERÄTE UND NETZWERK-ERKENNUNG ===
    @{Name="ARP-Cache (Lokale Geräte)"; Command="Get-NetNeighbor | Where-Object State -ne 'Unreachable' | Select-Object IPAddress, MacAddress, State, InterfaceAlias | Sort-Object IPAddress"; Type="PowerShell"; Category="Local-Devices"},
    @{Name="MAC-Adressen-Analyse"; Command="Get-NetNeighbor | Where-Object { `$_.MacAddress -ne '00-00-00-00-00-00' -and `$_.State -ne 'Unreachable' } | ForEach-Object { `$vendor = switch -Regex (`$_.MacAddress.Substring(0,8)) { '^00-50-56' {'VMware'}; '^00-0C-29' {'VMware'}; '^08-00-27' {'VirtualBox'}; '^00-15-5D' {'Microsoft Hyper-V'}; '^00-1B-21' {'Dell'}; '^00-25-90' {'Dell'}; '^D4-BE-D9' {'Dell'}; '^B8-2A-72' {'Dell'}; '^70-B3-D5' {'HP'}; '^3C-D9-2B' {'HP'}; '^94-57-A5' {'HP'}; default {'Unknown'} }; [PSCustomObject]@{ IP=`$_.IPAddress; MAC=`$_.MacAddress; Vendor=`$vendor; State=`$_.State; Interface=`$_.InterfaceAlias } } | Sort-Object Vendor, IP"; Type="PowerShell"; Category="Local-Devices"},
    @{Name="DHCP IPv4 Bereiche"; Command="Get-DhcpServerv4Scope | Select-Object ScopeId, Name, StartRange, EndRange, SubnetMask, State, LeaseDuration"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP IPv6 Bereiche"; Command="Get-DhcpServerv6Scope | Select-Object Prefix, Name, State, PreferredLifetime"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Reservierungen"; Command="try { `$scopes = Get-DhcpServerv4Scope -ErrorAction Stop; if (`$scopes) { `$scopes | ForEach-Object { Get-DhcpServerv4Reservation -ScopeId `$_.ScopeId -ErrorAction SilentlyContinue } | Select-Object ScopeId, IPAddress, ClientId, Name, Description } else { 'Keine DHCP-Bereiche konfiguriert' } } catch { 'DHCP-Reservierungen: DHCP-Server nicht verfügbar oder keine Bereiche konfiguriert' }"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Lease-Statistiken"; Command="try { Get-DhcpServerv4ScopeStatistics -ErrorAction Stop | Select-Object ScopeId, AddressesFree, AddressesInUse, PercentageInUse } catch { 'DHCP-Lease-Statistiken: Keine Bereiche konfiguriert oder DHCP-Server nicht verfügbar' }"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="Wireless-Netzwerke (Falls verfügbar)"; Command="try { netsh wlan show profiles } catch { 'Wireless-Adapter nicht verfügbar' }"; Type="CMD"; Category="Local-Devices"},

    # === DNS UND NETZWERK-AUFLÖSUNG ===
    @{Name="DNS-Cache-Analyse"; Command="Get-DnsClientCache | Where-Object { `$_.Type -eq 'A' } | Select-Object Name, Data, TTL, Section | Sort-Object Name"; Type="PowerShell"; Category="DNS-Info"},
    @{Name="DNS-Server-Konfiguration"; Command="Get-DnsClientServerAddress | Where-Object { `$_.ServerAddresses.Count -gt 0 } | Select-Object InterfaceAlias, @{Name='DNS_Server';Expression={`$_.ServerAddresses -join ', '}} | Sort-Object InterfaceAlias"; Type="PowerShell"; Category="DNS-Info"},
    @{Name="Reverse-DNS-Lookup (Top-IPs)"; Command="`$topIPs = Get-NetTCPConnection | Where-Object { `$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::' } | Group-Object RemoteAddress | Sort-Object Count -Descending | Select-Object -First 10; `$topIPs | ForEach-Object { `$ip = `$_.Name; try { `$hostname = [System.Net.Dns]::GetHostEntry(`$ip).HostName } catch { `$hostname = 'N/A' }; [PSCustomObject]@{ IP=`$ip; Hostname=`$hostname; Verbindungen=`$_.Count } } | Sort-Object Verbindungen -Descending"; Type="PowerShell"; Category="DNS-Info"},
    @{Name="Domänen-DNS-Informationen"; Command="nslookup `$env:USERDNSDOMAIN 2>null | Select-String -Pattern 'Address|Server'"; Type="CMD"; Category="DNS-Info"},
    @{Name="DNS Forwarders"; Command="Get-DnsServerForwarder"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},

    # === GEO-IP UND EXTERNE ANALYSE ===
    @{Name="Geo-IP-Analyse (Externe IPs)"; Command="`$externalIPs = Get-NetTCPConnection | Where-Object { `$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::' } | Select-Object -ExpandProperty RemoteAddress -Unique | Select-Object -First 5; `$externalIPs | ForEach-Object { `$ip = `$_; [PSCustomObject]@{ IP=`$ip; Land='Online-Analyse erforderlich'; Region='API-Limit'; Stadt='Verfügbar'; ISP='ipinfo.io'; Hostname='Manual-Check' } }"; Type="PowerShell"; Category="Geo-IP"},
    @{Name="Bedrohungsanalyse (Blacklist-Check)"; Command="`$suspiciousIPs = Get-NetTCPConnection | Where-Object { `$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::' } | Select-Object -ExpandProperty RemoteAddress -Unique; `$suspiciousIPs | ForEach-Object { `$ip = `$_; `$isPrivate = `$ip -match '^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.'; `$isSuspicious = `$ip -match '^(185\.243\.|91\.234\.|77\.83\.)' -or `$ip -match '^(193\.0\.14\.|208\.67\.222\.)'; [PSCustomObject]@{ IP=`$ip; Type=if(`$isPrivate){'Privat'}elseif(`$isSuspicious){'⚠️ Verdächtig'}else{'Public'}; Port_Count=(Get-NetTCPConnection | Where-Object RemoteAddress -eq `$ip).Count } } | Sort-Object Type, IP"; Type="PowerShell"; Category="Geo-IP"},

    # === FIREWALL UND SICHERHEIT ===
    @{Name="Windows-Firewall-Status"; Command="Get-NetFirewallProfile | Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction, LogFileName"; Type="PowerShell"; Category="Firewall-Logs"},
    @{Name="Firewall-Regeln (Aktiv)"; Command="Get-NetFirewallRule | Where-Object { `$_.Enabled -eq 'True' -and `$_.Direction -eq 'Inbound' } | Select-Object DisplayName, Direction, Action, Protocol, LocalPort | Sort-Object Protocol, LocalPort | Select-Object -First 50"; Type="PowerShell"; Category="Firewall-Logs"},
    @{Name="Firewall-Verbindungslogs"; Command="try { `$events = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Windows Firewall With Advanced Security/Firewall'} -MaxEvents 50 -ErrorAction Stop 2>`$null; if(`$events) { `$events | Select-Object TimeCreated, Id, LevelDisplayName, Message | Sort-Object TimeCreated -Descending } else { 'Keine Firewall-Events gefunden' } } catch [System.Exception] { 'Firewall-Logs nicht verfügbar oder deaktiviert - ' + `$_.Exception.Message.Split('.')[0] }"; Type="PowerShell"; Category="Firewall-Logs"},
    @{Name="Blockierte Verbindungen"; Command="try { `$events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=5157} -MaxEvents 20 -ErrorAction Stop 2>`$null; if(`$events) { `$events | ForEach-Object { `$xml = [xml]`$_.ToXml(); [PSCustomObject]@{ Zeit=`$_.TimeCreated; Prozess=`$xml.Event.EventData.Data | Where-Object Name -eq 'Application' | Select-Object -ExpandProperty '#text'; Quelle=`$xml.Event.EventData.Data | Where-Object Name -eq 'SourceAddress' | Select-Object -ExpandProperty '#text'; Ziel=`$xml.Event.EventData.Data | Where-Object Name -eq 'DestAddress' | Select-Object -ExpandProperty '#text'; Port=`$xml.Event.EventData.Data | Where-Object Name -eq 'DestPort' | Select-Object -ExpandProperty '#text' } } | Sort-Object Zeit -Descending } else { 'Keine blockierten Verbindungen (Event-ID 5157) gefunden' } } catch [System.Exception] { 'Sicherheitslogs für blockierte Verbindungen nicht verfügbar - Event-ID 5157 nicht aktiviert oder keine Events vorhanden' }"; Type="PowerShell"; Category="Firewall-Logs"},

    # === NETZWERK-EVENTS UND MONITORING ===
    @{Name="Netzwerk-Sicherheitsereignisse"; Command="try { `$events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=5156} -MaxEvents 30 -ErrorAction Stop 2>`$null; if(`$events) { `$events | Select-Object TimeCreated, Id, LevelDisplayName, Message | Sort-Object TimeCreated -Descending } else { 'Keine Netzwerk-Sicherheitsereignisse (Event-ID 5156) gefunden - Logging möglicherweise deaktiviert' } } catch [System.Exception] { 'Netzwerk-Sicherheitslogs nicht verfügbar - Event-ID 5156 erfordert aktivierte Firewall-Logging-Richtlinie' }"; Type="PowerShell"; Category="Network-Events"},
    @{Name="Netzwerk-Adapter-Events"; Command="try { `$events = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Kernel-Network/Analytic'} -MaxEvents 20 -ErrorAction Stop 2>`$null; if(`$events) { `$events | Select-Object TimeCreated, Id, LevelDisplayName, Message | Sort-Object TimeCreated -Descending } else { 'Keine Kernel-Network-Events gefunden - Analytisches Log möglicherweise deaktiviert' } } catch [System.Exception] { 'Kernel-Network-Logs nicht verfügbar - Analytische Logs müssen in der Ereignisanzeige aktiviert werden' }"; Type="PowerShell"; Category="Network-Events"},
    @{Name="Prozess-Netzwerk-Events"; Command="try { `$events = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Kernel-Process/Analytic'} -MaxEvents 25 -ErrorAction Stop 2>`$null; if(`$events) { `$networkEvents = `$events | Where-Object { `$_.Message -like '*network*' -or `$_.Message -like '*socket*' }; if(`$networkEvents) { `$networkEvents | Select-Object TimeCreated, Id, ProcessId, Message | Sort-Object TimeCreated -Descending } else { 'Keine Prozess-Netzwerk-Events in den letzten 25 Events gefunden' } } else { 'Keine Kernel-Process-Events gefunden - Analytisches Log möglicherweise deaktiviert' } } catch [System.Exception] { 'Prozess-Events nicht verfügbar - Analytische Logs müssen in der Ereignisanzeige aktiviert werden' }"; Type="PowerShell"; Category="Network-Events"},

    # === ACTIVE DIRECTORY UND DOMÄNEN-INFORMATIONEN ===
    @{Name="Domänen-Controller-Informationen"; Command="try { Get-ADDomainController -Discover -Service ADWS,KDC,TimeService | Select-Object Name, IPv4Address, Site, OperatingSystem, Domain } catch { try { nltest /dclist:`$env:USERDNSDOMAIN } catch { 'AD-Modul nicht verfügbar' } }"; Type="PowerShell"; Category="Domain-Users"},
    @{Name="Privilegierte AD-Gruppen"; Command="try { @('Domain Admins', 'Enterprise Admins', 'Schema Admins', 'Administrators') | ForEach-Object { `$group = `$_; try { Get-ADGroupMember -Identity `$group | Select-Object @{Name='Group';Expression={`$group}}, Name, SamAccountName, objectClass } catch { [PSCustomObject]@{Group=`$group; Name='Gruppe nicht gefunden'; SamAccountName='N/A'; objectClass='N/A'} } } } catch { 'AD-PowerShell-Modul nicht verfügbar' }"; Type="PowerShell"; Category="Domain-Users"},
    @{Name="Kürzliche AD-Anmeldungen"; Command="try { `$events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4624} -MaxEvents 50 -ErrorAction Stop 2>`$null; if(`$events) { `$filtered = `$events | Where-Object { `$_.Message -notlike '*ANONYMOUS*' }; if(`$filtered) { `$filtered | ForEach-Object { `$msg = `$_.Message; `$user = if(`$msg -match 'Account Name:\\s+([^\\r\\n]+)') { `$matches[1] } else { 'Unknown' }; `$workstation = if(`$msg -match 'Workstation Name:\\s+([^\\r\\n]+)') { `$matches[1] } else { 'Unknown' }; [PSCustomObject]@{ Zeit=`$_.TimeCreated; Benutzer=`$user; Workstation=`$workstation; LogonType=if(`$msg -match 'Logon Type:\\s+(\\d+)') { `$matches[1] } else { 'Unknown' } } } } | Where-Object { `$_.Benutzer -ne '-' -and `$_.Benutzer -ne 'ANONYMOUS LOGON' } | Sort-Object Zeit -Descending | Select-Object -First 20 } else { 'Keine relevanten Anmelde-Events (ohne ANONYMOUS) gefunden' } } else { 'Keine Anmelde-Events (Event-ID 4624) gefunden' } } catch [System.Exception] { 'Anmelde-Sicherheitslogs nicht verfügbar oder deaktiviert' }"; Type="PowerShell"; Category="Domain-Users"},
    @{Name="LDAP-Verbindungstests"; Command="try { `$domain = `$env:USERDNSDOMAIN; if(`$domain) { `$dcIP = (nslookup `$domain 2>null | Select-String -Pattern '\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}' | Select-Object -First 1).Matches.Value; if(`$dcIP) { Test-NetConnection -ComputerName `$dcIP -Port 389; Test-NetConnection -ComputerName `$dcIP -Port 636 } else { 'Domain-Controller-IP nicht ermittelbar' } } else { 'Nicht in einer Domäne' } } catch { 'LDAP-Test fehlgeschlagen' }"; Type="PowerShell"; Category="Domain-Users"},

    # === REMOTE-SESSIONS UND RDP ===
    @{Name="Remote-Desktop-Verbindungen"; Command="try { `$events = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'} -MaxEvents 30 -ErrorAction Stop 2>`$null; if(`$events) { `$events | Select-Object TimeCreated, Id, LevelDisplayName, Message | Sort-Object TimeCreated -Descending } else { 'Keine Terminal-Services-Events gefunden' } } catch [System.Exception] { 'Terminal-Services-Logs nicht verfügbar oder deaktiviert' }"; Type="PowerShell"; Category="Remote-Sessions"},
    @{Name="SMB-Verbindungen"; Command="try { Get-SmbConnection | Select-Object ServerName, ShareName, UserName, Dialect } catch { 'SMB-Informationen nicht verfügbar' }"; Type="PowerShell"; Category="Remote-Sessions"},
    @{Name="SMB-Freigaben-Zugriffe"; Command="try { `$events = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-SmbServer/Security'} -MaxEvents 50 -ErrorAction Stop 2>`$null; if(`$events) { `$events | Select-Object TimeCreated, Id, Message | Sort-Object TimeCreated -Descending } else { 'Keine SMB-Server-Security-Events gefunden' } } catch [System.Exception] { 'SMB-Security-Logs nicht verfügbar oder deaktiviert' }"; Type="PowerShell"; Category="Remote-Sessions"},
    @{Name="Aktive Terminal-Sessions"; Command="try { quser 2>`$null } catch { try { query session } catch { 'Terminal-Session-Abfrage nicht verfügbar' } }"; Type="CMD"; Category="Remote-Sessions"},

    # === ERWEITERTE NETZWERK-TOPOLOGIE ===
    @{Name="Routing-Tabelle (Detailliert)"; Command="Get-NetRoute | Select-Object DestinationPrefix, NextHop, InterfaceAlias, RouteMetric, Protocol, @{Name='NetworkCategory';Expression={if(`$_.DestinationPrefix -eq '0.0.0.0/0'){'Default Gateway'}elseif(`$_.DestinationPrefix -like '169.254.*'){'APIPA'}elseif(`$_.DestinationPrefix -like '224.*'){'Multicast'}else{'Network Route'}}} | Sort-Object RouteMetric, DestinationPrefix"; Type="PowerShell"; Category="Network-Topology"},
    @{Name="Netzwerk-Interface-Statistiken"; Command="Get-NetAdapterStatistics | Select-Object Name, @{Name='Empfangen_GB';Expression={[math]::Round(`$_.BytesReceived/1GB,3)}}, @{Name='Gesendet_GB';Expression={[math]::Round(`$_.BytesSent/1GB,3)}}, @{Name='Pakete_Empfangen';Expression={`$_.PacketsReceived}}, @{Name='Pakete_Gesendet';Expression={`$_.PacketsSent}}, @{Name='Fehler_Eingehend';Expression={`$_.InboundErrors}}, @{Name='Fehler_Ausgehend';Expression={`$_.OutboundErrors}} | Sort-Object Empfangen_GB -Descending"; Type="PowerShell"; Category="Network-Topology"},
    @{Name="Gateway- und DNS-Konfiguration"; Command="Get-NetIPConfiguration | Where-Object {`$_.IPv4DefaultGateway -or `$_.IPv6DefaultGateway} | Select-Object InterfaceAlias, @{Name='IPv4_Adresse';Expression={(`$_.IPv4Address | Select-Object -First 1).IPAddress}}, @{Name='IPv4_Gateway';Expression={(`$_.IPv4DefaultGateway | Select-Object -First 1).NextHop}}, @{Name='DNS_Server';Expression={`$_.DNSServer.ServerAddresses -join '; '}}, @{Name='DHCP_Aktiviert';Expression={`$_.NetProfile.NetworkCategory}} | Sort-Object InterfaceAlias"; Type="PowerShell"; Category="Network-Topology"},
    @{Name="Netzwerk-Troubleshooting-Infos"; Command="`$networkInfo = @(); Get-NetAdapter | Where-Object Status -eq 'Up' | ForEach-Object { `$adapter = `$_; `$tcpStats = Get-NetTCPConnection | Where-Object { try { (Get-NetAdapter -InterfaceIndex `$_.LocalAddress -ErrorAction SilentlyContinue).InterfaceIndex -eq `$adapter.InterfaceIndex } catch { `$false } }; `$networkInfo += [PSCustomObject]@{ Interface=`$adapter.Name; MAC=`$adapter.MacAddress; Status=`$adapter.Status; LinkSpeed=`$adapter.LinkSpeed; MediaType=`$adapter.MediaType; TCP_Verbindungen=(`$tcpStats | Measure-Object).Count; Typ=if(`$adapter.Virtual){'Virtual'}else{'Physical'} } }; `$networkInfo | Sort-Object Interface"; Type="PowerShell"; Category="Network-Topology"}
)

# Variable fuer die Verbindungsaudit-Ergebnisse
$global:connectionAuditResults = @{}

# Erweiterte Netzwerk-Verbindungsbaum-Analyse-Funktionen
function Get-NetworkConnectionTree {
    Write-DebugLog "Erstelle erweiterten Netzwerk-Verbindungsbaum" "ConnectionAudit"
    
    try {
        # Sammle alle aktiven Verbindungen
        $tcpConnections = Get-NetTCPConnection | Where-Object { $_.State -eq 'Established' }
        $udpConnections = Get-NetUDPEndpoint # Diese Variable wird aktuell nicht weiter verwendet, könnte für zukünftige Erweiterungen sein
        $processes = Get-Process
        
        # Erstelle Verbindungsbaum-Struktur
        $connectionTree = @{
            Timestamp = Get-Date
            ServerInfo = @{
                ComputerName = $env:COMPUTERNAME
                Domain = $env:USERDNSDOMAIN
                User = $env:USERNAME
                OS = (Get-CimInstance Win32_OperatingSystem).Caption
            }
            NetworkTopology = @{}
            ActiveConnections = @()
            ProcessMapping = @{} # Diese Struktur wird gefüllt, aber nicht explizit im Rückgabewert verwendet, könnte implizit durch $processInfo sein
            ExternalConnections = @()
            SecurityAnalysis = @{} # Diese Struktur wird initialisiert, aber nicht gefüllt
        }
        
        # Netzwerk-Topologie analysieren
        $adapters = Get-NetAdapter | Where-Object Status -eq 'Up'
        foreach ($adapter in $adapters) {
            $ipConfig = Get-NetIPConfiguration -InterfaceIndex $adapter.InterfaceIndex -ErrorAction SilentlyContinue
            if ($ipConfig) {
                $connectionTree.NetworkTopology[$adapter.Name] = @{
                    MAC = $adapter.MacAddress
                    IP = ($ipConfig.IPv4Address | Select-Object -First 1).IPAddress
                    Gateway = ($ipConfig.IPv4DefaultGateway | Select-Object -First 1).NextHop
                    DNS = $ipConfig.DNSServer.ServerAddresses -join ', '
                    Speed = $adapter.LinkSpeed
                    Type = if ($adapter.Virtual) { "Virtual" } else { "Physical" }
                }
            }
        }
        
        # Prozess-Mapping erstellen
        foreach ($connection in $tcpConnections) {
            $process = $processes | Where-Object Id -eq $connection.OwningProcess | Select-Object -First 1
            $processInfo = if ($process) {
                @{
                    Name = $process.ProcessName
                    Path = $(try { $process.MainModule.FileName } catch { "N/A" })
                    StartTime = $process.StartTime
                    Company = $(try { $process.Company } catch { "N/A" })
                }
            } else {
                @{ Name = "Unknown"; Path = "N/A"; StartTime = "N/A"; Company = "N/A" }
            }
            
            $connectionInfo = @{
                LocalAddress = $connection.LocalAddress
                LocalPort = $connection.LocalPort
                RemoteAddress = $connection.RemoteAddress
                RemotePort = $connection.RemotePort
                State = $connection.State
                ProcessInfo = $processInfo
                IsExternal = $connection.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:'
            }
            
            $connectionTree.ActiveConnections += $connectionInfo
            
            # Externe Verbindungen separat sammeln
            if ($connectionInfo.IsExternal -and $connection.RemoteAddress -ne '0.0.0.0' -and $connection.RemoteAddress -ne '::') { # Zusätzliche Prüfung für IPv6 unspecified
                $connectionTree.ExternalConnections += $connectionInfo
            }
        }
        
        return $connectionTree
    }
    catch {
        Write-DebugLog "FEHLER beim Erstellen des Verbindungsbaums: $($_.Exception.Message) $($_.ScriptStackTrace)" "ConnectionAudit"
        return $null
    }
}

function Format-ConnectionTreeHTML {
    param(
        [hashtable]$ConnectionTree,
        [hashtable]$Results
    )
    
    if (-not $ConnectionTree) {
        return "<p>Verbindungsbaum konnte nicht erstellt werden.</p>"
    }
    
    $html = @"
<div class="connection-tree-container">
    <h2>🌐 Netzwerk-Verbindungsbaum - $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</h2>
    
    <div class="server-summary">
        <h3>📊 Server-Übersicht</h3>
        <div class="info-grid">
            <div class="info-item"><strong>Server:</strong> $($ConnectionTree.ServerInfo.ComputerName)</div>
            <div class="info-item"><strong>Domäne:</strong> $($ConnectionTree.ServerInfo.Domain)</div>
            <div class="info-item"><strong>Benutzer:</strong> $($ConnectionTree.ServerInfo.User)</div>
            <div class="info-item"><strong>OS:</strong> $($ConnectionTree.ServerInfo.OS)</div>
        </div>
    </div>
    
    <div class="network-topology">
        <h3>🔗 Netzwerk-Topologie</h3>
        <table class="styled-table">
            <thead>
                <tr><th>Interface</th><th>IP-Adresse</th><th>Gateway</th><th>DNS</th><th>Typ</th><th>Geschwindigkeit</th></tr>
            </thead>
            <tbody>
"@
    
    foreach ($interface in $ConnectionTree.NetworkTopology.Keys) {
        $topo = $ConnectionTree.NetworkTopology[$interface]
        $html += @"
                <tr>
                    <td>$interface</td>
                    <td>$($topo.IP)</td>
                    <td>$($topo.Gateway)</td>
                    <td>$($topo.DNS)</td>
                    <td>$($topo.Type)</td>
                    <td>$($topo.Speed)</td>
                </tr>
"@
    }
    
    $html += @"
            </tbody>
        </table>
    </div>
    
    <div class="active-connections">
        <h3>⚡ Aktive Verbindungen ($(($ConnectionTree.ActiveConnections | Measure-Object).Count))</h3>
        <table class="styled-table">
            <thead>
                <tr><th>Prozess</th><th>Lokal</th><th>Remote</th><th>Status</th><th>Typ</th><th>Firma</th></tr>
            </thead>
            <tbody>
"@
    
    foreach ($conn in ($ConnectionTree.ActiveConnections | Sort-Object { $_.ProcessInfo.Name })) {
        $connectionType = if ($conn.IsExternal) { "🌍 Extern" } else { "🏠 Lokal" }
        $html += @"
                <tr class="$(if ($conn.IsExternal) { 'external-connection' } else { 'local-connection' })">
                    <td><strong>$($conn.ProcessInfo.Name)</strong></td>
                    <td>$($conn.LocalAddress):$($conn.LocalPort)</td>
                    <td>$($conn.RemoteAddress):$($conn.RemotePort)</td>
                    <td>$($conn.State)</td>
                    <td>$connectionType</td>
                    <td>$($conn.ProcessInfo.Company)</td>
                </tr>
"@
    }
    
    $html += @"
            </tbody>
        </table>
    </div>
    
    <div class="external-analysis">
        <h3>🌍 Externe Verbindungen - Sicherheitsanalyse</h3>
        <p><strong>Anzahl externer Verbindungen:</strong> $(($ConnectionTree.ExternalConnections | Measure-Object).Count)</p>
        <table class="styled-table">
            <thead>
                <tr><th>Remote-IP</th><th>Port</th><th>Prozess</th><th>Pfad</th><th>Bewertung</th></tr>
            </thead>
            <tbody>
"@
    
    foreach ($extConn in ($ConnectionTree.ExternalConnections | Sort-Object RemoteAddress)) {
        $risk = "✅ Normal"
        if ($extConn.ProcessInfo.Path -notmatch "Windows|Program Files") {
            $risk = "⚠️ Prüfen"
        }
        if ($extConn.RemoteAddress -match "^(185\.243\.|91\.234\.|77\.83\.)") {
            $risk = "🚨 Verdächtig"
        }
        
        $html += @"
                <tr class="$(if ($risk -eq '🚨 Verdächtig') { 'suspicious-connection' } elseif ($risk -eq '⚠️ Prüfen') { 'warning-connection' } else { 'normal-connection' })">
                    <td>$($extConn.RemoteAddress)</td>
                    <td>$($extConn.RemotePort)</td>
                    <td>$($extConn.ProcessInfo.Name)</td>
                    <td>$($extConn.ProcessInfo.Path)</td>
                    <td>$risk</td>
                </tr>
"@
    }
    
    $html += @"
            </tbody>
        </table>
    </div>
</div>

<style>
.connection-tree-container { margin: 20px 0; }
.server-summary { background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 15px 0; }
.info-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 10px; }
.info-item { padding: 8px; background: white; border-radius: 4px; border-left: 4px solid #FD7E14; }
.network-topology, .active-connections, .external-analysis { margin: 20px 0; }
.styled-table { width: 100%; border-collapse: collapse; margin: 10px 0; }
.styled-table th { background: #FD7E14; color: white; padding: 12px; text-align: left; }
.styled-table td { padding: 10px; border-bottom: 1px solid #ddd; }
.external-connection { background-color: #fff3cd; }
.local-connection { background-color: #f8f9fa; }
.suspicious-connection { background-color: #f8d7da; }
.warning-connection { background-color: #fff3cd; }
.normal-connection { background-color: #d1f7d1; }
.styled-table tr:hover { background-color: #f5f5f5; }
</style>
"@
    
    return $html
}

# === VERBINDUNGSAUDIT SPEZIFISCHE KOMMANDOS ===
$connectionAuditCommands = @(
    # === AKTIVE NETZWERKVERBINDUNGEN ===
    @{Name="Alle TCP Verbindungen"; Command="Get-NetTCPConnection | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess, CreationTime | Sort-Object State, LocalPort"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="Etablierte TCP Verbindungen"; Command="Get-NetTCPConnection -State Established | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, OwningProcess, CreationTime"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="Lauschende Ports (Listen)"; Command="Get-NetTCPConnection -State Listen | Select-Object LocalAddress, LocalPort, OwningProcess | Sort-Object LocalPort"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="UDP Endpunkte"; Command="Get-NetUDPEndpoint | Select-Object LocalAddress, LocalPort, OwningProcess | Sort-Object LocalPort"; Type="PowerShell"; Category="UDP-Connections"},
    @{Name="Externe Verbindungen (nicht lokal)"; Command="Get-NetTCPConnection | Where-Object {`$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::'} | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess"; Type="PowerShell"; Category="External-Connections"},
    
    # === PROZESS-NETZWERK ZUORDNUNG ===
    @{Name="Prozesse mit Netzwerkverbindungen"; Command="Get-NetTCPConnection | Where-Object {`$_.State -eq 'Established'} | ForEach-Object { `$conn = `$_; try { `$process = Get-Process -Id `$conn.OwningProcess -ErrorAction Stop; [PSCustomObject]@{ ProcessName = `$process.ProcessName; PID = `$process.Id; LocalAddress = `$conn.LocalAddress; LocalPort = `$conn.LocalPort; RemoteAddress = `$conn.RemoteAddress; RemotePort = `$conn.RemotePort; ProcessPath = `$process.Path } } catch { [PSCustomObject]@{ ProcessName = 'Unknown'; PID = `$conn.OwningProcess; LocalAddress = `$conn.LocalAddress; LocalPort = `$conn.LocalPort; RemoteAddress = `$conn.RemoteAddress; RemotePort = `$conn.RemotePort; ProcessPath = 'N/A' } } } | Sort-Object ProcessName"; Type="PowerShell"; Category="Process-Network"},
    @{Name="Top Prozesse nach Verbindungen"; Command="`$connections = Get-NetTCPConnection | Group-Object OwningProcess; `$connections | ForEach-Object { try { `$process = Get-Process -Id `$_.Name -ErrorAction Stop; [PSCustomObject]@{ ProcessName = `$process.ProcessName; PID = `$_.Name; ConnectionCount = `$_.Count; ProcessPath = `$process.Path } } catch { [PSCustomObject]@{ ProcessName = 'Unknown'; PID = `$_.Name; ConnectionCount = `$_.Count; ProcessPath = 'N/A' } } } | Sort-Object ConnectionCount -Descending | Select-Object -First 20"; Type="PowerShell"; Category="Process-Network"},
    @{Name="System-Prozesse mit Netzwerkzugriff"; Command="Get-NetTCPConnection | Where-Object {`$_.OwningProcess -lt 1000} | ForEach-Object { `$conn = `$_; try { `$process = Get-Process -Id `$conn.OwningProcess -ErrorAction Stop; [PSCustomObject]@{ ProcessName = `$process.ProcessName; PID = `$process.Id; LocalPort = `$conn.LocalPort; RemoteAddress = `$conn.RemoteAddress; State = `$conn.State } } catch { [PSCustomObject]@{ ProcessName = 'System/Unknown'; PID = `$conn.OwningProcess; LocalPort = `$conn.LocalPort; RemoteAddress = `$conn.RemoteAddress; State = `$conn.State } } } | Sort-Object PID"; Type="PowerShell"; Category="Process-Network"},
    
    # === LOKALE GERAETE (ARP-CACHE) ===
    @{Name="ARP Cache (alle Geraete)"; Command="Get-NetNeighbor | Select-Object IPAddress, MacAddress, State, InterfaceAlias | Sort-Object IPAddress"; Type="PowerShell"; Category="Local-Devices"},
    @{Name="ARP Cache (nur erreichbare)"; Command="Get-NetNeighbor -State Reachable | Select-Object IPAddress, MacAddress, InterfaceAlias"; Type="PowerShell"; Category="Local-Devices"},
    @{Name="MAC-Adressen im lokalen Netz"; Command="arp -a"; Type="CMD"; Category="Local-Devices"},
    @{Name="Netzwerk-Interfaces"; Command="Get-NetAdapter | Select-Object Name, InterfaceDescription, LinkSpeed, MacAddress, Status | Sort-Object Name"; Type="PowerShell"; Category="Local-Devices"},
    @{Name="DHCP-Leases (falls DHCP-Server)"; Command="if (Get-WindowsFeature -Name DHCP | Where-Object {`$_.Installed}) { Get-DhcpServerv4Lease -AllLeases | Select-Object IPAddress, ClientId, HostName, LeaseExpiryTime | Sort-Object IPAddress } else { 'DHCP-Server nicht installiert' }"; Type="PowerShell"; Category="Local-Devices"},
    
    # === DNS INFORMATIONEN ===
    @{Name="DNS Cache"; Command="Get-DnsClientCache | Select-Object Entry, Name, Data, TimeToLive | Sort-Object Name"; Type="PowerShell"; Category="DNS-Info"},
    @{Name="DNS Server Konfiguration"; Command="Get-DnsClientServerAddress | Select-Object InterfaceAlias, ServerAddresses"; Type="PowerShell"; Category="DNS-Info"},
    @{Name="Reverse DNS fuer externe IPs"; Command="`$extIPs = Get-NetTCPConnection | Where-Object {`$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::'} | Select-Object -ExpandProperty RemoteAddress -Unique; `$extIPs | ForEach-Object { try { `$hostname = [System.Net.Dns]::GetHostEntry(`$_).HostName; [PSCustomObject]@{ IPAddress = `$_; Hostname = `$hostname } } catch { [PSCustomObject]@{ IPAddress = `$_; Hostname = 'Aufloesung fehlgeschlagen' } } } | Sort-Object IPAddress"; Type="PowerShell"; Category="DNS-Info"},
    
    # === GEO-IP INFORMATIONEN ===
    @{Name="Geo-IP Analyse externer Verbindungen"; Command="`$extIPs = Get-NetTCPConnection | Where-Object {`$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::'} | Select-Object -ExpandProperty RemoteAddress -Unique | Select-Object -First 10; `$extIPs | ForEach-Object { try { `$geoInfo = Invoke-RestMethod `"http://ipinfo.io/`$_/json`" -TimeoutSec 5; [PSCustomObject]@{ IPAddress = `$_; Country = `$geoInfo.country; Region = `$geoInfo.region; City = `$geoInfo.city; Organization = `$geoInfo.org; ISP = `$geoInfo.isp } } catch { [PSCustomObject]@{ IPAddress = `$_; Country = 'N/A'; Region = 'N/A'; City = 'N/A'; Organization = 'Abfrage fehlgeschlagen'; ISP = 'N/A' } } }"; Type="PowerShell"; Category="Geo-IP"},
    
    # === FIREWALL UND LOGGING ===
    @{Name="Firewall Verbindungs-Logs"; Command="if (Test-Path `$env:SystemRoot\\system32\\LogFiles\\Firewall\\pfirewall.log) { Get-Content `$env:SystemRoot\\system32\\LogFiles\\Firewall\\pfirewall.log -Tail 50 | Where-Object {`$_ -match 'ALLOW|DROP'} } else { 'Firewall-Logging nicht aktiviert oder Log-Datei nicht gefunden' }"; Type="PowerShell"; Category="Firewall-Logs"},
    @{Name="Firewall Logging Status"; Command="Get-NetFirewallProfile | Select-Object Name, LogAllowed, LogBlocked, LogFileName, LogMaxSizeKilobytes"; Type="PowerShell"; Category="Firewall-Logs"},
    @{Name="Aktive Firewall Regeln"; Command="Get-NetFirewallRule | Where-Object {`$_.Enabled -eq 'True' -and `$_.Action -eq 'Allow'} | Select-Object DisplayName, Direction, Protocol, LocalPort, RemoteAddress | Sort-Object Direction, Protocol"; Type="PowerShell"; Category="Firewall-Logs"},
    
    # === EVENT LOGS FUER VERBINDUNGEN ===
    @{Name="Netzwerk-Events (Security Log)"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Security'; ID=5156,5157,5158} -MaxEvents 50 -ErrorAction Stop | Select-Object TimeCreated, Id, LevelDisplayName, Message } catch { 'Netzwerk-Events (Security Log): Event-IDs 5156/5157/5158 nicht verfügbar oder keine Events vorhanden - Firewall-Logging muss aktiviert sein' }"; Type="PowerShell"; Category="Network-Events"},
    @{Name="Prozessstart-Events"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4688} -MaxEvents 30 -ErrorAction Stop | Select-Object TimeCreated, Message } catch { 'Prozessstart-Events: Event-ID 4688 nicht verfügbar oder keine Events vorhanden - Prozess-Audit-Richtlinie muss aktiviert sein' }"; Type="PowerShell"; Category="Network-Events"},
    @{Name="Windows Firewall Events"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Windows Firewall With Advanced Security/Firewall'} -MaxEvents 30 -ErrorAction Stop | Select-Object TimeCreated, Id, LevelDisplayName, Message } catch { 'Windows Firewall Events: Firewall-Event-Log nicht verfügbar oder keine Events vorhanden' }"; Type="PowerShell"; Category="Network-Events"},
    
    # === DOMAENEN-USER AUDIT (DE/EN) ===
    @{Name="AD-User (aktuelle Domaene)"; Command="if (Get-Module -ListAvailable -Name ActiveDirectory) { Import-Module ActiveDirectory; Get-ADUser -Filter * -Properties LastLogonDate, PasswordLastSet, Enabled | Select-Object Name, SamAccountName, Enabled, LastLogonDate, PasswordLastSet, DistinguishedName | Sort-Object Name } else { 'ActiveDirectory PowerShell-Modul nicht verfuegbar' }"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Domain-Users"},
    @{Name="Aktuell angemeldete Domaenen-User"; Command="if (Get-Module -ListAvailable -Name ActiveDirectory) { Import-Module ActiveDirectory; Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4624} -MaxEvents 100 | Where-Object {`$_.Message -match 'Logon Type:\\s*[23]' -and `$_.Message -notmatch 'ANONYMOUS|`$'} | ForEach-Object { if (`$_.Message -match 'Account Name:\\s*([^\\r\\n]+)' -and `$_.Message -match 'Account Domain:\\s*([^\\r\\n]+)') { [PSCustomObject]@{ TimeCreated = `$_.TimeCreated; AccountName = `$matches[1].Trim(); AccountDomain = `$matches[2].Trim(); LogonType = if (`$_.Message -match 'Logon Type:\\s*(\\d+)') { `$matches[1] } else { 'Unknown' } } } } | Where-Object {`$_.AccountName -ne '-' -and `$_.AccountName -ne 'ANONYMOUS LOGON'} | Sort-Object TimeCreated -Descending | Select-Object -First 20 } else { 'ActiveDirectory PowerShell-Modul nicht verfuegbar' }"; Type="PowerShell"; Category="Domain-Users"},
    @{Name="Privilegierte AD-Gruppen Mitglieder"; Command="if (Get-Module -ListAvailable -Name ActiveDirectory) { Import-Module ActiveDirectory; `$privGroups = @('Domain Admins', 'Enterprise Admins', 'Schema Admins', 'Administrators', 'Domänen-Admins', 'Organisations-Admins', 'Schema-Admins'); `$results = @(); foreach (`$group in `$privGroups) { try { `$members = Get-ADGroupMember -Identity `$group -ErrorAction SilentlyContinue | Get-ADUser -Properties LastLogonDate -ErrorAction SilentlyContinue; foreach (`$member in `$members) { `$results += [PSCustomObject]@{ GroupName = `$group; UserName = `$member.Name; SamAccountName = `$member.SamAccountName; Enabled = `$member.Enabled; LastLogonDate = `$member.LastLogonDate } } } catch { } }; `$results | Sort-Object GroupName, UserName } else { 'ActiveDirectory PowerShell-Modul nicht verfuegbar' }"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Domain-Users"},
    @{Name="Letzten Anmeldungen (Domaene)"; Command="if (Get-Module -ListAvailable -Name ActiveDirectory) { Import-Module ActiveDirectory; Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4624,4625} -MaxEvents 100 | Where-Object {`$_.Message -notmatch 'ANONYMOUS|`$' -and `$_.Message -match '@|\\\\'} | ForEach-Object { if (`$_.Message -match 'Account Name:\\s*([^\\r\\n]+)' -and `$_.Message -match 'Account Domain:\\s*([^\\r\\n]+)') { [PSCustomObject]@{ TimeCreated = `$_.TimeCreated; EventID = `$_.Id; AccountName = `$matches[1].Trim(); AccountDomain = `$matches[2].Trim(); Status = if (`$_.Id -eq 4624) { 'Erfolg' } else { 'Fehlgeschlagen' } } } } | Where-Object {`$_.AccountName -ne '-'} | Sort-Object TimeCreated -Descending | Select-Object -First 30 } else { 'ActiveDirectory PowerShell-Modul nicht verfuegbar' }"; Type="PowerShell"; Category="Domain-Users"},
    
    # === REMOTE VERBINDUNGEN ===
    @{Name="RDP-Sitzungen"; Command="qwinsta"; Type="CMD"; Category="Remote-Sessions"},
    @{Name="Remote Desktop Events"; Command="Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'} -MaxEvents 30 | Select-Object TimeCreated, Id, LevelDisplayName, Message"; Type="PowerShell"; Category="Remote-Sessions"},
    @{Name="SMB-Verbindungen"; Command="Get-SmbConnection | Select-Object ServerName, ShareName, UserName, Dialect"; Type="PowerShell"; Category="Remote-Sessions"},
    @{Name="SMB-Freigaben Zugriffe"; Command="Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-SmbServer/Security'} -MaxEvents 50 | Select-Object TimeCreated, Id, Message"; Type="PowerShell"; Category="Remote-Sessions"},
    
    # === ROUTING UND NETZWERK-TOPOLOGIE ===
    @{Name="Routing Tabelle"; Command="Get-NetRoute | Select-Object DestinationPrefix, NextHop, InterfaceAlias, RouteMetric, Protocol | Sort-Object DestinationPrefix"; Type="PowerShell"; Category="Network-Topology"},
    @{Name="Netzwerk-Statistiken"; Command="Get-NetAdapterStatistics | Select-Object Name, BytesReceived, BytesSent, PacketsReceived, PacketsSent"; Type="PowerShell"; Category="Network-Topology"},
    @{Name="Gateway-Informationen"; Command="Get-NetIPConfiguration | Where-Object {`$_.IPv4DefaultGateway} | Select-Object InterfaceAlias, IPv4Address, IPv4DefaultGateway, DNSServer"; Type="PowerShell"; Category="Network-Topology"}
)

# Variable fuer die Verbindungsaudit-Ergebnisse
$global:connectionAuditResults = @{}

# Funktion zum Ausfuehren von PowerShell-Befehlen
function Invoke-PSCommand {
    param(
        [string]$Command
    )
    try {
        Write-DebugLog "Ausfuehren von PowerShell-Befehl: $Command" "CommandExec"
        
        # Spezielle Behandlung fuer bestimmte Befehle
        if ($Command -like "*Get-ComputerInfo*") {
            $result = Get-ComputerInfo | Format-List | Out-String
        } else {
            $result = Invoke-Expression -Command $Command | Format-Table -AutoSize | Out-String
        }
        
        Write-DebugLog "PowerShell-Befehl erfolgreich ausgefuehrt. Ergebnis-Laenge: $($result.Length)" "CommandExec"
        return $result
    }
    catch {
        $errorMsg = "Fehler bei der Ausfuehrung des Befehls: $Command`r`n$($_.Exception.Message)"
        Write-DebugLog "FEHLER: $errorMsg" "CommandExec"
        return $errorMsg
    }
}

# Funktion zum Ausfuehren von CMD-Befehlen
function Invoke-CMDCommand {
    param(
        [string]$Command
    )
    try {
        Write-DebugLog "Ausfuehren von CMD-Befehl: $Command" "CommandExec"
        $result = cmd /c $Command 2>&1 | Out-String
        Write-DebugLog "CMD-Befehl erfolgreich ausgefuehrt. Ergebnis-Laenge: $($result.Length)" "CommandExec"
        return $result
    }
    catch {
        $errorMsg = "Fehler bei der Ausfuehrung des Befehls: $Command`r`n$($_.Exception.Message)"
        Write-DebugLog "FEHLER: $errorMsg" "CommandExec"
        return $errorMsg
    }
}

# Funktion zum Pruefen, ob eine bestimmte Serverrolle installiert ist
function Test-ServerRole {
    param(
        [string]$FeatureName
    )
    
    try {
        Write-DebugLog "Pruefe Serverrolle: $FeatureName" "RoleCheck"
        $feature = Get-WindowsFeature -Name $FeatureName -ErrorAction SilentlyContinue
        if ($feature -and $feature.Installed) {
            Write-DebugLog "Serverrolle $FeatureName ist installiert" "RoleCheck"
            return $true
        }
        Write-DebugLog "Serverrolle $FeatureName ist NICHT installiert" "RoleCheck"
        return $false
    }
    catch {
        Write-DebugLog "FEHLER beim Pruefen der Serverrolle $FeatureName - $($_.Exception.Message)" "RoleCheck"
        return $false
    }
}
# SICHERE UI-UPDATE FUNKTION
function Invoke-SafeDispatcher {
    param(
        [scriptblock]$Action,
        [string]$Priority = "Normal"
    )
    
    try {
        # Validiere alle UI-Objekte vor dem Zugriff
        if ($null -eq $window) {
            Write-DebugLog "Warnung: Window ist null - überspringe UI-Update" "UI"
            return
        }
        
        if ($null -eq $window.Dispatcher) {
            Write-DebugLog "Warnung: Window.Dispatcher ist null - überspringe UI-Update" "UI"
            return
        }
        
        # Prüfe Thread-Zugriff
        if ($window.Dispatcher.CheckAccess()) {
            # Wir sind im UI-Thread - direkter Aufruf
            if ($Action) {
                & $Action
            }
        } else {
            # Wir sind nicht im UI-Thread - verwende Dispatcher
            if ($Action) {
                $window.Dispatcher.Invoke($Action, $Priority)
            } else {
                $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
            }
        }
    } catch [System.InvalidOperationException] {
        Write-DebugLog "Warnung: UI-Thread ungültig - überspringe Update: $($_.Exception.Message)" "UI"
    } catch [System.ComponentModel.Win32Exception] {
        Write-DebugLog "Warnung: Windows-API Fehler - überspringe Update: $($_.Exception.Message)" "UI"
    } catch {
        Write-DebugLog "Warnung: UI-Update fehlgeschlagen: $($_.Exception.Message)" "UI"
    }
}

# VERBINDUNGSAUDIT FUNKTIONEN

# Hauptfunktion fuer die Verbindungsaudit-Durchfuehrung
function Start-ConnectionAuditProcess {
    # UI vorbereiten
    $btnRunConnectionAudit.IsEnabled = $false
    $btnExportConnectionHTML.IsEnabled = $false
    $btnExportConnectionDrawIO.IsEnabled = $false
    $btnCopyConnectionToClipboard.IsEnabled = $false
    $cmbConnectionCategories.IsEnabled = $false
    
    $rtbConnectionResults.Document = New-Object System.Windows.Documents.FlowDocument
    $progressBarConnection.Value = 0
    Update-StatusText "Status: Verbindungsaudit laeuft..."
    
    # UI initial aktualisieren
    Invoke-SafeDispatcher -Action {
        $txtProgressConnection.Text = "Initialisiere Verbindungsaudit..."
        $progressBarConnection.Value = 0
    }
    
    # UI refresh erzwingen
    Invoke-SafeDispatcher
    Start-Sleep -Milliseconds 300
    
    # Sammle ausgewaehlte Befehle
    $selectedCommands = @()
    foreach ($cmd in $connectionAuditCommands) {
        if ($connectionCheckboxes[$cmd.Name].IsChecked) {
            $selectedCommands += $cmd
        }
    }
    
    Write-DebugLog "Starte Verbindungsaudit mit $($selectedCommands.Count) ausgewaehlten Befehlen" "ConnectionAudit"
    
    $global:connectionAuditResults = @{}
    $progressStep = 100.0 / $selectedCommands.Count
    $currentProgress = 0
    
    # UI Update mit Anzahl der Befehle
    Invoke-SafeDispatcher -Action {
        $txtProgressConnection.Text = "Bereite $($selectedCommands.Count) Verbindungsaudit-Befehle vor..."
    }
    
    # UI refresh erzwingen
    Invoke-SafeDispatcher
    Start-Sleep -Milliseconds 500
    
    for ($i = 0; $i -lt $selectedCommands.Count; $i++) {
        $cmd = $selectedCommands[$i]
        
        # UI aktualisieren - BEGINN des Befehls
        $window.Dispatcher.Invoke([Action]{
            $txtProgressConnection.Text = "Verarbeite: $($cmd.Name) ($($i+1)/$($selectedCommands.Count))"
            $progressBarConnection.Value = $currentProgress
        }, "Normal")
        
        # UI refresh erzwingen
        $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
        Invoke-SafeDoEvents
        Start-Sleep -Milliseconds 200
        
        Write-DebugLog "Fuehre Verbindungsaudit aus ($($i+1)/$($selectedCommands.Count)): $($cmd.Name)" "ConnectionAudit"
        
        try {
            if ($cmd.Type -eq "PowerShell") {
                $result = Invoke-PSCommand -Command $cmd.Command
            } else {
                $result = Invoke-CMDCommand -Command $cmd.Command
            }
            
            $global:connectionAuditResults[$cmd.Name] = $result
            
            # Erfolg und Fortschrittsbalken aktualisieren
            $currentProgress += $progressStep
            $window.Dispatcher.Invoke([Action]{
                $progressBarConnection.Value = $currentProgress
                $txtProgressConnection.Text = "Abgeschlossen: $($cmd.Name) ($($i+1)/$($selectedCommands.Count))"
            }, "Normal")
            
        } catch {
            $errorMsg = "Fehler: $($_.Exception.Message)"
            $global:connectionAuditResults[$cmd.Name] = $errorMsg
            
            # Fehler und Fortschrittsbalken trotzdem aktualisieren
            $currentProgress += $progressStep
            $window.Dispatcher.Invoke([Action]{
                $progressBarConnection.Value = $currentProgress
                $txtProgressConnection.Text = "Fehler bei: $($cmd.Name) ($($i+1)/$($selectedCommands.Count))"
            }, "Normal")
            
            Write-DebugLog "FEHLER bei Verbindungsaudit $($cmd.Name): $($_.Exception.Message)" "ConnectionAudit"
        }
        
        # UI refresh nach jedem Befehl erzwingen
        $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
        Invoke-SafeDoEvents
        Start-Sleep -Milliseconds 300
        
        # Zwischenstand der Ergebnisse aktualisieren
        if (($i + 1) % 3 -eq 0 -or $i -eq ($selectedCommands.Count - 1)) {
            $window.Dispatcher.Invoke([Action]{
                try {
                    Update-ConnectionResultsCategories
                    if ($cmbConnectionCategories.SelectedItem) {
                        $selectedCategory = $cmbConnectionCategories.SelectedItem.Tag
                        Show-ConnectionCategoryResults -Category $selectedCategory
                    } else {
                        Show-ConnectionCategoryResults -Category "Alle"
                    }
                }
                catch {
                    Write-DebugLog "FEHLER beim Zwischenupdate der Verbindungsaudit-Ergebnisanzeige: $($_.Exception.Message)" "ConnectionAudit"
                }
            }, "Normal")
            $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
            Invoke-SafeDoEvents
        }
    }
    
    # Verbindungsaudit abgeschlossen - Finale Updates
    $window.Dispatcher.Invoke([Action]{
        $progressBarConnection.Value = 100
        $txtProgressConnection.Text = "Verbindungsaudit vollstaendig abgeschlossen! $($selectedCommands.Count) Befehle ausgefuehrt."
        
        try {
            # Aktualisiere die Kategorien-Anzeige
            Update-ConnectionResultsCategories
            Show-ConnectionCategoryResults -Category "Alle"
        }
        catch {
            Write-DebugLog "FEHLER beim finalen Update der Verbindungsaudit-Ergebnisanzeige: $($_.Exception.Message)" "ConnectionAudit"
            try {
                Show-SimpleConnectionResults -Category "Alle"
            }
            catch {
                Write-DebugLog "FEHLER auch bei einfacher Verbindungsaudit-Anzeige: $($_.Exception.Message)" "ConnectionAudit"
            }
        }
        
        Update-StatusText "Status: Verbindungsaudit abgeschlossen - $($global:connectionAuditResults.Count) Ergebnisse"
        
        # Automatisch zu den Netzwerk-Ergebnissen wechseln
        Switch-Panel "networkResults"
        
        # Buttons wieder aktivieren
        $btnRunConnectionAudit.IsEnabled = $true
        $btnExportConnectionHTML.IsEnabled = $true
        $btnExportConnectionDrawIO.IsEnabled = $true
        $btnCopyConnectionToClipboard.IsEnabled = $true
        $cmbConnectionCategories.IsEnabled = $true
    }, "Normal")
    
    # Finaler UI refresh
    $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
    Invoke-SafeDoEvents
    
    Write-DebugLog "Verbindungsaudit abgeschlossen mit $($global:connectionAuditResults.Count) Ergebnissen" "ConnectionAudit"
}

# Funktion zum Aktualisieren der Verbindungsaudit-Kategorien-ComboBox
function Update-ConnectionResultsCategories {
    Write-DebugLog "Aktualisiere Verbindungsaudit-Kategorien-ComboBox" "UI"
    
    $cmbConnectionCategories.Items.Clear()
    
    # "Alle" Option hinzufügen
    $allItem = New-Object System.Windows.Controls.ComboBoxItem
    $allItem.Content = "Alle Kategorien"
    $allItem.Tag = "Alle"
    $cmbConnectionCategories.Items.Add($allItem)
    
    # Einzelne Kategorien hinzufügen
    $categories = @{}
    if ($null -ne $connectionAuditCommands) {
        foreach ($cmd in $connectionAuditCommands) {
            $category = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
            if (-not $categories.ContainsKey($category)) {
                $categories[$category] = 0
            }
            if ($global:connectionAuditResults.ContainsKey($cmd.Name)) {
                $categories[$category]++
            }
        }
    }
    
    foreach ($category in $categories.Keys | Sort-Object) {
        if ($categories[$category] -gt 0) {
            $categoryItem = New-Object System.Windows.Controls.ComboBoxItem
            $categoryItem.Content = "$category ($($categories[$category]))"
            $categoryItem.Tag = $category
            $cmbConnectionCategories.Items.Add($categoryItem)
        }
    }
    
    # Ersten Eintrag auswählen
    if ($cmbConnectionCategories.Items.Count -gt 0) {
        $cmbConnectionCategories.SelectedIndex = 0
    }
}

# Funktion zum Anzeigen der Verbindungsaudit-Ergebnisse
function Show-ConnectionCategoryResults {
    param([string]$Category = "Alle")
    
    Write-DebugLog "Zeige Verbindungsaudit-Ergebnisse fuer Kategorie: $Category" "UI"
    
    if ($global:connectionAuditResults.Count -eq 0) {
        $rtbConnectionResults.Document = New-Object System.Windows.Documents.FlowDocument
        $emptyParagraph = New-Object System.Windows.Documents.Paragraph
        $emptyRun = New-Object System.Windows.Documents.Run("Keine Verbindungsaudit-Ergebnisse verfügbar. Führen Sie zuerst ein Verbindungsaudit durch.")
        $emptyRun.FontStyle = "Italic"
        $emptyRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(108, 117, 125))
        $emptyParagraph.Inlines.Add($emptyRun)
        $rtbConnectionResults.Document.Blocks.Add($emptyParagraph)
        return
    }
    
    try {
        # Versuche die formatierte Anzeige
        $document = Format-ConnectionRichTextResults -Results $global:connectionAuditResults -CategoryFilter $Category
        
        # Sichere FlowDocument-Zuweisung
        $window.Dispatcher.Invoke([Action]{
            try {
                $rtbConnectionResults.Document = $document
                # Erzwinge Layout-Update mit Timeout
                $rtbConnectionResults.UpdateLayout()
            }
            catch {
                Write-DebugLog "FEHLER bei FlowDocument-Zuweisung: $($_.Exception.Message)" "UI"
                # Fallback zu einfacher Textanzeige
                Show-SimpleConnectionResults -Category $Category
            }
        }, "Normal", [TimeSpan]::FromSeconds(5))
        
        Write-DebugLog "Verbindungsaudit-Ergebnisse erfolgreich formatiert und angezeigt" "UI"
    }
    catch {
        Write-DebugLog "FEHLER bei der formatierten Verbindungsaudit-Anzeige: $($_.Exception.Message) - Verwende Fallback" "UI"
        
        # Fallback: Verwende einfache Textanzeige
        Show-SimpleConnectionResults -Category $Category
    }
}

# Funktion zum Formatieren der Verbindungsaudit-RichTextBox
function Format-ConnectionRichTextResults {
    param(
        [hashtable]$Results,
        [string]$CategoryFilter = "Alle"
    )
    
    Write-DebugLog "Formatiere Verbindungsaudit-Ergebnisse fuer Kategorie: $CategoryFilter" "UI"
    
    try {
        # Erstelle ein neues FlowDocument mit sichereren Einstellungen
        $document = New-Object System.Windows.Documents.FlowDocument
        $document.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
        $document.FontSize = 12
        $document.LineHeight = 18
        
        # Sichere Layout-Einstellungen ohne potentielle Memory-Probleme
        $document.PageWidth = [Double]::NaN
        $document.PageHeight = [Double]::NaN
        $document.ColumnWidth = [Double]::PositiveInfinity
        $document.TextAlignment = "Left"
        $document.PagePadding = New-Object System.Windows.Thickness(5)
        $document.IsOptimalParagraphEnabled = $false  # Deaktiviert für Stabilität
        $document.IsHyphenationEnabled = $false
        
        # Sicherheitsprüfung: Begrenzte Anzahl von Ergebnissen verarbeiten
        $maxItemsPerCategory = 50
        $processedItems = 0
        
        # Gruppiere Ergebnisse nach Kategorien
        $categorizedResults = @{}
        if ($null -ne $connectionAuditCommands) {
            foreach ($cmd in $connectionAuditCommands) {
                if ($processedItems -ge ($maxItemsPerCategory * 10)) { break } # Maximal 500 Items insgesamt
                
                $category = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
                if (-not $categorizedResults.ContainsKey($category)) {
                    $categorizedResults[$category] = @()
                }
                if ($Results.ContainsKey($cmd.Name)) {
                    $categorizedResults[$category] += @{
                        Name = $cmd.Name
                        Result = $Results[$cmd.Name]
                        Command = $cmd
                    }
                    $processedItems++
                }
            }
        }
        
        # Bestimme welche Kategorien angezeigt werden sollen
        $categoriesToShow = if ($CategoryFilter -eq "Alle") { 
            $categorizedResults.Keys | Sort-Object | Select-Object -First 20  # Maximal 20 Kategorien
        } else { 
            @($CategoryFilter) 
        }
        
        $totalItems = 0
        foreach ($category in $categoriesToShow) {
            if ($categorizedResults.ContainsKey($category)) {
                $categoryData = $categorizedResults[$category] | Select-Object -First $maxItemsPerCategory
                $totalItems += $categoryData.Count
                
                # Kategorie-Header mit sicherer Erstellung
                try {
                    $categoryParagraph = New-Object System.Windows.Documents.Paragraph
                    $categoryRun = New-Object System.Windows.Documents.Run("Kategorie: $category")
                    $categoryRun.FontWeight = "Bold"
                    $categoryRun.FontSize = 16
                    $categoryRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(253, 126, 20))
                    $categoryParagraph.Inlines.Add($categoryRun)
                    $categoryParagraph.Margin = New-Object System.Windows.Thickness(0, 15, 0, 10)
                    $categoryParagraph.TextAlignment = "Left"
                    $categoryParagraph.KeepTogether = $true
                    $document.Blocks.Add($categoryParagraph)
                }
                catch {
                    Write-DebugLog "FEHLER beim Erstellen der Kategorie-Header für ${category}: $($_.Exception.Message)" "UI"
                    continue
                }
                
                # Items in dieser Kategorie
                foreach ($item in $categoryData) {
                    try {
                        # Item-Header mit Textlängen-Begrenzung
                        $itemParagraph = New-Object System.Windows.Documents.Paragraph
                        $itemName = if ($item.Name.Length -gt 100) { $item.Name.Substring(0, 100) + "..." } else { $item.Name }
                        $itemRun = New-Object System.Windows.Documents.Run("Eintrag: $itemName")
                        $itemRun.FontWeight = "SemiBold"
                        $itemRun.FontSize = 14
                        $itemRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(44, 62, 80))
                        $itemParagraph.Inlines.Add($itemRun)
                        $itemParagraph.Margin = New-Object System.Windows.Thickness(0, 10, 0, 5)
                        $itemParagraph.TextAlignment = "Left"
                        $itemParagraph.KeepTogether = $true
                        $document.Blocks.Add($itemParagraph)
                        
                        # Command-Info (optional)
                        if ($item.Command -and $item.Command.Command) {
                            $cmdParagraph = New-Object System.Windows.Documents.Paragraph
                            $cmdText = if ($item.Command.Command.Length -gt 200) { $item.Command.Command.Substring(0, 200) + "..." } else { $item.Command.Command }
                            $cmdRun = New-Object System.Windows.Documents.Run("Befehl: $cmdText")
                            $cmdRun.FontSize = 11
                            $cmdRun.FontStyle = "Italic"
                            $cmdRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(108, 117, 125))
                            $cmdParagraph.Inlines.Add($cmdRun)
                            $cmdParagraph.Margin = New-Object System.Windows.Thickness(10, 0, 0, 5)
                            $cmdParagraph.TextAlignment = "Left"
                            $document.Blocks.Add($cmdParagraph)
                        }
                        
                        # Result mit Textlängen-Begrenzung für Stabilität
                        $resultParagraph = New-Object System.Windows.Documents.Paragraph
                        $resultText = if ($item.Result -and $item.Result.Length -gt 5000) { 
                            $item.Result.Substring(0, 5000) + "`r`n... (Ausgabe gekürzt für Stabilität)" 
                        } else { 
                            $item.Result 
                        }
                        $resultRun = New-Object System.Windows.Documents.Run($resultText)
                        $resultRun.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
                        $resultRun.FontSize = 11
                        $resultRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(73, 80, 87))
                        $resultParagraph.Inlines.Add($resultRun)
                        $resultParagraph.Margin = New-Object System.Windows.Thickness(10, 0, 0, 15)
                        $resultParagraph.TextAlignment = "Left"
                        $document.Blocks.Add($resultParagraph)
                    }
                    catch {
                        Write-DebugLog "FEHLER beim Erstellen des Items $($item.Name): $($_.Exception.Message)" "UI"
                        # Füge einen Fehler-Paragraph hinzu
                        try {
                            $errorParagraph = New-Object System.Windows.Documents.Paragraph
                            $errorRun = New-Object System.Windows.Documents.Run("FEHLER beim Anzeigen von: $($item.Name)")
                            $errorRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(220, 53, 69))
                            $errorParagraph.Inlines.Add($errorRun)
                            $document.Blocks.Add($errorParagraph)
                        }
                        catch {
                            Write-DebugLog "Kritischer FEHLER beim Erstellen des Fehler-Paragraphs" "UI"
                        }
                    }
                }
            }
        }
        
        Write-DebugLog "FlowDocument erfolgreich erstellt mit $totalItems Items" "UI"
        return $document
    }
    catch {
        Write-DebugLog "KRITISCHER FEHLER beim Erstellen des FlowDocuments: $($_.Exception.Message)" "UI"
        
        # Fallback: Einfaches Dokument mit Fehlermeldung
        try {
            $fallbackDocument = New-Object System.Windows.Documents.FlowDocument
            $fallbackParagraph = New-Object System.Windows.Documents.Paragraph
            $fallbackRun = New-Object System.Windows.Documents.Run("Fehler beim Erstellen der formatierten Anzeige. Verwenden Sie den Export oder die einfache Ansicht.")
            $fallbackRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(220, 53, 69))
            $fallbackParagraph.Inlines.Add($fallbackRun)
            $fallbackDocument.Blocks.Add($fallbackParagraph)
            return $fallbackDocument
        }
        catch {
            Write-DebugLog "Selbst Fallback-Dokument konnte nicht erstellt werden" "UI"
            return $null
        }
    }
}

# Sichere Fallback-Funktion für Verbindungsaudit-Ergebnisse
function Show-SimpleConnectionResults {
    param([string]$Category = "Alle")
    
    Write-DebugLog "Verwende einfache Textanzeige für Verbindungsaudit-Kategorie: $Category" "UI"
    
    # Erstelle einfaches FlowDocument
    $document = New-Object System.Windows.Documents.FlowDocument
    $document.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $document.FontSize = 11
    $document.PageWidth = [Double]::NaN
    $document.PageHeight = [Double]::NaN
    $document.ColumnWidth = [Double]::PositiveInfinity
    
    # Sammle alle relevanten Ergebnisse als einfachen Text
    $resultText = ""
    
    # Gruppiere nach Kategorien
    $categorizedResults = @{}
    foreach ($cmd in $connectionAuditCommands) {
        $cmdCategory = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
        if (-not $categorizedResults.ContainsKey($cmdCategory)) {
            $categorizedResults[$cmdCategory] = @()
        }
        if ($global:connectionAuditResults.ContainsKey($cmd.Name)) {
            $categorizedResults[$cmdCategory] += @{
                Name = $cmd.Name
                Result = $global:connectionAuditResults[$cmd.Name]
            }
        }
    }
    
    # Bestimme anzuzeigende Kategorien
    $categoriesToShow = if ($Category -eq "Alle") { 
        $categorizedResults.Keys | Sort-Object 
    } else { 
        @($Category) 
    }
    
    $totalItems = 0
    foreach ($cat in $categoriesToShow) {
        if ($categorizedResults.ContainsKey($cat)) {
            $categoryData = $categorizedResults[$cat]
            $totalItems += $categoryData.Count
            
            $resultText += "`n" + "="*60 + "`n"
            $resultText += "VERBINDUNGSAUDIT KATEGORIE: $cat`n"
            $resultText += "="*60 + "`n`n"
            
            foreach ($item in $categoryData) {
                $resultText += "-"*40 + "`n"
                $resultText += "EINTRAG: $($item.Name)`n"
                $resultText += "-"*40 + "`n"
                $resultText += "$($item.Result)`n`n"
            }
        }
    }
    
    # Erstelle einfachen Paragraph mit dem gesamten Text
    $paragraph = New-Object System.Windows.Documents.Paragraph
    $run = New-Object System.Windows.Documents.Run($resultText)
    $paragraph.Inlines.Add($run)
    $document.Blocks.Add($paragraph)
    
    $rtbConnectionResults.Document = $document
}

# Funktion zum Anzeigen der Verbindungsaudit-Ergebnisse mit Kategorisierung und WPF-Formatierung
function Show-ConnectionResults {
    param(
        [string]$Category = "Alle"
    )

    Write-DebugLog "Zeige Verbindungsergebnisse für Kategorie: $Category" "UI"

    # RichTextBox vorbereiten
    $rtbConnectionResults.Document.Blocks.Clear()
    $rtbConnectionResults.IsEnabled = $true # Sicherstellen, dass die RTB aktiviert ist

    # Prüfen, ob überhaupt Ergebnisse vorhanden sind
    if ($null -eq $global:connectionAuditResults -or $global:connectionAuditResults.Count -eq 0) {
        $paragraph = New-Object System.Windows.Documents.Paragraph
        $run = New-Object System.Windows.Documents.Run("Keine Verbindungsaudit-Ergebnisse vorhanden.")
        $paragraph.Inlines.Add($run)
        $rtbConnectionResults.Document.Blocks.Add($paragraph)
        Write-DebugLog "Keine Verbindungsaudit-Ergebnisse zum Anzeigen." "UI"
        return
    }

    # Prüfen, ob Befehlsdefinitionen für die Kategorisierung vorhanden sind
    if ($null -eq $connectionAuditCommands -or $connectionAuditCommands.Count -eq 0) {
        Write-DebugLog "Show-ConnectionResults: Keine Verbindungsaudit-Befehlsdefinitionen vorhanden. Zeige unkategorisierte Rohdaten." "UI-Warning"
        
        # Fallback: Unkategorisierte Rohdaten anzeigen, wenn keine Befehlsdefinitionen geladen sind
        $document = New-Object System.Windows.Documents.FlowDocument
        $document.PagePadding = New-Object System.Windows.Thickness(5)

        $headerParagraph = New-Object System.Windows.Documents.Paragraph
        $headerRun = New-Object System.Windows.Documents.Run("VERBINDUNGSAUDIT ERGEBNISSE (Unkategorisiert)")
        $headerRun.FontWeight = [System.Windows.FontWeights]::Bold
        $headerRun.FontSize = 14
        $headerParagraph.Inlines.Add($headerRun)
        $headerParagraph.Margin = New-Object System.Windows.Thickness(0,0,0,10) # Abstand nach unten
        $document.Blocks.Add($headerParagraph)

        foreach ($key in ($global:connectionAuditResults.Keys | Sort-Object)) {
            $itemNameParagraph = New-Object System.Windows.Documents.Paragraph
            $itemNameRun = New-Object System.Windows.Documents.Run("BEFEHL: $key") # Geändert zu "BEFEHL" für Klarheit
            $itemNameRun.FontWeight = [System.Windows.FontWeights]::SemiBold
            $itemNameParagraph.Inlines.Add($itemNameRun)
            $itemNameParagraph.Margin = New-Object System.Windows.Thickness(0,5,0,2)
            $document.Blocks.Add($itemNameParagraph)

            $itemResultParagraph = New-Object System.Windows.Documents.Paragraph
            $resultDisplayString = if ($null -ne $global:connectionAuditResults[$key]) {
                                       if ($global:connectionAuditResults[$key] -is [string]) { 
                                           $global:connectionAuditResults[$key]
                                       } else { 
                                           ($global:connectionAuditResults[$key] | Out-String).TrimEnd() 
                                       }
                                   } else { 
                                       "[Kein Ergebnis oder NULL]" 
                                   }
            $itemResultRun = New-Object System.Windows.Documents.Run($resultDisplayString)
            $itemResultRun.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas, Courier New, Lucida Console")
            $itemResultRun.FontSize = 11
            $itemResultParagraph.Inlines.Add($itemResultRun)
            $itemResultParagraph.Margin = New-Object System.Windows.Thickness(10,0,0,10) # Einzug links, Abstand unten
            $document.Blocks.Add($itemResultParagraph)
        }
        $rtbConnectionResults.Document = $document
        return
    }

    # Ergebnisse nach Kategorien gruppieren
    $categorizedResults = @{}
    foreach ($cmdDef in $connectionAuditCommands) {
        $cmdCategory = if ([string]::IsNullOrWhiteSpace($cmdDef.Category)) { "Allgemein" } else { $cmdDef.Category }
        
        if (-not $categorizedResults.ContainsKey($cmdCategory)) {
            $categorizedResults[$cmdCategory] = [System.Collections.Generic.List[object]]::new()
        }
        
        if ($global:connectionAuditResults.ContainsKey($cmdDef.Name)) {
            $categorizedResults[$cmdCategory].Add(@{
                Name = $cmdDef.Name
                Result = $global:connectionAuditResults[$cmdDef.Name]
            })
        }
    }

    # Zu anzeigende Kategorien bestimmen
    $categoriesToShow = if ($Category -eq "Alle") { 
        $categorizedResults.Keys | Where-Object { $categorizedResults[$_].Count -gt 0 } | Sort-Object 
    } else { 
        if ($categorizedResults.ContainsKey($Category) -and $categorizedResults[$Category].Count -gt 0) {
            @($Category) 
        } else {
            @() 
        }
    }
    
    $document = New-Object System.Windows.Documents.FlowDocument
    $document.PagePadding = New-Object System.Windows.Thickness(5)

    if ($categoriesToShow.Count -eq 0) {
        $message = if ($Category -eq "Alle") {
            "Keine kategorisierten Ergebnisse gefunden. Möglicherweise sind alle Ergebnisse ohne Kategorie oder die Befehlsdefinitionen passen nicht."
        } else {
            "Keine Ergebnisse für Kategorie '$Category' gefunden oder die Kategorie ist leer."
        }
        $paragraph = New-Object System.Windows.Documents.Paragraph
        $run = New-Object System.Windows.Documents.Run($message)
        $paragraph.Inlines.Add($run)
        $document.Blocks.Add($paragraph)
        Write-DebugLog $message "UI"
    } else {
        foreach ($catName in $categoriesToShow) {
            # Erneute Prüfung, obwohl $categoriesToShow bereits gefiltert sein sollte
            if ($categorizedResults.ContainsKey($catName) -and $categorizedResults[$catName].Count -gt 0) {
                $categoryItems = $categorizedResults[$catName]
                
                # Kategorie-Überschrift
                $headerParagraph = New-Object System.Windows.Documents.Paragraph
                $headerRun = New-Object System.Windows.Documents.Run("KATEGORIE: $($catName.ToUpper())")
                $headerRun.FontWeight = [System.Windows.FontWeights]::Bold
                $headerRun.FontSize = 14 
                $headerParagraph.Inlines.Add($headerRun)
                $headerParagraph.Margin = New-Object System.Windows.Thickness(0,10,0,5) # Oben, Rechts, Unten, Links
                $document.Blocks.Add($headerParagraph)

                foreach ($item in $categoryItems) {
                    # Eintragsname (Befehlsname)
                    $itemNameParagraph = New-Object System.Windows.Documents.Paragraph
                    $itemNameRun = New-Object System.Windows.Documents.Run($item.Name)
                    $itemNameRun.FontWeight = [System.Windows.FontWeights]::SemiBold
                    $itemNameRun.FontSize = 12
                    $itemNameParagraph.Inlines.Add($itemNameRun)
                    $itemNameParagraph.Margin = New-Object System.Windows.Thickness(0,5,0,2)
                    $document.Blocks.Add($itemNameParagraph)

                    # Eintragsergebnis
                    $itemResultParagraph = New-Object System.Windows.Documents.Paragraph
                    $resultDisplayString = ""
                    if ($null -ne $item.Result) {
                        if ($item.Result -is [string]) {
                            $resultDisplayString = $item.Result
                        } else {
                            $resultDisplayString = ($item.Result | Out-String).TrimEnd()
                        }
                    } else {
                        $resultDisplayString = "[Kein Ergebnis oder NULL]"
                    }

                    $itemResultRun = New-Object System.Windows.Documents.Run($resultDisplayString)
                    $itemResultRun.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas, Courier New, Lucida Console") 
                    $itemResultRun.FontSize = 11
                    $itemResultParagraph.Inlines.Add($itemResultRun)
                    $itemResultParagraph.Margin = New-Object System.Windows.Thickness(10,0,0,10) # Einzug links, Abstand unten
                    $document.Blocks.Add($itemResultParagraph)
                }
            }
        }
    }
    
    $rtbConnectionResults.Document = $document
    Write-DebugLog "Anzeige der Verbindungsergebnisse für Kategorie '$Category' abgeschlossen." "UI"
}

# Funktion zum sicheren Ausführen von Get-WinEvent Befehlen
function Invoke-SafeWinEvent {
    param(
        [hashtable]$FilterHashtable,
        [int]$MaxEvents = 50,
        [string]$Description = "Events"
    )
    
    try {
        Write-DebugLog "Versuche Get-WinEvent mit Filter: $($FilterHashtable | ConvertTo-Json -Compress)" "SafeWinEvent"
        
        # Versuche Event-Abfrage mit ErrorAction Stop. 2>$null unterdrückt die Standard-Fehlerausgabe in der Konsole.
        $events = Get-WinEvent -FilterHashtable $FilterHashtable -MaxEvents $MaxEvents -ErrorAction Stop 2>$null
        
        # Wenn $events nicht $null ist und Elemente enthält (und keine Exception ausgelöst wurde),
        # werden die Events zurückgegeben.
        if ($null -ne $events -and $events.Count -gt 0) {
            Write-DebugLog "Erfolgreich $($events.Count) Events gefunden" "SafeWinEvent"
            return $events
        } else {
            # Dieser Block wird erreicht, wenn Get-WinEvent $null oder ein leeres Array zurückgibt, 
            # ohne einen Fehler auszulösen, der von ErrorAction Stop abgefangen würde.
            Write-DebugLog "Get-WinEvent lieferte keine Events oder ein leeres Ergebnis (ohne Exception). Filter: $($FilterHashtable | ConvertTo-Json -Compress). Gebe leeres Array zurück." "SafeWinEvent"
            return @() 
        }
    }
    catch { # Fängt alle terminierenden Fehler von Get-WinEvent ab
        $ErrorRecord = $PSItem # $PSItem ist der ErrorRecord im Catch-Block (in PSv3+).
        
        Write-DebugLog "Get-WinEvent Fehler aufgetreten. Message: '$($ErrorRecord.Exception.Message)'. FullyQualifiedErrorId: '$($ErrorRecord.FullyQualifiedErrorId)'." "SafeWinEvent"
        
        # Spezifische Behandlung für häufige Fehler
        # Prüfung auf 'NoMatchingEventsFound' anhand der FullyQualifiedErrorId (bevorzugt)
        if ($ErrorRecord.FullyQualifiedErrorId -eq "NoMatchingEventsFound,Microsoft.PowerShell.Commands.GetWinEventCommand") {
            Write-DebugLog "Fehler 'NoMatchingEventsFound' (basierend auf FQID) abgefangen. Gebe leeres Array zurück." "SafeWinEvent"
            return @()
        }
        # Fallback: Prüfung auf 'NoMatchingEventsFound' oder deutschsprachige Entsprechung anhand der Fehlermeldung
        elseif ($ErrorRecord.Exception.Message -like "*NoMatchingEventsFound*" -or $ErrorRecord.Exception.Message -like "*Es wurden keine Ereignisse gefunden*") {
            Write-DebugLog "Fehler '$($ErrorRecord.Exception.Message)' (Nachricht ähnlich 'Keine Events gefunden') abgefangen. Gebe leeres Array zurück." "SafeWinEvent"
            return @()
        }
        # Zugriff verweigert (Access Denied)
        elseif ($ErrorRecord.Exception.Message -like "*Access is denied*" -or $ErrorRecord.Exception.Message -like "*Zugriff verweigert*") {
            Write-DebugLog "Fehler '$($ErrorRecord.Exception.Message)' (Zugriff verweigert) abgefangen." "SafeWinEvent"
            return "$Description nicht verfügbar - Keine Berechtigung für Event-Log-Zugriff"
        }
        # Kanal/Log nicht gefunden (Channel/Log not found)
        elseif (
            $ErrorRecord.Exception.Message -like "*The specified channel could not be found*" -or 
            $ErrorRecord.Exception.Message -like "*Der angegebene Kanal wurde nicht gefunden*" -or
            $ErrorRecord.Exception.Message -like "*log does not exist*" -or # Allgemeinere Prüfung auf Nichtexistenz des Logs
            $ErrorRecord.Exception.Message -like "*existiert nicht*" # Deutsche Variante für "existiert nicht"
        ) {
            Write-DebugLog "Fehler '$($ErrorRecord.Exception.Message)' (Kanal/Log existiert nicht) abgefangen." "SafeWinEvent"
            return "$Description nicht verfügbar - Event-Log-Kanal existiert nicht"
        }
        # Andere Fehler
        else {
            # Versuche, eine kurze, prägnante Fehlermeldung zu extrahieren (erster Satz oder erste Zeile)
            $shortErrorMessage = "Unbekannter Fehler" # Standardwert
            if ($ErrorRecord.Exception.Message) {
                $splitMessages = $ErrorRecord.Exception.Message -split '\r?\n|\. '
                if ($splitMessages.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($splitMessages[0])) {
                    $shortErrorMessage = $splitMessages[0]
                }
            }
            Write-DebugLog "Allgemeiner Get-WinEvent Fehler: '$shortErrorMessage'. Vollständige Exception: $($ErrorRecord.Exception)" "SafeWinEvent"
            return "$Description nicht verfügbar - Event-Log-Fehler: $shortErrorMessage"
        }
    }
}

# Funktion zum HTML-Export der Verbindungsaudit-Ergebnisse
function Export-ConnectionAuditToHTML {
    param(
        [hashtable]$Results,
        [string]$FilePath
    )
    
    Write-DebugLog "Starte Verbindungsaudit HTML-Export nach: $FilePath" "Export"

    # Helper to replace Umlaute and escape HTML
    function Convert-ToDisplayString {
        param([string]$Text)
        if ([string]::IsNullOrEmpty($Text)) { return "" }
        $processedText = $Text -replace 'ä', 'ae' -replace 'ö', 'oe' -replace 'ü', 'ue' -replace 'Ä', 'Ae' -replace 'Ö', 'Oe' -replace 'Ü', 'Ue' -replace 'ß', 'ss'
        return [System.Security.SecurityElement]::Escape($processedText)
    }
    
    # Erweiterte Serverinformationen sammeln
    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cpuInfoObj = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $totalRamBytes = (Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue | Measure-Object -Property Capacity -Sum).Sum
    
    $serverInfo = @{
        ServerName = $env:COMPUTERNAME
        ReportDate = Get-Date -Format "dd.MM.yyyy | HH:mm:ss"
        Domain = $env:USERDOMAIN
        User = $env:USERNAME
        OS = if ($osInfo) { "$($osInfo.Caption) $($osInfo.OSArchitecture)" } else { "N/A" }
        CPU = if ($cpuInfoObj) { $cpuInfoObj.Name } else { "N/A" }
        RAM = if ($totalRamBytes) { "{0:N2} GB" -f ($totalRamBytes / 1GB) } else { "N/A" }
    }

    # Gruppiere Ergebnisse nach Kategorien
    $groupedResults = @{}
    
    if ($null -ne $connectionAuditCommands) {
        foreach ($cmdDef in $connectionAuditCommands) {
            $categoryName = if ($cmdDef.Category) { $cmdDef.Category } else { "Allgemein" }
            
            if ($Results.ContainsKey($cmdDef.Name)) {
                if (-not $groupedResults.ContainsKey($categoryName)) {
                    $groupedResults[$categoryName] = @()
                }
                
                $groupedResults[$categoryName] += @{
                    Name = $cmdDef.Name
                    Result = $Results[$cmdDef.Name]
                    Command = $cmdDef
                }
            }
        }
    }

    # Navigationselemente und Tab-Inhalte generieren
    $sidebarNavLinks = ""
    $mainContentTabs = ""
    $firstTabId = $null
    
    $sortedCategories = $groupedResults.Keys | Sort-Object
    
    foreach ($categoryKey in $sortedCategories) {
        $items = $groupedResults[$categoryKey]
        $displayCategory = Convert-ToDisplayString $categoryKey
        
        $categoryIdPart = $categoryKey -replace '[^a-zA-Z0-9_]', ''
        if ($categoryIdPart.Length -eq 0) { 
            $categoryIdPart = "cat" + ($categoryKey.GetHashCode() | ForEach-Object ToString X) 
        }
        $tabId = "tab_$categoryIdPart"

        if ($null -eq $firstTabId) { $firstTabId = $tabId }

        $sidebarNavLinks += @"
<li class="nav-item category-nav">
    <a href="#" class="nav-link" onclick="showTab('$tabId', this)">
        $displayCategory ($($items.Count))
    </a>
</li>
"@
        
        $tabContent = "<div id='$tabId' class='tab-content'>"
        $tabContent += "<h2 class='content-category-title'>$displayCategory</h2>"

        foreach ($item in $items) {
            $displayItemName = Convert-ToDisplayString $item.Name
            $displayResult = Convert-ToDisplayString $item.Result
            
            $tabContent += @"
<div class="section">
    <div class="section-header">
        <h3 class="section-title">$displayItemName</h3>
    </div>
    <div class="section-content">
        <pre>$displayResult</pre>
    </div>
</div>
"@
        }
        
        $tabContent += "</div>"
        $mainContentTabs += $tabContent
    }

    # Erstelle die vollständige HTML-Ausgabe
    $htmlOutput = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Verbindungsaudit Bericht - $(Convert-ToDisplayString $serverInfo.ServerName)</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8f9fa; color: #333; line-height: 1.6; }
        .page-container { max-width: 1400px; margin: 0 auto; background-color: #ffffff; box-shadow: 0 0 20px rgba(0,0,0,0.1); }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px 40px; }
        .header-title { font-size: 2.2em; font-weight: 300; margin-bottom: 10px; }
        .header-info-cards-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-top: 25px; }
        .info-card { background-color: rgba(255,255,255,0.15); padding: 12px 18px; border-radius: 8px; backdrop-filter: blur(10px); }
        .main-content-wrapper { display: flex; min-height: 70vh; }
        .sidebar { width: 280px; background-color: #f8f9fa; border-right: 1px solid #e0e4e9; padding: 25px 0; }
        .nav-list { list-style: none; }
        .category-nav { margin: 0; }
        .nav-link { display: block; padding: 12px 25px; color: #495057; text-decoration: none; border-left: 4px solid transparent; transition: all 0.3s ease; }
        .nav-link:hover, .nav-link.active { background-color: #e3f2fd; color: #1976d2; border-left-color: #1976d2; }
        .content-area { flex: 1; padding: 25px 35px; overflow-y: auto; background-color: #ffffff; }
        .content-category-title { font-size: 1.6em; color: #005a9e; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 2px solid #eef1f5; }
        .tab-content { display: none; }
        .tab-content.active { display: block; animation: fadeIn 0.4s ease-in-out; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
        .section { margin-bottom: 30px; background: #ffffff; border-radius: 5px; border: 1px solid #e7eaf0; box-shadow: 0 1px 5px rgba(0,0,0,0.05); overflow: hidden; }
        .section-header { background: #f7f9fc; padding: 12px 18px; border-bottom: 1px solid #e7eaf0; }
        .section-title { font-size: 1.15em; font-weight: 600; color: #2c3e50; margin: 0; }
        .section-content { padding: 18px; }
        pre { background-color: #fdfdff; padding: 12px; border: 1px solid #e0e4e9; border-radius: 4px; white-space: pre-wrap; word-wrap: break-word; font-family: 'Consolas', 'Monaco', 'Courier New', monospace; font-size: 0.85em; line-height: 1.5; overflow-x: auto; color: #333; }
        .footer-timestamp { color: #505050; font-size: 0.8em; text-align: center; padding: 15px 40px; background-color: #e9ecef; border-top: 1px solid #d8dde3; }
        .footer-timestamp a { color: #005a9e; text-decoration: none; }
        .footer-timestamp a:hover { text-decoration: underline; }
    </style>
    <script>
        function showTab(tabId, clickedElement) {
            var i, contents, navLinks;
            contents = document.querySelectorAll('.tab-content');
            for (i = 0; i < contents.length; i++) {
                contents[i].classList.remove('active');
            }
            
            navLinks = document.querySelectorAll('.sidebar .category-nav .nav-link');
            for (i = 0; i < navLinks.length; i++) {
                navLinks[i].classList.remove('active');
            }
            
            var selectedTabContent = document.getElementById(tabId);
            if (selectedTabContent) {
                selectedTabContent.classList.add('active');
            }
            
            if (clickedElement) {
                clickedElement.classList.add('active');
            }
        }
        
        window.onload = function() {
            var firstNavLink = document.querySelector('.sidebar .nav-list .category-nav .nav-link');
            if (firstNavLink) {
                firstNavLink.click(); 
            } else {
                var firstContent = document.querySelector('.tab-content');
                if (firstContent) {
                    firstContent.classList.add('active');
                }
            }
        }
    </script>
</head>
<body>
    <div class="page-container">
        <header class="header">
            <h1 class="header-title">🌐 Verbindungsaudit Bericht</h1>
            <div class="header-info-cards-container">
                <div class="info-card"><strong>Hostname:</strong> $(Convert-ToDisplayString $serverInfo.ServerName)</div>
                <div class="info-card"><strong>Domaene:</strong> $(Convert-ToDisplayString $serverInfo.Domain)</div>
                <div class="info-card"><strong>Betriebssystem:</strong> $(Convert-ToDisplayString $serverInfo.OS)</div>
                <div class="info-card"><strong>CPU:</strong> $(Convert-ToDisplayString $serverInfo.CPU)</div>
                <div class="info-card"><strong>RAM:</strong> $(Convert-ToDisplayString $serverInfo.RAM)</div>
                <div class="info-card"><strong>Berichtsdatum:</strong> $($serverInfo.ReportDate)</div>
                <div class="info-card"><strong>Benutzer:</strong> $(Convert-ToDisplayString $serverInfo.User)</div>
            </div>
        </header>
        
        <div class="main-content-wrapper">
            <nav class="sidebar">
                <ul class="nav-list">
                    $sidebarNavLinks
                </ul>
            </nav>
            <main class="content-area">
                $mainContentTabs
            </main>
        </div>
        
        <footer class="footer-timestamp">
            Verbindungsaudit Bericht erstellt von easyWSAudit am $($serverInfo.ReportDate) | <a href="https://psscripts.de" target="_blank">PSscripts.de</a> | Andreas Hepp
        </footer>
    </div>
</body>
</html>
"@

    $htmlOutput | Out-File -FilePath $FilePath -Encoding utf8
    Write-DebugLog "Verbindungsaudit HTML-Export abgeschlossen" "Export"
}

# Funktion zum Generieren des HTML-Exports
function Export-AuditToHTML {
    param(
        [hashtable]$Results,
        [string]$FilePath
    )
    
    Write-DebugLog "Starte HTML-Export nach: $FilePath" "Export"

    # Helper to replace Umlaute and escape HTML
    function Convert-ToDisplayString {
        param([string]$Text)
        if ([string]::IsNullOrEmpty($Text)) { return "" }
        $processedText = $Text -replace 'ä', 'ae' -replace 'ö', 'oe' -replace 'ü', 'ue' -replace 'Ä', 'Ae' -replace 'Ö', 'Oe' -replace 'Ü', 'Ue' -replace 'ß', 'ss'
        # Weitere Sonderzeichen koennten hier bei Bedarf behandelt werden
        return [System.Security.SecurityElement]::Escape($processedText)
    }
    
    # Erweiterte Serverinformationen sammeln
    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cpuInfoObj = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $totalRamBytes = (Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue | Measure-Object -Property Capacity -Sum).Sum
    $driveC = Get-Volume -DriveLetter 'C' -ErrorAction SilentlyContinue
    
    $serverInfo = @{
        ServerName = $env:COMPUTERNAME
        ReportDate = Get-Date -Format "dd.MM.yyyy | HH:mm:ss" # Format beibehalten, da es ein Datum ist
        Domain = $env:USERDOMAIN
        User = $env:USERNAME
        OS = if ($osInfo) { "$($osInfo.Caption) $($osInfo.OSArchitecture)" } else { "N/A" }
        CPU = if ($cpuInfoObj) { $cpuInfoObj.Name } else { "N/A" }
        RAM = if ($totalRamBytes) { "{0:N2} GB" -f ($totalRamBytes / 1GB) } else { "N/A" }
        DiskCTotal = if ($driveC) { "{0:N2} GB" -f ($driveC.Size / 1GB) } else { "N/A" }
        DiskCFree = if ($driveC) { "{0:N2} GB" -f ($driveC.SizeRemaining / 1GB) } else { "N/A" }
    }

    # Gruppiere Ergebnisse nach Kategorien (gleiche Logik wie in der GUI)
    $groupedResults = @{}
    
    # Definiere die gewünschte Reihenfolge der Kategorien (System an erster Stelle)
    $categoryOrder = @(
        "System",
        "Hardware", 
        "Storage",
        "Network",
        "Security",
        "Services",
        "Tasks",
        "Events",
        "Features",
        "Software",
        "Updates",
        "Active-Directory",
        "DNS",
        "DHCP",
        "IIS",
        "WDS",
        "Hyper-V",
        "Cluster",
        "WSUS",
        "FileServices",
        "PrintServices",
        "RDS",
        "PKI",
        "ADFS",
        "ADLDS",
        "ADRMS",
        "DeviceAttestation",
        "VolumeActivation",
        "Backup",
        "NPAS",
        "HGS",
        "RemoteAccess",
        "InternalDB",
        "WindowsDefender",
        "WAS",
        "SearchService",
        "ServerEssentials",
        "Migration",
        "Identity",
        "FileSharing",
        "UserProfiles",
        "Firewall",
        "PowerManagement",
        "CredentialManager",
        "AuditPolicy",
        "GroupPolicy",
        "InstalledSoftware",
        "Environment"
    )
    
    if ($null -ne $commands) {
        foreach ($cmdDef in $commands) {
            # Verwende die Category direkt aus dem Befehl
            $categoryName = if ($cmdDef.Category) { $cmdDef.Category } else { "Allgemein" }
            
            if ($Results.ContainsKey($cmdDef.Name)) {
                if (-not $groupedResults.ContainsKey($categoryName)) {
                    $groupedResults[$categoryName] = @()
                }
                
                $groupedResults[$categoryName] += @{
                    Name = $cmdDef.Name
                    Result = $Results[$cmdDef.Name]
                    Command = $cmdDef
                }
            }
        }
    }

    # Navigationselemente und Tab-Inhalte generieren mit gewünschter Reihenfolge
    $sidebarNavLinks = ""
    $mainContentTabs = ""
    $firstTabId = $null
    
    # Erstelle eine sortierte Liste der Kategorien basierend auf der gewünschten Reihenfolge
    $sortedCategories = @()
    
    # Füge Kategorien in der gewünschten Reihenfolge hinzu
    foreach ($orderCat in $categoryOrder) {
        if ($groupedResults.ContainsKey($orderCat)) {
            $sortedCategories += $orderCat
        }
    }
    
    # Füge alle anderen Kategorien alphabetisch sortiert hinzu
    foreach ($categoryKey in ($groupedResults.Keys | Sort-Object)) {
        if ($categoryKey -notin $sortedCategories) {
            $sortedCategories += $categoryKey
        }
    }
    
    foreach ($categoryKey in $sortedCategories) {
        $items = $groupedResults[$categoryKey]
        $displayCategory = Convert-ToDisplayString $categoryKey
        
        # Spezielle Anzeigenamen für bessere Lesbarkeit
        $displayCategoryName = switch ($categoryKey) {
            "System" { "System-Informationen" }
            "Hardware" { "Hardware-Informationen" }
            "Storage" { "Speicher & Festplatten" }
            "Network" { "Netzwerk-Konfiguration" }
            "Security" { "Sicherheits-Einstellungen" }
            "Services" { "Dienste & Services" }
            "Tasks" { "Geplante Aufgaben" }
            "Events" { "Ereignisprotokoll" }
            "Features" { "Installierte Features" }
            "Software" { "Software & Programme" }
            "Updates" { "Windows Updates" }
            "Active-Directory" { "Active Directory" }
            "DNS" { "DNS-Server" }
            "DHCP" { "DHCP-Server" }
            "IIS" { "Internet Information Services (IIS)" }
            "WDS" { "Windows Deployment Services" }
            "Hyper-V" { "Hyper-V Virtualisierung" }
            "Cluster" { "Failover Clustering" }
            "WSUS" { "Windows Server Update Services" }
            "FileServices" { "Datei-Services" }
            "PrintServices" { "Druck-Services" }
            "RDS" { "Remote Desktop Services" }
            "PKI" { "Zertifikat-Services (PKI)" }
            "ADFS" { "Active Directory Federation Services" }
            "ADLDS" { "AD Lightweight Directory Services" }
            "ADRMS" { "AD Rights Management Services" }
            "DeviceAttestation" { "Device Health Attestation" }
            "VolumeActivation" { "Volume Activation Services" }
            "Backup" { "Windows Server Backup" }
            "NPAS" { "Network Policy and Access Services" }
            "HGS" { "Host Guardian Service" }
            "RemoteAccess" { "Remote Access Services" }
            "InternalDB" { "Windows Internal Database" }
            "WindowsDefender" { "Windows Defender" }
            "WAS" { "Windows Process Activation Service" }
            "SearchService" { "Windows Search Service" }
            "ServerEssentials" { "Windows Server Essentials" }
            "Migration" { "Migration Services" }
            "Identity" { "Windows Identity Foundation" }
            "FileSharing" { "Dateifreigaben" }
            "UserProfiles" { "Benutzer-Profile" }
            "Firewall" { "Windows Firewall" }
            "PowerManagement" { "Energieverwaltung" }
            "CredentialManager" { "Anmeldeinformationsverwaltung" }
            "AuditPolicy" { "Audit-Richtlinien" }
            "GroupPolicy" { "Gruppenrichtlinien" }
            "InstalledSoftware" { "Installierte Software" }
            "Environment" { "Umgebungsvariablen" }
            default { $displayCategory }
        }
        
        $categoryIdPart = $categoryKey -replace '[^a-zA-Z0-9_]', ''
        if ($categoryIdPart.Length -eq 0) { 
            $categoryIdPart = "cat" + ($categoryKey.GetHashCode() | ForEach-Object ToString X) 
        }
        $tabId = "tab_$categoryIdPart"

        if ($null -eq $firstTabId) { $firstTabId = $tabId }

        $sidebarNavLinks += @"
<li class="nav-item category-nav">
    <a href="#" class="nav-link" onclick="showTab('$tabId', this)">
        $displayCategoryName ($($items.Count))
    </a>
</li>
"@
        
        $tabContent = "<div id='$tabId' class='tab-content'>"
        $tabContent += "<h2 class='content-category-title'>$displayCategoryName</h2>"

        foreach ($item in $items) {
            $displayItemName = Convert-ToDisplayString $item.Name
            $displayItemResult = Convert-ToDisplayString $item.Result
            
            # Füge Kommando-Information hinzu, falls verfügbar
            $commandInfo = ""
            if ($item.Command -and $item.Command.Command) {
                $commandInfo = "<p class='command-info'><strong>Befehl:</strong> <code>$(Convert-ToDisplayString $item.Command.Command)</code></p>"
            }
            
            $tabContent += @"
<div class="section">
    <div class="section-header">
        <h3 class="section-title">$displayItemName</h3>
    </div>
    <div class="section-content">
        $commandInfo
        <pre>$displayItemResult</pre>
    </div>
</div>
"@
        }
        $tabContent += "</div>"
        $mainContentTabs += $tabContent
    }

    # HTML-Gesamtstruktur
    $htmlOutput = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <title>$(Convert-ToDisplayString "Windows Server Audit Bericht - $($serverInfo.ServerName)")</title>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 0; 
            background-color: #f0f2f5; /* Etwas hellerer Hintergrund */
            color: #333;
            line-height: 1.6;
        }
        .page-container {
            display: flex;
            flex-direction: column;
            min-height: 100vh;
        }
        .header {
            background: linear-gradient(135deg, #005a9e, #003966); /* Dunkleres Blau */
            color: white;
            padding: 20px 40px;
            display: flex;
            flex-direction: column; /* Ermoeglicht Top-Row und Info-Cards untereinander */
            align-items: center; /* Zentriert Info-Cards, falls sie schmaler sind */
        }
        .header-top-row {
            display: flex;
            align-items: center;
            width: 100%;
            margin-bottom: 15px;
        }
        .header-logo {
            width: 125px;
            height: 75px;
            background-color: #e0e0e0;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            color: #333;
            margin-right: 20px;
            flex-shrink: 0;
            border-radius: 4px;
        }
        .header-title { 
            margin: 0; 
            font-size: 2em; 
            font-weight: 500; /* Etwas staerker */
        }
        
        .header-info-cards-container {
            display: flex;
            flex-wrap: wrap;
            justify-content: center; /* Zentriert die Karten */
            gap: 15px; /* Abstand zwischen den Karten */
            padding: 10px 0;
            width: 100%;
            max-width: 1200px; /* Begrenzt die Breite der Kartenreihe */
        }
        .info-card {
            background-color: rgba(255, 255, 255, 0.1); /* Leicht transparente Karten */
            border: 1px solid rgba(255, 255, 255, 0.2);
            border-radius: 6px;
            padding: 10px 15px;
            font-size: 0.85em;
            color: white; /* Textfarbe auf den Karten */
            min-width: 150px; /* Mindestbreite fuer Karten */
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .info-card strong {
            display: block;
            font-size: 0.9em;
            color: #b0dfff; /* Hellere Farbe fuer Label */
            margin-bottom: 3px;
        }

        .main-content-wrapper {
            display: flex;
            flex: 1;
            background-color: #f0f2f5; 
            margin: 0;
        }

        .sidebar {
            width: 280px; /* Etwas breiter fuer tiefere Navigation */
            background-color: #ffffff; 
            padding: 20px;
            border-right: 1px solid #d8dde3;
            overflow-y: auto; 
            box-shadow: 2px 0 5px rgba(0,0,0,0.05);
            flex-shrink: 0;
        }
        .sidebar .nav-list {
            list-style: none;
            padding: 0;
            margin: 0;
        }
        .sidebar .nav-item {
            margin-bottom: 6px;
        }
        .sidebar .category-nav .nav-link {
            display: block;
            padding: 12px 15px;
            text-decoration: none;
            color: #33475b; 
            border-radius: 6px;
            transition: background-color 0.2s ease, color 0.2s ease;
            font-size: 0.95em;
            font-weight: 500;
            word-break: break-word;
            border-left: 3px solid transparent;
        }
        .sidebar .category-nav .nav-link:hover {
            background-color: #e9ecef;
            color: #005a9e;
            border-left-color: #005a9e;
        }
        .sidebar .category-nav .nav-link.active {
            background-color: #0078d4;
            color: white;
            font-weight: 600;
            border-left-color: #ffffff;
        }

        .content-area {
            flex: 1; 
            padding: 25px 35px; 
            overflow-y: auto;
            background-color: #ffffff; 
        }
        .content-category-title { /* Stil fuer den Titel im Inhaltsbereich */
            font-size: 1.6em;
            color: #005a9e;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #eef1f5;
        }
        .tab-content {
            display: none;
        }
        .tab-content.active {
            display: block;
            animation: fadeIn 0.4s ease-in-out;
        }
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(8px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .section {
            margin-bottom: 30px;
            background: #ffffff; 
            border-radius: 5px;
            border: 1px solid #e7eaf0; 
            box-shadow: 0 1px 5px rgba(0,0,0,0.05); 
            overflow: hidden;
        }
        .section-header {
            background: #f7f9fc; 
            padding: 12px 18px; /* Etwas kompakter */
            border-bottom: 1px solid #e7eaf0;
        }
        .section-title {
            font-size: 1.15em; 
            font-weight: 600;
            color: #2c3e50; 
            margin: 0;
        }
        .section-content {
            padding: 18px;
        }
        .command-info {
            background-color: #f8f9fa;
            border-left: 4px solid #28a745;
            padding: 10px 15px;
            margin-bottom: 15px;
            font-size: 0.9em;
        }
        .command-info code {
            background-color: #e9ecef;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 0.85em;
            color: #495057;
        }
        pre { 
            background-color: #fdfdff; 
            padding: 12px; 
            border: 1px solid #e0e4e9; 
            border-radius: 4px;
            white-space: pre-wrap; 
            word-wrap: break-word;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 0.85em; 
            line-height: 1.5;
            overflow-x: auto;
            color: #333;
        }

        .footer-timestamp { 
            color: #505050; /* Dunklerer Text */
            font-size: 0.8em;
            text-align: center;
            padding: 15px 40px;
            background-color: #e9ecef; /* Passend zum Rest */
            border-top: 1px solid #d8dde3;
        }
        .footer-timestamp a {
            color: #005a9e;
            text-decoration: none;
        }
        .footer-timestamp a:hover {
            text-decoration: underline;
        }
    </style>
    <script>
        function showTab(tabId, clickedElement) {
            var i, contents, navLinks;
            contents = document.querySelectorAll('.tab-content');
            for (i = 0; i < contents.length; i++) {
                contents[i].classList.remove('active');
            }
            
            navLinks = document.querySelectorAll('.sidebar .category-nav .nav-link');
            for (i = 0; i < navLinks.length; i++) {
                navLinks[i].classList.remove('active');
            }
            
            var selectedTabContent = document.getElementById(tabId);
            if (selectedTabContent) {
                selectedTabContent.classList.add('active');
            }
            
            if (clickedElement) {
                clickedElement.classList.add('active');
            }
        }
        
        window.onload = function() {
            // Ersten Kategorie-Link automatisch aktivieren
            var firstNavLink = document.querySelector('.sidebar .nav-list .category-nav .nav-link');
            if (firstNavLink) {
                firstNavLink.click(); 
            } else {
                // Fallback, falls keine Links vorhanden sind
                var firstContent = document.querySelector('.tab-content');
                if (firstContent) {
                    firstContent.classList.add('active');
                }
            }
        }
    </script>
</head>
<body>
    <div class="page-container">
        <header class="header">
            <div class="header-top-row">
                <div class="header-logo">LOGO</div>
                <h1 class="header-title">$(Convert-ToDisplayString "Windows Server Audit Bericht")</h1>
            </div>
            <div class="header-info-cards-container">
                <div class="info-card"><strong>Hostname:</strong> $(Convert-ToDisplayString $serverInfo.ServerName)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Domäne"):</strong> $(Convert-ToDisplayString $serverInfo.Domain)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Betriebssystem"):</strong> $(Convert-ToDisplayString $serverInfo.OS)</div>
                <div class="info-card"><strong>CPU:</strong> $(Convert-ToDisplayString $serverInfo.CPU)</div>
                <div class="info-card"><strong>RAM:</strong> $(Convert-ToDisplayString $serverInfo.RAM)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Festplatte C: Gesamt"):</strong> $(Convert-ToDisplayString $serverInfo.DiskCTotal)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Festplatte C: Frei"):</strong> $(Convert-ToDisplayString $serverInfo.DiskCFree)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Berichtsdatum"):</strong> $($serverInfo.ReportDate)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Benutzer"):</strong> $(Convert-ToDisplayString $serverInfo.User)</div>
            </div>
        </header>
        
        <div class="main-content-wrapper">
            <nav class="sidebar">
                <ul class="nav-list">
                    $sidebarNavLinks
                </ul>
            </nav>
            <main class="content-area">
                $mainContentTabs
            </main>
        </div>
        
        <footer class="footer-timestamp">
            $(Convert-ToDisplayString "Audit Bericht erstellt von easyWSAudit am $($serverInfo.ReportDate)") | <a href="https://psscripts.de" target="_blank">PSscripts.de</a> | Andreas Hepp
        </footer>
    </div>
</body>
</html>
"@

    $htmlOutput | Out-File -FilePath $FilePath -Encoding utf8
    Write-DebugLog "HTML-Export abgeschlossen" "Export"
}

# Variable fuer die Audit-Ergebnisse
$global:auditResults = @{}

# XAML UI Definition - Vereinfachte Version
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="easyWSAudit - Windows Server Audit Tool"
    Width="1420"
    Height="1000"
    MinWidth="1200"
    MinHeight="700"
    Background="#F5F5F5"
    FontFamily="Segoe UI"
    FontSize="12"
    WindowStartupLocation="CenterScreen">

    <Window.Resources>
        <!--  Modern Button Style  -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#007acc" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="15,8" />
            <Setter Property="Margin" Value="5" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="FontSize" Value="12" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="MinWidth" Value="80" />
            <Setter Property="MinHeight" Value="32" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border
                            x:Name="border"
                            Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="3">
                            <ContentPresenter
                                Margin="{TemplateBinding Padding}"
                                HorizontalAlignment="Center"
                                VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#005a9e" />
                                <Setter TargetName="border" Property="BorderBrush" Value="#005a9e" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#004578" />
                                <Setter TargetName="border" Property="BorderBrush" Value="#004578" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.6" />
                                <Setter TargetName="border" Property="Background" Value="#cccccc" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!--  Navigation Button Style  -->
        <Style x:Key="NavButton" TargetType="Button">
            <Setter Property="Background" Value="Transparent" />
            <Setter Property="Foreground" Value="#FFFFFF" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="20,12" />
            <Setter Property="Margin" Value="5,2" />
            <Setter Property="HorizontalAlignment" Value="Stretch" />
            <Setter Property="HorizontalContentAlignment" Value="Left" />
            <Setter Property="FontSize" Value="14" />
            <Setter Property="FontWeight" Value="Normal" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border
                            x:Name="border"
                            Margin="2"
                            Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="4">
                            <ContentPresenter
                                Margin="{TemplateBinding Padding}"
                                HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#1A2332" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#007acc" />
                                <Setter Property="Foreground" Value="White" />
                            </Trigger>
                            <Trigger Property="Tag" Value="active">
                                <Setter TargetName="border" Property="Background" Value="#007acc" />
                                <Setter Property="FontWeight" Value="SemiBold" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!--  Card Style  -->
        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="White" />
            <Setter Property="BorderBrush" Value="#E1E5E9" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="CornerRadius" Value="4" />
            <Setter Property="Padding" Value="20" />
            <Setter Property="Margin" Value="10" />
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect
                        BlurRadius="3"
                        Direction="270"
                        Opacity="0.1"
                        ShadowDepth="1"
                        Color="#000000" />
                </Setter.Value>
            </Setter>
        </Style>

        <!--  TextBox Style  -->
        <Style x:Key="ModernTextBox" TargetType="TextBox">
            <Setter Property="BorderBrush" Value="#CCCCCC" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="Padding" Value="8,6" />
            <Setter Property="Margin" Value="5" />
            <Setter Property="Background" Value="White" />
            <Setter Property="FontSize" Value="12" />
            <Setter Property="MinHeight" Value="28" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border
                            Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="3">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#007acc" />
                                <Setter Property="BorderThickness" Value="2" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.6" />
                                <Setter Property="Background" Value="#F0F0F0" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="CategoryHeader" TargetType="TextBlock">
            <Setter Property="FontSize" Value="15" />
            <Setter Property="FontWeight" Value="Bold" />
            <Setter Property="Foreground" Value="#FFFFFF" />
            <Setter Property="Background" Value="#0078D4" />
            <Setter Property="Padding" Value="12,6" />
            <Setter Property="Margin" Value="0,12,0,4" />
        </Style>

        <Style x:Key="CheckboxStyle" TargetType="CheckBox">
            <Setter Property="Margin" Value="15,3,8,3" />
            <Setter Property="Padding" Value="6,0,0,0" />
            <Setter Property="VerticalAlignment" Value="Center" />
            <Setter Property="FontSize" Value="12" />
        </Style>

        <Style x:Key="CompactCheckboxStyle" TargetType="CheckBox">
            <Setter Property="Margin" Value="8,2,4,2" />
            <Setter Property="Padding" Value="4,0,0,0" />
            <Setter Property="VerticalAlignment" Value="Center" />
            <Setter Property="FontSize" Value="11" />
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="60" />
            <RowDefinition Height="*" />
            <RowDefinition Height="40" />
        </Grid.RowDefinitions>

        <!--  Header  -->
        <Border
            Grid.Row="0"
            Background="#FF1C323C"
            BorderBrush="#E0E0E0"
            BorderThickness="0,0,0,1">
            <Grid Margin="20,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" />
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>

                <!--  App Name  -->
                <StackPanel
                    Grid.Column="0"
                    VerticalAlignment="Center"
                    Orientation="Horizontal">
                    <TextBlock
                        VerticalAlignment="Center"
                        FontSize="20"
                        FontWeight="SemiBold"
                        Foreground="#e2e2e2"
                        Text="easyWSAudit" />
                    <TextBlock
                        Margin="10,0,0,0"
                        VerticalAlignment="Center"
                        FontSize="14"
                        Foreground="#CCE7FF"
                        Text="Windows Server Audit Tool" />
                </StackPanel>

                <TextBlock
                    x:Name="txtStatus"
                    Grid.Column="1"
                    VerticalAlignment="Center"
                    Foreground="#e2e2e2"
                    Text="Status: READY" Margin="100,0,0,0" />

                <!--  Server Info and Main Button  -->
                <StackPanel
                    Grid.Column="2"
                    VerticalAlignment="Center"
                    Orientation="Horizontal">
                    <TextBlock
                        x:Name="txtServerName"
                        Margin="0,0,20,0"
                        VerticalAlignment="Center"
                        FontSize="14"
                        Foreground="White"
                        Text="" />
                    <TextBlock
                        x:Name="txtStatusHeader"
                        Margin="0,0,20,0"
                        VerticalAlignment="Center"
                        FontSize="12"
                        Foreground="#90EE90"
                        Text="Status: Ready" />
                    <Button
                        x:Name="btnFullAudit"
                        Background="#28A745"
                        Content="Complete Audit"
                        Style="{StaticResource ModernButton}" />
                </StackPanel>
            </Grid>
        </Border>
        
        <!--  Main Content  -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="250" />
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>

            <!--  Navigation Panel  -->
            <Border
                Grid.Column="0"
                Background="#FF334556"
                BorderBrush="#1E2A38"
                BorderThickness="0,0,1,0">
                <StackPanel Margin="0,10">
                    <!--  User Info  -->
                    <GroupBox
                        Margin="5,8,5,8"
                        Padding="8"
                        BorderBrush="#4A5C6E"
                        BorderThickness="1"
                        Foreground="#E0E0E0"
                        Header="Server">
                        <StackPanel>
                            <TextBlock
                                x:Name="txtNavServerName"
                                HorizontalAlignment="Left"
                                VerticalAlignment="Center"
                                FontSize="12"
                                FontWeight="SemiBold"
                                Foreground="#FFFFFF"
                                Text="Local" />
                        </StackPanel>
                    </GroupBox>

                    <!--  Trennlinie  -->
                    <Border
                        Height="1"
                        Margin="16,0,16,25"
                        Background="#4A5C6E"
                        Opacity="0.5" />

                    <!--  Updated Navigation Buttons  -->
                    <TextBlock
                        Margin="15,0,0,5"
                        FontSize="11"
                        Foreground="#A0A0A0"
                        Text="SERVER" />
                    <Button
                        x:Name="btnNavServerAudit"
                        Content="Server Audit"
                        Style="{StaticResource NavButton}"
                        Tag="serverAudit" />
                    <Button
                        x:Name="btnNavServerResults"
                        Content="Server Audit Results"
                        Style="{StaticResource NavButton}"
                        Tag="serverResults" />

                    <TextBlock
                        Margin="15,15,0,5"
                        FontSize="11"
                        Foreground="#A0A0A0"
                        Text="NETWORK" />
                    <Button
                        x:Name="btnNavNetworkAudit"
                        Content="Netzwerk Audit"
                        Style="{StaticResource NavButton}"
                        Tag="networkAudit" />
                    <Button
                        x:Name="btnNavNetworkResults"
                        Content="Netzwerk Audit Results"
                        Style="{StaticResource NavButton}"
                        Tag="networkResults" />

                    <TextBlock
                        Margin="15,15,0,5"
                        FontSize="11"
                        Foreground="#A0A0A0"
                        Text="OTHER" />
                    <Button
                        x:Name="btnNavTools"
                        Content="Tools"
                        Style="{StaticResource NavButton}"
                        Tag="tools" />

                    <TextBlock
                        Margin="15,15,0,5"
                        FontSize="11"
                        Foreground="#A0A0A0"
                        Text="SYSTEM" />
                    <Button
                        x:Name="btnNavDebug"
                        Content="Debug"
                        Style="{StaticResource NavButton}"
                        Tag="debug" />
                </StackPanel>
            </Border>
            <!--  Content Panel  -->
            <Border
                Grid.Column="1"
                Padding="20"
                Background="#FFFFFF">
                <Grid Name="contentGrid">
                    <!--  Server Audit Panel  -->
                    <Grid Name="serverAuditPanel" Visibility="Visible">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="400" />
                            <ColumnDefinition Width="*" />
                        </Grid.ColumnDefinitions>

                        <!--  Linke Seite - Optionen  -->
                        <Border
                            Grid.Column="0"
                            Margin="0,0,10,0"
                            Style="{StaticResource Card}">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto" />
                                    <RowDefinition Height="*" />
                                    <RowDefinition Height="Auto" />
                                </Grid.RowDefinitions>
                                <StackPanel Grid.Row="0">
                                    <TextBlock
                                        Margin="0,0,0,10"
                                        FontSize="18"
                                        FontWeight="SemiBold"
                                        Text="Server-Audit-Categories" />
                                    <TextBlock
                                        Margin="0,0,0,15"
                                        FontSize="12"
                                        Foreground="#6C757D"
                                        Text="Please select a Audit"
                                        TextWrapping="Wrap" />
                                    <Separator Margin="0,0,0,10" Background="#E1E5E9" />
                                </StackPanel>

                                <ScrollViewer
                                    Grid.Row="1"
                                    MaxHeight="400"
                                    Margin="0,0,0,15"
                                    HorizontalScrollBarVisibility="Disabled"
                                    VerticalScrollBarVisibility="Auto">
                                    <StackPanel x:Name="spOptions" />
                                </ScrollViewer>

                                <StackPanel Grid.Row="2">
                                    <Button
                                        x:Name="btnSelectAll"
                                        Margin="0,0,0,5"
                                        Background="#28A745"
                                        Content="Alle auswaehlen"
                                        Style="{StaticResource ModernButton}" />
                                    <Button
                                        x:Name="btnSelectNone"
                                        Margin="0,0,0,25"
                                        Background="#DC3545"
                                        Content="Alle abwaehlen"
                                        Style="{StaticResource ModernButton}" />
                                    <Button
                                        x:Name="btnRunAudit"
                                        Background="#0078D4"
                                        Content="Server-Audit starten"
                                        FontWeight="Bold"
                                        Foreground="White"
                                        Style="{StaticResource ModernButton}" />
                                </StackPanel>
                            </Grid>
                        </Border>

                        <!--  Rechte Seite - Fortschritt  -->
                        <Border
                            Grid.Column="1"
                            Margin="10,0,0,0"
                            Style="{StaticResource Card}">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto" />
                                    <RowDefinition Height="Auto" />
                                    <RowDefinition Height="*" />
                                </Grid.RowDefinitions>

                                <TextBlock
                                    Grid.Row="0"
                                    Margin="0,0,0,15"
                                    FontSize="18"
                                    FontWeight="SemiBold"
                                    Text="Server-Audit-Progress" />

                                <StackPanel Grid.Row="1" Margin="0,0,0,15">
                                    <ProgressBar
                                        x:Name="progressBar"
                                        Height="20"
                                        Margin="0,0,0,10" />
                                    <TextBlock
                                        x:Name="txtProgress"
                                        HorizontalAlignment="Center"
                                        FontSize="12"
                                        Foreground="#666"
                                        Text="Ready for Server-Audit" />
                                </StackPanel>

                                <Border
                                    Grid.Row="2"
                                    Padding="15"
                                    Background="#F8F9FA"
                                    CornerRadius="4">
                                    <ScrollViewer x:Name="scrollStatusLog" VerticalScrollBarVisibility="Auto">
                                        <TextBlock
                                            x:Name="txtStatusLog"
                                            FontFamily="Consolas"
                                            FontSize="11"
                                            Foreground="#495057"
                                            Text="Bereit..." />
                                    </ScrollViewer>
                                </Border>
                            </Grid>
                        </Border>
                    </Grid>

                    <!--  Server Audit Results Panel  -->
                    <Grid Name="serverResultsPanel" Visibility="Collapsed">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="*" />
                        </Grid.RowDefinitions>

                        <TextBlock
                            Grid.Row="0"
                            Margin="0,0,0,20"
                            FontSize="24"
                            FontWeight="SemiBold"
                            Foreground="#1C1C1C"
                            Text="Server-Audit-Ergebnisse" />

                        <!--  Toolbar mit Buttons  -->
                        <StackPanel
                            Grid.Row="1"
                            Margin="0,0,0,15"
                            Orientation="Horizontal">
                            <Button
                                x:Name="btnExportHTML"
                                Background="#17A2B8"
                                Content="📊 HTML Export"
                                IsEnabled="False"
                                Style="{StaticResource ModernButton}" />
                            <Button
                                x:Name="btnExportCSV"
                                Margin="5,0,0,0"
                                Background="#28A745"
                                Content="CSV Export"
                                IsEnabled="False"
                                Style="{StaticResource ModernButton}" />
                            <Button
                                x:Name="btnExportJSON"
                                Margin="5,0,0,0"
                                Background="#6F42C1"
                                Content="JSON Export"
                                IsEnabled="False"
                                Style="{StaticResource ModernButton}" />
                            <Button
                                x:Name="btnExportSummary"
                                Margin="5,0,0,0"
                                Background="#FD7E14"
                                Content="Executive Summary"
                                IsEnabled="False"
                                Style="{StaticResource ModernButton}" />
                            <Button
                                x:Name="btnCopyToClipboard"
                                Margin="10,0,0,0"
                                Background="#6C757D"
                                Content="Zwischenablage"
                                IsEnabled="False"
                                Style="{StaticResource ModernButton}" />
                            <Button
                                x:Name="btnRefreshResults"
                                Margin="5,0,0,0"
                                Background="#20C997"
                                Content="Aktualisieren"
                                IsEnabled="False"
                                Style="{StaticResource ModernButton}" />
                        </StackPanel>

                        <!--  Kategorien-Auswahl  -->
                        <Grid Grid.Row="2" Margin="0,0,0,15">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto" />
                                <ColumnDefinition Width="250" />
                                <ColumnDefinition Width="*" />
                                <ColumnDefinition Width="Auto" />
                            </Grid.ColumnDefinitions>

                            <TextBlock
                                Grid.Column="0"
                                Margin="0,0,10,0"
                                VerticalAlignment="Center"
                                FontWeight="SemiBold"
                                Text="Kategorie:" />
                            <ComboBox
                                x:Name="cmbResultCategories"
                                Grid.Column="1"
                                Height="30"
                                VerticalAlignment="Center" />
                            <TextBlock
                                x:Name="txtResultsSummary"
                                Grid.Column="3"
                                VerticalAlignment="Center"
                                FontSize="12"
                                Foreground="#6C757D"
                                Text="" />
                        </Grid>

                        <!--  Ergebnisse-Anzeige  -->
                        <Border Grid.Row="3" Style="{StaticResource Card}">
                            <ScrollViewer
                                Margin="0"
                                Padding="0"
                                HorizontalScrollBarVisibility="Disabled"
                                VerticalScrollBarVisibility="Auto">
                                <RichTextBox
                                    x:Name="rtbResults"
                                    Padding="20"
                                    HorizontalAlignment="Stretch"
                                    VerticalAlignment="Stretch"
                                    Background="Transparent"
                                    BorderThickness="0"
                                    FontFamily="Segoe UI"
                                    FontSize="12"
                                    IsReadOnly="True" />
                            </ScrollViewer>
                        </Border>
                    </Grid>

                    <!--  Network Audit Panel  -->
                    <Grid Name="networkAuditPanel" Visibility="Collapsed">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="400" />
                            <ColumnDefinition Width="*" />
                        </Grid.ColumnDefinitions>

                        <!--  Linke Seite - Netzwerkaudit Optionen  -->
                        <Border
                            Grid.Column="0"
                            Margin="0,0,10,0"
                            Style="{StaticResource Card}">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto" />
                                    <RowDefinition Height="*" />
                                    <RowDefinition Height="Auto" />
                                </Grid.RowDefinitions>
                                <StackPanel Grid.Row="0">
                                    <TextBlock
                                        Margin="0,0,0,10"
                                        FontSize="18"
                                        FontWeight="SemiBold"
                                        Text="Netzwerk-Audit-Kategorien" />
                                    <TextBlock
                                        Margin="0,0,0,15"
                                        FontSize="12"
                                        Foreground="#6C757D"
                                        Text="Analyse aktiver Netzwerkverbindungen, Geraete und Benutzer"
                                        TextWrapping="Wrap" />
                                    <Separator Margin="0,0,0,10" Background="#E1E5E9" />
                                </StackPanel>

                                <ScrollViewer
                                    Grid.Row="1"
                                    MaxHeight="400"
                                    Margin="0,0,0,15"
                                    HorizontalScrollBarVisibility="Disabled"
                                    VerticalScrollBarVisibility="Auto">
                                    <StackPanel x:Name="spConnectionOptions" />
                                </ScrollViewer>

                                <StackPanel Grid.Row="2">
                                    <Button
                                        x:Name="btnSelectAllConnection"
                                        Margin="0,0,0,5"
                                        Background="#28A745"
                                        Content="Alle auswaehlen"
                                        Style="{StaticResource ModernButton}" />
                                    <Button
                                        x:Name="btnSelectNoneConnection"
                                        Margin="0,0,0,25"
                                        Background="#DC3545"
                                        Content="Alle abwaehlen"
                                        Style="{StaticResource ModernButton}" />
                                    <Button
                                        x:Name="btnRunConnectionAudit"
                                        Background="#FD7E14"
                                        Content="Netzwerk-Audit starten"
                                        FontWeight="Bold"
                                        Foreground="White"
                                        Style="{StaticResource ModernButton}" />
                                </StackPanel>
                            </Grid>
                        </Border>

                        <!--  Rechte Seite - Fortschritt  -->
                        <Border
                            Grid.Column="1"
                            Margin="10,0,0,0"
                            Style="{StaticResource Card}">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto" />
                                    <RowDefinition Height="Auto" />
                                    <RowDefinition Height="*" />
                                </Grid.RowDefinitions>

                                <TextBlock
                                    Grid.Row="0"
                                    Margin="0,0,0,15"
                                    FontSize="18"
                                    FontWeight="SemiBold"
                                    Text="Netzwerk-Audit-Fortschritt" />

                                <StackPanel Grid.Row="1" Margin="0,0,0,15">
                                    <ProgressBar
                                        x:Name="progressBarConnection"
                                        Height="20"
                                        Margin="0,0,0,10" />
                                    <TextBlock
                                        x:Name="txtProgressConnection"
                                        HorizontalAlignment="Center"
                                        FontSize="12"
                                        Foreground="#666"
                                        Text="Bereit fuer Netzwerk-Audit" />
                                </StackPanel>

                                <Border
                                    Grid.Row="2"
                                    Padding="15"
                                    Background="#F8F9FA"
                                    CornerRadius="4">
                                    <ScrollViewer x:Name="scrollNetworkStatusLog" VerticalScrollBarVisibility="Auto">
                                        <TextBlock
                                            x:Name="txtNetworkStatusLog"
                                            FontFamily="Consolas"
                                            FontSize="11"
                                            Foreground="#495057"
                                            Text="Bereit..." />
                                    </ScrollViewer>
                                </Border>
                            </Grid>
                        </Border>
                    </Grid>

                    <!--  Network Audit Results Panel  -->
                    <Grid Name="networkResultsPanel" Visibility="Collapsed">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="*" />
                        </Grid.RowDefinitions>

                        <TextBlock
                            Grid.Row="0"
                            Margin="0,0,0,20"
                            FontSize="24"
                            FontWeight="SemiBold"
                            Foreground="#1C1C1C"
                            Text="Netzwerk-Audit-Ergebnisse" />

                        <!--  Toolbar mit Buttons  -->
                        <StackPanel
                            Grid.Row="1"
                            Margin="0,0,0,15"
                            Orientation="Horizontal">
                            <Button
                                x:Name="btnExportConnectionHTML"
                                Background="#17A2B8"
                                Content="HTML Export"
                                IsEnabled="False"
                                Style="{StaticResource ModernButton}" />
                            <Button
                                x:Name="btnExportConnectionDrawIO"
                                Margin="10,0,0,0"
                                Background="#28A745"
                                Content="DRAW.IO Export"
                                IsEnabled="False"
                                Style="{StaticResource ModernButton}" />
                            <Button
                                x:Name="btnCopyConnectionToClipboard"
                                Margin="10,0,0,0"
                                Background="#6C757D"
                                Content="Zwischenablage"
                                IsEnabled="False"
                                Style="{StaticResource ModernButton}" />
                        </StackPanel>

                        <!--  Kategorien-Auswahl  -->
                        <Grid Grid.Row="2" Margin="0,0,0,15">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto" />
                                <ColumnDefinition Width="250" />
                                <ColumnDefinition Width="*" />
                            </Grid.ColumnDefinitions>

                            <TextBlock
                                Grid.Column="0"
                                Margin="0,0,10,0"
                                VerticalAlignment="Center"
                                FontWeight="SemiBold"
                                Text="Kategorie:" />
                            <ComboBox
                                x:Name="cmbConnectionCategories"
                                Grid.Column="1"
                                Height="30"
                                VerticalAlignment="Center" />
                        </Grid>

                        <!--  Ergebnisse-Anzeige  -->
                        <Border
                            Grid.Row="3"
                            Background="#F8F9FA"
                            BorderBrush="#DEE2E6"
                            BorderThickness="1"
                            CornerRadius="4">
                            <ScrollViewer
                                Margin="0"
                                Padding="0"
                                HorizontalScrollBarVisibility="Disabled"
                                VerticalScrollBarVisibility="Auto">
                                <RichTextBox
                                    x:Name="rtbConnectionResults"
                                    Padding="20"
                                    HorizontalAlignment="Stretch"
                                    VerticalAlignment="Stretch"
                                    Background="Transparent"
                                    BorderThickness="0"
                                    FontFamily="Segoe UI"
                                    FontSize="12"
                                    IsReadOnly="True" />
                            </ScrollViewer>
                        </Border>
                    </Grid>

                    <!--  Debug Panel  -->
                    <Grid Name="debugPanel" Visibility="Collapsed">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="*" />
                            <RowDefinition Height="Auto" />
                        </Grid.RowDefinitions>

                        <TextBlock
                            Grid.Row="0"
                            Margin="0,0,0,20"
                            FontSize="24"
                            FontWeight="SemiBold"
                            Foreground="#1C1C1C"
                            Text="Debug-Informationen" />

                        <Border Grid.Row="1" Style="{StaticResource Card}">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto" />
                                    <RowDefinition Height="*" />
                                </Grid.RowDefinitions>

                                <TextBlock
                                    Grid.Row="0"
                                    Margin="0,0,0,15"
                                    FontSize="16"
                                    FontWeight="SemiBold"
                                    Foreground="#1C1C1C"
                                    Text="Terminal-Benachrichtigungen und Fehler" />

                                <Border
                                    Grid.Row="1"
                                    Background="#1E1E1E"
                                    BorderBrush="#333333"
                                    BorderThickness="1"
                                    CornerRadius="8">
                                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                                        <TextBox
                                            x:Name="txtDebugOutput"
                                            Padding="15"
                                            Background="Transparent"
                                            BorderThickness="0"
                                            FontFamily="Consolas"
                                            FontSize="11"
                                            Foreground="#00FF00"
                                            IsReadOnly="True"
                                            TextWrapping="Wrap" />
                                    </ScrollViewer>
                                </Border>
                            </Grid>
                        </Border>

                        <StackPanel
                            Grid.Row="2"
                            Margin="0,15,0,0"
                            HorizontalAlignment="Left"
                            Orientation="Horizontal">
                            <Button
                                x:Name="btnOpenLog"
                                Content="Log-Datei öffnen"
                                Style="{StaticResource ModernButton}" />
                            <Button
                                x:Name="btnClearLog"
                                Margin="10,0,0,0"
                                Background="#DC3545"
                                Content="Log leeren"
                                Style="{StaticResource ModernButton}" />
                            <Button
                                x:Name="btnExportDebug"
                                Margin="10,0,0,0"
                                Background="#6F42C1"
                                Content="Log exportieren"
                                Style="{StaticResource ModernButton}" />
                        </StackPanel>
                    </Grid>

                    <!--  Tools Panel  -->
                    <Grid Name="toolsPanel" Visibility="Collapsed">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="*" />
                        </Grid.RowDefinitions>

                        <TextBlock
                            Grid.Row="0"
                            Margin="0,0,0,20"
                            FontSize="24"
                            FontWeight="SemiBold"
                            Foreground="#1C1C1C"
                            Text="easyIT Tools" />

                        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                            <WrapPanel Orientation="Horizontal">
                                <!--  Tool Cards  -->
                                <Border
                                    Width="300"
                                    Height="200"
                                    Style="{StaticResource Card}">
                                    <StackPanel>
                                        <TextBlock
                                            Margin="0,0,0,10"
                                            FontSize="18"
                                            FontWeight="Bold"
                                            Text="easyADReport" />
                                        <TextBlock
                                            Height="80"
                                            Margin="0,0,0,15"
                                            Text="Active Directory Reporting Tool"
                                            TextWrapping="Wrap" />
                                        <Button
                                            x:Name="btnLaunchADReport"
                                            Content="Tool starten"
                                            Style="{StaticResource ModernButton}" />
                                    </StackPanel>
                                </Border>
                                <Border
                                    Width="300"
                                    Height="200"
                                    Style="{StaticResource Card}">
                                    <StackPanel>
                                        <TextBlock
                                            Margin="0,0,0,10"
                                            FontSize="18"
                                            FontWeight="Bold"
                                            Text="easyExchange" />
                                        <TextBlock
                                            Height="80"
                                            Margin="0,0,0,15"
                                            Text="Exchange Server Management und Reporting"
                                            TextWrapping="Wrap" />
                                        <Button
                                            x:Name="btnLaunchExchange"
                                            Content="Tool starten"
                                            Style="{StaticResource ModernButton}" />
                                    </StackPanel>
                                </Border>
                                <Border
                                    Width="300"
                                    Height="200"
                                    Style="{StaticResource Card}">
                                    <StackPanel>
                                        <TextBlock
                                            Margin="0,0,0,10"
                                            FontSize="18"
                                            FontWeight="Bold"
                                            Text="easyNetworkScan" />
                                        <TextBlock
                                            Height="80"
                                            Margin="0,0,0,15"
                                            Text="Netzwerk-Scan und Inventarisierung"
                                            TextWrapping="Wrap" />
                                        <Button
                                            x:Name="btnLaunchNetScan"
                                            Content="Tool starten"
                                            Style="{StaticResource ModernButton}" />
                                    </StackPanel>
                                </Border>
                                <Border
                                    Width="300"
                                    Height="200"
                                    Style="{StaticResource Card}">
                                    <StackPanel>
                                        <TextBlock
                                            Margin="0,0,0,10"
                                            FontSize="18"
                                            FontWeight="Bold"
                                            Text="easySQLAudit" />
                                        <TextBlock
                                            Height="80"
                                            Margin="0,0,0,15"
                                            Text="SQL Server Audit und Performance Analyse"
                                            TextWrapping="Wrap" />
                                        <Button
                                            x:Name="btnLaunchSQLAudit"
                                            Content="Tool starten"
                                            Style="{StaticResource ModernButton}" />
                                    </StackPanel>
                                </Border>
                                <Border
                                    Width="300"
                                    Height="200"
                                    Style="{StaticResource Card}">
                                    <StackPanel>
                                        <TextBlock
                                            Margin="0,0,0,10"
                                            FontSize="18"
                                            FontWeight="Bold"
                                            Text="easyPermissionReport" />
                                        <TextBlock
                                            Height="80"
                                            Margin="0,0,0,15"
                                            Text="Berechtigungsanalyse für Dateien und Ordner"
                                            TextWrapping="Wrap" />
                                        <Button
                                            x:Name="btnLaunchPermissionReport"
                                            Content="Tool starten"
                                            Style="{StaticResource ModernButton}" />
                                    </StackPanel>
                                </Border>
                                <Border
                                    Width="300"
                                    Height="200"
                                    Style="{StaticResource Card}">
                                    <StackPanel>
                                        <TextBlock
                                            Margin="0,0,0,10"
                                            FontSize="18"
                                            FontWeight="Bold"
                                            Text="easyFileSearch" />
                                        <TextBlock
                                            Height="80"
                                            Margin="0,0,0,15"
                                            Text="Erweiterte Dateisuche mit Filteroptionen"
                                            TextWrapping="Wrap" />
                                        <Button
                                            x:Name="btnLaunchFileSearch"
                                            Content="Tool starten"
                                            Style="{StaticResource ModernButton}" />
                                    </StackPanel>
                                </Border>
                            </WrapPanel>
                        </ScrollViewer>
                    </Grid>
                </Grid>
            </Border>
        </Grid>
        
        <!--  Footer  -->
        <Border
            Grid.Row="2"
            Background="#FF1C323C"
            BorderBrush="#E0E0E0"
            BorderThickness="0,1,0,0">
            <Grid Margin="20,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>
                <StackPanel
                    Grid.Column="2"
                    HorizontalAlignment="Right"
                    VerticalAlignment="Center"
                    Orientation="Horizontal">
                    <TextBlock
                        Margin="0,0,35,0"
                        FontSize="11"
                        Foreground="#e2e2e2"
                        Text="Copyright 2025  @  by PhinIT | PSscripts.de" />
                    <TextBlock
                        Name="linkWebsite"
                        Cursor="Hand"
                        FontSize="11"
                        Foreground="#d0e8ff"
                        Text="www.phinit.de  or  www.psscripts.de" />
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Lade das XAML
Write-DebugLog "Lade XAML fuer UI..." "UI"
try {
    # Stelle sicher, dass wir im STA-Thread sind
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
        Write-DebugLog "WARNUNG: Nicht im STA-Thread - UI könnte instabil sein" "UI"
    }
    
    # Validiere XAML vor dem Laden
    Write-DebugLog "Validiere XAML-Struktur..." "UI"
    if ($null -eq $xaml) {
        throw "XAML-Variable ist null"
    }
    
    Write-DebugLog "Erstelle XmlNodeReader..." "UI"
    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    
    if ($null -eq $reader) {
        throw "XmlNodeReader konnte nicht erstellt werden"
    }
    
    Write-DebugLog "Lade XAML mit XamlReader..." "UI"
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    if ($null -eq $window) {
        throw "Window konnte nicht erstellt werden - XamlReader.Load gibt null zurück"
    }
    
    # Zusätzliche Window-Validierung
    if (-not ($window -is [System.Windows.Window])) {
        throw "Geladenes Objekt ist kein gültiges Window-Objekt: $($window.GetType().FullName)"
    }
    
    Write-DebugLog "XAML erfolgreich geladen, Fenster erstellt. Type: $($window.GetType().FullName)" "UI"
    
} catch {
    Write-DebugLog "FEHLER beim Laden des XAML: $($_.Exception.Message)" "UI"
    Write-Host "KRITISCHER FEHLER beim Laden der UI: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Details: $($_.Exception.ToString())" -ForegroundColor Red
    
    # Detaillierte Fehleranalyse
    if ($_.Exception.InnerException) {
        Write-Host "Innere Exception: $($_.Exception.InnerException.Message)" -ForegroundColor Red
        Write-DebugLog "Innere Exception: $($_.Exception.InnerException.ToString())" "UI"
    }
    
    # XAML-Debugging-Informationen
    if ($null -ne $xaml) {
        Write-DebugLog "XAML-Type: $($xaml.GetType().FullName)" "UI"
        Write-DebugLog "XAML Root-Element: $($xaml.DocumentElement.Name)" "UI"
    } else {
        Write-DebugLog "XAML-Variable ist null!" "UI"
    }
    
    Read-Host "Drücken Sie eine Taste zum Beenden"
    return
}

# Hole die UI-Elemente
Write-DebugLog "Suche UI-Elemente..." "UI"

try {
    # Validiere Window vor FindName-Aufrufen
    if ($null -eq $window) {
        throw "Window-Objekt ist null - kann UI-Elemente nicht finden"
    }
    
    # Hauptelemente suchen
    Write-DebugLog "Suche Hauptelemente..." "UI"
    $txtServerName = $window.FindName("txtServerName")
    $btnFullAudit = $window.FindName("btnFullAudit")
    $spOptions = $window.FindName("spOptions")
    $btnSelectAll = $window.FindName("btnSelectAll")
    $btnSelectNone = $window.FindName("btnSelectNone")
    $btnRunAudit = $window.FindName("btnRunAudit")
    $progressBar = $window.FindName("progressBar")
    $txtProgress = $window.FindName("txtProgress")
    $txtStatusLog = $window.FindName("txtStatusLog")
    $cmbResultCategories = $window.FindName("cmbResultCategories")
    $rtbResults = $window.FindName("rtbResults")
    $txtResultsSummary = $window.FindName("txtResultsSummary")
    $btnExportHTML = $window.FindName("btnExportHTML")
    $btnCopyToClipboard = $window.FindName("btnCopyToClipboard")
    $btnRefreshResults = $window.FindName("btnRefreshResults")
    $txtStatus = $window.FindName("txtStatus")
    $txtStatusHeader = $window.FindName("txtStatusHeader")

    # Navigation-Elemente
    Write-DebugLog "Suche Navigation-Elemente..." "UI"
    $btnNavServerAudit = $window.FindName("btnNavServerAudit")
    $btnNavServerResults = $window.FindName("btnNavServerResults") 
    $btnNavNetworkAudit = $window.FindName("btnNavNetworkAudit")
    $btnNavNetworkResults = $window.FindName("btnNavNetworkResults")
    $btnNavTools = $window.FindName("btnNavTools")
    $btnNavDebug = $window.FindName("btnNavDebug")

    # Panel-Elemente
    Write-DebugLog "Suche Panel-Elemente..." "UI"
    $serverAuditPanel = $window.FindName("serverAuditPanel")
    $serverResultsPanel = $window.FindName("serverResultsPanel")
    $networkAuditPanel = $window.FindName("networkAuditPanel")
    $networkResultsPanel = $window.FindName("networkResultsPanel")
    $toolsPanel = $window.FindName("toolsPanel")
    $debugPanel = $window.FindName("debugPanel")

    # Verbindungsaudit-Elemente
    Write-DebugLog "Suche Verbindungsaudit-Elemente..." "UI"
    $spConnectionOptions = $window.FindName("spConnectionOptions")
    $btnSelectAllConnection = $window.FindName("btnSelectAllConnection")
    $btnSelectNoneConnection = $window.FindName("btnSelectNoneConnection")
    $btnRunConnectionAudit = $window.FindName("btnRunConnectionAudit")
    $progressBarConnection = $window.FindName("progressBarConnection")
    $txtProgressConnection = $window.FindName("txtProgressConnection")
    $btnExportConnectionHTML = $window.FindName("btnExportConnectionHTML")
    $btnExportConnectionDrawIO = $window.FindName("btnExportConnectionDrawIO")
    $btnCopyConnectionToClipboard = $window.FindName("btnCopyConnectionToClipboard")
    $cmbConnectionCategories = $window.FindName("cmbConnectionCategories")
    $rtbConnectionResults = $window.FindName("rtbConnectionResults")

    # Debug-Elemente
    Write-DebugLog "Suche Debug-Elemente..." "UI"
    $script:txtDebugOutput = $window.FindName("txtDebugOutput")
    $btnOpenLog = $window.FindName("btnOpenLog")
    $btnClearLog = $window.FindName("btnClearLog")

    Write-DebugLog "UI-Element-Suche abgeschlossen" "UI"

} catch {
    Write-DebugLog "FEHLER beim Suchen der UI-Elemente: $($_.Exception.Message)" "UI"
    Write-Host "KRITISCHER FEHLER bei der UI-Initialisierung: $($_.Exception.Message)" -ForegroundColor Red
    
    if ($_.Exception.InnerException) {
        Write-DebugLog "Innere Exception: $($_.Exception.InnerException.ToString())" "UI"
    }
    
    Read-Host "Drücken Sie eine Taste zum Beenden"
    return
}

# === HILFSFUNKTIONEN ===

# Funktion zum Aktualisieren des Status-Texts
function Update-StatusText {
    param([string]$StatusMessage)
    
    Write-DebugLog "Status-Update: $StatusMessage" "Status"
    
    try {
        if ($txtStatus) {
            $window.Dispatcher.Invoke([Action]{
                $txtStatus.Text = $StatusMessage
            }, "Normal")
        }
    }
    catch {
        Write-DebugLog "FEHLER beim Aktualisieren des Status-Texts: $($_.Exception.Message)" "Status"
        # Fallback: Direkter Zugriff
        try {
            $txtStatus.Text = $StatusMessage
        }
        catch {
            Write-DebugLog "FEHLER beim direkten Zugriff auf txtStatus: $($_.Exception.Message)" "Status"
        }
    }
}

# Überprüfe kritische UI-Elemente
Write-DebugLog "Überprüfe UI-Elemente..." "UI"

# Validiere kritische Elemente
$criticalElements = @{
    "window" = $window
    "spOptions" = $spOptions
    "txtServerName" = $txtServerName
    "spConnectionOptions" = $spConnectionOptions
}

foreach ($elementName in $criticalElements.Keys) {
    $element = $criticalElements[$elementName]
    if ($null -eq $element) { 
        Write-DebugLog "FEHLER: $elementName ist NULL!" "UI"
        Write-Host "KRITISCHER FEHLER: UI-Element '$elementName' konnte nicht gefunden werden!" -ForegroundColor Red
        Read-Host "Drücken Sie eine Taste zum Beenden"
        return
    } else {
        Write-DebugLog "✓ $elementName gefunden (Typ: $($element.GetType().Name))" "UI"
    }
}

Write-DebugLog "UI-Elemente erfolgreich initialisiert" "UI"

# Servername anzeigen
try {
    $txtServerName.Text = "Server: $env:COMPUTERNAME"
    Write-DebugLog "Servername gesetzt: $env:COMPUTERNAME" "UI"
} catch {
    Write-DebugLog "FEHLER beim Setzen des Servernamens: $($_.Exception.Message)" "UI"
}

# Dictionary fuer Checkboxen (beide Audits)
$checkboxes = @{}
$connectionCheckboxes = @{}

# Erstelle die Checkboxen fuer die Verbindungsaudit-Optionen
Write-DebugLog "Erstelle Checkboxen fuer Verbindungsaudit-Optionen..." "UI"

$connectionCategories = @{}
foreach ($cmd in $connectionAuditCommands) {
    $categoryName = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
    if (-not $connectionCategories.ContainsKey($categoryName)) {
        $connectionCategories[$categoryName] = @()
    }
    $connectionCategories[$categoryName] += $cmd
}

# Iteriere über die Verbindungsaudit-Kategorien in alphabetischer Reihenfolge
foreach ($categoryKey in ($connectionCategories.Keys | Sort-Object)) {
    # Kategorie-Header
    $categoryHeader = New-Object System.Windows.Controls.TextBlock
    $categoryHeader.Text = "$categoryKey"
    
    # Setze Style direkt
    $categoryHeader.FontSize = 16
    $categoryHeader.FontWeight = "Bold"
    $categoryHeader.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(253, 126, 20)) # Orange für Verbindungsaudit
    $categoryHeader.Margin = New-Object System.Windows.Thickness(0, 15, 0, 5)
    
    try {
        $spConnectionOptions.Children.Add($categoryHeader)
        Write-DebugLog "Verbindungsaudit Kategorie-Header '$categoryKey' hinzugefügt" "UI"
    } catch {
        Write-DebugLog "FEHLER beim Hinzufügen des Verbindungsaudit Kategorie-Headers '$categoryKey': $($_.Exception.Message)" "UI"
        continue
    }
    
    # Checkboxen fuer diese Kategorie
    foreach ($cmd in $connectionCategories[$categoryKey]) {
        $checkbox = New-Object System.Windows.Controls.CheckBox
        $checkbox.Content = $cmd.Name
        $checkbox.IsChecked = $true # Standardmäßig aktiviert
        
        # Setze Style direkt
        $checkbox.Margin = New-Object System.Windows.Thickness(20, 3, 5, 3)
        $checkbox.Padding = New-Object System.Windows.Thickness(5, 0, 0, 0)
        
        # Ueberpruefe, ob diese Option mit einer Serverrolle verbunden ist
        if ($cmd.ContainsKey("FeatureName")) {
            $isRoleInstalled = Test-ServerRole -FeatureName $cmd.FeatureName
            if (-not $isRoleInstalled) {
                $checkbox.IsEnabled = $false
                $checkbox.Content = "$($cmd.Name) (Nicht installiert)"
                $checkbox.IsChecked = $false
            }
        }
        
        try {
            $spConnectionOptions.Children.Add($checkbox)
            $connectionCheckboxes[$cmd.Name] = $checkbox
            Write-DebugLog "Verbindungsaudit Checkbox '$($cmd.Name)' hinzugefügt (Kategorie: $categoryKey)" "UI"
        } catch {
            Write-DebugLog "FEHLER beim Hinzufügen der Verbindungsaudit Checkbox '$($cmd.Name)' (Kategorie: $categoryKey): $($_.Exception.Message)" "UI"
        }
    }
}
Write-DebugLog "Verbindungsaudit Checkboxen erstellt für $($connectionCheckboxes.Count) Optionen" "UI"

# Dictionary fuer Checkboxen
$checkboxes = @{}

# Erstelle die Checkboxen fuer die Audit-Optionen gruppiert nach Kategorien
Write-DebugLog "Erstelle Checkboxen fuer Audit-Optionen..." "UI"

# Überprüfe, ob spOptions verfügbar ist
if ($null -eq $spOptions) {
    Write-DebugLog "FEHLER: spOptions nicht gefunden - kann keine Checkboxen erstellen!" "UI"
    [System.Windows.MessageBox]::Show("UI-Initialisierungsfehler: spOptions nicht gefunden.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    return
}

$categories = @{} # Verwende ein normales Hashtable für bessere Kompatibilität
foreach ($cmd in $commands) {
    $categoryName = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
    if (-not $categories.ContainsKey($categoryName)) {
        # Initialisiere den Eintrag für eine neue Kategorie mit einer Liste.
        $categories[$categoryName] = @()
    }
    # Füge den Befehl zur Liste der entsprechenden Kategorie hinzu
    $categories[$categoryName] += $cmd
}

# Iteriere über die Kategorien in alphabetischer Reihenfolge
foreach ($categoryKey in ($categories.Keys | Sort-Object)) {
    # Kategorie-Header
    $categoryHeader = New-Object System.Windows.Controls.TextBlock
    $categoryHeader.Text = "$categoryKey"
    
    # Setze Style direkt statt über FindResource
    $categoryHeader.FontSize = 16
    $categoryHeader.FontWeight = "Bold"
    $categoryHeader.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 120, 212))
    $categoryHeader.Margin = New-Object System.Windows.Thickness(0, 15, 0, 5)
    
    try {
        $spOptions.Children.Add($categoryHeader)
        Write-DebugLog "Kategorie-Header '$categoryKey' hinzugefügt" "UI"
    } catch {
        Write-DebugLog "FEHLER beim Hinzufügen des Kategorie-Headers '$categoryKey': $($_.Exception.Message)" "UI"
        continue # Springe zur nächsten Kategorie, falls der Header nicht hinzugefügt werden kann
    }
    
    # Checkboxen fuer diese Kategorie
    # $categories[$categoryKey] enthält eine Liste von Befehls-Hashtables für die aktuelle Kategorie
    foreach ($cmd in $categories[$categoryKey]) {
        $checkbox = New-Object System.Windows.Controls.CheckBox
        $checkbox.Content = $cmd.Name
        $checkbox.IsChecked = $true # Standardmäßig aktiviert
        
        # Setze Style direkt statt über FindResource
        $checkbox.Margin = New-Object System.Windows.Thickness(20, 3, 5, 3)
        $checkbox.Padding = New-Object System.Windows.Thickness(5, 0, 0, 0)
        
        # Ueberpruefe, ob diese Option mit einer Serverrolle verbunden ist
        if ($cmd.ContainsKey("FeatureName")) {
            # Test-ServerRole ist eine Funktion, die prüft, ob eine Windows-Funktion/Rolle installiert ist.
            # Diese Funktion ist außerhalb dieses Codeblocks definiert.
            $isRoleInstalled = Test-ServerRole -FeatureName $cmd.FeatureName
            if (-not $isRoleInstalled) {
                $checkbox.IsEnabled = $false
                $checkbox.Content = "$($cmd.Name) (Nicht installiert)"
                $checkbox.IsChecked = $false # Deaktiviere Checkbox, wenn zugehörige Rolle nicht installiert ist
            }
        }
        
        try {
            $spOptions.Children.Add($checkbox)
            $checkboxes[$cmd.Name] = $checkbox # Speichere eine Referenz zur Checkbox im globalen Hashtable
            Write-DebugLog "Checkbox '$($cmd.Name)' hinzugefügt (Kategorie: $categoryKey)" "UI"
        } catch {
            Write-DebugLog "FEHLER beim Hinzufügen der Checkbox '$($cmd.Name)' (Kategorie: $categoryKey): $($_.Exception.Message)" "UI"
            # Fahre mit der nächsten Checkbox fort, auch wenn eine fehlschlägt
        }
    }
}
Write-DebugLog "Checkboxen erstellt fuer $($checkboxes.Count) Optionen" "UI"

# Button-Event-Handler

# "Alle auswaehlen" Button
$btnSelectAll.Add_Click({
    Write-DebugLog "Alle Optionen auswaehlen" "UI"
    foreach ($key in $checkboxes.Keys) {
        if ($checkboxes[$key].IsEnabled) {
            $checkboxes[$key].IsChecked = $true
        }
    }
})

# "Alle abwaehlen" Button
$btnSelectNone.Add_Click({
    Write-DebugLog "Alle Optionen abwaehlen" "UI"
    foreach ($key in $checkboxes.Keys) {
        $checkboxes[$key].IsChecked = $false
    }
})

# Verbindungsaudit Button-Event-Handler

# "Alle auswaehlen" Button (Verbindungsaudit)
$btnSelectAllConnection.Add_Click({
    Write-DebugLog "Alle Verbindungsaudit-Optionen auswaehlen" "UI"
    foreach ($key in $connectionCheckboxes.Keys) {
        if ($connectionCheckboxes[$key].IsEnabled) {
            $connectionCheckboxes[$key].IsChecked = $true
        }
    }
})

# "Alle abwaehlen" Button (Verbindungsaudit)
$btnSelectNoneConnection.Add_Click({
    Write-DebugLog "Alle Verbindungsaudit-Optionen abwaehlen" "UI"
    foreach ($key in $connectionCheckboxes.Keys) {
        $connectionCheckboxes[$key].IsChecked = $false
    }
})

# "Verbindungsaudit starten" Button
$btnRunConnectionAudit.Add_Click({
    Write-DebugLog "Verbindungsaudit gestartet" "ConnectionAudit"
    
    # UI vorbereiten
    $btnRunConnectionAudit.IsEnabled = $false
    $btnExportConnectionHTML.IsEnabled = $false
    $btnExportConnectionDrawIO.IsEnabled = $false
    $btnCopyConnectionToClipboard.IsEnabled = $false
    $cmbConnectionCategories.IsEnabled = $false
    
    $rtbConnectionResults.Document = New-Object System.Windows.Documents.FlowDocument
    
    # Hole die UI-Elemente für das Netzwerk-Log
    $txtNetworkStatusLog = $window.FindName("txtNetworkStatusLog")
    $scrollNetworkStatusLog = $window.FindName("scrollNetworkStatusLog")

    # Leere das Log-Fenster sicher im UI-Thread
    $window.Dispatcher.Invoke([Action]{
        if ($txtNetworkStatusLog) {
            $txtNetworkStatusLog.Text = ""
        }
    })
    
    $progressBarConnection.Value = 0

    # Hilfsfunktion zum sicheren Hinzufügen von Text zum Netzwerk-Log
    $addNetworkLog = { param($text)
        $window.Dispatcher.Invoke([Action]{
            if ($txtNetworkStatusLog) {
                $txtNetworkStatusLog.Text += $text
                if ($scrollNetworkStatusLog) {
                    $scrollNetworkStatusLog.ScrollToEnd()
                }
            }
        })
    }
    
    # UI initial aktualisieren
    $window.Dispatcher.Invoke([Action]{
        $txtProgressConnection.Text = "Initialisiere Verbindungsaudit..."
        $progressBarConnection.Value = 0
    }, "Normal")
    $addNetworkLog.Invoke("=== Verbindungsaudit gestartet ===`r`n")
    
    # UI refresh erzwingen
    $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
    Start-Sleep -Milliseconds 300
    
    # Sammle ausgewählte Befehle
    $selectedConnectionCommands = @()
    foreach ($cmd in $connectionAuditCommands) {
        if ($connectionCheckboxes[$cmd.Name].IsChecked) {
            $selectedConnectionCommands += $cmd
        }
    }
    
    if ($selectedConnectionCommands.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie mindestens eine Verbindungsaudit-Option aus.", "Keine Auswahl", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        $btnRunConnectionAudit.IsEnabled = $true
        return
    }
    
    $global:connectionAuditResults = @{}
    $progressStep = 100.0 / $selectedConnectionCommands.Count
    $currentProgress = 0
    
    # UI Update mit Anzahl der Befehle
    $window.Dispatcher.Invoke([Action]{
        $txtProgressConnection.Text = "Bereite $($selectedConnectionCommands.Count) Verbindungsaudit-Befehle vor..."
    }, "Normal")
    $addNetworkLog.Invoke("Anzahl ausgewaehlter Befehle: $($selectedConnectionCommands.Count)`r`n`r`n")
    
    # UI refresh erzwingen
    $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
    Start-Sleep -Milliseconds 500
    
    for ($i = 0; $i -lt $selectedConnectionCommands.Count; $i++) {
        $cmd = $selectedConnectionCommands[$i]
        
        # UI aktualisieren - BEGINN des Befehls
        $window.Dispatcher.Invoke([Action]{
            $txtProgressConnection.Text = "Verarbeite: $($cmd.Name) ($($i+1)/$($selectedConnectionCommands.Count))"
            $progressBarConnection.Value = $currentProgress
        }, "Normal")
        $addNetworkLog.Invoke("[$($i+1)/$($selectedConnectionCommands.Count)] $($cmd.Name)...`r`n")
        
        # UI refresh erzwingen
        $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
        Invoke-SafeDoEvents
        Start-Sleep -Milliseconds 200
        
        Write-DebugLog "Fuehre Verbindungsaudit aus ($($i+1)/$($selectedConnectionCommands.Count)): $($cmd.Name)" "ConnectionAudit"
        
        try {
            if ($cmd.Type -eq "PowerShell") {
                $result = Invoke-PSCommand -Command $cmd.Command
            } else {
                $result = Invoke-CMDCommand -Command $cmd.Command
            }
            
            $global:connectionAuditResults[$cmd.Name] = $result
            $currentProgress += $progressStep
            
            $window.Dispatcher.Invoke([Action]{
                $progressBarConnection.Value = $currentProgress
                $txtProgressConnection.Text = "Abgeschlossen: $($cmd.Name) ($($i+1)/$($selectedConnectionCommands.Count))"
            }, "Normal")
            $addNetworkLog.Invoke("  [OK] Erfolgreich abgeschlossen`r`n")
            
        } catch {
            $errorMsg = "Fehler: $($_.Exception.Message)"
            $global:connectionAuditResults[$cmd.Name] = $errorMsg
            $currentProgress += $progressStep
            
            $window.Dispatcher.Invoke([Action]{
                $progressBarConnection.Value = $currentProgress
                $txtProgressConnection.Text = "Fehler bei: $($cmd.Name) ($($i+1)/$($selectedConnectionCommands.Count))"
            }, "Normal")
            $addNetworkLog.Invoke("  [FEHLER] $($_.Exception.Message)`r`n")
            
            Write-DebugLog "FEHLER bei Verbindungsaudit $($cmd.Name): $($_.Exception.Message)" "ConnectionAudit"
        }
        
        # UI refresh erzwingen
        $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
        Invoke-SafeDoEvents
        Start-Sleep -Milliseconds 300
    }
    
    # Verbindungsaudit abgeschlossen
    $window.Dispatcher.Invoke([Action]{
        $progressBarConnection.Value = 100
        $txtProgressConnection.Text = "Verbindungsaudit abgeschlossen! $($selectedConnectionCommands.Count) Befehle ausgefuehrt."
        
        # Zeige Ergebnisse an
        try {
            Update-ConnectionResultsCategories
            Show-ConnectionCategoryResults -Category "Alle"
        } catch {
            Write-DebugLog "FEHLER beim Anzeigen der Verbindungsaudit-Ergebnisse: $($_.Exception.Message)" "ConnectionAudit"
            Show-SimpleConnectionResults -Category "Alle"
        }
        
        # Automatisch zu den Netzwerk-Ergebnissen wechseln
        Switch-Panel "networkResults"
        
        # Buttons wieder aktivieren
        $btnRunConnectionAudit.IsEnabled = $true
        $btnExportConnectionHTML.IsEnabled = $true
        $btnExportConnectionDrawIO.IsEnabled = $true
        $btnCopyConnectionToClipboard.IsEnabled = $true
        $cmbConnectionCategories.IsEnabled = $true
        
    }, "Normal")
    
    $addNetworkLog.Invoke("`r`n" + "="*50 + "`r`n")
    $addNetworkLog.Invoke("[FERTIG] Verbindungsaudit erfolgreich abgeschlossen!`r`n")
    $addNetworkLog.Invoke("Ergebnisse: $($global:connectionAuditResults.Count) Eintraege`r`n")
    $addNetworkLog.Invoke("="*50 + "`r`n")
    
    Write-DebugLog "Verbindungsaudit abgeschlossen mit $($global:connectionAuditResults.Count) Ergebnissen" "ConnectionAudit"
})

# ComboBox Selection Changed Event
$cmbResultCategories.Add_SelectionChanged({
    if ($cmbResultCategories.SelectedItem) {
        $selectedCategory = $cmbResultCategories.SelectedItem.Tag
        Show-CategoryResults -Category $selectedCategory
    }
})

# "Ergebnisse aktualisieren" Button
$btnRefreshResults.Add_Click({
    Write-DebugLog "Aktualisiere Ergebnisse-Anzeige" "UI"
    
    try {
        # Versuche die RichTextBox zurückzusetzen
        Reset-ResultsDisplay
        
        # Aktualisiere Kategorien
        Update-ResultsCategories
        
        # Zeige die ausgewählte Kategorie erneut an
        if ($cmbResultCategories.SelectedItem) {
            $selectedCategory = $cmbResultCategories.SelectedItem.Tag
            Show-CategoryResults -Category $selectedCategory
        } else {
            Show-CategoryResults -Category "Alle"
        }
        
        Update-StatusText "Status: Ergebnisse erfolgreich aktualisiert"
    }
    catch {
        Write-DebugLog "FEHLER beim Aktualisieren der Ergebnisse: $($_.Exception.Message)" "UI"
        Update-StatusText "Status: Fehler beim Aktualisieren der Ergebnisse"
        [System.Windows.MessageBox]::Show("Fehler beim Aktualisieren der Ergebnisse:`r`n$($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        
        # Fallback: Zeige Ergebnisse in einfacher Form
        Show-SimpleResults -Category "Alle"
    }
})

# Funktion zum Zurücksetzen der Ergebnisanzeige
function Reset-ResultsDisplay {
    Write-DebugLog "Setze Ergebnisanzeige zurück" "UI"
    
    try {
        # Erstelle ein komplett neues, leeres FlowDocument
        $newDocument = New-Object System.Windows.Documents.FlowDocument
        $newDocument.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
        $newDocument.FontSize = 12
        $newDocument.PageWidth = [Double]::NaN
        $newDocument.PageHeight = [Double]::NaN
        $newDocument.ColumnWidth = [Double]::PositiveInfinity
        
        # Setze das neue Dokument
        $rtbResults.Document = $newDocument
        
        # Erzwinge Update
        $rtbResults.UpdateLayout()
        
        Write-DebugLog "Ergebnisanzeige erfolgreich zurückgesetzt" "UI"
    }
    catch {
        Write-DebugLog "FEHLER beim Zurücksetzen der Ergebnisanzeige: $($_.Exception.Message)" "UI"
        throw $_
    }
}

# Hilfsfunktion zum Optimieren der RichTextBox für besseres Text-Wrapping
function Optimize-RichTextBoxLayout {
    try {
        Write-DebugLog "Optimiere RichTextBox-Layout" "UI"
        
        # Setze die RichTextBox-Eigenschaften für optimales Wrapping
        $rtbResults.HorizontalScrollBarVisibility = "Disabled"
        $rtbResults.VerticalScrollBarVisibility = "Disabled" # ScrollViewer übernimmt das Scrolling
        
        # Stelle sicher, dass das Document korrekt konfiguriert ist
        if ($rtbResults.Document) {
            $rtbResults.Document.PageWidth = [Double]::NaN
            $rtbResults.Document.PageHeight = [Double]::NaN
            $rtbResults.Document.ColumnWidth = [Double]::PositiveInfinity
            $rtbResults.Document.TextAlignment = "Left"
        }
        
        Write-DebugLog "RichTextBox-Layout optimiert" "UI"
    }
    catch {
        Write-DebugLog "WARNUNG: Konnte RichTextBox-Layout nicht optimieren: $($_.Exception.Message)" "UI"
        # Fehler ignorieren, da dies nur eine Optimierung ist
    }
}

# Vollstaendiges Audit Button
$btnFullAudit.Add_Click({
    Write-DebugLog "Vollstaendiges Audit gestartet" "Audit"
    # Alle verfuegbaren Optionen auswaehlen
    foreach ($key in $checkboxes.Keys) {
        if ($checkboxes[$key].IsEnabled) {
            $checkboxes[$key].IsChecked = $true
        }
    }
    # Audit starten
    Start-AuditProcess
})

# "Audit starten" Button
$btnRunAudit.Add_Click({
    Write-DebugLog "Benutzerdefiniertes Audit gestartet" "Audit"
    Start-AuditProcess
})

# Hauptfunktion fuer die Audit-Durchfuehrung (Synchron)
function Start-AuditProcess {
    # UI vorbereiten
    $btnRunAudit.IsEnabled = $false
    $btnFullAudit.IsEnabled = $false
    $btnExportHTML.IsEnabled = $false
    $btnCopyToClipboard.IsEnabled = $false
    $btnRefreshResults.IsEnabled = $false
    
    $rtbResults.Document = New-Object System.Windows.Documents.FlowDocument
    $txtStatusLog.Text = ""
    $progressBar.Value = 0
    Update-StatusText "Status: Audit laeuft..."
    
    # UI initial aktualisieren
    $window.Dispatcher.Invoke([Action]{
        $txtProgress.Text = "Initialisiere Audit..."
        $txtStatusLog.Text = "=== Audit gestartet ===`r`n"
        $progressBar.Value = 0
    }, "Normal")
    
    # UI refresh erzwingen
    $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
    Start-Sleep -Milliseconds 300
    
    # Sammle ausgewaehlte Befehle
    $selectedCommands = @()
    foreach ($cmd in $commands) {
        if ($checkboxes[$cmd.Name].IsChecked) {
            $selectedCommands += $cmd
        }
    }
    
    Write-DebugLog "Starte Audit mit $($selectedCommands.Count) ausgewaehlten Befehlen" "Audit"
    
    $global:auditResults = @{}
    $allResults = ""
    $progressStep = 100.0 / $selectedCommands.Count
    $currentProgress = 0
    
    # UI Update mit Anzahl der Befehle
    $window.Dispatcher.Invoke([Action]{
        $txtProgress.Text = "Bereite $($selectedCommands.Count) Audit-Befehle vor..."
        $txtStatusLog.Text += "Anzahl ausgewaehlter Befehle: $($selectedCommands.Count)`r`n`r`n"
    }, "Normal")
    
    # UI refresh erzwingen
    $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
    Start-Sleep -Milliseconds 500
    
    for ($i = 0; $i -lt $selectedCommands.Count; $i++) {
        $cmd = $selectedCommands[$i]
        
        # UI aktualisieren - BEGINN des Befehls
        $window.Dispatcher.Invoke([Action]{
            $txtProgress.Text = "Verarbeite: $($cmd.Name) ($($i+1)/$($selectedCommands.Count))"
            $txtStatusLog.Text += "[$($i+1)/$($selectedCommands.Count)] $($cmd.Name)...`r`n"
            # Fortschritt am Anfang des Befehls anzeigen
            $progressBar.Value = $currentProgress
        }, "Normal")
        
        # UI refresh erzwingen - das ist der Schluessel!
        $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
        Invoke-SafeDoEvents
        Start-Sleep -Milliseconds 200
        
        Write-DebugLog "Fuehre aus ($($i+1)/$($selectedCommands.Count)): $($cmd.Name)" "Audit"
        
        try {
            if ($cmd.Type -eq "PowerShell") {
                $result = Invoke-PSCommand -Command $cmd.Command
            } else {
                $result = Invoke-CMDCommand -Command $cmd.Command
            }
            
            $global:auditResults[$cmd.Name] = $result
            $allResults += "`r`n=== $($cmd.Name) ===`r`n$result`r`n"
            
            # Erfolg in Status-Log UND Fortschrittsbalken aktualisieren
            $currentProgress += $progressStep
            $window.Dispatcher.Invoke([Action]{
                $txtStatusLog.Text += "  [OK] Erfolgreich abgeschlossen`r`n"
                # Fortschrittsbalken NACH erfolgreichem Befehl aktualisieren
                $progressBar.Value = $currentProgress
                $txtProgress.Text = "Abgeschlossen: $($cmd.Name) ($($i+1)/$($selectedCommands.Count))"
            }, "Normal")
            
        } catch {
            $errorMsg = "Fehler: $($_.Exception.Message)"
            $global:auditResults[$cmd.Name] = $errorMsg
            $allResults += "`r`n=== $($cmd.Name) ===`r`n$errorMsg`r`n"
            
            # Fehler in Status-Log UND Fortschrittsbalken trotzdem aktualisieren
            $currentProgress += $progressStep
            $window.Dispatcher.Invoke([Action]{
                $txtStatusLog.Text += "  [FEHLER] $($_.Exception.Message)`r`n"
                # Fortschrittsbalken auch bei Fehler aktualisieren
                $progressBar.Value = $currentProgress
                $txtProgress.Text = "Fehler bei: $($cmd.Name) ($($i+1)/$($selectedCommands.Count))"
            }, "Normal")
            
            Write-DebugLog "FEHLER bei $($cmd.Name): $($_.Exception.Message)" "Audit"
        }
        
        # UI refresh nach jedem Befehl erzwingen - SEHR WICHTIG!
        $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
        Invoke-SafeDoEvents
        Start-Sleep -Milliseconds 300
        
        # Zwischenstand der Ergebnisse aktualisieren (optional)
        if (($i + 1) % 3 -eq 0 -or $i -eq ($selectedCommands.Count - 1)) {
            $window.Dispatcher.Invoke([Action]{
                try {
                    # Aktualisiere die schöne Anzeige statt der alten Textbox
                    Update-ResultsCategories
                    if ($cmbResultCategories.SelectedItem) {
                        $selectedCategory = $cmbResultCategories.SelectedItem.Tag
                        Show-CategoryResults -Category $selectedCategory
                    } else {
                        Show-CategoryResults -Category "Alle"
                    }
                }
                catch {
                    Write-DebugLog "FEHLER beim Zwischenupdate der Ergebnisanzeige: $($_.Exception.Message)" "Audit"
                    # Bei Fehlern trotzdem fortfahren
                }
            }, "Normal")
            # UI refresh auch hier
            $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
            Invoke-SafeDoEvents
        }
    }
    
    # Audit abgeschlossen - Finale Updates
    $window.Dispatcher.Invoke([Action]{
        $progressBar.Value = 100
        $txtProgress.Text = "Audit vollstaendig abgeschlossen! $($selectedCommands.Count) Befehle ausgefuehrt."
        $txtStatusLog.Text += "`r`n" + "="*50 + "`r`n"
        $txtStatusLog.Text += "[FERTIG] Audit erfolgreich abgeschlossen!`r`n"
        $txtStatusLog.Text += "Ergebnisse: $($global:auditResults.Count) Eintraege`r`n"
        $txtStatusLog.Text += "="*50 + "`r`n"
        
        try {
            Write-DebugLog "Starte finales Update der Ergebnisanzeige" "Audit"
            Write-DebugLog "Anzahl Ergebnisse vor Update: $($global:auditResults.Count)" "Audit"
            
            # Aktualisiere die schöne Kategorien-Anzeige
            Update-ResultsCategories
            
            # Kleine Pause für UI-Stabilität
            Start-Sleep -Milliseconds 500
            
            # Zeige Ergebnisse an
            Show-CategoryResults -Category "Alle"
            
            Write-DebugLog "Finales Update der Ergebnisanzeige erfolgreich" "Audit"
        }
        catch {
            Write-DebugLog "FEHLER beim finalen Update der Ergebnisanzeige: $($_.Exception.Message)" "Audit"
            Write-DebugLog "Fehlerstapel: $($_.Exception.StackTrace)" "Audit"
            
            # Fallback auf einfache Anzeige
            try {
                Write-DebugLog "Versuche Fallback: Show-SimpleResults" "Audit"
                Show-SimpleResults -Category "Alle"
                $txtStatusLog.Text += "[WARNUNG] Verwendet einfache Ergebnisanzeige aufgrund von Formatierungsproblemen`r`n"
                Write-DebugLog "Fallback erfolgreich" "Audit"
            }
            catch {
                Write-DebugLog "FEHLER auch bei einfacher Anzeige: $($_.Exception.Message)" "Audit"
                $txtStatusLog.Text += "[FEHLER] Konnte Ergebnisse nicht anzeigen - siehe Debug-Log`r`n"
                
                # Letzter Notfall: Zeige Debug-Informationen
                $txtStatusLog.Text += "Debug: $($global:auditResults.Count) Ergebnisse verfügbar`r`n"
                $txtStatusLog.Text += "Debug: Erste 3 Ergebnisse: $($global:auditResults.Keys | Select-Object -First 3 | ForEach-Object { $_ })`r`n"
            }
        }
        
        $txtStatus.Text = "Status: Audit abgeschlossen - $($global:auditResults.Count) Ergebnisse"
        
        # Automatisch zu den Ergebnissen wechseln
        Switch-Panel "serverResults"
        
        # Buttons wieder aktivieren
        $btnRunAudit.IsEnabled = $true
        $btnFullAudit.IsEnabled = $true
        $btnExportHTML.IsEnabled = $true
        $btnCopyToClipboard.IsEnabled = $true
        $btnRefreshResults.IsEnabled = $true
    }, "Normal")
    
    # Finaler UI refresh
    Invoke-SafeDispatcher
    Invoke-SafeDoEvents
    
    Write-DebugLog "Audit abgeschlossen mit $($global:auditResults.Count) Ergebnissen" "Audit"
}

# Export-Button-Funktionalitaet
$btnExportHTML.Add_Click({
    Write-DebugLog "HTML-Export gestartet" "Export"
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "HTML Files (*.html)|*.html"
    $saveFileDialog.Title = "Speichern Sie den Audit-Bericht"
    $saveFileDialog.FileName = "ServerAudit_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtStatus.Text = "Status: Exportiere HTML..."
        
        try {
            Export-AuditToHTML -Results $global:auditResults -FilePath $saveFileDialog.FileName
            $txtStatus.Text = "Status: Export erfolgreich abgeschlossen"
            [System.Windows.MessageBox]::Show("Bericht wurde erfolgreich exportiert:`r`n$($saveFileDialog.FileName)", "Export erfolgreich", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } catch {
            $txtStatus.Text = "Status: Fehler beim Export"
            [System.Windows.MessageBox]::Show("Fehler beim Export:`r`n$($_.Exception.Message)", "Export Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Zwischenablage-Button
$btnCopyToClipboard.Add_Click({
    Write-DebugLog "Kopiere Ergebnisse in Zwischenablage" "UI"
    try {
        # Extrahiere Text aus der RichTextBox
        $textRange = New-Object System.Windows.Documents.TextRange($rtbResults.Document.ContentStart, $rtbResults.Document.ContentEnd)
        $plainText = $textRange.Text
        
        if ([string]::IsNullOrWhiteSpace($plainText)) {
            # Fallback: Erstelle Text aus den Rohdaten
            $allResults = ""
            foreach ($key in $global:auditResults.Keys | Sort-Object) {
                $allResults += "`r`n=== $key ===`r`n$($global:auditResults[$key])`r`n"
            }
            $plainText = $allResults
        }
        
        $plainText | Set-Clipboard
        $txtStatus.Text = "Status: Ergebnisse in Zwischenablage kopiert"
        [System.Windows.MessageBox]::Show("Audit-Ergebnisse wurden in die Zwischenablage kopiert.", "Kopiert", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Kopieren in die Zwischenablage: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# Debug-Funktionen
$btnOpenLog.Add_Click({
    Write-DebugLog "Oeffne Log-Datei" "Debug"
    if (Test-Path $DebugLogPath) {
        Start-Process notepad.exe -ArgumentList $DebugLogPath
    } else {
        [System.Windows.MessageBox]::Show("Log-Datei nicht gefunden.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$btnClearLog.Add_Click({
    Write-DebugLog "Debug-Log wird geleert" "Debug"
    $script:txtDebugOutput.Text = ""
    if ($DEBUG) {
        $clearMessage = "=== Debug-Log geloescht: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ==="
        Set-Content -Path $DebugLogPath -Value $clearMessage -Force
        Write-Host $clearMessage -ForegroundColor Cyan
    }
})

Write-DebugLog "UI-Initialisierung abgeschlossen" "Init"

# Funktion zum Formatieren der RichTextBox (vereinfacht für Stabilität)
function Format-RichTextResults {
    param(
        [hashtable]$Results,
        [string]$CategoryFilter = "Alle"
    )
    
    Write-DebugLog "WARNUNG: Format-RichTextResults ist deprecated - verwende Show-CategoryResults stattdessen" "UI"
    
    # Leite an die sichere Show-CategoryResults Funktion weiter
    Show-CategoryResults -Category $CategoryFilter
}

# Funktion zum Aktualisieren der Kategorien-ComboBox
function Update-ResultsCategories {
    Write-DebugLog "Aktualisiere Kategorien-ComboBox" "UI"
    Write-DebugLog "Verfügbare Audit-Ergebnisse für Kategorisierung: $($global:auditResults.Count)" "UI"
    
    if (-not $cmbResultCategories) {
        Write-DebugLog "FEHLER: cmbResultCategories nicht gefunden!" "UI"
        return
    }
    
    $cmbResultCategories.Items.Clear()
    
    # "Alle" Option hinzufügen
    $allItem = New-Object System.Windows.Controls.ComboBoxItem
    $allItem.Content = "Alle Kategorien ($($global:auditResults.Count))" 
    $allItem.Tag = "Alle"
    $cmbResultCategories.Items.Add($allItem)
    
    # Einzelne Kategorien hinzufügen
    $categories = @{}
    if ($null -ne $commands) {
        Write-DebugLog "Verarbeite $($commands.Count) Commands für Kategorisierung" "UI"
        foreach ($cmd in $commands) {
            $category = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
            if (-not $categories.ContainsKey($category)) {
                $categories[$category] = 0
            }
            if ($global:auditResults.ContainsKey($cmd.Name)) {
                $categories[$category]++
            }
        }
        
        Write-DebugLog "Gefundene Kategorien: $($categories.Keys -join ', ')" "UI"
    } else {
        Write-DebugLog "WARNUNG: commands Variable ist NULL!" "UI"
    }
    
    foreach ($category in $categories.Keys | Sort-Object) {
        if ($categories[$category] -gt 0) {
            $categoryItem = New-Object System.Windows.Controls.ComboBoxItem
            $categoryItem.Content = "$category ($($categories[$category]))"
            $categoryItem.Tag = $category
            $cmbResultCategories.Items.Add($categoryItem)
            Write-DebugLog "Kategorie hinzugefügt: $category mit $($categories[$category]) Ergebnissen" "UI"
        }
    }
    
    # Ersten Eintrag auswählen
    if ($cmbResultCategories.Items.Count -gt 0) {
        $cmbResultCategories.SelectedIndex = 0
        Write-DebugLog "ComboBox initialisiert mit $($cmbResultCategories.Items.Count) Einträgen" "UI"
    } else {
        Write-DebugLog "WARNUNG: Keine Kategorien für ComboBox gefunden!" "UI"
    }
}

# Funktion zum Anzeigen der Ergebnisse
function Show-CategoryResults {
    param([string]$Category = "Alle")
    
    Write-DebugLog "Zeige Ergebnisse fuer Kategorie: $Category" "UI"
    Write-DebugLog "Anzahl verfügbare Audit-Ergebnisse: $($global:auditResults.Count)" "UI"
    
    # Sicherheitsvalidierung
    if ($null -eq $rtbResults) {
        Write-DebugLog "FEHLER: rtbResults ist null - kann Ergebnisse nicht anzeigen" "UI"
        return
    }
    
    try {
        # Erstelle immer ein neues, einfaches FlowDocument
        Write-DebugLog "Erstelle neues FlowDocument..." "UI"
        $newDocument = New-Object System.Windows.Documents.FlowDocument
        
        # Sichere, minimale Konfiguration
        $newDocument.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
        $newDocument.FontSize = 11
        $newDocument.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::White)
        
        # Kritische Layout-Eigenschaften setzen um Rendering-Probleme zu vermeiden
        $newDocument.PageWidth = 800  # Feste Breite statt NaN
        $newDocument.PageHeight = [Double]::NaN
        $newDocument.ColumnWidth = 780  # Feste Spaltenbreite
        $newDocument.PagePadding = New-Object System.Windows.Thickness(10)
        $newDocument.TextAlignment = "Left"
        
        # Debug: Liste alle verfügbaren Ergebnisse auf
        if ($global:auditResults.Count -gt 0) {
            Write-DebugLog "Verfügbare Ergebnisse: $($global:auditResults.Keys -join ', ')" "UI"
        }
        
        if ($global:auditResults.Count -eq 0) {
            Write-DebugLog "Keine Audit-Ergebnisse verfügbar - zeige Leer-Meldung" "UI"
            
            $emptyParagraph = New-Object System.Windows.Documents.Paragraph
            $emptyRun = New-Object System.Windows.Documents.Run("Keine Audit-Ergebnisse verfügbar. Führen Sie zuerst ein Audit durch.")
            $emptyRun.FontWeight = "Bold"
            $emptyRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::DarkBlue)
            $emptyParagraph.Inlines.Add($emptyRun)
            $newDocument.Blocks.Add($emptyParagraph)
            
        } else {
            Write-DebugLog "Erstelle Ergebnisanzeige für Kategorie: $Category" "UI"
            
            # Filtere Ergebnisse nach Kategorie
            $filteredResults = @{}
            
            if ($Category -eq "Alle") {
                $filteredResults = $global:auditResults.Clone()
                Write-DebugLog "Zeige alle $($filteredResults.Count) Ergebnisse" "UI"
            } else {
                # Finde Ergebnisse für spezifische Kategorie
                foreach ($cmd in $commands) {
                    if ($cmd.Category -eq $Category -and $global:auditResults.ContainsKey($cmd.Name)) {
                        $filteredResults[$cmd.Name] = $global:auditResults[$cmd.Name]
                    }
                }
                Write-DebugLog "Zeige $($filteredResults.Count) Ergebnisse für Kategorie '$Category'" "UI"
            }
            
            # Erstelle einfache Text-Ausgabe
            $resultCount = 0
            foreach ($resultName in ($filteredResults.Keys | Sort-Object)) {
                $resultCount++
                
                # Begrenze die Anzahl der Ergebnisse um Performance-Probleme zu vermeiden
                if ($resultCount -gt 50) {
                    $moreParagraph = New-Object System.Windows.Documents.Paragraph
                    $moreRun = New-Object System.Windows.Documents.Run("... und $($filteredResults.Count - 50) weitere Ergebnisse")
                    $moreRun.FontStyle = "Italic"
                    $moreRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Gray)
                    $moreParagraph.Inlines.Add($moreRun)
                    $newDocument.Blocks.Add($moreParagraph)
                    break
                }
                
                # Titel-Paragraph
                $titleParagraph = New-Object System.Windows.Documents.Paragraph
                $titleRun = New-Object System.Windows.Documents.Run("=== $resultName ===")
                $titleRun.FontWeight = "Bold"
                $titleRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::DarkBlue)
                $titleParagraph.Inlines.Add($titleRun)
                $titleParagraph.Margin = New-Object System.Windows.Thickness(0, 10, 0, 5)
                $newDocument.Blocks.Add($titleParagraph)
                
                # Content-Paragraph mit sicherer Textbehandlung
                $contentParagraph = New-Object System.Windows.Documents.Paragraph
                $resultContent = $filteredResults[$resultName]
                
                # Sichere Textbehandlung
                if ($null -eq $resultContent) {
                    $resultContent = "[Keine Daten verfügbar]"
                } elseif ($resultContent -is [string]) {
                    # Begrenze die Textlänge um Rendering-Probleme zu vermeiden
                    if ($resultContent.Length -gt 5000) {
                        $resultContent = $resultContent.Substring(0, 5000) + "`n`n[Inhalt gekürzt - zu lang für Anzeige]"
                    }
                    # Entferne problematische Zeichen
                    $resultContent = $resultContent -replace '[^\x20-\x7E\r\n\t]', '?'
                } else {
                    $resultContent = $resultContent.ToString()
                }
                
                $contentRun = New-Object System.Windows.Documents.Run($resultContent)
                $contentRun.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
                $contentParagraph.Inlines.Add($contentRun)
                $contentParagraph.Margin = New-Object System.Windows.Thickness(10, 0, 0, 15)
                $newDocument.Blocks.Add($contentParagraph)
            }
            
            # Summary-Information
            $summaryParagraph = New-Object System.Windows.Documents.Paragraph
            $summaryRun = New-Object System.Windows.Documents.Run("`n--- Angezeigt: $resultCount von $($filteredResults.Count) Ergebnissen ---")
            $summaryRun.FontStyle = "Italic"
            $summaryRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Gray)
            $summaryParagraph.Inlines.Add($summaryRun)
            $newDocument.Blocks.Add($summaryParagraph)
        }
        
        # Setze das Document sicher
        Write-DebugLog "Setze FlowDocument in RichTextBox..." "UI"
        $rtbResults.Document = $newDocument
        
        # Update Summary
        if ($txtResultsSummary) {
            $itemCount = if ($Category -eq "Alle") { $global:auditResults.Count } else { $filteredResults.Count }
            $txtResultsSummary.Text = "Zeige $itemCount Einträge für Kategorie: $Category"
        }
        
        Write-DebugLog "Ergebnisanzeige erfolgreich aktualisiert" "UI"
        
    } catch {
        Write-DebugLog "FEHLER beim Anzeigen der Ergebnisse: $($_.Exception.Message)" "UI"
        Write-DebugLog "Exception-Type: $($_.Exception.GetType().FullName)" "UI"
        
        # Fallback: Verwende einfache Textanzeige
        try {
            Write-DebugLog "Verwende Fallback-Textanzeige..." "UI"
            $fallbackDocument = New-Object System.Windows.Documents.FlowDocument
            $fallbackParagraph = New-Object System.Windows.Documents.Paragraph
            $fallbackRun = New-Object System.Windows.Documents.Run("Ergebnisse können nicht angezeigt werden.`n`nFehler: $($_.Exception.Message)`n`nVerfügbare Ergebnisse: $($global:auditResults.Count)")
            $fallbackParagraph.Inlines.Add($fallbackRun)
            $fallbackDocument.Blocks.Add($fallbackParagraph)
            $rtbResults.Document = $fallbackDocument
        } catch {
            Write-DebugLog "KRITISCHER FEHLER: Auch Fallback-Anzeige fehlgeschlagen: $($_.Exception.Message)" "UI"
        }
    }
}

# Fallback-Funktion für einfache Textanzeige
function Show-SimpleResults {
    param([string]$Category = "Alle")
    
    Write-DebugLog "Verwende einfache Textanzeige für Kategorie: $Category" "UI"
    
    # Erstelle einfaches FlowDocument
    $document = New-Object System.Windows.Documents.FlowDocument
    $document.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $document.FontSize = 11
    $document.PageWidth = [Double]::NaN
    $document.PageHeight = [Double]::NaN
    $document.ColumnWidth = [Double]::PositiveInfinity
    
    # Sammle alle relevanten Ergebnisse als einfachen Text
    $resultText = ""
    
    # Gruppiere nach Kategorien
    $categorizedResults = @{}
    foreach ($cmd in $commands) {
        $cmdCategory = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
        if (-not $categorizedResults.ContainsKey($cmdCategory)) {
            $categorizedResults[$cmdCategory] = @()
        }
        if ($global:auditResults.ContainsKey($cmd.Name)) {
            $categorizedResults[$cmdCategory] += @{
                Name = $cmd.Name
                Result = $global:auditResults[$cmd.Name]
            }
        }
    }
    
    # Bestimme anzuzeigende Kategorien
    $categoriesToShow = if ($Category -eq "Alle") { 
        $categorizedResults.Keys | Sort-Object 
    } else { 
        @($Category) 
    }
    
    $totalItems = 0
    foreach ($cat in $categoriesToShow) {
        if ($categorizedResults.ContainsKey($cat)) {
            $categoryData = $categorizedResults[$cat]
            $totalItems += $categoryData.Count
            
            $resultText += "`n" + "="*60 + "`n"
            $resultText += "KATEGORIE: $cat`n"
            $resultText += "="*60 + "`n`n"
            
            foreach ($item in $categoryData) {
                $resultText += "-"*40 + "`n"
                $resultText += "EINTRAG: $($item.Name)`n"
                $resultText += "-"*40 + "`n"
                $resultText += "$($item.Result)`n`n"
            }
        }
    }
    
    # Erstelle einfachen Paragraph mit dem gesamten Text
    $paragraph = New-Object System.Windows.Documents.Paragraph
    $run = New-Object System.Windows.Documents.Run($resultText)
    $paragraph.Inlines.Add($run)
    $document.Blocks.Add($paragraph)
    
    $rtbResults.Document = $document
    
    # Update Summary
    $window.Dispatcher.Invoke([Action]{
        $txtResultsSummary.Text = "Zeige $totalItems Einträge (einfache Ansicht)"
    }, "Normal")
}

# Initialisiere die Ergebnisse-Anzeige
Show-CategoryResults -Category "Alle"

# Optimiere die RichTextBox für besseres Text-Wrapping
Optimize-RichTextBoxLayout

# Funktion zum Wechseln zwischen den Panels
function Switch-Panel {
    param([string]$PanelName)
    
    # Alle Panels ausblenden
    if ($serverAuditPanel) { $serverAuditPanel.Visibility = "Collapsed" }
    if ($serverResultsPanel) { $serverResultsPanel.Visibility = "Collapsed" }
    if ($networkAuditPanel) { $networkAuditPanel.Visibility = "Collapsed" }
    if ($networkResultsPanel) { $networkResultsPanel.Visibility = "Collapsed" }
    if ($toolsPanel) { $toolsPanel.Visibility = "Collapsed" }
    if ($debugPanel) { $debugPanel.Visibility = "Collapsed" }
    
    # Alle Navigation-Buttons zurücksetzen
    if ($btnNavServerAudit) { $btnNavServerAudit.Tag = "" }
    if ($btnNavServerResults) { $btnNavServerResults.Tag = "" }
    if ($btnNavNetworkAudit) { $btnNavNetworkAudit.Tag = "" }
    if ($btnNavNetworkResults) { $btnNavNetworkResults.Tag = "" }
    if ($btnNavTools) { $btnNavTools.Tag = "" }
    if ($btnNavDebug) { $btnNavDebug.Tag = "" }
    
    # Gewähltes Panel anzeigen und Button aktivieren
    switch ($PanelName) {
        "serverAudit" {
            if ($serverAuditPanel) { $serverAuditPanel.Visibility = "Visible" }
            if ($btnNavServerAudit) { $btnNavServerAudit.Tag = "active" }
        }
        "serverResults" {
            if ($serverResultsPanel) { $serverResultsPanel.Visibility = "Visible" }
            if ($btnNavServerResults) { $btnNavServerResults.Tag = "active" }
        }
        "networkAudit" {
            if ($networkAuditPanel) { $networkAuditPanel.Visibility = "Visible" }
            if ($btnNavNetworkAudit) { $btnNavNetworkAudit.Tag = "active" }
        }
        "networkResults" {
            if ($networkResultsPanel) { $networkResultsPanel.Visibility = "Visible" }
            if ($btnNavNetworkResults) { $btnNavNetworkResults.Tag = "active" }
        }
        "tools" {
            if ($toolsPanel) { $toolsPanel.Visibility = "Visible" }
            if ($btnNavTools) { $btnNavTools.Tag = "active" }
        }
        "debug" {
            if ($debugPanel) { $debugPanel.Visibility = "Visible" }
            if ($btnNavDebug) { $btnNavDebug.Tag = "active" }
        }
    }
}

# Initialisiere die GUI - Startpanel setzen
Write-DebugLog "Initialisiere GUI - setze Server Audit Panel als Standard" "Init"
Switch-Panel "serverAudit"

# Setze Status-Text
$txtStatus.Text = "Status: Bereit für Server-Audit"

# Initialisiere die Ergebnisse-Anzeige
Show-CategoryResults -Category "Alle"

# Verbindungsaudit Export-Button-Funktionalitaet
$btnExportConnectionHTML.Add_Click({
    Write-DebugLog "Verbindungsaudit HTML-Export gestartet" "Export"
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "HTML Files (*.html)|*.html"
    $saveFileDialog.Title = "Speichern Sie den Verbindungsaudit-Bericht"
    $saveFileDialog.FileName = "ConnectionAudit_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtStatus.Text = "Status: Exportiere Verbindungsaudit-HTML..."
        
        try {
            Export-ConnectionAuditToHTML -Results $global:connectionAuditResults -FilePath $saveFileDialog.FileName
            $txtStatus.Text = "Status: Verbindungsaudit-Export erfolgreich abgeschlossen"
            [System.Windows.MessageBox]::Show("Verbindungsaudit-Bericht wurde erfolgreich exportiert:`r`n$($saveFileDialog.FileName)", "Export erfolgreich", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } catch {
            $txtStatus.Text = "Status: Fehler beim Verbindungsaudit-Export"
            [System.Windows.MessageBox]::Show("Fehler beim Export:`r`n$($_.Exception.Message)", "Export Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Verbindungsaudit Zwischenablage-Button
$btnCopyConnectionToClipboard.Add_Click({
    Write-DebugLog "Kopiere Verbindungsaudit-Ergebnisse in Zwischenablage" "UI"
    try {
        # Extrahiere Text aus der RichTextBox
        $textRange = New-Object System.Windows.Documents.TextRange($rtbConnectionResults.Document.ContentStart, $rtbConnectionResults.Document.ContentEnd)
        $plainText = $textRange.Text
        
        if ([string]::IsNullOrWhiteSpace($plainText)) {
            # Fallback: Erstelle Text aus den Rohdaten
            $allResults = ""
            foreach ($key in $global:connectionAuditResults.Keys | Sort-Object) {
                $allResults += "`r`n=== $key ===`r`n$($global:connectionAuditResults[$key])`r`n"
            }
            $plainText = $allResults
        }
        
        $plainText | Set-Clipboard
        $txtStatus.Text = "Status: Verbindungsaudit-Ergebnisse in Zwischenablage kopiert"
        [System.Windows.MessageBox]::Show("Verbindungsaudit-Ergebnisse wurden in die Zwischenablage kopiert.", "Kopiert", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Kopieren in die Zwischenablage: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# Verbindungsaudit ComboBox Selection Changed Event
$cmbConnectionCategories.Add_SelectionChanged({
    if ($cmbConnectionCategories.SelectedItem) {
        $selectedCategory = $cmbConnectionCategories.SelectedItem.Tag
        Show-ConnectionCategoryResults -Category $selectedCategory
    }
})

# Initialisiere die Verbindungsaudit-Ergebnisse-Anzeige
Show-ConnectionCategoryResults -Category "Alle"

# Initialisiere die Ergebnisse-Anzeige
Show-CategoryResults -Category "Alle"

# Funktion zum Bereinigen von Sonderzeichen, Umlauten und Symbolen
function Clean-StringForDiagram {
    param(
        [string]$InputString
    )
    
    if ([string]::IsNullOrWhiteSpace($InputString)) {
        return "Unbekannt"
    }
    
    # Umlaute und Sonderzeichen ersetzen
    $cleanString = $InputString -replace 'ä', 'ae' -replace 'ö', 'oe' -replace 'ü', 'ue' -replace 'Ä', 'Ae' -replace 'Ö', 'Oe' -replace 'Ü', 'Ue' -replace 'ß', 'ss'
    
    # Sonderzeichen und Symbole entfernen oder ersetzen
    $cleanString = $cleanString -replace '[^\w\s\.\-_:]', '' -replace '\s+', ' '
    $cleanString = $cleanString.Trim()
    
    # XML-sichere Zeichen
    $cleanString = $cleanString -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&apos;'
    
    if ([string]::IsNullOrWhiteSpace($cleanString)) {
        return "Bereinigt"
    }
    
    return $cleanString
}

# Funktion zum Erstellen eines DRAW.IO XML Exports der Netzwerk-Topologie
function Export-NetworkTopologyToDrawIO {
    param(
        [hashtable]$Results,
        [string]$FilePath,
        [string]$ServerName = $env:COMPUTERNAME
    )
    
    Write-DebugLog "Starte DRAW.IO Netzwerk-Topologie Export nach: $FilePath" "DrawIO-Export"
    
    try {
        $cleanServerName = Clean-StringForDiagram -InputString $ServerName
        
        # --- Datenextraktion und -verarbeitung ---
        $processedData = @{
            TCPConnections = @()
            NetworkAdapters = @() # Wird unten neu befüllt
            ListeningPorts = @()
            ExternalConnections = @()
            GatewayIP = "N/A"
            DnsServers = "N/A"
            PrimaryIP = "IP nicht ermittelt"
            ServerOS = "Windows Server" # Generisch, da OS-Info nicht direkt in $Results erwartet wird
        }

        # TCP Verbindungen verarbeiten (für Statistiken und ggf. IP-Ermittlung)
        if ($Results.ContainsKey("Alle TCP-Verbindungen (Performance)") -or $Results.ContainsKey("Etablierte TCP-Verbindungen")) {
            $tcpData = $Results["Alle TCP-Verbindungen (Performance)"]
            if (-not $tcpData) { $tcpData = $Results["Etablierte TCP-Verbindungen"] }
            if ($tcpData) {
                $tcpLines = $tcpData -split "`n" | Where-Object { $_ -match '\d+\.\d+\.\d+\.\d+' }
                foreach ($line in $tcpLines) {
                    if ($line -match '(\d+\.\d+\.\d+\.\d+)\s+(\d+)\s+(\d+\.\d+\.\d+\.\d+)\s+(\d+)\s+(\w+)') {
                        $processedData.TCPConnections += @{
                            LocalIP = $matches[1]; LocalPort = $matches[2]; RemoteIP = $matches[3]; RemotePort = $matches[4]; State = $matches[5]
                        }
                    }
                }
            }
        }

        # Netzwerkadapter-Daten detaillierter parsen
        $parsedAdapters = @()
        $adapterDataSource = $null
        if ($Results.ContainsKey("Erweiterte Netzwerk-Adapter-Infos")) { $adapterDataSource = $Results["Erweiterte Netzwerk-Adapter-Infos"] }
        elseif ($Results.ContainsKey("Netzwerkadapter")) { $adapterDataSource = $Results["Netzwerkadapter"] }

        if ($adapterDataSource) {
            $currentAdapter = $null
            $adapterLines = $adapterDataSource -split '\r?\n'
            foreach ($line in $adapterLines) {
                if ($line -match '^(Name|InterfaceAlias)\s*:\s*(.+)$' -or $line -match '^\s*Beschreibung\.+:\s*(.+)$') { # Deutsch: Beschreibung
                    if ($currentAdapter) { $parsedAdapters += $currentAdapter }
                    $currentAdapter = @{ Name = Clean-StringForDiagram ($matches[1]).Trim(); Status = "Unknown"; Description = ""; IPAddress = "N/A"; SubnetMask = "N/A" }
                    if ($line -match '^(Name|InterfaceAlias)\s*:\s*(.+)$') { $currentAdapter.Name = Clean-StringForDiagram ($matches[2]).Trim() }
                    else { $currentAdapter.Description = Clean-StringForDiagram ($matches[1]).Trim() } # Fallback für Name wenn nur Beschreibung da
                } elseif ($currentAdapter) {
                    if ($line -match '^\s*(Status|Status der Verbindung)\s*:\s*(.+)$') { $currentAdapter.Status = Clean-StringForDiagram ($matches[2]).Trim() } # Deutsch: Status der Verbindung
                    elseif ($line -match '^\s*(InterfaceDescription|Beschreibung)\s*:\s*(.+)$') { $currentAdapter.Description = Clean-StringForDiagram ($matches[2]).Trim() }
                    elseif ($line -match '^\s*IPv4-Adresse\.+:\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}).*') { # Deutsch: IPv4-Adresse
                         if ($matches[1] -ne "0.0.0.0") { $currentAdapter.IPAddress = $matches[1] }
                    } elseif ($line -match '^\s*IPv4Address\s*:\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})') { # Englisch
                         if ($matches[1] -ne "0.0.0.0") { $currentAdapter.IPAddress = $matches[1] }
                    } elseif ($line -match '^\s*Subnetzmaske\.+:\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})') { # Deutsch
                        $currentAdapter.SubnetMask = $matches[1]
                    } elseif ($line -match '^\s*SubnetMask\s*:\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})') { # Englisch
                        $currentAdapter.SubnetMask = $matches[1]
                    }
                }
            }
            if ($currentAdapter) { $parsedAdapters += $currentAdapter }
            $processedData.NetworkAdapters = $parsedAdapters | Where-Object { -not ([string]::IsNullOrWhiteSpace($_.Name)) }
        }
        
        # Primäre IP-Adresse ermitteln
        $activeAdapterWithIP = $processedData.NetworkAdapters | Where-Object { ($_.Status -eq "Up" -or $_.Status -eq "Aktiviert") -and $_.IPAddress -ne "N/A" -and $_.IPAddress -ne "0.0.0.0" -and $_.IPAddress -notmatch "^169\.254\." -and $_.IPAddress -ne "127.0.0.1"} | Select-Object -First 1
        if ($activeAdapterWithIP) {
            $processedData.PrimaryIP = $activeAdapterWithIP.IPAddress
        } elseif ($processedData.TCPConnections.Count -gt 0) {
            $firstLocalTCP_IP = ($processedData.TCPConnections | Where-Object {$_.LocalIP -ne "0.0.0.0" -and $_.LocalIP -ne "127.0.0.1"} | Select-Object -First 1).LocalIP
            if ($firstLocalTCP_IP) { $processedData.PrimaryIP = $firstLocalTCP_IP }
        }


        # Gateway IP ermitteln
        if ($Results.ContainsKey("Netzwerkkonfiguration (ipconfig)")) {
            $ipConfigData = $Results["Netzwerkkonfiguration (ipconfig)"]
            if ($ipConfigData -match '(Standardgateway|Default Gateway).+:\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})') {
                if ($matches[2] -ne "0.0.0.0") { $processedData.GatewayIP = $matches[2] }
            }
        } elseif ($Results.ContainsKey("Routing Tabelle")) {
            $routingData = $Results["Routing Tabelle"]
            $routeLines = $routingData -split "`n" | Where-Object { $_ -match '^\s*0\.0\.0\.0\s+0\.0\.0\.0\s+(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})' }
            if ($routeLines.Count -gt 0) {
                # $matches is not available outside the Where-Object script block in this context directly
                # Need to re-match or extract differently
                $gwMatch = $routeLines[0] | Select-String -Pattern '^\s*0\.0\.0\.0\s+0\.0\.0\.0\s+(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})'
                if ($gwMatch) { $processedData.GatewayIP = $gwMatch.Matches[0].Groups[1].Value }
            }
        }
         if ([string]::IsNullOrWhiteSpace($processedData.GatewayIP) -or $processedData.GatewayIP -match "0.0.0.0") { $processedData.GatewayIP = "N/A" }


        # DNS Server ermitteln
        if ($Results.ContainsKey("Netzwerkkonfiguration (ipconfig)")) {
            $ipConfigData = $Results["Netzwerkkonfiguration (ipconfig)"]
            $dnsMatches = $ipConfigData | Select-String -Pattern '(DNS-Server|DNS Servers).+:\s*((\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\s*)+)' -AllMatches
            if ($dnsMatches) {
                $dnsIPs = $dnsMatches.Matches.Groups[2].Value -split '\s+' | Where-Object {$_ -match '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'} | Get-Unique
                if ($dnsIPs.Count -gt 0) { $processedData.DnsServers = $dnsIPs -join ', ' }
            }
        }
        if ([string]::IsNullOrWhiteSpace($processedData.DnsServers)) { $processedData.DnsServers = "N/A" }

        # Lauschende Ports verarbeiten
        if ($Results.ContainsKey("Lauschende Ports (Listen)")) {
            $listenData = $Results["Lauschende Ports (Listen)"]
            if ($listenData) {
                $listenLines = $listenData -split "`n" | Where-Object { $_ -match '(\d+\.\d+\.\d+\.\d+|\[::\]|0\.0\.0\.0)\s*:\s*(\d+)' } # Adjusted regex for IP:Port format
                foreach ($line in $listenLines) {
                     if ($line -match '(\d+\.\d+\.\d+\.\d+|\[::\]|0\.0\.0\.0)\s*:\s*(\d+)') { # More specific for IP:Port
                        $processedData.ListeningPorts += @{ IP = $matches[1]; Port = $matches[2] }
                    } elseif ($line -match '(\S+)\s+(\d+)\s+LISTENING') { # Fallback for netstat like format if IP:Port fails
                        $processedData.ListeningPorts += @{ IP = $matches[1]; Port = $matches[2] }
                    }
                }
            }
        }
        
        # Externe Verbindungen verarbeiten
        if ($Results.ContainsKey("Externe Verbindungen (Internet)")) {
            $externalData = $Results["Externe Verbindungen (Internet)"]
            if ($externalData) {
                $externalLines = $externalData -split "`n" | Where-Object { $_ -match '\d+\.\d+\.\d+\.\d+' }
                foreach ($line in $externalLines) {
                    if ($line -match '(\d+\.\d+\.\d+\.\d+)\s*:\s*(\d+)\s+(\d+\.\d+\.\d+\.\d+)\s*:\s*(\d+)') { # Adjusted for IP:Port format
                        $processedData.ExternalConnections += @{ LocalIP = $matches[1]; LocalPort = $matches[2]; RemoteIP = $matches[3]; RemotePort = $matches[4] }
                    }
                }
            }
        }

        # --- XML Generierung ---
        $cellIdCounter = 10 # Start counter for unique IDs

        # XML Header
        $xmlContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<mxfile host="Electron" agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) draw.io/27.0.5 Chrome/134.0.6998.205 Electron/35.3.0 Safari/537.36" version="27.0.5">
  <diagram name="Netzwerk-Topologie $(Clean-StringForDiagram $ServerName)" id="$(New-Guid)">
    <mxGraphModel dx="1700" dy="1000" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="1169" pageHeight="827" math="0" shadow="0">
      <root>
        <mxCell id="0" />
        <mxCell id="1" parent="0" />
"@
        # Zentraler Server
        $serverValue = "Hostname: $cleanServerName&#xa;IP: $(Clean-StringForDiagram $processedData.PrimaryIP)&#xa;OS: $(Clean-StringForDiagram $processedData.ServerOS)"
        $xmlContent += @"
        <!-- Zentraler Server -->
        <mxCell id="server-central" value="$serverValue" style="shape=mxgraph.windows.server;html=1;whiteSpace=wrap;fontSize=12;fontStyle=1;fillColor=#dae8fc;strokeColor=#6c8ebf;strokeWidth=2;" vertex="1" parent="1">
          <mxGeometry x="750" y="450" width="180" height="100" as="geometry" />
        </mxCell>
"@
        # Gateway
        $gatewayValue = "Gateway&#xa;IP: $(Clean-StringForDiagram $processedData.GatewayIP)"
        $xmlContent += @"
        <!-- Gateway -->
        <mxCell id="gateway-node" value="$gatewayValue" style="shape=mxgraph.cisco.routers.router_with_firewall_symbol;html=1;whiteSpace=wrap;fontSize=10;fillColor=#f8cecc;strokeColor=#b85450;" vertex="1" parent="1">
          <mxGeometry x="780" y="100" width="120" height="70" as="geometry" />
        </mxCell>
        <mxCell id="edge-server-gateway" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1;" edge="1" parent="1" source="server-central" target="gateway-node">
          <mxGeometry relative="1" as="geometry" />
        </mxCell>
"@
        # DNS Server
        $dnsValue = "DNS Server&#xa;$(Clean-StringForDiagram $processedData.DnsServers)"
        $xmlContent += @"
        <!-- DNS Server -->
        <mxCell id="dns-node" value="$dnsValue" style="shape=mxgraph.cisco.servers.dns_server;html=1;whiteSpace=wrap;fontSize=10;fillColor=#fff2cc;strokeColor=#d6b656;" vertex="1" parent="1">
          <mxGeometry x="1050" y="250" width="140" height="80" as="geometry" />
        </mxCell>
        <mxCell id="edge-server-dns" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1;" edge="1" parent="1" source="server-central" target="dns-node">
          <mxGeometry relative="1" as="geometry" />
        </mxCell>
"@
        # NAS/SAN (Placeholder)
        $xmlContent += @"
        <!-- NAS/SAN -->
        <mxCell id="nas-node" value="NAS / SAN&#xa;(Details manuell eintragen)" style="shape=mxgraph.cisco.storage.nas_icon;html=1;whiteSpace=wrap;fontSize=10;fillColor=#e1d5e7;strokeColor=#9673a6;" vertex="1" parent="1">
          <mxGeometry x="450" y="250" width="140" height="80" as="geometry" />
        </mxCell>
        <mxCell id="edge-server-nas" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1;" edge="1" parent="1" source="server-central" target="nas-node">
          <mxGeometry relative="1" as="geometry" />
        </mxCell>
"@
        # Internet Cloud (für externe Verbindungen)
        $xmlContent += @"
        <!-- Internet Cloud -->
        <mxCell id="internet-cloud" value="Internet" style="shape=cloud;html=1;whiteSpace=wrap;fontSize=12;fillColor=#f5f5f5;strokeColor=#666666;" vertex="1" parent="1">
          <mxGeometry x="780" y="750" width="120" height="80" as="geometry" />
        </mxCell>
        <mxCell id="edge-gateway-internet" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1;" edge="1" parent="1" source="gateway-node" target="internet-cloud">
          <mxGeometry relative="1" as="geometry" />
        </mxCell>
"@
        # Netzwerkadapter (als NETWORK / NETWORK2 etc.)
        $adapterXStart = 100
        $adapterY = 480
        $adapterCount = 0
        foreach ($adapter in ($processedData.NetworkAdapters | Where-Object {$_.Status -ne "Disabled" -and $_.Status -ne "Deaktiviert"} | Select-Object -First 3)) {
            $cellIdCounter++
            $adapterNameClean = Clean-StringForDiagram $adapter.Name
            $adapterIpClean = Clean-StringForDiagram $adapter.IPAddress
            $adapterSubnetClean = Clean-StringForDiagram $adapter.SubnetMask
            $adapterStatusClean = Clean-StringForDiagram $adapter.Status
            
            $adapterLabel = "Adapter: $adapterNameClean&#xa;IP: $adapterIpClean&#xa;Subnetz: $adapterSubnetClean&#xa;Status: $adapterStatusClean"
            $fillColor = if ($adapter.Status -eq "Up" -or $adapter.Status -eq "Aktiviert") { "#d5e8d4" } else { "#f8cecc" } # Grün für Up, Rot für Down
            $strokeColor = if ($adapter.Status -eq "Up" -or $adapter.Status -eq "Aktiviert") { "#82b366" } else { "#b85450" }

            $xmlContent += @"
        <!-- Netzwerkadapter: $adapterNameClean -->
        <mxCell id="adapter-$cellIdCounter" value="$adapterLabel" style="shape=card;html=1;whiteSpace=wrap;fontSize=9;fillColor=$fillColor;strokeColor=$strokeColor;" vertex="1" parent="1">
          <mxGeometry x="$($adapterXStart + ($adapterCount * 200))" y="$adapterY" width="160" height="90" as="geometry" />
        </mxCell>
        <mxCell id="edge-server-adapter-$cellIdCounter" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=1;" edge="1" parent="1" source="server-central" target="adapter-$cellIdCounter">
          <mxGeometry relative="1" as="geometry" />
        </mxCell>
"@
            $adapterCount++
        }

        # Lauschende Ports (als kleine verbundene Elemente)
        $listenPortXStart = 1050
        $listenPortYStart = 400
        $listenPortMax = 25
        $listenPortCurrent = 0
        $uniqueListeningPorts = $processedData.ListeningPorts | Sort-Object Port -Unique | Select-Object -First $listenPortMax
        
        foreach ($portInfo in $uniqueListeningPorts) {
            $cellIdCounter++
            $portClean = Clean-StringForDiagram $portInfo.Port
            $portLabel = "Port: $portClean"
            
            $xmlContent += @"
        <!-- Lauschender Port: $portClean -->
        <mxCell id="lport-$cellIdCounter" value="$portLabel" style="ellipse;shape=doubleEllipse;html=1;whiteSpace=wrap;fontSize=9;fillColor=#dae8fc;strokeColor=#6c8ebf;perimeter=ellipsePerimeter;" vertex="1" parent="1">
          <mxGeometry x="$($listenPortXStart + ($listenPortCurrent % 2 * 70))" y="$($listenPortYStart + ([Math]::Floor($listenPortCurrent / 2) * 50))" width="60" height="40" as="geometry" />
        </mxCell>
        <mxCell id="edge-server-lport-$cellIdCounter" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#6c8ebf;strokeWidth=1;endArrow=none;" edge="1" parent="1" source="server-central" target="lport-$cellIdCounter">
          <mxGeometry relative="1" as="geometry" />
        </mxCell>
"@
            $listenPortCurrent++
        }

        # Externe Verbindungen (gruppiert nach Remote IP)
        $extConnXStart = 450
        $extConnYStart = 700 
        $extConnMaxGroups = 3
        $extConnGroups = $processedData.ExternalConnections | Group-Object RemoteIP | Select-Object -First $extConnMaxGroups
        $extConnCounter = 0

        foreach ($group in $extConnGroups) {
            $cellIdCounter++
            $remoteIpClean = Clean-StringForDiagram $group.Name
            $ports = ($group.Group | Select-Object -ExpandProperty RemotePort -Unique | Select-Object -First 3) -join ", "
            if (($group.Group | Select-Object -ExpandProperty RemotePort -Unique).Count -gt 3) { $ports += ", ..." }
            $extConnLabel = "Extern: $remoteIpClean&#xa;Ports: $ports"

            $xmlContent += @"
        <!-- Externe Verbindung: $remoteIpClean -->
        <mxCell id="extconn-$cellIdCounter" value="$extConnLabel" style="rounded=0;whiteSpace=wrap;html=1;fontSize=9;fillColor=#fad7ac;strokeColor=#b46504;" vertex="1" parent="1">
          <mxGeometry x="$($extConnXStart + ($extConnCounter * 130))" y="$($extConnYStart - 100)" width="120" height="60" as="geometry" />
        </mxCell>
        <mxCell id="edge-server-extconn-$cellIdCounter" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#b46504;strokeWidth=1;dashed=1;" edge="1" parent="1" source="server-central" target="extconn-$cellIdCounter">
          <mxGeometry relative="1" as="geometry" />
        </mxCell>
        <mxCell id="edge-extconn-internet-$cellIdCounter" style="edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#b46504;strokeWidth=1;dashed=1;" edge="1" parent="1" source="extconn-$cellIdCounter" target="internet-cloud">
          <mxGeometry relative="1" as="geometry" />
        </mxCell>
"@
            $extConnCounter++
        }
        
        # Legende und Statistiken
        $statsLabel = "Statistiken&#xa;TCP Gesamt: $($processedData.TCPConnections.Count)&#xa;Externe Verb.: $($processedData.ExternalConnections.Count)&#xa;Lauschende Ports: $($processedData.ListeningPorts.Count)"
        $xmlContent += @"
        <!-- Statistiken -->
        <mxCell id="stats-node" value="$statsLabel" style="text;html=1;strokeColor=none;fillColor=none;align=left;verticalAlign=middle;whiteSpace=wrap;rounded=0;fontSize=10;fontStyle=0;" vertex="1" parent="1">
          <mxGeometry x="50" y="50" width="180" height="70" as="geometry" />
        </mxCell>
        
        <!-- Legende -->
        <mxCell id="legend-title" value="Legende" style="text;html=1;strokeColor=none;fillColor=none;align=left;verticalAlign=middle;whiteSpace=wrap;rounded=0;fontSize=12;fontStyle=1;" vertex="1" parent="1">
          <mxGeometry x="50" y="130" width="150" height="20" as="geometry" />
        </mxCell>
        <mxCell id="legend-server" value="Zentraler Server" style="shape=mxgraph.windows.server;html=1;fontSize=9;fillColor=#dae8fc;strokeColor=#6c8ebf;" vertex="1" parent="1">
          <mxGeometry x="50" y="160" width="120" height="30" as="geometry" />
        </mxCell>
        <mxCell id="legend-gateway" value="Gateway" style="shape=mxgraph.cisco.routers.router_with_firewall_symbol;html=1;fontSize=9;fillColor=#f8cecc;strokeColor=#b85450;" vertex="1" parent="1">
          <mxGeometry x="50" y="200" width="120" height="30" as="geometry" />
        </mxCell>
        <mxCell id="legend-dns" value="DNS Server" style="shape=mxgraph.cisco.servers.dns_server;html=1;fontSize=9;fillColor=#fff2cc;strokeColor=#d6b656;" vertex="1" parent="1">
          <mxGeometry x="50" y="240" width="120" height="30" as="geometry" />
        </mxCell>
        <mxCell id="legend-adapter" value="Netzwerkadapter" style="shape=card;html=1;fontSize=9;fillColor=#d5e8d4;strokeColor=#82b366;" vertex="1" parent="1">
          <mxGeometry x="50" y="280" width="120" height="30" as="geometry" />
        </mxCell>
        <mxCell id="legend-lport" value="Lauschender Port" style="ellipse;shape=doubleEllipse;html=1;fontSize=9;fillColor=#dae8fc;strokeColor=#6c8ebf;" vertex="1" parent="1">
          <mxGeometry x="50" y="320" width="120" height="30" as="geometry" />
        </mxCell>
        <mxCell id="legend-extconn" value="Externe Verbindung" style="rounded=0;html=1;fontSize=9;fillColor=#fad7ac;strokeColor=#b46504;" vertex="1" parent="1">
          <mxGeometry x="50" y="360" width="120" height="30" as="geometry" />
        </mxCell>
"@

        # XML Footer
        $xmlContent += @"
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>
"@
        
        # Datei schreiben
        $xmlContent | Out-File -FilePath $FilePath -Encoding UTF8 -Force
        
        Write-DebugLog "DRAW.IO Netzwerk-Topologie Export erfolgreich abgeschlossen" "DrawIO-Export"
        Write-DebugLog "Verarbeitete TCP-Verbindungen: $($processedData.TCPConnections.Count)" "DrawIO-Export"
        
        return $true
    }
    catch {
        Write-DebugLog "FEHLER beim DRAW.IO Export: $($_.Exception.Message)" "DrawIO-Export"
        Write-DebugLog "Stack Trace: $($_.ScriptStackTrace)" "DrawIO-Export"
        throw # Re-throw original exception to preserve details
    }
}

# Verbindungsaudit DRAW.IO Export-Button-Funktionalitaet
$btnExportConnectionDrawIO.Add_Click({
    Write-DebugLog "Verbindungsaudit DRAW.IO-Export gestartet" "Export"
    
    if ($global:connectionAuditResults.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Keine Verbindungsaudit-Ergebnisse zum Exportieren vorhanden.", "Keine Daten", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "DRAW.IO XML Files (*.xml)|*.xml"
    $saveFileDialog.Title = "Speichern Sie die Netzwerk-Topologie"
    $saveFileDialog.FileName = "ConnectionTopology_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtProgressConnection.Text = "Exportiere Netzwerk-Topologie..."
        
        try {
            Export-NetworkTopologyToDrawIO -Results $global:connectionAuditResults -FilePath $saveFileDialog.FileName -ServerName $env:COMPUTERNAME
            $txtProgressConnection.Text = "DRAW.IO-Export erfolgreich abgeschlossen"
            [System.Windows.MessageBox]::Show("Netzwerk-Topologie wurde erfolgreich exportiert:`r`n$($saveFileDialog.FileName)", "Export erfolgreich", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } catch {
            $txtProgressConnection.Text = "Fehler beim DRAW.IO-Export"
            [System.Windows.MessageBox]::Show("Fehler beim DRAW.IO-Export:`r`n$($_.Exception.Message)", "Export Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Initialisiere die Ergebnisse-Anzeige
Show-CategoryResults -Category "Alle"

# Navigation Event-Handler
$btnNavServerAudit.Add_Click({
    Switch-Panel "serverAudit"
})

$btnNavServerResults.Add_Click({
    Switch-Panel "serverResults"
    # Aktualisiere Ergebnisse falls vorhanden
    if ($global:auditResults.Count -gt 0) {
        try {
            Update-ResultsCategories
            Show-CategoryResults -Category "Alle"
        }
        catch {
            Write-DebugLog "FEHLER beim Aktualisieren der Server-Audit-Ergebnisse: $($_.Exception.Message)" "Navigation"
        }
    }
})

$btnNavNetworkAudit.Add_Click({
    Switch-Panel "networkAudit"
})

$btnNavNetworkResults.Add_Click({
    Switch-Panel "networkResults"
    # Aktualisiere Verbindungsaudit-Ergebnisse falls vorhanden
    if ($global:connectionAuditResults.Count -gt 0) {
        try {
            Update-ConnectionResultsCategories
            Show-ConnectionCategoryResults -Category "Alle"
        }
        catch {
            Write-DebugLog "FEHLER beim Aktualisieren der Verbindungsaudit-Ergebnisse: $($_.Exception.Message)" "Navigation"
        }
    }
})

$btnNavTools.Add_Click({
    Switch-Panel "tools"
})

$btnNavDebug.Add_Click({
    Switch-Panel "debug"
})

Write-DebugLog "GUI initialisiert, warte auf Benutzerinteraktion" "Startup"

# Zeige das Fenster an
$txtStatus.Text = "Status: Bereit fuer Audit"
Write-DebugLog "Zeige Hauptfenster" "UI"

try {
    # Umfassende Window-Validierung vor dem Anzeigen
    if ($null -eq $window) {
        Write-DebugLog "FEHLER: Window-Objekt ist null, kann nicht angezeigt werden" "UI"
        Write-Host "KRITISCHER FEHLER: Fenster konnte nicht angezeigt werden" -ForegroundColor Red
        Read-Host "Drücken Sie eine Taste zum Beenden"
        return
    }
    
    # Prüfe ob Window korrekt initialisiert ist
    if (-not ($window -is [System.Windows.Window])) {
        Write-DebugLog "FEHLER: Window-Objekt ist nicht vom korrekten Typ: $($window.GetType().FullName)" "UI"
        Write-Host "KRITISCHER FEHLER: Fenster ist nicht vom korrekten Typ" -ForegroundColor Red
        Read-Host "Drücken Sie eine Taste zum Beenden"
        return
    }
    
    # Prüfe wichtige Window-Eigenschaften
    Write-DebugLog "Window-Validierung: Title='$($window.Title)', Type='$($window.GetType().FullName)'" "UI"
    
    # Versuche das Fenster anzuzeigen
    Write-DebugLog "Rufe ShowDialog() auf..." "UI"
    [VOID]$window.ShowDialog()
    Write-DebugLog "ShowDialog() abgeschlossen. Ergebnis: $result" "UI"
    
} catch [System.InvalidOperationException] {
    Write-DebugLog "InvalidOperationException beim Anzeigen des Fensters: $($_.Exception.Message)" "UI"
    Write-Host "FEHLER: Fenster kann nicht in diesem Thread angezeigt werden (Threading-Problem)" -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    
    # Versuche STA-Thread-Information zu geben
    Write-Host "Aktueller Thread-State: $([System.Threading.Thread]::CurrentThread.ApartmentState)" -ForegroundColor Yellow
    
} catch [System.ArgumentException] {
    Write-DebugLog "ArgumentException beim Anzeigen des Fensters: $($_.Exception.Message)" "UI"
    Write-Host "FEHLER: Ungültiges Window-Objekt oder Parameter" -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    
} catch {
    Write-DebugLog "FEHLER beim Anzeigen des Fensters: $($_.Exception.Message)" "UI"
    Write-DebugLog "Exception-Type: $($_.Exception.GetType().FullName)" "UI"
    Write-DebugLog "Fehlerstapel: $($_.Exception.StackTrace)" "UI"
    
    Write-Host "FEHLER beim Anzeigen des Hauptfensters: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Exception-Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    
    # Detaillierte Debugging-Informationen
    if ($_.Exception.InnerException) {
        Write-Host "Innere Exception: $($_.Exception.InnerException.Message)" -ForegroundColor Red
        Write-DebugLog "Innere Exception: $($_.Exception.InnerException.ToString())" "UI"
    }
    
    # Window-Status-Informationen ausgeben
    if ($null -ne $window) {
        try {
            Write-DebugLog "Window-Debug: IsLoaded=$($window.IsLoaded), IsVisible=$($window.IsVisible)" "UI"
        } catch {
            Write-DebugLog "Kann Window-Status nicht abfragen: $($_.Exception.Message)" "UI"
        }
    }
}

Write-DebugLog "Anwendung geschlossen" "Shutdown"



# SIG # Begin signature block
# MIIoiQYJKoZIhvcNAQcCoIIoejCCKHYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAqtdYPfEQCi13f
# 63c2z08ujc7mDeqVD1w3krIumAwWe6CCILswggXJMIIEsaADAgECAhAbtY8lKt8j
# AEkoya49fu0nMA0GCSqGSIb3DQEBDAUAMH4xCzAJBgNVBAYTAlBMMSIwIAYDVQQK
# ExlVbml6ZXRvIFRlY2hub2xvZ2llcyBTLkEuMScwJQYDVQQLEx5DZXJ0dW0gQ2Vy
# dGlmaWNhdGlvbiBBdXRob3JpdHkxIjAgBgNVBAMTGUNlcnR1bSBUcnVzdGVkIE5l
# dHdvcmsgQ0EwHhcNMjEwNTMxMDY0MzA2WhcNMjkwOTE3MDY0MzA2WjCBgDELMAkG
# A1UEBhMCUEwxIjAgBgNVBAoTGVVuaXpldG8gVGVjaG5vbG9naWVzIFMuQS4xJzAl
# BgNVBAsTHkNlcnR1bSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTEkMCIGA1UEAxMb
# Q2VydHVtIFRydXN0ZWQgTmV0d29yayBDQSAyMIICIjANBgkqhkiG9w0BAQEFAAOC
# Ag8AMIICCgKCAgEAvfl4+ObVgAxknYYblmRnPyI6HnUBfe/7XGeMycxca6mR5rlC
# 5SBLm9qbe7mZXdmbgEvXhEArJ9PoujC7Pgkap0mV7ytAJMKXx6fumyXvqAoAl4Va
# qp3cKcniNQfrcE1K1sGzVrihQTib0fsxf4/gX+GxPw+OFklg1waNGPmqJhCrKtPQ
# 0WeNG0a+RzDVLnLRxWPa52N5RH5LYySJhi40PylMUosqp8DikSiJucBb+R3Z5yet
# /5oCl8HGUJKbAiy9qbk0WQq/hEr/3/6zn+vZnuCYI+yma3cWKtvMrTscpIfcRnNe
# GWJoRVfkkIJCu0LW8GHgwaM9ZqNd9BjuiMmNF0UpmTJ1AjHuKSbIawLmtWJFfzcV
# WiNoidQ+3k4nsPBADLxNF8tNorMe0AZa3faTz1d1mfX6hhpneLO/lv403L3nUlbl
# s+V1e9dBkQXcXWnjlQ1DufyDljmVe2yAWk8TcsbXfSl6RLpSpCrVQUYJIP4ioLZb
# MI28iQzV13D4h1L92u+sUS4Hs07+0AnacO+Y+lbmbdu1V0vc5SwlFcieLnhO+Nqc
# noYsylfzGuXIkosagpZ6w7xQEmnYDlpGizrrJvojybawgb5CAKT41v4wLsfSRvbl
# jnX98sy50IdbzAYQYLuDNbdeZ95H7JlI8aShFf6tjGKOOVVPORa5sWOd/7cCAwEA
# AaOCAT4wggE6MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFLahVDkCw6A/joq8
# +tT4HKbROg79MB8GA1UdIwQYMBaAFAh2zcsH/yT2xc3tu5C84oQ3RnX3MA4GA1Ud
# DwEB/wQEAwIBBjAvBgNVHR8EKDAmMCSgIqAghh5odHRwOi8vY3JsLmNlcnR1bS5w
# bC9jdG5jYS5jcmwwawYIKwYBBQUHAQEEXzBdMCgGCCsGAQUFBzABhhxodHRwOi8v
# c3ViY2Eub2NzcC1jZXJ0dW0uY29tMDEGCCsGAQUFBzAChiVodHRwOi8vcmVwb3Np
# dG9yeS5jZXJ0dW0ucGwvY3RuY2EuY2VyMDkGA1UdIAQyMDAwLgYEVR0gADAmMCQG
# CCsGAQUFBwIBFhhodHRwOi8vd3d3LmNlcnR1bS5wbC9DUFMwDQYJKoZIhvcNAQEM
# BQADggEBAFHCoVgWIhCL/IYx1MIy01z4S6Ivaj5N+KsIHu3V6PrnCA3st8YeDrJ1
# BXqxC/rXdGoABh+kzqrya33YEcARCNQOTWHFOqj6seHjmOriY/1B9ZN9DbxdkjuR
# mmW60F9MvkyNaAMQFtXx0ASKhTP5N+dbLiZpQjy6zbzUeulNndrnQ/tjUoCFBMQl
# lVXwfqefAcVbKPjgzoZwpic7Ofs4LphTZSJ1Ldf23SIikZbr3WjtP6MZl9M7JYjs
# NhI9qX7OAo0FmpKnJ25FspxihjcNpDOO16hO0EoXQ0zF8ads0h5YbBRRfopUofbv
# n3l6XYGaFpAP4bvxSgD5+d2+7arszgowggaDMIIEa6ADAgECAhEAnpwE9lWotKcC
# bUmMbHiNqjANBgkqhkiG9w0BAQwFADBWMQswCQYDVQQGEwJQTDEhMB8GA1UEChMY
# QXNzZWNvIERhdGEgU3lzdGVtcyBTLkEuMSQwIgYDVQQDExtDZXJ0dW0gVGltZXN0
# YW1waW5nIDIwMjEgQ0EwHhcNMjUwMTA5MDg0MDQzWhcNMzYwMTA3MDg0MDQzWjBQ
# MQswCQYDVQQGEwJQTDEhMB8GA1UECgwYQXNzZWNvIERhdGEgU3lzdGVtcyBTLkEu
# MR4wHAYDVQQDDBVDZXJ0dW0gVGltZXN0YW1wIDIwMjUwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDHKV9n+Kwr3ZBF5UCLWOQ/NdbblAvQeGMjfCi/bibT
# 71hPkwKV4UvQt1MuOwoaUCYtsLhw8jrmOmoz2HoHKKzEpiS3A1rA3ssXUZMnSrbi
# iVpDj+5MtnbXSVEJKbccuHbmwcjl39N4W72zccoC/neKAuwO1DJ+9SO+YkHncRiV
# 95idWhxRAcDYv47hc9GEFZtTFxQXLbrL4N7N90BqLle3ayznzccEPQ+E6H6p00zE
# 9HUp++3bZTF4PfyPRnKCLc5ezAzEqqbbU5F/nujx69T1mm02jltlFXnTMF1vlake
# QXWYpGIjtrR7WP7tIMZnk78nrYSfeAp8le+/W/5+qr7tqQZufW9invsRTcfk7P+m
# nKjJLuSbwqgxelvCBryz9r51bT0561aR2c+joFygqW7n4FPCnMLOj40X4ot7wP2u
# 8kLRDVHbhsHq5SGLqr8DbFq14ws2ALS3tYa2GGiA7wX79rS5oDMnSY/xmJO5cupu
# SvqpylzO7jzcLOwWiqCrq05AXp51SRrj9xRt8KdZWpDdWhWmE8MFiFtmQ0AqODLJ
# Bn1hQAx3FvD/pte6pE1Bil0BOVC2Snbeq/3NylDwvDdAg/0CZRJsQIaydHswJwyY
# BlYUDyaQK2yUS57hobnYx/vStMvTB96ii4jGV3UkZh3GvwdDCsZkbJXaU8ATF/z6
# DwIDAQABo4IBUDCCAUwwdQYIKwYBBQUHAQEEaTBnMDsGCCsGAQUFBzAChi9odHRw
# Oi8vc3ViY2EucmVwb3NpdG9yeS5jZXJ0dW0ucGwvY3RzY2EyMDIxLmNlcjAoBggr
# BgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2VydHVtLmNvbTAfBgNVHSMEGDAW
# gBS+VAIvv0Bsc0POrAklTp5DRBru4DAMBgNVHRMBAf8EAjAAMDkGA1UdHwQyMDAw
# LqAsoCqGKGh0dHA6Ly9zdWJjYS5jcmwuY2VydHVtLnBsL2N0c2NhMjAyMS5jcmww
# FgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQDAgeAMCIGA1UdIAQb
# MBkwCAYGZ4EMAQQCMA0GCyqEaAGG9ncCBQELMB0GA1UdDgQWBBSBjAagKFP8AD/b
# fp5KwR8i7LISiTANBgkqhkiG9w0BAQwFAAOCAgEAmQ8ZDBvrBUPnaL87AYc4Jlmf
# H1ZP5yt65MtzYu8fbmsL3d3cvYs+Enbtfu9f2wMehzSyved3Rc59a04O8NN7plw4
# PXg71wfSE4MRFM1EuqL63zq9uTjm/9tA73r1aCdWmkprKp0aLoZolUN0qGcvr9+Q
# G8VIJVMcuSqFeEvRrLEKK2xVkMSdTTbDhseUjI4vN+BrXm5z45EA3aDpSiZQuoNd
# 4RFnDzddbgfcCQPaY2UyXqzNBjnuz6AyHnFzKtNlCevkMBgh4dIDt/0DGGDOaTEA
# WZtUEqK5AlHd0PBnd40Lnog4UATU3Bt6GHfeDmWEHFTjHKsmn9Q8wiGj906bVgL8
# 35tfEH9EgYDklqrOUxWxDf1cOA7ds/r8pIc2vjLQ9tOSkm9WXVbnTeLG3Q57frTg
# CvTObd/qf3UzE97nTNOU7vOMZEo41AgmhuEbGsyQIDM/V6fJQX1RnzzJNoqfTTkU
# zUoP2tlNHnNsjFo2YV+5yZcoaawmNWmR7TywUXG2/vFgJaG0bfEoodeeXp7A4I4H
# aDDpfRa7ypgJEPeTwHuBRJpj9N+1xtri+6BzHPwsAAvUJm58PGoVsteHAXwvpg4N
# VgvUk3BKbl7xFulWU1KHqH/sk7T0CFBQ5ohuKPmFf1oqAP4AO9a3Yg2wBMwEg1zP
# Oh6xbUXskzs9iSa9yGwwgga5MIIEoaADAgECAhEAmaOACiZVO2Wr3G6EprPqOTAN
# BgkqhkiG9w0BAQwFADCBgDELMAkGA1UEBhMCUEwxIjAgBgNVBAoTGVVuaXpldG8g
# VGVjaG5vbG9naWVzIFMuQS4xJzAlBgNVBAsTHkNlcnR1bSBDZXJ0aWZpY2F0aW9u
# IEF1dGhvcml0eTEkMCIGA1UEAxMbQ2VydHVtIFRydXN0ZWQgTmV0d29yayBDQSAy
# MB4XDTIxMDUxOTA1MzIxOFoXDTM2MDUxODA1MzIxOFowVjELMAkGA1UEBhMCUEwx
# ITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5BLjEkMCIGA1UEAxMbQ2Vy
# dHVtIENvZGUgU2lnbmluZyAyMDIxIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAnSPPBDAjO8FGLOczcz5jXXp1ur5cTbq96y34vuTmflN4mSAfgLKT
# vggv24/rWiVGzGxT9YEASVMw1Aj8ewTS4IndU8s7VS5+djSoMcbvIKck6+hI1shs
# ylP4JyLvmxwLHtSworV9wmjhNd627h27a8RdrT1PH9ud0IF+njvMk2xqbNTIPsnW
# tw3E7DmDoUmDQiYi/ucJ42fcHqBkbbxYDB7SYOouu9Tj1yHIohzuC8KNqfcYf7Z4
# /iZgkBJ+UFNDcc6zokZ2uJIxWgPWXMEmhu1gMXgv8aGUsRdaCtVD2bSlbfsq7Biq
# ljjaCun+RJgTgFRCtsuAEw0pG9+FA+yQN9n/kZtMLK+Wo837Q4QOZgYqVWQ4x6cM
# 7/G0yswg1ElLlJj6NYKLw9EcBXE7TF3HybZtYvj9lDV2nT8mFSkcSkAExzd4prHw
# YjUXTeZIlVXqj+eaYqoMTpMrfh5MCAOIG5knN4Q/JHuurfTI5XDYO962WZayx7AC
# Ff5ydJpoEowSP07YaBiQ8nXpDkNrUA9g7qf/rCkKbWpQ5boufUnq1UiYPIAHlezf
# 4muJqxqIns/kqld6JVX8cixbd6PzkDpwZo4SlADaCi2JSplKShBSND36E/ENVv8u
# rPS0yOnpG4tIoBGxVCARPCg1BnyMJ4rBJAcOSnAWd18Jx5n858JSqPECAwEAAaOC
# AVUwggFRMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFN10XUwA23ufoHTKsW73
# PMAywHDNMB8GA1UdIwQYMBaAFLahVDkCw6A/joq8+tT4HKbROg79MA4GA1UdDwEB
# /wQEAwIBBjATBgNVHSUEDDAKBggrBgEFBQcDAzAwBgNVHR8EKTAnMCWgI6Ahhh9o
# dHRwOi8vY3JsLmNlcnR1bS5wbC9jdG5jYTIuY3JsMGwGCCsGAQUFBwEBBGAwXjAo
# BggrBgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2VydHVtLmNvbTAyBggrBgEF
# BQcwAoYmaHR0cDovL3JlcG9zaXRvcnkuY2VydHVtLnBsL2N0bmNhMi5jZXIwOQYD
# VR0gBDIwMDAuBgRVHSAAMCYwJAYIKwYBBQUHAgEWGGh0dHA6Ly93d3cuY2VydHVt
# LnBsL0NQUzANBgkqhkiG9w0BAQwFAAOCAgEAdYhYD+WPUCiaU58Q7EP89DttyZqG
# Yn2XRDhJkL6P+/T0IPZyxfxiXumYlARMgwRzLRUStJl490L94C9LGF3vjzzH8Jq3
# iR74BRlkO18J3zIdmCKQa5LyZ48IfICJTZVJeChDUyuQy6rGDxLUUAsO0eqeLNhL
# Vsgw6/zOfImNlARKn1FP7o0fTbj8ipNGxHBIutiRsWrhWM2f8pXdd3x2mbJCKKtl
# 2s42g9KUJHEIiLni9ByoqIUul4GblLQigO0ugh7bWRLDm0CdY9rNLqyA3ahe8Wlx
# VWkxyrQLjH8ItI17RdySaYayX3PhRSC4Am1/7mATwZWwSD+B7eMcZNhpn8zJ+6MT
# yE6YoEBSRVrs0zFFIHUR08Wk0ikSf+lIe5Iv6RY3/bFAEloMU+vUBfSouCReZwSL
# o8WdrDlPXtR0gicDnytO7eZ5827NS2x7gCBibESYkOh1/w1tVxTpV2Na3PR7nxYV
# lPu1JPoRZCbH86gc96UTvuWiOruWmyOEMLOGGniR+x+zPF/2DaGgK2W1eEJfo2qy
# rBNPvF7wuAyQfiFXLwvWHamoYtPZo0LHuH8X3n9C+xN4YaNjt2ywzOr+tKyEVAot
# nyU9vyEVOaIYMk3IeBrmFnn0gbKeTTyYeEEUz/Qwt4HOUBCrW602NCmvO1nm+/80
# nLy5r0AZvCQxaQ4wgga5MIIEoaADAgECAhEA5/9pxzs1zkuRJth0fGilhzANBgkq
# hkiG9w0BAQwFADCBgDELMAkGA1UEBhMCUEwxIjAgBgNVBAoTGVVuaXpldG8gVGVj
# aG5vbG9naWVzIFMuQS4xJzAlBgNVBAsTHkNlcnR1bSBDZXJ0aWZpY2F0aW9uIEF1
# dGhvcml0eTEkMCIGA1UEAxMbQ2VydHVtIFRydXN0ZWQgTmV0d29yayBDQSAyMB4X
# DTIxMDUxOTA1MzIwN1oXDTM2MDUxODA1MzIwN1owVjELMAkGA1UEBhMCUEwxITAf
# BgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5BLjEkMCIGA1UEAxMbQ2VydHVt
# IFRpbWVzdGFtcGluZyAyMDIxIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA6RIfBDXtuV16xaaVQb6KZX9Od9FtJXXTZo7b+GEof3+3g0ChWiKnO7R4
# +6MfrvLyLCWZa6GpFHjEt4t0/GiUQvnkLOBRdBqr5DOvlmTvJJs2X8ZmWgWJjC7P
# BZLYBWAs8sJl3kNXxBMX5XntjqWx1ZOuuXl0R4x+zGGSMzZ45dpvB8vLpQfZkfMC
# /1tL9KYyjU+htLH68dZJPtzhqLBVG+8ljZ1ZFilOKksS79epCeqFSeAUm2eMTGpO
# iS3gfLM6yvb8Bg6bxg5yglDGC9zbr4sB9ceIGRtCQF1N8dqTgM/dSViiUgJkcv5d
# LNJeWxGCqJYPgzKlYZTgDXfGIeZpEFmjBLwURP5ABsyKoFocMzdjrCiFbTvJn+bD
# 1kq78qZUgAQGGtd6zGJ88H4NPJ5Y2R4IargiWAmv8RyvWnHr/VA+2PrrK9eXe5q7
# M88YRdSTq9TKbqdnITUgZcjjm4ZUjteq8K331a4P0s2in0p3UubMEYa/G5w6jSWP
# UzchGLwWKYBfeSu6dIOC4LkeAPvmdZxSB1lWOb9HzVWZoM8Q/blaP4LWt6JxjkI9
# yQsYGMdCqwl7uMnPUIlcExS1mzXRxUowQref/EPaS7kYVaHHQrp4XB7nTEtQhkP0
# Z9Puz/n8zIFnUSnxDof4Yy650PAXSYmK2TcbyDoTNmmt8xAxzcMCAwEAAaOCAVUw
# ggFRMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFL5UAi+/QGxzQ86sCSVOnkNE
# Gu7gMB8GA1UdIwQYMBaAFLahVDkCw6A/joq8+tT4HKbROg79MA4GA1UdDwEB/wQE
# AwIBBjATBgNVHSUEDDAKBggrBgEFBQcDCDAwBgNVHR8EKTAnMCWgI6Ahhh9odHRw
# Oi8vY3JsLmNlcnR1bS5wbC9jdG5jYTIuY3JsMGwGCCsGAQUFBwEBBGAwXjAoBggr
# BgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2VydHVtLmNvbTAyBggrBgEFBQcw
# AoYmaHR0cDovL3JlcG9zaXRvcnkuY2VydHVtLnBsL2N0bmNhMi5jZXIwOQYDVR0g
# BDIwMDAuBgRVHSAAMCYwJAYIKwYBBQUHAgEWGGh0dHA6Ly93d3cuY2VydHVtLnBs
# L0NQUzANBgkqhkiG9w0BAQwFAAOCAgEAuJNZd8lMFf2UBwigp3qgLPBBk58BFCS3
# Q6aJDf3TISoytK0eal/JyCB88aUEd0wMNiEcNVMbK9j5Yht2whaknUE1G32k6uld
# 7wcxHmw67vUBY6pSp8QhdodY4SzRRaZWzyYlviUpyU4dXyhKhHSncYJfa1U75cXx
# Ce3sTp9uTBm3f8Bj8LkpjMUSVTtMJ6oEu5JqCYzRfc6nnoRUgwz/GVZFoOBGdrSE
# tDN7mZgcka/tS5MI47fALVvN5lZ2U8k7Dm/hTX8CWOw0uBZloZEW4HB0Xra3qE4q
# zzq/6M8gyoU/DE0k3+i7bYOrOk/7tPJg1sOhytOGUQ30PbG++0FfJioDuOFhj99b
# 151SqFlSaRQYz74y/P2XJP+cF19oqozmi0rRTkfyEJIvhIZ+M5XIFZttmVQgTxfp
# fJwMFFEoQrSrklOxpmSygppsUDJEoliC05vBLVQ+gMZyYaKvBJ4YxBMlKH5ZHkRd
# loRYlUDplk8GUa+OCMVhpDSQurU6K1ua5dmZftnvSSz2H96UrQDzA6DyiI1V3ejV
# tvn2azVAXg6NnjmuRZ+wa7Pxy0H3+V4K4rOTHlG3VYA6xfLsTunCz72T6Ot4+tkr
# DYOeaU1pPX1CBfYj6EW2+ELq46GP8KCNUQDirWLU4nOmgCat7vN0SD6RlwUiSsMe
# CiQDmZwgwrUwggbpMIIE0aADAgECAhBiOsZKIV2oSfsf25d4iu6HMA0GCSqGSIb3
# DQEBCwUAMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhBc3NlY28gRGF0YSBTeXN0
# ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2RlIFNpZ25pbmcgMjAyMSBDQTAe
# Fw0yNTA3MzExMTM4MDhaFw0yNjA3MzExMTM4MDdaMIGOMQswCQYDVQQGEwJERTEb
# MBkGA1UECAwSQmFkZW4tV8O8cnR0ZW1iZXJnMRQwEgYDVQQHDAtCYWllcnNicm9u
# bjEeMBwGA1UECgwVT3BlbiBTb3VyY2UgRGV2ZWxvcGVyMSwwKgYDVQQDDCNPcGVu
# IFNvdXJjZSBEZXZlbG9wZXIsIEhlcHAgQW5kcmVhczCCAiIwDQYJKoZIhvcNAQEB
# BQADggIPADCCAgoCggIBAOt2txKXx2UtfBNIw2kVihIAcgPkK3lp7np/qE0evLq2
# J/L5kx8m6dUY4WrrcXPSn1+W2/PVs/XBFV4fDfwczZnQ/hYzc8Ot5YxPKLx6hZxK
# C5v8LjNIZ3SRJvMbOpjzWoQH7MLIIj64n8mou+V0CMk8UElmU2d0nxBQyau1njQP
# CLvlfInu4tDndyp3P87V5bIdWw6MkZFhWDkILTYInYicYEkut5dN9hT02t/3rXu2
# 30DEZ6S1OQtm9loo8wzvwjRoVX3IxnfpCHGW8Z9ie9I9naMAOG2YpvpoUbLG3fL/
# B6JVNNR1mm/AYaqVMtAXJpRlqvbIZyepcG0YGB+kOQLdoQCWlIp3a14Z4kg6bU9C
# U1KNR4ueA+SqLNu0QGtgBAdTfqoWvyiaeyEogstBHglrZ39y/RW8OOa50pSleSRx
# SXiGW+yH+Ps5yrOopTQpKHy0kRincuJpYXgxGdGxxKHwuVJHKXL0nWScEku0C38p
# M9sYanIKncuF0Ed7RvyNqmPP5pt+p/0ZG+zLNu/Rce0LE5FjAIRtW2hFxmYMyohk
# afzyjCCCG0p2KFFT23CoUfXx59nCU+lyWx/iyDMV4sqrcvmZdPZF7lkaIb5B4PYP
# vFFE7enApz4Niycj1gPUFlx4qTcXHIbFLJDp0ry6MYelX+SiMHV7yDH/rnWXm5d3
# AgMBAAGjggF4MIIBdDAMBgNVHRMBAf8EAjAAMD0GA1UdHwQ2MDQwMqAwoC6GLGh0
# dHA6Ly9jY3NjYTIwMjEuY3JsLmNlcnR1bS5wbC9jY3NjYTIwMjEuY3JsMHMGCCsG
# AQUFBwEBBGcwZTAsBggrBgEFBQcwAYYgaHR0cDovL2Njc2NhMjAyMS5vY3NwLWNl
# cnR1bS5jb20wNQYIKwYBBQUHMAKGKWh0dHA6Ly9yZXBvc2l0b3J5LmNlcnR1bS5w
# bC9jY3NjYTIwMjEuY2VyMB8GA1UdIwQYMBaAFN10XUwA23ufoHTKsW73PMAywHDN
# MB0GA1UdDgQWBBQYl6R41hwxInb9JVvqbCTp9ILCcTBLBgNVHSAERDBCMAgGBmeB
# DAEEATA2BgsqhGgBhvZ3AgUBBDAnMCUGCCsGAQUFBwIBFhlodHRwczovL3d3dy5j
# ZXJ0dW0ucGwvQ1BTMBMGA1UdJQQMMAoGCCsGAQUFBwMDMA4GA1UdDwEB/wQEAwIH
# gDANBgkqhkiG9w0BAQsFAAOCAgEAQ4guyo7zysB7MHMBOVKKY72rdY5hrlxPci8u
# 1RgBZ9ZDGFzhnUM7iIivieAeAYLVxP922V3ag9sDVNR+mzCmu1pWCgZyBbNXykue
# KJwOfE8VdpmC/F7637i8a7Pyq6qPbcfvLSqiXtVrT4NX4NIvODW3kIqf4nGwd0h3
# 1tuJVHLkdpGmT0q4TW0gAxnNoQ+lO8uNzCrtOBk+4e1/3CZXSDnjR8SUsHrHdhnm
# qkAnYb40vf69dfDR148tToUj872yYeBUEGUsQUDgJ6HSkMVpLQz/Nb3xy9qkY33M
# 7CBWKuBVwEcbGig/yj7CABhIrY1XwRddYQhEyozUS4mXNqXydAD6Ylt143qrECD2
# s3MDQBgP2sbRHdhVgzr9+n1iztXkPHpIlnnXPkZrt89E5iGL+1PtjETrhTkr7nxj
# yMFjrbmJ8W/XglwopUTCGfopDFPlzaoFf5rH/v3uzS24yb6+dwQrvCwFA9Y9ZHy2
# ITJx7/Ll6AxWt7Lz9JCJ5xRyYeRUHs6ycB8EuMPAKyGpzdGtjWv2rkTXbkIYUjkl
# FTpquXJBc/kO5L+Quu0a0uKn4ea16SkABy052XHQqd87cSJg3rGxsagi0IAfxGM6
# 08oupufSS/q9mpQPgkDuMJ8/zdre0st8OduAoG131W+XJ7mm0gIuh2zNmSIet5RD
# oa8THmwxggckMIIHIAIBATBqMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhBc3Nl
# Y28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2RlIFNpZ25p
# bmcgMjAyMSBDQQIQYjrGSiFdqEn7H9uXeIruhzANBglghkgBZQMEAgEFAKCBhDAY
# BgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEi
# BCAbtvpQA4wC+Fwb0cbVCz62BhfdIRwnKORSRlgwGARv4DANBgkqhkiG9w0BAQEF
# AASCAgDENaN6iq0nFwmHbVa9TrQWkaa0YJC5W2BjhA1rBiv7QeoA8uLuetIHgWef
# 9tYkhdGYH52dKY7yNKTkKx5IAFYGSqjKdC0630+Wneijtr3GEEcu2/PmoVBFhaEg
# QlyUFnPQ0j8M/LwE9+iyqN2+RmrbnPm1giXII9raalU2jzHRf6QyLiN4tlzTQdzo
# eBxRRwha1HATVT0P3H0Y9fp10Txi+Xkqy7h458fT+M7HMd+5iJkJZUYXg1xgDoaM
# x7Z7CLvOvJowvMtHHPyJjrAEQVktvt5HU6L39srux89F7jkdfdN63Mwr9p+TeM6c
# 7xkXJip8/UjzmwEWq1yzPcXmTKr8/VAICpKbxNsHZN2og2naVaAvKd/4I60hHr4z
# PumVPqbtshkEYr3S2UEMnkSIensSrWLM3GgTOPgcrfLjqF0A10ZR37/IXP/c5YBT
# M5DczXeS2YMI2wkcvj+N57gLASMtbTser1fR3zYveBVoGLUaxl5PvqKi90UqXDhT
# 49hrhT+34L6LVWlrRlu0eeAKBVwdNTPPPNrxJ4s8J1UV2xwUZI6M1t5hg6U6HgBU
# 6eoy68pPKi+7/iJCuKnv33eBQl9wQ1JZBCm/mXJVXeCsF4h2BO/9m4nmis6G12nF
# RpAAUi/c/GsN9RIuLgXJCaBhWunI6+6l0GS4uVzaApun2nXM56GCBAQwggQABgkq
# hkiG9w0BCQYxggPxMIID7QIBATBrMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhB
# c3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBUaW1lc3Rh
# bXBpbmcgMjAyMSBDQQIRAJ6cBPZVqLSnAm1JjGx4jaowDQYJYIZIAWUDBAICBQCg
# ggFXMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcN
# MjUxMDI1MTgwMDI1WjA3BgsqhkiG9w0BCRACLzEoMCYwJDAiBCDPodw1ne0rw8uJ
# D6Iw5dr3e1QPGm4rI93PF1ThjPqg1TA/BgkqhkiG9w0BCQQxMgQwOSQOO2fpu7ZT
# AdG538S+N1Wmv5iHD+idVAoKP/XefX+FMx6w/XyFfE7PNO/BwWeQMIGgBgsqhkiG
# 9w0BCRACDDGBkDCBjTCBijCBhwQUwyW4mxf8xQJgYc4rcXtFB92camowbzBapFgw
# VjELMAkGA1UEBhMCUEwxITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5B
# LjEkMCIGA1UEAxMbQ2VydHVtIFRpbWVzdGFtcGluZyAyMDIxIENBAhEAnpwE9lWo
# tKcCbUmMbHiNqjANBgkqhkiG9w0BAQEFAASCAgBanYdMKMN7taRrxt37VY3zYJaI
# XFj2/gbm3+Sk45B81ZN4r/QLzMiO3XuX/ubegMXB1pSOqXBJY4i7QlirtiE0uXZi
# /IeycsHj8bE29dqv2hCphGiXQiD/VoTdQhlLZBQ/zK8FuCZ0je/9/LkVBFMlML6G
# e8cgoxkr/MHt44pIWxAUar04Z6mbr/t93uPwvPCtnDaghEfyK2b0/0Abx9H9gEy/
# /E6lKwjG6i1SGcRCV9EoWLHOPu41wdzs1/RLBd5z/2OYogO9kRN6iePQUm2lRnDH
# 9251zrZLhdY7/VbYIiJQ6fMDTTVxBPZk/1s1aM1LyDKdyKarEHBiut2EkmabqlZD
# 7FFjt11+K7ATPPgEXXWPfHAhcMio4C+pEVJOgPTcTvytiUrUP+aFe36c4IcjbB7Y
# LnkV8kt/DwGx3OflWKJxlwx/Vzcxetp1+zvRERS4iD6FS9FW4SoxkRz6b/x86s+m
# SHAhWFqXmNsXpEXL7Wg6ddZH1SQFiR/l3fBPiwnb6o2Bp77NriR8pO5T4+biKTCP
# yBhcZQws1vSdlsDqhUZGEErHiWSAgStUy6RT77ReHECvABErZ9JqiQccgF85t/B4
# bpIDdQFrZFpoem0kNiNSl9YNbS2PSY2fY8wta+1JRNAfkJgjoAj2fAKM+VTQ1c1Z
# 4Y4BrPBrvHcxB7MsEQ==
# SIG # End signature block
