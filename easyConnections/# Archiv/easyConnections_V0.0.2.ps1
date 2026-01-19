# easyConnections_V0.0.2.ps1
# Advanced Network Connections Monitor with WPF GUI - Admin Edition
# Enhanced with comprehensive protocol support for system administrators

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms # Für SaveFileDialog

# XAML Content
$xamlContent = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="easyConnections - Network Connections Monitor"
    Width="1100"
    Height="750"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize"
    MinWidth="1100"
    MinHeight="600"
    Background="#F3F3F3"
    WindowStyle="SingleBorderWindow">

    <Window.Resources>
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                BorderBrush="{TemplateBinding BorderBrush}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#106EBE"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#005A9E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
            <Button x:Name="btnRefresh" Content="Aktualisieren" Style="{StaticResource ModernButton}" Margin="0,0,10,0"/>
            <ComboBox x:Name="cmbProtocol" Width="180" Margin="0,0,10,0">
                <ComboBoxItem Content="Alle" IsSelected="True"/>
                <ComboBoxItem Content="HTTP"/>
                <ComboBoxItem Content="HTTPS"/>
                <ComboBoxItem Content="FTP"/>
                <ComboBoxItem Content="FTPS"/>
                <ComboBoxItem Content="SMTP"/>
                <ComboBoxItem Content="POP3"/>
                <ComboBoxItem Content="IMAP"/>
                <ComboBoxItem Content="DNS"/>
                <ComboBoxItem Content="DHCP"/>
                <ComboBoxItem Content="LDAP"/>
                <ComboBoxItem Content="LDAPS"/>
                <ComboBoxItem Content="SSH"/>
                <ComboBoxItem Content="Telnet"/>
                <ComboBoxItem Content="RDP"/>
                <ComboBoxItem Content="SMB/CIFS"/>
                <ComboBoxItem Content="NFS"/>
                <ComboBoxItem Content="SNMP"/>
                <ComboBoxItem Content="NTP"/>
                <ComboBoxItem Content="Kerberos"/>
                <ComboBoxItem Content="WinRM"/>
                <ComboBoxItem Content="PowerShell"/>
                <ComboBoxItem Content="SQL Server"/>
                <ComboBoxItem Content="MySQL"/>
                <ComboBoxItem Content="PostgreSQL"/>
                <ComboBoxItem Content="Oracle"/>
                <ComboBoxItem Content="MongoDB"/>
                <ComboBoxItem Content="Redis"/>
                <ComboBoxItem Content="VNC"/>
                <ComboBoxItem Content="TeamViewer"/>
                <ComboBoxItem Content="Syslog"/>
                <ComboBoxItem Content="RADIUS"/>
                <ComboBoxItem Content="TACACS+"/>
                <ComboBoxItem Content="NetBIOS"/>
                <ComboBoxItem Content="AD Replication"/>
                <ComboBoxItem Content="Exchange"/>
                <ComboBoxItem Content="Hyper-V"/>
                <ComboBoxItem Content="VMware"/>
            </ComboBox>
            <ComboBox x:Name="cmbDirection" Width="100" Margin="0,0,10,0">
                <ComboBoxItem Content="Beide" IsSelected="True"/>
                <ComboBoxItem Content="Eingehend"/>
                <ComboBoxItem Content="Ausgehend"/>
            </ComboBox>
            <ComboBox x:Name="cmbCategory" Width="125" Margin="0,0,10,0">
                <ComboBoxItem Content="Alle Kategorien" IsSelected="True"/>
                <ComboBoxItem Content="Web Services"/>
                <ComboBoxItem Content="Email Services"/>
                <ComboBoxItem Content="Database Services"/>
                <ComboBoxItem Content="Directory Services"/>
                <ComboBoxItem Content="Remote Access"/>
                <ComboBoxItem Content="File Services"/>
                <ComboBoxItem Content="Monitoring"/>
                <ComboBoxItem Content="Virtualization"/>
            </ComboBox>
            <Button x:Name="btnCapturePackets" Content="Pakete erfassen" Style="{StaticResource ModernButton}" Margin="0,0,10,0"/>
            <Button x:Name="btnStartCapture" Content="Erfassung starten" Style="{StaticResource ModernButton}" Margin="0,0,10,0"/>
            <Button x:Name="btnHideStatus" Content="Status ausblenden" Style="{StaticResource ModernButton}" Margin="0,0,10,0"/>
            <Button x:Name="btnExportHTML" Content="HTML Export" Style="{StaticResource ModernButton}"/>
        </StackPanel>

        <DataGrid x:Name="dgConnections" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Typ" Binding="{Binding Type}" Width="60"/>
                <DataGridTextColumn Header="Lokale Adresse" Binding="{Binding LocalAddress}" Width="*"/>
                <DataGridTextColumn Header="Lokaler Port" Binding="{Binding LocalPort}" Width="80"/>
                <DataGridTextColumn Header="Remote Adresse" Binding="{Binding RemoteAddress}" Width="*"/>
                <DataGridTextColumn Header="Remote Port" Binding="{Binding RemotePort}" Width="80"/>
                <DataGridTextColumn Header="Remote Hostname" Binding="{Binding RemoteHostname}" Width="*"/>
                <DataGridTextColumn Header="Status" Binding="{Binding State}" Width="80"/>
                <DataGridTextColumn Header="Prozess ID" Binding="{Binding OwningProcess}" Width="80"/>
                <DataGridTextColumn Header="Prozess Name" Binding="{Binding ProcessName}" Width="120"/>
            </DataGrid.Columns>
        </DataGrid>

        <TextBox x:Name="txtPackets" Grid.Row="2" Height="120" Margin="0,10,0,10" IsReadOnly="True" VerticalScrollBarVisibility="Auto" AcceptsReturn="True" Visibility="Collapsed"/>

        <!-- Footer -->
        <Border Grid.Row="3" Background="#F0F0F0" BorderBrush="#CCCCCC" BorderThickness="0,1,0,0" Padding="10,5">
            <StackPanel Orientation="Vertical" HorizontalAlignment="Center">
                <TextBlock Text="easyConnections V0.0.2 - Network Connections Monitor (Admin Edition)" FontSize="12" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                <TextBlock FontSize="11" HorizontalAlignment="Center" Margin="0,2,0,0">
                    <Run Text="Copyright 2025 | Autor: Andreas Hepp | "/>
                    <Hyperlink x:Name="linkPhinit" NavigateUri="https://www.phinit.de" Foreground="#0078D4">www.phinit.de</Hyperlink>
                    <Run Text=" / "/>
                    <Hyperlink x:Name="linkPsscripts" NavigateUri="https://www.psscripts.de" Foreground="#0078D4">www.psscripts.de</Hyperlink>
                </TextBlock>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xamlContent)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$btnRefresh = $window.FindName("btnRefresh")
$cmbProtocol = $window.FindName("cmbProtocol")
$cmbDirection = $window.FindName("cmbDirection")
$cmbCategory = $window.FindName("cmbCategory")
$btnCapturePackets = $window.FindName("btnCapturePackets")
$btnStartCapture = $window.FindName("btnStartCapture")
$btnHideStatus = $window.FindName("btnHideStatus")
$btnExportHTML = $window.FindName("btnExportHTML")
$dgConnections = $window.FindName("dgConnections")
$txtPackets = $window.FindName("txtPackets")
$linkPhinit = $window.FindName("linkPhinit")
$linkPsscripts = $window.FindName("linkPsscripts")

# Global variables for continuous capture
$Global:CaptureTimer = $null
$Global:CapturedConnections = @()
$Global:CapturedConnectionsHash = @{}
$Global:CaptureActive = $false
$Global:CategoryTimer = $null
$Global:ProcessCache = @{}

# Protocol port mappings for Admin use
$protocolPorts = @{
    "HTTP" = @(80, 8080, 8008, 8000)
    "HTTPS" = @(443, 8443, 8444)
    "FTP" = @(21, 20)
    "FTPS" = @(990, 989)
    "SMTP" = @(25, 587, 465)
    "POP3" = @(110, 995)
    "IMAP" = @(143, 993)
    "DNS" = @(53)
    "DHCP" = @(67, 68)
    "LDAP" = @(389)
    "LDAPS" = @(636)
    "SSH" = @(22)
    "Telnet" = @(23)
    "RDP" = @(3389)
    "SMB/CIFS" = @(445, 139, 137, 138)
    "NFS" = @(2049, 111)
    "SNMP" = @(161, 162)
    "NTP" = @(123)
    "Kerberos" = @(88, 464)
    "WinRM" = @(5985, 5986)
    "PowerShell" = @(5985, 5986)
    "SQL Server" = @(1433, 1434)
    "MySQL" = @(3306)
    "PostgreSQL" = @(5432)
    "Oracle" = @(1521, 1522)
    "MongoDB" = @(27017, 27018, 27019)
    "Redis" = @(6379)
    "VNC" = @(5900, 5901, 5902, 5903, 5904, 5905)
    "TeamViewer" = @(5938)
    "Syslog" = @(514)
    "RADIUS" = @(1812, 1813)
    "TACACS+" = @(49)
    "NetBIOS" = @(137, 138, 139)
    "AD Replication" = @(135, 389, 636, 3268, 3269, 53, 88, 445)
    "Exchange" = @(25, 80, 110, 143, 443, 993, 995, 587, 465)
    "Hyper-V" = @(2179, 5985, 5986)
    "VMware" = @(443, 902, 903, 8080, 8443)
}

# Protocol categories for better organization
$protocolCategories = @{
    "Web Services" = @("HTTP", "HTTPS")
    "Email Services" = @("SMTP", "POP3", "IMAP", "Exchange")
    "Database Services" = @("SQL Server", "MySQL", "PostgreSQL", "Oracle", "MongoDB", "Redis")
    "Directory Services" = @("LDAP", "LDAPS", "AD Replication", "Kerberos")
    "Remote Access" = @("SSH", "Telnet", "RDP", "VNC", "TeamViewer", "WinRM", "PowerShell")
    "File Services" = @("FTP", "FTPS", "SMB/CIFS", "NFS", "NetBIOS")
    "Monitoring" = @("SNMP", "Syslog", "NTP")
    "Virtualization" = @("Hyper-V", "VMware")
}



# Function to get process name with caching
function Get-ProcessNameCached {
    param([int]$ProcessId)
    
    if ($Global:ProcessCache.ContainsKey($ProcessId)) {
        return $Global:ProcessCache[$ProcessId]
    }
    
    try {
        $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
        $processName = if ($process) { $process.ProcessName } else { "Unknown" }
        $Global:ProcessCache[$ProcessId] = $processName
        return $processName
    } catch {
        $Global:ProcessCache[$ProcessId] = "Unknown"
        return "Unknown"
    }
}

# Function to get connections
function Get-NetworkConnections {
    param(
        [string]$protocol = "Alle", 
        [string]$direction = "Beide",
        [string]$category = "Alle Kategorien"
    )

    # Get TCP connections (faster without process names initially)
    $tcpConnections = Get-NetTCPConnection | Select-Object @{Name="Type";Expression={"TCP"}}, LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess, @{Name="ProcessName";Expression={""}}, @{Name="RemoteHostname";Expression={""}}

    # Get UDP endpoints (faster without process names initially)
    $udpConnections = Get-NetUDPEndpoint | Select-Object @{Name="Type";Expression={"UDP"}}, @{Name="LocalAddress";Expression={$_.LocalAddress}}, @{Name="LocalPort";Expression={$_.LocalPort}}, @{Name="RemoteAddress";Expression={"N/A"}}, @{Name="RemotePort";Expression={"N/A"}}, @{Name="State";Expression={"Listen"}}, @{Name="OwningProcess";Expression={$_.OwningProcess}}, @{Name="ProcessName";Expression={""}}, @{Name="RemoteHostname";Expression={"N/A"}}

    $connections = $tcpConnections + $udpConnections

    # Filter by category first
    if ($category -ne "Alle Kategorien") {
        $categoryProtocols = $protocolCategories[$category]
        $categoryPorts = @()
        foreach ($prot in $categoryProtocols) {
            if ($protocolPorts.ContainsKey($prot)) {
                $categoryPorts += $protocolPorts[$prot]
            }
        }
        $connections = $connections | Where-Object { 
            ($_.LocalPort -in $categoryPorts) -or 
            ($_.RemotePort -in $categoryPorts -and $_.Type -eq "TCP")
        }
    }

    # Filter by specific protocol
    if ($protocol -ne "Alle") {
        $ports = $protocolPorts[$protocol]
        if ($direction -eq "Eingehend") {
            $connections = $connections | Where-Object { $_.LocalPort -in $ports }
        } elseif ($direction -eq "Ausgehend") {
            $connections = $connections | Where-Object { $_.RemotePort -in $ports -and $_.Type -eq "TCP" }
        } else { # Beide
            $connections = $connections | Where-Object { ($_.LocalPort -in $ports -or ($_.RemotePort -in $ports -and $_.Type -eq "TCP")) }
        }
    }

    # Only resolve process names and hostnames for filtered results (much faster)
    foreach ($conn in $connections) {
        if (-not $conn.ProcessName) {
            $conn.ProcessName = Get-ProcessNameCached -ProcessId $conn.OwningProcess
        }
        # Skip hostname resolution for now to improve performance
        # Only resolve for localhost and known IPs
        if ($conn.Type -eq "TCP" -and $conn.RemoteAddress -ne "N/A" -and -not $conn.RemoteHostname) {
            if ($conn.RemoteAddress -match "^127\.|^192\.168\.|^10\.|^172\.(1[6-9]|2[0-9]|3[01])\.") {
                $conn.RemoteHostname = Get-HostnameFast -ip $conn.RemoteAddress
            } else {
                $conn.RemoteHostname = $conn.RemoteAddress
            }
        }
    }

    return $connections
}

# Fast hostname resolution (simplified)
function Get-HostnameFast {
    param([string]$ip)
    try {
        if ($ip -eq "0.0.0.0" -or $ip -eq "::" -or $ip -eq "127.0.0.1" -or $ip -eq "::1" -or $ip -eq "N/A") {
            return $ip
        }
        # Simple DNS lookup without jobs (faster for WPF)
        $hostEntry = [System.Net.Dns]::GetHostEntry($ip)
        return $hostEntry.HostName
    } catch {
        return $ip
    }
}

# Function to export captured data as HTML
function Export-CapturedDataHTML {
    $selectedProtocol = $cmbProtocol.SelectedItem.Content
    $selectedDirection = $cmbDirection.SelectedItem.Content
    $selectedCategory = $cmbCategory.SelectedItem.Content
    
    $DateTimeNow = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $totalConnections = $Global:CapturedConnections.Count
    $tcpCount = ($Global:CapturedConnections | Where-Object { $_.Type -eq "TCP" }).Count
    $udpCount = ($Global:CapturedConnections | Where-Object { $_.Type -eq "UDP" }).Count
    $computerName = $env:COMPUTERNAME
    $userName = $env:USERNAME

    $html = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Kontinuierliche Netzwerk-Erfassung - easyConnections</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f3f3f3; color: #333; }
        .container { max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 0 20px rgba(0,0,0,0.1); }
        .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #0078D4; padding-bottom: 20px; }
        h1 { color: #0078D4; margin: 0; font-size: 2.5em; }
        .subtitle { color: #666; font-size: 1.2em; margin: 10px 0; }
        .info-box { background-color: #e6f3ff; padding: 15px; border-radius: 6px; margin-bottom: 20px; border-left: 4px solid #0078D4; }
        .summary { display: flex; justify-content: space-around; margin-bottom: 20px; }
        .summary-item { text-align: center; background: #f8f9fa; padding: 10px; border-radius: 4px; }
        .timestamp { text-align: center; color: #666; margin-bottom: 20px; font-style: italic; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 14px; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #0078D4; color: white; font-weight: bold; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        tr:hover { background-color: #e6f3ff; }
        .footer { text-align: center; margin-top: 30px; color: #666; border-top: 1px solid #ddd; padding-top: 20px; }
        .footer p { margin: 5px 0; }
        .footer a { color: #0078D4; text-decoration: none; }
        .footer a:hover { text-decoration: underline; }
        .capture-info { background-color: #fff3cd; padding: 10px; border-radius: 4px; margin-bottom: 20px; border-left: 4px solid #ffc107; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Kontinuierliche Netzwerk-Erfassung</h1>
            <div class="subtitle">easyConnections Tool</div>
            <div class="info-box">
                <strong>Computer:</strong> $computerName | <strong>Benutzer:</strong> $userName<br>
                <strong>Filter:</strong> Kategorie: $selectedCategory | Protokoll: $selectedProtocol | Richtung: $selectedDirection
            </div>
            <div class="capture-info">
                <strong>Erfassungszeitraum:</strong> Kontinuierliche Überwachung bis $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')
            </div>
        </div>
        <div class="summary">
            <div class="summary-item">
                <strong>Gesamt erfasst</strong><br>$totalConnections
            </div>
            <div class="summary-item">
                <strong>TCP</strong><br>$tcpCount
            </div>
            <div class="summary-item">
                <strong>UDP</strong><br>$udpCount
            </div>
        </div>
        <div class="timestamp">Erstellt am $DateTimeNow</div>
        <table>
            <thead>
                <tr>
                    <th>Typ</th>
                    <th>Lokale Adresse</th>
                    <th>Lokaler Port</th>
                    <th>Remote Adresse</th>
                    <th>Remote Port</th>
                    <th>Remote Hostname</th>
                    <th>Status</th>
                    <th>Prozess ID</th>
                    <th>Prozess Name</th>
                    <th>Erfasst um</th>
                </tr>
            </thead>
            <tbody>
"@

    foreach ($conn in $Global:CapturedConnections) {
        $captureTime = if ($conn.CaptureTime) { $conn.CaptureTime.ToString("dd.MM.yyyy HH:mm:ss") } else { "N/A" }
        $processName = if ($conn.ProcessName) { $conn.ProcessName } else { "N/A" }
        $html += "<tr><td>$($conn.Type)</td><td>$($conn.LocalAddress)</td><td>$($conn.LocalPort)</td><td>$($conn.RemoteAddress)</td><td>$($conn.RemotePort)</td><td>$($conn.RemoteHostname)</td><td>$($conn.State)</td><td>$($conn.OwningProcess)</td><td>$processName</td><td>$captureTime</td></tr>"
    }

    $html += @"
            </tbody>
        </table>
    </div>
    <div class="footer">
        <p>Copyright - $(Get-Date -Format "yyyy") | Autor: Andreas Hepp | Webseite: <a href="https://www.phinit.de">www.phinit.de</a> / <a href="https://www.psscripts.de">www.psscripts.de</a></p>
        <p>Generiert mit easyConnections Tool - Kontinuierliche Erfassung</p>
    </div>
</body>
</html>
"@

    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "HTML files (*.html)|*.html"
    $saveFileDialog.Title = "Kontinuierliche Erfassung Report speichern"
    $saveFileDialog.FileName = "ContinuousCaptureReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $html | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
        [System.Windows.MessageBox]::Show("Kontinuierliche Erfassung Report gespeichert: $($saveFileDialog.FileName)", "Erfolg", "OK", "Information")
    }
}

# Function to capture connections continuously
function Start-ContinuousCapture {
    $Global:CaptureActive = $true
    $btnStartCapture.Content = "Erfassung stoppen"
    $btnStartCapture.Background = [System.Windows.Media.Brushes]::Red
    $Global:CapturedConnections = @()
    $Global:CapturedConnectionsHash = @{}
    $txtPackets.Visibility = [System.Windows.Visibility]::Visible
    
    # Create timer for continuous capture
    $Global:CaptureTimer = New-Object System.Windows.Threading.DispatcherTimer
    $Global:CaptureTimer.Interval = [System.TimeSpan]::FromSeconds(5)
    
    $Global:CaptureTimer.Add_Tick({
        $selectedProtocol = $cmbProtocol.SelectedItem.Content
        $selectedDirection = $cmbDirection.SelectedItem.Content
        $selectedCategory = $cmbCategory.SelectedItem.Content
        $connections = Get-NetworkConnections -protocol $selectedProtocol -direction $selectedDirection -category $selectedCategory
        
        # Use hashtable for faster duplicate checking
        if (-not $Global:CapturedConnectionsHash) {
            $Global:CapturedConnectionsHash = @{}
        }
        
        foreach ($conn in $connections) {
            $conn | Add-Member -MemberType NoteProperty -Name "CaptureTime" -Value (Get-Date) -Force
            
            # Create unique key for connection
            $key = "$($conn.LocalAddress):$($conn.LocalPort)-$($conn.RemoteAddress):$($conn.RemotePort)-$($conn.Type)"
            
            # Check if connection already exists (faster with hashtable)
            if (-not $Global:CapturedConnectionsHash.ContainsKey($key)) {
                $Global:CapturedConnections += $conn
                $Global:CapturedConnectionsHash[$key] = $true
            }
        }
        
        $txtPackets.Text = "Erfassung läuft... Bisher $($Global:CapturedConnections.Count) einzigartige Verbindungen erfasst.`nLetztes Update: $(Get-Date -Format 'HH:mm:ss')"
    })
    
    $Global:CaptureTimer.Start()
    $txtPackets.Text = "Kontinuierliche Erfassung gestartet für Protokoll: $($cmbProtocol.SelectedItem.Content)"
}

function Stop-ContinuousCapture {
    $Global:CaptureActive = $false
    $btnStartCapture.Content = "Erfassung starten"
    $btnStartCapture.Background = [System.Windows.Media.Brushes]::Green
    
    if ($Global:CaptureTimer) {
        $Global:CaptureTimer.Stop()
        $Global:CaptureTimer = $null
    }
    
    $txtPackets.Text = "Erfassung beendet. $($Global:CapturedConnections.Count) Verbindungen erfasst."
    # Keep txtPackets visible to show the results
    
    # Auto export captured data
    if ($Global:CapturedConnections.Count -gt 0) {
        Export-CapturedDataHTML
    }
}

# Function to export HTML report
function Export-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]$connections,
        [Parameter(Mandatory=$true)][string]$protocol,
        [Parameter(Mandatory=$true)][string]$direction,
        [Parameter(Mandatory=$false)][string]$category = "Alle Kategorien"
    )

    $DateTimeNow = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $totalConnections = $connections.Count
    $tcpCount = ($connections | Where-Object { $_.Type -eq "TCP" }).Count
    $udpCount = ($connections | Where-Object { $_.Type -eq "UDP" }).Count
    $computerName = $env:COMPUTERNAME
    $userName = $env:USERNAME

    $html = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Network Connections Report - easyConnections</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f3f3f3; color: #333; }
        .container { max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 0 20px rgba(0,0,0,0.1); }
        .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #0078D4; padding-bottom: 20px; }
        h1 { color: #0078D4; margin: 0; font-size: 2.5em; }
        .subtitle { color: #666; font-size: 1.2em; margin: 10px 0; }
        .info-box { background-color: #e6f3ff; padding: 15px; border-radius: 6px; margin-bottom: 20px; border-left: 4px solid #0078D4; }
        .summary { display: flex; justify-content: space-around; margin-bottom: 20px; }
        .summary-item { text-align: center; background: #f8f9fa; padding: 10px; border-radius: 4px; }
        .timestamp { text-align: center; color: #666; margin-bottom: 20px; font-style: italic; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 14px; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #0078D4; color: white; font-weight: bold; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        tr:hover { background-color: #e6f3ff; }
        .footer { text-align: center; margin-top: 30px; color: #666; border-top: 1px solid #ddd; padding-top: 20px; }
        .footer p { margin: 5px 0; }
        .footer a { color: #0078D4; text-decoration: none; }
        .footer a:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Network Connections Report</h1>
            <div class="subtitle">easyConnections Tool</div>
            <div class="info-box">
                <strong>Computer:</strong> $computerName | <strong>Benutzer:</strong> $userName<br>
                <strong>Filter:</strong> Kategorie: $category | Protokoll: $protocol | Richtung: $direction
            </div>
        </div>
        <div class="summary">
            <div class="summary-item">
                <strong>Gesamt</strong><br>$totalConnections
            </div>
            <div class="summary-item">
                <strong>TCP</strong><br>$tcpCount
            </div>
            <div class="summary-item">
                <strong>UDP</strong><br>$udpCount
            </div>
        </div>
        <div class="timestamp">Erstellt am $DateTimeNow</div>
        <table>
            <thead>
                <tr>
                    <th>Typ</th>
                    <th>Lokale Adresse</th>
                    <th>Lokaler Port</th>
                    <th>Remote Adresse</th>
                    <th>Remote Port</th>
                    <th>Remote Hostname</th>
                    <th>Status</th>
                    <th>Prozess ID</th>
                    <th>Prozess Name</th>
                </tr>
            </thead>
            <tbody>
"@

    foreach ($conn in $connections) {
        $processName = if ($conn.ProcessName) { $conn.ProcessName } else { "N/A" }
        $html += "<tr><td>$($conn.Type)</td><td>$($conn.LocalAddress)</td><td>$($conn.LocalPort)</td><td>$($conn.RemoteAddress)</td><td>$($conn.RemotePort)</td><td>$($conn.RemoteHostname)</td><td>$($conn.State)</td><td>$($conn.OwningProcess)</td><td>$processName</td></tr>"
    }

    $html += @"
            </tbody>
        </table>
    </div>
    <div class="footer">
        <p>Copyright © $(Get-Date -Format "yyyy") | Autor: Andreas Hepp | Webseite: <a href="https://www.phinit.de">www.phinit.de</a> / <a href="https://www.psscripts.de">www.psscripts.de</a></p>
        <p>Generiert mit easyConnections Tool</p>
    </div>
</body>
</html>
"@

    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "HTML files (*.html)|*.html"
    $saveFileDialog.Title = "HTML Report speichern"
    $saveFileDialog.FileName = "NetworkConnectionsReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $html | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
        [System.Windows.MessageBox]::Show("HTML Report gespeichert: $($saveFileDialog.FileName)", "Erfolg", "OK", "Information")
    }
}

# Event handlers
$btnRefresh.Add_Click({
    try {
        # Show loading status
        $originalText = $btnRefresh.Content
        $btnRefresh.Content = "Lädt..."
        $btnRefresh.IsEnabled = $false
        
        $selectedProtocol = $cmbProtocol.SelectedItem.Content
        $selectedDirection = $cmbDirection.SelectedItem.Content
        $selectedCategory = $cmbCategory.SelectedItem.Content
        $connections = Get-NetworkConnections -protocol $selectedProtocol -direction $selectedDirection -category $selectedCategory
        $dgConnections.ItemsSource = $connections
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Laden der Verbindungen: $($_.Exception.Message)", "Fehler", "OK", "Error")
    } finally {
        # Restore button
        $btnRefresh.Content = $originalText
        $btnRefresh.IsEnabled = $true
    }
})

# Category selection event (optimized)
$cmbCategory.Add_SelectionChanged({
    if ($cmbCategory.SelectedItem) {
        $selectedCategory = $cmbCategory.SelectedItem.Content
        if ($selectedCategory -ne "Alle Kategorien") {
            # Clear protocol selection when category is selected
            $cmbProtocol.SelectedIndex = 0
        }
        
        # Delayed auto-refresh to prevent multiple rapid calls
        if ($Global:CategoryTimer) {
            $Global:CategoryTimer.Stop()
        }
        $Global:CategoryTimer = New-Object System.Windows.Threading.DispatcherTimer
        $Global:CategoryTimer.Interval = [System.TimeSpan]::FromMilliseconds(500)
        $Global:CategoryTimer.Add_Tick({
            $Global:CategoryTimer.Stop()
            $btnRefresh.RaiseEvent((New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Button]::ClickEvent)))
        })
        $Global:CategoryTimer.Start()
    }
})

# Hyperlink event handlers
$linkPhinit.Add_Click({
    try {
        Start-Process "https://www.phinit.de"
    } catch {
        [System.Diagnostics.Process]::Start("https://www.phinit.de")
    }
})

$linkPsscripts.Add_Click({
    try {
        Start-Process "https://www.psscripts.de"
    } catch {
        [System.Diagnostics.Process]::Start("https://www.psscripts.de")
    }
})

# Function to capture network packets (placeholder for future implementation)
function Capture-Packets {
    $txtPackets.Visibility = [System.Windows.Visibility]::Visible
    $txtPackets.Text = "Paket-Erfassung ist in dieser Version noch nicht implementiert.`nVerwenden Sie die kontinuierliche Erfassung für detaillierte Überwachung."
    [System.Windows.MessageBox]::Show("Paket-Erfassung ist in dieser Version noch nicht verfügbar.`nVerwenden Sie stattdessen die kontinuierliche Erfassung.", "Information", "OK", "Information")
}

$btnCapturePackets.Add_Click({
    Capture-Packets
})

$btnStartCapture.Add_Click({
    if ($Global:CaptureActive) {
        Stop-ContinuousCapture
    } else {
        Start-ContinuousCapture
    }
})

$btnHideStatus.Add_Click({
    if ($txtPackets.Visibility -eq [System.Windows.Visibility]::Visible) {
        $txtPackets.Visibility = [System.Windows.Visibility]::Collapsed
        $btnHideStatus.Content = "Status anzeigen"
    } else {
        $txtPackets.Visibility = [System.Windows.Visibility]::Visible
        $btnHideStatus.Content = "Status ausblenden"
    }
})

$btnExportHTML.Add_Click({
    $selectedProtocol = $cmbProtocol.SelectedItem.Content
    $selectedDirection = $cmbDirection.SelectedItem.Content
    $selectedCategory = $cmbCategory.SelectedItem.Content
    $connections = Get-NetworkConnections -protocol $selectedProtocol -direction $selectedDirection -category $selectedCategory
    Export-HTMLReport -connections $connections -protocol $selectedProtocol -direction $selectedDirection -category $selectedCategory
})

# Window closing event to stop capture and export
$window.add_Closing({
    if ($Global:CaptureActive) {
        Stop-ContinuousCapture
    }
})

# Initial load with error handling
try {
    $btnRefresh.RaiseEvent((New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Button]::ClickEvent)))
} catch {
    Write-Host "Fehler beim initialen Laden: $($_.Exception.Message)"
    # Fallback: Load without filters
    try {
        $connections = Get-NetTCPConnection | Select-Object @{Name="Type";Expression={"TCP"}}, LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess, @{Name="ProcessName";Expression={""}}, @{Name="RemoteHostname";Expression={""}}
        $dgConnections.ItemsSource = $connections
    } catch {
        Write-Host "Fallback auch fehlgeschlagen: $($_.Exception.Message)"
    }
}

# Show window
$window.ShowDialog() | Out-Null
