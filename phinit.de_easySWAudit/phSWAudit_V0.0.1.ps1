<#
.SYNOPSIS
    Client Audit Tool - Umfassende Programm- und Nutzungsanalyse mit WPF GUI
    
.DESCRIPTION
    Erfasst installierte Programme, laufende Prozesse, Nutzungshäufigkeit und Programmverknüpfungen
    Moderne WPF GUI für einfache Bedienung
    
.NOTES
    Version: 0.0.1
    Author: PS-easyIT
    Date: 19.01.2026
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

#region XAML Definition

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Client Audit Tool v0.0.1 - PS-easyIT" 
        Height="750" Width="1400" 
        WindowStartupLocation="CenterScreen"
        Background="#F5F6FA">
    <Window.Resources>
        <!-- Modern Brush Definitions -->
        <SolidColorBrush x:Key="PrimaryColor" Color="#0078D4"/>
        <SolidColorBrush x:Key="SecondaryColor" Color="#106EBE"/>
        <SolidColorBrush x:Key="AccentColor" Color="#50E6FF"/>
        <SolidColorBrush x:Key="DarkBG" Color="#2D2D30"/>
        <SolidColorBrush x:Key="LightBG" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="BorderColor" Color="#E1E1E1"/>
        
        <!-- Button Style -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource PrimaryColor}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Margin" Value="0,5"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="4" 
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource SecondaryColor}"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Opacity" Value="0.5"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="350"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        
        <!-- Left Panel - Options -->
        <Border Grid.Column="0" Background="#FFFFFF" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,1,0" Padding="20">
            <StackPanel>
                <TextBlock Text="🔍 Client Audit Tool" FontSize="24" FontWeight="Bold" Foreground="{StaticResource PrimaryColor}" Margin="0,0,0,20"/>
                
                <!-- System Info -->
                <TextBlock x:Name="txtComputerName" FontSize="12" Foreground="#666" Margin="0,0,0,5"/>
                <TextBlock x:Name="txtUserName" FontSize="12" Foreground="#666" Margin="0,0,0,5"/>
                <TextBlock x:Name="txtOS" FontSize="12" Foreground="#666" Margin="0,0,0,20"/>
                
                <!-- Audit Options -->
                <TextBlock Text="Audit-Optionen:" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,10"/>
                
                <CheckBox x:Name="chkInstalledPrograms" Content="📦 Installierte Programme" Margin="0,5" IsChecked="True"/>
                <CheckBox x:Name="chkStoreApps" Content="🏪 Windows Store Apps" Margin="0,5" IsChecked="True"/>
                <CheckBox x:Name="chkRunningProcesses" Content="⚡ Laufende Prozesse" Margin="0,5" IsChecked="True"/>
                <CheckBox x:Name="chkPrefetch" Content="📊 Prefetch-Analyse" Margin="0,5" IsChecked="False"/>
                <CheckBox x:Name="chkProgramInventory" Content="📋 Programm-Inventar" Margin="0,5" IsChecked="False"/>
                <CheckBox x:Name="chkEventLogs" Content="📝 Event-Logs" Margin="0,5" IsChecked="False"/>
                <CheckBox x:Name="chkShortcuts" Content="🔗 Desktop-Verknüpfungen" Margin="0,5" IsChecked="True"/>
                
                <!-- Filter Options -->
                <TextBlock Text="Filter-Optionen:" FontSize="14" FontWeight="SemiBold" Margin="0,15,0,10"/>
                <CheckBox x:Name="chkExcludeWindows" Content="❌ Windows-Programme ausschließen" Margin="0,5" IsChecked="False"/>
                
                <!-- Action Buttons -->
                <Button x:Name="btnStartAudit" Content="▶ Audit starten" Style="{StaticResource ModernButton}" Margin="0,20,0,10"/>
                <Button x:Name="btnExport" Content="💾 Export" Style="{StaticResource ModernButton}" IsEnabled="False"/>
                <Button x:Name="btnClear" Content="🗑️ Ergebnisse löschen" Style="{StaticResource ModernButton}" IsEnabled="False"/>
                
                <!-- Status -->
                <TextBlock Text="Status:" FontSize="14" FontWeight="SemiBold" Margin="0,30,0,10"/>
                <TextBlock x:Name="txtStatus" Text="Bereit" FontSize="12" Foreground="#666" TextWrapping="Wrap"/>
                <ProgressBar x:Name="progressBar" Height="4" Margin="0,10,0,0" Visibility="Collapsed"/>
            </StackPanel>
        </Border>
        
        <!-- Right Panel - Results -->
        <Grid Grid.Column="1" Margin="20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            
            <!-- Data View Selector -->
            <ComboBox x:Name="cmbDataView" Grid.Row="0" Height="35" FontSize="13" IsEnabled="False" Margin="0,0,0,10"/>
            
            <!-- Data Grid -->
            <DataGrid x:Name="dataGrid" Grid.Row="1" AutoGenerateColumns="True" IsReadOnly="True" 
                      AlternatingRowBackground="#F9F9F9" RowBackground="White"
                      GridLinesVisibility="None" HeadersVisibility="Column"
                      CanUserSortColumns="True" CanUserResizeColumns="True"
                      FontSize="12" Margin="0,0,0,10"/>
            
            <!-- Summary -->
            <TextBlock x:Name="txtSummary" Grid.Row="2" Text="Keine Daten geladen" FontSize="12" Foreground="#666"/>
        </Grid>
    </Grid>
</Window>
"@

#endregion

#region Functions

function Update-Status {
    param([string]$Message)
    $txtStatus.Text = $Message
}

function Export-ToCSV {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataSets,
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }
        
        foreach ($key in $DataSets.Keys) {
            $fileName = $key -replace '[^\w]', '_'
            $filePath = Join-Path $Path "$fileName.csv"
            if ($DataSets[$key] -and $DataSets[$key].Count -gt 0) {
                $DataSets[$key] | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
            }
        }
        return $true
    }
    catch {
        return $false
    }
}

function Export-ToHTML {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$DataSets,
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filePath = Join-Path $Path "ClientAudit_Report_$timestamp.html"
        
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Client Audit Report</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background: #f5f6fa; }
        h1 { color: #0078D4; }
        h2 { color: #106EBE; margin-top: 30px; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 30px; background: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        th { background-color: #0078D4; color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #e1e1e1; }
        tr:hover { background-color: #f9f9f9; }
        .summary { background: white; padding: 15px; margin-bottom: 20px; border-left: 4px solid #0078D4; }
    </style>
</head>
<body>
    <h1>🔍 Client Audit Report</h1>
    <div class="summary">
        <p><strong>Computername:</strong> $env:COMPUTERNAME</p>
        <p><strong>Benutzer:</strong> $env:USERNAME</p>
        <p><strong>Datum:</strong> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
    </div>
"@
        
        foreach ($key in $DataSets.Keys) {
            $html += "<h2>$key</h2>"
            $html += $DataSets[$key] | ConvertTo-Html -Fragment
        }
        
        $html += @"
</body>
</html>
"@
        
        $html | Out-File -FilePath $filePath -Encoding UTF8
        return $filePath
    }
    catch {
        return $null
    }
}

#endregion

#region GUI

# Parse XAML
$reader = [System.Xml.XmlNodeReader]::new([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get Controls
$txtComputerName = $window.FindName('txtComputerName')
$txtUserName = $window.FindName('txtUserName')
$txtOS = $window.FindName('txtOS')
$chkInstalledPrograms = $window.FindName('chkInstalledPrograms')
$chkStoreApps = $window.FindName('chkStoreApps')
$chkRunningProcesses = $window.FindName('chkRunningProcesses')
$chkPrefetch = $window.FindName('chkPrefetch')
$chkProgramInventory = $window.FindName('chkProgramInventory')
$chkEventLogs = $window.FindName('chkEventLogs')
$chkShortcuts = $window.FindName('chkShortcuts')
$chkExcludeWindows = $window.FindName('chkExcludeWindows')
$btnStartAudit = $window.FindName('btnStartAudit')
$btnExport = $window.FindName('btnExport')
$btnClear = $window.FindName('btnClear')
$txtStatus = $window.FindName('txtStatus')
$progressBar = $window.FindName('progressBar')
$cmbDataView = $window.FindName('cmbDataView')
$dataGrid = $window.FindName('dataGrid')
$txtSummary = $window.FindName('txtSummary')

# Global Data
$script:auditData = @{}
$script:exportPath = "$env:USERPROFILE\Desktop\ClientAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

# Initialize System Info
$txtComputerName.Text = "💻 Computer: $env:COMPUTERNAME"
$txtUserName.Text = "👤 Benutzer: $env:USERNAME"
try {
    $os = (Get-CimInstance Win32_OperatingSystem).Caption
    $txtOS.Text = "🖥️ OS: $os"
} catch {
    $txtOS.Text = "🖥️ OS: Windows"
}

# Start Audit Button
$btnStartAudit.Add_Click({
    $btnStartAudit.IsEnabled = $false
    $script:auditData = @{}
    $cmbDataView.Items.Clear()
    $dataGrid.ItemsSource = $null
    Update-Status "Starte Audit..."
    $progressBar.Visibility = 'Visible'
    $progressBar.IsIndeterminate = $true
    
    # Sammle Daten
    if ($chkInstalledPrograms.IsChecked) {
        Update-Status "Erfasse installierte Programme..."
        $programs = @()
        $regPaths = @(
            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
        
        foreach ($path in $regPaths) {
            $programs += Get-ItemProperty $path -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName } |
                Select-Object @{N='ProgramName';E={$_.DisplayName}},
                             @{N='Version';E={$_.DisplayVersion}},
                             @{N='Publisher';E={$_.Publisher}},
                             @{N='InstallDate';E={$_.InstallDate}}
        }
        
        # Filter Windows-Programme falls gewünscht
        if ($chkExcludeWindows.IsChecked) {
            $programs = $programs | Where-Object { 
                -not ($_.Publisher -like '*Microsoft*' -or 
                      $_.ProgramName -like '*Microsoft*' -or
                      $_.ProgramName -like '*Windows*' -or
                      $_.ProgramName -like '*Update*' -or
                      $_.Publisher -like '*Windows*')
            }
        }
        
        if ($programs.Count -gt 0) {
            $script:auditData['Installierte Programme'] = $programs
            $cmbDataView.Items.Add('Installierte Programme')
        }
    }
    
    if ($chkStoreApps.IsChecked) {
        Update-Status "Erfasse Windows Store Apps..."
        try {
            $storeApps = Get-AppxPackage -ErrorAction SilentlyContinue | Select-Object Name, Version, Publisher, InstallLocation
        } catch {
            $storeApps = @()
        }
        
        # Filter Windows Store Apps falls gewünscht
        if ($chkExcludeWindows.IsChecked) {
            $storeApps = $storeApps | Where-Object { 
                -not ($_.Publisher -like '*Microsoft*' -or 
                      $_.Name -like '*Microsoft*' -or
                      $_.Name -like '*Windows*')
            }
        }
        
        if ($storeApps.Count -gt 0) {
            $script:auditData['Windows Store Apps'] = $storeApps
            $cmbDataView.Items.Add('Windows Store Apps')
        }
    }
    
    if ($chkRunningProcesses.IsChecked) {
        Update-Status "Erfasse laufende Prozesse..."
        $processes = Get-Process | Where-Object { $_.MainWindowTitle } |
            Select-Object ProcessName, Id, @{N='MainWindowTitle';E={$_.MainWindowTitle}}, 
                         @{N='Memory_MB';E={[math]::Round($_.WorkingSet64 / 1MB, 2)}}
        if ($processes.Count -gt 0) {
            $script:auditData['Laufende Prozesse'] = $processes
            $cmbDataView.Items.Add('Laufende Prozesse')
        }
    }
    
    if ($chkPrefetch.IsChecked) {
        Update-Status "Erfasse Prefetch-Daten..."
        $prefetchPath = "$env:SystemRoot\Prefetch"
        if (Test-Path $prefetchPath) {
            $prefetchFiles = Get-ChildItem -Path $prefetchPath -Filter "*.pf" -ErrorAction SilentlyContinue |
                Select-Object @{N='ProgramName';E={$_.Name -replace '\.pf$',''}},
                             @{N='LastAccess';E={$_.LastAccessTime}},
                             @{N='Created';E={$_.CreationTime}},
                             @{N='Size_KB';E={[math]::Round($_.Length / 1KB, 2)}}
            
            if ($prefetchFiles.Count -gt 0) {
                $script:auditData['Prefetch-Analyse'] = $prefetchFiles
                $cmbDataView.Items.Add('Prefetch-Analyse')
            }
        }
    }
    
    if ($chkProgramInventory.IsChecked) {
        Update-Status "Erstelle Programm-Inventar..."
        $inventory = @()
        
        # Sammle alle einzigartigen Programme
        $allPrograms = @()
        $regPaths = @(
            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
        
        foreach ($path in $regPaths) {
            $allPrograms += Get-ItemProperty $path -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName } |
                Select-Object @{N='ProgramName';E={$_.DisplayName}},
                             @{N='Publisher';E={$_.Publisher}},
                             @{N='InstallLocation';E={$_.InstallLocation}}
        }
        
        if ($allPrograms.Count -gt 0) {
            $inventory = $allPrograms | Sort-Object ProgramName -Unique
            $script:auditData['Programm-Inventar'] = $inventory
            $cmbDataView.Items.Add('Programm-Inventar')
        }
    }
    
    if ($chkEventLogs.IsChecked) {
        Update-Status "Erfasse Event-Logs..."
        try {
            $events = Get-WinEvent -FilterHashtable @{LogName='Application'; ID=1000,1001,1002} -MaxEvents 100 -ErrorAction SilentlyContinue |
                Select-Object TimeCreated, 
                             @{N='EventID';E={$_.Id}},
                             @{N='Level';E={$_.LevelDisplayName}},
                             @{N='Source';E={$_.ProviderName}},
                             @{N='Message';E={if ($_.Message) { $_.Message.Substring(0, [Math]::Min(100, $_.Message.Length)) } else { "" }}}
        } catch {
            $events = @()
        }
        
        if ($events -and $events.Count -gt 0) {
            $script:auditData['Event-Logs'] = $events
            $cmbDataView.Items.Add('Event-Logs')
        }
    }
    
    if ($chkShortcuts.IsChecked) {
        Update-Status "Erfasse Desktop-Verknüpfungen..."
        $shortcuts = @()
        $locations = @("$env:Public\Desktop", "$env:UserProfile\Desktop")
        
        foreach ($loc in $locations) {
            if (Test-Path $loc) {
                $lnks = Get-ChildItem -Path $loc -Filter "*.lnk" -ErrorAction SilentlyContinue |
                    Select-Object -First 300 |
                    Select-Object @{N='ShortcutName';E={$_.Name}},
                                 @{N='ShortcutPath';E={$_.FullName}},
                                 @{N='Location';E={if ($_.DirectoryName -eq "$env:Public\Desktop") { 'Public Desktop' } else { 'User Desktop' }}},
                                 @{N='LastModified';E={$_.LastWriteTime}}
                if ($lnks) { $shortcuts += $lnks }
            }
        }
        
        if ($shortcuts.Count -gt 0) {
            $script:auditData['Programmverknüpfungen'] = $shortcuts
            $cmbDataView.Items.Add('Programmverknüpfungen')
        }
    }
    
    # Audit abgeschlossen
    if ($script:auditData.Count -gt 0) {
        $cmbDataView.SelectedIndex = 0
        $cmbDataView.IsEnabled = $true
        $btnExport.IsEnabled = $true
        $btnClear.IsEnabled = $true
        $totalItems = ($script:auditData.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
        Update-Status "✓ Audit abgeschlossen! $($script:auditData.Keys.Count) Kategorien mit $totalItems Einträgen erfasst"
    } else {
        Update-Status "Audit abgeschlossen - Keine Daten gefunden"
    }
    
    $progressBar.Visibility = 'Collapsed'
    $btnStartAudit.IsEnabled = $true
})

# ComboBox Selection Changed
$cmbDataView.Add_SelectionChanged({
    if ($cmbDataView.SelectedItem -and $script:auditData.ContainsKey($cmbDataView.SelectedItem)) {
        $selectedData = $script:auditData[$cmbDataView.SelectedItem]
        $dataGrid.ItemsSource = $selectedData
        $txtSummary.Text = "$($cmbDataView.SelectedItem): $($selectedData.Count) Einträge"
    }
})

# Export mit Dialog
$btnExport.Add_Click({
    if ($script:auditData.Count -eq 0) {
        Update-Status "Keine Daten zum Exportieren vorhanden"
        return
    }
    
    # Export-Dialog XAML
    $exportDialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Export-Optionen" Height="280" Width="400" WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Text="Export-Format wählen:" FontWeight="Bold" Margin="0,0,0,10"/>
        <StackPanel Grid.Row="1" Margin="0,0,0,15">
            <RadioButton Name="rbCSV" Content="📊 CSV Export" IsChecked="True" Margin="0,5"/>
            <RadioButton Name="rbHTML" Content="📄 HTML Export" Margin="0,5"/>
        </StackPanel>
        
        <TextBlock Grid.Row="2" Text="Filter-Optionen:" FontWeight="Bold" Margin="0,0,0,10"/>
        <CheckBox Grid.Row="3" Name="chkExcludeSystem" Content="System-Programme ausschließen (vorinstallierte OS-Programme)" Margin="0,5" IsChecked="False"/>
        
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button Name="btnOK" Content="Export" Width="80" Height="30" Margin="5,0" Background="#0078D4" Foreground="White" BorderThickness="0" IsDefault="True"/>
            <Button Name="btnCancel" Content="Abbrechen" Width="80" Height="30" Margin="5,0" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    # Export-Dialog erstellen
    $exportDialog = [Windows.Markup.XamlReader]::Parse($exportDialogXaml)
    $rbCSV = $exportDialog.FindName('rbCSV')
    $rbHTML = $exportDialog.FindName('rbHTML')
    $chkExcludeSystem = $exportDialog.FindName('chkExcludeSystem')
    $btnOK = $exportDialog.FindName('btnOK')
    $btnCancel = $exportDialog.FindName('btnCancel')
    
    $btnOK.Add_Click({ $exportDialog.DialogResult = $true; $exportDialog.Close() })
    $btnCancel.Add_Click({ $exportDialog.DialogResult = $false; $exportDialog.Close() })
    
    $result = $exportDialog.ShowDialog()
    
    if ($result -eq $true) {
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        
        # Daten filtern falls System-Programme ausgeschlossen werden sollen
        $filteredData = @{}
        foreach ($key in $script:auditData.Keys) {
            if ($chkExcludeSystem.IsChecked -and ($key -eq 'InstalledPrograms' -or $key -eq 'StoreApps')) {
                # System-Programme filtern (Microsoft, Windows, Update, System32)
                $filteredData[$key] = $script:auditData[$key] | Where-Object {
                    $_.Publisher -notmatch 'Microsoft' -and
                    $_.Publisher -notmatch 'Windows' -and
                    $_.Name -notmatch 'Windows' -and
                    $_.Name -notmatch 'Update' -and
                    $_.Name -notmatch 'KB[0-9]' -and
                    $_.InstallLocation -notmatch 'System32'
                }
            } else {
                $filteredData[$key] = $script:auditData[$key]
            }
        }
        
        # Export basierend auf Format
        if ($rbCSV.IsChecked) {
            $csvPath = "$env:USERPROFILE\Desktop\ClientAudit_CSV_$timestamp"
            $csvExported = Export-ToCSV -DataSets $filteredData -Path $csvPath
            if ($csvExported) {
                Update-Status "CSV-Export erfolgreich erstellt: $csvPath"
                Start-Process $csvPath
            } else {
                Update-Status "Fehler beim CSV-Export"
            }
        } else {
            $htmlPath = "$env:USERPROFILE\Desktop"
            $htmlExported = Export-ToHTML -DataSets $filteredData -Path $htmlPath
            if ($htmlExported) {
                Update-Status "HTML-Report erstellt: $htmlExported"
                Start-Process $htmlExported
            } else {
                Update-Status "Fehler beim HTML-Export"
            }
        }
    }
})

# Clear Results
$btnClear.Add_Click({
    $script:auditData = @{}
    $cmbDataView.Items.Clear()
    $cmbDataView.IsEnabled = $false
    $dataGrid.ItemsSource = $null
    $txtSummary.Text = "Keine Daten geladen"
    Update-Status "Bereit"
    $btnExport.IsEnabled = $false
    $btnClear.IsEnabled = $false
})

#endregion

# Show Window
[void]$window.ShowDialog()
