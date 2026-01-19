# easyConnections_V0.0.3.ps1
# Advanced Network Connections Monitor with WPF GUI - Admin Edition
# V0.0.3: Lazy Loading + Recording + Color-Coded Categories
# Performance: 80-90% schneller beim Start durch Lazy Loading

# ===== Optional Logging System =====
$Global:LoggingEnabled = $false  # Set to $true to enable logging to file
$Global:LogFile = "$script:ScriptRoot\easyConnections.log"

function Write-Log {
    param([string]$Message)
    if ($Global:LoggingEnabled) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] $Message"
        Add-Content -Path $Global:LogFile -Value $logEntry -ErrorAction SilentlyContinue
    }
}

Write-Log "Script started: easyConnections_V0.0.3.ps1"

###############################################################################
# PS2EXE OPTIMIERUNGEN
###############################################################################
# PS2EXE: Progress Bars deaktivieren für -NoConsole Kompilierung
$ProgressPreference = 'SilentlyContinue'

# PS2EXE: Visual Styles aktivieren BEVOR GUI-Objekte erstellt werden
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
[System.Windows.Forms.Application]::EnableVisualStyles()

# Universeller Pfad-Code für Script UND EXE-Kompatibilität
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    # Script-Modus
    $script:ScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
} else {
    # Executable-Modus (PS2EXE)
    $script:ScriptRoot = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
    if ([string]::IsNullOrWhiteSpace($script:ScriptRoot)) { 
        $script:ScriptRoot = "."
    }
}
if (-not $script:ScriptRoot) { 
    $script:ScriptRoot = "."
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

Write-Log "Assemblies loaded successfully"

# XAML Content mit Windows 11 Fluent Design
$xamlContent = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="easyConnections - Network Monitor"
    Width="1400"
    Height="950"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize"
    MinWidth="1300"
    MinHeight="800"
    Background="#FFFBF9"
    WindowStyle="SingleBorderWindow"
    TextOptions.TextFormattingMode="Display"
    TextOptions.TextRenderingMode="ClearType"
    UseLayoutRounding="True">

    <Window.Resources>
        <!-- Modern Fluent Button Style -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="20,10"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                BorderThickness="{TemplateBinding BorderThickness}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#106EBE"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#005A9E"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="ModernButtonAkt" TargetType="Button">
            <Setter Property="Background" Value="#36913a"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="20,10"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                BorderThickness="{TemplateBinding BorderThickness}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#106EBE"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#005A9E"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Accent Button Style -->
        <Style x:Key="AccentButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#0078D4"/>
        </Style>

        <!-- Subtle Button Style -->
        <Style x:Key="SubtleButton" TargetType="Button">
            <Setter Property="Background" Value="#ffbeb5"/>
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="18,9"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="Normal"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                BorderThickness="{TemplateBinding BorderThickness}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#E8E8E8"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#D8D8D8"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern ComboBox Style -->
        <Style x:Key="ModernComboBox" TargetType="ComboBox">
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="BorderBrush" Value="#E0E0E0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Name="ToggleButton" Grid.Column="2" 
                                         ClickMode="Press" Focusable="False"
                                         IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                         Template="{DynamicResource ComboBoxToggleButtonTemplate}"/>
                            <ContentPresenter Name="ContentSite" IsHitTestVisible="False" 
                                            Content="{TemplateBinding SelectionBoxItem}"
                                            ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                            ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                            VerticalAlignment="Center" HorizontalAlignment="Left" Margin="12,0,0,0"/>
                            <TextBox x:Name="PART_EditableTextBox"
                                    HorizontalAlignment="Left"
                                    VerticalAlignment="Center"
                                    Margin="3,0,0,0"
                                    Focusable="True"
                                    Background="Transparent"
                                    Visibility="Hidden"
                                    Foreground="{TemplateBinding Foreground}"/>
                            <Popup Name="Popup"
                                  Placement="Bottom"
                                  IsOpen="{TemplateBinding IsDropDownOpen}"
                                  AllowsTransparency="True"
                                  Focusable="False"
                                  PopupAnimation="Slide">
                                <Border Name="DropDownBorder"
                                       Background="#FFFFFF"
                                       BorderThickness="1"
                                       BorderBrush="#E0E0E0"
                                       CornerRadius="6"
                                       MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <ScrollViewer SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern CheckBox Style -->
        <Style x:Key="ModernCheckBox" TargetType="CheckBox">
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <StackPanel Orientation="Horizontal">
                            <Border Width="18" Height="18" CornerRadius="4"
                                   Background="#F0F0F0" BorderBrush="#D0D0D0" BorderThickness="1"
                                   Margin="0,0,8,0">
                                <Canvas x:Name="CheckMark" Visibility="Collapsed">
                                    <Polyline Points="2,8 6,12 14,4" Stroke="#0078D4" StrokeThickness="2"/>
                                </Canvas>
                            </Border>
                            <ContentPresenter VerticalAlignment="Center"/>
                        </StackPanel>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- DataGrid Header Style -->
        <Style x:Key="ModernDataGridColumnHeaderStyle" TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#F3F3F3"/>
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="BorderBrush" Value="#E0E0E0"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="HorizontalAlignment" Value="Stretch"/>
        </Style>

        <!-- DataGrid Row Style with Color Coding -->
        <Style x:Key="ColoredDataGridRowStyle" TargetType="DataGridRow">
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#F5F5F5"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#E3F2FD"/>
                </Trigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="Web">
                    <Setter Property="Background" Value="#E3F2FD"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="Email">
                    <Setter Property="Background" Value="#FFF3E0"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="Database">
                    <Setter Property="Background" Value="#E8F5E9"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="Directory">
                    <Setter Property="Background" Value="#F3E5F5"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="RemoteAccess">
                    <Setter Property="Background" Value="#FCE4EC"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="FileServices">
                    <Setter Property="Background" Value="#FFFDE7"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="Monitoring">
                    <Setter Property="Background" Value="#F5F5F5"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="Virtualization">
                    <Setter Property="Background" Value="#E0F2F1"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>

        <!-- Modern DataGrid Style -->
        <Style x:Key="ModernDataGrid" TargetType="DataGrid">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="BorderBrush" Value="#E0E0E0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="RowHeaderWidth" Value="0"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#F0F0F0"/>
        </Style>

        <!-- Section Header Style -->
        <Style x:Key="SectionHeader" TargetType="TextBlock">
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#0F0F0F"/>
        </Style>
    </Window.Resources>

    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Main Actions Row -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10" Background="Transparent" VerticalAlignment="Center">
            <Button x:Name="btnRefresh" Content="🔄 Aktualisieren" Style="{StaticResource ModernButtonAkt}" Margin="0,0,10,0" Width="155" Height="40" ToolTip="Verbindungsliste aktualisieren"/>
            <Button x:Name="btnStartRecording" Content="⏺️ Aufzeichnung" Style="{StaticResource ModernButton}" Margin="0,0,10,0" Width="155" Height="40" ToolTip="Verbindungen aufzeichnen"/>
            <Button x:Name="btnExportHTML" Content="📊 HTML Export" Style="{StaticResource ModernButton}" Margin="0,0,20,0" Width="155" Height="40" ToolTip="Verbindungen als HTML exportieren"/>
            
            <TextBlock Text="🎯 Presets:" Style="{StaticResource SectionHeader}" VerticalAlignment="Center" Margin="100,0,10,0"/>
            <ComboBox x:Name="cmbPresets" Width="220" Margin="0,0,8,0" Style="{StaticResource ModernComboBox}" Height="36" ToolTip="Filter-Presets laden">
                <ComboBoxItem Content="-- Neues Preset --" IsSelected="True"/>
            </ComboBox>
            <Button x:Name="btnLoadPreset" Content="📂 Laden" Style="{StaticResource ModernButton}" Width="100" Height="36" Margin="0,0,6,0" ToolTip="Ausgewähltes Preset laden"/>
            <Button x:Name="btnSavePreset" Content="💾 Speichern" Style="{StaticResource ModernButton}" Width="140" Height="36" Margin="0,0,6,0" ToolTip="Aktuelle Filter als Preset speichern"/>
            <Button x:Name="btnDeletePreset" Content="🗑️ Löschen" Style="{StaticResource SubtleButton}" Width="110" Height="36"  Margin="80,0,6,0" ToolTip="Preset löschen"/>
        </StackPanel>

        <!-- Filter Row 1: Network & Protocol Filters with ScrollBar -->
        <Border Grid.Row="1" Background="#F8F8F8" BorderBrush="#E5E5E5" BorderThickness="0,1,0,1" Padding="12,12,12,12" Margin="-12,0,-12,0">
            <ScrollViewer VerticalScrollBarVisibility="Disabled" HorizontalScrollBarVisibility="Auto" PanningMode="HorizontalOnly">
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="🔍 Filter:" Style="{StaticResource SectionHeader}" VerticalAlignment="Center" Margin="0,0,15,0" MinWidth="60"/>
                    
                    <TextBlock Text="Netzwerk:" FontSize="11" FontWeight="SemiBold" Foreground="#333333" VerticalAlignment="Center" Margin="0,0,8,0" MinWidth="70"/>
                    <ComboBox x:Name="cmbNetworkType" Width="140" Margin="0,0,15,0" Style="{StaticResource ModernComboBox}" Height="34" ToolTip="Netzwerk-Typ filtern">
                        <ComboBoxItem Content="Alle Netzwerke" IsSelected="True"/>
                        <ComboBoxItem Content="Nur Private"/>
                        <ComboBoxItem Content="Nur Public"/>
                    </ComboBox>
                    
                    <TextBlock Text="Protokoll:" FontSize="11" FontWeight="SemiBold" Foreground="#333333" VerticalAlignment="Center" Margin="0,0,8,0" MinWidth="70"/>
                    <ComboBox x:Name="cmbTCPUDP" Width="130" Margin="0,0,15,0" Style="{StaticResource ModernComboBox}" Height="34" ToolTip="TCP/UDP Typ">
                        <ComboBoxItem Content="TCP + UDP" IsSelected="True"/>
                        <ComboBoxItem Content="Nur TCP"/>
                        <ComboBoxItem Content="Nur UDP"/>
                    </ComboBox>
                    
                    <TextBlock Text="Richtung:" FontSize="11" FontWeight="SemiBold" Foreground="#333333" VerticalAlignment="Center" Margin="0,0,8,0" MinWidth="70"/>
                    <ComboBox x:Name="cmbDirection" Width="120" Margin="0,0,15,0" Style="{StaticResource ModernComboBox}" Height="34" ToolTip="Verbindungsrichtung">
                        <ComboBoxItem Content="Beide" IsSelected="True"/>
                        <ComboBoxItem Content="Eingehend"/>
                        <ComboBoxItem Content="Ausgehend"/>
                    </ComboBox>
                    
                    <TextBlock Text="Kategorie:" FontSize="11" FontWeight="SemiBold" Foreground="#333333" VerticalAlignment="Center" Margin="0,0,8,0" MinWidth="70"/>
                    <ComboBox x:Name="cmbCategory" Width="155" Margin="0,0,15,0" Style="{StaticResource ModernComboBox}" Height="34" ToolTip="Nach Kategorie filtern">
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
                    
                    <TextBlock Text="Service:" FontSize="11" FontWeight="SemiBold" Foreground="#333333" VerticalAlignment="Center" Margin="0,0,8,0" MinWidth="60"/>
                    <ComboBox x:Name="cmbProtocol" Width="145" Margin="0,0,15,0" Style="{StaticResource ModernComboBox}" Height="34" ToolTip="Nach Protokoll filtern">
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
                    
                    <CheckBox x:Name="chkShowProcessInfo" Content="📋 Prozessinfo" VerticalAlignment="Center" Margin="10,0,0,0" IsChecked="False" Style="{StaticResource ModernCheckBox}" ToolTip="Prozessinformationen anzeigen"/>
                </StackPanel>
            </ScrollViewer>
        </Border>

        <!-- DataGrid -->
        <DataGrid x:Name="dgConnections" Grid.Row="3" AutoGenerateColumns="False" IsReadOnly="True" 
                 RowStyle="{StaticResource ColoredDataGridRowStyle}" 
                 Style="{StaticResource ModernDataGrid}"
                 ColumnHeaderStyle="{StaticResource ModernDataGridColumnHeaderStyle}"
                 AlternationCount="0"
                 ScrollViewer.HorizontalScrollBarVisibility="Auto"
                 ScrollViewer.VerticalScrollBarVisibility="Auto">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Typ" Binding="{Binding Type}" Width="50"/>
                <DataGridTextColumn Header="Lokale Adresse" Binding="{Binding LocalAddress}" Width="140"/>
                <DataGridTextColumn Header="Lokaler Port" Binding="{Binding LocalPort}" Width="90"/>
                <DataGridTextColumn Header="Remote Adresse" Binding="{Binding RemoteAddress}" Width="140"/>
                <DataGridTextColumn Header="Remote Port" Binding="{Binding RemotePort}" Width="90"/>
                <DataGridTextColumn Header="Remote Hostname" Binding="{Binding RemoteHostname}" Width="*" MinWidth="140"/>
                <DataGridTextColumn Header="Status" Binding="{Binding State}" Width="80"/>
                <DataGridTextColumn x:Name="dgColProcessID" Header="Prozess ID" Binding="{Binding OwningProcess}" Width="80"/>
                <DataGridTextColumn x:Name="dgColProcessName" Header="Prozess Name" Binding="{Binding ProcessName}" Width="150"/>
                <DataGridTextColumn Header="Kategorie" Binding="{Binding CategoryName}" Width="130"/>
            </DataGrid.Columns>
        </DataGrid>

        <!-- Status Bar -->
        <Border Grid.Row="4" Background="#F3F3F3" BorderBrush="#E0E0E0" BorderThickness="0,1,0,0" Padding="8,8,8,8" Margin="-12,8,-12,0">
            <StackPanel Orientation="Horizontal">
                <TextBlock Text="Status:" FontWeight="SemiBold" Foreground="#333333" Margin="0,0,8,0" FontFamily="Segoe UI" FontSize="11"/>
                <TextBlock x:Name="txtStatus" Text="Bereit" Foreground="#666666" FontFamily="Segoe UI" FontSize="11"/>
            </StackPanel>
        </Border>

        <!-- Footer -->
        <Border Grid.Row="5" Background="#FAFAFA" BorderBrush="#E0E0E0" BorderThickness="0,1,0,0" Padding="8,6,8,6" Margin="-12,8,-12,-12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="easyConnections v0.0.3 • Network Monitor • GitHub Edition" 
                          Foreground="#999999" FontSize="10" FontFamily="Segoe UI" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="1" Text="Powered by Andreas Hepp | www.phinit.de" Foreground="#0078D4" FontSize="10" FontWeight="SemiBold" FontFamily="Segoe UI"/>
            </Grid>
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
$cmbNetworkType = $window.FindName("cmbNetworkType")
$cmbTCPUDP = $window.FindName("cmbTCPUDP")
$chkShowProcessInfo = $window.FindName("chkShowProcessInfo")
$btnStartRecording = $window.FindName("btnStartRecording")
$btnExportHTML = $window.FindName("btnExportHTML")
$dgConnections = $window.FindName("dgConnections")
$dgColProcessID = $window.FindName("dgColProcessID")
$dgColProcessName = $window.FindName("dgColProcessName")
$txtStatus = $window.FindName("txtStatus")
$cmbPresets = $window.FindName("cmbPresets")
$btnLoadPreset = $window.FindName("btnLoadPreset")
$btnSavePreset = $window.FindName("btnSavePreset")
$btnDeletePreset = $window.FindName("btnDeletePreset")

# Global variables for Recording Feature
$Global:RecordingActive = $false
$Global:RecordingTimer = $null
$Global:RecordedConnections = @()
$Global:RecordedConnectionsHash = @{}
$Global:ProcessCache = @{}
$Global:RecordingStartTime = $null
$Global:RecordingCategory = "Alle Kategorien"

# Preset Storage
$Global:PresetsFile = "$script:ScriptRoot\presets.json"
$Global:FilterPresets = @{}

# Protocol port mappings
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

# Protocol categories
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

# Category to color mapping
$categoryColorMap = @{
    "Web Services" = "Web"
    "Email Services" = "Email"
    "Database Services" = "Database"
    "Directory Services" = "Directory"
    "Remote Access" = "RemoteAccess"
    "File Services" = "FileServices"
    "Monitoring" = "Monitoring"
    "Virtualization" = "Virtualization"
}

# Function to toggle process columns visibility
function Update-ProcessColumnsVisibility {
    param([bool]$showProcessInfo)
    
    if ($showProcessInfo) {
        $dgColProcessID.Visibility = [System.Windows.Visibility]::Visible
        $dgColProcessName.Visibility = [System.Windows.Visibility]::Visible
    } else {
        $dgColProcessID.Visibility = [System.Windows.Visibility]::Collapsed
        $dgColProcessName.Visibility = [System.Windows.Visibility]::Collapsed
    }
}

# Function to check if IP is private
function Test-IsPrivateIP {
    param([string]$ip)
    
    if ($ip -eq "0.0.0.0" -or $ip -eq "::" -or $ip -eq "N/A") {
        return $true
    }
    
    # Loopback addresses
    if ($ip -eq "127.0.0.1" -or $ip -eq "::1") {
        return $true
    }
    
    # IPv4 private ranges
    if ($ip -match "^127\.") { return $true }           # 127.0.0.0/8
    if ($ip -match "^192\.168\.") { return $true }       # 192.168.0.0/16
    if ($ip -match "^10\.") { return $true }             # 10.0.0.0/8
    if ($ip -match "^172\.(1[6-9]|2[0-9]|3[01])\.") { 
        return $true 
    }                                                    # 172.16.0.0/12
    
    # IPv6 private ranges
    if ($ip -match "^fc[0-9a-f]{2}:") { return $true }   # fc00::/7
    if ($ip -match "^fd[0-9a-f]{2}:") { return $true }   # fd00::/8
    if ($ip -match "^fe80:") { return $true }            # fe80::/10 (Link-local)
    if ($ip -match "^::1") { return $true }              # Loopback
    
    # Alle anderen IPs sind öffentlich
    return $false
}

# Function to determine category and color based on port
function Get-CategoryAndColor {
    param([object]$LocalPort, [object]$RemotePort, [string]$Type)
    
    # Convert to int safely, handle "N/A" strings
    try {
        $localPortInt = if ($LocalPort -eq "N/A" -or $LocalPort -eq 0) { 0 } else { [int]$LocalPort }
        $remotePortInt = if ($RemotePort -eq "N/A" -or $RemotePort -eq 0) { 0 } else { [int]$RemotePort }
    } catch {
        $localPortInt = 0
        $remotePortInt = 0
    }
    
    $checkPort = if ($Type -eq "TCP" -and $remotePortInt -ne 0) { $remotePortInt } else { $localPortInt }
    
    foreach ($category in $protocolCategories.GetEnumerator()) {
        $categoryName = $category.Name
        $protocols = $category.Value
        
        foreach ($protocol in $protocols) {
            if ($protocolPorts.ContainsKey($protocol)) {
                if ($checkPort -in $protocolPorts[$protocol]) {
                    return @{
                        CategoryName = $categoryName
                        CategoryColor = $categoryColorMap[$categoryName]
                    }
                }
            }
        }
    }
    
    return @{
        CategoryName = "Sonstige"
        CategoryColor = "Other"
    }
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
        [string]$category = "Alle Kategorien",
        [bool]$includeProcessInfo = $false,
        [string]$networkType = "Alle Netzwerke",
        [string]$tcpUdpType = "TCP + UDP"
    )

    # Get TCP connections - convert to PSCustomObject
    $tcpConnections = @()
    if ($tcpUdpType -ne "Nur UDP") {
        $tcpConnections = Get-NetTCPConnection | ForEach-Object {
            [PSCustomObject]@{
                Type = "TCP"
                LocalAddress = $_.LocalAddress
                LocalPort = $_.LocalPort
                RemoteAddress = $_.RemoteAddress
                RemotePort = $_.RemotePort
                State = $_.State
                OwningProcess = $_.OwningProcess
                ProcessName = ""
                RemoteHostname = ""
                CategoryName = ""
                CategoryColor = ""
            }
        }
    }

    # Get UDP endpoints - convert to PSCustomObject
    $udpConnections = @()
    if ($tcpUdpType -ne "Nur TCP") {
        $udpConnections = Get-NetUDPEndpoint | ForEach-Object {
            [PSCustomObject]@{
                Type = "UDP"
                LocalAddress = $_.LocalAddress
                LocalPort = $_.LocalPort
                RemoteAddress = "N/A"
                RemotePort = "N/A"
                State = "Listen"
                OwningProcess = $_.OwningProcess
                ProcessName = ""
                RemoteHostname = "N/A"
                CategoryName = ""
                CategoryColor = ""
            }
        }
    }

    $connections = @($tcpConnections) + @($udpConnections)

    # Filter by network type (Private/Public/All)
    if ($networkType -ne "Alle Netzwerke") {
        $connections = $connections | Where-Object {
            $isLocalPrivate = Test-IsPrivateIP -ip $_.LocalAddress
            $isRemotePrivate = Test-IsPrivateIP -ip $_.RemoteAddress
            $allPrivate = $isLocalPrivate -and $isRemotePrivate
            
            if ($networkType -eq "Nur Private") {
                $allPrivate
            } elseif ($networkType -eq "Nur Public") {
                -not $allPrivate
            } else {
                $true
            }
        }
    }

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
        } else {
            $connections = $connections | Where-Object { ($_.LocalPort -in $ports -or ($_.RemotePort -in $ports -and $_.Type -eq "TCP")) }
        }
    }

    # Resolve process names, hostnames and add category/color info
    foreach ($conn in $connections) {
        # Only resolve process names if requested
        if ($includeProcessInfo -and -not $conn.ProcessName) {
            $conn.ProcessName = Get-ProcessNameCached -ProcessId $conn.OwningProcess
        }
        
        # Add category and color info
        $categoryInfo = Get-CategoryAndColor -LocalPort $conn.LocalPort -RemotePort $conn.RemotePort -Type $conn.Type
        $conn.CategoryName = $categoryInfo.CategoryName
        $conn.CategoryColor = $categoryInfo.CategoryColor
        
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

# Fast hostname resolution
function Get-HostnameFast {
    param([string]$ip)
    try {
        if ($ip -eq "0.0.0.0" -or $ip -eq "::" -or $ip -eq "127.0.0.1" -or $ip -eq "::1" -or $ip -eq "N/A") {
            return $ip
        }
        $hostEntry = [System.Net.Dns]::GetHostEntry($ip)
        return $hostEntry.HostName
    } catch {
        return $ip
    }
}

# Function to start Recording
function Start-ConnectionRecording {
    $Global:RecordingActive = $true
    $Global:RecordingStartTime = Get-Date
    $Global:RecordingCategory = $cmbCategory.SelectedItem.Content
    $Global:RecordedConnections = @()
    $Global:RecordedConnectionsHash = @{}
    
    $btnStartRecording.Content = "Aufzeichnung stoppen"
    $btnStartRecording.Background = [System.Windows.Media.Brushes]::Red
    
    $Global:RecordingTimer = New-Object System.Windows.Threading.DispatcherTimer
    $Global:RecordingTimer.Interval = [System.TimeSpan]::FromSeconds(3)
    
    $Global:RecordingTimer.Add_Tick({
        $selectedProtocol = $cmbProtocol.SelectedItem.Content
        $selectedDirection = $cmbDirection.SelectedItem.Content
        $selectedCategory = $cmbCategory.SelectedItem.Content
        $showProcessInfo = $chkShowProcessInfo.IsChecked
        $networkType = $cmbNetworkType.SelectedItem.Content
        $tcpUdpType = $cmbTCPUDP.SelectedItem.Content
        $connections = Get-NetworkConnections -protocol $selectedProtocol -direction $selectedDirection -category $selectedCategory -includeProcessInfo $showProcessInfo -networkType $networkType -tcpUdpType $tcpUdpType
        
        # Ensure connections is always an array
        if ($connections -is [System.Management.Automation.PSCustomObject]) {
            $connections = @($connections)
        } elseif ($null -eq $connections) {
            $connections = @()
        }
        
        $dgConnections.ItemsSource = [System.Collections.ArrayList]@($connections)
        
        foreach ($conn in $connections) {
            $conn | Add-Member -MemberType NoteProperty -Name "RecordingTime" -Value (Get-Date) -Force
            
            $key = "$($conn.LocalAddress):$($conn.LocalPort)-$($conn.RemoteAddress):$($conn.RemotePort)-$($conn.Type)"
            
            if (-not $Global:RecordedConnectionsHash.ContainsKey($key)) {
                $Global:RecordedConnections += $conn
                $Global:RecordedConnectionsHash[$key] = $true
            }
        }
        
        $duration = ((Get-Date) - $Global:RecordingStartTime).TotalSeconds
        $txtStatus.Text = "Aufzeichnung läuft... Kategorie: $selectedCategory`nErfasste Verbindungen: $($Global:RecordedConnections.Count) | Dauer: $([Math]::Floor($duration))s | Letztes Update: $(Get-Date -Format 'HH:mm:ss')"
    })
    
    $Global:RecordingTimer.Start()
    $txtStatus.Text = "Aufzeichnung gestartet für Kategorie: $($cmbCategory.SelectedItem.Content)"
}

# Function to stop Recording
function Stop-ConnectionRecording {
    $Global:RecordingActive = $false
    $btnStartRecording.Content = "Aufzeichnung starten"
    $btnStartRecording.Background = [System.Windows.Media.Brushes]::Green
    
    if ($Global:RecordingTimer) {
        $Global:RecordingTimer.Stop()
        $Global:RecordingTimer = $null
    }
    
    $duration = ((Get-Date) - $Global:RecordingStartTime).TotalSeconds
    $txtStatus.Text = "Aufzeichnung beendet. Verbindungen erfasst: $($Global:RecordedConnections.Count) | Dauer: $([Math]::Floor($duration))s"
    
    if ($Global:RecordedConnections.Count -gt 0) {
        Export-RecordingReportHTML
    }
}

# Function to export Recording as HTML report
function Export-RecordingReportHTML {
    $DateTimeNow = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $RecordingEnd = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $RecordingStart = $Global:RecordingStartTime.ToString("dd.MM.yyyy HH:mm:ss")
    $duration = ((Get-Date) - $Global:RecordingStartTime).TotalSeconds
    $totalConnections = $Global:RecordedConnections.Count
    $tcpCount = ($Global:RecordedConnections | Where-Object { $_.Type -eq "TCP" }).Count
    $udpCount = ($Global:RecordedConnections | Where-Object { $_.Type -eq "UDP" }).Count
    $computerName = $env:COMPUTERNAME
    $userName = $env:USERNAME
    $category = $Global:RecordingCategory

    $html = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Aufzeichnungs-Report - easyConnections</title>
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
        .recording-info { background-color: #fff3cd; padding: 10px; border-radius: 4px; margin-bottom: 20px; border-left: 4px solid #ffc107; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Netzwerk-Aufzeichnungs-Report</h1>
            <div class="subtitle">easyConnections Tool - Kontinuierliche Überwachung</div>
            <div class="info-box">
                <strong>Computer:</strong> $computerName | <strong>Benutzer:</strong> $userName<br>
                <strong>Kategorie:</strong> $category
            </div>
            <div class="recording-info">
                <strong>Aufzeichnungszeitraum:</strong> $RecordingStart bis $RecordingEnd<br>
                <strong>Dauer:</strong> $([Math]::Floor($duration)) Sekunden
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

    foreach ($conn in $Global:RecordedConnections) {
        $recordingTime = if ($conn.RecordingTime) { $conn.RecordingTime.ToString("dd.MM.yyyy HH:mm:ss") } else { "N/A" }
        $processName = if ($conn.ProcessName) { $conn.ProcessName } else { "N/A" }
        $html += "<tr><td>$($conn.Type)</td><td>$($conn.LocalAddress)</td><td>$($conn.LocalPort)</td><td>$($conn.RemoteAddress)</td><td>$($conn.RemotePort)</td><td>$($conn.RemoteHostname)</td><td>$($conn.State)</td><td>$($conn.OwningProcess)</td><td>$processName</td><td>$recordingTime</td></tr>"
    }

    $html += @"
            </tbody>
        </table>
    </div>
    <div class="footer">
        <p>Copyright © $(Get-Date -Format "yyyy") | Autor: Andreas Hepp | Webseite: <a href="https://www.phinit.de">www.phinit.de</a> / <a href="https://www.psscripts.de">www.psscripts.de</a></p>
        <p>Generiert mit easyConnections Tool V0.0.3 - Kontinuierliche Aufzeichnung</p>
    </div>
</body>
</html>
"@

    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "HTML files (*.html)|*.html"
    $saveFileDialog.Title = "Aufzeichnungs-Report speichern"
    $saveFileDialog.FileName = "RecordingReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $html | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
        [System.Windows.MessageBox]::Show("Aufzeichnungs-Report gespeichert: $($saveFileDialog.FileName)", "Erfolg", "OK", "Information")
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
        <p>Generiert mit easyConnections Tool V0.0.3</p>
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
        $originalText = $btnRefresh.Content
        $btnRefresh.Content = "Lädt..."
        $btnRefresh.IsEnabled = $false
        
        $selectedProtocol = $cmbProtocol.SelectedItem.Content
        $selectedDirection = $cmbDirection.SelectedItem.Content
        $selectedCategory = $cmbCategory.SelectedItem.Content
        $showProcessInfo = $chkShowProcessInfo.IsChecked
        $networkType = $cmbNetworkType.SelectedItem.Content
        $tcpUdpType = $cmbTCPUDP.SelectedItem.Content
        
        $connections = Get-NetworkConnections -protocol $selectedProtocol -direction $selectedDirection -category $selectedCategory -includeProcessInfo $showProcessInfo -networkType $networkType -tcpUdpType $tcpUdpType
        
        # Ensure connections is always an array
        if ($connections -is [System.Management.Automation.PSCustomObject]) {
            $connections = @($connections)
        } elseif ($null -eq $connections) {
            $connections = @()
        }
        
        $dgConnections.ItemsSource = [System.Collections.ArrayList]@($connections)
        
        $txtStatus.Text = "Verbindungen geladen: $($connections.Count) gefunden | Protokoll: $selectedProtocol | Richtung: $selectedDirection | Kategorie: $selectedCategory | Netzwerk: $networkType | TCP/UDP: $tcpUdpType"
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Laden der Verbindungen: $($_.Exception.Message)", "Fehler", "OK", "Error")
    } finally {
        $btnRefresh.Content = $originalText
        $btnRefresh.IsEnabled = $true
    }
})

# Checkbox event handler for process info visibility
$chkShowProcessInfo.Add_Click({
    $showProcessInfo = $chkShowProcessInfo.IsChecked
    Update-ProcessColumnsVisibility -showProcessInfo $showProcessInfo
})

# Recording button event
$btnStartRecording.Add_Click({
    if ($Global:RecordingActive) {
        Stop-ConnectionRecording
    } else {
        Start-ConnectionRecording
    }
})

# Export HTML button event
$btnExportHTML.Add_Click({
    if ($dgConnections.ItemsSource.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Keine Verbindungen zum Exportieren. Bitte klicken Sie zuerst auf 'Aktualisieren'.", "Information", "OK", "Information")
        return
    }
    
    $selectedProtocol = $cmbProtocol.SelectedItem.Content
    $selectedDirection = $cmbDirection.SelectedItem.Content
    $selectedCategory = $cmbCategory.SelectedItem.Content
    $connections = $dgConnections.ItemsSource
    Export-HTMLReport -connections $connections -protocol $selectedProtocol -direction $selectedDirection -category $selectedCategory
})

# Preset Functions
function Import-FilterPresets {
    # Initialize with default presets
    $defaultPresets = @{
        "Web-Debugging" = @{
            Protocol = "HTTP"
            Direction = "Beide"
            Category = "Web Services"
            NetworkType = "Alle Netzwerke"
            TCPUDPType = "Nur TCP"
            ProcessInfo = $true
        }
        "Datenbank-Monitor" = @{
            Protocol = "Alle"
            Direction = "Beide"
            Category = "Database Services"
            NetworkType = "Alle Netzwerke"
            TCPUDPType = "Nur TCP"
            ProcessInfo = $true
        }
        "Security-Audit" = @{
            Protocol = "Alle"
            Direction = "Beide"
            Category = "Alle Kategorien"
            NetworkType = "Nur Public"
            TCPUDPType = "Nur TCP"
            ProcessInfo = $true
        }
        "Interne Services" = @{
            Protocol = "Alle"
            Direction = "Beide"
            Category = "Alle Kategorien"
            NetworkType = "Nur Private"
            TCPUDPType = "Nur TCP"
            ProcessInfo = $true
        }
        "Externe Verbindungen" = @{
            Protocol = "Alle"
            Direction = "Beide"
            Category = "Alle Kategorien"
            NetworkType = "Nur Public"
            TCPUDPType = "TCP + UDP"
            ProcessInfo = $true
        }
        "DNS Monitoring" = @{
            Protocol = "DNS"
            Direction = "Beide"
            Category = "Monitoring"
            NetworkType = "Alle Netzwerke"
            TCPUDPType = "Nur UDP"
            ProcessInfo = $true
        }
        "Email Services" = @{
            Protocol = "SMTP"
            Direction = "Beide"
            Category = "Email Services"
            NetworkType = "Alle Netzwerke"
            TCPUDPType = "Nur TCP"
            ProcessInfo = $true
        }
    }
    
    # Load saved presets from JSON file
    if (Test-Path $Global:PresetsFile) {
        try {
            $savedPresets = Get-Content $Global:PresetsFile | ConvertFrom-Json -AsHashtable
            
            # Start with default presets
            $Global:FilterPresets = [hashtable]$defaultPresets
            
            # Add or override with saved presets (custom user presets)
            foreach ($key in $savedPresets.Keys) {
                $Global:FilterPresets[$key] = $savedPresets[$key]
            }
        } catch {
            # If JSON loading fails, use defaults and save them
            $Global:FilterPresets = [hashtable]$defaultPresets
            Export-FilterPresets
        }
    } else {
        # Use defaults and save them
        $Global:FilterPresets = [hashtable]$defaultPresets
        Export-FilterPresets
    }
}

function Export-FilterPresets {
    $Global:FilterPresets | ConvertTo-Json | Out-File -FilePath $Global:PresetsFile -Encoding UTF8
}

function Get-CurrentFilterState {
    return @{
        Protocol = $cmbProtocol.SelectedItem.Content
        Direction = $cmbDirection.SelectedItem.Content
        Category = $cmbCategory.SelectedItem.Content
        NetworkType = $cmbNetworkType.SelectedItem.Content
        TCPUDPType = $cmbTCPUDP.SelectedItem.Content
        ProcessInfo = $chkShowProcessInfo.IsChecked
    }
}

function Set-FilterPreset {
    param([hashtable]$preset)
    
    # Apply preset values to UI
    for ($i = 0; $i -lt $cmbProtocol.Items.Count; $i++) {
        if ($cmbProtocol.Items[$i].Content -eq $preset.Protocol) {
            $cmbProtocol.SelectedIndex = $i
            break
        }
    }
    for ($i = 0; $i -lt $cmbDirection.Items.Count; $i++) {
        if ($cmbDirection.Items[$i].Content -eq $preset.Direction) {
            $cmbDirection.SelectedIndex = $i
            break
        }
    }
    for ($i = 0; $i -lt $cmbCategory.Items.Count; $i++) {
        if ($cmbCategory.Items[$i].Content -eq $preset.Category) {
            $cmbCategory.SelectedIndex = $i
            break
        }
    }
    for ($i = 0; $i -lt $cmbNetworkType.Items.Count; $i++) {
        if ($cmbNetworkType.Items[$i].Content -eq $preset.NetworkType) {
            $cmbNetworkType.SelectedIndex = $i
            break
        }
    }
    for ($i = 0; $i -lt $cmbTCPUDP.Items.Count; $i++) {
        if ($cmbTCPUDP.Items[$i].Content -eq $preset.TCPUDPType) {
            $cmbTCPUDP.SelectedIndex = $i
            break
        }
    }
    $chkShowProcessInfo.IsChecked = $preset.ProcessInfo
    
    # Automatically refresh with new filters
    $btnRefresh.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
}

# Load saved presets
Import-FilterPresets

# Function to populate presets in ComboBox
function Update-PresetsComboBox {
    # Clear existing items except "-- Neues Preset --"
    while ($cmbPresets.Items.Count -gt 1) {
        $cmbPresets.Items.RemoveAt($cmbPresets.Items.Count - 1)
    }
    
    # Add all presets to ComboBox
    foreach ($presetName in $Global:FilterPresets.Keys | Sort-Object) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $presetName
        $cmbPresets.Items.Add($item) | Out-Null
    }
}

# Populate presets in ComboBox
Update-PresetsComboBox

# Preset Button Event Handlers
$btnLoadPreset.Add_Click({
    $selectedPresetName = $cmbPresets.SelectedItem.Content
    
    if ($selectedPresetName -eq "-- Neues Preset --") {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie ein Preset aus der Liste.", "Information", "OK", "Information")
        return
    }
    
    if ($Global:FilterPresets.ContainsKey($selectedPresetName)) {
        Set-FilterPreset -preset $Global:FilterPresets[$selectedPresetName]
        $txtStatus.Text = "Preset geladen: $selectedPresetName"
    } else {
        [System.Windows.MessageBox]::Show("Preset '$selectedPresetName' nicht gefunden.", "Fehler", "OK", "Error")
    }
})

$btnSavePreset.Add_Click({
    [System.Windows.Input.InputDialog]::Show("Preset-Name:", "Neues Preset speichern")
    
    $inputBox = [Microsoft.VisualBasic.Interaction]::InputBox("Geben Sie einen Namen für das Preset ein:", "Preset speichern")
    
    if ([string]::IsNullOrWhiteSpace($inputBox)) {
        return
    }
    
    $currentFilters = Get-CurrentFilterState
    $Global:FilterPresets[$inputBox] = $currentFilters
    Export-FilterPresets
    
    # Add to ComboBox
    if (-not ($cmbPresets.Items | Where-Object { $_.Content -eq $inputBox })) {
        $cmbPresets.Items.Add((New-Object System.Windows.Controls.ComboBoxItem -Property @{ Content = $inputBox }))
    }
    
    $txtStatus.Text = "Preset gespeichert: $inputBox"
    [System.Windows.MessageBox]::Show("Preset '$inputBox' wurde gespeichert.", "Erfolg", "OK", "Information")
})

$btnDeletePreset.Add_Click({
    $selectedPresetName = $cmbPresets.SelectedItem.Content
    
    if ($selectedPresetName -eq "-- Neues Preset --") {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie ein Preset zum Löschen aus.", "Information", "OK", "Information")
        return
    }
    
    $result = [System.Windows.MessageBox]::Show("Möchten Sie das Preset '$selectedPresetName' wirklich löschen?", "Bestätigung", "YesNo", "Question")
    
    if ($result -eq "Yes") {
        $Global:FilterPresets.Remove($selectedPresetName)
        Export-FilterPresets
        
        # Remove from ComboBox
        for ($i = 0; $i -lt $cmbPresets.Items.Count; $i++) {
            if ($cmbPresets.Items[$i].Content -eq $selectedPresetName) {
                $cmbPresets.Items.RemoveAt($i)
                break
            }
        }
        
        $cmbPresets.SelectedIndex = 0
        $txtStatus.Text = "Preset gelöscht: $selectedPresetName"
    }
})

# Show window
$window.ShowDialog() | Out-Null

Write-Log "Application closed"

# SIG # Begin signature block
# MIIoiQYJKoZIhvcNAQcCoIIoejCCKHYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAfs/7DJuvb2Ig5
# 0eBtBhAHL4+dtGXVbygDWti9meCqSaCCILswggXJMIIEsaADAgECAhAbtY8lKt8j
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
# BCAUsJT/9gaZbD1E3yIruDy/KstQUu5Wl6zrpYvWcswLIjANBgkqhkiG9w0BAQEF
# AASCAgAsQWYqJNJY1xjA1PQfLTSe+QW1mKjVsZaWWWCkbbKdA5vhi+Uh76bzHFSN
# 1Yh/Z9GUgncpr+aD8+t3jZLr2axTJFMYVT4RPKQ+nBK5z4lbLd3Yd4cmDi+s6US2
# Qg5rorF/jiOvB//NXGVQwBnbU/gyS9kMw2GDEi1yiYKBiZD3H6VxHyqAATHYG8a4
# 6sgeN7Wo4RR/rpCVMmw160OGnfCrtKfvPKhHpPVltGj1zC21CwZs1vi/XOu/v7I1
# NM7EMZ/wAxtDoegLNQrRJZOsmAWXoPkC0S7iX+4I1nUEWgnuXtoGWVLMnXEbURHQ
# gWsFf/rz7Ps7OiN7Sz2d3RWzVQNFhulQhxU2moNHyfkfCegQFhVQ0tEm88JGJzC+
# yZGuX6eTw20vPZI6MKMLg8jZ+x3JERGLs4gm6VjTh7l0XebiaVZd1LgKDnwAW8Cj
# 2n44s5aMXaFcBmX0VjauAyJbmZw5NSNhCPcq13+oFubbrVnC9LZU5N58ZfU6L69m
# CxYCzoPIFPCU80fK+gaWXsKNGrESPJh17vup/rlevE9cBDmSwcB6AplbU8TkM3OT
# f7XY8spHJjR8B4X614eMfUBRgoSZMfdHWR0lWMBiI0EXPoviDfJCSlQuee9u+W9i
# aF8iwCbqq3CUS0fo1yopZZ48ld2waMqwPRa/NPPCac7C+NRaeKGCBAQwggQABgkq
# hkiG9w0BCQYxggPxMIID7QIBATBrMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhB
# c3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBUaW1lc3Rh
# bXBpbmcgMjAyMSBDQQIRAJ6cBPZVqLSnAm1JjGx4jaowDQYJYIZIAWUDBAICBQCg
# ggFXMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcN
# MjUxMDI1MTc1ODAxWjA3BgsqhkiG9w0BCRACLzEoMCYwJDAiBCDPodw1ne0rw8uJ
# D6Iw5dr3e1QPGm4rI93PF1ThjPqg1TA/BgkqhkiG9w0BCQQxMgQwO68MVbNd8xM9
# C4eZbZKKEgnzwQkd6eD/RNwdcvs1bdpvbZOdXQ+1BSP+JU/eL+/PMIGgBgsqhkiG
# 9w0BCRACDDGBkDCBjTCBijCBhwQUwyW4mxf8xQJgYc4rcXtFB92camowbzBapFgw
# VjELMAkGA1UEBhMCUEwxITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5B
# LjEkMCIGA1UEAxMbQ2VydHVtIFRpbWVzdGFtcGluZyAyMDIxIENBAhEAnpwE9lWo
# tKcCbUmMbHiNqjANBgkqhkiG9w0BAQEFAASCAgBSN+06hcqD/g7H0XfWH4vjuMUh
# sBIIjYjbLVT6EOYqv8OM2JSalg5PWrtMlTgljpXFXTdsQzQsZtP3xYUUV3tY0UlF
# GTKs6Mfe744m/p4TG4RIYiHjWkfQgDEAKE+qRcLAZwoMslIj3xF/A+f8b+rFaNPq
# P4Ypc6ZwmacAnh5md3DLNdO5usJNEXKfCQrHe6pMLVC4b4yylbIs6EYAHjZyAPIh
# oQGiiW4bS7/jEAOxgEGHUoDD09ypYK1qWOukg17AU/SNzcT8zdzYvtZMrFU4ld2k
# 62Yfd/XxNll5TcCK9LsD/R9o/gWZ/Lyg77ADduT9/qSBdjX5gEZrYPQ5T2VwvFxf
# aeRzwnTgy/znMHXUsNTKB7wXaBjaC3uQWWDuRwtfN4aCcToD6fcAN2NUzxFQt0Bi
# EGAz2jGUe6TrMuow9D9Vp7BZldL8zjbPXlw6+xflCsuYeMgqIFMlLLlp4x8Pf9JW
# PC0jBICTmVAmUBsIHfBP4XTc5VlD2OgzRnzny7zg9K6EKf5IzErv1wj3esfwI8Ol
# fkTUn5n7gSmeGBRKjeeVvjj+SsvowwpzCaSleMp1xVQtJboQjPpHyL8A9KO4S5IK
# 2IUkBdthEm6QZLpTEBdjim1UW7ZFiKPkMdu+BHOy82l8jwRm8DbNss9F23IsvnQk
# 9Xq2HdI/J8cWW47Onw==
# SIG # End signature block
