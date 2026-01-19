# easyConnections_V0.0.3.ps1
# Advanced Network Connections Monitor with WPF GUI - Admin Edition
# V0.0.3: Lazy Loading + Recording + Color-Coded Categories
# Performance: 80-90% schneller beim Start durch Lazy Loading

# ===== Optional Logging System =====
$Global:LoggingEnabled = $false  # Set to $true to enable logging to file
$Global:LogFile = "$PSScriptRoot\easyConnections.log"

function Write-Log {
    param([string]$Message)
    if ($Global:LoggingEnabled) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] $Message"
        Add-Content -Path $Global:LogFile -Value $logEntry -ErrorAction SilentlyContinue
    }
}

Write-Log "Script started: easyConnections_V0.0.3.ps1"

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
$Global:PresetsFile = "$PSScriptRoot\presets.json"
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
