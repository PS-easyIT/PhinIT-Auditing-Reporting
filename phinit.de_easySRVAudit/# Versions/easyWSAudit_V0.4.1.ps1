<#
================================================================================
    easyWSAudit - Windows Server Audit Tool (Redesigned)
    Version: 0.4.0
    
    Beschreibung:
    Umfassendes Server-Audit-Tool mit moderner WPF-GUI.
    Bietet detaillierte Analysen von Systemkonfiguration, Sicherheit, Diensten,
    Rollen und Replikation mit Export-Funktionen (CSV, HTML).
    
    Design: Sidebar-Menü (Links) mit Audit-Kategorien + Output-Bereich (Rechts)
    
================================================================================
#>

[xml]$Global:XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="easyWSAudit v0.4.0 - Windows Server Audit Tool"
    Width="1850"
    Height="1000"
    Background="#F8FAFC"
    FontFamily="Segoe UI"
    ResizeMode="CanResizeWithGrip"
    WindowStartupLocation="CenterScreen">

    <!--  Window Resources for Modern Styling  -->
    <Window.Resources>
        <!--  Modern Card Style  -->
        <Style x:Key="ModernCard" TargetType="Border">
            <Setter Property="Background" Value="White" />
            <Setter Property="CornerRadius" Value="6" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="BorderBrush" Value="#E5E7EB" />
        </Style>

        <!--  Primary Button Style  -->
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#6366F1" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="24,14" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="FontSize" Value="14" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border
                            x:Name="border"
                            Background="{TemplateBinding Background}"
                            CornerRadius="4">
                            <ContentPresenter
                                Margin="{TemplateBinding Padding}"
                                HorizontalAlignment="Center"
                                VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#5B5BD6" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#4F46E5" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!--  Category Header Style  -->
        <Style x:Key="CategoryHeader" TargetType="TextBlock">
            <Setter Property="FontSize" Value="14" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="Foreground" Value="#1F2937" />
            <Setter Property="Margin" Value="0,8,0,4" />
        </Style>

        <!--  Sidebar Menu Button Style  -->
        <Style x:Key="SidebarButton" TargetType="Button">
            <Setter Property="Background" Value="Transparent" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="12,8" />
            <Setter Property="FontWeight" Value="Normal" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="HorizontalAlignment" Value="Stretch" />
            <Setter Property="HorizontalContentAlignment" Value="Left" />
            <Setter Property="Margin" Value="0,1" />
            <Setter Property="Foreground" Value="#374151" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border
                            x:Name="border"
                            Padding="{TemplateBinding Padding}"
                            Background="{TemplateBinding Background}"
                            CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F3F4F6" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E5E7EB" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!--  Active Sidebar Button Style  -->
        <Style
            x:Key="SidebarButtonActive"
            BasedOn="{StaticResource SidebarButton}"
            TargetType="Button">
            <Setter Property="Background" Value="#EEF2FF" />
            <Setter Property="Foreground" Value="#6366F1" />
            <Setter Property="FontWeight" Value="Medium" />
        </Style>

        <!--  Expandable Section Style  -->
        <Style x:Key="ExpanderStyle" TargetType="Expander">
            <Setter Property="IsExpanded" Value="True" />
            <Setter Property="Margin" Value="0,2,0,6" />
            <Setter Property="Background" Value="#FFFFFF" />
            <Setter Property="BorderBrush" Value="#E5E7EB" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="FontWeight" Value="Medium" />
            <Setter Property="Foreground" Value="#374151" />
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="50" />
            <!--  Header  -->
            <RowDefinition Height="*" />
            <!--  Content Area  -->
            <RowDefinition Height="25" />
            <!--  Footer  -->
        </Grid.RowDefinitions>

        <!--  Header (Grid.Row="0")  -->
        <Border
            Grid.Row="0"
            Background="#374151"
            BorderBrush="#E5E7EB"
            BorderThickness="0,0,0,1">
            <Grid Margin="20,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" />
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>

                <!--  App Title  -->
                <StackPanel
                    Grid.Column="0"
                    VerticalAlignment="Center"
                    Orientation="Horizontal">
                    <TextBlock
                        FontSize="20"
                        FontWeight="SemiBold"
                        Foreground="#F9FAFB"
                        Text="⚙️ easyWSAudit v0.4.0 - Windows Server Audit" />
                </StackPanel>

                <!--  Status Bar  -->
                <StackPanel
                    Grid.Column="1"
                    VerticalAlignment="Center"
                    Orientation="Vertical"
                    Background="#FF4C5B73"
                    Height="49"
                    Margin="1475,0,-20,0"
                    Grid.ColumnSpan="2">
                    <TextBlock Text="Results" FontSize="14" Foreground="#FFC1FFDB" HorizontalAlignment="Center"/>
                    <TextBlock x:Name="TotalResultCountText" Text="0" FontSize="20" FontWeight="Bold" 
                               Foreground="#FFEAEAEA" HorizontalAlignment="Center" Margin="0,2,0,0" Width="60" TextAlignment="Center"/>
                </StackPanel>
            </Grid>
        </Border>

        <!--  Content Area (Grid.Row="1")  -->
        <Grid Grid.Row="1" Margin="24,20,24,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="320" />
                <!--  Sidebar  -->
                <ColumnDefinition Width="24" />
                <!--  Spacing  -->
                <ColumnDefinition Width="*" />
                <!--  Main Content  -->
            </Grid.ColumnDefinitions>

            <!--  Modern Sidebar with Categorized Reports  -->
            <ScrollViewer
                Grid.Column="0"
                Margin="0,0,0,10"
                HorizontalScrollBarVisibility="Disabled"
                VerticalScrollBarVisibility="Auto">
                <Border
                    Padding="16,16"
                    Background="#F8FAFC"
                    BorderBrush="#E5E7EB"
                    BorderThickness="1"
                    CornerRadius="6">
                    <StackPanel>
                        <!--  Sidebar Header  -->
                        <Grid Margin="0,0,0,16">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto" />
                                <RowDefinition Height="Auto" />
                            </Grid.RowDefinitions>
                            <TextBlock
                                Grid.Row="0"
                                Margin="0,0,0,8"
                                FontSize="16"
                                FontWeight="Bold"
                                Foreground="#374151"
                                Text="📊 Audit Reports" />
                        </Grid>

                        <!--  System Information  -->
                        <Expander
                            Header="🖥️ System Information"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonSystemInfo" Content="System Overview" Style="{StaticResource SidebarButton}" ToolTip="Retrieve system information" />
                                <Button x:Name="ButtonOSInfo" Content="OS Details" Style="{StaticResource SidebarButton}" ToolTip="Get operating system details" />
                                <Button x:Name="ButtonHardwareInfo" Content="Hardware Summary" Style="{StaticResource SidebarButton}" ToolTip="Display hardware information" />
                                <Button x:Name="ButtonCPUInfo" Content="CPU Details" Style="{StaticResource SidebarButton}" ToolTip="CPU specifications" />
                                <Button x:Name="ButtonMemoryInfo" Content="Memory Details" Style="{StaticResource SidebarButton}" ToolTip="Memory information" />
                                <Button x:Name="ButtonStorageInfo" Content="Storage Summary" Style="{StaticResource SidebarButton}" ToolTip="Disk and volume information" />
                            </StackPanel>
                        </Expander>

                        <!--  Network Configuration  -->
                        <Expander
                            Header="🌐 Network Configuration"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonNetConfig" Content="IP Configuration" Style="{StaticResource SidebarButton}" ToolTip="Network IP settings" />
                                <Button x:Name="ButtonNetAdapters" Content="Network Adapters" Style="{StaticResource SidebarButton}" ToolTip="Network adapter status" />
                                <Button x:Name="ButtonTCPConnections" Content="Active Connections" Style="{StaticResource SidebarButton}" ToolTip="Listen ports and connections" />
                                <Button x:Name="ButtonFirewallRules" Content="Firewall Rules" Style="{StaticResource SidebarButton}" ToolTip="Active firewall rules" />
                            </StackPanel>
                        </Expander>

                        <!--  Services & Tasks  -->
                        <Expander
                            Header="⚙️ Services &amp; Tasks"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonAutomaticServices" Content="Automatic Services" Style="{StaticResource SidebarButton}" ToolTip="Services set to automatic" />
                                <Button x:Name="ButtonRunningServices" Content="Running Services" Style="{StaticResource SidebarButton}" ToolTip="Currently running services" />
                                <Button x:Name="ButtonScheduledTasks" Content="Scheduled Tasks" Style="{StaticResource SidebarButton}" ToolTip="Ready scheduled tasks" />
                            </StackPanel>
                        </Expander>

                        <!--  Roles &amp; Features  -->
                        <Expander
                            Header="📦 Roles &amp; Features"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonInstalledFeatures" Content="Installed Features" Style="{StaticResource SidebarButton}" ToolTip="Windows features in use" />
                                <Button x:Name="ButtonInstalledPrograms" Content="Installed Programs" Style="{StaticResource SidebarButton}" ToolTip="Software inventory" />
                                <Button x:Name="ButtonWindowsUpdates" Content="Recent Updates" Style="{StaticResource SidebarButton}" ToolTip="Latest Windows patches" />
                            </StackPanel>
                        </Expander>

                        <!--  IIS  -->
                        <Expander
                            Header="🌐 IIS Web Server"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonIISWebsites" Content="Websites" Style="{StaticResource SidebarButton}" ToolTip="IIS websites" />
                                <Button x:Name="ButtonIISAppPools" Content="Application Pools" Style="{StaticResource SidebarButton}" ToolTip="App pool configuration" />
                                <Button x:Name="ButtonIISBindings" Content="SSL Bindings" Style="{StaticResource SidebarButton}" ToolTip="SSL certificates and bindings" />
                            </StackPanel>
                        </Expander>

                        <!--  RDS / WTS  -->
                        <Expander
                            Header="🖥️ Remote Desktop Services"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonRDSCollections" Content="Session Collections" Style="{StaticResource SidebarButton}" ToolTip="RDS session collections" />
                                <Button x:Name="ButtonRDSSessionHosts" Content="Session Hosts" Style="{StaticResource SidebarButton}" ToolTip="RDS session hosts" />
                                <Button x:Name="ButtonRDSLicensing" Content="RDS Licensing" Style="{StaticResource SidebarButton}" ToolTip="RDS licensing configuration" />
                            </StackPanel>
                        </Expander>

                        <!--  DFS  -->
                        <Expander
                            Header="🔀 Distributed File System"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonDFSNamespaces" Content="DFS Namespaces" Style="{StaticResource SidebarButton}" ToolTip="DFS namespace configuration" />
                                <Button x:Name="ButtonDFSReplication" Content="Replication Groups" Style="{StaticResource SidebarButton}" ToolTip="DFS replication groups" />
                            </StackPanel>
                        </Expander>

                        <!--  Print Server  -->
                        <Expander
                            Header="🖨️ Print Server"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonPrinters" Content="Printers" Style="{StaticResource SidebarButton}" ToolTip="Print server printers" />
                                <Button x:Name="ButtonPrinterDrivers" Content="Printer Drivers" Style="{StaticResource SidebarButton}" ToolTip="Installed printer drivers" />
                            </StackPanel>
                        </Expander>

                        <!--  WSUS  -->
                        <Expander
                            Header="🔄 WSUS (Update Services)"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonWSUSConfig" Content="WSUS Configuration" Style="{StaticResource SidebarButton}" ToolTip="WSUS server configuration" />
                                <Button x:Name="ButtonWSUSGroups" Content="Computer Groups" Style="{StaticResource SidebarButton}" ToolTip="WSUS computer target groups" />
                                <Button x:Name="ButtonWSUSUpdates" Content="Available Updates" Style="{StaticResource SidebarButton}" ToolTip="Available updates in WSUS" />
                            </StackPanel>
                        </Expander>

                        <!--  Hyper-V  -->
                        <Expander
                            Header="🔧 Hyper-V Virtualization"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonHyperVVMs" Content="Virtual Machines" Style="{StaticResource SidebarButton}" ToolTip="Hyper-V virtual machines" />
                                <Button x:Name="ButtonHyperVSwitches" Content="Virtual Switches" Style="{StaticResource SidebarButton}" ToolTip="Hyper-V virtual network switches" />
                                <Button x:Name="ButtonHyperVSnapshots" Content="Snapshots" Style="{StaticResource SidebarButton}" ToolTip="VM snapshots" />
                            </StackPanel>
                        </Expander>

                        <!--  NRAS / NPS  -->
                        <Expander
                            Header="🔐 Network Access Protection"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonNPASConfig" Content="NRAS/NPS Configuration" Style="{StaticResource SidebarButton}" ToolTip="Network Policy Server configuration" />
                                <Button x:Name="ButtonNASClients" Content="RADIUS NAS Clients" Style="{StaticResource SidebarButton}" ToolTip="RADIUS network access servers" />
                            </StackPanel>
                        </Expander>

                        <!--  KMS (Volume Activation)  -->
                        <Expander
                            Header="🔑 KMS (Volume Activation)"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonKMSConfig" Content="KMS Configuration" Style="{StaticResource SidebarButton}" ToolTip="Key Management Service configuration" />
                            </StackPanel>
                        </Expander>

                        <!--  WDS (Windows Deployment)  -->
                        <Expander
                            Header="📦 WDS (Deployment Services)"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonWDSConfig" Content="WDS Configuration" Style="{StaticResource SidebarButton}" ToolTip="Windows Deployment Services configuration" />
                                <Button x:Name="ButtonWDSBootImages" Content="Boot Images" Style="{StaticResource SidebarButton}" ToolTip="WDS boot images" />
                                <Button x:Name="ButtonWDSInstallImages" Content="Install Images" Style="{StaticResource SidebarButton}" ToolTip="WDS install images" />
                            </StackPanel>
                        </Expander>

                        <!--  File Services &amp; SMB  -->
                        <Expander
                            Header="📁 File Services (SMB)"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonFileShares" Content="File Shares" Style="{StaticResource SidebarButton}" ToolTip="SMB file shares" />
                                <Button x:Name="ButtonSharePermissions" Content="Share Permissions" Style="{StaticResource SidebarButton}" ToolTip="Share ACLs and permissions" />
                                <Button x:Name="ButtonFileQuotas" Content="File Quotas (FSRM)" Style="{StaticResource SidebarButton}" ToolTip="File Server Resource Manager quotas" />
                                <Button x:Name="ButtonShadowCopies" Content="Shadow Copies" Style="{StaticResource SidebarButton}" ToolTip="Volume shadow copies / snapshots" />
                            </StackPanel>
                        </Expander>

                        <!--  Advanced Active Directory  -->
                        <Expander
                            Header="🌳 AD Advanced Info"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonADDCExtended" Content="Domain Controllers (Extended)" Style="{StaticResource SidebarButton}" ToolTip="Detailed DC information" />
                                <Button x:Name="ButtonADFunctionalLevels" Content="Functional Levels" Style="{StaticResource SidebarButton}" ToolTip="Domain and forest functional levels" />
                                <Button x:Name="ButtonADSites" Content="Sites &amp; Subnets" Style="{StaticResource SidebarButton}" ToolTip="AD sites and subnets configuration" />
                                <Button x:Name="ButtonADGPO" Content="Group Policy Summary" Style="{StaticResource SidebarButton}" ToolTip="GPO overview and statistics" />
                                <Button x:Name="ButtonADClustering" Content="Failover Clustering" Style="{StaticResource SidebarButton}" ToolTip="Failover cluster configuration" />
                            </StackPanel>
                        </Expander>

                        <!--  Event Logs  -->
                        <Expander
                            Header="📋 Event Logs"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonSystemEvents" Content="System Events (24h)" Style="{StaticResource SidebarButton}" ToolTip="System event log" />
                                <Button x:Name="ButtonAppEvents" Content="Application Events (24h)" Style="{StaticResource SidebarButton}" ToolTip="Application event log" />
                                <Button x:Name="ButtonSecurityEvents" Content="Security Events (100)" Style="{StaticResource SidebarButton}" ToolTip="Security event log" />
                                <Button x:Name="ButtonFailedLogons" Content="Failed Logon Attempts" Style="{StaticResource SidebarButton}" ToolTip="Event ID 4625" />
                                <Button x:Name="ButtonAccountLockouts" Content="Account Lockouts" Style="{StaticResource SidebarButton}" ToolTip="Event ID 4740" />
                            </StackPanel>
                        </Expander>

                        <!--  Security  -->
                        <Expander
                            Header="🔒 Security &amp; Users"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonLocalUsers" Content="Local Users" Style="{StaticResource SidebarButton}" ToolTip="Local user accounts" />
                                <Button x:Name="ButtonLocalGroups" Content="Local Groups" Style="{StaticResource SidebarButton}" ToolTip="Local group membership" />
                                <Button x:Name="ButtonPrivilegeAudit" Content="Privilege Use Audit" Style="{StaticResource SidebarButton}" ToolTip="Event IDs 4672/4673/4674" />
                            </StackPanel>
                        </Expander>

                        <!--  Active Directory  -->
                        <Expander
                            Header="🌳 Active Directory"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonADDC" Content="Domain Controllers" Style="{StaticResource SidebarButton}" ToolTip="AD domain controller status" />
                                <Button x:Name="ButtonADDomain" Content="Domain Info" Style="{StaticResource SidebarButton}" ToolTip="AD domain properties" />
                                <Button x:Name="ButtonADForest" Content="Forest Info" Style="{StaticResource SidebarButton}" ToolTip="AD forest properties" />
                                <Button x:Name="ButtonADOUs" Content="Organizational Units" Style="{StaticResource SidebarButton}" ToolTip="AD OUs" />
                                <Button x:Name="ButtonADAdmins" Content="Domain Admins" Style="{StaticResource SidebarButton}" ToolTip="Domain admin members" />
                                <Button x:Name="ButtonADComputers" Content="Computer Accounts" Style="{StaticResource SidebarButton}" ToolTip="AD computer objects" />
                                <Button x:Name="ButtonADReplStatus" Content="Replication Status" Style="{StaticResource SidebarButton}" ToolTip="AD replication status" />
                                <Button x:Name="ButtonADTrusts" Content="Trust Relationships" Style="{StaticResource SidebarButton}" ToolTip="AD domain trusts" />
                            </StackPanel>
                        </Expander>

                        <!--  DNS  -->
                        <Expander
                            Header="🔗 DNS Server"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonDNSConfig" Content="DNS Configuration" Style="{StaticResource SidebarButton}" ToolTip="DNS server settings" />
                                <Button x:Name="ButtonDNSZones" Content="DNS Zones" Style="{StaticResource SidebarButton}" ToolTip="DNS zones" />
                                <Button x:Name="ButtonDNSForwarders" Content="DNS Forwarders" Style="{StaticResource SidebarButton}" ToolTip="DNS forwarders" />
                                <Button x:Name="ButtonDNSCache" Content="DNS Cache" Style="{StaticResource SidebarButton}" ToolTip="DNS cache entries" />
                            </StackPanel>
                        </Expander>

                        <!--  DHCP  -->
                        <Expander
                            Header="📡 DHCP Server"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonDHCPConfig" Content="DHCP Configuration" Style="{StaticResource SidebarButton}" ToolTip="DHCP server settings" />
                                <Button x:Name="ButtonDHCPv4Scopes" Content="IPv4 Scopes" Style="{StaticResource SidebarButton}" ToolTip="DHCP IPv4 scopes" />
                                <Button x:Name="ButtonDHCPv6Scopes" Content="IPv6 Scopes" Style="{StaticResource SidebarButton}" ToolTip="DHCP IPv6 scopes" />
                                <Button x:Name="ButtonDHCPReservations" Content="DHCP Reservations" Style="{StaticResource SidebarButton}" ToolTip="DHCP reservations" />
                            </StackPanel>
                        </Expander>

                        <!--  Export Options  -->
                        <StackPanel Margin="0,20,0,0">
                            <Button
                                x:Name="ButtonExportCSV"
                                Content="📥 Export to CSV"
                                Style="{StaticResource PrimaryButton}"
                                Margin="0,0,0,8" />
                            <Button
                                x:Name="ButtonClearOutput"
                                Content="🗑️ Clear Output"
                                Style="{StaticResource PrimaryButton}" />
                        </StackPanel>
                    </StackPanel>
                </Border>
            </ScrollViewer>

            <!--  Output Area  -->
            <Border
                Grid.Column="2"
                Padding="20"
                Background="White"
                BorderBrush="#E5E7EB"
                BorderThickness="1"
                CornerRadius="6">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="*" />
                    </Grid.RowDefinitions>

                    <TextBlock
                        Grid.Row="0"
                        Margin="0,0,0,12"
                        FontSize="14"
                        FontWeight="SemiBold"
                        Foreground="#374151"
                        Text="📊 Audit Results" />

                    <TextBox
                        x:Name="OutputTextBox"
                        Grid.Row="1"
                        Padding="12"
                        Background="#F9FAFB"
                        BorderBrush="#E5E7EB"
                        BorderThickness="1"
                        FontFamily="Consolas"
                        FontSize="11"
                        Foreground="#1F2937"
                        IsReadOnly="True"
                        TextWrapping="Wrap"
                        VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Auto" />
                </Grid>
            </Border>
        </Grid>

        <!--  Footer (Grid.Row="2")  -->
        <Border
            Grid.Row="2"
            Background="#E5E7EB"
            Padding="12,4">
            <TextBlock
                x:Name="StatusBarText"
                FontSize="11"
                Foreground="#6B7280"
                Text="Ready" />
        </Border>
    </Grid>
</Window>
"@

# Add required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Create WPF window
$reader = New-Object System.Xml.XmlNodeReader $Global:XAML
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Get control references
$OutputTextBox = $window.FindName("OutputTextBox")
$StatusBarText = $window.FindName("StatusBarText")
$TotalResultCountText = $window.FindName("TotalResultCountText")

# Variables for UI state
$script:ResultsCount = 0
$script:AllResults = @()
$script:CurrentButton = $null

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

function Update-Output {
    param([string]$Text)
    $OutputTextBox.Text = $Text
    [int]$lineCount = ($Text | Measure-Object -Line).Lines
    $script:ResultsCount = $lineCount
    $TotalResultCountText.Text = $lineCount.ToString()
}

function Add-OutputText {
    param([string]$Text)
    $OutputTextBox.Text += $Text
    [int]$lineCount = ($OutputTextBox.Text | Measure-Object -Line).Lines
    $script:ResultsCount = $lineCount
    $TotalResultCountText.Text = $lineCount.ToString()
}

function Clear-Output {
    $OutputTextBox.Text = ""
    $script:ResultsCount = 0
    $TotalResultCountText.Text = "0"
}

function Update-Status {
    param([string]$Status)
    $StatusBarText.Text = $Status
    $window.Dispatcher.Invoke([Action]{}, "Background")
}

# ============================================================================
# TABLE FORMATTING FUNCTIONS
# ============================================================================

function Format-TableOutput {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Title,
        
        [Parameter(Mandatory=$false)]
        [PSObject[]]$Data,
        
        [Parameter(Mandatory=$false)]
        [string[]]$Columns
    )
    
    $output = @()
    $output += ""
    $output += "╔" + ("═" * 100) + "╗"
    $output += "║ $($Title.PadRight(98)) ║"
    $output += "╠" + ("═" * 100) + "╣"
    
    if (-not $Data -or $Data.Count -eq 0) {
        $output += "║ Keine Daten verfügbar" + (" " * 79) + "║"
        $output += "╚" + ("═" * 100) + "╝"
        return ($output -join "`n")
    }
    
    # Wenn Columns nicht angegeben, automatisch ermitteln
    if (-not $Columns) {
        $Columns = @($Data[0].PSObject.Properties | Select-Object -ExpandProperty Name)
    }
    
    # Spalten-Header
    $headerRow = "║ " + ($Columns | ForEach-Object { $_.PadRight(18) } | Join-String -Separator " | ") + " ║"
    $output += $headerRow
    $output += "╠" + ("═" * 100) + "╣"
    
    # Daten-Zeilen
    foreach ($item in $Data) {
        $row = "║ "
        foreach ($col in $Columns) {
            $value = $item.$col
            if ($null -eq $value) { $value = "-" }
            $value = $value.ToString().Substring(0, [Math]::Min($value.ToString().Length, 16))
            $row += $value.PadRight(18) + " | "
        }
        $row = $row.Substring(0, $row.Length - 3) + " ║"
        $output += $row
    }
    
    $output += "╚" + ("═" * 100) + "╝"
    $output += ""
    
    return ($output -join "`n")
}

function Format-SimpleTable {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Title,
        
        [Parameter(Mandatory=$false)]
        [Hashtable[]]$Data
    )
    
    $output = @()
    $output += ""
    $output += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    $output += "  $Title"
    $output += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    
    if (-not $Data -or $Data.Count -eq 0) {
        $output += "  ⚠ Keine Daten verfügbar"
        $output += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        return ($output -join "`n")
    }
    
    foreach ($item in $Data) {
        foreach ($key in $item.Keys) {
            $value = $item[$key]
            $output += "  $($key): $value"
        }
        $output += ""
    }
    
    $output += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    $output += ""
    
    return ($output -join "`n")
}

# ============================================================================
# AUDIT FUNCTIONS
# ============================================================================

function Get-SystemInformation {
    try {
        $info = Get-ComputerInfo -ErrorAction Stop
        $data = @(
            @{
                "Attribut" = "Computername"
                "Wert" = $info.CsComputerName
            },
            @{
                "Attribut" = "Domäne"
                "Wert" = $info.CsDomain
            },
            @{
                "Attribut" = "Betriebssystem"
                "Wert" = $info.OsName
            },
            @{
                "Attribut" = "Installationsdatum"
                "Wert" = $info.OsInstallDate
            },
            @{
                "Attribut" = "Letzter Boot"
                "Wert" = $info.OsLastBootUpTime
            }
        )
        return Format-SimpleTable -Title "📊 SYSTEM INFORMATION" -Data $data
    } catch {
        return Format-SimpleTable -Title "📊 SYSTEM INFORMATION" -Data @(@{"Fehler" = $_.Exception.Message})
    }
}

function Get-OSDetails {
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $data = @(
            @{
                "Eigenschaft" = "Caption"
                "Wert" = $os.Caption
            },
            @{
                "Eigenschaft" = "Version"
                "Wert" = $os.Version
            },
            @{
                "Eigenschaft" = "Build"
                "Wert" = $os.BuildNumber
            },
            @{
                "Eigenschaft" = "Gesamtarbeitsspeicher"
                "Wert" = "$('{0:N0}' -f ($os.TotalVisibleMemorySize / 1024)) MB"
            },
            @{
                "Eigenschaft" = "Freier Arbeitsspeicher"
                "Wert" = "$('{0:N0}' -f ($os.FreePhysicalMemory / 1024)) MB"
            },
            @{
                "Eigenschaft" = "Letzter Boot"
                "Wert" = $os.LastBootUpTime
            }
        )
        return Format-SimpleTable -Title "💻 OPERATING SYSTEM DETAILS" -Data $data
    } catch {
        return Format-SimpleTable -Title "💻 OPERATING SYSTEM DETAILS" -Data @(@{"Fehler" = $_.Exception.Message})
    }
}

function Get-HardwareSummary {
    try {
        $hw = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $data = @(
            @{
                "Eigenschaft" = "Hersteller"
                "Wert" = $hw.Manufacturer
            },
            @{
                "Eigenschaft" = "Modell"
                "Wert" = $hw.Model
            },
            @{
                "Eigenschaft" = "Prozessoren"
                "Wert" = $hw.NumberOfProcessors
            },
            @{
                "Eigenschaft" = "Logische Kerne"
                "Wert" = $hw.NumberOfLogicalProcessors
            },
            @{
                "Eigenschaft" = "RAM (GB)"
                "Wert" = "$('{0:N2}' -f ($hw.TotalPhysicalMemory / 1GB))"
            },
            @{
                "Eigenschaft" = "Systemtyp"
                "Wert" = $hw.SystemType
            }
        )
        return Format-SimpleTable -Title "🖥️ HARDWARE SUMMARY" -Data $data
    } catch {
        return Format-SimpleTable -Title "🖥️ HARDWARE SUMMARY" -Data @(@{"Fehler" = $_.Exception.Message})
    }
}

function Get-CPUDetails {
    try {
        $cpu = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1
        $data = @(
            @{
                "Eigenschaft" = "Name"
                "Wert" = $cpu.Name
            },
            @{
                "Eigenschaft" = "Kerne"
                "Wert" = $cpu.NumberOfCores
            },
            @{
                "Eigenschaft" = "Logische Prozessoren"
                "Wert" = $cpu.NumberOfLogicalProcessors
            },
            @{
                "Eigenschaft" = "Geschwindigkeit (GHz)"
                "Wert" = "$('{0:N2}' -f ($cpu.MaxClockSpeed / 1000))"
            },
            @{
                "Eigenschaft" = "Architektur"
                "Wert" = $cpu.Architecture
            },
            @{
                "Eigenschaft" = "Cache (KB)"
                "Wert" = $cpu.L3CacheSize
            }
        )
        return Format-SimpleTable -Title "⚡ CPU INFORMATION" -Data $data
    } catch {
        return Format-SimpleTable -Title "⚡ CPU INFORMATION" -Data @(@{"Fehler" = $_.Exception.Message})
    }
}

function Get-MemoryDetails {
    try {
        $mem = Get-CimInstance Win32_PhysicalMemory -ErrorAction Stop
        $data = @()
        $data += @{"Eigenschaft" = "Anzahl Module"; "Wert" = $mem.Count}
        foreach ($m in $mem) {
            $data += @{
                "Eigenschaft" = $m.PartNumber
                "Wert" = "$('{0:N2}' -f ($m.Capacity / 1GB)) GB @ $($m.Speed) MHz"
            }
        }
        return Format-SimpleTable -Title "🧠 MEMORY DETAILS" -Data $data
    } catch {
        return Format-SimpleTable -Title "🧠 MEMORY DETAILS" -Data @(@{"Fehler" = $_.Exception.Message})
    }
}

function Get-StorageSummary {
    try {
        $disks = Get-CimInstance Win32_LogicalDisk -ErrorAction Stop | Where-Object DriveType -eq 3
        $data = foreach ($disk in $disks) {
            $used = '{0:N2}' -f (($disk.Size - $disk.FreeSpace) / 1GB)
            $total = '{0:N2}' -f ($disk.Size / 1GB)
            $percent = [math]::Round(($used / $total) * 100, 2)
            @{
                "Laufwerk" = $disk.Name
                "Gesamt (GB)" = $total
                "Belegt (GB)" = $used
                "Frei (%)" = "$(100 - $percent)%"
            }
        }
        return Format-TableOutput -Title "💾 STORAGE INFORMATION" -Data $data -Columns @("Laufwerk", "Gesamt (GB)", "Belegt (GB)", "Frei (%)")
    } catch {
        return Format-SimpleTable -Title "💾 STORAGE INFORMATION" -Data @(@{"Fehler" = $_.Exception.Message})
    }
}

function Get-NetworkConfiguration {
    try {
        $config = Get-NetIPConfiguration -ErrorAction Stop
        $data = foreach ($cfg in $config) {
            @{
                "Interface" = $cfg.InterfaceAlias
                "IPv4" = ($cfg.IPv4Address.IPAddress -join ', ')
                "Gateway" = ($cfg.IPv4DefaultGateway.NextHopAddress -join ', ')
            }
        }
        return Format-TableOutput -Title "🌐 NETWORK IP CONFIGURATION" -Data $data -Columns @("Interface", "IPv4", "Gateway")
    } catch {
        return Format-SimpleTable -Title "🌐 NETWORK IP CONFIGURATION" -Data @(@{"Fehler" = $_.Exception.Message})
    }
}

function Get-NetworkAdapters {
    try {
        $adapters = Get-NetAdapter -ErrorAction Stop
        $data = foreach ($adapter in $adapters) {
            $speedGbps = '{0:N2}' -f ($adapter.Speed / 1000000000)
            @{
                "Name" = $adapter.Name
                "Status" = $adapter.Status
                "Geschwindigkeit" = "$speedGbps Gbps"
                "MAC" = $adapter.MacAddress
            }
        }
        return Format-TableOutput -Title "🔌 NETWORK ADAPTERS" -Data $data -Columns @("Name", "Status", "Geschwindigkeit", "MAC")
    } catch {
        return Format-SimpleTable -Title "🔌 NETWORK ADAPTERS" -Data @(@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ActiveConnections {
    try {
        $connections = Get-NetTCPConnection -State Listen -ErrorAction Stop | Select-Object LocalAddress, LocalPort, OwningProcess -First 50
        $data = foreach ($conn in $connections) {
            $process = Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue
            @{
                "IP:Port" = "$($conn.LocalAddress):$($conn.LocalPort)"
                "PID" = $conn.OwningProcess
                "Prozess" = $process.ProcessName
            }
        }
        return Format-TableOutput -Title "🔗 ACTIVE NETWORK CONNECTIONS" -Data $data -Columns @("IP:Port", "PID", "Prozess")
    } catch {
        return Format-SimpleTable -Title "🔗 ACTIVE NETWORK CONNECTIONS" -Data @(@{"Fehler" = $_.Exception.Message})
    }
}

function Get-FirewallRules {
    try {
        $rules = Get-NetFirewallRule -Enabled $true -ErrorAction Stop | Select-Object DisplayName, Direction, Action -First 50
        $data = foreach ($rule in $rules) {
            @{
                "Name" = $rule.DisplayName.Substring(0, [Math]::Min(30, $rule.DisplayName.Length))
                "Richtung" = $rule.Direction
                "Aktion" = $rule.Action
            }
        }
        return Format-TableOutput -Title "🔥 FIREWALL RULES" -Data $data -Columns @("Name", "Richtung", "Aktion")
    } catch {
        return Format-SimpleTable -Title "🔥 FIREWALL RULES" -Data @(@{"Fehler" = $_.Exception.Message})
    }
}

function Get-AutomaticServices {
    $result = "=== AUTOMATIC SERVICES ===`n"
    try {
        $services = Get-Service -ErrorAction Stop | Where-Object StartType -eq 'Automatic' | Sort-Object Status, Name | Select-Object -First 50
        foreach ($svc in $services) {
            $result += "$($svc.Name) | Status: $($svc.Status) | Display: $($svc.DisplayName)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-RunningServices {
    $result = "=== RUNNING SERVICES ===`n"
    try {
        $services = Get-Service -ErrorAction Stop | Where-Object Status -eq 'Running' | Sort-Object Name | Select-Object -First 50
        foreach ($svc in $services) {
            $result += "$($svc.Name) | $($svc.DisplayName)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ScheduledTasks {
    $result = "=== SCHEDULED TASKS (READY) ===`n"
    try {
        $tasks = Get-ScheduledTask -ErrorAction Stop | Where-Object State -eq 'Ready' | Select-Object TaskName, TaskPath, State -First 50
        foreach ($task in $tasks) {
            $result += "$($task.TaskPath)$($task.TaskName)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-InstalledFeatures {
    $result = "=== INSTALLED WINDOWS FEATURES ===`n"
    try {
        $features = Get-WindowsFeature -ErrorAction Stop | Where-Object Installed -eq $true | Select-Object Name, DisplayName
        foreach ($feat in $features) {
            $result += "$($feat.Name) - $($feat.DisplayName)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-InstalledPrograms {
    $result = "=== INSTALLED PROGRAMS ===`n"
    try {
        $programs = Get-CimInstance Win32_Product -ErrorAction Stop | Select-Object Name, Version, Vendor | Sort-Object Name | Select-Object -First 100
        foreach ($prog in $programs) {
            $result += "$($prog.Name) | Ver: $($prog.Version) | Vendor: $($prog.Vendor)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-WindowsUpdates {
    $result = "=== RECENT WINDOWS UPDATES ===`n"
    try {
        $updates = Get-HotFix -ErrorAction Stop | Sort-Object InstalledOn -Descending | Select-Object -First 20
        foreach ($upd in $updates) {
            $result += "KB$($upd.HotFixID) - $($upd.InstalledOn)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-SystemEvents {
    $result = "=== SYSTEM EVENT LOG (LAST 24 HOURS - TOP 50) ===`n"
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='System'; StartTime=(Get-Date).AddDays(-1)} -MaxEvents 50 -ErrorAction Stop | Select-Object TimeCreated, Id, LevelDisplayName, Message
        foreach ($evt in $events) {
            $result += "$($evt.TimeCreated) [$($evt.LevelDisplayName)] ID:$($evt.Id) - $($evt.Message.Substring(0, [Math]::Min(100, $evt.Message.Length)))`n"
        }
    } catch {
        $result += "Error retrieving System events: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ApplicationEvents {
    $result = "=== APPLICATION EVENT LOG (LAST 24 HOURS - TOP 50) ===`n"
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddDays(-1)} -MaxEvents 50 -ErrorAction Stop | Select-Object TimeCreated, Id, LevelDisplayName, Message
        foreach ($evt in $events) {
            $result += "$($evt.TimeCreated) [$($evt.LevelDisplayName)] ID:$($evt.Id) - $($evt.Message.Substring(0, [Math]::Min(100, $evt.Message.Length)))`n"
        }
    } catch {
        $result += "Error retrieving Application events: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-SecurityEvents {
    $result = "=== SECURITY EVENT LOG (TOP 100) ===`n"
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security'} -MaxEvents 100 -ErrorAction Stop | Select-Object TimeCreated, Id, LevelDisplayName
        foreach ($evt in $events) {
            $result += "$($evt.TimeCreated) [ID:$($evt.Id)] - $($evt.LevelDisplayName)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-FailedLogons {
    $result = "=== FAILED LOGON ATTEMPTS (EVENT ID 4625) ===`n"
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4625} -MaxEvents 50 -ErrorAction Stop
        foreach ($evt in $events) {
            $result += "$($evt.TimeCreated) - $($evt.Message.Substring(0, [Math]::Min(150, $evt.Message.Length)))`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-AccountLockouts {
    $result = "=== ACCOUNT LOCKOUTS (EVENT ID 4740) ===`n"
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4740} -MaxEvents 50 -ErrorAction Stop
        foreach ($evt in $events) {
            $result += "$($evt.TimeCreated) - $($evt.Message.Substring(0, [Math]::Min(150, $evt.Message.Length)))`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-LocalUsers {
    $result = "=== LOCAL USERS ===`n"
    try {
        $users = Get-LocalUser -ErrorAction Stop | Select-Object Name, Enabled, LastLogon, PasswordRequired
        foreach ($user in $users) {
            $result += "$($user.Name) | Enabled: $($user.Enabled) | LastLogon: $($user.LastLogon) | PwdRequired: $($user.PasswordRequired)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-LocalGroups {
    $result = "=== LOCAL GROUPS ===`n"
    try {
        $groups = Get-LocalGroup -ErrorAction Stop | Select-Object Name, Description
        foreach ($group in $groups) {
            $result += "$($group.Name) - $($group.Description)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-PrivilegeAudit {
    $result = "=== PRIVILEGE USE AUDIT (EVENT IDS 4672/4673/4674) ===`n"
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4672,4673,4674} -MaxEvents 50 -ErrorAction Stop
        foreach ($evt in $events) {
            $result += "$($evt.TimeCreated) [ID:$($evt.Id)] - $($evt.Message.Substring(0, [Math]::Min(150, $evt.Message.Length)))`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# AD Functions
function Get-ADDomainControllers {
    $result = "=== ACTIVE DIRECTORY - DOMAIN CONTROLLERS ===`n"
    try {
        $dcs = Get-ADDomainController -Filter * -ErrorAction Stop | Select-Object Name, Site, IPv4Address, OperatingSystem, IsGlobalCatalog, IsReadOnly
        foreach ($dc in $dcs) {
            $result += "$($dc.Name) | Site: $($dc.Site) | IP: $($dc.IPv4Address) | GC: $($dc.IsGlobalCatalog)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ADDomainInfo {
    $result = "=== ACTIVE DIRECTORY - DOMAIN INFORMATION ===`n"
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $result += "Name: $($domain.Name)`n"
        $result += "NetBIOS: $($domain.NetBIOSName)`n"
        $result += "Mode: $($domain.DomainMode)`n"
        $result += "PDC: $($domain.PDCEmulator)`n"
        $result += "RID Master: $($domain.RIDMaster)`n"
        $result += "Infrastructure Master: $($domain.InfrastructureMaster)`n"
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ADForestInfo {
    $result = "=== ACTIVE DIRECTORY - FOREST INFORMATION ===`n"
    try {
        $forest = Get-ADForest -ErrorAction Stop
        $result += "Name: $($forest.Name)`n"
        $result += "Mode: $($forest.ForestMode)`n"
        $result += "Domain Naming Master: $($forest.DomainNamingMaster)`n"
        $result += "Schema Master: $($forest.SchemaMaster)`n"
        $result += "Sites: $(($forest.Sites | Measure-Object).Count)`n"
        $result += "Domains: $(($forest.Domains | Measure-Object).Count)`n"
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ADOUs {
    $result = "=== ACTIVE DIRECTORY - ORGANIZATIONAL UNITS ===`n"
    try {
        $ous = Get-ADOrganizationalUnit -Filter * -ErrorAction Stop | Select-Object Name, DistinguishedName | Sort-Object Name
        foreach ($ou in $ous | Select-Object -First 50) {
            $result += "$($ou.Name)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ADDomainAdmins {
    $result = "=== ACTIVE DIRECTORY - DOMAIN ADMINS ===`n"
    try {
        $admins = Get-ADGroupMember -Identity 'Domain Admins' -ErrorAction Stop | Get-ADUser -Properties LastLogonDate, PasswordLastSet, Enabled | Select-Object Name, SamAccountName, Enabled, LastLogonDate, PasswordLastSet
        foreach ($admin in $admins) {
            $result += "$($admin.Name) | $($admin.SamAccountName) | Enabled: $($admin.Enabled) | LastLogon: $($admin.LastLogonDate)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ADComputers {
    $result = "=== ACTIVE DIRECTORY - COMPUTER ACCOUNTS ===`n"
    try {
        $computers = Get-ADComputer -Filter * -Properties OperatingSystem, LastLogonDate -ErrorAction Stop | Select-Object Name, OperatingSystem, LastLogonDate, Enabled | Sort-Object LastLogonDate -Descending | Select-Object -First 50
        foreach ($comp in $computers) {
            $result += "$($comp.Name) | OS: $($comp.OperatingSystem) | LastLogon: $($comp.LastLogonDate) | Enabled: $($comp.Enabled)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ADReplicationStatus {
    $result = "=== ACTIVE DIRECTORY - REPLICATION STATUS ===`n"
    try {
        $repl = &cmd /c "repadmin /replsummary" 2>&1
        $result += $repl | Out-String
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ADTrusts {
    $result = "=== ACTIVE DIRECTORY - TRUST RELATIONSHIPS ===`n"
    try {
        $trusts = Get-ADTrust -Filter * -ErrorAction Stop | Select-Object Name, Direction, TrustType, DisallowTransivity
        foreach ($trust in $trusts) {
            $result += "$($trust.Name) | $($trust.Direction) | $($trust.TrustType)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# DNS Functions
function Get-DNSConfiguration {
    $result = "=== DNS SERVER CONFIGURATION ===`n"
    try {
        $dns = Get-DnsServer -ErrorAction Stop | Select-Object ComputerName, ZoneScavenging, EnableDnsSec
        $result += "Computer: $($dns.ComputerName)`n"
        $result += "Zone Scavenging: $($dns.ZoneScavenging)`n"
        $result += "DNSSEC Enabled: $($dns.EnableDnsSec)`n"
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-DNSZones {
    $result = "=== DNS SERVER ZONES ===`n"
    try {
        $zones = Get-DnsServerZone -ErrorAction Stop | Select-Object ZoneName, ZoneType, IsAutoCreated, IsDsIntegrated
        foreach ($zone in $zones) {
            $result += "$($zone.ZoneName) | Type: $($zone.ZoneType) | DS: $($zone.IsDsIntegrated)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-DNSForwarders {
    $result = "=== DNS FORWARDERS ===`n"
    try {
        $fwd = Get-DnsServerForwarder -ErrorAction Stop
        foreach ($f in $fwd.IPAddress) {
            $result += "$f`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-DNSCache {
    $result = "=== DNS CACHE (TOP 20 ENTRIES) ===`n"
    try {
        $cache = Get-DnsServerCache -ErrorAction Stop | Select-Object -First 20
        foreach ($entry in $cache) {
            $result += "$entry`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# DHCP Functions
function Get-DHCPConfiguration {
    $result = "=== DHCP SERVER CONFIGURATION ===`n"
    try {
        $dhcp = Get-DhcpServerInDC -ErrorAction Stop
        $result += $dhcp | Out-String
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-DHCPv4Scopes {
    $result = "=== DHCP IPv4 SCOPES ===`n"
    try {
        $scopes = Get-DhcpServerv4Scope -ErrorAction Stop | Select-Object ScopeId, Name, StartRange, EndRange, State
        foreach ($scope in $scopes) {
            $result += "$($scope.ScopeId) - $($scope.Name) | $($scope.StartRange) - $($scope.EndRange) | $($scope.State)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-DHCPv6Scopes {
    $result = "=== DHCP IPv6 SCOPES ===`n"
    try {
        $scopes = Get-DhcpServerv6Scope -ErrorAction Stop | Select-Object Prefix, Name, State
        foreach ($scope in $scopes) {
            $result += "$($scope.Prefix) - $($scope.Name) | $($scope.State)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-DHCPReservations {
    $result = "=== DHCP RESERVATIONS ===`n"
    try {
        $scopes = Get-DhcpServerv4Scope -ErrorAction Stop
        if ($scopes) {
            $scopes | ForEach-Object {
                $reservations = Get-DhcpServerv4Reservation -ScopeId $_.ScopeId -ErrorAction SilentlyContinue
                foreach ($res in $reservations | Select-Object -First 20) {
                    $result += "$($res.ScopeId) | $($res.IPAddress) | $($res.Name)`n"
                }
            }
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# ============================================================================
# IIS FUNCTIONS
# ============================================================================

function Get-IISWebsites {
    $result = "=== IIS WEBSITES ===`n"
    try {
        $sites = Get-IISSite -ErrorAction Stop | Select-Object Name, Id, State, PhysicalPath
        foreach ($site in $sites) {
            $result += "$($site.Name) | State: $($site.State) | Path: $($site.PhysicalPath)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-IISAppPools {
    $result = "=== IIS APPLICATION POOLS ===`n"
    try {
        $pools = Get-IISAppPool -ErrorAction Stop | Select-Object Name, State, StartMode
        foreach ($pool in $pools) {
            $result += "$($pool.Name) | State: $($pool.State) | StartMode: $($pool.StartMode)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-IISBindings {
    $result = "=== IIS SSL/BINDINGS ===`n"
    try {
        $bindings = Get-ChildItem IIS:SslBindings -ErrorAction Stop | Select-Object IPAddress, Port, Host, Thumbprint
        foreach ($binding in $bindings) {
            $result += "IP: $($binding.IPAddress):$($binding.Port) | Host: $($binding.Host) | Thumbprint: $($binding.Thumbprint)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# ============================================================================
# RDS / WTS FUNCTIONS
# ============================================================================

function Get-RDSCollections {
    $result = "=== RDS SESSION COLLECTIONS ===`n"
    try {
        $collections = Get-RDSessionCollection -ErrorAction Stop | Select-Object CollectionName, CollectionDescription
        foreach ($collection in $collections) {
            $result += "Name: $($collection.CollectionName) | Desc: $($collection.CollectionDescription)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-RDSSessionHosts {
    $result = "=== RDS SESSION HOSTS ===`n"
    try {
        $hosts = Get-RDSessionHost -ErrorAction Stop | Select-Object SessionHost, NewConnectionAllowed
        foreach ($rdHost in $hosts) {
            $result += "$($rdHost.SessionHost) | New Connections: $($rdHost.NewConnectionAllowed)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-RDSActiveLicensing {
    $result = "=== RDS LICENSING ===`n"
    try {
        $licensing = Get-RDLicenseConfiguration -ErrorAction Stop | Select-Object Mode, LicenseServer
        $result += "Mode: $($licensing.Mode) | Server: $($licensing.LicenseServer)`n"
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# ============================================================================
# DFS FUNCTIONS
# ============================================================================

function Get-DFSNamespaces {
    $result = "=== DFS NAMESPACES ===`n"
    try {
        $namespaces = Get-DfsnRoot -ErrorAction Stop | Select-Object Path, Type, State
        foreach ($ns in $namespaces) {
            $result += "$($ns.Path) | Type: $($ns.Type) | State: $($ns.State)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-DFSReplicationGroups {
    $result = "=== DFS REPLICATION GROUPS ===`n"
    try {
        $groups = Get-DfsReplicationGroup -ErrorAction Stop | Select-Object GroupName, State, Description
        foreach ($group in $groups) {
            $result += "$($group.GroupName) | State: $($group.State)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# ============================================================================
# PRINT SERVER FUNCTIONS
# ============================================================================

function Get-PrintServers {
    $result = "=== PRINT SERVERS ===`n"
    try {
        $printers = Get-Printer -ErrorAction Stop | Select-Object Name, DriverName, Shared, Published
        foreach ($printer in $printers) {
            $result += "$($printer.Name) | Driver: $($printer.DriverName) | Shared: $($printer.Shared)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-PrinterDrivers {
    $result = "=== PRINTER DRIVERS ===`n"
    try {
        $drivers = Get-PrinterDriver -ErrorAction Stop | Select-Object Name, Manufacturer, DriverVersion
        foreach ($driver in $drivers) {
            $result += "$($driver.Name) | Vendor: $($driver.Manufacturer) | Ver: $($driver.DriverVersion)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# ============================================================================
# WSUS FUNCTIONS
# ============================================================================

function Get-WSUSConfiguration {
    $result = "=== WSUS SERVER CONFIGURATION ===`n"
    try {
        $wsus = Get-WsusServer -ErrorAction Stop | Select-Object Name, PortNumber, ServerProtocolVersion
        $result += "Name: $($wsus.Name) | Port: $($wsus.PortNumber) | Protocol: $($wsus.ServerProtocolVersion)`n"
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-WSUSComputerTargetGroups {
    $result = "=== WSUS COMPUTER TARGET GROUPS ===`n"
    try {
        $server = Get-WsusServer -ErrorAction Stop
        $groups = $server | Get-WsusComputerTargetGroup -ErrorAction Stop | Select-Object Name
        foreach ($group in $groups) {
            $result += "$($group.Name)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-WSUSUpdates {
    $result = "=== WSUS AVAILABLE UPDATES ===`n"
    try {
        $server = Get-WsusServer -ErrorAction Stop
        $updates = $server.GetUpdates() | Select-Object Title, Classification, ApprovedCount | Select-Object -First 50
        foreach ($update in $updates) {
            $result += "$($update.Title) | Class: $($update.Classification) | Approved: $($update.ApprovedCount)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# ============================================================================
# HYPER-V FUNCTIONS
# ============================================================================

function Get-HyperVVirtualMachines {
    $result = "=== HYPER-V VIRTUAL MACHINES ===`n"
    try {
        $vms = Get-VM -ErrorAction Stop | Select-Object Name, State, MemoryAssigned, ProcessorCount
        foreach ($vm in $vms | Select-Object -First 50) {
            $result += "$($vm.Name) | State: $($vm.State) | RAM: $('{0:N0}' -f ($vm.MemoryAssigned / 1GB)) GB | CPUs: $($vm.ProcessorCount)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-HyperVSwitches {
    $result = "=== HYPER-V VIRTUAL SWITCHES ===`n"
    try {
        $switches = Get-VMSwitch -ErrorAction Stop | Select-Object Name, SwitchType, NetAdapterInterfaceDescription
        foreach ($switch in $switches) {
            $result += "$($switch.Name) | Type: $($switch.SwitchType) | Adapter: $($switch.NetAdapterInterfaceDescription)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-HyperVSnapshots {
    $result = "=== HYPER-V SNAPSHOTS ===`n"
    try {
        $snapshots = Get-VMSnapshot -ErrorAction Stop | Select-Object VMName, Name, CreationTime
        foreach ($snap in $snapshots | Select-Object -First 50) {
            $result += "VM: $($snap.VMName) | Snapshot: $($snap.Name) | Created: $($snap.CreationTime)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# ============================================================================
# NRAS / NPS FUNCTIONS
# ============================================================================

function Get-NPASConfiguration {
    $result = "=== NRAS / NPS CONFIGURATION ===`n"
    try {
        $nps = netsh nps show config -ErrorAction Stop 2>&1
        $result += ($nps | Out-String)
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-NASClients {
    $result = "=== NPS RADIUS NAS CLIENTS ===`n"
    try {
        $nasclients = Get-NpsRadiusClient -ErrorAction Stop | Select-Object Name, Address, SharedSecret
        foreach ($nas in $nasclients) {
            $result += "$($nas.Name) | IP: $($nas.Address)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# ============================================================================
# KMS FUNCTIONS (Volume Activation)
# ============================================================================

function Get-KMSConfiguration {
    $result = "=== KMS (VOLUME ACTIVATION) CONFIGURATION ===`n"
    try {
        $kmshost = cscript.exe $env:SystemRoot\system32\slmgr.vbs /dlv all -ErrorAction Stop 2>&1
        $result += ($kmshost | Out-String)
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# ============================================================================
# WDS FUNCTIONS
# ============================================================================

function Get-WDSConfiguration {
    $result = "=== WDS (WINDOWS DEPLOYMENT SERVICES) CONFIGURATION ===`n"
    try {
        $wdsconfig = wdsutil /get-server /show:config 2>&1
        $result += ($wdsconfig | Out-String)
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-WDSBootImages {
    $result = "=== WDS BOOT IMAGES ===`n"
    try {
        $bootimages = wdsutil /get-allimages /show:boot 2>&1
        $result += ($bootimages | Out-String | Select-Object -First 50 | Out-String)
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-WDSInstallImages {
    $result = "=== WDS INSTALL IMAGES ===`n"
    try {
        $installimages = wdsutil /get-allimages /show:install 2>&1
        $result += ($installimages | Out-String | Select-Object -First 50 | Out-String)
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# ============================================================================
# FILE SERVICES & SMB FUNCTIONS
# ============================================================================

function Get-FileShares {
    $result = "=== FILE SHARES (SMB) ===`n"
    try {
        $shares = Get-SmbShare -ErrorAction Stop | Select-Object Name, Path, Description, ShareType
        foreach ($share in $shares) {
            $result += "$($share.Name) | Path: $($share.Path) | Type: $($share.ShareType)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-FileSharePermissions {
    $result = "=== FILE SHARE PERMISSIONS ===`n"
    try {
        $shares = Get-SmbShare -ErrorAction Stop | Select-Object Name
        foreach ($share in $shares) {
            $perms = Get-SmbShareAccess -Name $share.Name -ErrorAction SilentlyContinue
            foreach ($perm in $perms) {
                $result += "$($share.Name) | $($perm.AccountName) | $($perm.AccessRight)`n"
            }
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-FileServerQuotas {
    $result = "=== FILE SERVER QUOTAS ===`n"
    try {
        $quotas = Get-FsrmQuota -ErrorAction Stop | Select-Object Path, Size, SoftLimit
        foreach ($quota in $quotas | Select-Object -First 50) {
            $result += "Path: $($quota.Path) | Size: $('{0:N0}' -f ($quota.Size / 1GB)) GB | SoftLimit: $($quota.SoftLimit)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ShadowCopies {
    $result = "=== SHADOW COPIES (VOLUME SNAPSHOTS) ===`n"
    try {
        $shadows = Get-CimInstance -ClassName Win32_ShadowCopy -ErrorAction Stop | Select-Object VolumeName, InstallDate, ID
        foreach ($shadow in $shadows) {
            $result += "Volume: $($shadow.VolumeName) | Date: $($shadow.InstallDate)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

# ============================================================================
# EXTENDED ACTIVE DIRECTORY FUNCTIONS
# ============================================================================

function Get-ADDomainControllerExtended {
    $result = "=== ACTIVE DIRECTORY - DOMAIN CONTROLLER STATUS (EXTENDED) ===`n"
    try {
        $dcs = Get-ADDomainController -Filter * -ErrorAction Stop | Select-Object Name, Site, IPv4Address, OperatingSystem, IsGlobalCatalog, IsReadOnly
        foreach ($dc in $dcs) {
            $result += "Name: $($dc.Name) | OS: $($dc.OperatingSystem) | GC: $($dc.IsGlobalCatalog) | RODC: $($dc.IsReadOnly)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ADDomainFunctionalLevel {
    $result = "=== ACTIVE DIRECTORY - DOMAIN & FOREST FUNCTIONAL LEVELS ===`n"
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $forest = Get-ADForest -ErrorAction Stop
        $result += "Domain Name: $($domain.Name)`n"
        $result += "Domain Mode: $($domain.DomainMode)`n"
        $result += "Forest Name: $($forest.Name)`n"
        $result += "Forest Mode: $($forest.ForestMode)`n"
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ADSiteConfiguration {
    $result = "=== ACTIVE DIRECTORY - SITES & SUBNETS ===`n"
    try {
        $sites = Get-ADReplicationSite -ErrorAction Stop | Select-Object Name, Description
        $result += "=== SITES ===`n"
        foreach ($site in $sites) {
            $result += "Site: $($site.Name)`n"
        }
        $result += "`n=== SUBNETS ===`n"
        $subnets = Get-ADReplicationSubnet -Filter * -ErrorAction Stop | Select-Object Name, Site
        foreach ($subnet in $subnets | Select-Object -First 50) {
            $result += "$($subnet.Name) | Site: $($subnet.Site)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ADGroupPolicySummary {
    $result = "=== ACTIVE DIRECTORY - GROUP POLICY SUMMARY ===`n"
    try {
        $gpos = Get-GPO -All -ErrorAction Stop | Measure-Object
        $result += "Total GPOs: $($gpos.Count)`n"
        $result += "Enabled GPOs: (use Get-GPO | Where GpoStatus -eq All | Measure).Count`n"
        $result += "Top 20 GPOs:`n"
        $topGPOs = Get-GPO -All -ErrorAction Stop | Select-Object DisplayName, CreationTime | Sort-Object CreationTime -Descending | Select-Object -First 20
        foreach ($gpo in $topGPOs) {
            $result += "  - $($gpo.DisplayName) | Created: $($gpo.CreationTime)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message)`n"
    }
    return $result
}

function Get-ADClusterInformation {
    $result = "=== ACTIVE DIRECTORY - FAILOVER CLUSTERING ===`n"
    try {
        $cluster = Get-Cluster -ErrorAction Stop
        $result += "Cluster Name: $($cluster.Name) | Domain: $($cluster.Domain)`n"
        $nodes = Get-ClusterNode -ErrorAction Stop | Select-Object Name, State
        $result += "`n=== CLUSTER NODES ===`n"
        foreach ($node in $nodes) {
            $result += "$($node.Name) | State: $($node.State)`n"
        }
    } catch {
        $result += "Error: $($_.Exception.Message) - Failover Clustering nicht konfiguriert`n"
    }
    return $result
}

# ============================================================================
# EVENT HANDLERS
# ============================================================================

# System Information Buttons
$window.FindName("ButtonSystemInfo").Add_Click({
    Update-Status "Running: System Overview..."
    Update-Output (Get-SystemInformation)
    Update-Status "Complete: System Overview"
})

$window.FindName("ButtonOSInfo").Add_Click({
    Update-Status "Running: OS Details..."
    Update-Output (Get-OSDetails)
    Update-Status "Complete: OS Details"
})

$window.FindName("ButtonHardwareInfo").Add_Click({
    Update-Status "Running: Hardware Summary..."
    Update-Output (Get-HardwareSummary)
    Update-Status "Complete: Hardware Summary"
})

$window.FindName("ButtonCPUInfo").Add_Click({
    Update-Status "Running: CPU Details..."
    Update-Output (Get-CPUDetails)
    Update-Status "Complete: CPU Details"
})

$window.FindName("ButtonMemoryInfo").Add_Click({
    Update-Status "Running: Memory Details..."
    Update-Output (Get-MemoryDetails)
    Update-Status "Complete: Memory Details"
})

$window.FindName("ButtonStorageInfo").Add_Click({
    Update-Status "Running: Storage Summary..."
    Update-Output (Get-StorageSummary)
    Update-Status "Complete: Storage Summary"
})

# Network Buttons
$window.FindName("ButtonNetConfig").Add_Click({
    Update-Status "Running: Network IP Configuration..."
    Update-Output (Get-NetworkConfiguration)
    Update-Status "Complete: Network IP Configuration"
})

$window.FindName("ButtonNetAdapters").Add_Click({
    Update-Status "Running: Network Adapters..."
    Update-Output (Get-NetworkAdapters)
    Update-Status "Complete: Network Adapters"
})

$window.FindName("ButtonTCPConnections").Add_Click({
    Update-Status "Running: Active Connections..."
    Update-Output (Get-ActiveConnections)
    Update-Status "Complete: Active Connections"
})

$window.FindName("ButtonFirewallRules").Add_Click({
    Update-Status "Running: Firewall Rules..."
    Update-Output (Get-FirewallRules)
    Update-Status "Complete: Firewall Rules"
})

# Services & Tasks Buttons
$window.FindName("ButtonAutomaticServices").Add_Click({
    Update-Status "Running: Automatic Services..."
    Update-Output (Get-AutomaticServices)
    Update-Status "Complete: Automatic Services"
})

$window.FindName("ButtonRunningServices").Add_Click({
    Update-Status "Running: Running Services..."
    Update-Output (Get-RunningServices)
    Update-Status "Complete: Running Services"
})

$window.FindName("ButtonScheduledTasks").Add_Click({
    Update-Status "Running: Scheduled Tasks..."
    Update-Output (Get-ScheduledTasks)
    Update-Status "Complete: Scheduled Tasks"
})

# Roles & Features Buttons
$window.FindName("ButtonInstalledFeatures").Add_Click({
    Update-Status "Running: Installed Features..."
    Update-Output (Get-InstalledFeatures)
    Update-Status "Complete: Installed Features"
})

$window.FindName("ButtonInstalledPrograms").Add_Click({
    Update-Status "Running: Installed Programs..."
    Update-Output (Get-InstalledPrograms)
    Update-Status "Complete: Installed Programs"
})

$window.FindName("ButtonWindowsUpdates").Add_Click({
    Update-Status "Running: Recent Updates..."
    Update-Output (Get-WindowsUpdates)
    Update-Status "Complete: Recent Updates"
})

# Event Logs Buttons
$window.FindName("ButtonSystemEvents").Add_Click({
    Update-Status "Running: System Events..."
    Update-Output (Get-SystemEvents)
    Update-Status "Complete: System Events"
})

$window.FindName("ButtonAppEvents").Add_Click({
    Update-Status "Running: Application Events..."
    Update-Output (Get-ApplicationEvents)
    Update-Status "Complete: Application Events"
})

$window.FindName("ButtonSecurityEvents").Add_Click({
    Update-Status "Running: Security Events..."
    Update-Output (Get-SecurityEvents)
    Update-Status "Complete: Security Events"
})

$window.FindName("ButtonFailedLogons").Add_Click({
    Update-Status "Running: Failed Logon Attempts..."
    Update-Output (Get-FailedLogons)
    Update-Status "Complete: Failed Logon Attempts"
})

$window.FindName("ButtonAccountLockouts").Add_Click({
    Update-Status "Running: Account Lockouts..."
    Update-Output (Get-AccountLockouts)
    Update-Status "Complete: Account Lockouts"
})

# Security & Users Buttons
$window.FindName("ButtonLocalUsers").Add_Click({
    Update-Status "Running: Local Users..."
    Update-Output (Get-LocalUsers)
    Update-Status "Complete: Local Users"
})

$window.FindName("ButtonLocalGroups").Add_Click({
    Update-Status "Running: Local Groups..."
    Update-Output (Get-LocalGroups)
    Update-Status "Complete: Local Groups"
})

$window.FindName("ButtonPrivilegeAudit").Add_Click({
    Update-Status "Running: Privilege Use Audit..."
    Update-Output (Get-PrivilegeAudit)
    Update-Status "Complete: Privilege Use Audit"
})

# Active Directory Buttons
$window.FindName("ButtonADDC").Add_Click({
    Update-Status "Running: AD Domain Controllers..."
    Update-Output (Get-ADDomainControllers)
    Update-Status "Complete: AD Domain Controllers"
})

$window.FindName("ButtonADDomain").Add_Click({
    Update-Status "Running: AD Domain Info..."
    Update-Output (Get-ADDomainInfo)
    Update-Status "Complete: AD Domain Info"
})

$window.FindName("ButtonADForest").Add_Click({
    Update-Status "Running: AD Forest Info..."
    Update-Output (Get-ADForestInfo)
    Update-Status "Complete: AD Forest Info"
})

$window.FindName("ButtonADOUs").Add_Click({
    Update-Status "Running: AD Organizational Units..."
    Update-Output (Get-ADOUs)
    Update-Status "Complete: AD Organizational Units"
})

$window.FindName("ButtonADAdmins").Add_Click({
    Update-Status "Running: AD Domain Admins..."
    Update-Output (Get-ADDomainAdmins)
    Update-Status "Complete: AD Domain Admins"
})

$window.FindName("ButtonADComputers").Add_Click({
    Update-Status "Running: AD Computers..."
    Update-Output (Get-ADComputers)
    Update-Status "Complete: AD Computers"
})

$window.FindName("ButtonADReplStatus").Add_Click({
    Update-Status "Running: AD Replication Status..."
    Update-Output (Get-ADReplicationStatus)
    Update-Status "Complete: AD Replication Status"
})

$window.FindName("ButtonADTrusts").Add_Click({
    Update-Status "Running: AD Trust Relationships..."
    Update-Output (Get-ADTrusts)
    Update-Status "Complete: AD Trust Relationships"
})

# DNS Buttons
$window.FindName("ButtonDNSConfig").Add_Click({
    Update-Status "Running: DNS Configuration..."
    Update-Output (Get-DNSConfiguration)
    Update-Status "Complete: DNS Configuration"
})

$window.FindName("ButtonDNSZones").Add_Click({
    Update-Status "Running: DNS Zones..."
    Update-Output (Get-DNSZones)
    Update-Status "Complete: DNS Zones"
})

$window.FindName("ButtonDNSForwarders").Add_Click({
    Update-Status "Running: DNS Forwarders..."
    Update-Output (Get-DNSForwarders)
    Update-Status "Complete: DNS Forwarders"
})

$window.FindName("ButtonDNSCache").Add_Click({
    Update-Status "Running: DNS Cache..."
    Update-Output (Get-DNSCache)
    Update-Status "Complete: DNS Cache"
})

# DHCP Buttons
$window.FindName("ButtonDHCPConfig").Add_Click({
    Update-Status "Running: DHCP Configuration..."
    Update-Output (Get-DHCPConfiguration)
    Update-Status "Complete: DHCP Configuration"
})

$window.FindName("ButtonDHCPv4Scopes").Add_Click({
    Update-Status "Running: DHCP IPv4 Scopes..."
    Update-Output (Get-DHCPv4Scopes)
    Update-Status "Complete: DHCP IPv4 Scopes"
})

$window.FindName("ButtonDHCPv6Scopes").Add_Click({
    Update-Status "Running: DHCP IPv6 Scopes..."
    Update-Output (Get-DHCPv6Scopes)
    Update-Status "Complete: DHCP IPv6 Scopes"
})

$window.FindName("ButtonDHCPReservations").Add_Click({
    Update-Status "Running: DHCP Reservations..."
    Update-Output (Get-DHCPReservations)
    Update-Status "Complete: DHCP Reservations"
})

# Export & Clear Buttons
$window.FindName("ButtonExportCSV").Add_Click({
    try {
        $csv_path = "$([System.IO.Path]::GetTempPath())easyWSAudit_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $OutputTextBox.Text | Out-File -FilePath $csv_path -Encoding UTF8
        Update-Status "Export complete: $csv_path"
        [System.Diagnostics.Process]::Start("notepad", $csv_path)
    } catch {
        Update-Status "Export error: $($_.Exception.Message)"
    }
})

$window.FindName("ButtonClearOutput").Add_Click({
    Clear-Output
    Update-Status "Output cleared"
})

# Export & Clear Buttons
$window.FindName("ButtonIISWebsites").Add_Click({Update-Status "Running: IIS Websites..."; Update-Output (Get-IISWebsites); Update-Status "Complete"})
$window.FindName("ButtonIISAppPools").Add_Click({Update-Status "Running: AppPools..."; Update-Output (Get-IISAppPools); Update-Status "Complete"})
$window.FindName("ButtonIISBindings").Add_Click({Update-Status "Running: SSL..."; Update-Output (Get-IISBindings); Update-Status "Complete"})
$window.FindName("ButtonRDSCollections").Add_Click({Update-Status "Running: RDS..."; Update-Output (Get-RDSCollections); Update-Status "Complete"})
$window.FindName("ButtonRDSSessionHosts").Add_Click({Update-Status "Running: Hosts..."; Update-Output (Get-RDSSessionHosts); Update-Status "Complete"})
$window.FindName("ButtonRDSLicensing").Add_Click({Update-Status "Running: Licensing..."; Update-Output (Get-RDSActiveLicensing); Update-Status "Complete"})
$window.FindName("ButtonDFSNamespaces").Add_Click({Update-Status "Running: DFS NS..."; Update-Output (Get-DFSNamespaces); Update-Status "Complete"})
$window.FindName("ButtonDFSReplication").Add_Click({Update-Status "Running: Replication..."; Update-Output (Get-DFSReplicationGroups); Update-Status "Complete"})
$window.FindName("ButtonPrinters").Add_Click({Update-Status "Running: Printers..."; Update-Output (Get-PrintServers); Update-Status "Complete"})
$window.FindName("ButtonPrinterDrivers").Add_Click({Update-Status "Running: Drivers..."; Update-Output (Get-PrinterDrivers); Update-Status "Complete"})
$window.FindName("ButtonWSUSConfig").Add_Click({Update-Status "Running: WSUS Config..."; Update-Output (Get-WSUSConfiguration); Update-Status "Complete"})
$window.FindName("ButtonWSUSGroups").Add_Click({Update-Status "Running: Groups..."; Update-Output (Get-WSUSComputerTargetGroups); Update-Status "Complete"})
$window.FindName("ButtonWSUSUpdates").Add_Click({Update-Status "Running: Updates..."; Update-Output (Get-WSUSUpdates); Update-Status "Complete"})
$window.FindName("ButtonHyperVVMs").Add_Click({Update-Status "Running: VMs..."; Update-Output (Get-HyperVVirtualMachines); Update-Status "Complete"})
$window.FindName("ButtonHyperVSwitches").Add_Click({Update-Status "Running: Switches..."; Update-Output (Get-HyperVSwitches); Update-Status "Complete"})
$window.FindName("ButtonHyperVSnapshots").Add_Click({Update-Status "Running: Snapshots..."; Update-Output (Get-HyperVSnapshots); Update-Status "Complete"})
$window.FindName("ButtonNPASConfig").Add_Click({Update-Status "Running: NRAS/NPS..."; Update-Output (Get-NPASConfiguration); Update-Status "Complete"})
$window.FindName("ButtonNASClients").Add_Click({Update-Status "Running: NAS..."; Update-Output (Get-NASClients); Update-Status "Complete"})
$window.FindName("ButtonKMSConfig").Add_Click({Update-Status "Running: KMS..."; Update-Output (Get-KMSConfiguration); Update-Status "Complete"})
$window.FindName("ButtonWDSConfig").Add_Click({Update-Status "Running: WDS..."; Update-Output (Get-WDSConfiguration); Update-Status "Complete"})
$window.FindName("ButtonWDSBootImages").Add_Click({Update-Status "Running: Boot..."; Update-Output (Get-WDSBootImages); Update-Status "Complete"})
$window.FindName("ButtonWDSInstallImages").Add_Click({Update-Status "Running: Install..."; Update-Output (Get-WDSInstallImages); Update-Status "Complete"})
$window.FindName("ButtonFileShares").Add_Click({Update-Status "Running: Shares..."; Update-Output (Get-FileShares); Update-Status "Complete"})
$window.FindName("ButtonSharePermissions").Add_Click({Update-Status "Running: Perms..."; Update-Output (Get-FileSharePermissions); Update-Status "Complete"})
$window.FindName("ButtonFileQuotas").Add_Click({Update-Status "Running: Quotas..."; Update-Output (Get-FileServerQuotas); Update-Status "Complete"})
$window.FindName("ButtonShadowCopies").Add_Click({Update-Status "Running: Shadows..."; Update-Output (Get-ShadowCopies); Update-Status "Complete"})
$window.FindName("ButtonADDCExtended").Add_Click({Update-Status "Running: DC Ext..."; Update-Output (Get-ADDomainControllerExtended); Update-Status "Complete"})
$window.FindName("ButtonADFunctionalLevels").Add_Click({Update-Status "Running: Levels..."; Update-Output (Get-ADDomainFunctionalLevel); Update-Status "Complete"})
$window.FindName("ButtonADSites").Add_Click({Update-Status "Running: Sites..."; Update-Output (Get-ADSiteConfiguration); Update-Status "Complete"})
$window.FindName("ButtonADGPO").Add_Click({Update-Status "Running: GPO..."; Update-Output (Get-ADGroupPolicySummary); Update-Status "Complete"})
$window.FindName("ButtonADClustering").Add_Click({Update-Status "Running: Cluster..."; Update-Output (Get-ADClusterInformation); Update-Status "Complete"})

# ============================================================================
# MAIN
# ============================================================================

$window.ShowDialog()
