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

                    <DataGrid
                        x:Name="DataGridResults"
                        Grid.Row="1"
                        AlternatingRowBackground="#F8FAFC"
                        AutoGenerateColumns="True"
                        Background="White"
                        BorderBrush="#E5E7EB"
                        BorderThickness="1"
                        CanUserReorderColumns="True"
                        CanUserResizeColumns="True"
                        CanUserSortColumns="True"
                        GridLinesVisibility="Horizontal"
                        HeadersVisibility="Column"
                        IsReadOnly="True"
                        RowBackground="White">
                        <DataGrid.ColumnHeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#F3F4F6" />
                                <Setter Property="Foreground" Value="#374151" />
                                <Setter Property="FontWeight" Value="SemiBold" />
                                <Setter Property="BorderBrush" Value="#E5E7EB" />
                                <Setter Property="BorderThickness" Value="0,0,1,1" />
                                <Setter Property="Padding" Value="12,8" />
                            </Style>
                        </DataGrid.ColumnHeaderStyle>
                        <DataGrid.CellStyle>
                            <Style TargetType="DataGridCell">
                                <Setter Property="Padding" Value="12,6" />
                                <Setter Property="BorderThickness" Value="0" />
                                <Style.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter Property="Background" Value="#EEF2FF" />
                                        <Setter Property="Foreground" Value="#6366F1" />
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.CellStyle>
                    </DataGrid>
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
$DataGridResults = $window.FindName("DataGridResults")
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
    param([PSObject[]]$Data)
    try {
        # Konvertiere alle Eingaben zu einem konsistenten Array-Format
        $processedData = @()
        
        if ($null -eq $Data -or $Data.Count -eq 0) {
            $script:ResultsCount = 0
        } else {
            # Handle verschiedene Input-Typen
            foreach ($item in $Data) {
                if ($null -ne $item) {
                    # Wenn einzelnes Objekt (nicht in einem Array)
                    if ($item -is [System.Collections.IEnumerable] -and $item -isnot [string] -and $item -isnot [System.Collections.Specialized.OrderedDictionary]) {
                        foreach ($subitem in $item) {
                            $processedData += $subitem
                        }
                    } else {
                        $processedData += $item
                    }
                }
            }
            
            # Erstelle ObservableCollection
            $collection = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
            foreach ($item in $processedData) {
                $collection.Add($item)
            }
            $script:ResultsCount = $collection.Count
            
            # UI-Update auf dem Dispatcher-Thread
            $window.Dispatcher.Invoke({
                try {
                    $DataGridResults.ItemsSource = $null
                    
                    # Auto-Generate Columns basierend auf den Daten
                    $DataGridResults.AutoGenerateColumns = $true
                    
                    $DataGridResults.ItemsSource = $collection
                    $TotalResultCountText.Text = $script:ResultsCount.ToString()
                    
                    # Force Refresh
                    $DataGridResults.Items.Refresh()
                } catch {
                    Write-Host "UI-Update Fehler: $_" -ForegroundColor Red
                }
            }, "Normal")
            
            return
        }
        
        # Fallback wenn Daten leer sind
        $window.Dispatcher.Invoke({
            $DataGridResults.ItemsSource = $null
            $TotalResultCountText.Text = "0"
        }, "Normal")
        
    } catch {
        Write-Host "Update-Output Fehler: $_" -ForegroundColor Red
        $window.Dispatcher.Invoke({
            try {
                $DataGridResults.ItemsSource = $null
                $TotalResultCountText.Text = "0"
            } catch {}
        }, "Normal")
    }
}

function Clear-Output {
    $DataGridResults.ItemsSource = @()
    $script:ResultsCount = 0
    $TotalResultCountText.Text = "0"
}

function Update-Status {
    param([string]$Status)
    $StatusBarText.Text = $Status
    $window.Dispatcher.Invoke([Action]{}, "Background")
}

# ============================================================================
# TABLE FORMATTING FUNCTIONS - REMOVED (Using DataGrid Now)
# ============================================================================

# Alle Ausgaben erfolgen jetzt direkt über DataGrid mit PSObjects
# Die alten Format-* Funktionen werden nicht mehr benötigt

# ============================================================================
# AUDIT FUNCTIONS
# ============================================================================

function Get-SystemInformation {
    try {
        $info = Get-ComputerInfo -ErrorAction Stop
        $data = @(
            @{"Attribut" = "Computername"; "Wert" = $info.CsComputerName},
            @{"Attribut" = "Domäne"; "Wert" = $info.CsDomain},
            @{"Attribut" = "Betriebssystem"; "Wert" = $info.OsName},
            @{"Attribut" = "Installationsdatum"; "Wert" = $info.OsInstallDate},
            @{"Attribut" = "Letzter Boot"; "Wert" = $info.OsLastBootUpTime}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-OSDetails {
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $data = @(
            @{"Eigenschaft" = "Caption"; "Wert" = $os.Caption},
            @{"Eigenschaft" = "Version"; "Wert" = $os.Version},
            @{"Eigenschaft" = "Build"; "Wert" = $os.BuildNumber},
            @{"Eigenschaft" = "Gesamtarbeitsspeicher"; "Wert" = "$('{0:N0}' -f ($os.TotalVisibleMemorySize / 1024)) MB"},
            @{"Eigenschaft" = "Freier Arbeitsspeicher"; "Wert" = "$('{0:N0}' -f ($os.FreePhysicalMemory / 1024)) MB"},
            @{"Eigenschaft" = "Letzter Boot"; "Wert" = $os.LastBootUpTime}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-HardwareSummary {
    try {
        $hw = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $data = @(
            @{"Eigenschaft" = "Hersteller"; "Wert" = $hw.Manufacturer},
            @{"Eigenschaft" = "Modell"; "Wert" = $hw.Model},
            @{"Eigenschaft" = "Prozessoren"; "Wert" = $hw.NumberOfProcessors},
            @{"Eigenschaft" = "Logische Kerne"; "Wert" = $hw.NumberOfLogicalProcessors},
            @{"Eigenschaft" = "RAM (GB)"; "Wert" = "$('{0:N2}' -f ($hw.TotalPhysicalMemory / 1GB))"},
            @{"Eigenschaft" = "Systemtyp"; "Wert" = $hw.SystemType}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-CPUDetails {
    try {
        $cpu = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1
        $data = @(
            @{"Eigenschaft" = "Name"; "Wert" = $cpu.Name},
            @{"Eigenschaft" = "Kerne"; "Wert" = $cpu.NumberOfCores},
            @{"Eigenschaft" = "Logische Prozessoren"; "Wert" = $cpu.NumberOfLogicalProcessors},
            @{"Eigenschaft" = "Geschwindigkeit (GHz)"; "Wert" = "$('{0:N2}' -f ($cpu.MaxClockSpeed / 1000))"},
            @{"Eigenschaft" = "Architektur"; "Wert" = $cpu.Architecture},
            @{"Eigenschaft" = "Cache (KB)"; "Wert" = $cpu.L3CacheSize}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
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
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-StorageSummary {
    try {
        $disks = Get-CimInstance Win32_LogicalDisk -ErrorAction Stop | Where-Object DriveType -eq 3
        $data = @()
        foreach ($disk in $disks) {
            if ($disk.Size -gt 0) {
                $used = '{0:N2}' -f (($disk.Size - $disk.FreeSpace) / 1GB)
                $total = '{0:N2}' -f ($disk.Size / 1GB)
                $percent = [math]::Round((($disk.Size - $disk.FreeSpace) / $disk.Size) * 100, 2)
                $freePercent = [math]::Round(100 - $percent, 2)
            } else {
                $used = "0"
                $total = "0"
                $percent = "0"
                $freePercent = "100"
            }
            $data += @{
                "Laufwerk" = $disk.Name
                "Gesamt(GB)" = $total
                "Belegt(GB)" = $used
                "Frei(%)" = $freePercent
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-NetworkConfiguration {
    try {
        $config = Get-NetIPConfiguration -ErrorAction Stop
        $data = @()
        foreach ($cfg in $config) {
            $data += @{
                "Interface" = $cfg.InterfaceAlias
                "IPv4" = ($cfg.IPv4Address.IPAddress -join ', ')
                "Gateway" = ($cfg.IPv4DefaultGateway.NextHopAddress -join ', ')
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-NetworkAdapters {
    try {
        $adapters = Get-NetAdapter -ErrorAction Stop
        $data = @()
        foreach ($adapter in $adapters) {
            $speedGbps = '{0:N2}' -f ($adapter.Speed / 1000000000)
            $data += @{
                "Name" = $adapter.Name
                "Status" = $adapter.Status
                "Speed(Gbps)" = $speedGbps
                "MAC" = $adapter.MacAddress
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ActiveConnections {
    try {
        $connections = Get-NetTCPConnection -State Listen -ErrorAction Stop | Select-Object LocalAddress, LocalPort, OwningProcess -First 50
        $data = @()
        foreach ($conn in $connections) {
            $process = Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue
            $data += @{
                "IP:Port" = "$($conn.LocalAddress):$($conn.LocalPort)"
                "PID" = $conn.OwningProcess
                "Process" = $process.ProcessName
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-FirewallRules {
    try {
        $rules = Get-NetFirewallRule -ErrorAction Stop | Where-Object { $_.Enabled -eq 'True' } | Select-Object DisplayName, Direction, Action -First 50
        if ($null -eq $rules) {
            return @([PSCustomObject]@{"Information" = "Keine aktivierten Firewall-Regeln gefunden"})
        }
        $data = @()
        foreach ($rule in $rules) {
            $name = $rule.DisplayName.Substring(0, [Math]::Min(40, $rule.DisplayName.Length))
            $data += @{
                "Name" = $name
                "Direction" = $rule.Direction
                "Action" = $rule.Action
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-AutomaticServices {
    try {
        $services = Get-Service -ErrorAction Stop | Where-Object StartType -eq 'Automatic' | Sort-Object Status, Name | Select-Object -First 50
        return $services | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Status" = $_.Status; "Display" = $_.DisplayName} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-RunningServices {
    try {
        $services = Get-Service -ErrorAction Stop | Where-Object Status -eq 'Running' | Sort-Object Name | Select-Object -First 50
        return $services | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Display" = $_.DisplayName} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ScheduledTasks {
    try {
        $tasks = Get-ScheduledTask -ErrorAction Stop | Where-Object State -eq 'Ready' | Select-Object -First 50
        return $tasks | ForEach-Object { [PSCustomObject]@{"Path" = $_.TaskPath; "Name" = $_.TaskName; "State" = $_.State} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-InstalledFeatures {
    try {
        $features = Get-WindowsFeature -ErrorAction Stop | Where-Object Installed -eq $true | Select-Object Name, DisplayName
        return $features | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Display" = $_.DisplayName} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-InstalledPrograms {
    try {
        $programs = Get-CimInstance Win32_Product -ErrorAction Stop | Select-Object Name, Version, Vendor | Sort-Object Name | Select-Object -First 100
        return $programs | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Version" = $_.Version; "Vendor" = $_.Vendor} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-WindowsUpdates {
    try {
        $updates = Get-HotFix -ErrorAction Stop | Sort-Object InstalledOn -Descending | Select-Object -First 20
        return $updates | ForEach-Object { [PSCustomObject]@{"KB" = $_.HotFixID; "InstalledOn" = $_.InstalledOn} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-SystemEvents {
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='System'; StartTime=(Get-Date).AddDays(-1)} -MaxEvents 50 -ErrorAction Stop
        if ($null -eq $events) {
            return @([PSCustomObject]@{"Information" = "Keine System Events gefunden"})
        }
        return $events | Select-Object -First 50 | ForEach-Object { 
            [PSCustomObject]@{
                "Zeit" = $_.TimeCreated
                "ID" = $_.Id
                "Level" = $_.LevelDisplayName
                "Message" = $_.Message.Substring(0, [Math]::Min(50, $_.Message.Length))
            } 
        }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ApplicationEvents {
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddDays(-1)} -MaxEvents 50 -ErrorAction Stop
        if ($null -eq $events) {
            return @([PSCustomObject]@{"Information" = "Keine Application Events gefunden"})
        }
        return $events | ForEach-Object { 
            [PSCustomObject]@{
                "Zeit" = $_.TimeCreated
                "ID" = $_.Id
                "Level" = $_.LevelDisplayName
            } 
        }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-SecurityEvents {
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security'} -MaxEvents 100 -ErrorAction Stop
        if ($null -eq $events) {
            return @([PSCustomObject]@{"Information" = "Keine Security Events gefunden"})
        }
        return $events | ForEach-Object { 
            [PSCustomObject]@{
                "Zeit" = $_.TimeCreated
                "ID" = $_.Id
                "Level" = $_.LevelDisplayName
            } 
        }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-FailedLogons {
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4625} -MaxEvents 50 -ErrorAction Stop
        if ($null -eq $events) {
            return @([PSCustomObject]@{"Information" = "Keine fehlgeschlagenen Anmeldungen gefunden"})
        }
        return $events | ForEach-Object { 
            [PSCustomObject]@{
                "Zeit" = $_.TimeCreated
                "Message" = $_.Message.Substring(0, [Math]::Min(80, $_.Message.Length))
            } 
        }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-AccountLockouts {
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4740} -MaxEvents 50 -ErrorAction Stop
        if ($null -eq $events) {
            return @([PSCustomObject]@{"Information" = "Keine Account-Sperren gefunden"})
        }
        return $events | ForEach-Object { 
            [PSCustomObject]@{
                "Zeit" = $_.TimeCreated
                "Message" = $_.Message.Substring(0, [Math]::Min(80, $_.Message.Length))
            } 
        }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-LocalUsers {
    try {
        $users = Get-LocalUser -ErrorAction Stop
        if ($null -eq $users) {
            return @([PSCustomObject]@{"Information" = "Keine lokalen Benutzer gefunden"})
        }
        $data = @()
        foreach ($user in $users) {
            $data += @{
                "Name" = $user.Name
                "Enabled" = if($user.Enabled) {"Ja"} else {"Nein"}
                "LastLogon" = $user.LastLogon
                "PwdRequired" = if($user.PasswordRequired) {"Ja"} else {"Nein"}
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-LocalGroups {
    try {
        $groups = Get-LocalGroup -ErrorAction Stop
        if ($null -eq $groups) {
            return @([PSCustomObject]@{"Information" = "Keine lokalen Gruppen gefunden"})
        }
        return $groups | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Description" = $_.Description} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-PrivilegeAudit {
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4672,4673,4674} -MaxEvents 50 -ErrorAction Stop
        if ($null -eq $events) {
            return @([PSCustomObject]@{"Information" = "Keine Privilege-Audit Events gefunden"})
        }
        return $events | ForEach-Object { 
            [PSCustomObject]@{
                "Zeit" = $_.TimeCreated
                "ID" = $_.Id
                "Message" = $_.Message.Substring(0, [Math]::Min(60, $_.Message.Length))
            } 
        }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# AD Functions (String alt noch umgestellt auf PSObjects)
function Get-ADDomainControllers {
    try {
        $dcs = Get-ADDomainController -Filter * -ErrorAction Stop | Select-Object Name, Site, IPv4Address, OperatingSystem, IsGlobalCatalog, IsReadOnly
        return $dcs | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Site" = $_.Site; "IP" = $_.IPv4Address; "GC" = $_.IsGlobalCatalog} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADDomainInfo {
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $data = @(
            @{"Property" = "Name"; "Value" = $domain.Name},
            @{"Property" = "NetBIOS"; "Value" = $domain.NetBIOSName},
            @{"Property" = "Mode"; "Value" = $domain.DomainMode},
            @{"Property" = "PDC"; "Value" = $domain.PDCEmulator}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DNSZones {
    try {
        $zones = Get-DnsServerZone -ErrorAction Stop | Select-Object -First 50
        return $zones | ForEach-Object { [PSCustomObject]@{"Zone" = $_.ZoneName; "Type" = $_.ZoneType; "DS" = $_.IsDsIntegrated} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-FileShares {
    try {
        $shares = Get-SmbShare -ErrorAction Stop | Select-Object Name, Path, Description, ShareType
        return $shares | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Path" = $_.Path; "Type" = $_.ShareType} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-IISWebsites {
    try {
        Import-Module WebAdministration -ErrorAction Stop
        $sites = Get-Website -ErrorAction Stop | Select-Object Name, State, PhysicalPath
        return $sites | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "State" = $_.State; "Path" = $_.PhysicalPath} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = "IIS: " + $_.Exception.Message})
    }
}

function Get-RDSCollections {
    try {
        $collections = Get-RDSessionCollection -ErrorAction Stop | Select-Object CollectionName, CollectionDescription
        return $collections | ForEach-Object { [PSCustomObject]@{"Name" = $_.CollectionName; "Description" = $_.CollectionDescription} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = "RDS: " + $_.Exception.Message})
    }
}

# ============================================================================
# EXTENDED AD FUNCTIONS
# ============================================================================

function Get-ADForestInfo {
    try {
        $forest = Get-ADForest -ErrorAction Stop
        $data = @(
            @{"Property" = "Name"; "Value" = $forest.Name},
            @{"Property" = "Mode"; "Value" = $forest.ForestMode},
            @{"Property" = "Domains"; "Value" = ($forest.Domains | Measure-Object).Count},
            @{"Property" = "Sites"; "Value" = ($forest.Sites | Measure-Object).Count}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADOUs {
    try {
        $ous = Get-ADOrganizationalUnit -Filter * -ErrorAction Stop | Select-Object Name, DistinguishedName | Sort-Object Name | Select-Object -First 50
        return $ous | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "DN" = $_.DistinguishedName} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADDomainAdmins {
    try {
        $admins = Get-ADGroupMember -Identity 'Domain Admins' -ErrorAction Stop | Get-ADUser -Properties LastLogonDate, PasswordLastSet, Enabled -ErrorAction SilentlyContinue | Select-Object Name, SamAccountName, Enabled, LastLogonDate
        return $admins | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Account" = $_.SamAccountName; "Enabled" = $_.Enabled; "LastLogon" = $_.LastLogonDate} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADComputers {
    try {
        $computers = Get-ADComputer -Filter * -Properties OperatingSystem, LastLogonDate -ErrorAction Stop | Select-Object Name, OperatingSystem, LastLogonDate, Enabled | Sort-Object LastLogonDate -Descending | Select-Object -First 50
        return $computers | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "OS" = $_.OperatingSystem; "LastLogon" = $_.LastLogonDate; "Enabled" = $_.Enabled} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADReplicationStatus {
    try {
        $output = repadmin /replsummary 2>&1
        $data = @()
        
        # Parse der repadmin Ausgabe
        foreach ($line in $output) {
            if (-not [string]::IsNullOrWhiteSpace($line) -and $line -match '\S') {
                $data += @{
                    "Information" = $line.Trim()
                }
            }
        }
        
        if ($data.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "Keine Replikationsinformationen verfügbar"})
        }
        
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADTrusts {
    try {
        $trusts = Get-ADTrust -Filter * -ErrorAction Stop | Select-Object Name, Direction, TrustType
        return $trusts | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Direction" = $_.Direction; "Type" = $_.TrustType} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# DNS FUNCTIONS
# ============================================================================

function Get-DNSConfiguration {
    try {
        # Versuche Get-DnsServerSetting zuerst (besser als Get-DnsServer)
        $dnsSettings = Get-DnsServerSetting -ErrorAction SilentlyContinue
        
        if ($null -ne $dnsSettings) {
            $data = @(
                @{"Setting" = "Listen Addresses"; "Value" = ($dnsSettings.ListeningIPAddress -join ", ")},
                @{"Setting" = "All Zones Writeable"; "Value" = if($dnsSettings.AllowZoneEditing) {"Ja"} else {"Nein"}},
                @{"Setting" = "DNSSEC Enabled"; "Value" = if($dnsSettings.EnableDnsSec) {"Ja"} else {"Nein"}},
                @{"Setting" = "Log Queries"; "Value" = if($dnsSettings.LogQueries) {"Ja"} else {"Nein"}},
                @{"Setting" = "Write to Log"; "Value" = if($dnsSettings.WriteToLog) {"Ja"} else {"Nein"}}
            )
            return $data | ForEach-Object { [PSCustomObject]$_ }
        } else {
            # Fallback: Get-DnsServer (ältere Windows-Versionen)
            $dns = Get-DnsServer -ErrorAction Stop
            $data = @(
                @{"Setting" = "Computer"; "Value" = $dns.ComputerName},
                @{"Setting" = "Version"; "Value" = $dns.Version}
            )
            return $data | ForEach-Object { [PSCustomObject]$_ }
        }
    } catch {
        return @([PSCustomObject]@{"Fehler" = "DNS nicht verfügbar oder kein Zugriff: $($_.Exception.Message)"})
    }
}

function Get-DNSForwarders {
    try {
        $fwd = Get-DnsServerForwarder -ErrorAction Stop
        if ($null -ne $fwd.IPAddress) {
            return $fwd.IPAddress | ForEach-Object { [PSCustomObject]@{"Forwarder" = $_} }
        } else {
            return @([PSCustomObject]@{"Forwarder" = "Keine Forwarder konfiguriert"})
        }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DNSCache {
    try {
        $cache = Get-DnsServerCache -ErrorAction Stop
        return @([PSCustomObject]@{"CacheSize" = $cache.CacheSize; "MaxTTL" = $cache.MaxTTL; "MaxNegativeTTL" = $cache.MaxNegativeTTL})
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# DHCP FUNCTIONS
# ============================================================================

function Get-DHCPConfiguration {
    try {
        $dhcp = Get-DhcpServerInDC -ErrorAction Stop | Select-Object -First 10
        return $dhcp | ForEach-Object { [PSCustomObject]@{"Server" = $_.ToString()} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DHCPv4Scopes {
    try {
        $scopes = Get-DhcpServerv4Scope -ErrorAction Stop | Select-Object ScopeId, Name, StartRange, EndRange, State | Select-Object -First 50
        return $scopes | ForEach-Object { [PSCustomObject]@{"ScopeID" = $_.ScopeId; "Name" = $_.Name; "Start" = $_.StartRange; "End" = $_.EndRange; "State" = $_.State} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DHCPv6Scopes {
    try {
        $scopes = Get-DhcpServerv6Scope -ErrorAction Stop | Select-Object Prefix, Name, State | Select-Object -First 50
        return $scopes | ForEach-Object { [PSCustomObject]@{"Prefix" = $_.Prefix; "Name" = $_.Name; "State" = $_.State} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DHCPReservations {
    try {
        $scopes = Get-DhcpServerv4Scope -ErrorAction Stop
        $reservations = @()
        foreach ($scope in $scopes) {
            $scopeRes = Get-DhcpServerv4Reservation -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue | Select-Object -First 20
            $reservations += $scopeRes
        }
        return $reservations | ForEach-Object { [PSCustomObject]@{"ScopeID" = $_.ScopeId; "IP" = $_.IPAddress; "Name" = $_.Name; "Type" = $_.ReservationType} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# IIS FUNCTIONS
# ============================================================================

function Get-IISAppPools {
    try {
        Import-Module WebAdministration -ErrorAction Stop
        $pools = Get-ChildItem IIS:\AppPools -ErrorAction Stop | Select-Object Name, State, ManagedRuntimeVersion
        return $pools | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "State" = $_.State; "Runtime" = $_.ManagedRuntimeVersion} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = "IIS: " + $_.Exception.Message})
    }
}

function Get-IISBindings {
    try {
        Import-Module WebAdministration -ErrorAction Stop
        $sites = Get-ChildItem IIS:\Sites -ErrorAction Stop
        $bindings = @()
        foreach ($site in $sites) {
            $site.Bindings.Collection | ForEach-Object { $bindings += $_ }
        }
        return $bindings | Select-Object -First 50 | ForEach-Object { [PSCustomObject]@{"Protocol" = $_.Protocol; "IP" = $_.BindingInformation} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = "IIS: " + $_.Exception.Message})
    }
}

# ============================================================================
# RDS FUNCTIONS
# ============================================================================

function Get-RDSSessionHosts {
    try {
        $hosts = Get-RDSessionHost -ErrorAction Stop | Select-Object SessionHost, NewConnectionAllowed | Select-Object -First 50
        return $hosts | ForEach-Object { [PSCustomObject]@{"Host" = $_.SessionHost; "NewConnections" = $_.NewConnectionAllowed} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = "RDS: " + $_.Exception.Message})
    }
}

function Get-RDSActiveLicensing {
    try {
        $licensing = Get-RDLicenseConfiguration -ErrorAction Stop
        return @([PSCustomObject]@{"Mode" = $licensing.Mode; "Server" = $licensing.LicenseServer; "IssuedLicenses" = $licensing.IssuedLicenses})
    } catch {
        return @([PSCustomObject]@{"Fehler" = "RDS: " + $_.Exception.Message})
    }
}

# ============================================================================
# DFS FUNCTIONS
# ============================================================================

function Get-DFSNamespaces {
    try {
        $namespaces = Get-DfsnRoot -ErrorAction Stop | Select-Object Path, Type, State | Select-Object -First 50
        return $namespaces | ForEach-Object { [PSCustomObject]@{"Path" = $_.Path; "Type" = $_.Type; "State" = $_.State} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DFSReplicationGroups {
    try {
        $groups = Get-DfsReplicationGroup -ErrorAction Stop | Select-Object GroupName, State, Description | Select-Object -First 50
        return $groups | ForEach-Object { [PSCustomObject]@{"Name" = $_.GroupName; "State" = $_.State; "Description" = $_.Description} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# PRINT SERVER FUNCTIONS
# ============================================================================

function Get-PrintServers {
    try {
        $printers = Get-Printer -ErrorAction Stop | Select-Object Name, DriverName, Shared, Published | Select-Object -First 50
        return $printers | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Driver" = $_.DriverName; "Shared" = $_.Shared; "Published" = $_.Published} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-PrinterDrivers {
    try {
        $drivers = Get-PrinterDriver -ErrorAction Stop | Select-Object Name, Manufacturer, DriverVersion | Select-Object -First 50
        return $drivers | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Vendor" = $_.Manufacturer; "Version" = $_.DriverVersion} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# WSUS FUNCTIONS
# ============================================================================

function Get-WSUSConfiguration {
    try {
        $wsus = Get-WsusServer -ErrorAction Stop | Select-Object Name, PortNumber, ServerProtocolVersion
        return @([PSCustomObject]@{"Name" = $wsus.Name; "Port" = $wsus.PortNumber; "Protocol" = $wsus.ServerProtocolVersion})
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-WSUSComputerTargetGroups {
    try {
        $server = Get-WsusServer -ErrorAction Stop
        $groups = $server | Get-WsusComputerTargetGroup -ErrorAction Stop | Select-Object Name | Select-Object -First 50
        return $groups | ForEach-Object { [PSCustomObject]@{"Group" = $_.Name} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-WSUSUpdates {
    try {
        $server = Get-WsusServer -ErrorAction Stop
        $updates = $server.GetUpdates() | Select-Object Title, Classification, ApprovedCount | Select-Object -First 50
        return $updates | ForEach-Object { [PSCustomObject]@{"Title" = $_.Title; "Classification" = $_.Classification; "Approved" = $_.ApprovedCount} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# HYPER-V FUNCTIONS
# ============================================================================

function Get-HyperVVirtualMachines {
    try {
        $vms = Get-VM -ErrorAction Stop | Select-Object Name, State, MemoryAssigned, ProcessorCount | Select-Object -First 50
        return $vms | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "State" = $_.State; "RAM(GB)" = [Math]::Round($_.MemoryAssigned / 1GB, 2); "CPUs" = $_.ProcessorCount} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-HyperVSwitches {
    try {
        $switches = Get-VMSwitch -ErrorAction Stop | Select-Object Name, SwitchType, NetAdapterInterfaceDescription | Select-Object -First 50
        return $switches | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Type" = $_.SwitchType; "Adapter" = $_.NetAdapterInterfaceDescription} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-HyperVSnapshots {
    try {
        $snapshots = Get-VMSnapshot -ErrorAction Stop | Select-Object VMName, Name, CreationTime | Select-Object -First 50
        return $snapshots | ForEach-Object { [PSCustomObject]@{"VM" = $_.VMName; "Snapshot" = $_.Name; "Created" = $_.CreationTime} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# NRAS / NPS FUNCTIONS
# ============================================================================

function Get-NPASConfiguration {
    try {
        $nasclients = Get-NpsRadiusClient -ErrorAction Stop | Select-Object Name, Address | Select-Object -First 50
        return $nasclients | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "IP" = $_.Address} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-NASClients {
    try {
        $nasclients = Get-NpsRadiusClient -ErrorAction Stop | Select-Object Name, Address | Select-Object -First 50
        return $nasclients | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "IP" = $_.Address} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# KMS FUNCTIONS
# ============================================================================

function Get-KMSConfiguration {
    try {
        $kms = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction Stop
        return @([PSCustomObject]@{"ServiceRunning" = $kms.IsS_Running; "Version" = $kms.Version; "VLActivationInterval" = $kms.VLActivationInterval})
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# WDS FUNCTIONS
# ============================================================================

function Get-WDSConfiguration {
    try {
        $wdsReg = Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\WDSServer\Providers\WDSPXE" -ErrorAction SilentlyContinue
        if ($wdsReg) {
            return @([PSCustomObject]@{"Status" = "WDS Service existiert"; "Version" = $wdsReg.Version})
        } else {
            return @([PSCustomObject]@{"Status" = "WDS Service nicht gefunden"})
        }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-WDSBootImages {
    try {
        $wdsPath = "C:\RemoteInstall\Boot" # Standard WDS Path
        if (Test-Path $wdsPath) {
            $images = Get-ChildItem -Path $wdsPath -Filter "*.wim" -ErrorAction SilentlyContinue | Select-Object -First 50
            return $images | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Size(MB)" = [Math]::Round($_.Length / 1MB, 2)} }
        } else {
            return @([PSCustomObject]@{"Info" = "WDS Boot Images Pfad nicht gefunden"})
        }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-WDSInstallImages {
    try {
        $wdsPath = "C:\RemoteInstall\Images" # Standard WDS Path
        if (Test-Path $wdsPath) {
            $images = Get-ChildItem -Path $wdsPath -Filter "*.wim" -ErrorAction SilentlyContinue | Select-Object -First 50
            return $images | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Size(MB)" = [Math]::Round($_.Length / 1MB, 2)} }
        } else {
            return @([PSCustomObject]@{"Info" = "WDS Install Images Pfad nicht gefunden"})
        }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# FILE SERVICES & SMB FUNCTIONS
# ============================================================================

function Get-FileSharePermissions {
    try {
        $shares = Get-SmbShare -ErrorAction Stop | Select-Object Name | Select-Object -First 20
        $perms = @()
        foreach ($share in $shares) {
            $sharePerm = Get-SmbShareAccess -Name $share.Name -ErrorAction SilentlyContinue
            $perms += $sharePerm
        }
        return $perms | Select-Object -First 100 | ForEach-Object { [PSCustomObject]@{"Share" = $_.Name; "Account" = $_.AccountName; "Access" = $_.AccessRight} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-FileServerQuotas {
    try {
        $quotas = Get-FsrmQuota -ErrorAction Stop | Select-Object Path, Size, SoftLimit | Select-Object -First 50
        return $quotas | ForEach-Object { [PSCustomObject]@{"Path" = $_.Path; "Size(MB)" = [Math]::Round($_.Size / 1MB, 2); "SoftLimit" = $_.SoftLimit} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ShadowCopies {
    try {
        $shadows = Get-CimInstance -ClassName Win32_ShadowCopy -ErrorAction Stop | Select-Object VolumeName, InstallDate, ID | Select-Object -First 50
        return $shadows | ForEach-Object { [PSCustomObject]@{"Volume" = $_.VolumeName; "Date" = $_.InstallDate; "ID" = $_.ID.Substring(0, [Math]::Min(20, $_.ID.Length))} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# EXTENDED AD FUNCTIONS
# ============================================================================

function Get-ADDomainControllerExtended {
    try {
        $dcs = Get-ADDomainController -Filter * -ErrorAction Stop
        if ($null -eq $dcs) {
            return @([PSCustomObject]@{"Information" = "Keine Domain Controller gefunden"})
        }
        $data = @()
        foreach ($dc in $dcs) {
            $data += @{
                "Name" = $dc.Name
                "Site" = $dc.Site
                "OS" = $dc.OperatingSystem
                "GC" = if($dc.IsGlobalCatalog) {"Ja"} else {"Nein"}
                "RODC" = if($dc.IsReadOnly) {"Ja"} else {"Nein"}
                "IPv4" = $dc.IPv4Address
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADDomainFunctionalLevel {
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $forest = Get-ADForest -ErrorAction Stop
        return @(
            [PSCustomObject]@{"Property" = "Domain Name"; "Value" = $domain.Name},
            [PSCustomObject]@{"Property" = "Domain Mode"; "Value" = $domain.DomainMode},
            [PSCustomObject]@{"Property" = "Forest Name"; "Value" = $forest.Name},
            [PSCustomObject]@{"Property" = "Forest Mode"; "Value" = $forest.ForestMode}
        )
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADSiteConfiguration {
    try {
        $sites = Get-ADReplicationSite -ErrorAction Stop | Select-Object Name, Description | Select-Object -First 50
        return $sites | ForEach-Object { [PSCustomObject]@{"Site" = $_.Name; "Description" = $_.Description} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADGroupPolicySummary {
    try {
        $gpos = Get-GPO -All -ErrorAction Stop | Measure-Object
        $topGPOs = Get-GPO -All -ErrorAction Stop | Select-Object DisplayName, CreationTime | Sort-Object CreationTime -Descending | Select-Object -First 30
        return $topGPOs | ForEach-Object { [PSCustomObject]@{"Name" = $_.DisplayName; "Created" = $_.CreationTime} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADClusterInformation {
    try {
        $cluster = Get-Cluster -ErrorAction Stop
        $nodes = Get-ClusterNode -ErrorAction Stop | Select-Object Name, State | Select-Object -First 50
        return $nodes | ForEach-Object { [PSCustomObject]@{"Node" = $_.Name; "State" = $_.State; "Cluster" = $cluster.Name} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = "Failover Clustering: " + $_.Exception.Message})
    }
}

# ============================================================================
# EXTENDED AD FUNCTIONS (ADDITIONAL)
# ============================================================================

function Get-ADUserAccounts {
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $forest = Get-ADForest -ErrorAction Stop
        $data = @(
            @{"FSMO-Rolle" = "PDC Emulator"; "Inhaber" = $domain.PDCEmulator},
            @{"FSMO-Rolle" = "RID Master"; "Inhaber" = $domain.RIDMaster},
            @{"FSMO-Rolle" = "Infrastructure Master"; "Inhaber" = $domain.InfrastructureMaster},
            @{"FSMO-Rolle" = "Schema Master"; "Inhaber" = $forest.SchemaMaster},
            @{"FSMO-Rolle" = "Domain Naming Master"; "Inhaber" = $forest.DomainNamingMaster}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADGroupAccounts {
    try {
        $forest = Get-ADForest -ErrorAction Stop
        $data = @(
            @{"Property" = "Schema Version"; "Value" = $forest.SchemaVersion},
            @{"Property" = "Forest Mode"; "Value" = $forest.ForestMode},
            @{"Property" = "Exchange Version"; "Value" = $forest.ExchangeVersion},
            @{"Property" = "Domains Count"; "Value" = ($forest.Domains | Measure-Object).Count},
            @{"Property" = "Sites Count"; "Value" = ($forest.Sites | Measure-Object).Count}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADServiceAccounts {
    try {
        $features = Get-ADOptionalFeature -Filter * -ErrorAction Stop | Select-Object Name, @{Name='Status';Expression={if($_.EnabledScopes.Count -gt 0) {"Enabled"} else {"Disabled"}}} | Sort-Object Name | Select-Object -First 50
        if ($null -eq $features) {
            return @([PSCustomObject]@{"Information" = "Keine optionalen Features gefunden"})
        }
        return $features | ForEach-Object { [PSCustomObject]@{"Feature" = $_.Name; "Status" = $_.Status} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADComputerAccounts {
    try {
        $dcs = Get-ADDomainController -Filter * -ErrorAction Stop | Select-Object Name, Site, IPv4Address, OperatingSystem, IsGlobalCatalog, IsReadOnly
        if ($null -eq $dcs) {
            return @([PSCustomObject]@{"Information" = "Keine DCs gefunden"})
        }
        return $dcs | ForEach-Object { [PSCustomObject]@{"DC" = $_.Name; "Site" = $_.Site; "IP" = $_.IPv4Address; "OS" = $_.OperatingSystem; "GC" = $_.IsGlobalCatalog; "RODC" = $_.IsReadOnly} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADPasswordPolicy {
    try {
        $trusts = Get-ADTrust -Filter * -ErrorAction Stop | Select-Object Name, Direction, TrustType, TrustAttributes | Sort-Object Name | Select-Object -First 50
        if ($null -eq $trusts) {
            return @([PSCustomObject]@{"Information" = "Keine Vertrauensbeziehungen gefunden"})
        }
        return $trusts | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Direction" = $_.Direction; "Type" = $_.TrustType} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADReplicationSites {
    try {
        $gpos = Get-GPO -All -ErrorAction Stop | Select-Object DisplayName, CreationTime, ModificationTime | Sort-Object ModificationTime -Descending | Select-Object -First 50
        if ($null -eq $gpos) {
            return @([PSCustomObject]@{"Information" = "Keine GPOs gefunden"})
        }
        return $gpos | ForEach-Object { [PSCustomObject]@{"Name" = $_.DisplayName; "Created" = $_.CreationTime; "Modified" = $_.ModificationTime} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADReplicationSubnets {
    try {
        $dcs = Get-ADDomainController -Filter * -ErrorAction Stop
        $data = @()
        foreach ($dc in $dcs) {
            try {
                $services = @("NTDS", "DFSR", "DNS", "KDC") | ForEach-Object { 
                    $svc = Get-Service -Name $_ -ComputerName $dc.HostName -ErrorAction SilentlyContinue
                    if ($null -ne $svc) { "$_`:$($svc.Status.ToString())" }
                }
                $data += @{
                    "DC" = $dc.Name
                    "Reachable" = "Yes"
                    "Services" = ($services -join ", ")
                }
            } catch {
                $data += @{
                    "DC" = $dc.Name
                    "Reachable" = "No"
                    "Services" = "Unable to check"
                }
            }
        }
        if ($data.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "Keine DC-Gesundheitsdaten verfügbar"})
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ADReplicationSiteLinks {
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $dcSites = Get-ADDomainController -Filter * -ErrorAction Stop | Group-Object Site
        $data = @()
        foreach ($siteGroup in $dcSites) {
            $data += @{
                "Site" = $siteGroup.Name
                "DC_Count" = ($siteGroup.Group | Measure-Object).Count
                "DCs" = ($siteGroup.Group.Name -join ", ")
            }
        }
        if ($data.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "Keine Site-Informationen verfügbar"})
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# EXTENDED DNS FUNCTIONS (ADDITIONAL)
# ============================================================================

function Get-DNSResourceRecords {
    try {
        $zones = Get-DnsServerZone -ErrorAction Stop | Select-Object -ExpandProperty ZoneName
        $records = @()
        foreach ($zone in $zones | Select-Object -First 10) {
            $zoneRecords = Get-DnsServerResourceRecord -ZoneName $zone -ErrorAction SilentlyContinue | Select-Object -First 30
            $records += $zoneRecords | Select-Object -Property Name, RecordType, @{Name='TTL';Expression={$_.TTL}}, @{Name='Zone';Expression={$zone}} | Select-Object -First 50
        }
        if ($null -eq $records -or $records.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "Keine DNS-Ressourcen-Einträge gefunden"})
        }
        return $records | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Type" = $_.RecordType; "TTL" = $_.TTL; "Zone" = $_.Zone} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DNSServerStatistics {
    try {
        $stats = Get-DnsServerStatistics -ErrorAction Stop
        if ($null -eq $stats) {
            return @([PSCustomObject]@{"Information" = "Keine DNS-Statistiken verfügbar"})
        }
        return @(
            [PSCustomObject]@{"Statistik" = "Abfragen (Total)"; "Wert" = $stats.Stats.TotalQueries},
            [PSCustomObject]@{"Statistik" = "Responses (Total)"; "Wert" = $stats.Stats.TotalResponses},
            [PSCustomObject]@{"Statistik" = "Cache Hits"; "Wert" = $stats.Stats.CacheHits},
            [PSCustomObject]@{"Statistik" = "Cache Misses"; "Wert" = $stats.Stats.CacheMisses}
        )
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DNSServerScavengingSettings {
    try {
        $scavenging = Get-DnsServerScavenging -ErrorAction Stop
        if ($null -eq $scavenging) {
            return @([PSCustomObject]@{"Information" = "Keine Scavenging-Einstellungen gefunden"})
        }
        return @(
            [PSCustomObject]@{"Setting" = "ScavengingInterval"; "Wert" = $scavenging.ScavengingInterval},
            [PSCustomObject]@{"Setting" = "RefreshInterval"; "Wert" = $scavenging.RefreshInterval},
            [PSCustomObject]@{"Setting" = "NoRefreshInterval"; "Wert" = $scavenging.NoRefreshInterval},
            [PSCustomObject]@{"Setting" = "ScavengingState"; "Wert" = $scavenging.ScavengingState}
        )
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DNSSecSettings {
    try {
        $dnssec = Get-DnsServerDnsSec -ErrorAction Stop | Select-Object -First 50
        if ($null -eq $dnssec) {
            return @([PSCustomObject]@{"Information" = "Keine DNSSEC-Einstellungen gefunden"})
        }
        return $dnssec | ForEach-Object { [PSCustomObject]@{"Zone" = $_.ZoneName; "DNSSEC" = $_.DnsSecState; "KSKCount" = $_.KSKCount; "ZSKCount" = $_.ZSKCount} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# EXTENDED DHCP FUNCTIONS (ADDITIONAL)
# ============================================================================

function Get-DHCPServerInformation {
    try {
        $servers = Get-DhcpServerInDC -ErrorAction Stop
        if ($null -eq $servers) {
            return @([PSCustomObject]@{"Information" = "Keine DHCP-Server in AD gefunden"})
        }
        return $servers | ForEach-Object { [PSCustomObject]@{"Server" = $_.ToString()} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DHCPServerLeases {
    try {
        $scopes = Get-DhcpServerv4Scope -ErrorAction Stop
        $leases = @()
        foreach ($scope in $scopes | Select-Object -First 5) {
            $scopeLeases = Get-DhcpServerv4Lease -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue | Select-Object -First 50
            $leases += $scopeLeases
        }
        if ($null -eq $leases -or $leases.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "Keine aktiven DHCP-Leases gefunden"})
        }
        return $leases | ForEach-Object { [PSCustomObject]@{"IP" = $_.IPAddress; "Scope" = $_.ScopeId; "Client" = $_.HostName; "LeaseExpires" = $_.LeaseExpiryTime} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DHCPServerOptions {
    try {
        $options = Get-DhcpServerv4OptionValue -ErrorAction Stop | Select-Object OptionId, Name, Value | Sort-Object OptionId | Select-Object -First 50
        if ($null -eq $options) {
            return @([PSCustomObject]@{"Information" = "Keine DHCP-Server-Optionen gefunden"})
        }
        return $options | ForEach-Object { [PSCustomObject]@{"OptionID" = $_.OptionId; "Name" = $_.Name; "Value" = $_.Value} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DHCPv4ScopeStatistics {
    try {
        $scopes = Get-DhcpServerv4Scope -ErrorAction Stop
        if ($null -eq $scopes) {
            return @([PSCustomObject]@{"Information" = "Keine DHCP-Scopes gefunden"})
        }
        return $scopes | ForEach-Object { 
            $stats = Get-DhcpServerv4ScopeStatistics -ScopeId $_.ScopeId -ErrorAction SilentlyContinue
            [PSCustomObject]@{
                "Scope" = $_.ScopeId
                "Name" = $_.Name
                "AddressesInUse" = $stats.AddressesInUse
                "AddressesAvailable" = $stats.AddressesAvailable
                "PercentageInUse" = $stats.PercentageInUse
            }
        }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DHCPServerAuthorization {
    try {
        $auth = Get-DhcpServerv4Failover -ErrorAction Stop | Select-Object Name, State, Mode, ServerRole | Select-Object -First 50
        if ($null -eq $auth) {
            return @([PSCustomObject]@{"Information" = "Keine DHCP-Failover-Konfiguration gefunden"})
        }
        return $auth | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "State" = $_.State; "Mode" = $_.Mode; "Role" = $_.ServerRole} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# HTML EXPORT FUNCTIONS
# ============================================================================

function Export-AuditResultsToHtml {
    param(
        [string]$Title,
        [PSObject[]]$Data,
        [string]$Category
    )
    
    try {
        if ($null -eq $Data -or $Data.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Keine Daten zum Exportieren vorhanden!", "Warnung", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return $false
        }
        
        # HTML Header mit modernem Styling
        $htmlHeader = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>easyWSAudit - $Title</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #333;
            padding: 20px;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 8px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 {
            font-size: 28px;
            margin-bottom: 10px;
        }
        .header p {
            font-size: 14px;
            opacity: 0.9;
        }
        .content {
            padding: 30px;
        }
        .audit-info {
            background: #f8f9fa;
            border-left: 4px solid #667eea;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
        }
        .audit-info p {
            margin: 5px 0;
            font-size: 14px;
            line-height: 1.6;
        }
        .audit-info strong {
            color: #667eea;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        thead {
            background: #f3f4f6;
        }
        th {
            padding: 12px;
            text-align: left;
            font-weight: 600;
            color: #374151;
            border-bottom: 2px solid #667eea;
        }
        td {
            padding: 12px;
            border-bottom: 1px solid #e5e7eb;
        }
        tr:nth-child(even) {
            background: #f9fafb;
        }
        tr:hover {
            background: #eff6ff;
        }
        .footer {
            background: #f3f4f6;
            padding: 20px 30px;
            text-align: center;
            font-size: 12px;
            color: #6b7280;
            border-top: 1px solid #e5e7eb;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>⚙️ easyWSAudit - Audit Bericht</h1>
            <p>Windows Server Audit Tool v0.4.0</p>
        </div>
        <div class="content">
            <div class="audit-info">
                <p><strong>Kategorie:</strong> $Category</p>
                <p><strong>Prüfung:</strong> $Title</p>
                <p><strong>Server:</strong> $(hostname)</p>
                <p><strong>Domain:</strong> $([System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName)</p>
                <p><strong>Erstellt:</strong> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
                <p><strong>Ergebnisse:</strong> $($Data.Count) Einträge</p>
            </div>
"@

        # Tabelle mit Daten konvertieren
        $htmlTable = $Data | ConvertTo-Html -Fragment -As Table
        
        # HTML Footer
        $htmlFooter = @"
            $htmlTable
        </div>
        <div class="footer">
            <p>easyWSAudit © 2025 | Generiert: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
        </div>
    </div>
</body>
</html>
"@

        # Dateiname mit Zeitstempel erstellen
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $safeTitle = ($Title -replace '[^a-zA-Z0-9äöüÄÖÜß_]', '_').Substring(0, [Math]::Min(30, $Title.Length))
        $fileName = "AuditReport_${safeTitle}_${timestamp}.html"
        
        # SaveFileDialog anzeigen
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = "HTML Files (*.html)|*.html|All Files (*.*)|*.*"
        $dialog.FileName = $fileName
        $dialog.DefaultExt = "html"
        $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
        
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $fullContent = $htmlHeader + "`r`n" + $htmlTable + "`r`n" + $htmlFooter
            $fullContent | Out-File -FilePath $dialog.FileName -Encoding UTF8 -Force
            
            [System.Windows.MessageBox]::Show("HTML-Export erfolgreich!`n`nDatei: $($dialog.FileName)", "Erfolg", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            
            # Optional: HTML-Datei im Standard-Browser öffnen
            try {
                Start-Process $dialog.FileName
            } catch { }
            
            return $true
        }
        
        return $false
    } catch {
        [System.Windows.MessageBox]::Show("HTML-Export Fehler: $_", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

# ============================================================================
# EVENT HANDLERS
# ============================================================================

# System Information Buttons
$window.FindName("ButtonSystemInfo").Add_Click({
    Update-Status "System Information wird geladen..."
    Update-Output (Get-SystemInformation)
    Update-Status "Fertig"
})

$window.FindName("ButtonOSInfo").Add_Click({
    Update-Status "OS Details werden geladen..."
    Update-Output (Get-OSDetails)
    Update-Status "Fertig"
})

$window.FindName("ButtonHardwareInfo").Add_Click({
    Update-Status "Hardware Info wird geladen..."
    Update-Output (Get-HardwareSummary)
    Update-Status "Fertig"
})

$window.FindName("ButtonCPUInfo").Add_Click({
    Update-Status "CPU Details werden geladen..."
    Update-Output (Get-CPUDetails)
    Update-Status "Fertig"
})

$window.FindName("ButtonMemoryInfo").Add_Click({
    Update-Status "Memory Details werden geladen..."
    Update-Output (Get-MemoryDetails)
    Update-Status "Fertig"
})

$window.FindName("ButtonStorageInfo").Add_Click({
    Update-Status "Storage Info wird geladen..."
    Update-Output (Get-StorageSummary)
    Update-Status "Fertig"
})

# Network Buttons
$window.FindName("ButtonNetConfig").Add_Click({
    Update-Status "Netzwerk IP Konfiguration wird geladen..."
    Update-Output (Get-NetworkConfiguration)
    Update-Status "Fertig"
})

$window.FindName("ButtonNetAdapters").Add_Click({
    Update-Status "Netzwerk Adapter werden geladen..."
    Update-Output (Get-NetworkAdapters)
    Update-Status "Fertig"
})

$window.FindName("ButtonTCPConnections").Add_Click({
    Update-Status "Aktive Verbindungen werden geladen..."
    Update-Output (Get-ActiveConnections)
    Update-Status "Fertig"
})

$window.FindName("ButtonFirewallRules").Add_Click({
    Update-Status "Firewall Regeln werden geladen..."
    Update-Output (Get-FirewallRules)
    Update-Status "Fertig"
})

# Services Buttons
$window.FindName("ButtonAutomaticServices").Add_Click({
    Update-Status "Automatische Services werden geladen..."
    Update-Output (Get-AutomaticServices)
    Update-Status "Fertig"
})

$window.FindName("ButtonRunningServices").Add_Click({
    Update-Status "Laufende Services werden geladen..."
    Update-Output (Get-RunningServices)
    Update-Status "Fertig"
})

$window.FindName("ButtonScheduledTasks").Add_Click({
    Update-Status "Geplante Tasks werden geladen..."
    Update-Output (Get-ScheduledTasks)
    Update-Status "Fertig"
})

# Roles & Features Buttons
$window.FindName("ButtonInstalledFeatures").Add_Click({
    Update-Status "Installierte Features werden geladen..."
    Update-Output (Get-InstalledFeatures)
    Update-Status "Fertig"
})

$window.FindName("ButtonInstalledPrograms").Add_Click({
    Update-Status "Installierte Programme werden geladen..."
    Update-Output (Get-InstalledPrograms)
    Update-Status "Fertig"
})

$window.FindName("ButtonWindowsUpdates").Add_Click({
    Update-Status "Windows Updates werden geladen..."
    Update-Output (Get-WindowsUpdates)
    Update-Status "Fertig"
})

# Event Logs Buttons
$window.FindName("ButtonSystemEvents").Add_Click({
    Update-Status "System Events werden geladen..."
    Update-Output (Get-SystemEvents)
    Update-Status "Fertig"
})

$window.FindName("ButtonAppEvents").Add_Click({
    Update-Status "Application Events werden geladen..."
    Update-Output (Get-ApplicationEvents)
    Update-Status "Fertig"
})

$window.FindName("ButtonSecurityEvents").Add_Click({
    Update-Status "Security Events werden geladen..."
    Update-Output (Get-SecurityEvents)
    Update-Status "Fertig"
})

$window.FindName("ButtonFailedLogons").Add_Click({
    Update-Status "Failed Logons werden geladen..."
    Update-Output (Get-FailedLogons)
    Update-Status "Fertig"
})

$window.FindName("ButtonAccountLockouts").Add_Click({
    Update-Status "Account Lockouts werden geladen..."
    Update-Output (Get-AccountLockouts)
    Update-Status "Fertig"
})

# Security & Users Buttons
$window.FindName("ButtonLocalUsers").Add_Click({
    Update-Status "Lokale Benutzer werden geladen..."
    Update-Output (Get-LocalUsers)
    Update-Status "Fertig"
})

$window.FindName("ButtonLocalGroups").Add_Click({
    Update-Status "Lokale Gruppen werden geladen..."
    Update-Output (Get-LocalGroups)
    Update-Status "Fertig"
})

$window.FindName("ButtonPrivilegeAudit").Add_Click({
    Update-Status "Privilege Audit wird geladen..."
    Update-Output (Get-PrivilegeAudit)
    Update-Status "Fertig"
})

# Active Directory Buttons
$window.FindName("ButtonADDC").Add_Click({
    Update-Status "AD Domain Controller werden geladen..."
    Update-Output (Get-ADDomainControllers)
    Update-Status "Fertig"
})

$window.FindName("ButtonADDomain").Add_Click({
    Update-Status "AD Domain Info wird geladen..."
    Update-Output (Get-ADDomainInfo)
    Update-Status "Fertig"
})

$window.FindName("ButtonADForest").Add_Click({
    Update-Status "AD Forest Info wird geladen..."
    Update-Output (Get-ADForestInfo)
    Update-Status "Fertig"
})

$window.FindName("ButtonADOUs").Add_Click({
    Update-Status "AD OUs werden geladen..."
    Update-Output (Get-ADOUs)
    Update-Status "Fertig"
})

$window.FindName("ButtonADAdmins").Add_Click({
    Update-Status "AD Admins werden geladen..."
    Update-Output (Get-ADDomainAdmins)
    Update-Status "Fertig"
})

$window.FindName("ButtonADComputers").Add_Click({
    Update-Status "AD Computer werden geladen..."
    Update-Output (Get-ADComputers)
    Update-Status "Fertig"
})

$window.FindName("ButtonADReplStatus").Add_Click({
    Update-Status "AD Replication Status wird geladen..."
    Update-Output (Get-ADReplicationStatus)
    Update-Status "Fertig"
})

$window.FindName("ButtonADTrusts").Add_Click({
    Update-Status "AD Trusts werden geladen..."
    Update-Output (Get-ADTrusts)
    Update-Status "Fertig"
})

$window.FindName("ButtonADDCExtended").Add_Click({
    Update-Status "AD DC Extended wird geladen..."
    Update-Output (Get-ADDomainControllerExtended)
    Update-Status "Fertig"
})

$window.FindName("ButtonADFunctionalLevels").Add_Click({
    Update-Status "AD Functional Levels werden geladen..."
    Update-Output (Get-ADDomainFunctionalLevel)
    Update-Status "Fertig"
})

$window.FindName("ButtonADSites").Add_Click({
    Update-Status "AD Sites werden geladen..."
    Update-Output (Get-ADSiteConfiguration)
    Update-Status "Fertig"
})

$window.FindName("ButtonADGPO").Add_Click({
    Update-Status "AD GPOs werden geladen..."
    Update-Output (Get-ADGroupPolicySummary)
    Update-Status "Fertig"
})

$window.FindName("ButtonADClustering").Add_Click({
    Update-Status "AD Clustering wird geladen..."
    Update-Output (Get-ADClusterInformation)
    Update-Status "Fertig"
})

# DNS Buttons
$window.FindName("ButtonDNSConfig").Add_Click({
    Update-Status "DNS Konfiguration wird geladen..."
    Update-Output (Get-DNSConfiguration)
    Update-Status "Fertig"
})

$window.FindName("ButtonDNSZones").Add_Click({
    Update-Status "DNS Zones werden geladen..."
    Update-Output (Get-DNSZones)
    Update-Status "Fertig"
})

$window.FindName("ButtonDNSForwarders").Add_Click({
    Update-Status "DNS Forwarders werden geladen..."
    Update-Output (Get-DNSForwarders)
    Update-Status "Fertig"
})

$window.FindName("ButtonDNSCache").Add_Click({
    Update-Status "DNS Cache wird geladen..."
    Update-Output (Get-DNSCache)
    Update-Status "Fertig"
})

# DHCP Buttons
$window.FindName("ButtonDHCPConfig").Add_Click({
    Update-Status "DHCP Konfiguration wird geladen..."
    Update-Output (Get-DHCPConfiguration)
    Update-Status "Fertig"
})

$window.FindName("ButtonDHCPv4Scopes").Add_Click({
    Update-Status "DHCP IPv4 Scopes werden geladen..."
    Update-Output (Get-DHCPv4Scopes)
    Update-Status "Fertig"
})

$window.FindName("ButtonDHCPv6Scopes").Add_Click({
    Update-Status "DHCP IPv6 Scopes werden geladen..."
    Update-Output (Get-DHCPv6Scopes)
    Update-Status "Fertig"
})

$window.FindName("ButtonDHCPReservations").Add_Click({
    Update-Status "DHCP Reservations werden geladen..."
    Update-Output (Get-DHCPReservations)
    Update-Status "Fertig"
})

# IIS Buttons
$window.FindName("ButtonIISWebsites").Add_Click({
    Update-Status "IIS Websites werden geladen..."
    Update-Output (Get-IISWebsites)
    Update-Status "Fertig"
})

$window.FindName("ButtonIISAppPools").Add_Click({
    Update-Status "IIS App Pools werden geladen..."
    Update-Output (Get-IISAppPools)
    Update-Status "Fertig"
})

$window.FindName("ButtonIISBindings").Add_Click({
    Update-Status "IIS Bindings werden geladen..."
    Update-Output (Get-IISBindings)
    Update-Status "Fertig"
})

# RDS Buttons
$window.FindName("ButtonRDSCollections").Add_Click({
    Update-Status "RDS Collections werden geladen..."
    Update-Output (Get-RDSCollections)
    Update-Status "Fertig"
})

$window.FindName("ButtonRDSSessionHosts").Add_Click({
    Update-Status "RDS Session Hosts werden geladen..."
    Update-Output (Get-RDSSessionHosts)
    Update-Status "Fertig"
})

$window.FindName("ButtonRDSLicensing").Add_Click({
    Update-Status "RDS Licensing wird geladen..."
    Update-Output (Get-RDSActiveLicensing)
    Update-Status "Fertig"
})

# DFS Buttons
$window.FindName("ButtonDFSNamespaces").Add_Click({
    Update-Status "DFS Namespaces werden geladen..."
    Update-Output (Get-DFSNamespaces)
    Update-Status "Fertig"
})

$window.FindName("ButtonDFSReplication").Add_Click({
    Update-Status "DFS Replication Groups werden geladen..."
    Update-Output (Get-DFSReplicationGroups)
    Update-Status "Fertig"
})

# Print Server Buttons
$window.FindName("ButtonPrinters").Add_Click({
    Update-Status "Printers werden geladen..."
    Update-Output (Get-PrintServers)
    Update-Status "Fertig"
})

$window.FindName("ButtonPrinterDrivers").Add_Click({
    Update-Status "Printer Drivers werden geladen..."
    Update-Output (Get-PrinterDrivers)
    Update-Status "Fertig"
})

# WSUS Buttons
$window.FindName("ButtonWSUSConfig").Add_Click({
    Update-Status "WSUS Konfiguration wird geladen..."
    Update-Output (Get-WSUSConfiguration)
    Update-Status "Fertig"
})

$window.FindName("ButtonWSUSGroups").Add_Click({
    Update-Status "WSUS Computer Target Groups werden geladen..."
    Update-Output (Get-WSUSComputerTargetGroups)
    Update-Status "Fertig"
})

$window.FindName("ButtonWSUSUpdates").Add_Click({
    Update-Status "WSUS Updates werden geladen..."
    Update-Output (Get-WSUSUpdates)
    Update-Status "Fertig"
})

# Hyper-V Buttons
$window.FindName("ButtonHyperVVMs").Add_Click({
    Update-Status "Hyper-V VMs werden geladen..."
    Update-Output (Get-HyperVVirtualMachines)
    Update-Status "Fertig"
})

$window.FindName("ButtonHyperVSwitches").Add_Click({
    Update-Status "Hyper-V Switches werden geladen..."
    Update-Output (Get-HyperVSwitches)
    Update-Status "Fertig"
})

$window.FindName("ButtonHyperVSnapshots").Add_Click({
    Update-Status "Hyper-V Snapshots werden geladen..."
    Update-Output (Get-HyperVSnapshots)
    Update-Status "Fertig"
})

# NRAS/NPS Buttons
$window.FindName("ButtonNPASConfig").Add_Click({
    Update-Status "NRAS/NPS Konfiguration wird geladen..."
    Update-Output (Get-NPASConfiguration)
    Update-Status "Fertig"
})

$window.FindName("ButtonNASClients").Add_Click({
    Update-Status "NAS Clients werden geladen..."
    Update-Output (Get-NASClients)
    Update-Status "Fertig"
})

# KMS Buttons
$window.FindName("ButtonKMSConfig").Add_Click({
    Update-Status "KMS Konfiguration wird geladen..."
    Update-Output (Get-KMSConfiguration)
    Update-Status "Fertig"
})

# WDS Buttons
$window.FindName("ButtonWDSConfig").Add_Click({
    Update-Status "WDS Konfiguration wird geladen..."
    Update-Output (Get-WDSConfiguration)
    Update-Status "Fertig"
})

$window.FindName("ButtonWDSBootImages").Add_Click({
    Update-Status "WDS Boot Images werden geladen..."
    Update-Output (Get-WDSBootImages)
    Update-Status "Fertig"
})

$window.FindName("ButtonWDSInstallImages").Add_Click({
    Update-Status "WDS Install Images werden geladen..."
    Update-Output (Get-WDSInstallImages)
    Update-Status "Fertig"
})

# File Services Buttons
$window.FindName("ButtonFileShares").Add_Click({
    Update-Status "File Shares werden geladen..."
    Update-Output (Get-FileShares)
    Update-Status "Fertig"
})

$window.FindName("ButtonSharePermissions").Add_Click({
    Update-Status "Share Permissions werden geladen..."
    Update-Output (Get-FileSharePermissions)
    Update-Status "Fertig"
})

$window.FindName("ButtonFileQuotas").Add_Click({
    Update-Status "File Quotas werden geladen..."
    Update-Output (Get-FileServerQuotas)
    Update-Status "Fertig"
})

$window.FindName("ButtonShadowCopies").Add_Click({
    Update-Status "Shadow Copies werden geladen..."
    Update-Output (Get-ShadowCopies)
    Update-Status "Fertig"
})

# Export & Clear Buttons
$window.FindName("ButtonExportCSV").Add_Click({
    try {
        if ($null -eq $DataGridResults.ItemsSource -or $DataGridResults.ItemsSource.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Keine Daten zum Exportieren vorhanden!", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # Export-Optionen anzeigen
        $result = [System.Windows.MessageBox]::Show("Wählen Sie das Export-Format:`n`nJa = CSV Export`nNein = HTML Export", "Export-Format", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            # CSV Export
            $dialog = New-Object System.Windows.Forms.SaveFileDialog
            $dialog.Filter = "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            $dialog.DefaultExt = "csv"
            $dialog.FileName = "easyWSAudit_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
            
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $exportPath = $dialog.FileName
                
                # Daten zu CSV exportieren
                $DataGridResults.ItemsSource | Export-Csv -Path $exportPath -Encoding UTF8 -NoTypeInformation -Force
                
                Update-Status "CSV-Export erfolgreich: $exportPath"
                [System.Windows.MessageBox]::Show("CSV-Export erfolgreich abgeschlossen!`n`nDatei: $exportPath", "Erfolg", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
        } else {
            # HTML Export
            $title = $StatusBarText.Text
            if ([string]::IsNullOrWhiteSpace($title)) { $title = "Audit Results" }
            
            Export-AuditResultsToHtml -Title $title -Data $DataGridResults.ItemsSource -Category "Audit"
        }
    } catch {
        Update-Status "Export Fehler: $_"
        [System.Windows.MessageBox]::Show("Fehler beim Export: $_", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$window.FindName("ButtonClearOutput").Add_Click({
    Clear-Output
    Update-Status "Output gelöscht"
})

# ============================================================================
# MAIN
# ============================================================================

$window.ShowDialog()

# ============================================================================
# EXTENDED CERTIFICATE FUNCTIONS
# ============================================================================

function Get-InstalledCertificates {
    try {
        $certs = Get-ChildItem -Path Cert:\LocalMachine\My -ErrorAction Stop | Select-Object -First 50
        if ($null -eq $certs) {
            return @([PSCustomObject]@{"Information" = "Keine Zertifikate gefunden"})
        }
        $data = @()
        foreach ($cert in $certs) {
            $data += @{
                "Thumbprint" = $cert.Thumbprint.Substring(0, 16)
                "Subject" = $cert.Subject
                "Issuer" = $cert.Issuer
                "ValidFrom" = $cert.NotBefore
                "ValidTo" = $cert.NotAfter
                "DaysLeft" = ($cert.NotAfter - (Get-Date)).Days
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-CertificateAuthorities {
    try {
        $cas = Get-ChildItem -Path Cert:\LocalMachine\Root -ErrorAction Stop | Select-Object -First 50
        if ($null -eq $cas) {
            return @([PSCustomObject]@{"Information" = "Keine CAs gefunden"})
        }
        $data = @()
        foreach ($ca in $cas) {
            $data += @{
                "CA-Name" = $ca.Subject
                "Thumbprint" = $ca.Thumbprint.Substring(0, 16)
                "ValidFrom" = $ca.NotBefore
                "ValidTo" = $ca.NotAfter
                "Issuer" = $ca.Issuer
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ExpiringCertificates {
    try {
        $certs = Get-ChildItem -Path Cert:\LocalMachine\My -ErrorAction Stop
        $expiring = @()
        foreach ($cert in $certs) {
            $daysLeft = ($cert.NotAfter - (Get-Date)).Days
            if ($daysLeft -le 90 -and $daysLeft -ge 0) {
                $expiring += @{
                    "Subject" = $cert.Subject
                    "DaysLeft" = $daysLeft
                    "ExpiresOn" = $cert.NotAfter
                    "Status" = if ($daysLeft -le 30) {"KRITISCH"} else {"WARNUNG"}
                }
            }
        }
        if ($expiring.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "Keine ablaufenden Zertifikate in den nächsten 90 Tagen"})
        }
        return $expiring | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

# ============================================================================
# MICROSOFT SERVER ROLES FUNCTIONS
# ============================================================================

function Get-InstalledRoles {
    try {
        $roles = Get-WindowsFeature -ErrorAction Stop | Where-Object { $_.Installed -eq $true } | Select-Object Name, DisplayName, FeatureType
        if ($null -eq $roles) {
            return @([PSCustomObject]@{"Information" = "Keine Rollen/Features installiert"})
        }
        return $roles | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "DisplayName" = $_.DisplayName; "Type" = $_.FeatureType} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-SQLServerInfo {
    try {
        $sqlServices = Get-Service -Name MSSQLSERVER -ErrorAction SilentlyContinue
        if ($null -eq $sqlServices) {
            return @([PSCustomObject]@{"Information" = "SQL Server nicht installiert"})
        }
        
        $data = @(
            @{"Service" = "MSSQLSERVER"; "Status" = $sqlServices.Status},
            @{"Service" = "SQL Server Agent"; "Status" = (Get-Service -Name SQLSERVERAGENT -ErrorAction SilentlyContinue).Status},
            @{"Service" = "SQL Browser"; "Status" = (Get-Service -Name SQLBrowser -ErrorAction SilentlyContinue).Status}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ExchangeServerInfo {
    try {
        $exService = Get-Service -Name MSExchangeServiceHost -ErrorAction SilentlyContinue
        if ($null -eq $exService) {
            return @([PSCustomObject]@{"Information" = "Exchange Server nicht installiert"})
        }
        
        $data = @(
            @{"Component" = "Exchange Service Host"; "Status" = $exService.Status},
            @{"Component" = "Information Store"; "Status" = (Get-Service -Name MSExchangeIS -ErrorAction SilentlyContinue).Status},
            @{"Component" = "Transport"; "Status" = (Get-Service -Name MSExchangeTransport -ErrorAction SilentlyContinue).Status}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-SharePointInfo {
    try {
        $spService = Get-Service -Name SPAdminV4 -ErrorAction SilentlyContinue
        if ($null -eq $spService) {
            return @([PSCustomObject]@{"Information" = "SharePoint nicht installiert"})
        }
        
        $data = @(
            @{"Service" = "SP Admin"; "Status" = $spService.Status},
            @{"Service" = "SP Timer"; "Status" = (Get-Service -Name SPTimerV4 -ErrorAction SilentlyContinue).Status}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-DomainControllerHealth {
    try {
        $dcHealth = dcdiag /v 2>&1 | Select-Object -First 50
        if ($null -eq $dcHealth) {
            return @([PSCustomObject]@{"Information" = "Keine DC Health Informationen verfügbar"})
        }
        $data = @()
        foreach ($line in $dcHealth) {
            if ($line -match "passed|failed") {
                $data += @{"Test" = $line}
            }
        }
        if ($data.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "DC Health Check durchgeführt"})
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-HyperVInfo {
    try {
        $hvService = Get-Service -Name vmms -ErrorAction SilentlyContinue
        if ($null -eq $hvService) {
            return @([PSCustomObject]@{"Information" = "Hyper-V nicht installiert"})
        }
        
        $vms = Get-VM -ErrorAction SilentlyContinue | Measure-Object
        $data = @(
            @{"Component" = "Hyper-V Service"; "Status" = $hvService.Status},
            @{"Component" = "VMs installiert"; "Status" = $vms.Count}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-ServerUpdates {
    try {
        $updates = Get-HotFix -ErrorAction Stop | Sort-Object InstalledOn -Descending | Select-Object -First 20
        if ($null -eq $updates) {
            return @([PSCustomObject]@{"Information" = "Keine Updates installiert"})
        }
        return $updates | ForEach-Object { [PSCustomObject]@{"KB" = $_.HotFixID; "InstalledOn" = $_.InstalledOn; "Description" = $_.Description} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-NetworkSecurity {
    try {
        $firewallEnabled = Get-NetFirewallProfile -All -ErrorAction Stop | Select-Object Name, Enabled
        if ($null -eq $firewallEnabled) {
            return @([PSCustomObject]@{"Information" = "Firewall Status nicht verfügbar"})
        }
        return $firewallEnabled | ForEach-Object { [PSCustomObject]@{"Profile" = $_.Name; "Enabled" = if($_.Enabled) {"Ja"} else {"Nein"}} }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-WindowsDefender {
    try {
        $defender = Get-Service -Name WinDefend -ErrorAction SilentlyContinue
        $defenderStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
        
        if ($null -eq $defender) {
            return @([PSCustomObject]@{"Information" = "Windows Defender nicht verfügbar"})
        }
        
        $data = @(
            @{"Component" = "Service Status"; "Status" = $defender.Status},
            @{"Component" = "Real-time Protection"; "Status" = if($defenderStatus.RealTimeProtectionEnabled) {"Aktiviert"} else {"Deaktiviert"}},
            @{"Component" = "Definitions Update"; "Status" = $defenderStatus.AntivirusSignatureLastUpdated}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-UserAccountControl {
    try {
        $uac = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name EnableLUA -ErrorAction SilentlyContinue
        $data = @(
            @{"Setting" = "User Account Control"; "Status" = if($uac.EnableLUA -eq 1) {"Aktiviert"} else {"Deaktiviert"}}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Fehler" = $_.Exception.Message})
    }
}

function Get-SystemBackupConfig {
    try {
        $backup = Get-WBBackupSet -ErrorAction SilentlyContinue | Select-Object -First 5
        if ($null -eq $backup) {
            return @([PSCustomObject]@{"Information" = "Keine Backups konfiguriert oder Windows Backup nicht aktiv"})
        }
        return $backup | ForEach-Object { [PSCustomObject]@{"BackupSetId" = $_.BackupSetId; "BackupTime" = $_.BackupTime; "Items" = ($_.Items.Count)} }
    } catch {
        return @([PSCustomObject]@{"Information" = "Windows Backup nicht verfügbar - verwende andere Backup-Lösung"})
    }
}
