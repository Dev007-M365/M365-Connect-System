Add-Type -AssemblyName PresentationFramework

$ScriptRoot    = $PSScriptRoot
$TenantCsvPath = Join-Path $ScriptRoot "tenants.csv"

#-------------------------
# Helper: Write to log box
#-------------------------
function Write-Log {
    param([string]$Message, [string]$Color = "Black")
    $null = $LogBox.Dispatcher.Invoke([action]{
        $LogBox.AppendText("$Message`r`n")
        $LogBox.ScrollToEnd()
    })
}

#-------------------------
# Helper: Ensure module
#-------------------------
function Ensure-Module {
    param([string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Log "Installing module: $Name..." "DarkGoldenrod"
        Install-Module $Name -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Name -ErrorAction Stop
    Write-Log "Module ready: $Name" "DarkGreen"
}

#-------------------------
# Connect functions
#-------------------------
function Connect-GraphDelegated {
    param([string]$TenantId)

    Ensure-Module "Microsoft.Graph"

    $Scopes = @(
        "User.Read",
        "User.ReadWrite.All",
        "Directory.ReadWrite.All",
        "Group.ReadWrite.All",
        "Sites.ReadWrite.All",
        "Files.ReadWrite.All",
        "Mail.ReadWrite",
        "Calendars.ReadWrite",
        "Reports.Read.All",
        "AuditLog.Read.All",
        "Team.ReadWrite.All",
        "Channel.ReadWrite.All",
        "OnlineMeetings.ReadWrite",
        "Presence.Read.All"
    )

    Write-Log "Connecting to Microsoft Graph (delegated)..."
    Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -NoWelcome
    Write-Log "Connected to Microsoft Graph (delegated)." "DarkGreen"
}

function Connect-ExchangeOnline {
    param([string]$AdminUPN)

    Ensure-Module "ExchangeOnlineManagement"
    Write-Log "Connecting to Exchange Online as $AdminUPN..."
    Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false
    Write-Log "Connected to Exchange Online." "DarkGreen"
}

function Connect-PnP {
    param([string]$SPOUrl, [string]$AdminUPN)

    if (-not $SPOUrl) {
        Write-Log "No SPO URL defined for this tenant. Skipping PnP." "DarkGoldenrod"
        return
    }

    Ensure-Module "PnP.PowerShell"
    Write-Log "Connecting to SharePoint Online (PnP) at $SPOUrl..."
    Connect-PnPOnline -Url $SPOUrl -Interactive -LoginName $AdminUPN
    Write-Log "Connected to SharePoint Online (PnP)." "DarkGreen"
}

function Connect-Teams {
    param([string]$AdminUPN)

    Ensure-Module "MicrosoftTeams"
    Write-Log "Connecting to Microsoft Teams as $AdminUPN..."
    Connect-MicrosoftTeams -AccountId $AdminUPN
    Write-Log "Connected to Microsoft Teams." "DarkGreen"
}

#-------------------------
# Load tenants
#-------------------------
if (-not (Test-Path $TenantCsvPath)) {
    [System.Windows.MessageBox]::Show("tenant.csv not found in $ScriptRoot","M365-Connect","OK","Error") | Out-Null
    exit
}

$Tenants = Import-Csv $TenantCsvPath
if (-not $Tenants -or $Tenants.Count -eq 0) {
    [System.Windows.MessageBox]::Show("tenant.csv is empty.","M365-Connect","OK","Error") | Out-Null
    exit
}

#-------------------------
# XAML UI
#-------------------------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="M365-Connect SYSTEM" Height="420" Width="700"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <!-- Header -->
        <TextBlock Grid.Row="0" Grid.ColumnSpan="2"
                   Text="M365-Connect SYSTEM"
                   FontSize="20" FontWeight="Bold"
                   Foreground="DarkCyan" Margin="0,0,0,10"/>

        <!-- Tenant selection -->
        <StackPanel Grid.Row="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock Text="Tenant:" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <ComboBox x:Name="TenantCombo" Width="350" DisplayMemberPath="TenantName"/>
        </StackPanel>

        <!-- Profile selection -->
        <StackPanel Grid.Row="2" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock Text="Profile:" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <ComboBox x:Name="ProfileCombo" Width="200">
                <ComboBoxItem Content="Graph only" Tag="Graph"/>
                <ComboBoxItem Content="Teams only" Tag="Teams"/>
                <ComboBoxItem Content="Full Admin (Graph + EXO + PnP + Teams)" Tag="Full"/>
            </ComboBox>
            <Button x:Name="HelpButton" Content="Global Admin / Graph Note" Margin="10,0,0,0" Padding="8,2"/>
        </StackPanel>

        <!-- Log box -->
        <TextBox x:Name="LogBox" Grid.Row="3" Grid.ColumnSpan="2"
                 Margin="0,0,0,10" IsReadOnly="True"
                 VerticalScrollBarVisibility="Auto"
                 HorizontalScrollBarVisibility="Auto"
                 TextWrapping="Wrap"/>

        <!-- Buttons -->
        <StackPanel Grid.Row="4" Grid.Column="0" Orientation="Horizontal" HorizontalAlignment="Left">
            <Button x:Name="ConnectButton" Content="Connect" Width="120" Margin="0,0,10,0" Padding="8,2"/>
            <Button x:Name="CloseButton" Content="Close" Width="80" Padding="8,2"/>
        </StackPanel>

        <TextBlock Grid.Row="4" Grid.Column="1" HorizontalAlignment="Right" VerticalAlignment="Center"
                   Text="PowerShell 7 + Modern Modules (Graph / EXO / PnP / Teams)"
                   Foreground="Gray" FontSize="10"/>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$TenantCombo   = $Window.FindName("TenantCombo")
$ProfileCombo  = $Window.FindName("ProfileCombo")
$LogBox        = $Window.FindName("LogBox")
$ConnectButton = $Window.FindName("ConnectButton")
$CloseButton   = $Window.FindName("CloseButton")
$HelpButton    = $Window.FindName("HelpButton")

# Bind tenants
$Tenants | ForEach-Object { [void]$TenantCombo.Items.Add($_) }
$TenantCombo.SelectedIndex = 0
$ProfileCombo.SelectedIndex = 0

# Help button (GA / Graph note)
$HelpButton.Add_Click({
    $msg = @"
Being a Global Administrator does NOT automatically grant full Microsoft Graph permissions.

Microsoft Graph permissions are granted to APPLICATIONS, not users.

Even if you sign in as a Global Admin, your app registration must still be explicitly
assigned the required Graph permissions (delegated or application) and must have
ADMIN CONSENT granted.

Global Admin = full access in the portal
Global Admin ≠ full access in Microsoft Graph
"@
    [System.Windows.MessageBox]::Show($msg,"Global Admin / Graph Permissions","OK","Information") | Out-Null
})

# Close button
$CloseButton.Add_Click({
    $Window.Close()
})

# Connect button
$ConnectButton.Add_Click({
    $selectedTenant = $TenantCombo.SelectedItem
    if (-not $selectedTenant) {
        Write-Log "No tenant selected." "Red"
        return
    }

    $tenantId = $selectedTenant.TenantId
    $adminUPN = $selectedTenant.AdminUPN
    $spoUrl   = $selectedTenant.SPOUrl

    $profileItem = $ProfileCombo.SelectedItem
    $profileTag  = $profileItem.Tag

    Write-Log "----------------------------------------"
    Write-Log "Tenant : $($selectedTenant.TenantName) [$tenantId]"
    Write-Log "Admin  : $adminUPN"
    Write-Log "SPO    : $spoUrl"
    Write-Log "Profile: $profileTag"
    Write-Log "----------------------------------------"

    try {
        switch ($profileTag) {
            "Graph" {
                Connect-GraphDelegated -TenantId $tenantId
            }
            "Teams" {
                Connect-GraphDelegated -TenantId $tenantId
                Connect-Teams -AdminUPN $adminUPN
            }
            "Full" {
                Connect-GraphDelegated -TenantId $tenantId
                Connect-ExchangeOnline -AdminUPN $adminUPN
                Connect-PnP -SPOUrl $spoUrl -AdminUPN $adminUPN
                Connect-Teams -AdminUPN $adminUPN
            }
        }

        Write-Log "M365-Connect SYSTEM is ready for this tenant/profile." "DarkCyan"
    }
    catch {
        Write-Log "ERROR: $($_.Exception.Message)" "Red"
    }
})

# Show window
$Window.ShowDialog() | Out-Null
