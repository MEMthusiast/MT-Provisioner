<#
.SYNOPSIS
    Multi-Tenant Provisioner (MTP) collects the hardware hash in WinPE and uploads it to Intune Autopilot of the selected tenant and installs an operating system. Optionally run a custom script before OOBE.

.DESCRIPTION
    This script is intended for Managed Service Providers (MSPs) to streamline bare-metal operating system deployment and the upload of device hardware hashes to Microsoft Intune Autopilot for multiple tenants. It is designed to run in conjunction with OSDCloud WinPE.
    The Microsoft Graph authentication works with a multi-tenant Entra ID app registration, enabling the same App ID and App Secret across multiple tenants.

    It performs the following functions:
        1. Displays a selection UI for the user to choose which tenant to upload the hardware hash to. Optionally allow for only installing an operating system with configured parameters in the config file (ParametersUrl or hardcoded parameters).
        2. Optionally collects the hardware hash using OA3Tool, including TPM information by registering the PCPKsp.dll.
        3. Optionally uploads the hardware hash to the selected tenant's Intune Autopilot via Microsoft Graph API.
        4. Optionally waits for the Autopilot profile assignment to complete.
        5. Optionally downloads a SetupComplete.ps1 script and saves it to the appropriate location in the WinPE environment where it later will be copied to be executed before OOBE.
        6. Optionally downloads a tenant configuration JSON file, allowing for dynamic configuration without modifying the script.
        7. Supports retrieving the (multi-tenant) enterprise app secret from Azure Key Vault for enhanced security when allowing the keyvault to only be accessed from a trusted public IP address.

.PARAMETER AppId
    Required. The app registration ID for authentication. This is the AppID of the multitenant Entra ID enterprise app.

.PARAMETER UploadToAutopilot
    Conditionally required. Set this switch to $true or $false to enable or $false to disable the Autopilot upload functionality. This will be overridden by the tenant-specific value if specified.

.PARAMETER GroupTag
    Conditionally required. Set this parameter for tenants that don't have a GroupTag property in the config file (ParametersUrl or hardcoded parameters).

.PARAMETER KeyVault
    Optional. Indicates whether to retrieve the AppSecret from an Azure Key Vault.

.PARAMETER AppSecret
    Conditionally required. The app secret for authentication to the multitenant enterprise app. Required if Key Vault retrieval is disabled ($KeyVault = $false).

.PARAMETER SPNAppID
    Conditionally required. The app ID for authentication to Key Vault. Required if Key Vault retrieval is enabled ($KeyVault = $true).

.PARAMETER SPNSecret
    Conditionally required. The app secret for authentication to Key Vault. Required if Key Vault retrieval is enabled ($KeyVault = $true).

.PARAMETER SPNTenantID
    Conditionally required. The tenant ID for authentication to Key Vault. Required if Key Vault retrieval is enabled ($KeyVault = $true).

.PARAMETER VaultName
    Conditionally required. The name of the Azure Key Vault to retrieve the secret from. Required if Key Vault retrieval is enabled ($KeyVault = $true).

.PARAMETER SecretName
    Conditionally required. The name of the secret in Azure Key Vault that contains the app secret. Required if Key Vault retrieval is enabled ($KeyVault = $true).

.PARAMETER SetupCompleteUrl
    Optional. A URL to download a SetupComplete.ps1 script. This script will be copied to the WinPE image and will run during the SetupComplete phase before OOBE. Example: https://raw.githubusercontent.com/MEMthusiast/MT-Provisioner/refs/heads/main/SetupComplete/SetupComplete.ps1

.PARAMETER ParametersUrl
    Optional. If provided, the script will read tenant information from the downloaded JSON instead of using hardcoded tenant parameters inside this script. Example: https://raw.githubusercontent.com/MEMthusiast/MT-Provisioner/refs/heads/main/TenantsConfig.json

.NOTES
    File Name   : Start-MTP.ps1
    Author      : https://github.com/MEMthusiast
    Version     : 4.05
    Purpose     : Upload device hashes to the selected tenant and install an operating system.
    Requires    : The OSDCloud PowerShell module, a multi-tenant Entra ID enterprise application in each tenant, and optionally an Azure Key Vault for secret retrieval, along with hosting the SetupComplete.ps1 and TenantsConfig.json files in an Azure Blob (that is only accessible from a trusted public IP address).
    References  : Autopilot upload logic in this script is based on: https://github.com/blawal/WinPEAP
    Usage       : Designed for use with OSDCloud WinPE. This script runs automatically when placed in the StartNet folder of OSDCloud WinPE. Adjust parameters as required prior to deployment.
#>

#region: Required Parameters
    $AppId              = ""
#endregion

#region: Conditionally Required Parameters
    $UploadToAutopilot  = $true
    $GroupTag           = ""
    $AppSecret          = ""
#endregion

#region: Optional Parameters
    $SPNAppID           = ""
    $SPNSecret          = ""
    $SPNTenantID        = ""
    $VaultName          = ""
    $SecretName         = ""
    $SetupCompleteUrl   = ""
    $ParametersUrl      = ""
#endregion

#region: Logging
# Tracks whether the previous output was a separator, so Write-Separator can dedupe back-to-back separators caused by skipped if/else branches.
$script:LastWasSeparator = $false

# Write log function with current date and time
function Write-Log {
    param(
        [Parameter(Mandatory)] [string] $Message,
        [ConsoleColor] $Color = [ConsoleColor]::Gray
    )

    $stamp = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
    $line = "[$stamp] $Message"

    Write-Host $line -ForegroundColor $Color
    $script:LastWasSeparator = $false
}

# Plain (non-timestamped) line — used for formatted summary output. Resets the separator flag so a following Write-Separator will still print.
function Write-Plain {
    param(
        [string] $Text = "",
        [ConsoleColor] $Color = [ConsoleColor]::Gray
    )
    Write-Host $Text -ForegroundColor $Color
    $script:LastWasSeparator = $false
}

# Print a horizontal separator line. Skips itself if the previous output was already a separator.
function Write-Separator {
    param(
        [ValidateSet('=','-')] [string] $Char = '=',
        [int] $Length = 70,
        [ConsoleColor] $Color = [ConsoleColor]::DarkGray
    )
    if ($script:LastWasSeparator) { return }
    Write-Host ($Char * $Length) -ForegroundColor $Color
    $script:LastWasSeparator = $true
}

# Start transcript
$Global:Transcript = "$(Get-Date -Format 'dd-MM-yyyy_HHmmss')-Start-MTP.log"
Start-Transcript -Path (Join-Path "X:\OSDCloud\Config\Scripts\SetupComplete" $Global:Transcript) -ErrorAction Ignore
$script:TranscriptActive = $true

# Start timer
$startTime = Get-Date
Write-Log "Multi-Tenant Provisioner Started" DarkMagenta
#endregion

#region: Test internet connectity
Write-Separator
if ($SPNSecret -or $SetupCompleteUrl -or $ParametersUrl -or $UploadToAutopilot) {
    Write-Log "Internet connectivity is required. Checking connection..." Cyan

    $RetryInterval = 10
    $firstAttempt = $true

    while ($true) {
        try {
            # Test internet connectivity
            Invoke-WebRequest -Uri "https://www.microsoft.com/" -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop | Out-Null
            Write-Log "Internet connection detected. Continuing..." Green
            break
        }
        catch {
            if ($firstAttempt) {
                Write-Log "No internet connection detected. Retrying..." DarkRed
                $firstAttempt = $false
            }
                for ($i = $RetryInterval; $i -gt 0; $i--) {
                    Write-Progress -Activity "Connect the device to the network to continue" -Status "Retrying in $i seconds..." -PercentComplete ((($RetryInterval - $i) / $RetryInterval) * 100)
                    Start-Sleep -Seconds 1
                }
                Write-Progress -Activity "Connect the device to the network to continue" -Completed      
        }
    }
}
else {
    Write-Log "No internet-dependent parameters configured" DarkGray
    Write-Log "[WARNING]: Continuing without internet connectivity. Script functionality may be limited" Yellow
}
# Get public IP address
try {
$publicIP = (Invoke-RestMethod -Uri "https://api.ipify.org?format=json").ip
    } catch {} 
#endregion

#region: Injected Tenant Parameters
# If ParametersUrl is provided, read tenant information from the config file instead of using hardcoded parameters.

Write-Separator
if (-not [string]::IsNullOrWhiteSpace($ParametersUrl)) {
    
$DestinationFolder = "X:\OSDCloud\Config\Scripts\SetupComplete"

try {
    # Ensure TLS 1.2
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Extract filename and construct destination path
    $FileName        = Split-Path -Path $ParametersUrl -Leaf
    $DestinationFile = Join-Path -Path $DestinationFolder -ChildPath $FileName

    # Download parameter file
    Write-Log "Downloading $FileName from public IP: $publicIP" DarkGray

    Invoke-WebRequest -Uri $ParametersUrl -OutFile $DestinationFile -UseBasicParsing -ErrorAction Stop

    Write-Log "$FileName downloaded successfully" Green
}
catch {
    Write-Log "Failed to download $FileName : $($_.Exception.Message)" Red
}

$config = Get-Content "$DestinationFile" -Raw | ConvertFrom-Json
$Tenants = $config.Tenants

}
else {
    Write-Log "ParametersUrl not provided. Using hardcoded parameters" DarkGray
}
#endregion

#region: Hardcoded Tenant Parameters
# If ParametersUrl is not provided, use hardcoded tenant information.

if ([string]::IsNullOrWhiteSpace($ParametersUrl)) {

    $Tenants = @(
    [PSCustomObject]@{
        Name = "Tenant A"
        TenantId = "21212121-2121-2121-2121-212121212121"
        UploadToAutopilot = $false
        GroupTag = ""
        OSBuild = "25H2"
        OSEdition = "Pro"
        OSVersion = "Windows 11"
        OSLanguage = "nl-nl"
        OSActivation = "Volume"
    },
    [PSCustomObject]@{
        Name = "Tenant B"
        TenantId = "21212121-2121-2121-2121-212121212121"
        UploadToAutopilot = $true
        GroupTag = "TAG1"
        OSBuild = "25H2"
        OSEdition = "Pro"
        OSVersion = "Windows 11"
        OSLanguage = "nl-nl"
        OSActivation = "Volume"
    },
    [PSCustomObject]@{
        Name = "Tenant C"
        TenantId = "21212121-2121-2121-2121-212121212121"
        UploadToAutopilot = $true
        GroupTag = "TAG2"
        OSBuild = "25H2"
        OSEdition = "Pro"
        OSVersion = "Windows 11"
        OSLanguage = "en-us"
        OSActivation = "Volume"
        Pinned = $true # This tenant will be pinned to the top of the list in the UI
        SetupCompleteUrl = "" # Optionally add a tenant-specific SetupComplete.ps1 URL.
    }
)

}
else {
    Write-Log "ParametersUrl provided. Using downloaded config file" DarkGray
}
#endregion

#region: Tenant Selector
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# Sort tenants alphabetically and pinned tenants on top
$AllTenants = @(
    $Tenants | Sort-Object @{ Expression = { -not [bool]$_.Pinned } }, Name
)

# Unicode normalization for names with special characters (mojibake)
$Tenants | ForEach-Object {
    if ($null -ne $_ -and $null -ne $_.Name) {
        if ($_.Name -match 'Ã.|Â.| ') {
            $_.Name = [System.Text.Encoding]::UTF8.GetString(
                [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes([string]$_.Name)
            )
        }
    }
}

# Validation Rule Engine
$ValidationRules = @(
    @{
        Level   = "Error"
        Test    = {
            param($t)
            ($t.UploadToAutopilot) -and [string]::IsNullOrWhiteSpace([string]$t.TenantId)
        }
        Message = { "TenantID is required for Autopilot upload. Unable to start." }
    },
    @{
        Level   = "Warning"
        Test    = {
            param($t)
            ($t.UploadToAutopilot) -and
            (-not [string]::IsNullOrWhiteSpace([string]$t.TenantId)) -and
            ([string]::IsNullOrWhiteSpace([string]$t.GroupTag))
        }
        Message = { "GroupTag is missing for Autopilot upload." }
    },
    @{
        Level   = "Warning"
        Test    = { param($t) [string]::IsNullOrWhiteSpace([string]$t.OSVersion) }
        Message = { "OS Version not set. Continuing will open a selection menu." }
    },
    @{
        Level   = "Warning"
        Test    = { param($t) [string]::IsNullOrWhiteSpace([string]$t.OSEdition) }
        Message = { "OS Edition not set. Continuing will open a selection menu." }
    },
    @{
        Level   = "Warning"
        Test    = { param($t) [string]::IsNullOrWhiteSpace([string]$t.OSBuild) }
        Message = { "OS Build not set. Continuing will open a selection menu." }
    },
    @{
        Level   = "Warning"
        Test    = { param($t) [string]::IsNullOrWhiteSpace([string]$t.OSLanguage) }
        Message = { "OS Language not set. Continuing will open a selection menu." }
    },
    @{
        Level   = "Warning"
        Test    = { param($t) [string]::IsNullOrWhiteSpace([string]$t.OSActivation) }
        Message = { "OS Activation not set. Continuing will open a selection menu." }
    }
)

function Get-ValidationResults {
    param($Tenant)

    $results = @()

    foreach ($rule in $ValidationRules) {
        if (& $rule.Test $Tenant) {
            $results += [PSCustomObject]@{
                Level   = $rule.Level
                Message = & $rule.Message
            }
        }
    }

    if ($results.Count -eq 0) {
        $results += [PSCustomObject]@{
            Level   = "Success"
            Message = "No issues detected."
        }
    }

    return $results
}

# UI
$form = New-Object System.Windows.Forms.Form
$form.Text = "Tenant Selector"
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'FixedDialog'
$form.ControlBox = $false
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None

# DPI-aware scaling helpers
$g = $form.CreateGraphics()
$script:UiScale = [Math]::Max(1.0, [Math]::Min(2.0, ($g.DpiX / 96.0)))
$g.Dispose()

function S([double]$v) {
    return [int][Math]::Round($v * $script:UiScale)
}

function SFP([double]$pt) {
    return [float]$pt
}

$form.Font = New-Object System.Drawing.Font("Segoe UI Light", (SFP 9.5))
$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea

# Base design size scaled by DPI, capped to screen
$targetWidth  = [int][Math]::Min((S 700), ($wa.Width * 0.92))
$targetHeight = [int][Math]::Min((S 600), ($wa.Height * 0.90))

$form.MinimumSize = New-Object System.Drawing.Size((S 700), (S 600))
$form.Size = New-Object System.Drawing.Size($targetWidth, $targetHeight)

# Search
$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "Search"
$searchLabel.Location = New-Object System.Drawing.Point((S 10), (S 10))
$searchLabel.Size = New-Object System.Drawing.Size((S 120), (S 26))
$searchLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", (SFP 10))
$searchLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Width = S 250
$searchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$searchBox.Font = $form.Font
$searchBox.Height = $searchBox.PreferredHeight
$searchBox.Location = New-Object System.Drawing.Point((S 10), ($searchLabel.Bottom + (S 2)))

# Tenant list
$list = New-Object System.Windows.Forms.ListBox
$list.Location = New-Object System.Drawing.Point((S 10), (S 60))
$list.Size = New-Object System.Drawing.Size((S 250), (S 400))
$list.DisplayMember = "Name"
$list.IntegralHeight = $false
$list.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
$list.ItemHeight = S 22
$list.Font = New-Object System.Drawing.Font("Segoe UI Light", (SFP 9.5))

# Colors
$RowBackA = [System.Drawing.Color]::White
$RowBackB = [System.Drawing.Color]::FromArgb(225, 228, 232)
$SelBack  = [System.Drawing.Color]::FromArgb(0, 120, 212)
$SelFore  = [System.Drawing.Color]::White
$RowFore  = [System.Drawing.Color]::Black

$list.Add_DrawItem({
    param($src, $e)

    if ($e.Index -lt 0) { return }

    $item = $src.Items[$e.Index]
    $text = if ($null -ne $item) { [string]$item.Name } else { "" }
    $text = if ([bool]$item.Pinned) { "★ " + [string]$item.Name } else { [string]$item.Name }

    $isSelected = (($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -ne 0)

    if ($isSelected) {
        $bg = $SelBack
        $fg = $SelFore
    }
    else {
        $bg = if (($e.Index % 2) -eq 0) { $RowBackA } else { $RowBackB }
        $fg = $RowFore
    }

    $bgBrush = New-Object System.Drawing.SolidBrush($bg)
    $e.Graphics.FillRectangle($bgBrush, $e.Bounds)
    $bgBrush.Dispose()

    $flags = [System.Windows.Forms.TextFormatFlags]::Left -bor
             [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor
             [System.Windows.Forms.TextFormatFlags]::EndEllipsis -bor
             [System.Windows.Forms.TextFormatFlags]::NoPrefix

    $x = [int]$e.Bounds.X + (S 6)
    $y = [int]$e.Bounds.Y
    $w = [int]$e.Bounds.Width - (S 8)
    $h = [int]$e.Bounds.Height

    $textRect = New-Object System.Drawing.Rectangle($x, $y, $w, $h)

    [System.Windows.Forms.TextRenderer]::DrawText(
        $e.Graphics,
        $text,
        $e.Font,
        $textRect,
        $fg,
        $flags
    )

    $e.DrawFocusRectangle()
})

# Right headers
$detailsHeader = New-Object System.Windows.Forms.Label
$detailsHeader.Text = "Provisioning details"
$detailsHeader.Location = New-Object System.Drawing.Point((S 270), (S 60))
$detailsHeader.Size = New-Object System.Drawing.Size((S 400), (S 26))
$detailsHeader.Font = New-Object System.Drawing.Font("Segoe UI Semibold", (SFP 10))
$detailsHeader.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$detailsHeader.AutoEllipsis = $true

$validationHeader = New-Object System.Windows.Forms.Label
$validationHeader.Text = "Validation"
$validationHeader.Location = New-Object System.Drawing.Point((S 270), (S 270))
$validationHeader.Size = New-Object System.Drawing.Size((S 400), (S 26))
$validationHeader.Font = New-Object System.Drawing.Font("Segoe UI Semibold", (SFP 10))
$validationHeader.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$validationHeader.AutoEllipsis = $true

# Provisioning details panel
$details = New-Object System.Windows.Forms.TextBox
$details.Location = New-Object System.Drawing.Point((S 270), (S 80))
$details.Size = New-Object System.Drawing.Size((S 400), (S 170))
$details.Font = New-Object System.Drawing.Font("Consolas", (SFP 9.5))
$details.Multiline = $true
$details.ReadOnly = $true
$details.ScrollBars = 'Vertical'
$details.WordWrap = $false
$details.Text = "Select a tenant from the list."

# Validation panel
$validation = New-Object System.Windows.Forms.TextBox
$validation.Location = New-Object System.Drawing.Point((S 270), (S 290))
$validation.Size = New-Object System.Drawing.Size((S 400), (S 80))
$validation.Font = New-Object System.Drawing.Font("Consolas", (SFP 9.5))
$validation.Multiline = $true
$validation.ReadOnly = $true
$validation.ScrollBars = 'Vertical'
$validation.WordWrap = $false

# Start Button
$start = New-Object System.Windows.Forms.Button
$start.Text = "Start"
$start.Font = New-Object System.Drawing.Font("Segoe UI Semibold", (SFP 11), [System.Drawing.FontStyle]::Regular)
$start.Size = New-Object System.Drawing.Size((S 140), (S 40))
$start.Enabled = $false
$start.FlatStyle = 'Flat'
$start.FlatAppearance.BorderSize = 0
$start.UseVisualStyleBackColor = $false
$start.Cursor = [System.Windows.Forms.Cursors]::Hand

$colorNormal   = [System.Drawing.Color]::FromArgb(22,102,172)
$colorHover    = [System.Drawing.Color]::FromArgb(32,122,192)
$colorPressed  = [System.Drawing.Color]::FromArgb(14,82,142)

$start.BackColor = $colorNormal
$start.ForeColor = [System.Drawing.Color]::White

# Brand label (top-right)
$brandLabel = New-Object System.Windows.Forms.Label
$brandLabel.Text = "Multi-Tenant Provisioner"
$brandLabel.AutoSize = $true
$brandLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", (SFP 13), [System.Drawing.FontStyle]::Regular)
$brandLabel.ForeColor = $colorNormal
$brandLabel.BackColor = [System.Drawing.Color]::Transparent
$brandLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

function Set-BrandLabelPosition {
    $paddingRight = S 14
    $paddingTop   = S 10
    $brandLabel.Location = New-Object System.Drawing.Point(
        ([int]($form.ClientSize.Width - $brandLabel.Width - $paddingRight)),
        $paddingTop
    )
}

# Rounded corners
$start.Add_Paint({
    $radius = S 10
    $rect = New-Object System.Drawing.Rectangle(0,0,$start.Width,$start.Height)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc($rect.X, $rect.Y, $radius, $radius, 180, 90)
    $path.AddArc($rect.Right - $radius, $rect.Y, $radius, $radius, 270, 90)
    $path.AddArc($rect.Right - $radius, $rect.Bottom - $radius, $radius, $radius, 0, 90)
    $path.AddArc($rect.X, $rect.Bottom - $radius, $radius, $radius, 90, 90)
    $path.CloseFigure()
    $start.Region = New-Object System.Drawing.Region($path)
})

# Hover → lighter blue
$start.Add_MouseEnter({
    if ($start.Enabled) {
        $start.BackColor = $colorHover
    }
})

# Leave → normal blue
$start.Add_MouseLeave({
    if ($start.Enabled) {
        $start.BackColor = $colorNormal
    }
})

# Press → darker blue
$start.Add_MouseDown({
    if ($start.Enabled) {
        $start.BackColor = $colorPressed
    }
})

# Release → go back to hover or normal
$start.Add_MouseUp({
    if ($start.Enabled) {
        if ($start.ClientRectangle.Contains($start.PointToClient([System.Windows.Forms.Cursor]::Position))) {
            $start.BackColor = $colorHover
        } else {
            $start.BackColor = $colorNormal
        }
    }
})

function Update-Layout {
    $margin = S 10
    $gap = S 10
    $headerHeight = S 26
    $topSearchLabelY = S 10
    $contentTopY = S 60
    $bottomPadding = S 16
    $sectionGap = S 12
    $headerToBoxGap = S 8

    # Left column width scales slightly but stays sensible
    $leftWidth = [int][Math]::Max((S 250), [Math]::Min((S 320), ($form.ClientSize.Width * 0.36)))
    $rightX = $margin + $leftWidth + $gap
    $rightWidth = [int]($form.ClientSize.Width - $rightX - $margin)

    # Search
    $searchLabel.Location = New-Object System.Drawing.Point($margin, $topSearchLabelY)
    $searchLabel.Size = New-Object System.Drawing.Size((S 120), (S 26))

    $searchBox.Location = New-Object System.Drawing.Point($margin, ($searchLabel.Bottom + (S 2)))
    $searchBox.Width = $leftWidth
    $searchBox.Height = $searchBox.PreferredHeight
    $contentTopY = $searchBox.Bottom + (S 8)
    
    # Tenant list fills available height
    $list.Location = New-Object System.Drawing.Point($margin, $contentTopY)
    $list.Size = New-Object System.Drawing.Size($leftWidth, ($form.ClientSize.Height - $contentTopY - $bottomPadding))

    # Right pane total vertical area
    $rightTop = $contentTopY
    $rightBottom = $form.ClientSize.Height - $bottomPadding

    # Reserve a larger bottom action area so the button has breathing room
    $buttonAreaHeight = [int][Math]::Max((S 120), ($form.ClientSize.Height * 0.28))

    # Available height for details + validation + headers
    $availableTextArea = $rightBottom - $rightTop - $buttonAreaHeight - ($headerHeight * 2) - ($headerToBoxGap * 2) - $sectionGap

    if ($availableTextArea -lt (S 300)) { $availableTextArea = S 300 }

    # Better split
    $detailsHeight = [int]($availableTextArea * 0.62)
    $validationHeight = [int]($availableTextArea - $detailsHeight)

    # Details header + box
    $detailsHeader.Location = New-Object System.Drawing.Point($rightX, $rightTop)
    $detailsHeader.Size = New-Object System.Drawing.Size($rightWidth, $headerHeight)

    $details.Location = New-Object System.Drawing.Point($rightX, ($detailsHeader.Bottom + $headerToBoxGap))
    $details.Size = New-Object System.Drawing.Size($rightWidth, $detailsHeight)

    # Validation header + box
    $validationHeader.Location = New-Object System.Drawing.Point($rightX, ($details.Bottom + $sectionGap))
    $validationHeader.Size = New-Object System.Drawing.Size($rightWidth, $headerHeight)

    $validation.Location = New-Object System.Drawing.Point($rightX, ($validationHeader.Bottom + $headerToBoxGap))
    $validation.Size = New-Object System.Drawing.Size($rightWidth, $validationHeight)

    # Start button centered in the remaining space below Validation
    $startX = $validation.Left + [int](($validation.Width - $start.Width) / 2)
    $spaceBelowValidation = $rightBottom - $validation.Bottom
    $startY = $validation.Bottom + [int](($spaceBelowValidation - $start.Height) / 2)
    $start.Location = New-Object System.Drawing.Point($startX, $startY)

    # Brand top-right
    Set-BrandLabelPosition
}

# Add controls
$form.Controls.AddRange(@(
    $brandLabel,
    $searchLabel, $searchBox,
    $list,
    $detailsHeader, $details,
    $validationHeader, $validation,
    $start
))

# Apply layout once after controls are added
Update-Layout

# Re-apply after show / resize + focus search box on startup
$form.Add_Shown({
    Update-Layout
    $form.ActiveControl = $searchBox
    $searchBox.Focus() | Out-Null
    $searchBox.SelectionStart = $searchBox.TextLength
    $searchBox.SelectionLength = 0
})
$form.Add_Resize({ Update-Layout })

# Populate list helper
function Update-TenantList {
    param([string]$Filter)

    $list.BeginUpdate()
    $list.Items.Clear()

    $filtered = if ([string]::IsNullOrWhiteSpace([string]$Filter)) {
        $AllTenants
    }
    else {
        $AllTenants | Where-Object {
            $_.Name -and $_.Name.IndexOf($Filter, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
        }
    }

    foreach ($t in $filtered) { [void]$list.Items.Add($t) }
    $list.EndUpdate()

    $list.ClearSelected()
    $details.Text = "Select a tenant from the list."
    $validation.Text = ""
    $start.Enabled = $false
}

Update-TenantList ""

# Events
$searchBox.Add_TextChanged({
    Update-TenantList $searchBox.Text.Trim()
})

$list.Add_SelectedIndexChanged({
    $t = $list.SelectedItem
    if (-not $t) { return }

$script:SetupCompleteDisplay = if (
    ($t.PSObject.Properties.Name -contains 'SetupCompleteUrl') -and
    (-not [string]::IsNullOrWhiteSpace([string]$t.SetupCompleteUrl))
) {
    "Tenant Specific"
}
elseif (-not [string]::IsNullOrWhiteSpace([string]$SetupCompleteUrl)) {
    "Default"
}
else {
    "Not used"
}

    $details.Text = @"
Name           : $($t.Name)
TenantId       : $($t.TenantId)
Autopilot      : $($t.UploadToAutopilot)
GroupTag       : $($t.GroupTag)
PreOOBEScript  : $script:SetupCompleteDisplay

OSVersion      : $($t.OSVersion)
OSEdition      : $($t.OSEdition)
OSBuild        : $($t.OSBuild)
OSLanguage     : $($t.OSLanguage)
OSActivation   : $($t.OSActivation)
"@

    $results = Get-ValidationResults $t

    $validation.Text = ($results | ForEach-Object {
        "[$($_.Level)] $($_.Message)"
    }) -join "`r`n"

    # Block Start on Error
    $levels = @($results | ForEach-Object { $_.Level })
    $start.Enabled = -not ($levels -contains "Error")
})

$start.Add_Click({
    $t = $list.SelectedItem
    if (-not $t) { return }

    $script:TenantId          = $t.TenantId
    $script:TenantName        = $t.Name
    $script:OSVersion         = $t.OSVersion
    $script:OSEdition         = $t.OSEdition
    $script:OSBuild           = $t.OSBuild
    $script:OSLanguage        = $t.OSLanguage
    $script:OSActivation      = $t.OSActivation
    $script:GroupTag          = $t.GroupTag
    $script:UploadToAutopilot = [bool]$t.UploadToAutopilot
    if ( ($t.PSObject.Properties.Name -contains 'SetupCompleteUrl') -and (-not [string]::IsNullOrWhiteSpace([string]$t.SetupCompleteUrl))) {$script:SetupCompleteUrl = [string]$t.SetupCompleteUrl}

    $form.Close()
})

# Run
Write-Separator
Write-Log "Starting Tenant Selector..." Cyan
[void]$form.ShowDialog()
Write-Log "Tenant selected continuing..." Green
Write-Separator
#endregion

#region: Key Vault
if ([string]::IsNullOrWhiteSpace($AppSecret) -and $UploadToAutopilot -eq $true) {

Write-Separator
try {
    # Get token for Key Vault
    Write-Log "Requesting Key Vault token..." DarkGray

    $tokenResponse = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$SPNTenantID/oauth2/v2.0/token" `
        -Body @{
            client_id     = $SPNAppID
            client_secret = $SPNSecret
            scope         = "https://vault.azure.net/.default"
            grant_type    = "client_credentials"
        } `
        -ContentType "application/x-www-form-urlencoded"

    $accessToken = $tokenResponse.access_token

    $uri = "https://$VaultName.vault.azure.net/secrets/$SecretName/?api-version=7.3"

    Write-Log "Calling Key Vault from public IP: $publicIP" DarkGray

    # Call Key Vault
    $response = Invoke-RestMethod -Method Get `
        -Uri $uri `
        -Headers @{
            Authorization = "Bearer $accessToken"
        }

    Write-Log "Secret retrieved successfully" Green
    $AppSecret = $response.value
}
catch {
    Write-Log "Key Vault error:" Red
    Write-Log $_ Red

    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $body = $reader.ReadToEnd()
        Write-Log "Response body:" DarkRed
        Write-Log $body DarkRed
    }

    throw
}

}
#endregion

#region: Autopilot Upload
Write-Separator
    if ($UploadToAutopilot)
    {
# Functions for Autopilot API operations
function Get-AuthToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] [String] $TenantId,
        [Parameter(Mandatory=$true)] [String] $AppId,
        [Parameter(Mandatory=$true)] [String] $AppSecret
    )
    try {
        # Define auth body
        $body = @{
            grant_type    = "client_credentials"
            client_id     = $AppId
            client_secret = $AppSecret
            scope         = "https://graph.microsoft.com/.default"
        }

        # Get OAuth token
        $response = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body
        
        # Return the token
        return $response.access_token
    }
    catch {
        Write-Log "Error getting auth token: $_" Red
        if ($_.Exception.Response) {
            $errorResponse = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Log $responseBody Red
        }
        throw
    }
}

function Add-AutopilotImportedDevice {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] [String] $SerialNumber,
        [Parameter(Mandatory=$true)] [String] $HardwareHash,
        [Parameter(Mandatory=$false)] [String] $GroupTag = "",
        [Parameter(Mandatory=$true)] [String] $AuthToken
    )

    try {
        # Create the device object
        $deviceObject = @{
            serialNumber = $SerialNumber
            hardwareIdentifier = $HardwareHash
        }

        # Add GroupTag if specified
        if (-not [string]::IsNullOrEmpty($GroupTag)) {
            $deviceObject.groupTag = $GroupTag
        }

        # Convert to JSON
        $deviceJson = $deviceObject | ConvertTo-Json

        # Set up API request
        $headers = @{
            "Authorization" = "Bearer $AuthToken"
            "Content-Type" = "application/json"
        }

        # Upload to Autopilot using the importedWindowsAutopilotDeviceIdentities endpoint
        Write-Log "Uploading device to Autopilot..." DarkGray
        $response = Invoke-RestMethod -Method Post `
            -Uri "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities" `
            -Headers $headers `
            -Body $deviceJson

        return $response
    }
    catch {
        Write-Log "Error adding device to Autopilot: $_" Red
        if ($_.Exception.Response) {
            $errorResponse = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Log $responseBody Red
        }
        throw
    }
}

function Get-AutopilotImportedDevice {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] [String] $Id,
        [Parameter(Mandatory=$true)] [String] $AuthToken
    )

    try {
        # Set up API request
        $headers = @{
            "Authorization" = "Bearer $AuthToken"
            "Content-Type" = "application/json"
        }

        # Get device status from Autopilot
        $response = Invoke-RestMethod -Method Get `
            -Uri "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities/$Id" `
            -Headers $headers

        return $response
    }
    catch {
        Write-Log "Error getting device status: $_" Red
        if ($_.Exception.Response) {
            $errorResponse = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Log $responseBody Red
        }
        throw
    }
}

function Get-AutopilotDevice {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [String] $Serial,
        [Parameter(Mandatory = $true)] [String] $AuthToken
    )

    try {
        # Set up API request
        $headers = @{
            "Authorization" = "Bearer $AuthToken"
            "Content-Type"  = "application/json"
        }

        # Get device status from Autopilot
        $response = Invoke-RestMethod -Method Get `
            -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,%27$serial%27)" `
            -Headers $headers
        $device = $response.value | Where-Object { $_.serialNumber -eq $serial }
        return $device
    }
    catch {
        Write-Log "Error getting device status: $_" Red
        if ($_.Exception.Response) {
            $errorResponse = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Log $responseBody Red
        }
        throw
    }
}

# Check if we're in WinPE and have the required PCPKsp.dll file
If ((Test-Path X:\Windows\System32\wpeutil.exe) -and (Test-Path $PSScriptRoot\PCPKsp.dll))
{
    Write-Log "Running in WinPE, installing PCPKsp.dll for TPM support..." DarkGray
    Copy-Item "$PSScriptRoot\PCPKsp.dll" "X:\Windows\System32\PCPKsp.dll"
    # Register PCPKsp
    rundll32 X:\Windows\System32\PCPKsp.dll,DllInstall
}

# Change Current Directory so OA3Tool finds the files written in the Config File 
Push-Location $PSScriptRoot

# Delete old Files if exits
if (Test-Path $PSScriptRoot\OA3.xml) 
{
    Remove-Item $PSScriptRoot\OA3.xml -Force
}

# Get SN from WMI
$serial = (Get-WmiObject -Class Win32_BIOS).SerialNumber
Write-Log "Device Serial Number: $serial" DarkGray

# Run OA3Tool
Write-Log "Running OA3Tool to gather hardware hash..." DarkGray
&$PSScriptRoot\oa3tool.exe /Report /ConfigFile=$PSScriptRoot\OA3.cfg /NoKeyCheck

# Check if Hash was found
If (Test-Path $PSScriptRoot\OA3.xml)
{
    # Read Hash from generated XML File
    [xml]$xmlhash = Get-Content -Path "$PSScriptRoot\OA3.xml"
    $hash = $xmlhash.Key.HardwareHash
    Write-Log "Hardware Hash successfully retrieved" Green

    # Delete XML File
    Remove-Item $PSScriptRoot\OA3.xml -Force

    # Output the hash information to screen
    Write-Log "Serial Number: $serial" DarkGray
    Write-Log "Group Tag: $GroupTag" DarkGray
    Write-Log "Hardware Hash length: $(($hash).Length) characters" DarkGray

    # Create temporary CSV file in case it's needed
    $TempCSVPath = "X:\Windows\Temp\AutopilotHash.csv"

    # Create the CSV object
    $computers = @()
    $product = ""

    if ($GroupTag -ne "")
    {
        # Create a pipeline object with Group Tag
        $c = New-Object psobject -Property @{
            "Device Serial Number" = $serial
            "Windows Product ID"   = $product
            "Hardware Hash"        = $hash
            "Group Tag"            = $GroupTag
        }

        # Save to temp CSV
        $computers += $c
        $computers |
            Select-Object "Device Serial Number", "Windows Product ID", "Hardware Hash", "Group Tag" |
            ConvertTo-Csv -NoTypeInformation |
            ForEach-Object { $_ -replace '"','' } |
            Out-File $TempCSVPath
    }
    else
    {
        # Create a pipeline object without Group Tag
        $c = New-Object psobject -Property @{
            "Device Serial Number" = $serial
            "Windows Product ID"   = $product
            "Hardware Hash"        = $hash
        }

        # Save to temp CSV
        $computers += $c
        $computers |
            Select-Object "Device Serial Number", "Windows Product ID", "Hardware Hash" |
            ConvertTo-Csv -NoTypeInformation |
            ForEach-Object { $_ -replace '"','' } |
            Out-File $TempCSVPath
    }

    Write-Log "CSV file created at: $TempCSVPath" DarkGray

    # Upload to Autopilot if requested
    $authToken = $null
    $authSucceeded = $false
    $uploadSucceeded = $false
    $deviceAlreadyExists = $false


        if (-not $TenantId -or -not $AppId -or -not $AppSecret)
        {
            Write-Log "Error: TenantId, AppId, and AppSecret parameters are required for Autopilot upload" Red
        }
        else
        {
            try
            {
                # Get auth token
                Write-Log "Getting authorization token..." DarkGray
                $authToken = Get-AuthToken -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret
                $authSucceeded = $true

                # Check if device already exists in Autopilot before upload
                $existingDevice = Get-AutopilotDevice -Serial $serial -AuthToken $authToken
                if ($existingDevice)
                {
                    $deviceAlreadyExists = $true
                    Write-Log "Device already exists in Autopilot with SerialNumber: $serial" Green
                    Write-Log "Skipping upload and profile assignment check." DarkGray
                }
                else
                {
                    # Upload device to Autopilot
                    Write-Log "Adding device to Autopilot..." DarkGray
                    $importedDevice = Add-AutopilotImportedDevice -SerialNumber $serial -HardwareHash $hash -GroupTag $GroupTag -AuthToken $authToken

                    if ($importedDevice)
                    {
                        Write-Log "Device added successfully with ID: $($importedDevice.id)" Green

                        # Wait for import processing to complete
                        Write-Log "Waiting for import to complete..." DarkGray
                        $processingComplete = $false
                        $maxRetries       = 20
                        $retryCount       = 0
                        $lastStatus       = $null
                        $pollInterval     = 15
                        $progressActivity = "Waiting for Autopilot import"

                        while (-not $processingComplete -and $retryCount -lt $maxRetries)
                        {
                            # Countdown bar between polls
                            for ($i = $pollInterval; $i -gt 0; $i--)
                            {
                                $statusText = if ($lastStatus) { "Status: $lastStatus. Next check in $i seconds..." } else { "Next check in $i seconds..." }
                                Write-Progress -Activity $progressActivity `
                                               -Status $statusText `
                                               -PercentComplete (($retryCount / $maxRetries) * 100)
                                Start-Sleep -Seconds 1
                            }

                            $device = Get-AutopilotImportedDevice -Id $importedDevice.id -AuthToken $authToken
                            $status = $device.state.deviceImportStatus

                            if ($status -eq "complete")
                            {
                                Write-Progress -Activity $progressActivity -Completed
                                $processingComplete = $true
                                $uploadSucceeded    = $true
                                Write-Log "Import completed successfully!" Green
                                Write-Log "Device Registration ID: $($device.state.deviceRegistrationId)" DarkGray
                                break
                            }
                            elseif ($status -eq "error")
                            {
                                Write-Progress -Activity $progressActivity -Completed
                                Write-Log "Import failed with error: $($device.state.deviceErrorCode) - $($device.state.deviceErrorName)" Red
                                break
                            }

                            # Output only when status changes
                            if ($status -ne $lastStatus)
                            {
                                Write-Log "Import status: $status" DarkGray
                                $lastStatus = $status
                            }

                            $retryCount++
                        }

                        Write-Progress -Activity $progressActivity -Completed

                        if (-not $processingComplete)
                        {
                            Write-Log "[WARNING]: Import did not complete within the expected time" Yellow
                        }
                    }
                    else
                    {
                        Write-Log "[WARNING]: Import request did not return a device object" Yellow
                    }
                }
            }
            catch
            {
                Write-Log "An error occurred during the Autopilot upload process: $_" Red
            }
        }
    }
    else
    {
        Write-Log "[WARNING]: Skipping Autopilot upload. Use -UploadToAutopilot switch with required parameters to upload" Yellow
    }

    # Only check profile assignment if:
    # 1) authentication succeeded
    # 2) the device was uploaded successfully in THIS run
    # 3) the device did not already exist before upload
    if ($authSucceeded -and $uploadSucceeded -and -not $deviceAlreadyExists)
    {
        Write-Log "Waiting for Autopilot profile assignment..." DarkGray
        $profileAssigned  = $false
        $maxRetries       = 50
        $retryCount       = 0
        $lastStatus       = $null
        $pollInterval     = 15
        $progressActivity = "Waiting for Autopilot profile assignment"

        while (-not $profileAssigned -and $retryCount -lt $maxRetries)
        {
            # Countdown bar between polls
            for ($i = $pollInterval; $i -gt 0; $i--)
            {
                $statusText = if ($lastStatus) { "Status: $lastStatus. Next check in $i seconds..." } else { "Next check in $i seconds..." }
                Write-Progress -Activity $progressActivity `
                               -Status $statusText `
                               -PercentComplete (($retryCount / $maxRetries) * 100)
                Start-Sleep -Seconds 1
            }

            $apDevice = Get-AutopilotDevice -Serial $serial -AuthToken $authToken
            $status   = if ($apDevice) { $apDevice.deploymentProfileAssignmentStatus } else { 'notVisible' }

            if ($status -in @("assignedUnkownSyncState","assignedInSync","assignedOutOfSync"))
            {
                Write-Progress -Activity $progressActivity -Completed
                Write-Log "Autopilot profile assigned successfully!" Green
                $profileAssigned = $true
                break
            }

            # Output only when status changes
            if ($status -ne $lastStatus)
            {
                $message = switch ($status) {
                    'notVisible' { "Device not yet visible in Autopilot device list" }
                    'pending'    { "Profile assignment in progress" }
                    default      { "Profile status: $status" }
                }
                Write-Log $message DarkGray
                $lastStatus = $status
            }

            $retryCount++
        }

        Write-Progress -Activity $progressActivity -Completed

        if (-not $profileAssigned)
        {
            Write-Log "[WARNING]: Autopilot profile was not assigned within expected time." Yellow
        }
    }
    else
    {
        Write-Log "[WARNING]: Skipping Autopilot profile assignment check" Yellow

        if ($deviceAlreadyExists)
        {
            Write-Plain "Reason: device was already present in Autopilot." Green
        }
        elseif (-not $authSucceeded)
        {
            Write-Plain "Reason: authentication did not succeed." Red
        }
        elseif (-not $uploadSucceeded)
        {
            Write-Plain "Reason: upload/import did not complete successfully in this run." Red
        }
    }
}
else
{
    Write-Log "Autopilot upload is disabled." DarkGray
}
#endregion

#region: Download SetupComplete.ps1
if (-not [string]::IsNullOrWhiteSpace($SetupCompleteUrl)) {

    $DestinationFolder = "X:\OSDCloud\Config\Scripts\SetupComplete"

    try {
        # Ensure TLS 1.2+
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        # Extract filename and construct destination path
        $FileName        = Split-Path -Path $SetupCompleteUrl -Leaf
        $DestinationFile = Join-Path -Path $DestinationFolder -ChildPath $FileName

        Write-Separator
        Write-Log "Downloading $FileName from public IP: $publicIP" DarkGray

        # Hardcode the filename to SetupComplete.ps1 to ensure OSDCloud picks it up, regardless of the original name
        Invoke-WebRequest -Uri $SetupCompleteUrl -OutFile "$DestinationFolder\SetupComplete.ps1" -UseBasicParsing -ErrorAction Stop

        Write-Log "$FileName downloaded successfully" Green
    }
    catch {
        Write-Log "Failed to download $FileName $($_.Exception.Message)" Red
    }   

    }
else {
Write-Log "SetupCompleteUrl not provided. Skipping SetupComplete.ps1 download." DarkGray
}
Write-Separator
#endregion

#region: Start OSDCloud
# Output provisioning details before starting OSDCloud
Write-Plain "Provisioning Details" DarkCyan
Write-Separator -Char '-'

$ProvisioningDetails = [ordered]@{
    Tenant          = [string]$TenantName
    OSVersion       = [string]$OSVersion
    OSEdition       = [string]$OSEdition
    OSBuild         = [string]$OSBuild
    OSLanguage      = [string]$OSLanguage
    OSActivation    = [string]$OSActivation
    PreOOBEScript   = if ([string]::IsNullOrWhiteSpace([string]$script:SetupCompleteDisplay)) { "Not Used" } else { [string]$script:SetupCompleteDisplay }
    GroupTag        = if ([string]::IsNullOrWhiteSpace([string]$GroupTag)) { "Not Used" } else { [string]$GroupTag }
}

foreach ($item in $ProvisioningDetails.GetEnumerator()) {
    $value = if ([string]::IsNullOrWhiteSpace([string]$item.Value)) {
        "Not Set"
    }
    else {
        ([string]$item.Value).Replace("`r", " ").Replace("`n", " ")
    }

    $line = ("{0,-14}: {1}" -f $item.Key, $value)
    Write-Plain $line Yellow
}

# Autopilot status separately
$AutopilotStatus =
    if (-not $UploadToAutopilot) {
        "Not used"
    }
    elseif ($deviceAlreadyExists) {
        "Already uploaded"
    }
    elseif (-not $authSucceeded) {
        "Authentication failed"
    }
    elseif (-not $uploadSucceeded) {
        "Upload failed"
    }
    elseif (-not $profileAssigned) {
        "Profile not assigned in allowed time"
    }
    else {
        "Upload succeeded"
    }

$AutopilotLine = ("{0,-14}: {1}" -f "Autopilot", $AutopilotStatus)

if (-not $UploadToAutopilot) {
    Write-Plain $AutopilotLine Yellow
}
elseif ($deviceAlreadyExists) {
    Write-Plain $AutopilotLine Green
}
elseif (-not $authSucceeded) {
    Write-Plain $AutopilotLine Red
}
elseif (-not $uploadSucceeded) {
    Write-Plain $AutopilotLine Red
}
elseif (-not $profileAssigned) {
    Write-Plain $AutopilotLine Red
}
else {
    Write-Plain $AutopilotLine Green
}
Write-Separator -Char '-'

# Manual confirm to continue after Autopilot issue
if (
    $UploadToAutopilot -and
    -not $deviceAlreadyExists -and
        (
            -not $authSucceeded -or
            -not $uploadSucceeded -or
            -not $profileAssigned
        )
    ) 
    {
        do {
            $continueChoice = Read-Host "Autopilot encountered an error. Continue provisioning without Autopilot anyway? (Y/N)"
        } until ($continueChoice -match '^[YyNn]$')

        if ($continueChoice -match '^[Nn]$') {
            Write-Plain "Chose not to continue. Restarting computer in 10 seconds..." DarkGray
            Start-Sleep -Seconds 10
            Restart-Computer
            Start-Sleep -Seconds 10
        }
    }

# Setting parameters from selected tenant for Start-OSDCloud.
$StartOSDCloudParams = @{
    OSEdition      = $OSEdition
    OSLanguage     = $OSLanguage
    OSVersion      = $OSVersion
    OSBuild        = $OSBuild    
    OSActivation   = $OSActivation
    ZTI            = $true
    Restart        = $true
    SkipAutopilot  = $true
}
    
    try {
        # Show end/runtime lines right before handoff
        $endTime = Get-Date
        $duration = $endTime - $startTime 

        # Duration
        $mins = [int]$duration.TotalMinutes
        $secs = [int]$duration.Seconds
        $durationText = ("{0:00} minutes {1:00} seconds" -f $mins, $secs)

        Write-Separator
        Write-Plain "Multi-Tenant Provisioner Finished" Cyan
        Write-Plain "Completed in $durationText" DarkGray

        Write-Separator
        Write-Log "Starting OSDCloud in 10 seconds..." Cyan
        Start-Sleep -Seconds 10

        # Stop MTP transcript before handing off to OSDCloud / OSDCloudGUI
        if ($script:TranscriptActive) {
            Write-Log "Further actions are logged by OSDCloud itself." DarkGray
            Stop-Transcript
            $script:TranscriptActive = $false
        }
        
        Update-Module -Name OSD -ErrorAction Ignore
        Start-OSDCloud @StartOSDCloudParams -ErrorAction Stop

    }
    catch {
        
        Write-Log "Start-OSDCloud failed: $($_.Exception.Message)" Red
        Write-Log "Falling back to Start-OSDCloudGUI..." DarkGray

        try {
            Start-OSDCloudGUI -ErrorAction Stop     
        }
        catch {

        Write-Log "Start-OSDCloudGUI also failed: $($_.Exception.Message)" Red
        throw

        }
}
#endregion

#region: Stop Logging
if ($script:TranscriptActive) {
    Stop-Transcript
    $script:TranscriptActive = $false
}
#endregion