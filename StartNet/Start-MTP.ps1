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

#region: Start Logging
# Start transcript
$Global:Transcript = "$((Get-Date).ToString('dd-MM-yyyy-HHmmss'))-Start-MTP.log"
Start-Transcript -Path (Join-Path "X:\OSDCloud\Config\Scripts\SetupComplete" $Global:Transcript) -ErrorAction Ignore

# Start timer
$startTime = Get-Date
Write-Host "Multi-Tenant Provisioner started at: $($startTime.ToString('dd-MM-yyyy-HHmmss'))" -ForegroundColor Yellow
#endregion

#region: Test internet connectity
if ($SPNSecret -or $SetupCompleteUrl -or $ParametersUrl -or $UploadToAutopilot) {
    Write-Host "Internet connectivity is required. Checking connection..." -ForegroundColor Cyan

    $RetryInterval = 10
    $firstAttempt = $true

    while ($true) {
        try {
            # Test internet connectivity
            Invoke-WebRequest -Uri "https://www.microsoft.com/" -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop | Out-Null
            Write-Host "Internet connection detected. Continuing script..." -ForegroundColor Green
            break
        }
        catch {
            if ($firstAttempt) {
                Write-Host "No internet connection detected." -ForegroundColor Yellow
                Write-Host "Please connect the device to the network." -ForegroundColor Yellow
                $firstAttempt = $false
            }

            # Countdown
            for ($i = $RetryInterval; $i -gt 0; $i--) {
                Write-Host -NoNewline "`rRetrying in $i seconds..."
                Start-Sleep -Seconds 1
            }

            # Overwrite the countdown line with "Retrying now..."
            Write-Host -NoNewline "`rRetrying now...           " # Extra spaces to overwrite previous text
        }
    }
}
else {
    Write-Host "No internet-dependent parameters configured." -ForegroundColor Cyan
    Write-Host "Continuing without internet connectivity. Script functionality may be limited." -ForegroundColor DarkRed
}
# Get public IP address
try {
$publicIP = (Invoke-RestMethod -Uri "https://api.ipify.org?format=json").ip
    } catch {} 
#endregion

#region: Injected Tenant Parameters
# If ParametersUrl is provided, read tenant information from the config file instead of using hardcoded parameters. This allows for dynamic configuration without modifying the script.

if (-not [string]::IsNullOrWhiteSpace($ParametersUrl)) {
    
$DestinationFolder = "X:\OSDCloud\Config\Scripts\SetupComplete"

try {
    # Ensure TLS 1.2
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Extract filename and construct destination path
    $FileName        = Split-Path -Path $ParametersUrl -Leaf
    $DestinationFile = Join-Path -Path $DestinationFolder -ChildPath $FileName

    # Download parameter file
    Write-Host "Downloading $FileName from public IP: $publicIP" -ForegroundColor Cyan

    Invoke-WebRequest -Uri $ParametersUrl -OutFile $DestinationFile -UseBasicParsing -ErrorAction Stop

    Write-Host "$FileName downloaded successfully to $DestinationFolder." -ForegroundColor Green
}
catch {
    Write-Error "Failed to download $FileName : $($_.Exception.Message)" -ForegroundColor Red
}

$config = Get-Content "$DestinationFile" -Raw | ConvertFrom-Json
$Tenants = $config.Tenants

}
else {
    Write-Host "ParametersUrl not provided. Using hardcoded parameters." -ForegroundColor Yellow
}
#endregion

#region: Hardcoded Tenant Parameters
# If ParametersUrl is not provided, use hardcoded tenant information. This is useful if you prefer to have the tenant information directly in the script.

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
    Write-Host "ParametersUrl provided. Skipping hardcoded parameters and using downloaded config file." -ForegroundColor Yellow
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
        if ($_.Name -match 'Ã.|Â.|�') {
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
OOBEScript     : $script:SetupCompleteDisplay

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
Write-Host "Starting Tenant Selector..." -ForegroundColor Cyan
[void]$form.ShowDialog()
Write-Host "Tenant selected continuing script.." -ForegroundColor Green
#endregion

#region: Key Vault
if ([string]::IsNullOrWhiteSpace($AppSecret) -and $UploadToAutopilot -eq $true) {

try {
    # Get token for Key Vault
    Write-Host "Requesting Key Vault token..." -ForegroundColor Yellow

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

    Write-Host "Calling Key Vault from public IP: $publicIP" -ForegroundColor Cyan

    # Call Key Vault
    $response = Invoke-RestMethod -Method Get `
        -Uri $uri `
        -Headers @{
            Authorization = "Bearer $accessToken"
        }

    Write-Host "Secret retrieved successfully." -ForegroundColor Green
    $AppSecret = $response.value
}
catch {
    Write-Host "Key Vault error:" -ForegroundColor Red
    Write-Host $_ -ForegroundColor Red

    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $body = $reader.ReadToEnd()
        Write-Host "Response body:" -ForegroundColor DarkRed
        Write-Host $body -ForegroundColor DarkRed
    }

    throw
}

}
#endregion

#region: Autopilot Upload
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
        Write-Host "Error getting auth token: $_" -ForegroundColor Red
        if ($_.Exception.Response) {
            $errorResponse = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Host $responseBody -ForegroundColor Red
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
        Write-Host "Uploading device to Autopilot..." -ForegroundColor Yellow
        $response = Invoke-RestMethod -Method Post `
            -Uri "https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities" `
            -Headers $headers `
            -Body $deviceJson

        return $response
    }
    catch {
        Write-Host "Error adding device to Autopilot: $_" -ForegroundColor Red
        if ($_.Exception.Response) {
            $errorResponse = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Host $responseBody -ForegroundColor Red
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
        Write-Host "Error getting device status: $_" -ForegroundColor Red
        if ($_.Exception.Response) {
            $errorResponse = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Host $responseBody -ForegroundColor Red
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
        Write-Host "Error getting device status: $_" -ForegroundColor Red
        if ($_.Exception.Response) {
            $errorResponse = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Host $responseBody -ForegroundColor Red
        }
        throw
    }
}

# Check if we're in WinPE and have the required PCPKsp.dll file
If ((Test-Path X:\Windows\System32\wpeutil.exe) -and (Test-Path $PSScriptRoot\PCPKsp.dll))
{
    Write-Host "Running in WinPE, installing PCPKsp.dll for TPM support..." -ForegroundColor Yellow
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
Write-Host "Device Serial Number: $serial" -ForegroundColor Cyan

# Run OA3Tool
Write-Host "Running OA3Tool to gather hardware hash..." -ForegroundColor Green
&$PSScriptRoot\oa3tool.exe /Report /ConfigFile=$PSScriptRoot\OA3.cfg /NoKeyCheck

# Check if Hash was found
If (Test-Path $PSScriptRoot\OA3.xml)
{
    # Read Hash from generated XML File
    [xml]$xmlhash = Get-Content -Path "$PSScriptRoot\OA3.xml"
    $hash = $xmlhash.Key.HardwareHash
    Write-Host "Hardware Hash successfully retrieved" -ForegroundColor Green

    # Delete XML File
    Remove-Item $PSScriptRoot\OA3.xml -Force

    # Output the hash information to screen
    Write-Host "Serial Number: $serial" -ForegroundColor Cyan
    Write-Host "Group Tag: $GroupTag" -ForegroundColor Cyan
    Write-Host "Hardware Hash length: $(($hash).Length) characters" -ForegroundColor Cyan

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

    Write-Host "CSV file created at: $TempCSVPath" -ForegroundColor Green

    # Upload to Autopilot if requested
    $authToken = $null
    $authSucceeded = $false
    $uploadSucceeded = $false
    $deviceAlreadyExists = $false


        if (-not $TenantId -or -not $AppId -or -not $AppSecret)
        {
            Write-Host "Error: TenantId, AppId, and AppSecret parameters are required for Autopilot upload" -ForegroundColor Red
        }
        else
        {
            try
            {
                # Get auth token
                Write-Host "Getting authorization token..." -ForegroundColor Yellow
                $authToken = Get-AuthToken -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret
                $authSucceeded = $true

                # Check if device already exists in Autopilot before upload
                $existingDevice = Get-AutopilotDevice -Serial $serial -AuthToken $authToken
                if ($existingDevice)
                {
                    $deviceAlreadyExists = $true
                    Write-Host "Device already exists in Autopilot with SerialNumber: $serial" -ForegroundColor Green
                    Write-Host "Skipping upload and profile assignment check." -ForegroundColor Yellow
                }
                else
                {
                    # Upload device to Autopilot
                    Write-Host "Adding device to Autopilot..." -ForegroundColor Yellow
                    $importedDevice = Add-AutopilotImportedDevice -SerialNumber $serial -HardwareHash $hash -GroupTag $GroupTag -AuthToken $authToken

                    if ($importedDevice)
                    {
                        Write-Host "Device added successfully with ID: $($importedDevice.id)" -ForegroundColor Green

                        # Wait for import processing to complete
                        Write-Host "Waiting for import to complete..." -ForegroundColor Yellow
                        $processingComplete = $false
                        $maxRetries = 20
                        $retryCount = 0

                        while (-not $processingComplete -and $retryCount -lt $maxRetries)
                        {
                            Start-Sleep -Seconds 15
                            $device = Get-AutopilotImportedDevice -Id $importedDevice.id -AuthToken $authToken

                            if ($device.state.deviceImportStatus -eq "complete")
                            {
                                $processingComplete = $true
                                $uploadSucceeded = $true
                                Write-Host "Import completed successfully!" -ForegroundColor Green
                                Write-Host "Device Registration ID: $($device.state.deviceRegistrationId)" -ForegroundColor Cyan
                            }
                            elseif ($device.state.deviceImportStatus -eq "error")
                            {
                                Write-Host "Import failed with error: $($device.state.deviceErrorCode) - $($device.state.deviceErrorName)" -ForegroundColor Red
                                break
                            }
                            else
                            {
                                Write-Host "Import status: $($device.state.deviceImportStatus). Waiting..." -ForegroundColor Yellow
                                $retryCount++
                            }
                        }

                        if (-not $processingComplete)
                        {
                            Write-Host "Import did not complete within the expected time." -ForegroundColor Yellow
                        }
                    }
                    else
                    {
                        Write-Host "Import request did not return a device object." -ForegroundColor Yellow
                    }
                }
            }
            catch
            {
                Write-Host "An error occurred during the Autopilot upload process: $_" -ForegroundColor Red
            }
        }
    }
    else
    {
        Write-Host "Skipping Autopilot upload. Use -UploadToAutopilot switch with required parameters to upload." -ForegroundColor Yellow
    }

    # Only check profile assignment if:
    # 1) authentication succeeded
    # 2) the device was uploaded successfully in THIS run
    # 3) the device did not already exist before upload
    if ($authSucceeded -and $uploadSucceeded -and -not $deviceAlreadyExists)
    {
        Write-Host "Waiting for Autopilot profile assignment..." -ForegroundColor Yellow
        $profileAssigned = $false
        $maxRetries = 50
        $retryCount = 0

        while (-not $profileAssigned -and $retryCount -lt $maxRetries)
        {
            Start-Sleep -Seconds 15
            $apDevice = Get-AutopilotDevice -Serial $serial -AuthToken $authToken

            if ($apDevice)
            {
                $status = $apDevice.deploymentProfileAssignmentStatus

                if ($status -in @("assignedUnkownSyncState","assignedInSync","assignedOutOfSync"))
                {
                    Write-Host "Autopilot profile assigned successfully!" -ForegroundColor Green                        
                    $profileAssigned = $true
                }
                elseif ($status -eq "pending")
                {
                    Write-Host "Profile assignment in progress..." -ForegroundColor Yellow
                }
                else
                {
                    Write-Host "Profile status: $status - waiting..." -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "Device not yet visible in Autopilot device list..." -ForegroundColor Yellow
            }

            $retryCount++
        }

        if (-not $profileAssigned)
        {
            Write-Host "Autopilot profile was not assigned within expected time." -ForegroundColor Yellow
        }
    }
    else
    {
        Write-Host "Skipping Autopilot profile assignment check." -ForegroundColor Yellow

        if ($deviceAlreadyExists)
        {
            Write-Host "Reason: device was already present in Autopilot." -ForegroundColor Yellow
        }
        elseif (-not $authSucceeded)
        {
            Write-Host "Reason: authentication did not succeed." -ForegroundColor Yellow
        }
        elseif (-not $uploadSucceeded)
        {
            Write-Host "Reason: upload/import did not complete successfully in this run." -ForegroundColor Yellow
        }
    }
}
else
{
    Write-Host "Autopilot upload is disabled." -ForegroundColor Yellow
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

        # Download only if folder exists
            Write-Host "Downloading $FileName from public IP: $publicIP" -ForegroundColor Cyan

            Invoke-WebRequest -Uri $SetupCompleteUrl -OutFile "$DestinationFolder\SetupComplete.ps1" -UseBasicParsing -ErrorAction Stop

            Write-Host "$FileName downloaded successfully to $DestinationFolder." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to download $FileName $($_.Exception.Message)" -ForegroundColor Red
    }   

    }
else {
Write-Host "SetupCompleteUrl not provided. Skipping SetupComplete.ps1 download." -ForegroundColor Yellow
}
#endregion

#region: Start OSDCloud
# Output provisioning details before starting OSDCloud
Write-Host "`nProvisioning Details" -ForegroundColor DarkCyan
Write-Host "---------------------------------------------" -ForegroundColor DarkGray

Write-Host ("{0,-18}: " -f "Tenant") -NoNewline -ForegroundColor DarkGray
Write-Host $TenantName -ForegroundColor Yellow

Write-Host ("{0,-18}: " -f "OSVersion") -NoNewline -ForegroundColor DarkGray
Write-Host $OSVersion -ForegroundColor Yellow

Write-Host ("{0,-18}: " -f "OSEdition") -NoNewline -ForegroundColor DarkGray
Write-Host $OSEdition -ForegroundColor Yellow

Write-Host ("{0,-18}: " -f "OSLanguage") -NoNewline -ForegroundColor DarkGray
Write-Host $OSBuild -ForegroundColor Yellow

Write-Host ("{0,-18}: " -f "OSLanguage") -NoNewline -ForegroundColor DarkGray
Write-Host $OSLanguage -ForegroundColor Yellow

Write-Host ("{0,-18}: " -f "OSActivation") -NoNewline -ForegroundColor DarkGray
Write-Host $OSActivation -ForegroundColor Yellow

Write-Host ("{0,-18}: " -f "OOBEScript") -NoNewline -ForegroundColor DarkGray
Write-Host $script:SetupCompleteDisplay -ForegroundColor Yellow

Write-Host ("{0,-18}: " -f "GroupTag") -NoNewline -ForegroundColor DarkGray
Write-Host $GroupTag -ForegroundColor Yellow

Write-Host ("{0,-18}: " -f "Autopilot") -NoNewline -ForegroundColor DarkGray

if (-not $UploadToAutopilot) {
    Write-Host "Not used" -ForegroundColor Yellow
}
elseif ($deviceAlreadyExists) {
    Write-Host "Already uploaded" -ForegroundColor Green
}
elseif (-not $authSucceeded) {
    Write-Host "Authentication failed" -ForegroundColor Red
}
elseif (-not $uploadSucceeded) {
    Write-Host "Upload failed" -ForegroundColor Red
}
elseif (-not $profileAssigned) {
    Write-Host "Profile not assigned in allowed time" -ForegroundColor Red
}
else {
    Write-Host "Upload succeeded" -ForegroundColor Green
}

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
        Write-Host "`nAutopilot encountered an issue." -ForegroundColor DarkRed

        if (-not $authSucceeded) {
            Write-Host "Reason: authentication did not succeed." -ForegroundColor Red
        }

        if ($authSucceeded -and -not $uploadSucceeded) {
            Write-Host "Reason: upload did not succeed." -ForegroundColor Red
        }

        if ($authSucceeded -and $uploadSucceeded -and -not $profileAssigned) {
            Write-Host "Reason: profile is not assigned in the allowed time." -ForegroundColor Red
        }
        
        do {
            $continueChoice = Read-Host "Continue provisioning without Autopilot anyway? (Y/N)"
        } until ($continueChoice -match '^[YyNn]$')

        if ($continueChoice -match '^[Nn]$') {
            Write-Host "Chose not to continue. Restarting computer in 10 seconds..." -ForegroundColor Yellow
            Start-Sleep -Seconds 10
            Restart-Computer
            Start-Sleep -Seconds 10
        }
    }

Write-Host "`nStarting OSDCloud in 10 seconds..." -ForegroundColor Yellow
Start-Sleep -Seconds 10

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
        Write-Host "Now launching Start-OSDCloud..." -ForegroundColor Yellow            
        Update-Module -Name OSD -ErrorAction Ignore
        Start-OSDCloud @StartOSDCloudParams -ErrorAction Stop
    }
    catch {
        Write-Host "Start-OSDCloud failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Falling back to Start-OSDCloudGUI..." -ForegroundColor Yellow

        try {
            Start-OSDCloudGUI -ErrorAction Stop
        }
        catch {
            Write-Host "Start-OSDCloudGUI also failed: $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
}
#endregion

#region: Stop Logging
# End timer
$endTime = Get-Date
$duration = $endTime - $startTime
$durationFormatted = "{0:hh\:mm\:ss}" -f $duration

Write-Host "Multi-Tenant Provisioner ended at: $($endTime.ToString('dd-MM-yyyy HH:mm:ss'))" -ForegroundColor Yellow
Write-Host "Total runtime: $durationFormatted" -ForegroundColor Cyan
Stop-Transcript
#endregion