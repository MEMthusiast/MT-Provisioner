<#
.SYNOPSIS
    Collects Windows Autopilot hardware hash from the selected tenants in WinPE and uploads to Microsoft Intune Autopilot
.DESCRIPTION
    This script gathers the Windows Autopilot hardware hash using OA3Tool while in WinPE,
    including TPM information by registering the PCPKsp.dll, and then uploads the device to Windows Autopilot via Microsoft Graph API
    Place oa3tool.exe and PCPKsp.dll files in the same folder as this script.
.PARAMETER GroupTag
    Optional. Specifies the Autopilot group tag for all tenants to assign to the device.
.PARAMETER TenantId
    Required for upload. Specifies the Entra ID tenant ID for each tenant (line 44).
.PARAMETER AppId
    Required for upload. Specifies the app registration ID for authentication.
.PARAMETER AppSecret
    Required for upload. Specifies the app registration secret for authentication.
.PARAMETER UploadToAutopilot
    Optional. Indicates whether to upload the device to Autopilot. Default is $false.
.NOTES
    File Name: TenantSelectorAutopilotHashUpload.ps1
    Author: https://github.com/MEMthusiast
    Based on:
    https://github.com/blawalt/WinPEAP
    Mike Mdm's approach (https://mikemdm.de/2023/01/29/can-you-create-a-autopilot-hash-from-winpe-yes/)
#>

# -------------------------------------------------
# Tenant selection logic starts here
# -------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$GroupTag = "Personal"
$UploadToAutopilot = $true
$AppSecret = ""
$AppId = ""

# -------------------------------------------------
# TENANTS
# -------------------------------------------------

$Tenants = @(
    @{
        Name = "Tenant 1"
        TenantId = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
    },
    @{
        Name = "Tenant 2"
        TenantId = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
    },
    @{
        Name = "Tenant 3"
        TenantId = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
    }
)

# DROPDOWN UI

$form = New-Object System.Windows.Forms.Form
$form.Text = "Autopilot Tenant Selector"
$form.Size = "400,200"
$form.StartPosition = "CenterScreen"
$form.ControlBox = $false

$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.Location = "40,40"
$dropdown.Size = "300,30"
$dropdown.DropDownStyle = "DropDownList"

$Tenants.Name | ForEach-Object {
    $dropdown.Items.Add($_)
}

$form.Controls.Add($dropdown)

$button = New-Object System.Windows.Forms.Button
$button.Text = "Start"
$button.Location = "150,90"
$form.Controls.Add($button)

$form.Controls.Add($button)

# BUTTON ACTION

$button.Add_Click({

    if(!$dropdown.SelectedItem){
        [System.Windows.Forms.MessageBox]::Show("Select a tenant")
        return
    }

    $SelectedTenant = $Tenants | Where-Object Name -eq $dropdown.SelectedItem

    $script:TenantId = $SelectedTenant.TenantId
    $form.Close()
})

$form.ShowDialog()

Write-Host "Selected customer: $($dropdown.SelectedItem)" -ForegroundColor Cyan


# -------------------------------------------------
# Autopilot logic starts here
# -------------------------------------------------

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
            "Windows Product ID" = $product
            "Hardware Hash" = $hash
            "Group Tag" = $GroupTag
        }
        
        # Save to temp CSV
        $computers += $c
        $computers | Select "Device Serial Number", "Windows Product ID", "Hardware Hash", "Group Tag" | 
            ConvertTo-CSV -NoTypeInformation | % {$_ -replace '"',''} | Out-File $TempCSVPath
    }
    else
    {
        # Create a pipeline object without Group Tag
        $c = New-Object psobject -Property @{
            "Device Serial Number" = $serial
            "Windows Product ID" = $product
            "Hardware Hash" = $hash
        }
        
        # Save to temp CSV
        $computers += $c
        $computers | Select "Device Serial Number", "Windows Product ID", "Hardware Hash" | 
            ConvertTo-CSV -NoTypeInformation | % {$_ -replace '"',''} | Out-File $TempCSVPath
    }
    
    Write-Host "CSV file created at: $TempCSVPath" -ForegroundColor Green
    
    # Upload to Autopilot if requested
    if ($UploadToAutopilot)
    {
        if (-not $TenantId -or -not $AppId -or -not $AppSecret)
        {
            Write-Host "Error: TenantId , AppId, and AppSecret parameters are required for Autopilot upload" -ForegroundColor Red
        }
        else
        {
            try {
                # Get auth token
                Write-Host "Getting authorization token..." -ForegroundColor Yellow
                $authToken = Get-AuthToken -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret
                
                # Upload device to Autopilot
                Write-Host "Adding device to Autopilot..." -ForegroundColor Yellow
                $importedDevice = Add-AutopilotImportedDevice -SerialNumber $serial -HardwareHash $hash -GroupTag $GroupTag -AuthToken $authToken
                
                #Check if device already exists in Autopilot
                $device = Get-AutopilotDevice -Serial $serial -AuthToken $authToken
                if ($device) {
                    #Device already exists in Autopilot
                    Write-Host "Device already exists in Autopilot with SerialNumber: $serial" -ForegroundColor Green
                }
                else {
                    if ($importedDevice) {
                        Write-Host "Device added successfully with ID: $($importedDevice.id)" -ForegroundColor Green
                        
                        # Wait for processing to complete
                        Write-Host "Waiting for import to complete..." -ForegroundColor Yellow
                        $processingComplete = $false
                        $maxRetries = 20
                        $retryCount = 0
                        
                        while (-not $processingComplete -and $retryCount -lt $maxRetries) {
                            Start-Sleep -Seconds 15
                            $device = Get-AutopilotImportedDevice -Id $importedDevice.id -AuthToken $authToken
                            
                            if ($device.state.deviceImportStatus -eq "complete") {
                                $processingComplete = $true
                                Write-Host "Import completed successfully!" -ForegroundColor Green
                                Write-Host "Device Registration ID: $($device.state.deviceRegistrationId)" -ForegroundColor Cyan
                            }
                            elseif ($device.state.deviceImportStatus -eq "error") {
                                Write-Host "Import failed with error: $($device.state.deviceErrorCode) - $($device.state.deviceErrorName)" -ForegroundColor Red
                                break
                            }
                            else {
                                Write-Host "Import status: $($device.state.deviceImportStatus). Waiting..." -ForegroundColor Yellow
                                $retryCount++
                            }
                        }
                        
                        if (-not $processingComplete) {
                            Write-Host "Import did not complete within the expected time." -ForegroundColor Yellow
                        }
                    }
                }
            }
            catch {
                Write-Host "An error occurred during the Autopilot upload process: $_" -ForegroundColor Red
            }
        }
    }
    else {
        Write-Host "Skipping Autopilot upload. Use -UploadToAutopilot switch with required parameters to upload." -ForegroundColor Yellow
    }
}
else
{
    Write-Host "No Hardware Hash found" -ForegroundColor Red
    Pop-Location
    exit 1
}

Write-Host "Waiting for Autopilot profile assignment..." -ForegroundColor Yellow

$profileAssigned = $false
$maxRetries = 50
$retryCount = 0

while (-not $profileAssigned -and $retryCount -lt $maxRetries) {

    Start-Sleep -Seconds 15

    $apDevice = Get-AutopilotDevice -Serial $serial -AuthToken $authToken

    if ($apDevice) {

        $status = $apDevice.deploymentProfileAssignmentStatus

        if ($status -eq "assignedUnkownSyncState") {

            Write-Host "Autopilot profile assigned successfully!" -ForegroundColor Green
            Write-Host "Profile ID: $($apDevice.deploymentProfileId)" -ForegroundColor Cyan
            $profileAssigned = $true

        }
        elseif ($status -eq "assigning") {

            Write-Host "Profile assignment in progress..." -ForegroundColor Yellow

        }
        else {

            Write-Host "Profile status: $status - waiting..." -ForegroundColor Yellow
        }

    }
    else {

        Write-Host "Device not yet visible in Autopilot device list..." -ForegroundColor Yellow
    }

    $retryCount++
}

if (-not $profileAssigned) {
    Write-Host "Autopilot profile was not assigned within expected time." -ForegroundColor Yellow
}

# Return to original directory
Pop-Location
