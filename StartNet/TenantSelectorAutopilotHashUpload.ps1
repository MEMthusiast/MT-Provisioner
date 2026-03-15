<#
.SYNOPSIS
    Collects hardware hash in WinPE and uploads it to Intune Autopilot of the selected tenant.
    And downloads the SetupComplete.ps1 script to WINPE so it can later on be picked up by OSDCloud and executed before the OOBE phase.
.DESCRIPTION
    This script is designed for MSPs and to be run in combination with OSDCloud in a WinPE environment during the deployment of a Windows device.
    The Graph authentication logic is based on a multi-tenant app registration in Entra ID, allowing the same App ID and App Secret to be used across all tenants.
    
    It performs the following functions:
        1. Displays a tenant selection UI for the user to choose which tenant to upload the hardware hash to.
        2. Collects the hardware hash using OA3Tool, including TPM information by registering the PCPKsp.dll.
        3. Uploads the hardware hash to the selected tenant's Intune Autopilot via Microsoft Graph API.
        4. Waits for the Autopilot profile assignment to complete.
        5. Optionally downloads a SetupComplete.ps1 script and saves it to the appropriate location in the WinPE environment where it later will be copied to be executed before OOBE.
.PARAMETER GroupTag
    Required. The Autopilot group tag for all tenants to assign to the device.
.PARAMETER UploadToAutopilot
    Optional. Indicates whether to upload the device to Autopilot. Default is $false.
.PARAMETER AppSecret
    Required. The app registration secret for authentication.
.PARAMETER AppId
    Required. The app registration ID for authentication.
.PARAMETER Name
    Specifies the tenant name to display in the tenant selection UI. You can choose here for injected parameters or hardcoded parameters.
    If using hardcoded parameters, the tenant names will be defined in the script go to #region hardcoded parameters    
    If using injected parameters, the tenant names will be read from the config file.
.PARAMETER TenantId
    Specifies the Entra ID tenant ID for each tenant to upload the hardware hash to. You can choose here for injected parameters or hardcoded parameters.
    If using hardcoded parameters, the tenant IDs will be defined in the script go to #region hardcoded parameters
    If using injected parameters, the tenant IDs will be read from the config file.
.PARAMETER SetupCompleteUrl
    Optional. URL to download a SetupComplete.ps1 script. This script will be copied to the WinPE image and will run during the SetupComplete phase before OOBE.
    If not provided, the script will skip downloading. The default URL points to a sample SetupComplete.ps1 in this repository.
.PARAMETER ParametersUrl
    Optional. If provided, the script will read tenant information from the downloaded JSON instead of using hardcoded parameters.
.NOTES
    File Name: TenantSelectorAutopilotHashUpload.ps1
    Author: https://github.com/MEMthusiast
    Autopilot logic based on: https://github.com/blawalt/WinPEAP
#>

#region required parameters

    $GroupTag = ""
    $UploadToAutopilot = $true 
    $AppSecret = ""
    $AppId = ""

#endregion required parameters

#region optional parameters

$SetupCompleteUrl  = ""
#Example: https://raw.githubusercontent.com/MEMthusiast/Intune-Autopilot-MultiTenant/refs/heads/main/SetupComplete/SetupComplete.ps1

$ParametersUrl = ""
#Example: https://raw.githubusercontent.com/MEMthusiast/Intune-Autopilot-MultiTenant/refs/heads/main/config.json

#endregion optional parameters

#region injected parameters
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
        Invoke-WebRequest -Uri $ParametersUrl -OutFile $DestinationFile -UseBasicParsing -ErrorAction Stop

        Write-Host "$FileName downloaded successfully to $DestinationFolder." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to download $FileName : $($_.Exception.Message)"
    }

    $config = Get-Content "X:\OSDCloud\Config\config.json" -Raw | ConvertFrom-Json
    $Tenants = $config.Tenants

    }
    else {
        Write-Host "ParameterUrl not provided. Using hardcoded parameters." -ForegroundColor Yellow
    }

#endregion injected parameters

#region hardcoded parameters
# If ParametersUrl is not provided, use hardcoded tenant information. This is useful if you prefer to have the tenant information directly in the script.

    if ( [string]::IsNullOrWhiteSpace($ParametersUrl)) {
    
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
    }
    else {
        Write-Host "ParametersUrl provided. Skipping hardcoded parameters and using injected parameters." -ForegroundColor Yellow
    }

#endregion hardcoded parameters

#region Tenant Selection UI

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tenant Selector"
    $form.Size = "450,250"
    $form.StartPosition = "CenterScreen"
    $form.ControlBox = $false

    # Search TextBox
    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = "40,20"
    $searchBox.Size = "350,25"
    $form.Controls.Add($searchBox)

    # Dropdown ComboBox
    $dropdown = New-Object System.Windows.Forms.ComboBox
    $dropdown.Location = "40,60"
    $dropdown.Size = "350,30"
    $dropdown.DropDownStyle = "DropDownList"
    $form.Controls.Add($dropdown)

    # Convert all tenants to PSCustomObjects
    $TenantObjects = $Tenants | ForEach-Object {
        if ($_ -is [hashtable]) {
            [PSCustomObject]@{
                Name     = $_.Name
                TenantId = $_.TenantId
            }
        } else {
            $_
        }
    }

    # Sort alphabetically
    $SortedTenants = $TenantObjects | Sort-Object Name

    # Clone for search filtering
    $AllTenants = $SortedTenants.Clone()

    # Populate dropdown initially with all tenants
    foreach ($tenant in $AllTenants) {
        [void]$dropdown.Items.Add($tenant)
    }

    # Display tenant names
    $dropdown.DisplayMember = "Name"

    # Search box filtering logic
    $searchBox.Add_TextChanged({
        $filter = $searchBox.Text.Trim().ToLower()
        $dropdown.Items.Clear()

        # Filter tenants
        $filtered = if ([string]::IsNullOrWhiteSpace($filter)) {
            $AllTenants
        } else {
            $AllTenants | Where-Object { $_.Name.ToLower() -like "*$filter*" }
        }

        # Add filtered items
        foreach ($tenant in $filtered) {
            [void]$dropdown.Items.Add($tenant)
        }

        # Select first item if available
        if ($dropdown.Items.Count -gt 0) { 
            $dropdown.SelectedIndex = 0 
        }
    })

    # Start button
    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Start"
    $button.Location = "150,120"
    $form.Controls.Add($button)

    $button.Add_Click({
        if (!$dropdown.SelectedItem) {
            [System.Windows.Forms.MessageBox]::Show("Select a tenant")
            return
        }

        $SelectedTenant = $dropdown.SelectedItem

        # Set only tenant-related variables
        $script:TenantId   = $SelectedTenant.TenantId
        $script:TenantName = $SelectedTenant.Name

        $form.Close()
    })

    # Show the form
    $form.ShowDialog()

    Write-Host "Selected customer: $TenantName" -ForegroundColor Green

#endregion Tenant Selection UI

#region Autopilot Upload

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

#endregion Autopilot Upload

#region Download SetupComplete.ps1

    if (-not [string]::IsNullOrWhiteSpace($SetupCompleteUrl)) {

    Write-Host "Downloading SetupComplete.ps1..." -ForegroundColor Cyan

    $DestinationFolder = "X:\OSDCloud\Config\Scripts\SetupComplete"

    try {
        # Ensure TLS 1.2+
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        # Extract filename and construct destination path
        $FileName        = Split-Path -Path $SetupCompleteUrl -Leaf
        $DestinationFile = Join-Path -Path $DestinationFolder -ChildPath $FileName

        # Download only if folder exists
        if ((Test-Path -Path $DestinationFolder)) {

            Invoke-WebRequest -Uri $SetupCompleteUrl -OutFile $DestinationFile -UseBasicParsing -ErrorAction Stop

            Write-Host "$FileName downloaded successfully to $DestinationFolder." -ForegroundColor Green
        }
        else {
            Write-Host "Folder $DestinationFolder does not exist." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "Failed to download $FileName $($_.Exception.Message)" -ForegroundColor Red
    }   

    }
    else {
    Write-Host "SetupCompleteUrl not provided. Skipping SetupComplete.ps1 download." -ForegroundColor Yellow
    }
 
#endregion Download SetupComplete.ps1

# Return to original directory
Pop-Location