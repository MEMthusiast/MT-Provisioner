<#
.SYNOPSIS
    Collects hardware hash in WinPE and uploads it to Intune Autopilot of the selected tenant. Optionally run a script before OOBE.
.DESCRIPTION
    This script is designed for MSPs and to be run in combination with OSDCloud in a WinPE environment during the deployment of a Windows device.
    The Graph authentication logic is based on a multi-tenant app registration in Entra ID, allowing the same App ID and App Secret to be used across all tenants.    
    It performs the following functions:
        1. Displays a tenant selection UI for the user to choose which tenant to upload the hardware hash to.
        2. Collects the hardware hash using OA3Tool, including TPM information by registering the PCPKsp.dll.
        3. Uploads the hardware hash to the selected tenant's Intune Autopilot via Microsoft Graph API.
        4. Waits for the Autopilot profile assignment to complete.
        5. Optionally downloads a SetupComplete.ps1 script and saves it to the appropriate location in the WinPE environment where it later will be copied to be executed before OOBE.
        6. Optionally downloads a tenant configuration JSON file, allowing for dynamic configuration without modifying the script.
        7. Supports retrieving the multitenant enterprise app secret from Azure Key Vault for enhanced security when allowing the keyvault to only be accessed from a trusted public IP address.
.PARAMETER GroupTag
    Conditionally required. Set this parameter for tenants that don't have a GroupTag property in the config file (ParametersUrl or hardcoded parameters). The GroupTag value in the config will overwrite this value.
.PARAMETER UploadToAutopilot
    Optional. Indicates whether to upload the device to Autopilot. Default is $false.
.PARAMETER AppId
    Required. The app registration ID for authentication. This is the AppID of the multitenant Entra ID enterprise app.
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
    Optional. A URL to download a SetupComplete.ps1 script. This script will be copied to the WinPE image and will run during the SetupComplete phase before OOBE.
.PARAMETER ParametersUrl
    Optional. If provided, the script will read tenant information from the downloaded JSON instead of using hardcoded tenant parameters inside this script.
.NOTES
    File Name   : TenantSelectorAutopilotHashUpload.ps1
    Author      : https://github.com/MEMthusiast
    Version     : 2.05
    Purpose     : Upload device hashes to selected tenant's Autopilot.
    Requires    : OSDCloud, a multi-tenant Entra ID enterprise app in each customer tenant, and optionally an Azure Key Vault for secret retrieval and hosting the SetupComplete.ps1 and config.json files in an Azure Blob that is only accessable from a trusted public IP address.
    References  : Autopilot upload logic based on: https://github.com/blawal/WinPEAP
    Usage       : Can be used as a standalone script but used best in combiantion with OSDCloud. This is automatically executed as part of the OSDCloud process. Customize parameters as needed before deployment.
#>

#region: Required Parameters
    $UploadToAutopilot  = $true     # Set to $false to disable Autopilot upload, or $true to enable the upload step
    $AppId              = ""        # The app ID of the multitenant Entra ID app registration.
    $KeyVault           = $true     # Set to $false to skip Key Vault retrieval and use hardcoded $AppSecret. Set to $true to use Key Vault.
#endregion

#region: Conditionally Required Parameters
    $GroupTag           = ""    # Use this for tenants that don't have a GroupTag property in the config file (ParametersUrl or hardcoded parameters).
    $AppSecret          = ""    # Use this when not using Key Vault for secret retrieval.
#endregion

#region: Optional Parameters
    $SPNAppID           = ""    # The app ID for authentication to Key Vault.
    $SPNSecret          = ""    # The app secret for authentication to Key Vault.
    $SPNTenantID        = ""    # The tenant ID for authentication to Key Vault.
    $VaultName          = ""    # Name of your Key Vault
    $SecretName         = ""    # Name of the secret to retrieve
    $SetupCompleteUrl   = ""    # Example: https://raw.githubusercontent.com/MEMthusiast/Intune-Autopilot-MultiTenant/refs/heads/main/SetupComplete/SetupComplete.ps1
    $ParametersUrl      = ""    # Example: https://raw.githubusercontent.com/MEMthusiast/Intune-Autopilot-MultiTenant/refs/heads/main/config.json
#endregion

#region: Start Logging
    # Start transcript
    $Global:Transcript = "$((Get-Date).ToString('dd-MM-yyyy-HHmmss'))-Invoke-MT-AP.log"
    Start-Transcript -Path (Join-Path "X:\OSDCloud\Config\Scripts\SetupComplete" $Global:Transcript) -ErrorAction Ignore

    # Start timer
    $startTime = Get-Date
    Write-Host "Script started at: $($startTime.ToString('dd-MM-yyyy-HHmmss'))" -ForegroundColor Yellow
#endregion

#region: Test internet connectity
    if ($KeyVault -or $SetupCompleteUrl -or $ParametersUrl) {
        Write-Host "`nInternet connectivity is required. Checking connection..." -ForegroundColor Cyan

        $RetryInterval = 10
        $firstAttempt = $true

        while ($true) {
            try {
                # Test internet connectivity
                Invoke-WebRequest -Uri "https://www.microsoft.com/" -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop | Out-Null
                Write-Host "`nInternet connection detected. Continuing script..." -ForegroundColor Green
                break
            }
            catch {
                if ($firstAttempt) {
                    Write-Host "`nNo internet connection detected." -ForegroundColor Yellow
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
        Write-Host "Continuing without internet connectivity. Script functionality may be limited." -ForegroundColor Yellow
    }
    # Get public IP address
    try {
    $publicIP = (Invoke-RestMethod -Uri "https://api.ipify.org?format=json").ip
        } catch {} 
#endregion

#region: Key Vault
    if ($KeyVault) {

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

#region: Injected Parameters
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
        Write-Host "ParameterUrl not provided. Using hardcoded parameters." -ForegroundColor Yellow
    }
#endregion

#region: Hardcoded Tenant Parameters
    # If ParametersUrl is not provided, use hardcoded tenant information. This is useful if you prefer to have the tenant information directly in the script.

    if ( [string]::IsNullOrWhiteSpace($ParametersUrl)) {
    
    $Tenants = @(
    @{
        Name = "Tenant 1"
        TenantId = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
        GroupTag = "AP-Tenant1"
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
#endregion

#region: Tenant Selection UI
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
                GroupTag = $_.GroupTag
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

        # Only override the default GroupTag if the tenant has one filled in
        if (-not [string]::IsNullOrWhiteSpace($SelectedTenant.GroupTag)) {
            $script:GroupTag = $SelectedTenant.GroupTag
        }

        $form.Close()
    })

    # Show the form
    $form.ShowDialog()

    Write-Host "Selected Tenant: $TenantName" -ForegroundColor Green
#endregion

#region: Autopilot Upload
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
        $TempCSVPath = "X:\OSDCloud\Config\Scripts\SetupComplete\AutopilotHash.csv"

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

        if ($UploadToAutopilot)
        {
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
        Write-Host "No Hardware Hash found" -ForegroundColor Red
        Pop-Location
        exit 1
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

            Invoke-WebRequest -Uri $SetupCompleteUrl -OutFile $DestinationFile -UseBasicParsing -ErrorAction Stop

            Write-Host "$FileName downloaded successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to download $FileName $($_.Exception.Message)" -ForegroundColor Red
    }   

    }
    else {
    Write-Host "SetupCompleteUrl not provided. Skipping SetupComplete.ps1 download." -ForegroundColor Yellow
    }
#endregion

#region: Stop Logging
    # End timer
    $endTime = Get-Date
    $duration = $endTime - $startTime
    $durationFormatted = "{0:hh\:mm\:ss}" -f $duration

    Write-Host "Script ended at: $($endTime.ToString('dd-MM-yyyy HH:mm:ss'))" -ForegroundColor Yellow
    Write-Host "Total runtime: $durationFormatted" -ForegroundColor Cyan
    Stop-Transcript
#endregion

Write-Host "Starting OSDCloud in 10 seconds..." -ForegroundColor Yellow
Start-Sleep -Seconds 10

# Return to original directory
Pop-Location