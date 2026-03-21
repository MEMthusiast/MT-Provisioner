<#
.SYNOPSIS
    Cleans up OSDCloud deployment leftovers, enables OEM Product Key, enables BitLocker and downloads CMTrace to the system for easy log viewing.
.DESCRIPTION
    This script is automatically downloaded in combination with the TenantSelectorAutopilotHashUpload.ps1 script during the WinPE phase of OSDCloud deployment.
    It is designed to be executed before the OOBE phase of Windows setup, and performs several post-deployment configuration tasks to ensure the device is properly set up and secured before the user starts using it.
 
    It performs the following functions:
        1. Downloads CMTrace to the system for easy log viewing. 
        2. Cleans up OSDCloud leftovers and copies all logs to the Intune Management Extension log folder for easier troubleshooting.
        3. Enables the OEM Product Key if available.
        4. Enables BitLocker on all internal drives with TPM and XTS-AES 256 encryption.
.NOTES
    File Name: SetupComplete.ps1
    Author: https://github.com/MEMthusiast
#>

# Transcript log file name with timestamp

    $Global:Transcript = "$((Get-Date).ToString('yyyy-MM-dd-HHmmss'))-StartupComplete-Script.log"
    Start-Transcript -Path (Join-Path "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD\" $Global:Transcript) -ErrorAction Ignore

#region Download CMTrace

    Write-Host "Downloading CMTrace..."

    $Url               = "https://github.com/MEMthusiast/Intune-Autopilot-MultiTenant/raw/refs/heads/main/cmtrace.exe"
    $DestinationFolder = "C:\Windows\System32"

    try {
        # Ensure TLS 1.2+
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        # Extract filename and construct destination path
        $FileName        = Split-Path -Path $Url -Leaf
        $DestinationFile = Join-Path -Path $DestinationFolder -ChildPath $FileName

        # Download only if file doesn't already exist
        if (-not (Test-Path -Path $DestinationFile)) {

            Invoke-WebRequest -Uri $Url -OutFile $DestinationFile -UseBasicParsing -ErrorAction Stop

            Write-Host "CMTrace downloaded successfully to $DestinationFile"
        }
        else {
            Write-Host "CMTrace already exists at $DestinationFile"
        }

    }
    catch {
        Write-Error "Failed to download CMTrace: $($_.Exception.Message)"
    }

#endregion Download CMTrace

#region cleanup OSDCloud

    Write-Host "Cleaning up OSDCloud leftovers"

    # Copying OSDCloud Logs
    If (Test-Path -Path 'C:\OSDCloud\Logs') {
        Move-Item 'C:\OSDCloud\Logs\*.*' -Destination 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD' -Force
    }
    If (Test-Path -Path 'C:\OSDCloud\Scripts\SetupComplete') {
        Move-Item 'C:\OSDCloud\Scripts\SetupComplete\*.log' -Destination 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD' -Force
    }
    If (Test-Path -Path 'C:\Windows\Temp\osdcloud-logs') {
        Get-ChildItem 'C:\Windows\Temp\osdcloud-logs' | Copy-Item -Destination 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD' -Force
    }
    
    If (Test-Path -Path 'C:\ProgramData\OSDeploy') {
        Get-ChildItem 'C:\ProgramData\OSDeploy' | Copy-Item -Destination 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD' -Force
    }

    If (Test-Path -Path 'C:\Temp') {
        Get-ChildItem 'C:\Temp' -Filter *OOBE* | Copy-Item -Destination 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD' -Force
        Get-ChildItem 'C:\Windows\Temp' -Filter *Events* | Copy-Item -Destination 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD' -Force
        Get-ChildItem 'C:\Windows\Temp' -Filter *OOBE* | Copy-Item -Destination 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD' -Force
    }

    # Cleanup directories
    If (Test-Path -Path 'C:\OSDCloud') { Remove-Item -Path 'C:\OSDCloud' -Recurse -Force }
    If (Test-Path -Path 'C:\Drivers') { Remove-Item 'C:\Drivers' -Recurse -Force }
    If (Test-Path -Path 'C:\Intel') { Remove-Item 'C:\Intel' -Recurse -Force }
    If (Test-Path -Path 'C:\ProgramData\OSDeploy') { Remove-Item 'C:\ProgramData\OSDeploy' -Recurse -Force }

    # Cleanup Scripts
    Remove-Item C:\Windows\Setup\Scripts\*.* -Exclude *.TAG -Force | Out-Null

#endregion cleanup OSDCloud

#region Enable WIndows Product Key

    Write-Host "Enabling OEM Product Key"

    $key = (Get-CimInstance SoftwareLicensingService).OA3xOriginalProductKey; if ($key) { Write-Host "Installing $key"; changepk.exe /ProductKey $key } else { Write-Host "No key present" }

#endregion Enable WIndows Product Key

#region Enable BitLocker

    Write-Host "Enabling BitLocker"

    # Detect mounted CD/DVD/ISO
    $cdDrives = Get-CimInstance Win32_CDROMDrive
    if ($cdDrives) {
        Write-Host "Bootable media detected (CD/DVD/ISO). BitLocker cannot start. Remove the CD/DVD/ISO and reboot."
        Exit 1
    }

    # Get internal drive letters
    $disks = Get-Disk | Where-Object { $_.Bustype -notin @("USB","Unknown") }
    $driveletters = @()

    foreach ($disk in $disks) {
        $partitions = Get-Partition -DiskNumber $disk.Number
        foreach ($partition in $partitions) {
            if ($partition.DriveLetter) {
                $driveletters += $partition.DriveLetter
            }
        }
    }

    # Check TPM
    $TPM = Get-TPM
    if (-not $TPM.TpmPresent) {
        Write-Host "No TPM chip detected"
        Exit 1
    }

    foreach ($DriveLetter in $driveletters) {

        $MountPoint = "${DriveLetter}:"
        $BitlockerStatus = Get-BitLockerVolume -MountPoint $MountPoint

        if ($BitlockerStatus.ProtectionStatus -eq "Off" -or $BitlockerStatus.EncryptionMethod -ne "XtsAes256") {

            Write-Host "Processing drive $MountPoint"

            # Ensure BitLocker policy registry keys exist
            $Key = "HKLM:\SOFTWARE\Policies\Microsoft\FVE"
            if (!(Test-Path $Key)) {
                New-Item -Path $Key -Force | Out-Null
            }

            New-ItemProperty -Path $Key -Name "EncryptionMethod" -PropertyType DWord -Value 7 -Force | Out-Null
            New-ItemProperty -Path $Key -Name "EncryptionMethodWithXtsOs" -PropertyType DWord -Value 7 -Force | Out-Null
            New-ItemProperty -Path $Key -Name "EncryptionMethodWithXtsFdv" -PropertyType DWord -Value 7 -Force | Out-Null

            # Disable BitLocker if needed
            if ($BitlockerStatus.VolumeStatus -ne "FullyDecrypted") {

                Write-Host "Decrypting existing BitLocker configuration"

                try {
                    Clear-BitLockerAutoUnlock -ErrorAction SilentlyContinue
                    Disable-BitLocker -MountPoint $MountPoint
                }
                catch {
                    Write-Host "Error disabling BitLocker"
                }

                # Wait for decryption (max 10 minutes)
                $maxRetries = 60
                $retryCount = 0

                while ($retryCount -lt $maxRetries) {

                    Start-Sleep 10
                    $BitlockerStatus = Get-BitLockerVolume -MountPoint $MountPoint

                    Write-Host "DecryptionPercentage $($BitlockerStatus.EncryptionPercentage)"

                    if ($BitlockerStatus.VolumeStatus -eq "FullyDecrypted") {
                        break
                    }

                    $retryCount++
                }

                if ($retryCount -ge $maxRetries) {
                    Write-Host "Decryption timeout reached"
                    continue
                }
            }

            # Add TPM protector for OS disk
            if ($BitlockerStatus.VolumeType -eq "OperatingSystem") {
                try {
                    Add-BitLockerKeyProtector -MountPoint $MountPoint -TpmProtector
                }
                catch {
                    Write-Host "Error adding TPM protector"
                }
            }

            # Enable BitLocker
            try {
                Enable-BitLocker -MountPoint $MountPoint -SkipHardwareTest -RecoveryPasswordProtector
            }
            catch {
                Write-Host "Error enabling BitLocker"
                continue
            }

            # Wait for encryption (max 10 minutes)
            $maxRetries = 60
            $retryCount = 0

            while ($retryCount -lt $maxRetries) {

                Start-Sleep 10
                $BitlockerStatus = Get-BitLockerVolume -MountPoint $MountPoint

                Write-Host "EncryptionPercentage $($BitlockerStatus.EncryptionPercentage)"

                if ($BitlockerStatus.VolumeStatus -eq "FullyEncrypted") {
                    break
                }

                $retryCount++
            }

            if ($retryCount -ge $maxRetries) {
                Write-Host "Encryption timeout reached"
            }

            # Check Entra ID / Hybrid join
            $dsreg = dsregcmd /status | Out-String
            $IsEntraJoined = $false

            if ($dsreg -match "AzureAdJoined\s*:\s*YES" -or ($dsreg -match "DomainJoined\s*:\s*YES" -and $dsreg -match "AzureAdPrt\s*:\s*YES")) {

                $IsEntraJoined = $true
            }

            if ($IsEntraJoined -and $BitlockerStatus.VolumeType -eq "OperatingSystem") {

                Write-Host "Entra ID join detected, backing up BitLocker key"

                $RecoveryKey = $BitlockerStatus.KeyProtector |
                            Where-Object { $_.KeyProtectorType -eq "RecoveryPassword" }

                if ($RecoveryKey) {

                    try {
                        BackupToAAD-BitLockerKeyProtector `
                            -MountPoint $MountPoint `
                            -KeyProtectorId $RecoveryKey.KeyProtectorId

                        Write-Host "BitLocker key backed up to Entra ID"
                    }
                    catch {
                        Write-Host "Error backing up BitLocker key"
                    }
                }
            }

            # Resume BitLocker
            try {
                Resume-BitLocker -MountPoint $MountPoint
            }
            catch {
                Write-Host "Error resuming BitLocker"
            }

            # Enable auto unlock for data drives
            if ($BitlockerStatus.VolumeType -ne "OperatingSystem") {
                try {
                    Enable-BitLockerAutoUnlock -MountPoint $MountPoint
                }
                catch {
                    Write-Host "Error enabling auto unlock"
                }
            }

        }
        else {
            Write-Host "BitLocker already enabled on $MountPoint"
        }
    }
        
#endregion Enable Bitlocker

Stop-Transcript