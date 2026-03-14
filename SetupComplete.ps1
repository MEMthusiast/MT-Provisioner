# Transcript log file name with timestamp
    $Global:Transcript = "$((Get-Date).ToString('yyyy-MM-dd-HHmmss'))-StartupComplete-Script.log"
    Start-Transcript -Path (Join-Path "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD\" $Global:Transcript) -ErrorAction Ignore

#region cleanup OSDCloud
    Write-Host "Execute OSD Cloud Cleanup Script"

    # Copying the OOBEDeploy and AutopilotOOBE Logs
    Get-ChildItem 'C:\Windows\Temp' -Filter *OOBE* | Copy-Item -Destination 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD' -Force

    # Copying OSDCloud Logs
    If (Test-Path -Path 'C:\OSDCloud\Logs') {
        Move-Item 'C:\OSDCloud\Logs\*.*' -Destination 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD' -Force
    }
    If (Test-Path -Path 'C:\Windows\Temp\osdcloud-logs') {
        Move-Item 'C:\Windows\Temp\osdcloud-logs' -Destination 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD' -Force
    }
    Move-Item 'C:\ProgramData\OSDeploy\*.*' -Destination 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD' -Force

    If (Test-Path -Path 'C:\Temp') {
        Get-ChildItem 'C:\Temp' -Filter *OOBE* | Copy-Item -Destination 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD' -Force
        Get-ChildItem 'C:\Windows\Temp' -Filter *Events* | Copy-Item -Destination 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD' -Force
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
    Write-Host "Enable OEM Product Key"

    $(Get-WmiObject SoftwareLicensingService).OA3xOriginalProductKey | foreach{ if ( $null -ne $_ ) { Write-Host "Installing"$_;changepk.exe /Productkey $_ } else { Write-Host "No key present" } }

#endregion Enable WIndows Product Key

#region Enable BitLocker
    Write-Host "Enable BitLocker"

    # Get driveletters from Internaldrives
    $disks = Get-Disk | Where-Object -FilterScript {$_.Bustype -ne "USB"}
    $driveletters = @()

    foreach ($disk in $disks)  {
        $partitions = Get-Partition -DiskNumber $disk.Number

        foreach ($partition in $partitions) {

            if ($($partition.DriveLetter)) {
                $driveletters += "$($partition.DriveLetter)"
                }
        }
    }

    # Check if TPM chip is available
    $TPM = Get-TPM

    If ($TPM.TpmPresent -like "False"){
            Write-Host -Message "No TPM-chip detected"
            Exit 1
        }

    ForEach ($DriveLetter in $DriveLetters) {
        $DriveLetter2 = "$DriveLetter"+":"
        $BitlockerStatus = Get-BitLockerVolume -MountPoint $DriveLetter
        
        If ($BitlockerStatus.ProtectionStatus -like "*off*" -or $BitlockerStatus.EncryptionMethod -ne "XtsAes256")  {
            #set registerkeys to newest encryption method
            Set-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\FVE" -Name "EncryptionMethod" -Type "Dword" -Value "7"
            Set-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\FVE" -Name "EncryptionMethodWithXtsOs" -Type "Dword" -Value "7"
            Set-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\FVE" -Name "EncryptionMethodWithXtsFdv" -Type "Dword" -Value "7"
            
            #Disable bitlocker if not decrypted
            if ($BitlockerStatus.VolumeStatus -ne "FullyDecrypted"){
                try {
                    Clear-BitLockerAutoUnlock
                    Disable-Bitlocker -MountPoint $DriveLetter
                }
                catch{Write-Host -Message "Error disabling bitlocker"}
            }
            
            #wait until decryption is complete
            $DecryptionComplete = $false
            while (-not $DecryptionComplete) {
                Start-Sleep -Seconds 10
                $BitlockerStatus = Get-BitLockerVolume -MountPoint $DriveLetter
                if ($BitlockerStatus.VolumeStatus -eq "FullyDecrypted") {
                    $DecryptionComplete = $true
                }
                #view process in log
                Write-Host -Message "DecryptionPercentage $($BitlockerStatus.EncryptionPercentage)"
            }
            
            #Add TPM chip for autounlock OS-disk
            if ($BitlockerStatus.VolumeType -eq "OperatingSystem"){
                try {Add-BitLockerKeyProtector -MountPoint $DriveLetter -TpmProtector}
                catch {Write-Host -Message "Error TPM add to OperatingSystem drive"}
            }

            #Encrypt disk
            try {Enable-Bitlocker -MountPoint $DriveLetter -SkipHardwareTest -RecoveryPasswordProtector}
            catch {Write-Host -Message "Error enabling bitlocker"}

            #wait until encryption status 100%
            $encryptionComplete = $false
            while (-not $encryptionComplete) {
                Start-Sleep -Seconds 10
                $BitlockerStatus = Get-BitLockerVolume -MountPoint $DriveLetter
                if (($BitlockerStatus.EncryptionPercentage -eq 100) -and ($BitlockerStatus.VolumeStatus -eq "FullyEncrypted")) {
                    $encryptionComplete = $true
                }
                #view process in log
                Write-Host -Message "EncryptionPercentage $($BitlockerStatus.EncryptionPercentage)"
            }
            
            #backup bitlocker key to Microsoft
            if ($BitlockerStatus.VolumeType -eq "OperatingSystem"){
                try {BackupToAAD-BitLockerKeyProtector -MountPoint $DriveLetter -KeyProtectorId $BitlockerStatus.KeyProtector[1].KeyProtectorId}
                catch {Write-Host -Message "Error backup bitlocker key to Microsoft"}
            }
            else {
                try {BackupToAAD-BitLockerKeyProtector -MountPoint $DriveLetter -KeyProtectorId $BitlockerStatus.KeyProtector[0].KeyProtectorId}
                catch {Write-Host -Message "Error backup bitlocker key to Microsoft"}
            }
            
            #resume and enable bitlocker
            try {Resume-BitLocker -MountPoint $DriveLetter}
            catch {Write-Host -Message "error resuming bitlocker"}
                
            #autounlock bitlocker Data-drives
            If ($BitlockerStatus.VolumeType -ne "OperatingSystem")  {
                try {Enable-BitLockerAutoUnlock -MountPoint $DriveLetter}
                catch {Write-Host -Message "Error autounlock"}
            }
        }
        Else  {
            Write-Host -Message "Bitlocker already enabled $($DriveLetter)"
        }
    }

    Start-Process -FilePath "C:\Program Files (x86)\Microsoft Intune Management Extension\Microsoft.Management.Services.IntuneWindowsAgent.exe" -ArgumentList "intunemanagementextension://synccompliance"

#endregion Enable Bitlocker

Stop-Transcript