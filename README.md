# Intune-Autopilot-MultiTenant
WinPE-compatible PowerShell tool to collect Autopilot hardware hashes, select a tenant, and register devices in Microsoft Intune with automatic profile assignment validation.

Idea from: https://github.com/blawalt/WinPEAP

Best to Use in combination with OSDCloud and a Multi tenant enterprise application that has Autopilot rights.

OSDCloud: https://github.com/OSDeploy/OSDCloud
Multi tenant app: https://learningbytesblog.com/posts/Muiltitenant-Entra-APP-for-multitenant-managment/

# TenantSelectorAutopilotHashUpload.ps1


# SetupComplete.ps1

Default OSDCloud log path
C:\OSDCloud\Logs
C:\Windows\Temp\osdcloud-logs

All logs moved with SetupComplete script to: C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\OSD
