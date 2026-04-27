# Multi-Tenant Provisioner

> A PowerShell-based graphical user interface build on top of OSDCloud for selecting a provisioning profile

## 📋 Table of Contents

- [Overview](#Overview)
- [Optional prerequisites](#Optional-prerequisites)
- [Usage](#Prerequisites)
- [Tenants Configuration](#Tenants-Configuration)
- [OSDCloud](#osdcloud)

## Overview

WinPE-compatible PowerShell tool to collect Autopilot hardware hashes, select a tenant, and register devices in Microsoft Intune with automatic profile assignment validation.

This tool is designed for MSPs and to be run in combination with OSDCloud in a WinPE environment during the deployment of a Windows device.

The Graph authentication logic is based on a multi-tenant app registration in Entra ID, allowing the same App ID and App Secret to be used across all tenants.

OSDCloud: https://github.com/OSDeploy/OSDCloud

Multi tenant app: https://learningbytesblog.com/posts/Muiltitenant-Entra-APP-for-multitenant-managment/

Autopilot logic used in this tool and OSDCloud USB creation based on: https://github.com/blawalt/WinPEAP

## Screenhots

### Tenant Selector
*test*
<img width="719" height="564" alt="Image" src="https://github.com/user-attachments/assets/1b94467e-c879-486c-a2eb-b98818f32f51" />

## Optional prerequisites

* A (multi-tenant) Entra ID enterprise application in every tenant

* An Azure Key Vault: https://learn.microsoft.com/en-us/azure/key-vault/general/quick-create-portal

* An Azure Blob Storage: https://learn.microsoft.com/en-us/azure/storage/common/storage-account-create?tabs=azure-portal


## Prerequisites

* **Building the Tenants configuration:** Inside the *Start-MTP.ps1* or with *Export-TentansConfig.ps1*
    * Edit Start-MTP.ps1 and go to: *#region: Hardcoded Tenant Parameters* and fill in the parameters of every tenant you want to provision.

    If you only want to provision an OS you can set *UploadToAutopilot* to **$false** and change *Name* to for example **Windows 11 Pro**

    ```powershell
        Name = "Tenant 1"
        TenantId = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
        UploadToAutopilot = $true
        GroupTag = "TENANT1"
        OSBuild = "25H2"
        OSEdition = "Pro"
        OSVersion = "Windows 11"
        OSLanguage = "nl-nl"
        OSActivation = "Volume"
    ```
    ### test
    
    
    ```powershell
    Install-Module PartnerCenter
    ```

* **OSDCloud PowerShell module**
    ```powershell
    Install-Module OSD
    ```
* **Windows Assessment and Deployment Kit (ADK) and WinPE Add-on:** Install the Windows 10 ADK and the WinPE add-on. These provide deployment tools, including WinPE itself and the `oa3tool.exe` needed later.
    * Download link: [Windows ADK Download](https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install)
    * Ensure installation of **ADK** and the **WinPE Add-on** 

## Tenants Configuration

## OSDCloud

* Create an OSDCloud Template
    ```powershell
    New-OSDCloudTemplate -SetAllIntl en-us -SetInputLocale en-us
    ```

* Create an OSDCloud WorkSpace
    ```powershell
    New-OSDCloudWorkspace -WorkspacePath "C:\OSDCloud\MTP"
    ```
* Copy files

* Create WinPE
    ```powershell
    Edit-OSDCloudWinPE -Wallpaper "C:\path\to\your\background.jpg"
    ```
 * Update WinPE
    ```powershell
    Edit-OSDCloudWinPE
    ```

## Test in Hyper-V

## Create bootable USB
*Create a bootable USB
     ```powershell
    New-OSDCloudUSB
    ```

 *If you make changes to WinPE in your OSDCloud Workspace, you can easily update your OSDCloud USB WinPE volume by using Update-OSDCloudUSB
     ```powershell
    Update-OSDCloudUSB
    ```

## Use with WDS PXE