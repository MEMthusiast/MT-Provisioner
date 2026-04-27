# Multi-Tenant Provisioner

> 💻 A PowerShell-based graphical user interface for selecting a provisioning profile and, if needed, uploading the device hardware hash to Intune Autopilot before Windows is installed.

## 📋 Table of Contents

- [Overview](#-overview)
- [Configuration Choice](#️-configuration-choice)
- [Requirements](#-requirements)
- [Tenants Configuration](#Tenants-Configuration)
- [OSDCloud](#osdcloud)
- [Screenshots](#-screenshots)
- [Credits](#-credits)

## 🎯 Overview

**Multi-Tenant Provisioner** is developed in response to Microsoft’s announced retirement of the Microsoft Deployment Toolkit (MDT).

It is intended for organizations that need multi-tenant bare-metal deployment capabilities and are looking for a MDT replacement.

**Multi-Tenant Provisioner** adds a tenant selection layer on top of OSDCloud WinPE, allowing to choose from predefined tenant-specific provisioning configurations. The Windows installation is fully handled by OSDCloud.

 ### Key Capabilities
 
- 🏢 **Multi-Tenant Profile Selection**: Select from predefined tenant-specific provisioning profiles before deployment
- 🔍 **Search & Filter**: Quickly find the correct tenant or profile using the built-in search functionality
- 🖥️ **Autopilot Hardware Hash Upload**: Optionally upload the device hardware hash to Microsoft Intune Autopilot before Windows is installed
- ✅ **Profile Validation**: Validate tenant and provisioning settings before starting deployment
- 🌍 **Tenant-Specific Configuration**: Support tenant-specific values such as GroupTag, OS language, edition, build, activation, and optional provisioning scripts
- ☁️ **Flexible Configuration Sources**: Use either hardcoded parameters or centrally hosted configuration files in managed Cloud storage
- 🔐 **Multi-Tenant Graph Authentication**: Authenticate to multiple customer tenants using a multi-tenant Entra ID app registration
- 🧩 **Optional Custom Provisioning Logic**: Support optional tenant-specific provisioning scripts during deployment
- 🪟 **WinPE & OSDCloud Integration**: Designed to run in WinPE and work alongside OSDCloud for bare-metal deployment
- 📋 **Deployment Summary**: Show a clear overview of the selected provisioning settings before handing off to OSDCloud for Windows installation

## ❓ Configuration Choice

Before using **Mutli-Tenant Provisioner**, decide how you want to store the configuration and authentication details.

You can choose between:

### Option 1 - Hardcoded in `Start-MTP.ps1`

All tenant settings and authentication details are stored directly inside the script.

This is suitable when:

- the script is only used internally
- deployment is started from a **centralized Windows Deployment Server**
- there is no need to restrict usage outside your own environment

### Option 2 - Hosted in managed Cloud storage

Configuration files are stored a managed Cloud storage, for example in an Azure blob storage:

- `TenantsConfig.json`
- optional provisioning scripts such as `SetupComplete.ps1`

Authentication secrets can optionally be stored in **Azure Key Vault** instead of inside the script.

This is recommended when:

- you want an additional security layer
- you do not want to maintain Tenant configuration settings directly inside `Start-MTP.ps1`
- you are also using bootable USB sticks

---

## Why host the files in Azure Blob Storage and use Azure Key Vault?

A major benefit of hosting the configuration files in a managed cloud storage is that access can be restricted to a **specific public IP address**.

This means provisioning only works from an approved network location.

### Example

If you are using a **Bootable USB stick** and that USB stick is lost or stolen, the script cannot be used successfully outside the approved location, because:

- the **tenant configuration file**
- the **optional provisioning script**
- and the **authentication secret in Azure Key Vault**

can only be accessed from the allowed public IP address.

This creates an additional **safety net**.

---

## Practical Recommendation

- If you plan to use **Bootable USB sticks** for deployment, hosting the configuration in **Azure Blob Storage** and secrets in **Azure Key Vault** is the safer choice.
- If you only plan to use a **centralized Windows deployment server** in a controlled environment, you may choose to use **hardcoded parameters** in `Start-MTP.ps1` for simplicity.

## 📋 Requirements

### Required for both options

* **Multi-tenant Entra ID Enterprise Application**
The Graph authentication for the hardware hash upload to Intune Autopilot works with a multi-tenant app registration in Entra ID, allowing the same App ID and App Secret to be used across all tenants.

* **Partner Center PowerShell module**
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
    

### Option 1 - Hardcoded parameters

* **Building the Tenants configuration:** Inside the *Start-MTP.ps1* or with *Export-TentansConfig.ps1*
    * Edit Start-MTP.ps1 and go to: *#region: Hardcoded Tenant Parameters* and fill in the parameters of every tenant you want to provision.

    If you only want to provision an OS you can set *UploadToAutopilot* to **$false** and change *Name* to for example **Windows 11 Pro**

    ```powershell
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
    ```

### Option 2 - Hosted in managed Cloud storage


* A (multi-tenant) Entra ID enterprise application in every tenant

* An Azure Key Vault: https://learn.microsoft.com/en-us/azure/key-vault/general/quick-create-portal

* An Azure Blob Storage: https://learn.microsoft.com/en-us/azure/storage/common/storage-account-create?tabs=azure-portal

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
* Create a bootable USB
    ```powershell
    New-OSDCloudUSB
    ```

 * If you make changes to WinPE in your OSDCloud Workspace, you can easily update your OSDCloud USB WinPE volume by using Update-OSDCloudUSB
     ```powershell
    Update-OSDCloudUSB
    ```

## Use with WDS PXE


## 📸 Screenshots

### Tenant Selector

<img width="719" height="564" alt="Image" src="https://github.com/user-attachments/assets/1b94467e-c879-486c-a2eb-b98818f32f51" />

## 🙏 Credits

- OSDCloud: https://github.com/OSDeploy/OSDCloud
- Autopilot uploade logic is based on: https://github.com/blawalt/WinPEAP
- Additional troubleshooting, polishing, and a healthy amount of PowerShell brainstorming by Microsoft 365 Copilot (ChatGPT 5.4) 😉