# Multi-Tenant Autopilot Provisioner
WinPE-compatible PowerShell tool to collect Autopilot hardware hashes, select a tenant, and register devices in Microsoft Intune with automatic profile assignment validation.

This tool is designed for MSPs and to be run in combination with OSDCloud in a WinPE environment during the deployment of a Windows device.

The Graph authentication logic is based on a multi-tenant app registration in Entra ID, allowing the same App ID and App Secret to be used across all tenants.

OSDCloud: https://github.com/OSDeploy/OSDCloud

Multi tenant app: https://learningbytesblog.com/posts/Muiltitenant-Entra-APP-for-multitenant-managment/

Autopilot logic used in this tool and OSDCloud USB creation based on: https://github.com/blawalt/WinPEAP


# Install OSDCloud

