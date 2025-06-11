# ManageExchangeCert (Beta)

🚨 **This project is currently in BETA and not considered production ready.** Features may change, and bugs are expected.

## Overview

ManageExchangeCert is a PowerShell-based CLI tool designed to streamline the management of Microsoft Exchange SSL certificates. This utility provides a menu-driven interface to assist with:

- Renewing Exchange certificates
- Replacing certificates with new versions
- Exporting and importing certificates
- Deleting or revoking old or unused certificates
- Resuming pending certificate renewals
- Propagating changes across multiple Exchange servers
- Automatically detecting and preserving services bound to certificates

## Quick Installation

To install the module and launch the CLI menu:

```powershell
iwr https://raw.githubusercontent.com/r0tifer/Manage_Exchange_Certs/main/Install.ps1 -UseBasicParsing | iex
```

## Features

- Interactive main menu
- Auto-discovery of Exchange servers
- Smart detection and listing of available certificates
- Auto-assignment of current services to new certificates
- Full support for `.cer`, `.crt`, `.pem`, and `.p7b` formats
- Graceful error handling and session recovery

## Project Structure

```
ManageExchangeCert│
├── Public│   ├── Start-CertManager.ps1
│   ├── Renew-Certificate.ps1
│   ├── Replace-Certificate.ps1
│   ├── Export-Certificate.ps1
│   ├── Import-Certificate.ps1
│   ├── Select-ExchangeServer.ps1
│
├── Private│   ├── Show-MainMenu.ps1
│   ├── Prompt-CertSelection.ps1
│   ├── Resume-PendingCertificate.ps1
│   ├── Delete-Certificate.ps1
│   ├── Revoke-Certificate.ps1
│   └── Utility.ps1
│
├── ManageExchangeCert.psm1
└── ManageExchangeCert.psd1
```

## Usage

Once installed, the CLI will launch automatically. You can also launch it manually:

```powershell
Start-CertManager
```

Follow the on-screen menu to manage Exchange certificates with ease.

---

ManageExchangeCert Project | BETA Release