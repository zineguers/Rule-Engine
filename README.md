# Rule Engine for Office 365

Clean WPF GUI for auditing and managing inbox rules and SMTP forwarding across O365 mailboxes using Microsoft Graph API.

## Features
- Scan single/multiple mailboxes (or import .txt)
- Detect rules and forwarding settings
- Detailed rule preview
- Delete selected rules with text backup (`C:\Temp\Rule Engine Vault`)
- Bulk-create predefined security rules on selected mailboxes
- Grouped results view

## Prerequisites
- Azure AD app with Graph permissions:
  - `MailboxSettings.Read`
  - `MailboxSettings.ReadWrite` (for delete/create)
- PowerShell 5.1+ on Windows

## Setup & Usage
1. Register Azure AD app → client secret → admin consent
2. Download `Rule Engine.ps1`
3. Unblock file
4. Run:
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File ".\Rule Engine.ps1"
