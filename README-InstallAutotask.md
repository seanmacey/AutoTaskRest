# AutoTaskRest

A PowerShell 7 module for interacting with the [Autotask PSA REST API](https://webservices6.autotask.net/ATServicesRest/swagger/ui/index).  
Developed by [Kiss IT](https://kissit.co.nz).

---

## Requirements

| Requirement | Minimum version |
|-------------|----------------|
| PowerShell | 7.0 |
| Autotask API credentials | API user + password + Integration Code |
| Internet access | Required to reach Autotask REST endpoints |

> **Note:** The module uses Windows DPAPI to encrypt saved credentials.  
> This means credentials saved on one Windows account **cannot** be decrypted by another account or on another machine.

---

## Installation

### Option A — Automated (recommended)

Copy and run the following script in a **PowerShell 7** session.  
It downloads the module from the KISS IT GitLab repository and installs it into your user module path.
if required it will also install the self signed certificate for Powershell scripting

* this script is also available as **```https://rmm.imatec.co.nz/webdocs/AutoTaskRest/install-autotaskrest.ps1```**

```powershell

<#
    Auto-Install and Auto-Switch to PowerShell 7 Script
    Works when launched from Windows PowerShell 5.1
#>
$moduleUrl  = 'https://rmm.imatec.co.nz/webdocs/AutoTaskRest/AutoTaskRest.psm1'
$moduleName = 'AutoTaskRest'
$seanspubliccerturl = 'https://rmm.imatec.co.nz/webdocs/AutoTaskRest/SeanMacey-CodeSigning-public.cer'

# Resolve the correct PS7 user module directory
$moduleRoot = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules'
$moduleDir  = Join-Path $moduleRoot $moduleName

function Test-RequiredPermissions {
    $issues = @()

    # Check 1: Can we write to the user-scope module path?
    $userModulePath = ($env:PSModulePath -split ';') |
    Where-Object { $_ -like "*$env:USERPROFILE*" } |
    Select-Object -First 1

    if (-not $userModulePath) {
        $issues += "Cannot resolve a user-scoped module path from PSModulePath."
    }
    else {
        try {
            $testFile = Join-Path $userModulePath ".permtest_$(New-Guid)"
            $null = New-Item -Path $testFile -ItemType File -Force -ErrorAction Stop
            Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
        }
        catch {
            $issues += "No write access to user module path: $userModulePath"
        }
    }

    # Check 2: Can we write to the certificate stores?
    foreach ($store in @('Cert:\LocalMachine\Root', 'Cert:\LocalMachine\TrustedPublisher')) {
        try {
            $certStore = Get-Item $store -ErrorAction Stop
            $rawStore = New-Object System.Security.Cryptography.X509Certificates.X509Store(
                $certStore.Name,
                [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine
            )
            $rawStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
            $rawStore.Close()
        }
        catch {
            $issues += "No write access to certificate store: $store (requires elevation)"
        }
    }

    return $issues
}

function Get-PwshPath {
    $cmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $default = "C:\Program Files\PowerShell\7\pwsh.exe"
    if (Test-Path $default) { return $default }

    return $null
}

function Install-PowerShell7 {
    Write-Host "Downloading latest PowerShell 7 MSI installer..."

    $temp = Join-Path $env:TEMP "PowerShell7-latest.msi"

    $release = Invoke-RestMethod "https://api.github.com/repos/PowerShell/PowerShell/releases/latest"
    $asset = $release.assets | Where-Object { $_.name -match "win-x64\.msi$" }

    if (-not $asset) {
        Write-Host "Could not find MSI asset in release metadata." -ForegroundColor Red
        exit 1
    }

    Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $temp

    Write-Host "Installing PowerShell 7 silently..."
    Start-Process "msiexec.exe" -ArgumentList "/i `"$temp`" /qn /norestart" -Wait

    Remove-Item $temp -Force
    Write-Host "PowerShell 7 installation complete." -ForegroundColor Green
}
function install-it {

    # ============================================================
    # ENTRY POINT
    # ============================================================

    # Check if running under Windows PowerShell 5.1
    if ($PSVersionTable.PSEdition -eq "Desktop") {
        Write-Host ""
        Write-Host "This script must be run under PowerShell 7." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  1. Press Win+R" -ForegroundColor Cyan
        Write-Host "  2. Type: pwsh" -ForegroundColor Cyan
        Write-Host "  3. Right-click pwsh in the Start Menu and choose 'Run as Administrator'" -ForegroundColor Cyan
        Write-Host "  4. Navigate to: $PSScriptRoot" -ForegroundColor Cyan
        Write-Host "  5. Run: .\$($MyInvocation.MyCommand.Name)" -ForegroundColor Cyan
        Write-Host ""

        # Check if PS7 is even installed, and offer to install it before the user goes looking for it
        $pwshPath = Get-PwshPath
        if (-not $pwshPath) {
            Write-Host "PowerShell 7 does not appear to be installed." -ForegroundColor Yellow
            $choice = Read-Host "Install it now? (Y/N)"
            if ($choice -match '^[Yy]$') {
                Install-PowerShell7
            }
        }

        Read-Host "Press Enter to exit"
        # return
    }
    else {

        # From here we are guaranteed to be in PowerShell 7+
        Write-Host "Running under PowerShell $($PSVersionTable.PSVersion)" -ForegroundColor Green

        # ============================================================
        # PERMISSION CHECK
        # ============================================================




        #check if certificatre already exists, and if not then add
        $i = Get-ChildItem Cert:CurrentUser\TrustedPublisher | Where-Object { $_.Thumbprint -eq "DAFB37056F2C9900A0E8B1FDB62E2F75C21F73CE" }
        $u = Get-ChildItem Cert:CurrentUser\Root | Where-Object { $_.Thumbprint -eq "DAFB37056F2C9900A0E8B1FDB62E2F75C21F73CE" }
        if ((-not $i) -or (-not $u)) {
            $permissionIssues = Test-RequiredPermissions
            if ($permissionIssues.Count -gt 0) {
                Write-Host ""
                Write-Host "[PERMISSION CHECK FAILED]" -ForegroundColor Red
                $permissionIssues | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
                Write-Host ""
                Write-Host "Please re-run this script as Administrator." -ForegroundColor Cyan
                Write-Host "Right-click pwsh -> 'Run as Administrator', then re-run the script." -ForegroundColor Cyan
                Write-Host ""
                Read-Host "Press Enter to exit"
                exit 1
            }
        }

        if (-not $i) {
            try {
                $destFile = Join-Path $moduleDir "cert.cer"
                Invoke-WebRequest -Uri $seanspubliccerturl -OutFile $destFile -UseBasicParsing
                Import-Certificate  -FilePath $destFile -CertStoreLocation Cert:\CurrentUser\TrustedPublisher
                Write-Host "Cert saved to: TrustedPublisher" -ForegroundColor Green
            }
            catch {
                Write-Host "crt download fto trustedpublisher failed: $($_.Exception.Message)" -ForegroundColor Red
                exit 1
            }
        }

        if (-not $u) {
            try {
                $destFile = Join-Path $moduleDir "cert.cer"
                Invoke-WebRequest -Uri $seanspubliccerturl -OutFile $destFile -UseBasicParsing
                Import-Certificate  -FilePath $destFile -CertStoreLocation Cert:\CurrentUser\Root
                Write-Host "Cert saved to: Root" -ForegroundColor Green
            }
            catch {
                Write-Host "crt download to root failed: $($_.Exception.Message)" -ForegroundColor Red
                

                exit 1
            }
        }

        Write-Host "[OK] Required permissions confirmed." -ForegroundColor Green
        Write-Host ""

        # ============================================================
        # YOUR POWERSHELL 7 CODE HERE
        # ============================================================

        Write-Host "Hello from PowerShell Installer for AutoTaskRest"

        
        # Create the module folder if it doesn't exist
        if (-not (Test-Path $moduleDir)) {
            New-Item -ItemType Directory -Path $moduleDir -Force | Out-Null
            Write-Host "Created module directory: $moduleDir" -ForegroundColor Cyan
        }

        # Download and save as .psm1
        $destFile = Join-Path $moduleDir "$moduleName.psm1"
        Write-Host "Downloading $moduleName from GitLab..." -ForegroundColor Cyan

        try {
            Invoke-WebRequest -Uri $moduleUrl -OutFile $destFile -UseBasicParsing
            Write-Host "Module saved to: $destFile" -ForegroundColor Green
        }
        catch {
            Write-Host "Download failed: $($_.Exception.Message)" -ForegroundColor Red
            read-host "Press Enter to exit"
            exit 1
        }

        # Set execution policy for current user if needed
        $policy = Get-ExecutionPolicy -Scope CurrentUser
        if ($policy -notin 'RemoteSigned', 'Unrestricted', 'Bypass') {
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
            Write-Host "Execution policy set to RemoteSigned for current user." -ForegroundColor Yellow
        }

        # Import and verify
        Import-Module $moduleName -Force
        if (Get-Module -Name $moduleName) {
            Write-Host "`nAutoTaskRest module installed successfully!" -ForegroundColor Green
            Write-Host "Run 'Set-ATLogin' to configure your Autotask API credentials." -ForegroundColor Cyan
        }
        else {
            Write-Host "Module import failed. Check the file at: $destFile" -ForegroundColor Red
        }

        # ============================================================
        # DONE
        # ============================================================

        Write-Host ""
        Write-Host "Execution complete." -ForegroundColor Green
        Read-Host "Press Enter to close this window"
    }
}
install-it


```

---

### Option B — Manual Installation

**Step 1:** Download the file from GitLab (you must be authenticated):
or from RMM delivered copy @ ``https://rmm.imatec.co.nz/webdocs/AutoTaskRest/``


**Step 2:** Create the module folder:

```powershell
$moduleDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules\AutoTaskRest'
New-Item -ItemType Directory -Path $moduleDir -Force
```

**Step 3:** Copy the downloaded file into that folder and rename it:

```powershell
# Replace C:\Users\YourName\Downloads\AutoTaskRest.ps1 with your actual download path
Copy-Item 'C:\Users\YourName\Downloads\AutoTaskRest.ps1' `
          (Join-Path $moduleDir 'AutoTaskRest.psm1') -Force
```

**Step 4:** Set execution policy (if not already done):

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**Step 5:** Import the module:

```powershell
Import-Module AutoTaskRest
```

---

### Option C — Update an Existing Installation

Re-run Option A, or run this one-liner to pull the latest version and reload:

```powershell
$dest = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules\AutoTaskRest\AutoTaskRest.psm1'
Invoke-WebRequest -Uri 'https://gitlab.kissit.co.nz/kiss/autotaskrest/-/raw/main/AutoTaskRest.psm1' `
                  -OutFile $dest -UseBasicParsing


Import-Module AutoTaskRest -Force
Write-Host "AutoTaskRest updated and reloaded." -ForegroundColor Green
```

---

## First-Time Setup

After installing the module, you must configure your Autotask API credentials.  
Run this **once** per machine/user account:

```powershell
Import-Module AutoTaskRest
Set-ATLogin
```

You will be prompted for:

| Prompt | What to enter |
|--------|--------------|
| API username | Your Autotask API user email address |
| Password | The API user's password (entered as a SecureString — not echoed) |
| AT-API-ID | Your Autotask API Integration Code |

Credentials are encrypted using Windows DPAPI and saved to:  
`$HOME\kiss-atapi\kissAtapilogin.json`

They are automatically loaded on every subsequent call — you do not need to log in again.

**Optionally**, save which engineer's data to use as the default for reporting:

```powershell
Set-ATEngineerToReport -email "you@yourcompany.com"
```

---

## Auto-Import on Every Session

To have the module load automatically every time you open PowerShell 7, add it to your profile:

```powershell
# Open your PowerShell 7 profile in your editor
notepad $PROFILE

# Add this line to the profile file:
Import-Module AutoTaskRest
```

---

## Verifying the Installation

```powershell
Import-Module AutoTaskRest

# List all available commands
Get-Command -Module AutoTaskRest

# Test your saved credentials
Test-ATConnection

# Quick smoke test — fetch classification icons
Invoke-ATQuery -entityName 'v1.0/ClassificationIcons' -includeFields "id", "name"
```

---

## Common Commands

| Command | Description |
|---------|-------------|
| `Set-ATLogin` | Save/update Autotask API credentials |
| `Test-ATConnection` | Verify saved credentials work |
| `Set-ATEngineerToReport` | Set the default engineer for reporting |
| `Get-ATCompanies` | Retrieve company records |
| `Get-ATTickets` | Retrieve tickets |
| `Get-ATTimeEntries` | Retrieve time entries |
| `Get-ATEngineers` | Retrieve resource / engineer records |
| `Get-ATCustomerReport` | Full customer time/billing report (console + optional HTML) |
| `Get-ATTicketReport` | Full ticket report with notes and time entries |
| `Export-ATTimeRecords` | Export time data to CSV files (for Power BI etc.) |
| `Export-ATCompanies` | Export company data to CSV |
| `Sync-AT365Calendar` | Sync Autotask time entries to Microsoft 365 calendar |

Use `Get-Help <CommandName> -Full` for complete parameter documentation on any command.

---

## Examples

```powershell
# Get all active companies
Get-ATCompanies

# Get a specific company by ID
Get-ATCompanies -id 29762985

# Get companies by name (partial match)
Get-ATCompanies -CompanyName "Acme"

# Get time entries for the last 2 weeks (for currently logged-in engineer)
Get-ATTimeEntries -LastXWeeks 2 -ForMeOnly

# Get time entries for the last month
Get-ATTimeEntries -LastxMonths 1

# Generate a customer report for a company (console output + HTML)
Get-ATCustomerReport -CompanyID 29762985 -LastxMonths 1 -OpenInBrowser

# Generate a ticket report
Get-ATTicketReport -TicketNumbers 'T20260101.0001'

# Export all time records for last 3 months to CSV
Export-ATTimeRecords -LastxMonths 3 -path "C:\Reports"
```

---

## Credential Storage

Credentials are stored at:  
```
$HOME\kiss-atapi\kissAtapilogin.json
```

The JSON file contains:
- `UserName` — plain text (the API email address)
- `Secret` — DPAPI-encrypted password string
- `atapi` — DPAPI-encrypted API Integration Code
- `url` — the Autotask regional API base URL (e.g. `https://webservices6.autotask.net/ATServicesRest/`)

> ⚠️ The encrypted values are machine- and user-account-specific. They **cannot** be copied to another machine or used under a different Windows account.

---

## Troubleshooting

**"No global Autotask API is currently set"**  
Run `Set-ATLogin` to configure credentials.

**"Connection test failed"**  
- Verify your API username, password, and Integration Code in Autotask.  
- Ensure your API user has the required permissions in Autotask.  
- Check network connectivity to `webservices.autotask.net`.

**Module not found after installation**  
- Confirm you are running **PowerShell 7** (`pwsh`), not Windows PowerShell 5 (`powershell`).  
- Check the module exists: `Test-Path "$([Environment]::GetFolderPath('MyDocuments'))\PowerShell\Modules\AutoTaskRest\AutoTaskRest.psm1"`

**Execution policy error**  
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## License

Proprietary — Kiss IT Ltd. For internal use only.
