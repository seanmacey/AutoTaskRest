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

        Write-Host "Hello from PowerShell Installer for AutoTaskRest"

        
        # Create the module folder if it doesn't exist
        if (-not (Test-Path $moduleDir)) {
            New-Item -ItemType Directory -Path $moduleDir -Force | Out-Null
            Write-Host "Created module directory: $moduleDir" -ForegroundColor Cyan
        }


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
        # now install the module
        # ============================================================

        Write-Host "Hello from PowerShell Installer for AutoTaskRest"

        
        # # Create the module folder if it doesn't exist
        # if (-not (Test-Path $moduleDir)) {
        #     New-Item -ItemType Directory -Path $moduleDir -Force | Out-Null
        #     Write-Host "Created module directory: $moduleDir" -ForegroundColor Cyan
        # }

        # Download and save as .psm1
        $destFile = Join-Path $moduleDir "$moduleName.psm1"
        Write-Host "Downloading $moduleName from https://rmm.imatec.co.nz/webdocs/AutoTaskRest/..." -ForegroundColor Cyan

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

# SIG # Begin signature block
# MIIFgwYJKoZIhvcNAQcCoIIFdDCCBXACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBjzri+WW92qbX5
# 8Q7iNEGXATPvr49sovVefoYtHBOUCKCCAv4wggL6MIIB4qADAgECAhA7Wkn363I8
# sU+TMwSw335wMA0GCSqGSIb3DQEBCwUAMBUxEzARBgNVBAMMClNlYW4gTWFjZXkw
# HhcNMjYwNDA0MjIxNTM5WhcNMzYwNDA0MjIyNTM5WjAVMRMwEQYDVQQDDApTZWFu
# IE1hY2V5MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA6/RB8ks87nu6
# LqUgWXq02tdjYt427XKkEW7vFYFqr35woedz7nUwIgMcyDmbiTtOdzDAFJl4ld3/
# TJEVeyndCqePz+LsXRBk3nDxhouuh+ORnyn7ga3FFwp7jSmTiTr/LWMy8gZqhsvU
# sBCQWPA6OaJy8x0iGAjkKqWjwiO8lepPHR9MeTuRsiVI0GYbxdyf+2If8Lhhqq7R
# BwaNhTTvYjDGG95VaaIOngPYxDnz1UsWjLiCA0vrq+ZEeiT/gOvtAzRrH6NMZHVE
# JekVhuByAreI9StjTwyzmiIwZhK95vwHVaXpF4OXFzSpneGihJPeoU/M9PToeJnm
# EHw7rIWQrQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYB
# BQUHAwMwHQYDVR0OBBYEFC3Zqjwt22ikPMgP/7MK7ULso51/MA0GCSqGSIb3DQEB
# CwUAA4IBAQB6jHCzFEeK/R1TwNZotJtRIJX67GTeQvY/LmLuLCo74td0rzIUddp/
# DmZWszqlNnEcnNkRnWJ1A07ge7FTn51biripsHSxX7f6xfSc/5HbcUm9diANjYXV
# 18hEeSc2E0Yw2Xz1HI35owaQZotWZX9I7CKLiCXfOEEtWgbS/+Ff7PxQ7C60zwP+
# OSmthwdUeSeDPnSr6IXnTQ0/DKlEMW1wFfhinGvT20J/dJQxlm66vE4WfKDrrDln
# TAQaVWe5CvhZ1q84AV8o5zz13mO3HWJ+2+2bqj1+CYVwSqXtaYbuALVQTJSBaUU5
# DV9gdv/aK0f8k5TkeFr+S598G2l84JB9MYIB2zCCAdcCAQEwKTAVMRMwEQYDVQQD
# DApTZWFuIE1hY2V5AhA7Wkn363I8sU+TMwSw335wMA0GCWCGSAFlAwQCAQUAoIGE
# MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQB
# gjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkE
# MSIEIKOv+QOh/B92PvihPNcFI266s6Fs3Nd07FOU53+7xNGvMA0GCSqGSIb3DQEB
# AQUABIIBAG8Z6Jsw8m2WjxqFC41JwioxtTMzhLj3nOoJ53n4Y7XxazurdoZyQtwm
# Kf/wF/U9MehCsY7tS0LvPENiImaOzlBcPKFeke8aBFBU6qFdpITKWo/HfQjTRzxJ
# UqmWxNFNx3+JbxJSh/g+pfF3WsnjUn3u7TxtE3c0doraiKoMCeg3xZmtodTI5SLi
# UKvWs7wVnIVi6+edFXjw1N5boIhDR0r2ABPEYvARx5ZUZdRLlUcN0BP7ilLCSW6D
# 1dr8W08kJJXjBsz+5Wczaqqxxdkg5k4g2oqelbaU1rSiBeC5G70/FBr/HDANYArk
# H96BPJRa1f9W2HrmKrcBUOBJ3cqsgLc=
# SIG # End signature block
