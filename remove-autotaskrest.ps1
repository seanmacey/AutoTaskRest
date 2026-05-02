function Remove-AutotaskRest {
try {
   # Import-Module -Name AutoTaskRest
    $m = Get-Module -Name AutoTaskRest -ListAvailable
}
catch {
    <#Do this if a terminating exception happens#>
}
if ($m) {
    Write-Host "Module AutoTaskRest is found and is being uninstalled." -ForegroundColor Red
    # Unload from session
    Remove-Module -Name AutoTaskRest -Force -ErrorAction SilentlyContinue

    # Find and delete the folder
    $modulePath = Get-Module -Name AutoTaskRest -ListAvailable | Select-Object -ExpandProperty ModuleBase
    Remove-Item -Path $modulePath -Recurse -Force
}
else {
    Write-Host "Module AutoTaskRest is not found - so cannot be uninstalled." -ForegroundColor Green
}
}
Remove-AutotaskRest
# SIG # Begin signature block
# MIIFgwYJKoZIhvcNAQcCoIIFdDCCBXACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAfDiPZZhQzzlA1
# Bxwdp8Q0PJJH/Ro0PZVkx2Xali/XFKCCAv4wggL6MIIB4qADAgECAhA7Wkn363I8
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
# MSIEIKYe0mjsY8ICw37Hx+3wuAHnaWp43NfIQ3dhZZ39KjqtMA0GCSqGSIb3DQEB
# AQUABIIBAMgI5b+qRS0EVqV827EQHhamoPym7WzsDBashf8A75qSYURXkJ0vy88O
# y9y78esGptQzUVzVhgs2qtKxI6Ci353K/zg3ijstKoJ5aermR6KkE/IaKWKewnU6
# DDygUK7GGf9GmRYw0/ByCQDMexzzSk+gMY0waW7U6baLJIQBhfr5aoA46uRkHL3L
# /A5kNvdrPe/x7Ux8A9TtlwW1hIAB8gew4wPGLNav03LCxT0tSVNcst9gYCZdZ4g4
# RxXy8qBR3kPO+0KoY7UM422Vz1Y7hiwFKcenVRQtaCDI/QAV4js8/uAZU0zxdjRA
# W2RGcbosLpTk4fj8HsBRtU773obL/rI=
# SIG # End signature block
