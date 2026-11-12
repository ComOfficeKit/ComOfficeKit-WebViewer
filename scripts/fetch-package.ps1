# Puts the ComOffice .NET package of a release of ComOffice for Windows and Linux in packages/, where nuget.config
# looks for it, in place of any package there before.
#
#   scripts/fetch-package.ps1                  the latest release
#   scripts/fetch-package.ps1 -Version 1.6.0   that release
#
# GH_TOKEN or GITHUB_TOKEN, when set, is sent to GitHub's API, so a CI runner sharing an address stays under the
# anonymous rate limit.
param([string]$Version)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
# Windows PowerShell 5.1 may still offer only TLS 1.0/1.1 by default; GitHub needs 1.2.
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$repo = 'ComOfficeKit/ComOfficeKit-Windows'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
if (-not $Version) {
    $headers = @{ 'User-Agent' = 'comofficekit-viewer-fetch-package'; 'Accept' = 'application/vnd.github+json' }
    $token = if ($env:GH_TOKEN) { $env:GH_TOKEN } else { $env:GITHUB_TOKEN }
    if ($token) { $headers['Authorization'] = "Bearer $token" }
    $latest = Invoke-RestMethod -UseBasicParsing -Uri "https://api.github.com/repos/$repo/releases/latest" -Headers $headers
    $Version = $latest.tag_name -replace '^v', ''
}
$packages = Join-Path $root 'packages'
New-Item -ItemType Directory -Force $packages | Out-Null
$file = Join-Path $packages "ComOfficeKit.$Version.nupkg"
Invoke-WebRequest -UseBasicParsing -Uri "https://github.com/$repo/releases/download/v$Version/ComOfficeKit.$Version.nupkg" -OutFile "$file.part"
Get-ChildItem $packages -Filter 'ComOfficeKit.*.nupkg' | Remove-Item -Force
Move-Item "$file.part" $file
Write-Output $file
