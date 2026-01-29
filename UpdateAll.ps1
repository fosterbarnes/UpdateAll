if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Start-Process pwsh -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

$Host.UI.RawUI.WindowTitle = "UpdateAll"
. $PROFILE; Center-PSWindow

function Run-Task {
    param (
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory)]
        [ScriptBlock]$Action
    )

    $bar = "=" * $Host.UI.RawUI.WindowSize.Width
    Write-Host "$Message..." -ForegroundColor Yellow && Write-Host $bar
    & $Action
    Write-Host "Done.`n$bar`n`n"
}


Run-Task "Updating Windows" {
    Import-Module PSWindowsUpdate
    Get-WindowsUpdate -MicrosoftUpdate -Install -AcceptAll
}

Run-Task "Updating NVIDIA drivers" {
    powershell.exe -ExecutionPolicy Bypass -File "C:\Users\Foster\Documents\Powershell Scripts\nvidiaNEW.ps1" # can't detect drivers unless win powershell is used
}

Run-Task "Updating Python" { 
    $currentPythonVersion = $null
    try {
        if ((python --version 2>&1) -match '(\d+\.\d+\.\d+)') {
            $currentPythonVersion = $matches[1]
        }
    } catch { }

    $latestPythonURL = (Invoke-WebRequest https://www.python.org/downloads/windows/).Links | 
        Where-Object href -match 'python-(\d+\.\d+\.\d+)-amd64\.exe$' | 
        Select-Object -First 1
    $latestPythonVersion = if ($latestPythonURL.href -match '(\d+\.\d+\.\d+)') { $matches[1] }

    if ($currentPythonVersion -eq $latestPythonVersion) {
        Write-Host "Python $currentPythonVersion is already up to date" -ForegroundColor Green
        return
    }

    Write-Host "Updating from $currentPythonVersion to $latestPythonVersion"
    $pythonInstallerPath = "$env:TEMP\python-$latestPythonVersion-amd64.exe"
    Invoke-WebRequest ($latestPythonURL.href -replace '^/', 'https://www.python.org/') -OutFile $pythonInstallerPath
    Start-Process $pythonInstallerPath -ArgumentList "/passive", "InstallAllUsers=1", "PrependPath=1", "Include_pip=1", "Include_launcher=1", "AssociateFiles=1", "Shortcuts=1", "Include_test=0", "TargetDir=C:\Python" -Wait

    if (Test-Path $pythonInstallerPath) {
        Remove-Item $pythonInstallerPath -Force
    }
}

Run-Task "Updating VSCode" { 
    winget upgrade --id Microsoft.VisualStudioCode --accept-package-agreements --accept-source-agreements 
}

Run-Task "Updating Visual Studio" {
    dotnet tool update -g dotnet-vs
    vs update --all
}

Run-Task "Updating Cursor" {
    $installedCursorVersion = (cursor --version 2>&1 | Select-Object -First 1).Trim()
    $cursorVersionHistory = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/oslook/cursor-ai-downloads/main/version-history.json"
    $latestCursorVersion = $cursorVersionHistory.versions[0]
    $cursorInstallerPath = "$env:TEMP\CursorSetup.exe"

    if ($installedCursorVersion -eq $latestCursorVersion.version) {
        Write-Host "Cursor $installedCursorVersion is already up to date" -ForegroundColor Green
        return
    }

    Write-Host "Updating from $installedCursorVersion to $($latestCursorVersion.version)"
    Invoke-WebRequest -Uri $latestCursorVersion.platforms."win32-x64-user" -OutFile $cursorInstallerPath
    Start-Process -FilePath $cursorInstallerPath -ArgumentList "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/SP-" -Wait

    if (Test-Path $cursorInstallerPath) {
        Remove-Item $cursorInstallerPath -Force
    }
}

Run-Task "Updating Audacity" {
    $installedAudacityVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Audacity_is1").DisplayVersion
    $audacityRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/audacity/audacity/releases/latest" -Headers @{ 'User-Agent' = 'PowerShell' }
    $audacityInstaller = $audacityRelease.assets | Where-Object { $_.name -match "audacity-win-.*-64bit\.exe$" }
    $latestAudacityVersion = $audacityRelease.tag_name -replace '^(Audacity-|v)', ''
    $audacityInstallerPath = "$env:TEMP\$($audacityInstaller.name)"

    if ($installedAudacityVersion -eq $latestAudacityVersion) {
        Write-Host "Audacity $installedAudacityVersion is already up to date" -ForegroundColor Green
        return
    }

    Write-Host "Updating from $installedAudacityVersion to $latestAudacityVersion"
    Invoke-WebRequest -Uri $audacityInstaller.browser_download_url -OutFile $audacityInstallerPath
    Start-Process -FilePath $audacityInstallerPath -ArgumentList "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/SP-" -Wait

    if (Test-Path $audacityInstallerPath) {
        Remove-Item $audacityInstallerPath -Force
    }
}

Run-Task "Updating balenaEtcher" {
    $uninstallKeys = Get-ChildItem "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction SilentlyContinue
    foreach ($key in $uninstallKeys) {
        $app = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
        if ($app.DisplayName -like "*balenaEtcher*" -or $app.DisplayName -like "*Etcher*") {
            $installedEtcherVersion = $app.DisplayVersion
            break
        }
    }

    $etcherRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/balena-io/etcher/releases/latest" -Headers @{ 'User-Agent' = 'PowerShell' }
    $etcherInstaller = $etcherRelease.assets | Where-Object { $_.name -match "balenaEtcher-.*\.Setup\.exe$" }
    $latestEtcherVersion = $etcherRelease.tag_name -replace '^v', ''
    $etcherInstallerPath = "$env:TEMP\$($etcherInstaller.name)"

    if ($installedEtcherVersion -eq $latestEtcherVersion) {
        Write-Host "balenaEtcher $installedEtcherVersion is already up to date" -ForegroundColor Green
        return
    }

    Write-Host "Updating from $installedEtcherVersion to $latestEtcherVersion"
    Invoke-WebRequest -Uri $etcherInstaller.browser_download_url -OutFile $etcherInstallerPath
    Start-Process -FilePath $etcherInstallerPath -ArgumentList "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/SP-"
    WaitForProgram "balenaEtcher" && Start-Sleep -Seconds 1.5
    Write-Host "Killing balenaEtcher..." && Stop-Process -Name "balenaEtcher" -Force

    if (Test-Path $etcherInstallerPath) {
        Remove-Item $etcherInstallerPath -Force
    }
}

Run-Task "Updating Bulk Crap Uninstaller" {
    $bcuUninstallerLog = "D:\Users\Foster\Documents\Applications\BCUninstaller\win-x64\BCUninstaller.log"
    $bcuDir = "D:\Users\Foster\Documents\Applications\BCUninstaller"
    $currentBcuVersion = $null

    $logContent = Get-Content $bcuUninstallerLog -Tail 500  # Read last 500 lines
    foreach ($line in $logContent) {
        if ($line -match 'Bulk Crap Uninstaller v([\d.]+)') {
            $currentBcuVersion = $matches[1]
            break
        }
    }

    $bcuRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/Klocman/Bulk-Crap-Uninstaller/releases/latest" -Headers @{ 'User-Agent' = 'PowerShell' }
    $latestBcuVersion = $bcuRelease.tag_name -replace '^v', ''
    $bcuPortable = $bcuRelease.assets | Where-Object { $_.name -match "BCUninstaller_.*_portable\.zip$" }

    if ($currentBcuVersion -eq $latestBcuVersion) {
        Write-Host "BCUninstaller $currentBcuVersion is already up to date" -ForegroundColor Green
        return
    }
    
    Write-Host "Updating from $currentBcuVersion to $latestBcuVersion"
    $bcuZipPath = "$env:TEMP\$($bcuPortable.name)"
    Invoke-WebRequest -Uri $bcuPortable.browser_download_url -OutFile $bcuZipPath
    Expand-Archive -Path $bcuZipPath -DestinationPath $bcuDir -Force

    if (Test-Path $bcuZipPath) {
        Remove-Item $bcuZipPath -Force
    }
}

# Blender
# Brave
# Bridge
# Bulk rename utility
# Chatterino
# Citra
# CPU-Z
# DB Browser
# DDU
# DS4Windows
# Dolphin
# Elgato Camera Hub
# Everything
# FanControl
# Flow Launcher
# Geek Uninstaller
# Ghidra
# Git
# GitHub Desktop
# Google Chrome
# GrepWin
# HandBrake
# HWMonitor
# Inkscape
# Jackett
# Java
# JDownloader
# Last.fm Scrubbler
# JoyToKey
# Lumia Stream
# mGBA
# Minecraft
# Microsoft Store
# MixItUp
# MKVToolNix
# Moonscraper
# MP3Tag
# MPluginManager
# MSI Afterburner
# MTGA
# NVidia Profile Inspector
# O&O Shutup
# OBS
# Onyx
# ParkControl
# PDFGear
# PowerShell 7
# PowerToys
# Proton VPN
# Putty
# Rainmeter
# Rare
# Reaper
# RegCool
# rpcs3
# rufus
# rustdesk
# rustdesk server
# ShareX
# SoulSeek
# Spotify
# Streamer.bot
# Sudachi
# SysInternals Suite
# TouchPortal
# TreeSizeFree
# UVR
# VLC
# test
# VMWare Workstation Pro
# WinSCP

Run-Task "Updating YARG Nightly" {
    $dApplications = "D:\Users\Foster\Documents\Applications"
    $yargNightlyDir = "$dApplications\YARG Nightly"
    $yargVersionFile    = Join-Path $yargNightlyDir "version.txt"
    $yargVoxDir = "$yargNightlyDir\YARG_Data\StreamingAssets\vox"
    $fullComboOPUS     = "$yargVoxDir\FullCombo.opus"
    $highScoreOPUS     = "$yargVoxDir\HighScore.opus"
    $fullComboOPUSNew  = "$yargVoxDir\FullCombo.opus.new"
    $highScoreOPUSNew  = "$yargVoxDir\HighScore.opus.new"

    $yargRelease = Invoke-RestMethod `
        -Uri "https://api.github.com/repos/YARC-Official/YARG-BleedingEdge/releases/latest" `
        -Headers @{ "User-Agent" = "PowerShell" }

    $latestYargVersion = $yargRelease.tag_name.TrimStart("v")
    
    if (Test-Path $yargVersionFile) {
        $currentYargVersion = Get-Content $yargVersionFile -ErrorAction SilentlyContinue
        if ($currentYargVersion -eq $latestYargVersion) {
        Write-Host "YARG $currentYargVersion is already up to date" -ForegroundColor Green
        return
    }
        Write-Host "Installed version: $currentYargVersion"
    }

    $yargWinReleaseAsset = $yargRelease.assets | Where-Object { $_.name -match "Windows-x64\.zip$" }
    $yargZipPath = Join-Path $env:TEMP $yargWinReleaseAsset.name
    $yargExtract = Join-Path $env:TEMP "YARG-Nightly-Extract"

    Write-Host "Latest YARG version: $latestYargVersion" && Write-Host "Downloading..."
    Invoke-WebRequest -Uri $yargWinReleaseAsset.browser_download_url -OutFile $yargZipPath
    Remove-Item $yargExtract -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $yargExtract | Out-Null
    & 7z x $yargZipPath "-o$yargExtract" -y | Out-Null
    Write-Host "Copying files..."
    Copy-Item "$yargExtract\*" $yargNightlyDir -Recurse -Force
    Remove-Item $fullComboOPUS, $highScoreOPUS -ErrorAction SilentlyContinue
    Copy-Item $fullComboOPUSNew $fullComboOPUS -Force
    Copy-Item $highScoreOPUSNew $highScoreOPUS -Force
    Set-Content -Path $yargVersionFile -Value $latestYargVersion -Encoding UTF8
    Write-Host "Cleaning up..."
    Remove-Item $yargZipPath -Force
    Remove-Item $yargExtract -Recurse -Force
}

Pause