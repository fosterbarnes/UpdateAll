if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Start-Process pwsh -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

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

Run-Task "Updating NVIDIA drivers" {
    powershell.exe -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "modules\UpdateNvidiaDrivers.ps1") # can't detect drivers unless win powershell is used
}




Pause