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



  
