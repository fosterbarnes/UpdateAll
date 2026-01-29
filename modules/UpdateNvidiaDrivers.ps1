# Downloads and runs the latest NVIDIA Game Ready Driver interactively
. $PROFILE; Center-PSWindow

function Split-String (
    [Parameter(Mandatory = $True)][string]$String, 
    [Parameter(Mandatory = $True)][string]$Delimiter,
    [int]$MaxSubStrings = 0) {
    return $String -Split "$Delimiter", $MaxSubStrings, "SimpleMatch"
}

function Expand-NvidiaDriverPackage (
    [Parameter(Mandatory = $True)]$DriverPackage,
    [ValidateSet("Launch", "Install", "Open", $Null)]
    [string]$Post,
    [switch]$All,
    [array]$Components = @()) {
    $ComponentsFolders = "Display.Driver NVI2 EULA.txt ListDevices.txt setup.cfg setup.exe"
    $PostArguments, $Wait = " ", $False
    $DriverPackage = (Resolve-Path $DriverPackage)
    $Output = (Split-String $DriverPackage (Get-Item $DriverPackage).Extension 2)[0]
    $7Zip = "$ENV:TEMP\7zr.exe"
    $PostCfg = "$Output\setup.cfg"
    $PresentationsCfg = "$Output/NVI2/presentations.cfg"
    if ($All) {
        Write-Output "Extraction Options: All Driver Components" 
        $ComponentsFolders = "" 
    }
    elseif ($Components -and !$All) {
        Write-Output "Extraction Options: Display Driver | $($Components -Join " | ")"
        $Components | ForEach-Object {
            switch ($_) {
                "PhysX" { $ComponentsFolders += " $_" }
                "HDAudio" { $ComponentsFolders += " $_" }
                default { Write-Error "Invalid Component." -ErrorAction Stop } 
            }
        }
    }

    Write-Output "Extracting: `"$DriverPackage`""
    Write-Output "Extraction Directory: `"$Output`""
    Remove-Item $Output -Recurse -Force -ErrorAction SilentlyContinue
    (New-Object System.Net.WebClient).DownloadFile("https://www.7-zip.org/a/7zr.exe", $7Zip)
    Invoke-Expression "& `"$7Zip`" x -bso0 -bsp1 -bse1 -aoa `"$DriverPackage`" $ComponentsFolders -o`"$Output`"" 

    $PostCfgContent = [System.Collections.ArrayList](Get-Content $PostCfg -Encoding Ascii)
    foreach ($Index in 0..($PostCfgContent.Count - 1)) {
        if ($PostCfgContent[$Index].Trim() -in @('<file name="${{EulaHtmlFile}}"/>', 
                '<file name="${{FunctionalConsentFile}}"/>'
                '<file name="${{PrivacyPolicyFile}}"/>')) { 
            $PostCfgContent[$Index] = "" 
        }
    }
    Set-Content $PostCfg $PostCfgContent -Encoding Ascii

    $PresentationsCfgContent = [System.Collections.ArrayList](Get-Content $PresentationsCfg -Encoding Ascii)
    foreach ($Index in 0..($PresentationsCfgContent.Count - 1)) {
        foreach ($String in @('<string name="ProgressPresentationUrl" value=',
                '<string name="ProgressPresentationSelectedPackageUrl" value=')) {
            if ($PresentationsCfgContent[$Index] -like "`t`t$String*") {
                $PresentationsCfgContent[$Index] = "`t`t$String`"`"/>"
            }
        }
    }
    Set-Content $PresentationsCfg $PresentationsCfgContent -Encoding Ascii

    Write-Output "Finished: The specified Driver Package has been Extracted."
}

# Get user's Downloads folder correctly
$shell = New-Object -ComObject Shell.Application
$Downloads = $shell.Namespace('shell:Downloads').Self.Path

# Detect current installed NVIDIA driver version
Write-Host "Detecting currently installed NVIDIA driver..."
try {
    $VideoController = Get-WmiObject -ClassName Win32_VideoController | Where-Object { $_.Name -match "NVIDIA" }
    if ($VideoController) {
        # Extract version in form like 576.52 from DriverVersion string
        $ins_version = ($VideoController.DriverVersion.Replace('.', '')[-5..-1] -join '').Insert(3, '.')
        Write-Host "Installed version:`t$ins_version"
    }
    else {
        throw "NVIDIA GPU not found."
    }
}
catch {
    Write-Host -ForegroundColor Yellow "Unable to detect an installed NVIDIA GPU or driver version."
    $ins_version = $null
}

# Get latest NVIDIA driver version for Windows 10/11 64-bit, RTX 3080
$uri = 'https://gfwsl.geforce.com/services_toolkit/services/com/nvidia/services/AjaxDriverService.php' +
    '?func=DriverManualLookup' +
    '&psid=120' +  # GeForce RTX 30 Series
    '&pfid=929' +  # RTX 3080
    '&osID=57' +   # Windows 10 64-bit
    '&languageCode=1033' +
    '&isWHQL=1' +
    '&dch=1' +
    '&sort1=0' +
    '&numberOfResults=1'

try {
    $response = Invoke-WebRequest -Uri $uri -Method GET -UseBasicParsing
    $payload = $response.Content | ConvertFrom-Json
    $latestVersion = $payload.IDS[0].downloadInfo.Version
    Write-Host "Latest version:`t`t$latestVersion"
}
catch {
    Write-Error "Failed to get latest NVIDIA driver version."
    exit
}

# Compare with installed version
if ($ins_version -and $ins_version -eq $latestVersion) {
    Write-Host "You already have the latest driver installed."
    exit
}

# Construct download URL
$windowsVersion = "win10-win11"
$windowsArchitecture = "64bit"
$downloadUrl = "https://international.download.nvidia.com/Windows/$latestVersion/$latestVersion-desktop-$windowsVersion-$windowsArchitecture-international-dch-whql.exe"
$dlPath = Join-Path $Downloads "$latestVersion-NVIDIA.exe"

Write-Host "Downloading latest NVIDIA driver to:`n$dlPath"

# Function to download using aria2c
function Download-WithAria2c($url, $outFile) {
    $ariaPath = (Get-Command aria2c.exe -ErrorAction SilentlyContinue).Source
    if (-not $ariaPath) {
        return $false
    }

    Write-Host "aria2c found. Downloading with aria2c..."
    $args = @(
        "--console-log-level=error"
        "--file-allocation=none"
        "--dir=$(Split-Path $outFile)"
        "--out=$(Split-Path $outFile -Leaf)"
        "--max-connection-per-server=16"
        "--split=16"
        "--retry-wait=5"
        "--max-tries=5"
        $url
    )

    $process = Start-Process -FilePath $ariaPath -ArgumentList $args -NoNewWindow -Wait -PassThru
    return $process.ExitCode -eq 0
}

# Try download with aria2c, fallback to Invoke-WebRequest, fallback to RP version with aria2c again
$downloadSucceeded = Download-WithAria2c -url $downloadUrl -outFile $dlPath

if (-not $downloadSucceeded) {
    Write-Host "aria2c download failed. Falling back to Invoke-WebRequest..."
    try {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $dlPath -UseBasicParsing -ErrorAction Stop
        $downloadSucceeded = $true
        Write-Host "Download complete."
    }
    catch {
        Write-Host "Download failed. Trying fallback RP version..."
        $fallbackUrl = "https://international.download.nvidia.com/Windows/$latestVersion/$latestVersion-desktop-$windowsVersion-$windowsArchitecture-international-dch-whql-rp.exe"
        $downloadSucceeded = Download-WithAria2c -url $fallbackUrl -outFile $dlPath
        if (-not $downloadSucceeded) {
            Write-Error "Fallback RP version download failed. Exiting."
            exit 1
        }
    }
}

if (-not $downloadSucceeded) {
    Write-Error "Download failed. Exiting."
    exit 1
}

# Expand-NvidiaDriverPackage "$dlPath"
Write-Host "Unpacking..."
$unpackedFolder = [System.IO.Path]::GetFileNameWithoutExtension($dlPath) 
$unpackedFolder = Join-Path $Downloads $unpackedFolder
$setupEXE = "$unpackedFolder\setup.exe"

# Unpack driver
Expand-NvidiaDriverPackage "$dlPath"

# Install unpacked driver
Write-Host "Installing $dlPath..."
Start-Process -FilePath $setupEXE -ArgumentList " -passive -noeula -disable-nvtelemetry" -Wait

# Clean up
Write-Host "Cleaning up..."
Remove-Item "$dlPath" -Force -Recurse
Remove-Item "$unpackedFolder" -Force -Recurse
exit
