# Strip C2PA / Content Credentials metadata from LinkedIn infographic PNGs.
# LinkedIn shows a "Content credentials" (CR) badge when caBX/C2PA chunks are present.
# Run after image generation and before mirror/post upload.
#
# Usage:
#   .\strip-linkedin-image-c2pa.ps1 -ImagePath "C:\...\topic-infographic.png"
#   .\strip-linkedin-image-c2pa.ps1 -Folder "C:\...\2026-07-09\engineering-rca-evidence-structure"

[CmdletBinding()]
param(
    [Parameter(ParameterSetName = "File")]
    [string]$ImagePath,

    [Parameter(ParameterSetName = "Folder")]
    [string]$Folder,

    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

function Test-PngHasC2paChunk {
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $text = [System.Text.Encoding]::ASCII.GetString($bytes)
    return ($text -match 'caBX' -or $text -match 'c2pa' -or $text -match 'jumb')
}

function Remove-C2paFromPng {
    param(
        [string]$Path,
        [bool]$IsDryRun
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Image not found: $Path"
    }

    if ($Path -notmatch '\.png$') {
        Write-Host "SKIP (not PNG): $Path"
        return $false
    }

    $hadC2pa = Test-PngHasC2paChunk -Path $Path
    if (-not $hadC2pa) {
        Write-Host "OK   no C2PA metadata: $Path"
        return $false
    }

    if ($IsDryRun) {
        Write-Host "[dry-run] would strip C2PA from $Path"
        return $true
    }

    Add-Type -AssemblyName System.Drawing
    $image = $null
    try {
        $image = [System.Drawing.Image]::FromFile($Path)
        $tempPath = "$Path.tmp.png"
        $image.Save($tempPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $image.Dispose()
        $image = $null

        Move-Item -LiteralPath $tempPath -Destination $Path -Force

        if (Test-PngHasC2paChunk -Path $Path) {
            throw "C2PA metadata still present after strip: $Path"
        }

        Write-Host "STRIP C2PA removed: $Path"
        return $true
    }
    finally {
        if ($null -ne $image) { $image.Dispose() }
    }
}

$targets = @()
if ($PSCmdlet.ParameterSetName -eq "File") {
    if ([string]::IsNullOrWhiteSpace($ImagePath)) {
        throw "ImagePath is required"
    }
    $targets += (Resolve-Path -LiteralPath $ImagePath).Path
}
else {
    if ([string]::IsNullOrWhiteSpace($Folder)) {
        throw "Folder is required"
    }
    $resolvedFolder = (Resolve-Path -LiteralPath $Folder).Path
    $targets += @(Get-ChildItem -Path $resolvedFolder -Filter "*-infographic.png" -File | ForEach-Object { $_.FullName })
}

if ($targets.Count -eq 0) {
    Write-Host "No infographic PNG files found."
    exit 0
}

$stripped = 0
foreach ($target in $targets) {
    if (Remove-C2paFromPng -Path $target -IsDryRun:([bool]$DryRun)) {
        $stripped++
    }
}

Write-Host "C2PA strip complete: $stripped file(s) cleaned."
exit 0
