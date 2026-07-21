#requires -Version 5.1
<#
.SYNOPSIS
  Copy Navarino revised CV/strategy next to canonical originals on E:, then Word COM PDF + PNG QA.

.DESCRIPTION
  - NEVER overwrites canonical *_Professional_CV.docx or *_20260719.docx
  - Copies revised artifacts into: E:\이력서\Account Manager\
  - Exports PDF via Word COM (wdFormatPDF = 17)
  - Renders PNG pages via Word ExportAsFixedFormat + optional pdftoppm if available
  - Writes QA note under E:\이력서\Account Manager\qa-renders\

.NOTES
  Run on the Windows machine that has Microsoft Word and the E: drive:
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File `
      "C:\path\to\Copy-ToLocalAndWordComQa.ps1" `
      -SourceDir "C:\Users\namma\...\career\navarino-account-manager"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SourceDir = "",

    [Parameter(Mandatory = $false)]
    [string]$DestDir = "E:\이력서\Account Manager",

    [Parameter(Mandatory = $false)]
    [string]$RenderDir = ""
)

$ErrorActionPreference = "Stop"

function Resolve-SourceDir {
    param([string]$Hint)
    if ($Hint -and (Test-Path -LiteralPath $Hint)) { return (Resolve-Path -LiteralPath $Hint).Path }

    $candidates = @(
        (Join-Path $PSScriptRoot "."),
        (Join-Path $PSScriptRoot "..\..\career\navarino-account-manager"),
        "C:\Users\namma\Codex\career\navarino-account-manager",
        "C:\Users\namma\source\repos\Codex\career\navarino-account-manager",
        "D:\Codex\career\navarino-account-manager"
    )
    foreach ($c in $candidates) {
        $cv = Join-Path $c "Joonha_Lee_NAVARINO_Account_Manager_Professional_CV_Revised_JD_Fit_20260721.docx"
        if (Test-Path -LiteralPath $cv) { return (Resolve-Path -LiteralPath $c).Path }
    }
    throw "SourceDir not found. Pass -SourceDir to the folder containing the Revised JD Fit DOCX."
}

function Assert-NotCanonicalOverwrite {
    param([string]$Path)
    $name = [IO.Path]::GetFileName($Path)
    $blocked = @(
        "Joonha_Lee_NAVARINO_Account_Manager_Professional_CV.docx",
        "Job Ad_Account Manager.docx",
        "Navarino_한국_Account_Manager_회사조사_JD매칭_합격전략_20260719.docx"
    )
    if ($blocked -contains $name) {
        throw "Refusing to overwrite canonical file: $Path"
    }
}

function Copy-RevisedArtifact {
    param(
        [string]$From,
        [string]$ToDir
    )
    if (-not (Test-Path -LiteralPath $From)) {
        throw "Missing source file: $From"
    }
    if (-not (Test-Path -LiteralPath $ToDir)) {
        New-Item -ItemType Directory -Path $ToDir -Force | Out-Null
    }
    $dest = Join-Path $ToDir ([IO.Path]::GetFileName($From))
    Assert-NotCanonicalOverwrite -Path $dest
    Copy-Item -LiteralPath $From -Destination $dest -Force
    Write-Host "COPIED -> $dest"
    return $dest
}

function Export-WordPdf {
    param(
        [string]$InputPath,
        [string]$PdfPath
    )
    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($InputPath, $false, $true) # ReadOnly
        # 17 = wdFormatPDF
        $doc.SaveAs([ref]$PdfPath, [ref]17)
        Write-Host "PDF    -> $PdfPath"
    }
    finally {
        if ($doc -ne $null) {
            $doc.Close([ref]$false) | Out-Null
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
        }
        if ($word -ne $null) {
            $word.Quit() | Out-Null
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        }
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    }
}

function Export-PdfPagesPng {
    param(
        [string]$PdfPath,
        [string]$OutPrefix
    )
    $pdftoppm = Get-Command pdftoppm -ErrorAction SilentlyContinue
    if ($pdftoppm) {
        & $pdftoppm.Source -png -r 150 $PdfPath $OutPrefix
        Write-Host "PNG    -> ${OutPrefix}-*.png (pdftoppm)"
        return
    }

    # Fallback: Word can open PDF read-only in some installs; prefer Acrobat/ghostscript if present
    $gs = Get-Command gswin64c -ErrorAction SilentlyContinue
    if ($gs) {
        $dir = Split-Path -Parent $OutPrefix
        $base = Split-Path -Leaf $OutPrefix
        & $gs.Source -dSAFER -dBATCH -dNOPAUSE -sDEVICE=png16m -r150 `
            "-sOutputFile=$(Join-Path $dir ($base + '-%d.png'))" $PdfPath
        Write-Host "PNG    -> ${OutPrefix}-%d.png (ghostscript)"
        return
    }

    Write-Warning "Neither pdftoppm nor gswin64c found. PDF exported; open PDF manually for visual QA."
}

# --- main ---
$src = Resolve-SourceDir -Hint $SourceDir
if (-not $RenderDir) {
    $RenderDir = Join-Path $DestDir "qa-renders\navarino-am-20260721"
}
New-Item -ItemType Directory -Path $RenderDir -Force | Out-Null

Write-Host "SourceDir = $src"
Write-Host "DestDir   = $DestDir"
Write-Host "RenderDir = $RenderDir"

# Preserve check: list canonical files if present
$canonical = @(
    (Join-Path $DestDir "Joonha_Lee_NAVARINO_Account_Manager_Professional_CV.docx"),
    (Join-Path $DestDir "Job Ad_Account Manager.docx"),
    (Join-Path $DestDir "Navarino_한국_Account_Manager_회사조사_JD매칭_합격전략_20260719.docx")
)
foreach ($c in $canonical) {
    if (Test-Path -LiteralPath $c) {
        Write-Host "PRESERVED (exists, not touched): $c"
    }
    else {
        Write-Warning "Canonical not found (skip): $c"
    }
}

$cvSrc = Join-Path $src "Joonha_Lee_NAVARINO_Account_Manager_Professional_CV_Revised_JD_Fit_20260721.docx"
$stSrc = Join-Path $src "Navarino_한국_Account_Manager_회사조사_JD매칭_합격전략_20260721_Revised.docx"

$cvDest = Copy-RevisedArtifact -From $cvSrc -ToDir $DestDir
$stDest = Copy-RevisedArtifact -From $stSrc -ToDir $DestDir

$cvPdf = Join-Path $RenderDir "Joonha_Lee_NAVARINO_Account_Manager_Professional_CV_Revised_JD_Fit_20260721.pdf"
$stPdf = Join-Path $RenderDir "Navarino_한국_Account_Manager_회사조사_JD매칭_합격전략_20260721_Revised.pdf"

Export-WordPdf -InputPath $cvDest -PdfPath $cvPdf
Export-WordPdf -InputPath $stDest -PdfPath $stPdf

Export-PdfPagesPng -PdfPath $cvPdf -OutPrefix (Join-Path $RenderDir "cv-page")
Export-PdfPagesPng -PdfPath $stPdf -OutPrefix (Join-Path $RenderDir "strategy-page")

$note = @"
NAVARINO AM Word COM PDF QA — $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

Copied (new names only):
  $cvDest
  $stDest

PDF:
  $cvPdf
  $stPdf

Canonical originals were not overwritten.
Inspect every PNG under:
  $RenderDir

Checklist:
  [ ] Page count CV = 2 (preferred) or 3
  [ ] No clipped/overlapping text
  [ ] No fabricated Genohco satellite/TVAC / sales numbers / employed AVEVA Marine
  [ ] Title shows Account Manager
"@
$notePath = Join-Path $RenderDir "QA_NOTE.txt"
Set-Content -LiteralPath $notePath -Value $note -Encoding UTF8
Write-Host "NOTE   -> $notePath"
Write-Host "DONE."
