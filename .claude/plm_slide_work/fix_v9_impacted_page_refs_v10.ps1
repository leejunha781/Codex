$ErrorActionPreference = 'Stop'

$proposalDir = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$sourceDeck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_V9.pptx'
$v10Deck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_V10.pptx'
$defaultDeck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$finalDeck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$finalPdf = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf'
$workDir = 'C:\Users\namma\.claude\plm_slide_work'
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $proposalDir "_backup_20260607_impacted_refs_v10_$timestamp"
$reportPath = Join-Path $workDir 'future_direction_meeting_v10_impacted_refs_report.txt'

function Backup-IfExists {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$BackupDir
    )
    if (Test-Path -LiteralPath $Path) {
        Copy-Item -LiteralPath $Path -Destination (Join-Path $BackupDir ([System.IO.Path]::GetFileName($Path))) -Force
    }
}

function Replace-InShapes {
    param(
        $Shapes,
        [Parameter(Mandatory = $true)][string]$OldText,
        [Parameter(Mandatory = $true)][string]$NewText,
        [Parameter(Mandatory = $true)][ref]$Count
    )

    foreach ($shape in $Shapes) {
        try {
            if ($shape.Type -eq 6) {
                Replace-InShapes -Shapes $shape.GroupItems -OldText $OldText -NewText $NewText -Count $Count
            } elseif ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $textRange = $shape.TextFrame.TextRange
                $before = $textRange.Text
                if ($before -like "*$OldText*") {
                    try {
                        $null = $textRange.Replace($OldText, $NewText)
                    } catch {
                        $textRange.Text = $before.Replace($OldText, $NewText)
                    }
                    $after = $shape.TextFrame.TextRange.Text
                    if ($after -ne $before) {
                        $Count.Value = $Count.Value + 1
                    }
                }
            }
        } catch {}
    }
}

function Replace-InSlide {
    param(
        $Slide,
        [Parameter(Mandatory = $true)][string]$OldText,
        [Parameter(Mandatory = $true)][string]$NewText
    )
    $count = 0
    Replace-InShapes -Shapes $Slide.Shapes -OldText $OldText -NewText $NewText -Count ([ref]$count)
    return $count
}

function Set-TextContainingInShapes {
    param(
        $Shapes,
        [Parameter(Mandatory = $true)][string]$Needle,
        [Parameter(Mandatory = $true)][string]$NewText,
        [Parameter(Mandatory = $true)][ref]$Count
    )

    foreach ($shape in $Shapes) {
        try {
            if ($shape.Type -eq 6) {
                Set-TextContainingInShapes -Shapes $shape.GroupItems -Needle $Needle -NewText $NewText -Count $Count
            } elseif ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $textRange = $shape.TextFrame.TextRange
                $before = $textRange.Text
                if ($before -like "*$Needle*") {
                    $textRange.Text = $NewText
                    $textRange.Font.Size = [single]10.5
                    $textRange.Font.Color.RGB = 16777215
                    $textRange.Font.Bold = [int]-1
                    try { $shape.TextFrame.VerticalAnchor = 3 } catch {}
                    try { $shape.TextFrame.MarginTop = 4; $shape.TextFrame.MarginBottom = 4 } catch {}
                    $Count.Value = $Count.Value + 1
                }
            }
        } catch {}
    }
}

function Set-SlideTextContaining {
    param(
        $Slide,
        [Parameter(Mandatory = $true)][string]$Needle,
        [Parameter(Mandatory = $true)][string]$NewText
    )
    $count = 0
    Set-TextContainingInShapes -Shapes $Slide.Shapes -Needle $Needle -NewText $NewText -Count ([ref]$count)
    return $count
}

function Get-PowerPointApplication {
    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    } catch {
        try {
            return (New-Object -ComObject PowerPoint.Application)
        } catch {
            $exe = 'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE'
            Start-Process -FilePath $exe -WindowStyle Hidden
            Start-Sleep -Seconds 8
            return [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
        }
    }
}

New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
Backup-IfExists -Path $sourceDeck -BackupDir $backupDir
Backup-IfExists -Path $v10Deck -BackupDir $backupDir
Backup-IfExists -Path $defaultDeck -BackupDir $backupDir
Backup-IfExists -Path $finalDeck -BackupDir $backupDir
Backup-IfExists -Path $finalPdf -BackupDir $backupDir

if (-not (Test-Path -LiteralPath $sourceDeck)) {
    throw "Source deck not found: $sourceDeck"
}

Copy-Item -LiteralPath $sourceDeck -Destination $v10Deck -Force

$ppt = $null
$pres = $null
$changes = New-Object System.Collections.Generic.List[string]
try {
    $ppt = Get-PowerPointApplication
    $pres = $ppt.Presentations.Open($v10Deck, $false, $false, $false)

    $c = Replace-InSlide -Slide $pres.Slides.Item(1) -OldText 'Rev. V9 - 2026-06-06' -NewText 'Rev. V10 - 2026-06-07'
    if ($c -eq 0) {
        $c = Replace-InSlide -Slide $pres.Slides.Item(1) -OldText 'Rev. V9' -NewText 'Rev. V10'
        $c = $c + (Replace-InSlide -Slide $pres.Slides.Item(1) -OldText 'Rev. V10 - 2026-06-06' -NewText 'Rev. V10 - 2026-06-07')
    }
    $c = $c + (Replace-InSlide -Slide $pres.Slides.Item(1) -OldText '2026-06-06' -NewText '2026-06-07')
    [void]$changes.Add("Slide 1: revision/date updated to Rev. V10 - 2026-06-07 ($c replacement)")

    $oldSlide18Full = 'Pilot value: deterministic, auditable route proposals on existing Linux control-plane infrastructure; no automatic E3D promote without human/class validation. | Scope: future-phase concept; separate from Phase 2 Naval Assurance WBS 0–5 (p.28).'
    $newSlide18Full = 'Pilot value: auditable Linux-based route proposals with human/class validation. | Scope: future phase; WBS p.31.'
    $c = Replace-InSlide -Slide $pres.Slides.Item(18) -OldText $oldSlide18Full -NewText $newSlide18Full
    $oldSlide18 = 'Scope: future-phase concept; separate from Phase 2 Naval Assurance WBS 0–5 (p.28).'
    $newSlide18 = 'Scope: future phase; WBS p.31.'
    if ($c -eq 0) {
        $c = Replace-InSlide -Slide $pres.Slides.Item(18) -OldText $oldSlide18 -NewText $newSlide18
    }
    if ($c -eq 0) {
        $c = Replace-InSlide -Slide $pres.Slides.Item(18) -OldText 'separate from Phase 2 Naval Assurance WBS 0–5 (p.28).' -NewText 'WBS p.31.'
    }
    $c = $c + (Set-SlideTextContaining -Slide $pres.Slides.Item(18) -Needle 'Pilot value:' -NewText 'Pilot value: auditable Linux-based route proposals with human/class validation. | Scope: future phase; WBS p.31.')
    [void]$changes.Add("Slide 18: shortened WBS reference and corrected p.28 -> p.31 ($c replacement)")

    $c = Replace-InSlide -Slide $pres.Slides.Item(38) -OldText 'Future Industrial PLM / Meeting Procedure' -NewText 'Future Industrial PLM / Close'
    [void]$changes.Add("Slide 38: footer section Meeting Procedure -> Close ($c replacement)")

    $c = Replace-InSlide -Slide $pres.Slides.Item(2) -OldText 'Meeting procedure guide -> p.3-5' -NewText 'Meeting procedure guide -> p.3–5'
    [void]$changes.Add("Slide 2: normalized meeting guide range p.3-5 -> p.3–5 ($c replacement)")

    $pres.Save()
    $pres.Close()
    $pres = $null
} finally {
    if ($pres -ne $null) {
        try { $pres.Close() } catch {}
    }
    if ($ppt -ne $null) {
        try {
            if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() }
        } catch {}
    }
    if ($pres -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null 2>$null
    }
    if ($ppt -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null 2>$null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Copy-Item -LiteralPath $v10Deck -Destination $defaultDeck -Force
Copy-Item -LiteralPath $v10Deck -Destination $finalDeck -Force

$ppt = $null
$pres = $null
try {
    $ppt = Get-PowerPointApplication
    $pres = $ppt.Presentations.Open($finalDeck, $true, $false, $false)
    $pres.SaveAs($finalPdf, 32)
    $pres.Close()
    $pres = $null
} finally {
    if ($pres -ne $null) {
        try { $pres.Close() } catch {}
    }
    if ($ppt -ne $null) {
        try {
            if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() }
        } catch {}
    }
    if ($pres -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null 2>$null
    }
    if ($ppt -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null 2>$null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$lines = New-Object System.Collections.Generic.List[string]
[void]$lines.Add('Future Industrial PLM V10 - impacted page/content reference correction')
[void]$lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')")
[void]$lines.Add('Scope: correct content affected by adding/moving meeting-procedure pages 3-5.')
[void]$lines.Add('')
[void]$lines.Add('Changes:')
foreach ($change in $changes) { [void]$lines.Add("- $change") }
[void]$lines.Add('')
[void]$lines.Add("Backup folder: $backupDir")
[void]$lines.Add('')
[void]$lines.Add('Saved files:')
foreach ($path in @($v10Deck, $defaultDeck, $finalDeck, $finalPdf)) {
    $item = Get-Item -LiteralPath $path
    [void]$lines.Add(("{0}`t{1}`t{2}" -f $item.FullName, $item.Length, $item.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')))
}
[System.IO.File]::WriteAllLines($reportPath, $lines.ToArray(), [System.Text.Encoding]::UTF8)

Write-Host ([string]::Join("`n", $lines.ToArray()))
