$ErrorActionPreference = 'Stop'

$proposalDir = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$sourceDeck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$v12Deck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_V12.pptx'
$defaultDeck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$finalDeck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$finalPdf = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf'
$workDir = 'C:\Users\namma\.claude\plm_slide_work'
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $proposalDir "_backup_20260607_slides3_5_generated_bg_v12_$timestamp"
$reportPath = Join-Path $workDir 'future_direction_meeting_v12_slides3_5_generated_bg_report.txt'

$imageMap = @(
    @{
        Slide = 3
        Source = 'C:\Users\namma\.codex\generated_images\019e9c6d-6c87-7662-88d9-1f7f01934a80\ig_01db6af15a727d80016a244437c03881919336ed9f3e9cdc68.png'
        Project = (Join-Path $workDir 'slide03_decision_path_background_v12.png')
        PictureName = 'GeneratedDecisionPathBackgroundV12'
        WashName = 'GeneratedDecisionPathReadabilityWashV12'
        Transparency = 0.39
    },
    @{
        Slide = 4
        Source = 'C:\Users\namma\.codex\generated_images\019e9c6d-6c87-7662-88d9-1f7f01934a80\ig_01db6af15a727d80016a2444b50eb48191ab360f930e4cbaa2.png'
        Project = (Join-Path $workDir 'slide04_development_pipeline_background_v12.png')
        PictureName = 'GeneratedDevelopmentPipelineBackgroundV12'
        WashName = 'GeneratedDevelopmentPipelineReadabilityWashV12'
        Transparency = 0.43
    },
    @{
        Slide = 5
        Source = 'C:\Users\namma\.codex\generated_images\019e9c6d-6c87-7662-88d9-1f7f01934a80\ig_01db6af15a727d80016a24453d8fec819199c673b893528def.png'
        Project = (Join-Path $workDir 'slide05_planning_strategy_background_v12.png')
        PictureName = 'GeneratedPlanningStrategyBackgroundV12'
        WashName = 'GeneratedPlanningStrategyReadabilityWashV12'
        Transparency = 0.42
    }
)

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

function Backup-IfExists {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$BackupDir
    )
    if (Test-Path -LiteralPath $Path) {
        Copy-Item -LiteralPath $Path -Destination (Join-Path $BackupDir ([System.IO.Path]::GetFileName($Path))) -Force
    }
}

function Rgb {
    param([int]$R, [int]$G, [int]$B)
    return ($R + ($G * 256) + ($B * 65536))
}

function Set-TextContainingInShapes {
    param(
        $Shapes,
        [string]$Needle,
        [string]$NewText,
        [ref]$Count
    )
    foreach ($shape in $Shapes) {
        try {
            if ($shape.Type -eq 6) {
                Set-TextContainingInShapes -Shapes $shape.GroupItems -Needle $Needle -NewText $NewText -Count $Count
            } elseif ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                if ($shape.TextFrame.TextRange.Text -like "*$Needle*") {
                    $shape.TextFrame.TextRange.Text = $NewText
                    $Count.Value = $Count.Value + 1
                }
            }
        } catch {}
    }
}

function Remove-ExistingGeneratedBg {
    param($Slide)
    for ($i = $Slide.Shapes.Count; $i -ge 1; $i--) {
        $shape = $Slide.Shapes.Item($i)
        if ($shape.Name -like 'Generated*BackgroundV12' -or $shape.Name -like 'Generated*ReadabilityWashV12') {
            $shape.Delete()
        }
    }
}

if (-not (Test-Path -LiteralPath $sourceDeck)) { throw "Missing source deck: $sourceDeck" }
foreach ($entry in $imageMap) {
    if (-not (Test-Path -LiteralPath $entry.Source)) { throw "Missing generated image: $($entry.Source)" }
}

New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
Backup-IfExists -Path $sourceDeck -BackupDir $backupDir
Backup-IfExists -Path $v12Deck -BackupDir $backupDir
Backup-IfExists -Path $defaultDeck -BackupDir $backupDir
Backup-IfExists -Path $finalDeck -BackupDir $backupDir
Backup-IfExists -Path $finalPdf -BackupDir $backupDir

foreach ($entry in $imageMap) {
    Copy-Item -LiteralPath $entry.Source -Destination $entry.Project -Force
}
Copy-Item -LiteralPath $sourceDeck -Destination $v12Deck -Force

$ppt = $null
$pres = $null
try {
    $ppt = Get-PowerPointApplication
    $pres = $ppt.Presentations.Open($v12Deck, $false, $false, $false)

    $revCount = 0
    Set-TextContainingInShapes -Shapes $pres.Slides.Item(1).Shapes -Needle 'Rev. V11' -NewText 'Rev. V12 · 2026-06-07' -Count ([ref]$revCount)

    $slideW = [single]$pres.PageSetup.SlideWidth
    $slideH = [single]$pres.PageSetup.SlideHeight
    $navy = Rgb 4 17 31

    foreach ($entry in $imageMap) {
        $slide = $pres.Slides.Item([int]$entry.Slide)
        Remove-ExistingGeneratedBg -Slide $slide

        $wash = $slide.Shapes.AddShape(1, 0, 0, $slideW, $slideH)
        $wash.Name = $entry.WashName
        $wash.Fill.Visible = -1
        $wash.Fill.ForeColor.RGB = $navy
        $wash.Fill.Transparency = [single]$entry.Transparency
        $wash.Line.Visible = 0
        $wash.ZOrder(1)

        $pic = $slide.Shapes.AddPicture($entry.Project, 0, -1, 0, 0, $slideW, $slideH)
        $pic.Name = $entry.PictureName
        $pic.ZOrder(1)
    }

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

Copy-Item -LiteralPath $v12Deck -Destination $defaultDeck -Force
Copy-Item -LiteralPath $v12Deck -Destination $finalDeck -Force

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
[void]$lines.Add('Future Industrial PLM V12 - generated backgrounds for slides 3-5')
[void]$lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')")
[void]$lines.Add('Scope: add content-matched generated background images to p.3, p.4 and p.5.')
[void]$lines.Add("Backup folder: $backupDir")
[void]$lines.Add('')
[void]$lines.Add('Image mapping:')
foreach ($entry in $imageMap) {
    [void]$lines.Add(("Slide {0}: {1} -> {2}" -f $entry.Slide, $entry.Source, $entry.Project))
}
[void]$lines.Add('')
[void]$lines.Add('Changes:')
[void]$lines.Add('- Slide 3: added generated decision-path background under existing meeting-procedure cards.')
[void]$lines.Add('- Slide 4: added generated development pipeline background under BUILD/VALIDATE/AVOID cards.')
[void]$lines.Add('- Slide 5: added generated planning strategy background under planning-team argument cards.')
[void]$lines.Add('- Added dark readability wash on each changed slide; existing PowerPoint text/shapes remain editable.')
[void]$lines.Add('- Cover revision updated to Rev. V12 · 2026-06-07.')
[void]$lines.Add('')
[void]$lines.Add('Saved files:')
foreach ($path in @($v12Deck, $defaultDeck, $finalDeck, $finalPdf)) {
    $item = Get-Item -LiteralPath $path
    [void]$lines.Add(("{0}`t{1}`t{2}" -f $item.FullName, $item.Length, $item.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')))
}
foreach ($entry in $imageMap) {
    $item = Get-Item -LiteralPath $entry.Project
    [void]$lines.Add(("{0}`t{1}`t{2}" -f $item.FullName, $item.Length, $item.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')))
}
[System.IO.File]::WriteAllLines($reportPath, $lines.ToArray(), [System.Text.Encoding]::UTF8)
Write-Host ([string]::Join("`n", $lines.ToArray()))
