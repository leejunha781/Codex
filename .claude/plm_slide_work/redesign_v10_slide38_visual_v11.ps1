$ErrorActionPreference = 'Stop'

$proposalDir = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$sourceDeck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$v11Deck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_V11.pptx'
$defaultDeck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$finalDeck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$finalPdf = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf'
$workDir = 'C:\Users\namma\.claude\plm_slide_work'
$srcImage = 'C:\Users\namma\.codex\generated_images\019e9c6d-6c87-7662-88d9-1f7f01934a80\ig_01db6af15a727d80016a244111e51c8191aad498a09cbbfcc2.png'
$projectImage = Join-Path $workDir 'future_direction_control_plane_background_v11.png'
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $proposalDir "_backup_20260607_slide38_visual_v11_$timestamp"
$reportPath = Join-Path $workDir 'future_direction_meeting_v11_slide38_visual_report.txt'

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

function Add-TextBox {
    param(
        $Slide,
        [string]$Name,
        [float]$Left,
        [float]$Top,
        [float]$Width,
        [float]$Height,
        [string]$Text,
        [float]$Size,
        [int]$Color,
        [int]$Bold = 0
    )
    $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $shape.Name = $Name
    $shape.TextFrame.TextRange.Text = $Text
    $shape.TextFrame.TextRange.Font.Name = 'Aptos Display'
    $shape.TextFrame.TextRange.Font.Size = [single]$Size
    $shape.TextFrame.TextRange.Font.Color.RGB = $Color
    $shape.TextFrame.TextRange.Font.Bold = [int]$Bold
    $shape.TextFrame.MarginLeft = 0
    $shape.TextFrame.MarginRight = 0
    $shape.TextFrame.MarginTop = 0
    $shape.TextFrame.MarginBottom = 0
    return $shape
}

function Add-RoundedPanel {
    param(
        $Slide,
        [string]$Name,
        [float]$Left,
        [float]$Top,
        [float]$Width,
        [float]$Height,
        [int]$Fill,
        [int]$Line,
        [float]$Transparency = 0,
        [float]$LineWeight = 1.5
    )
    $shape = $Slide.Shapes.AddShape(5, $Left, $Top, $Width, $Height)
    $shape.Name = $Name
    $shape.Fill.Visible = -1
    $shape.Fill.ForeColor.RGB = $Fill
    $shape.Fill.Transparency = [single]$Transparency
    $shape.Line.Visible = -1
    $shape.Line.ForeColor.RGB = $Line
    $shape.Line.Weight = [single]$LineWeight
    return $shape
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

if (-not (Test-Path -LiteralPath $sourceDeck)) { throw "Missing source deck: $sourceDeck" }
if (-not (Test-Path -LiteralPath $srcImage)) { throw "Missing generated image: $srcImage" }

New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
Backup-IfExists -Path $sourceDeck -BackupDir $backupDir
Backup-IfExists -Path $v11Deck -BackupDir $backupDir
Backup-IfExists -Path $defaultDeck -BackupDir $backupDir
Backup-IfExists -Path $finalDeck -BackupDir $backupDir
Backup-IfExists -Path $finalPdf -BackupDir $backupDir

Copy-Item -LiteralPath $srcImage -Destination $projectImage -Force
Copy-Item -LiteralPath $sourceDeck -Destination $v11Deck -Force

$ppt = $null
$pres = $null
try {
    $ppt = Get-PowerPointApplication
    $pres = $ppt.Presentations.Open($v11Deck, $false, $false, $false)

    $revCount = 0
    Set-TextContainingInShapes -Shapes $pres.Slides.Item(1).Shapes -Needle 'Rev. V10' -NewText 'Rev. V11 · 2026-06-07' -Count ([ref]$revCount)

    $slide = $pres.Slides.Item(38)
    while ($slide.Shapes.Count -gt 0) {
        $slide.Shapes.Item(1).Delete()
    }

    $W = [single]$pres.PageSetup.SlideWidth
    $H = [single]$pres.PageSetup.SlideHeight
    $navy = Rgb 4 17 31
    $panel = Rgb 8 31 50
    $cyan = Rgb 34 214 230
    $teal = Rgb 50 214 156
    $amber = Rgb 255 176 64
    $white = Rgb 245 250 255
    $muted = Rgb 178 196 210

    $bg = $slide.Shapes.AddPicture($projectImage, 0, -1, 0, 0, $W, $H)
    $bg.Name = 'GeneratedControlPlaneBackground'
    $bg.ZOrder(1)

    $wash = $slide.Shapes.AddShape(1, 0, 0, $W, $H)
    $wash.Name = 'DarkReadabilityWash'
    $wash.Fill.ForeColor.RGB = $navy
    $wash.Fill.Transparency = [single]0.25
    $wash.Line.Visible = 0
    $wash.ZOrder(1)

    $leftPanel = $slide.Shapes.AddShape(1, 0, 0, 525, $H)
    $leftPanel.Name = 'LeftGradientReadabilityPanel'
    $leftPanel.Fill.ForeColor.RGB = $navy
    $leftPanel.Fill.Transparency = [single]0.08
    $leftPanel.Line.Visible = 0

    $title = Add-TextBox -Slide $slide -Name 'Slide38TitleV11' -Left 42 -Top 34 -Width 700 -Height 44 -Text 'Closing argument' -Size 24 -Color $white -Bold -1
    $subtitle = Add-TextBox -Slide $slide -Name 'Slide38SubtitleV11' -Left 44 -Top 78 -Width 670 -Height 28 -Text 'Future AVEVA Marine direction - one decision, one governed loop' -Size 10.5 -Color $muted -Bold 0

    $line = $slide.Shapes.AddShape(1, 44, 116, 470, 3)
    $line.Name = 'Slide38AccentLineV11'
    $line.Fill.ForeColor.RGB = $cyan
    $line.Line.Visible = 0

    $hero = Add-RoundedPanel -Slide $slide -Name 'HeroStatementPanelV11' -Left 44 -Top 146 -Width 452 -Height 148 -Fill $panel -Line $cyan -Transparency 0.06 -LineWeight 2.2
    $hero.Adjustments.Item(1) = 0.12
    $headline = 'AVEVA Marine should become the open, governed control plane for every shipbuilding state transition.'
    Add-TextBox -Slide $slide -Name 'HeroStatementTextV11' -Left 70 -Top 174 -Width 402 -Height 90 -Text $headline -Size 21 -Color $white -Bold -1 | Out-Null

    $proofY = 326
    $cardW = 140
    $gap = 13
    $cards = @(
        @{Name='Credible'; Header='CREDIBLE'; Body='Executable FastAPI + PostgreSQL + OIDC reference package.'; Color=$cyan},
        @{Name='Different'; Header='DIFFERENT'; Body='Decision, evidence and operations learning above specialist tools.'; Color=$teal},
        @{Name='Approve'; Header='APPROVE'; Body='Phase 2 core + one bounded Continuous Naval Assurance pilot.'; Color=$amber}
    )
    for ($i = 0; $i -lt $cards.Count; $i++) {
        $x = 44 + (($cardW + $gap) * $i)
        $card = Add-RoundedPanel -Slide $slide -Name ("Slide38" + $cards[$i].Name + "CardV11") -Left $x -Top $proofY -Width $cardW -Height 96 -Fill $panel -Line $cards[$i].Color -Transparency 0.08 -LineWeight 1.7
        $card.Adjustments.Item(1) = 0.16
        Add-TextBox -Slide $slide -Name ("Slide38" + $cards[$i].Name + "HeaderV11") -Left ($x + 14) -Top ($proofY + 15) -Width ($cardW - 28) -Height 16 -Text $cards[$i].Header -Size 9.8 -Color $cards[$i].Color -Bold -1 | Out-Null
        Add-TextBox -Slide $slide -Name ("Slide38" + $cards[$i].Name + "BodyV11") -Left ($x + 14) -Top ($proofY + 39) -Width ($cardW - 24) -Height 44 -Text $cards[$i].Body -Size 8.8 -Color $white -Bold 0 | Out-Null
    }

    $ask = Add-RoundedPanel -Slide $slide -Name 'MeetingAskPanelV11' -Left 44 -Top 450 -Width 868 -Height 46 -Fill (Rgb 33 18 4) -Line $amber -Transparency 0.08 -LineWeight 1.8
    $ask.Adjustments.Item(1) = 0.18
    Add-TextBox -Slide $slide -Name 'MeetingAskTextV11' -Left 66 -Top 462 -Width 820 -Height 20 -Text 'Meeting ask: align product direction, approve pilot scope, assign owners, and validate the first customer/class evidence scenario.' -Size 13 -Color $white -Bold -1 | Out-Null

    $hub = Add-RoundedPanel -Slide $slide -Name 'DecisionHubBadgeV11' -Left 602 -Top 58 -Width 238 -Height 68 -Fill $panel -Line $cyan -Transparency 0.12 -LineWeight 1.8
    $hub.Adjustments.Item(1) = 0.18
    Add-TextBox -Slide $slide -Name 'DecisionHubTextV11' -Left 624 -Top 75 -Width 198 -Height 32 -Text 'OPEN GOVERNED CONTROL PLANE' -Size 13.5 -Color $white -Bold -1 | Out-Null

    $footer = Add-TextBox -Slide $slide -Name 'Slide38FooterV11' -Left 38 -Top 515 -Width 250 -Height 12 -Text 'Future Industrial PLM / Close' -Size 6.5 -Color (Rgb 130 150 168) -Bold 0
    $pg = Add-TextBox -Slide $slide -Name 'Slide38PageNumV11' -Left 895 -Top 514 -Width 28 -Height 12 -Text '38' -Size 7.5 -Color (Rgb 160 178 193) -Bold 0
    $pg.TextFrame.TextRange.ParagraphFormat.Alignment = 3

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

Copy-Item -LiteralPath $v11Deck -Destination $defaultDeck -Force
Copy-Item -LiteralPath $v11Deck -Destination $finalDeck -Force

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
[void]$lines.Add('Future Industrial PLM V11 - slide 38 visual redesign')
[void]$lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')")
[void]$lines.Add('Scope: redesign slide 38 for higher visibility and add generated control-plane image.')
[void]$lines.Add("Generated source image: $srcImage")
[void]$lines.Add("Project image copy: $projectImage")
[void]$lines.Add("Backup folder: $backupDir")
[void]$lines.Add('')
[void]$lines.Add('Changes:')
[void]$lines.Add('- Slide 38 rebuilt as full-bleed generated marine control-plane visual with dark readability panel.')
[void]$lines.Add('- Enlarged closing argument into one high-contrast hero statement.')
[void]$lines.Add('- Reworked supporting content into three proof cards: CREDIBLE / DIFFERENT / APPROVE.')
[void]$lines.Add('- Added strong bottom meeting-ask banner and kept footer/page number.')
[void]$lines.Add('- Cover revision updated to Rev. V11 · 2026-06-07.')
[void]$lines.Add('')
[void]$lines.Add('Saved files:')
foreach ($path in @($v11Deck, $defaultDeck, $finalDeck, $finalPdf, $projectImage)) {
    $item = Get-Item -LiteralPath $path
    [void]$lines.Add(("{0}`t{1}`t{2}" -f $item.FullName, $item.Length, $item.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')))
}
[System.IO.File]::WriteAllLines($reportPath, $lines.ToArray(), [System.Text.Encoding]::UTF8)
Write-Host ([string]::Join("`n", $lines.ToArray()))
