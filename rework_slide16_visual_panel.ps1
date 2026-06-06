$file = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
$img = "C:\Users\namma\.claude\plm_slide_work\img17_hybrid_boundary.png"
$backupDir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\_backup_20260606_v4_slide16_visualpanel"
$renderDir = "C:\Users\namma\.claude\plm_slide_work\v4_slide16_visualpanel_render"
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
New-Item -ItemType Directory -Force -Path $renderDir | Out-Null

if (-not (Test-Path -LiteralPath $file)) { throw "Missing PPTX: $file" }
if (-not (Test-Path -LiteralPath $img)) { throw "Missing image: $img" }

function RgbInt([int]$r, [int]$g, [int]$b) {
    return ($r + ($g * 256) + ($b * 65536))
}

function SetText($shape, [string]$text, [single]$fontSize) {
    $shape.TextFrame.TextRange.Text = $text
    $shape.TextFrame.TextRange.Font.Name = "Aptos"
    $shape.TextFrame.TextRange.Font.Size = $fontSize
    $shape.TextFrame.TextRange.Font.Color.RGB = RgbInt 222 235 248
    $shape.TextFrame.MarginLeft = 10
    $shape.TextFrame.MarginRight = 10
    $shape.TextFrame.MarginTop = 6
    $shape.TextFrame.MarginBottom = 5
    try { $shape.TextFrame2.AutoSize = 0 } catch {}
    try { $shape.TextFrame2.WordWrap = -1 } catch {}
}

$pp = $null
$pres = $null

try {
    try {
        $pp = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
    } catch {
        $pp = New-Object -ComObject PowerPoint.Application
    }
    $pp.Visible = $true

    foreach ($p in $pp.Presentations) {
        if ($p.FullName -eq $file) {
            $pres = $p
            break
        }
    }
    if (-not $pres) {
        $pres = $pp.Presentations.Open($file, $false, $false, $true)
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backup = Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4_BEFORE_SLIDE16_VISUALPANEL_{0}.pptx" -f $stamp)
    $pres.SaveCopyAs($backup)
    Write-Output "BACKUP=$backup"

    $slide = $pres.Slides.Item(16)

    # Remove the previous background-image treatment and keep slide background native/dark.
    $slide.FollowMasterBackground = $false
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = 3678739

    for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
        $sh = $slide.Shapes.Item($i)
        if ($sh.Name -eq "GeneratedHybridBgScrim" -or
            $sh.Name -eq "HybridBoundaryVisual" -or
            $sh.Name -eq "HybridBoundaryCaption") {
            $sh.Delete()
        }
    }

    # Header fit-up.
    $title = $slide.Shapes.Item("TextBox 1")
    $title.Left = 34
    $title.Top = 18
    $title.Width = 800
    $title.Height = 34
    $title.TextFrame.TextRange.Font.Size = [single]24

    $subtitle = $slide.Shapes.Item("TextBox 2")
    $subtitle.Left = 34
    $subtitle.Top = 58
    $subtitle.Width = 882
    $subtitle.Height = 22
    $subtitle.TextFrame.TextRange.Font.Size = [single]12

    # Visible inserted image panel, not a background.
    $pic = $slide.Shapes.AddPicture($img, 0, -1, 40, 104, 400, 225)
    $pic.Name = "HybridBoundaryVisual"
    $pic.Line.Visible = -1
    $pic.Line.ForeColor.RGB = RgbInt 39 205 255
    $pic.Line.Weight = [single]1.4
    $pic.ZOrder(1) | Out-Null

    $capText = "Visual boundary: Windows authoring -> Local Agent/Plugin bridge -> Linux control plane -> Cloud/CONNECT adapters"
    $caption = $slide.Shapes.AddTextbox(1, 40, 335, 400, 18)
    $caption.Name = "HybridBoundaryCaption"
    $caption.TextFrame.TextRange.Text = $capText
    $caption.TextFrame.TextRange.Font.Name = "Aptos"
    $caption.TextFrame.TextRange.Font.Size = [single]8.5
    $caption.TextFrame.TextRange.Font.Color.RGB = RgbInt 77 213 255
    $caption.TextFrame.MarginLeft = 0
    $caption.TextFrame.MarginRight = 0
    $caption.TextFrame.MarginTop = 0
    $caption.TextFrame.MarginBottom = 0

    # Compact the original three zone cards to the right of the image.
    $cardSpecs = @(
        @{
            Name = "Rounded Rectangle 3"; Left = 468; Top = 104; Width = 454; Height = 70;
            Text = "Windows Authoring Zone`rAVEVA E3D / Marine / Hull · Draw / MTO`rLocal Agent / Plugin / PML.NET Bridge`rWindows designer PCs or supported AVEVA environments`rNo direct Linux hosting assumption"
        },
        @{
            Name = "Rounded Rectangle 4"; Left = 468; Top = 184; Width = 454; Height = 70;
            Text = "Linux Open PLM Control Plane`rFastAPI / OpenAPI · PostgreSQL / Evidence Store`rAI Gate + Rule Guardrails`rIntegration Gateway + Callback Receiver`rKeycloak / OIDC / RBAC"
        },
        @{
            Name = "Rounded Rectangle 5"; Left = 468; Top = 264; Width = 454; Height = 70;
            Text = "Cloud / CONNECT + Adapters`rCONNECT · AIM / PI Web API · AVEVA Edge / Adapter layer`rNAPA / Class / CFD / FEM adapters`rERP / MES / Document Vault`rFederates specialist tools without replacing them"
        }
    )

    foreach ($spec in $cardSpecs) {
        $card = $slide.Shapes.Item($spec.Name)
        $card.Left = [single]$spec.Left
        $card.Top = [single]$spec.Top
        $card.Width = [single]$spec.Width
        $card.Height = [single]$spec.Height
        SetText $card $spec.Text ([single]9.2)
        $card.Fill.Transparency = [single]0.03
        $card.Line.Weight = [single]1.2
        try {
            $card.TextFrame.TextRange.Paragraphs(1).Font.Bold = [int]1
            $card.TextFrame.TextRange.Paragraphs(1).Font.Size = [single]12.2
            $card.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = RgbInt 245 250 255
        } catch {}
    }

    # Resize correction table to preserve the user's correction logic in a compact lower band.
    $tableShape = $slide.Shapes.Item("Table 8")
    $tableShape.Left = 40
    $tableShape.Top = 360
    $tableShape.Width = 882
    $tableShape.Height = 118
    try {
        $tableShape.Table.Columns.Item(1).Width = 300
        $tableShape.Table.Columns.Item(2).Width = 582
        $tableShape.Table.Rows.Item(1).Height = 22
        for ($r = 2; $r -le $tableShape.Table.Rows.Count; $r++) {
            $tableShape.Table.Rows.Item($r).Height = 24
        }
    } catch {}
    foreach ($cell in $tableShape.Table.Range.Cells) {
        try {
            $cell.Shape.TextFrame.TextRange.Font.Name = "Aptos"
            $cell.Shape.TextFrame.TextRange.Font.Size = [single]8.7
            $cell.Shape.TextFrame.MarginLeft = 5
            $cell.Shape.TextFrame.MarginRight = 5
            $cell.Shape.TextFrame.MarginTop = 2
            $cell.Shape.TextFrame.MarginBottom = 2
        } catch {}
    }
    try {
        for ($c = 1; $c -le $tableShape.Table.Columns.Count; $c++) {
            $tableShape.Table.Cell(1,$c).Shape.TextFrame.TextRange.Font.Bold = [int]1
            $tableShape.Table.Cell(1,$c).Shape.TextFrame.TextRange.Font.Size = [single]9.5
        }
    } catch {}

    $interview = $slide.Shapes.Item("TextBox 9")
    $interview.Left = 40
    $interview.Top = 501
    $interview.Width = 880
    $interview.Height = 25
    $interview.TextFrame.TextRange.Font.Size = [single]9.6
    $interview.TextFrame.TextRange.Font.Color.RGB = RgbInt 190 204 220

    # Ensure brand/page number stay above all objects.
    try { $slide.Shapes.Item("Text 2").ZOrder(0) | Out-Null } catch {}
    try { $slide.Shapes.Item("Text 12").ZOrder(0) | Out-Null } catch {}

    $pres.Save()
    $png = Join-Path $renderDir "slide16_visual_panel.png"
    $slide.Export($png, "PNG", 1600, 900)
    Write-Output "SAVED=$file"
    Write-Output "PNG=$png"

    # Reopen from disk so the visible PowerPoint window cannot show a stale copy.
    try {
        $pres.Saved = $true
        $pres.Close()
    } catch {}
    $pres = $pp.Presentations.Open($file, $false, $false, $true)
    $pp.ActiveWindow.View.GotoSlide(16)
    Write-Output "REOPENED=$($pres.Name) SLIDES=$($pres.Slides.Count) READONLY=$($pres.ReadOnly)"
} finally {
    # Keep PowerPoint open for the user.
}
