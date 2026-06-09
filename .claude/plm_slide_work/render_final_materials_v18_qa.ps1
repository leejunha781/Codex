$ErrorActionPreference = 'Stop'

$scriptPath = 'C:\Users\namma\.codex\skills\office-com-doc-qa\scripts\Export-OfficePdfPng.ps1'
$scriptBlock = [scriptblock]::Create((Get-Content -LiteralPath $scriptPath -Raw))
$proposal = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$work = 'C:\Users\namma\.claude\plm_slide_work'

& $scriptBlock -InputPath (Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V18.pptx') -RenderDir (Join-Path $work 'final_materials_v18_deck_render')
& $scriptBlock -InputPath (Join-Path $proposal 'Future_Industrial_PLM_Presenter_Facilitation_Guide_EN.pdf') -RenderDir (Join-Path $work 'final_materials_v18_guide_en_render')
& $scriptBlock -InputPath (Join-Path $proposal 'Future_Industrial_PLM_Presenter_Facilitation_Guide_KO.pdf') -RenderDir (Join-Path $work 'final_materials_v18_guide_ko_render')
