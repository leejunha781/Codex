$ErrorActionPreference = 'Stop'

$scriptPath = 'C:\Users\namma\.codex\skills\office-com-doc-qa\scripts\Export-OfficePdfPng.ps1'
$inputPath = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$renderDir = 'C:\Users\namma\.claude\plm_slide_work\future_direction_meeting_v14_slide7_stack_cards_render'

$scriptBlock = [scriptblock]::Create((Get-Content -LiteralPath $scriptPath -Raw))
& $scriptBlock -InputPath $inputPath -RenderDir $renderDir
