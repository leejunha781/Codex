$ErrorActionPreference = 'Stop'

$proposal = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$backup = Join-Path $proposal ('_backup_20260609_final_materials_v18_review_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Path $backup -Force | Out-Null

$names = @(
    'Future_Industrial_PLM_Meeting_Deck_EN_V18.pptx',
    'Future_Industrial_PLM_Meeting_Deck_EN.pptx',
    'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx',
    'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf',
    'Future_Industrial_PLM_Presenter_Facilitation_Guide_EN.md',
    'Future_Industrial_PLM_Presenter_Facilitation_Guide_EN.docx',
    'Future_Industrial_PLM_Presenter_Facilitation_Guide_EN.pdf',
    'Future_Industrial_PLM_Presenter_Facilitation_Guide_KO.md',
    'Future_Industrial_PLM_Presenter_Facilitation_Guide_KO.docx',
    'Future_Industrial_PLM_Presenter_Facilitation_Guide_KO.pdf'
)

foreach ($name in $names) {
    $path = Join-Path $proposal $name
    if (Test-Path -LiteralPath $path) {
        Copy-Item -LiteralPath $path -Destination (Join-Path $backup $name) -Force
    }
}

Write-Output $backup
