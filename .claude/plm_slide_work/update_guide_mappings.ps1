$ErrorActionPreference = "Stop"
$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\"
$enc = New-Object System.Text.UTF8Encoding($false)  # no BOM

# ---------- EN ----------
$enMd = $dir + "Future_Industrial_PLM_Presenter_Facilitation_Guide_EN.md"
$lines = Get-Content $enMd -Encoding UTF8
for ($i=0; $i -lt $lines.Count; $i++) {
    $l = $lines[$i]
    if ($l -match '^#### Block 2 ') { $l = $l -replace 'PLM SME -> both teams$','Consultant (PLM SME) -> Dev + Planning' }
    elseif ($l -match '^#### Block 3 ') { $l = $l -replace 'PLM SME -> both teams$','Consultant (PLM SME) -> Planning (Dev on AI)' }
    elseif ($l -match '^#### Block 4 ') { $l = $l -replace 'PLM SME \+ Architect -> Development$','Consultant (PLM SME) + Architect -> Dev (Planning on scope)' }
    elseif ($l -match '^#### Block 5 ') { $l = $l -replace 'PLM SME -> Planning$','Consultant (PLM SME) -> Dev + Planning' }
    elseif ($l -match '^#### Block 6 ') { $l = $l -replace 'PLM SME -> Planning$','Consultant (PLM SME) -> Planning' }
    $l = $l -replace 'Facilitator -> all$','Facilitator -> All'
    $l = $l -replace '\(V17, 42 slides\)','(V18, 42 slides)'
    $lines[$i] = $l
}
[System.IO.File]::WriteAllText($enMd, ($lines -join "`r`n") + "`r`n", $enc)
Write-Output "EN md updated"

# ---------- KO ----------
$koMd = $dir + "Future_Industrial_PLM_Presenter_Facilitation_Guide_KO.md"
$lines = Get-Content $koMd -Encoding UTF8
for ($i=0; $i -lt $lines.Count; $i++) {
    $l = $lines[$i]
    if ($l -match '^#### 블록 2 ') { $l = $l -replace 'PLM SME -> 양 팀$','컨설턴트(PLM SME) -> 개발 + 기획' }
    elseif ($l -match '^#### 블록 3 ') { $l = $l -replace 'PLM SME -> 양 팀$','컨설턴트(PLM SME) -> 기획 (AI는 개발)' }
    elseif ($l -match '^#### 블록 4 ') { $l = $l -replace 'PLM SME \+ Architect -> 개발팀$','컨설턴트(PLM SME) + Architect -> 개발 (범위는 기획)' }
    elseif ($l -match '^#### 블록 5 ') { $l = $l -replace 'PLM SME -> 기획팀$','컨설턴트(PLM SME) -> 개발 + 기획' }
    elseif ($l -match '^#### 블록 6 ') { $l = $l -replace 'PLM SME -> 기획팀$','컨설턴트(PLM SME) -> 기획' }
    $l = $l -replace '\(V17, 42장\)','(V18, 42장)'
    $lines[$i] = $l
}
[System.IO.File]::WriteAllText($koMd, ($lines -join "`r`n") + "`r`n", $enc)
Write-Output "KO md updated"
