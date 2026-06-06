$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$fi = Get-Item $file
Write-Output ("FILE: {0}" -f $file)
Write-Output ("SIZE: {0}   LastWrite: {1}" -f $fi.Length, $fi.LastWriteTime)

# copy to temp so a PowerPoint lock doesn't block reading
$tmp = Join-Path $env:TEMP ("chk_" + [System.IO.Path]::GetFileName($file))
Copy-Item $file $tmp -Force
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($tmp)
$markers = @("Open PLM Control Plane","OPEN PLM CONTROL PLANE","Hybrid: Windows","Windows authoring tier","Integration Gateway","AI Gate / Evidence")
try {
  foreach ($e in $zip.Entries) {
    if ($e.FullName -like "ppt/slides/slide*.xml") {
      $sr = New-Object System.IO.StreamReader($e.Open())
      $xml = $sr.ReadToEnd(); $sr.Close()
      $hits = @()
      foreach ($m in $markers) { if ($xml -match [regex]::Escape($m)) { $hits += $m } }
      if ($hits.Count -gt 0) { Write-Output ("  {0}  ->  {1}" -f $e.FullName, ($hits -join ", ")) }
    }
  }
} finally { $zip.Dispose(); Remove-Item $tmp -Force }
Write-Output "DONE-CHECK"
