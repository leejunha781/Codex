$ErrorActionPreference = 'Stop'

$pp = $null
try {
    $pp = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
} catch {
    Write-Output 'No active PowerPoint COM object'
    exit 0
}

Write-Output ("Presentations: {0}" -f $pp.Presentations.Count)
for ($i = 1; $i -le $pp.Presentations.Count; $i++) {
    $p = $pp.Presentations.Item($i)
    [pscustomobject]@{
        Index = $i
        Name = $p.Name
        FullName = $p.FullName
        Saved = $p.Saved
        SlideCount = $p.Slides.Count
    } | Format-List
}
