$ErrorActionPreference = 'Stop'

try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
    if ($word.Documents.Count -eq 0) {
        $word.Quit()
        Write-Output 'Closed empty Word instance'
    } else {
        Write-Output ('Word has open docs: {0}' -f $word.Documents.Count)
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
} catch {
    Write-Output ('No active Word COM or close skipped: {0}' -f $_.Exception.Message)
}
