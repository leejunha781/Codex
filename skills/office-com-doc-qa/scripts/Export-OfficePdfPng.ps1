#requires -Version 5.1
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [string]$OutputPdf,

    [string]$RenderDir,

    [uint32]$Width = 1240
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

if ($PSVersionTable.PSEdition -ne "Desktop") {
    throw "Run this script with Windows PowerShell 5.1 (powershell.exe), not PowerShell Core/pwsh, because WinRT PDF rendering types are required."
}

function Resolve-FullPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $Path))
}

function Export-WordPdf {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$PdfPath
    )

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $null
    try {
        $doc = $word.Documents.Open($SourcePath, $false, $true)
        $doc.ExportAsFixedFormat($PdfPath, 17)
        $doc.Close([ref]$false)
        $doc = $null
    } finally {
        if ($doc -ne $null) {
            try { $doc.Close([ref]$false) } catch {}
        }
        if ($word -ne $null -and $word.Documents.Count -eq 0) {
            $word.Quit()
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Export-PowerPointPdf {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$PdfPath
    )

    $ppt = New-Object -ComObject PowerPoint.Application
    $presentation = $null
    try {
        $presentation = $ppt.Presentations.Open($SourcePath, $true, $false, $false)
        $presentation.SaveAs($PdfPath, 32)
        $presentation.Close()
        $presentation = $null
    } finally {
        if ($presentation -ne $null) {
            try { $presentation.Close() } catch {}
        }
        if ($ppt -ne $null) {
            $ppt.Quit()
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Initialize-WinRtPdfTypes {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime] | Out-Null
    [Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime] | Out-Null
    [Windows.Data.Pdf.PdfPageRenderOptions,Windows.Data.Pdf,ContentType=WindowsRuntime] | Out-Null
    [Windows.Storage.Streams.InMemoryRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime] | Out-Null
    [Windows.Storage.Streams.DataReader,Windows.Storage.Streams,ContentType=WindowsRuntime] | Out-Null
}

function Await-AsyncOperation {
    param(
        [Parameter(Mandatory = $true)]$Operation,
        [Parameter(Mandatory = $true)][Type]$ResultType
    )

    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object {
            $parameters = $_.GetParameters()
            if ($_.Name -ne "AsTask" -or -not $_.IsGenericMethodDefinition -or $parameters.Count -ne 1) {
                $false
            } else {
                $parameters[0].ParameterType.Name -eq 'IAsyncOperation`1'
            }
        } |
        Select-Object -First 1

    if ($method -eq $null) {
        throw "Could not find WindowsRuntime AsTask overload for IAsyncOperation."
    }

    $task = $method.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    $task.Wait()
    return $task.Result
}

function Await-AsyncAction {
    param([Parameter(Mandatory = $true)]$Action)

    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object {
            $parameters = $_.GetParameters()
            if ($_.Name -ne "AsTask" -or $_.IsGenericMethodDefinition -or $parameters.Count -ne 1) {
                $false
            } else {
                $parameters[0].ParameterType.FullName -eq "Windows.Foundation.IAsyncAction"
            }
        } |
        Select-Object -First 1

    if ($method -eq $null) {
        throw "Could not find WindowsRuntime AsTask overload for IAsyncAction."
    }

    $task = $method.Invoke($null, @($Action))
    $task.Wait()
}

function Render-PdfToPng {
    param(
        [Parameter(Mandatory = $true)][string]$PdfPath,
        [Parameter(Mandatory = $true)][string]$OutputDir,
        [Parameter(Mandatory = $true)][uint32]$DestinationWidth
    )

    Initialize-WinRtPdfTypes
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

    $storageFile = Await-AsyncOperation -Operation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($PdfPath)) -ResultType ([Windows.Storage.StorageFile])
    $pdf = Await-AsyncOperation -Operation ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($storageFile)) -ResultType ([Windows.Data.Pdf.PdfDocument])
    $written = New-Object System.Collections.Generic.List[string]

    for ($i = 0; $i -lt $pdf.PageCount; $i++) {
        $page = $pdf.GetPage([uint32]$i)
        $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
        $options = New-Object Windows.Data.Pdf.PdfPageRenderOptions
        $options.DestinationWidth = $DestinationWidth
        Await-AsyncAction -Action ($page.RenderToStreamAsync($stream, $options))

        $stream.Seek(0) | Out-Null
        $reader = New-Object Windows.Storage.Streams.DataReader($stream.GetInputStreamAt(0))
        $size = [uint32]$stream.Size
        Await-AsyncOperation -Operation ($reader.LoadAsync($size)) -ResultType ([uint32]) | Out-Null

        $bytes = New-Object byte[] $size
        $reader.ReadBytes($bytes)

        $pngPath = Join-Path $OutputDir ("page-{0:D2}.png" -f ($i + 1))
        [System.IO.File]::WriteAllBytes($pngPath, $bytes)
        $written.Add($pngPath) | Out-Null

        $reader.Dispose()
        $stream.Dispose()
        $page.Dispose()
    }

    return $written
}

$inputFullPath = Resolve-FullPath -Path $InputPath
if (-not (Test-Path -LiteralPath $inputFullPath)) {
    throw "Input file not found: $inputFullPath"
}

$extension = [System.IO.Path]::GetExtension($inputFullPath).ToLowerInvariant()
if ([string]::IsNullOrWhiteSpace($OutputPdf)) {
    if ($extension -eq ".pdf") {
        $OutputPdf = $inputFullPath
    } else {
        $OutputPdf = [System.IO.Path]::ChangeExtension($inputFullPath, ".pdf")
    }
} else {
    $OutputPdf = Resolve-FullPath -Path $OutputPdf
}

$pdfDir = [System.IO.Path]::GetDirectoryName($OutputPdf)
if (-not [string]::IsNullOrWhiteSpace($pdfDir)) {
    New-Item -ItemType Directory -Force -Path $pdfDir | Out-Null
}

switch ($extension) {
    ".doc" { Export-WordPdf -SourcePath $inputFullPath -PdfPath $OutputPdf }
    ".docx" { Export-WordPdf -SourcePath $inputFullPath -PdfPath $OutputPdf }
    ".ppt" { Export-PowerPointPdf -SourcePath $inputFullPath -PdfPath $OutputPdf }
    ".pptx" { Export-PowerPointPdf -SourcePath $inputFullPath -PdfPath $OutputPdf }
    ".pdf" {
        if ($inputFullPath -ne $OutputPdf) {
            Copy-Item -LiteralPath $inputFullPath -Destination $OutputPdf -Force
        }
    }
    default {
        throw "Unsupported input extension '$extension'. Use .doc, .docx, .ppt, .pptx, or .pdf."
    }
}

if ([string]::IsNullOrWhiteSpace($RenderDir)) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($OutputPdf)
    $RenderDir = Join-Path ([System.IO.Path]::GetDirectoryName($OutputPdf)) ($name + "_png")
} else {
    $RenderDir = Resolve-FullPath -Path $RenderDir
}

$pngFiles = @(Render-PdfToPng -PdfPath $OutputPdf -OutputDir $RenderDir -DestinationWidth $Width)

[PSCustomObject]@{
    InputPath = $inputFullPath
    PdfPath = $OutputPdf
    RenderDir = $RenderDir
    PageCount = $pngFiles.Count
    PngFiles = $pngFiles
}
