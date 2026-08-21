param(
  [Parameter(Mandatory=$true)][string]$ProjectRoot,
  [switch]$IncludeVendor
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath $ProjectRoot).Path
$files = Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction Stop
$excluded = '[\\/](build[^\\/]*|Drivers|Middlewares|\.git)[\\/]'
$projectFiles = if($IncludeVendor){ $files } else { @($files | Where-Object { $_.FullName -notmatch $excluded }) }
$artifactFiles = @($files | Where-Object { $_.FullName -notmatch '[\\/](Drivers|Middlewares|CMakeFiles)[\\/]' })
$source = @($projectFiles | Where-Object { $_.Extension -in '.c','.h','.cpp','.hpp' })
$ioc = @($files | Where-Object { $_.Extension -eq '.ioc' })
$build = @($files | Where-Object { $_.Name -in 'CMakeLists.txt','Makefile','CMakePresets.json' })
$iar = @($artifactFiles | Where-Object { $_.Extension -in '.ewp','.eww','.icf','.map','.out' })
$firmware = @($artifactFiles | Where-Object { $_.Extension -in '.hex','.bin','.elf' })
$pdf = @($artifactFiles | Where-Object { $_.Extension -eq '.pdf' })
$patterns = @(
  'HAL_Delay\s*\(', 'while\s*\(\s*1\s*\)', 'malloc\s*\(', 'free\s*\(',
  'sprintf\s*\(', 'strcpy\s*\(', 'USER CODE BEGIN', 'TODO|FIXME',
  'HAL_UART_Transmit\s*\([^,]+,[^,]+,[^,]+,\s*HAL_MAX_DELAY'
)
[pscustomobject]@{
  Root=$root; SourceFiles=$source.Count; CubeMxIoc=$ioc.FullName;
  BuildFiles=$build.FullName; IarArtifacts=$iar.FullName; FirmwareImages=$firmware.FullName;
  PdfInputs=$pdf.FullName; HasFreeRtos=[bool]($projectFiles.Name -match 'FreeRTOS|cmsis_os')
} | Format-List
foreach($pattern in $patterns){
  $hits = $source | Select-String -Pattern $pattern
  if($hits){ "--- $pattern"; $hits | Select-Object Path,LineNumber,Line }
}
