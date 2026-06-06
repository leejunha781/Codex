$ErrorActionPreference="Stop"
trap {
 Write-Output ("ERROR_LINE="+$_.InvocationInfo.ScriptLineNumber)
 Write-Output ("ERROR_TEXT="+$_.InvocationInfo.Line)
 Write-Output ("ERROR_MESSAGE="+$_.Exception.Message)
 break
}
$dir="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$deck=Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$stamp=Get-Date -Format "yyyyMMdd_HHmmss"
$backupDir=Join-Path $dir ("_backup_"+$stamp)
$backup=Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_BEFORE_NATIVE_COMPARISON_TABLES_"+$stamp+".pptx")
$renderDir="C:\Users\namma\.claude\plm_slide_work\meeting_current\comparison_tables"
New-Item -ItemType Directory -Force -Path $backupDir,$renderDir|Out-Null
Copy-Item -LiteralPath $deck -Destination $backup -Force

function Rgb([string]$h){$h=$h.TrimStart('#');[int]([Convert]::ToInt32($h.Substring(0,2),16)+[Convert]::ToInt32($h.Substring(2,2),16)*256+[Convert]::ToInt32($h.Substring(4,2),16)*65536)}
$WHITE="F4F7FA";$LIGHT="C7D6E5";$MUTED="93A8C0";$BG="132238";$PANEL="172C45";$HEAD="0D5067";$LINE="6D879F"
$CYAN="31C5F4";$GREEN="49C28C";$AMBER="E8B23A";$ORANGE="F0832F";$VIOLET="9B7BE0";$RED="E85A5A"
function StatusStyle($cell,[string]$value){
 $v=$value.ToLower()
 if($v-like"*strong - new*" -or$v-like"*federated*"){$color=$CYAN;$fill="10394A"}
 elseif($v-like"*strong*"){$color=$GREEN;$fill="123D36"}
 elseif($v-like"*partial*"){$color=$AMBER;$fill="3A3218"}
 elseif($v-like"*limited*"){$color=$ORANGE;$fill="3D2718"}
 else{$color=$LIGHT;$fill=$PANEL}
 $cell.Shape.Fill.Solid();$cell.Shape.Fill.ForeColor.RGB=[int](Rgb $fill)
 $cell.Shape.TextFrame.TextRange.Font.Color.RGB=[int](Rgb $color)
}
function SetCell($cell,$text,$size,$bold,$color,$fill,$align){
 $cell.Shape.Fill.Solid();$cell.Shape.Fill.ForeColor.RGB=[int](Rgb $fill)
 try{$cell.Borders.Item(1).ForeColor.RGB=[int](Rgb $LINE);$cell.Borders.Item(1).Weight=[single]0.55}catch{}
 try{$cell.Borders.Item(2).ForeColor.RGB=[int](Rgb $LINE);$cell.Borders.Item(2).Weight=[single]0.55}catch{}
 try{$cell.Borders.Item(3).ForeColor.RGB=[int](Rgb $LINE);$cell.Borders.Item(3).Weight=[single]0.55}catch{}
 try{$cell.Borders.Item(4).ForeColor.RGB=[int](Rgb $LINE);$cell.Borders.Item(4).Weight=[single]0.55}catch{}
 $tf=$cell.Shape.TextFrame;$tf.MarginLeft=4;$tf.MarginRight=4;$tf.MarginTop=2;$tf.MarginBottom=1
 $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Name="Aptos";$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Color.RGB=[int](Rgb $color);$tr.ParagraphFormat.Alignment=[int]$align
}
function BuildTable($sl,$x,$y,$w,$h,$headers,$rows,$widths){
 $tableShape=$sl.Shapes.AddTable($rows.Count+1,$headers.Count,[single]$x,[single]$y,[single]$w,[single]$h)
 $tableShape.Name=("NativeComparisonTable_"+$sl.SlideIndex)
 $tb=$tableShape.Table
 for($c=1;$c-le$headers.Count;$c++){$tb.Columns.Item($c).Width=[single]$widths[$c-1];SetCell $tb.Cell(1,$c) $headers[$c-1] 8.0 1 $WHITE $HEAD 2}
 for($r=0;$r-lt$rows.Count;$r++){
  for($c=0;$c-lt$headers.Count;$c++){
   $txt=[string]$rows[$r][$c]
   if($c-eq0){SetCell $tb.Cell($r+2,$c+1) $txt 7.3 1 $WHITE $PANEL 1}
   else{
    $symbol=if($txt-like"*Strong - new*" -or$txt-like"*Federated*"){"◆"}elseif($txt-like"*Strong*"){"●"}elseif($txt-like"*Partial*"){"◐"}else{"○"}
    SetCell $tb.Cell($r+2,$c+1) ($symbol+"  "+$txt) 7.1 1 $LIGHT $PANEL 2
    StatusStyle $tb.Cell($r+2,$c+1) $txt
   }
  }
 }
 return $tableShape
}
function DeleteComparisonBody($sl,$startTop,$endTop){
 for($i=$sl.Shapes.Count;$i-ge1;$i--){
  $sh=$sl.Shapes.Item($i)
  if($sh.Top-ge$startTop-and$sh.Top-lt$endTop){$sh.Delete()}
 }
}

$ppt=$null;$pres=$null
try{
 $ppt=New-Object -ComObject PowerPoint.Application;$pres=$ppt.Presentations.Open($deck,$false,$false,$false)

 # Slide 3
 $s=$pres.Slides.Item(3);DeleteComparisonBody $s 105 430
 $headers=@("Capability","AVEVA`nE3D/Marine + AIM/PI","Siemens / Dassault`nPLM platforms","Hexagon / CADMATIC`n/ FORAN","NAPA")
 $rows=@(
  @("Configuration / BOM / change","Partial","Strong","Partial","Limited"),
  @("Hull / outfitting / production","Strong","Partial","Strong","Partial"),
  @("Naval architecture / stability","Partial","Limited","Partial","Strong"),
  @("Operations / asset context","Strong","Partial","Strong","Partial"),
  @("Customer-owned API / process","Partial","Partial","Partial","Partial")
 )
 BuildTable $s 38 112 886 308 $headers $rows @(176,186,186,186,152)|Out-Null

 # Slide 26
 $s=$pres.Slides.Item(26);DeleteComparisonBody $s 100 452
 $headers=@("Capability","AVEVA + Continuous`nNaval Assurance","NAPA","Siemens / Dassault`nPLM platforms","Other ship`ndesign tools")
 $rows=@(
  @("Certified stability / hydrostatics","Federated","Strong","Limited","Partial"),
  @("Hull + outfitting + production","Strong","Partial","Partial","Strong"),
  @("Configuration / change / effectivity","Strong","Partial","Strong","Partial"),
  @("Cross-domain assurance impact graph","Strong - new","Partial","Partial","Limited"),
  @("Evidence + approval orchestration","Strong - new","Partial","Strong","Partial"),
  @("Operations calibration / learning","Strong - new","Partial","Partial","Partial"),
  @("Customer-owned API / process","Strong","Partial","Partial","Partial")
 )
 BuildTable $s 36 108 888 338 $headers $rows @(196,208,162,162,160)|Out-Null

 $pres.Save()
 foreach($i in 3,26){$pres.Slides.Item($i).Export((Join-Path $renderDir ("slide-{0:D2}.png"-f$i)),"PNG",1600,900)}
}finally{if($pres){try{$pres.Close()}catch{}};if($ppt){try{if($ppt.Presentations.Count-eq0){$ppt.Quit()}}catch{}}}
Write-Output("BACKUP="+$backup);Write-Output("RENDER="+$renderDir)
