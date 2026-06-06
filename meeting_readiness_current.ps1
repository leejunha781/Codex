$ErrorActionPreference="Stop"
$dir="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$src=Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$stamp=Get-Date -Format "yyyyMMdd_HHmmss"
$backupDir=Join-Path $dir ("_backup_"+$stamp)
$backup=Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_MEETING_REVIEW_BACKUP_"+$stamp+".pptx")
$work="C:\Users\namma\.claude\plm_slide_work\meeting_current"
$render=Join-Path $work "render_before"
$text=Join-Path $work "content.txt"
$audit=Join-Path $work "audit.txt"
New-Item -ItemType Directory -Force -Path $backupDir,$render|Out-Null
Get-ChildItem -LiteralPath $render -Filter *.png -ErrorAction SilentlyContinue|Remove-Item -Force
Copy-Item -LiteralPath $src -Destination $backup -Force

$ppt=$null;$pres=$null;$sb=New-Object Text.StringBuilder;$ab=New-Object Text.StringBuilder
try{
 $ppt=New-Object -ComObject PowerPoint.Application
 $pres=$ppt.Presentations.Open($src,$true,$false,$false)
 [void]$ab.AppendLine("SLIDES="+$pres.Slides.Count)
 foreach($sl in $pres.Slides){
  $title="";$pics=0;$tables=0;$texts=0;$chars=0
  [void]$sb.AppendLine(("===== SLIDE {0:D2} ====="-f$sl.SlideIndex))
  foreach($sh in $sl.Shapes){
   if($sh.Type-eq13){$pics++}
   try{if($sh.HasTable){$tables++}}catch{}
   try{
    if($sh.HasTextFrame-eq-1-and$sh.TextFrame.HasText-eq-1){
     $t=($sh.TextFrame.TextRange.Text-replace"[\r\n\v]"," ").Trim()
     if($t){if(-not$title){$title=$t};$texts++;$chars+=$t.Length;[void]$sb.AppendLine($t)}
    }
   }catch{}
  }
  [void]$ab.AppendLine(("[{0:D2}] pics={1} tables={2} textShapes={3} chars={4} :: {5}"-f$sl.SlideIndex,$pics,$tables,$texts,$chars,$title))
  $sl.Export((Join-Path $render ("slide-{0:D2}.png"-f$sl.SlideIndex)),"PNG",1280,720)
 }
}finally{
 if($pres){try{$pres.Close()}catch{}}
 if($ppt){try{if($ppt.Presentations.Count-eq0){$ppt.Quit()}}catch{}}
}
[IO.File]::WriteAllText($text,$sb.ToString(),[Text.Encoding]::UTF8)
[IO.File]::WriteAllText($audit,$ab.ToString(),[Text.Encoding]::UTF8)
Write-Output("BACKUP="+$backup);Write-Output("TEXT="+$text);Write-Output("AUDIT="+$audit);Write-Output("RENDER="+$render)
