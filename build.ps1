$ErrorActionPreference = "Stop"

$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt.pptx"
$out = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$png = "C:\Users\namma\.claude\plm_slide_work\slide19.png"

Copy-Item -LiteralPath $src -Destination $out -Force

function Rgb([string]$hex) {
    $hex = $hex.TrimStart('#')
    $r = [Convert]::ToInt32($hex.Substring(0,2),16)
    $g = [Convert]::ToInt32($hex.Substring(2,2),16)
    $b = [Convert]::ToInt32($hex.Substring(4,2),16)
    return [int]($r + $g*256 + $b*65536)
}

# ---- palette ----
$BG_TOP="16273E"; $BG_BOT="091120"
$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"
$CYAN="34C0EE"; $BLUE="4F8BE8"; $GREEN="49C28C"; $ORANGE="F0832F"
$PANEL="13243B"; $PANEL_LINE="2C4663"
$CARD="0F2034"; $CARD_LINE="2B5170"
$CODEBG="0A1626"; $CODE_LINE="27425C"; $CODE_TXT="D7E7FF"
$STEP="112338"; $STEP_LINE="2A4663"; $CHIPBG="17304C"

$ppt = New-Object -ComObject PowerPoint.Application
# PowerPoint is a single-instance COM server: this may ATTACH to the user's
# already-running instance (which has other presentations open). Therefore:
#  - open our file WITHOUT a window (WithWindow=msoFalse) so we don't disturb them
#  - never call $ppt.Quit() (that would close the user's other presentations)
$pres = $ppt.Presentations.Open($out, 0, 0, 0)  # ReadOnly=F, Untitled=F, WithWindow=F
$idx = [int]($pres.Slides.Count + 1)
$slide = $pres.Slides.Add($idx, 12)  # 12 = ppLayoutBlank

# ---------- helper functions ----------
function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }

function Panel($l,$t,$w,$h,$fill,$line,$lw,$trans,$rad){
    $s=$slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h)
    try{$s.Adjustments.Item(1)=[single]$rad}catch{}
    $s.Fill.Solid(); $s.Fill.ForeColor.RGB=[int](Rgb $fill); $s.Fill.Transparency=[single]$trans
    if($line){ $s.Line.Visible=-1; $s.Line.ForeColor.RGB=[int](Rgb $line); $s.Line.Weight=[single]$lw } else { $s.Line.Visible=0 }
    NoShadow $s
    return $s
}

function Bar($l,$t,$w,$h,$fill,$rad){
    $s=$slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h)
    try{$s.Adjustments.Item(1)=[single]$rad}catch{}
    $s.Fill.Solid(); $s.Fill.ForeColor.RGB=[int](Rgb $fill); $s.Line.Visible=0; NoShadow $s
    return $s
}

function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){
    $tb=$slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h)
    $tb.Fill.Visible=0; $tb.Line.Visible=0; NoShadow $tb
    $tf=$tb.TextFrame
    $tf.WordWrap=-1; $tf.AutoSize=0
    $tf.MarginLeft=[single]2; $tf.MarginRight=[single]2; $tf.MarginTop=[single]1; $tf.MarginBottom=[single]1
    $tf.VerticalAnchor=[int]$anchor
    $tr=$tf.TextRange; $tr.Text=$text
    $tr.Font.Size=[single]$size; $tr.Font.Bold=[int]$bold; $tr.Font.Name=[string]$font; $tr.Font.Color.RGB=[int](Rgb $color)
    $tr.ParagraphFormat.Alignment=[int]$align
    try{$tr.ParagraphFormat.SpaceBefore=[single]0; $tr.ParagraphFormat.SpaceAfter=[single]0}catch{}
    return $tb
}

function Chip($l,$t,$w,$h,$text,$size,$fill,$tcolor,$line,$font){
    $s=$slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h)
    try{$s.Adjustments.Item(1)=[single]0.35}catch{}
    $s.Fill.Solid(); $s.Fill.ForeColor.RGB=[int](Rgb $fill)
    if($line){ $s.Line.Visible=-1; $s.Line.ForeColor.RGB=[int](Rgb $line); $s.Line.Weight=[single]0.75 } else { $s.Line.Visible=0 }
    NoShadow $s
    $tf=$s.TextFrame; $tf.WordWrap=-1; $tf.AutoSize=0
    $tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]0;$tf.MarginBottom=[single]0; $tf.VerticalAnchor=[int]3
    $tr=$tf.TextRange; $tr.Text=$text; $tr.Font.Size=[single]$size; $tr.Font.Name=[string]$font
    $tr.Font.Color.RGB=[int](Rgb $tcolor); $tr.ParagraphFormat.Alignment=[int]2
    return $s
}

function Arrow($x1,$y1,$x2,$y2,$color,$w,$beg,$end){
    $c=$slide.Shapes.AddConnector(1,[single]$x1,[single]$y1,[single]$x2,[single]$y2)
    $c.Line.ForeColor.RGB=[int](Rgb $color); $c.Line.Weight=[single]$w
    $c.Line.BeginArrowheadStyle=[int]$beg; $c.Line.EndArrowheadStyle=[int]$end
    NoShadow $c
    return $c
}

function Oval($l,$t,$w,$h,$fill){
    $s=$slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h)
    $s.Fill.Solid(); $s.Fill.ForeColor.RGB=[int](Rgb $fill); $s.Line.Visible=0; NoShadow $s
    return $s
}

# ---------- background ----------
$bg=$slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540)
$bg.Line.Visible=0; NoShadow $bg
$bg.Fill.TwoColorGradient(2,1)
$bg.Fill.ForeColor.RGB=[int](Rgb $BG_TOP)
$bg.Fill.BackColor.RGB=[int](Rgb $BG_BOT)

# ---------- title row ----------
Txt 26 10 908 30 "Hardware  <->  Software (API) Architecture    &    HD Hyundai API Example" 22 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 26 42 908 16 "Organic data-center -> AVEVA PLM -> API stack, with a simple step-by-step client HD Hyundai can run and verify." 10.5 0 $LIGHT "Aptos" 1 1 | Out-Null

# =======================================================
#  REGION A (left): reference architecture stack
# =======================================================
$LX=26; $LW=494
$inL=36; $inW=476

Txt $LX 70 $LW 18 "A - Reference Architecture: organic Hardware <-> Software <-> API" 12 -1 $CYAN "Aptos" 1 1 | Out-Null

# ---- Layer 1: API ----
Panel $LX 96 $LW 104 $PANEL $PANEL_LINE 1 0.1 0.08 | Out-Null
Bar  $LX 96 6 104 $CYAN 0.5 | Out-Null
Txt $inL 100 $inW 16 "1.  API LAYER" 11 -1 $CYAN "Aptos" 1 1 | Out-Null
Txt $inL 100 $inW 16 "FastAPI  -  REST / OData  -  OpenAPI 3.1  -  OIDC (OAuth2)" 9 0 $LIGHT "Aptos" 3 1 | Out-Null
$ep=@("POST /oauth/token","GET /objects/{id}","POST /structures","POST /approvals")
$cw=[int][math]::Floor(($inW-18)/4); $cx=$inL
foreach($e in $ep){ Chip $cx 124 $cw 24 $e 8 $CHIPBG $CODE_TXT $CODE_LINE "Consolas" | Out-Null; $cx += $cw+6 }
Txt $inL 154 $inW 40 "Stateless, versioned, OpenAPI-documented endpoints, token-secured; every call returns JSON + an HTTP status, so each step is verifiable." 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null

# organic connection arrow (API <-> PLM)
Arrow 120 202 120 222 $MUTE 1.5 2 2 | Out-Null
Txt 132 200 380 22 "API exposes services   <->   AVEVA PLM returns data (JSON)" 8.5 0 $MUTE "Aptos" 1 3 | Out-Null

# ---- Layer 2: AVEVA PLM software ----
Panel $LX 224 $LW 104 $PANEL $PANEL_LINE 1 0.1 0.08 | Out-Null
Bar  $LX 224 6 104 $BLUE 0.5 | Out-Null
Txt $inL 228 $inW 16 "2.  AVEVA PLM SOFTWARE: application & engineering services" 10.5 -1 $BLUE "Aptos" 1 1 | Out-Null
$mods=@("Objects & Vault","Structures / eBOM","ECR & Impact","Approvals / Workflow","CAD / AVEVA E3D","Digital Thread")
$mcw=[int][math]::Floor(($inW-20)/3)
$mx=$inL; $my=252; $col=0
foreach($m in $mods){
    Chip $mx $my $mcw 22 $m 8.5 $CHIPBG $LIGHT $PANEL_LINE "Aptos" | Out-Null
    $col++; if($col -eq 3){ $col=0; $mx=$inL; $my=278 } else { $mx += $mcw+10 }
}
Txt $inL 304 $inW 20 "Versioned objects, structures/BOM & configuration rules, ECR/impact, approvals, all exposed through the API layer above." 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null

# organic connection arrow (PLM <-> Infra)
Arrow 120 330 120 350 $MUTE 1.5 2 2 | Out-Null
Txt 132 328 380 22 "PLM runs on   <->   data-center hosts, scales & secures it" 8.5 0 $MUTE "Aptos" 1 3 | Out-Null

# ---- Layer 3: Infrastructure / data center ----
Panel $LX 352 $LW 164 $PANEL $PANEL_LINE 1 0.1 0.08 | Out-Null
Bar  $LX 352 6 164 $GREEN 0.5 | Out-Null
Txt $inL 356 $inW 16 "3.  INFRASTRUCTURE - DATA CENTER (on-prem, HA, multi-yard, 1,000+ users)" 10 -1 $GREEN "Aptos" 1 1 | Out-Null

$cards=@(
 @("App / API Servers", @("4-node cluster - HA","16 vCPU - 64 GB/node","Linux - containers")),
 @("Database", @("PostgreSQL 16 - HA","32 vCPU - 256 GB","NVMe RAID10 + standby")),
 @("Vault / Storage", @("SAN (FC) + NAS","100 TB+ usable","Snapshots - WORM")),
 @("Gateway - Security", @("API GW + LB (HA)","TLS 1.3 - WAF","OIDC/OAuth2 - ZTNA"))
)
$kcw=[int][math]::Floor(($inW-30)/4); $kx=$inL; $ky=378; $kh=96
foreach($c in $cards){
    Panel $kx $ky $kcw $kh $CARD $CARD_LINE 0.75 0 0.08 | Out-Null
    Oval ($kx+8) ($ky+8) 8 8 $GREEN | Out-Null
    Txt ($kx+20) ($ky+5) ($kcw-24) 14 $c[0] 8.5 -1 $WHITE "Aptos" 1 1 | Out-Null
    $specTxt = ($c[1] -join "`r")
    $b=Txt ($kx+8) ($ky+24) ($kcw-14) ($kh-30) $specTxt 7.5 0 $LIGHT "Aptos" 1 1
    try{ $b.TextFrame.TextRange.ParagraphFormat.SpaceAfter=[single]3 }catch{}
    $kx += $kcw+10
}
Txt $inL 480 $inW 32 "Redundant 10/25 GbE backbone  -  Geo-DR site (async replica) + nightly/continuous backup  -  GPU node for 3D / digital-twin  -  Observability: metrics - logs - traces" 7.5 0 $MUTE "Aptos" 1 1 | Out-Null

# =======================================================
#  cross-link API -> client (short cyan arrow across the gap, no colliding label)
# =======================================================
Arrow 519 162 539 162 $CYAN 2.5 1 2 | Out-Null

# =======================================================
#  REGION B (right): HD Hyundai API example
# =======================================================
$RX=538; $RW=396

Txt $RX 70 $RW 30 "B - HD Hyundai: Using the API (simple, step-by-step, verifiable)" 12 -1 $ORANGE "Aptos" 1 1 | Out-Null

# ---- code window ----
$winT=108; $winH=132
Panel $RX $winT $RW $winH $CODEBG $CODE_LINE 1 0 0.05 | Out-Null
$tbar=$slide.Shapes.AddShape(1,[single]$RX,[single]$winT,[single]$RW,[single]18); $tbar.Line.Visible=0; NoShadow $tbar
$tbar.Fill.Solid(); $tbar.Fill.ForeColor.RGB=[int](Rgb "15273C")
Oval ($RX+10) ($winT+6) 7 7 "FF5F56" | Out-Null
Oval ($RX+22) ($winT+6) 7 7 "FFBD2E" | Out-Null
Oval ($RX+34) ($winT+6) 7 7 "27C93F" | Out-Null
Txt ($RX+48) ($winT+2) ($RW-56) 14 "plm_client.py  -  HD Hyundai PLM client  ->  calls REST API" 7.5 0 $MUTE "Consolas" 1 3 | Out-Null

$code = @(
 'import requests',
 'API = "https://plm.hhi.co.kr/api/v1"',
 'tok = requests.post(f"{API}/oauth/token",',
 '                    data=CRED).json()["access_token"]',
 'H = {"Authorization": f"Bearer {tok}"}',
 'r = requests.get(f"{API}/objects/HD-ETO-1001", headers=H)',
 'print(r.status_code, r.json()["name"])   # 200 OK -> verified'
) -join "`r"
$cb=Txt ($RX+10) ($winT+22) ($RW-20) ($winH-28) $code 8 0 $CODE_TXT "Consolas" 1 1
try{
  $tr=$cb.TextFrame.TextRange
  $tr.ParagraphFormat.SpaceBefore=[single]0; $tr.ParagraphFormat.SpaceAfter=[single]0
  $tr.Paragraphs(7,1).Font.Color.RGB=[int](Rgb $GREEN)
}catch{}

# ---- step flow ----
Txt $RX 246 $RW 14 "Each step calls one endpoint and returns a status you can check  (OK)" 8.5 0 $MUTE "Aptos" 1 1 | Out-Null

$steps=@(
 @("Authenticate (OIDC)","POST /oauth/token","200 - token  OK"),
 @("Read object","GET /objects/HD-ETO-1001","200 - object  OK"),
 @("Create structure (eBOM)","POST /structures","201 - struct_id  OK"),
 @("Raise ECR & run impact","POST /ecr/{id}/impact","200 - impact  OK"),
 @("Approve","POST /approvals","200 - APPROVED  OK")
)
$sy=262; $stH=44; $sgap=8; $n=1
foreach($st in $steps){
    Panel $RX $sy $RW $stH $STEP $STEP_LINE 0.75 0 0.12 | Out-Null
    Bar  $RX $sy 5 $stH $ORANGE 0.5 | Out-Null
    Oval ($RX+12) ($sy+9) 26 26 $ORANGE | Out-Null
    Txt ($RX+12) ($sy+9) 26 26 "$n" 12 -1 $WHITE "Aptos" 2 3 | Out-Null
    $tb=Txt ($RX+48) ($sy+5) 232 ($stH-8) ($st[0]+"`r"+$st[1]) 9 0 $WHITE "Aptos" 1 3
    try{
      $tr=$tb.TextFrame.TextRange
      $tr.ParagraphFormat.SpaceBefore=[single]0; $tr.ParagraphFormat.SpaceAfter=[single]0
      $p1=$tr.Paragraphs(1,1); $p1.Font.Size=[single]9.5; $p1.Font.Bold=[int](-1); $p1.Font.Color.RGB=[int](Rgb $WHITE); $p1.Font.Name="Aptos"
      $p2=$tr.Paragraphs(2,1); $p2.Font.Size=[single]8.5; $p2.Font.Bold=[int]0; $p2.Font.Color.RGB=[int](Rgb $CODE_TXT); $p2.Font.Name="Consolas"
    }catch{}
    Chip ($RX+288) ($sy+10) 100 24 $st[2] 8.5 "15402F" $GREEN $GREEN "Aptos" | Out-Null
    if($n -lt 5){ Arrow ($RX+25) ($sy+$stH) ($RX+25) ($sy+$stH+$sgap) $ORANGE 1.25 1 2 | Out-Null }
    $sy += $stH+$sgap; $n++
}

# ---------- bottom caption ----------
Txt 26 520 908 16 "Organic connection: data-center hardware hosts AVEVA PLM software -> AVEVA PLM exposes secured APIs -> HD Hyundai's client calls them in clear, status-verified steps." 8 -1 $MUTE "Aptos" 2 3 | Out-Null

# ---------- save + export ----------
$pres.Save()
$slide.Export($png,"PNG",1280,720)
$pres.Close()
# IMPORTANT: do NOT call $ppt.Quit() - the user has other presentations open in this instance.
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($slide) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "BUILT_OK $out"
