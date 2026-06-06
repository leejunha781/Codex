$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"

# ---- palette (BGR ints) ----
$NAVY=3678739; $WHITE=16777215; $LIGHT=15062727; $MUTED=12626067
$CYAN=15646772; $AMBER=3846888; $GREEN=9224777; $DIVID=6178356
$DASH=[char]0x2014; $MID=[char]0x00B7

try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
catch { $ppt = New-Object -ComObject PowerPoint.Application }
$pres = $ppt.Presentations.Open($src, $false, $false, $false)

function Retry($block){ for($k=0;$k -lt 6;$k++){ try{ & $block; return }catch{ Start-Sleep -Milliseconds 120 } }; & $block }

function Add-Text($slide,$x,$y,$w,$h,$text,$font,$size,$bold,$color,$align){
  $tb=$slide.Shapes.AddTextbox(1,$x,$y,$w,$h)
  $tf=$tb.TextFrame; $tf.WordWrap=-1; $tf.AutoSize=0
  $tf.MarginLeft=0;$tf.MarginRight=0;$tf.MarginTop=0;$tf.MarginBottom=0
  $tr=$tf.TextRange; $tr.Text=$text
  $tr.Font.Name=$font; $tr.Font.Size=[single]$size; $tr.Font.Bold=[int]$bold
  Retry { $tr.Font.Color.RGB=[int]$color }
  if($align){ $tr.ParagraphFormat.Alignment=[int]$align }
  return $tb
}
function Add-Bar($slide,$x,$y,$w,$h,$color){
  $r=$slide.Shapes.AddShape(1,$x,$y,$w,$h)
  $r.Fill.Solid(); Retry { $r.Fill.ForeColor.RGB=[int]$color }
  $r.Line.Visible=0
  return $r
}
function Add-Entry($slide,$x,$y,$w,$h,$term,$def,$accent){
  $tb=$slide.Shapes.AddTextbox(1,$x,$y,$w,$h)
  $tf=$tb.TextFrame; $tf.WordWrap=-1; $tf.AutoSize=0
  $tf.MarginLeft=0;$tf.MarginRight=0;$tf.MarginTop=0;$tf.MarginBottom=0
  $tr=$tf.TextRange
  $tr.Text = $term + "   " + $DASH + "  " + $def
  $tr.Font.Name="Aptos"; $tr.Font.Size=[single]9.2; $tr.Font.Bold=[int]0
  Retry { $tr.Font.Color.RGB=[int]$LIGHT }
  $run=$tr.Characters(1,$term.Length)
  $run.Font.Bold=[int]1
  Retry { $run.Font.Color.RGB=[int]$accent }
  return $tb
}

function Build-Glossary($idx,$pageno,$title,$colAhdr,$colAcol,$colA,$colBhdr,$colBcol,$colB){
  $layout = $pres.Slides.Item(8).CustomLayout
  $slide = $pres.Slides.AddSlide($idx,$layout)
  $toDel=@(); foreach($sh in $slide.Shapes){ $toDel += $sh }
  foreach($sh in $toDel){ try{ $sh.Delete() }catch{} }
  $slide.FollowMasterBackground = $false
  $slide.Background.Fill.Solid(); Retry { $slide.Background.Fill.ForeColor.RGB=[int]$NAVY }

  Add-Text $slide 34 14 892 30 $title "Aptos Display" 21.5 -1 $WHITE 1 | Out-Null
  $sub = "Appendix " + $MID + " quick reference " + $DASH + " plain-language key to the abbreviations used throughout this deck, for the development and planning teams."
  Add-Text $slide 34 50 892 24 $sub "Aptos" 9.6 0 $LIGHT 1 | Out-Null
  Add-Text $slide 800 24 124 16 "Future Industrial PLM" "Aptos" 7.6 0 $MUTED 3 | Out-Null
  Add-Bar  $slide 34 80 892 1 $DIVID | Out-Null
  Add-Text $slide 900 8 26 16 ([string]$pageno) "Aptos" 6.5 0 $MUTED 3 | Out-Null

  Add-Bar  $slide 40  92 5 14 $colAcol | Out-Null
  Add-Text $slide 50  91 320 14 $colAhdr "Aptos" 8.5 -1 $colAcol 1 | Out-Null
  Add-Bar  $slide 490 92 5 14 $colBcol | Out-Null
  Add-Text $slide 500 91 320 14 $colBhdr "Aptos" 8.5 -1 $colBcol 1 | Out-Null

  $startY=116
  $stepA=[int]((498-$startY)/$colA.Count)
  for($i=0;$i -lt $colA.Count;$i++){
    $p=$colA[$i] -split '\|\|\|'
    $y=$startY + $i*$stepA
    Add-Entry $slide 46 $y 432 $stepA $p[0] $p[1] $colAcol | Out-Null
  }
  $stepB=[int]((498-$startY)/$colB.Count)
  for($i=0;$i -lt $colB.Count;$i++){
    $p=$colB[$i] -split '\|\|\|'
    $y=$startY + $i*$stepB
    Add-Entry $slide 496 $y 430 $stepB $p[0] $p[1] $colBcol | Out-Null
  }

  $foot = "Quick-reference appendix " + $MID + " abbreviations are also defined where first used in the deck " + $MID + " product names are trademarks of their respective owners."
  Add-Text $slide 34 514 892 14 $foot "Aptos" 7 0 $MUTED 1 | Out-Null
  return $slide
}

# ============ GLOSSARY 1 ============
$g1A = @(
 "PLM|||Product Lifecycle Management: one governed source for product data and change.",
 "BOM|||Bill of Materials: structured list of parts, quantities and relationships.",
 "eBOM / E-BOM|||Engineering BOM: the as-designed structure taken from the 3D model.",
 "M-BOM|||Manufacturing BOM: the as-built structure used on the shop floor.",
 "MTO|||Material Take-Off: quantities and materials extracted from the model.",
 "ECR / ECO|||Engineering Change Request / Order: proposed vs. approved-and-executed change.",
 "MCO / MOC|||Mfg Change Order / Management of Change: a production change and its governance.",
 "Effectivity|||Rule for when/where a configuration applies (hull / block / option / date).",
 "Baseline|||A frozen, released configuration snapshot.",
 "Digital thread|||Connected data trail linking design, build and operations.",
 "VCRM|||Verification Cross-Reference Matrix: maps each requirement to its evidence."
)
$g1B = @(
 "CoG|||Centre of Gravity: a key weight-and-stability quantity.",
 "NAPA|||Naval-architecture & stability software; the certified solver of record.",
 "OCX|||Open Class eXchange: open standard for sending hull data to class societies.",
 "IMO D-2|||IMO ballast-water discharge-quality performance standard.",
 "P&ID|||Piping & Instrumentation Diagram.",
 "HVAC|||Heating, Ventilation & Air-Conditioning.",
 "CFD / FEM|||Computational Fluid Dynamics / Finite-Element Method (engineering analysis).",
 "Hydrostatics / Stability|||Calculations proving the ship floats safely as loaded.",
 "Class society|||Body (e.g. DNV, LR, KR) that approves designs against rules.",
 "Loading case|||A defined load/ballast condition checked for stability.",
 "Digital twin|||Live data model mirroring the real asset's state."
)
$g1title  = "Glossary  1 / 2   " + $DASH + "   PLM, Engineering & Marine terms"
$g1Ahdr   = "PLM " + $MID + " CONFIGURATION " + $MID + " CHANGE"
$g1Bhdr   = "MARINE " + $MID + " NAVAL " + $MID + " CLASS"
Build-Glossary 31 31 $g1title $g1Ahdr $CYAN $g1A $g1Bhdr $AMBER $g1B | Out-Null

# ============ GLOSSARY 2 ============
$g2A = @(
 "AVEVA E3D / Marine|||Data-centric 3D ship & plant design.",
 "AVEVA Draw|||Produces 2D production drawings (DWG/DXF) from the 3D model.",
 "AVEVA AIM|||Asset Information Management: handover & operations data.",
 "AVEVA PI System|||Operational time-series (historian) platform.",
 "AVEVA CONNECT|||AVEVA's industrial-intelligence cloud.",
 "Teamcenter|||Siemens PLM platform (configuration benchmark).",
 "3DEXPERIENCE / ENOVIA|||Dassault collaboration & PLM platform.",
 "Windchill / Codebeamer|||PTC PLM / requirements-traceability tools.",
 "Hexagon SDx / Smart 3D / j5|||Asset-lifecycle & operations tools.",
 "ERP / MES|||Enterprise Resource Planning / Manufacturing Execution systems.",
 "DWG / DXF|||AutoCAD 2D drawing file formats."
)
$g2B = @(
 "API|||Application Programming Interface: how systems call each other.",
 "REST / OData|||Common web-API styles for reading and writing data.",
 "OpenAPI 3.1|||Machine-readable contract describing every endpoint.",
 "FastAPI|||The Python API framework used in the reference build.",
 "OIDC / OAuth2|||Identity login & authorization standards.",
 "JWT|||Signed access token proving who is calling.",
 "RBAC|||Role-Based Access Control: who may do what.",
 "PostgreSQL / Keycloak|||The reference database / identity server.",
 "E2E / CI|||End-to-End tests / Continuous Integration.",
 "HA / DR|||High Availability / Disaster Recovery.",
 "TLS / WAF / ZTNA|||Transport encryption / web firewall / zero-trust access.",
 "SAN / NAS / WORM|||Storage network / network storage / write-once retention."
)
$g2title = "Glossary  2 / 2   " + $DASH + "   Products, Architecture, Security & Infrastructure"
$g2Ahdr  = "PRODUCTS " + $MID + " PLATFORMS"
$g2Bhdr  = "ARCHITECTURE " + $MID + " SECURITY " + $MID + " INFRA"
Build-Glossary 32 32 $g2title $g2Ahdr $GREEN $g2A $g2Bhdr $CYAN $g2B | Out-Null

$pres.Save()
$cnt = $pres.Slides.Count
$pres.Close()
Write-Output ("BUILD DONE. Total slides now: " + $cnt)
