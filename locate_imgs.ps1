$ErrorActionPreference = "Continue"
$roots = @(
  "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal",
  "C:\Users\namma\Downloads"
)
$targets = @("08_open_aveva_control_plane_generated.png","16_continuous_naval_assurance_loop_generated.png","생성된 이미지 1.png","생성된 이미지 2.png")
foreach($r in $roots){
  if(Test-Path $r){
    foreach($t in $targets){
      $hits = Get-ChildItem -Path $r -Recurse -Filter $t -File -ErrorAction SilentlyContinue
      foreach($h in $hits){ "{0,-9}KB  {1}" -f [int]($h.Length/1KB), $h.FullName }
    }
  }
}
