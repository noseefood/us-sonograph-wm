# Dump position / id / z-order of every top-level shape on every slide (via PowerPoint COM).
param([Parameter(Mandatory)][string]$Pptx, [Parameter(Mandatory)][string]$Work)
$ErrorActionPreference = "Stop"
$out = @()
$app = New-Object -ComObject PowerPoint.Application
$p = $app.Presentations.Open((Resolve-Path $Pptx).Path, $true, $false, $false)   # read-only, no window
try {
  foreach ($sl in $p.Slides) {
    $i = 0
    foreach ($sh in $sl.Shapes) {
      $i++
      $out += [pscustomobject]@{ slide=$sl.SlideIndex; idx=$i; id=$sh.Id; name=$sh.Name; type=$sh.Type; left=$sh.Left; top=$sh.Top; width=$sh.Width; height=$sh.Height; rot=$sh.Rotation; visible=$sh.Visible }
    }
  }
} finally { $p.Close(); $app.Quit() }
$out | ConvertTo-Json -Depth 3 | Out-File -Encoding utf8 (Join-Path $Work "shapes.json")
"dumped $($out.Count) shapes"
