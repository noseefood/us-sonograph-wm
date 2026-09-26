# Export every planned layer of every slide as SVG via PowerPoint (hidden Shape.Export, format 6 = SVG).
# Two fully transparent 1pt anchors at opposite slide corners are added to each range so that every
# export covers the whole slide in the same 96-dpi pixel frame (otherwise PowerPoint picks the frame per range).
# The deck is opened read-only and never saved, so the source .pptx is untouched.
param([Parameter(Mandatory)][string]$Pptx, [Parameter(Mandatory)][string]$Work)
$ErrorActionPreference = "Stop"
$outDir = Join-Path $Work "exp"
New-Item -ItemType Directory -Force $outDir | Out-Null
$plan = Get-Content (Join-Path $Work "plan.json") -Raw | ConvertFrom-Json
$app = New-Object -ComObject PowerPoint.Application
$p = $app.Presentations.Open((Resolve-Path $Pptx).Path, $true, $false, $false)
try {
  $W = $p.PageSetup.SlideWidth; $H = $p.PageSetup.SlideHeight
  foreach ($prop in $plan.PSObject.Properties) {
    $n = [int]$prop.Name; $sl = $p.Slides($n); $k = 0
    $a1 = $sl.Shapes.AddShape(1, 0, 0, 1, 1); $a2 = $sl.Shapes.AddShape(1, $W - 1, $H - 1, 1, 1)
    foreach ($a in $a1, $a2) { $a.Line.Visible = 0; $a.Fill.ForeColor.RGB = 16777215; $a.Fill.Transparency = 1.0 }
    $ia = $sl.Shapes.Count - 1; $ib = $sl.Shapes.Count
    foreach ($L in $prop.Value.layers) {
      if ($L.kind -eq "svg") {
        $idx = [int[]](@($L.indices) + @($ia, $ib))
        $sl.Shapes.Range($idx).Export((Join-Path $outDir ("s{0}_{1}.svg" -f $n, $k)), 6, 0, 0, 0)
      }
      $k++
    }
    "slide $n exported"
  }
} finally { $p.Close(); $app.Quit() }
