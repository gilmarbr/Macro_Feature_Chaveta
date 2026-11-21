param([string]$Saida = $PSScriptRoot)
Add-Type -AssemblyName System.Drawing
$ink = "C:\Program Files\Inkscape\bin\inkscape.exe"
$tmp = $env:TEMP
$svgOrig = [IO.File]::ReadAllText((Join-Path $Saida "Rasgo.svg"))

# Variante: tira o ponto solto (path2), preenche a face da ponta de branco
$svg = [regex]::Replace($svgOrig, '(?s)<circle\s+id="path2".*?/>', '')
$svg = $svg.Replace('style="fill:#000000;fill-opacity:0;stroke:#000000;stroke-width:0.2;stroke-linecap:round;stroke-linejoin:round"
       id="path1"', 'style="fill:#ffffff;fill-opacity:1;stroke:#000000;stroke-width:0.2;stroke-linecap:round;stroke-linejoin:round"
       id="path1"')

function Salvar24([Drawing.Bitmap]$b, [string]$p) {
    $o = New-Object Drawing.Bitmap $b.Width, $b.Height, ([Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $g = [Drawing.Graphics]::FromImage($o); $g.DrawImage($b, 0, 0, $b.Width, $b.Height); $g.Dispose()
    $o.Save($p, [Drawing.Imaging.ImageFormat]::Bmp); $o.Dispose()
}

# Monta um ícone NxN sobre a cor-chave (fora da silhueta) e branco quase puro (dentro)
function Compor([Drawing.Bitmap]$src, [int]$N, [Drawing.Color]$chave, [bool]$apagado) {
    $W = $src.Width; $H = $src.Height
    $ox = [int][math]::Floor(($N - $W) / 2); $oy = [int][math]::Floor(($N - $H) / 2)
    # máscara "fora": flood fill a partir da borda do quadrado, por pixels de alfa baixo
    $alfa = New-Object 'int[,]' $N, $N
    for ($y = 0; $y -lt $N; $y++) { for ($x = 0; $x -lt $N; $x++) {
        $sx = $x - $ox; $sy = $y - $oy
        if ($sx -ge 0 -and $sy -ge 0 -and $sx -lt $W -and $sy -lt $H) { $alfa[$x, $y] = $src.GetPixel($sx, $sy).A } } }
    $fora = New-Object 'bool[,]' $N, $N
    $fila = New-Object 'System.Collections.Generic.Queue[int]'
    for ($i = 0; $i -lt $N; $i++) { foreach ($q in @(@($i,0),@($i,($N-1)),@(0,$i),@(($N-1),$i))) {
        if ($alfa[$q[0],$q[1]] -lt 128 -and -not $fora[$q[0],$q[1]]) { $fora[$q[0],$q[1]] = $true; $fila.Enqueue($q[1]*$N+$q[0]) } } }
    while ($fila.Count) { $k = $fila.Dequeue(); $x = $k % $N; $y = [int][math]::Floor($k / $N)
        foreach ($d in @(@(1,0),@(-1,0),@(0,1),@(0,-1))) { $nx = $x+$d[0]; $ny = $y+$d[1]
            if ($nx -ge 0 -and $ny -ge 0 -and $nx -lt $N -and $ny -lt $N -and -not $fora[$nx,$ny] -and $alfa[$nx,$ny] -lt 128) {
                $fora[$nx,$ny] = $true; $fila.Enqueue($ny*$N+$nx) } } }
    $out = New-Object Drawing.Bitmap $N, $N
    for ($y = 0; $y -lt $N; $y++) { for ($x = 0; $x -lt $N; $x++) {
        if ($fora[$x,$y]) { $out.SetPixel($x, $y, $chave); continue }
        $sx = $x - $ox; $sy = $y - $oy
        $c = [Drawing.Color]::FromArgb(0,0,0,0)
        if ($sx -ge 0 -and $sy -ge 0 -and $sx -lt $W -and $sy -lt $H) { $c = $src.GetPixel($sx, $sy) }
        $a = $c.A / 255.0
        $r = $c.R*$a + 254*(1-$a); $gg = $c.G*$a + 254*(1-$a); $bb = $c.B*$a + 254*(1-$a)
        if ($apagado) { $l = 0.3*$r + 0.59*$gg + 0.11*$bb; $l = 110 + $l*0.55; $r = $l; $gg = $l; $bb = $l }
        $px = [Drawing.Color]::FromArgb([int][math]::Min(254,$r), [int][math]::Min(254,$gg), [int][math]::Min(254,$bb))
        if ($px.R -eq $chave.R -and $px.G -eq $chave.G -and $px.B -eq $chave.B) { $px = [Drawing.Color]::FromArgb($px.R, $px.G, [math]::Max(0,$px.B-1)) }
        $out.SetPixel($x, $y, $px) } }
    return $out
}

$branco = [Drawing.Color]::FromArgb(255,255,255)
$cinza  = [Drawing.Color]::FromArgb(192,192,192)
foreach ($N in 20, 32, 40) {
    $w = [math]::Max(0.2, 1.1 / ($N / 9.86)).ToString("0.###", [Globalization.CultureInfo]::InvariantCulture)
    $svgN = $svg.Replace("stroke-width:0.2;", "stroke-width:$w;")
    $fSvg = Join-Path $tmp "rasgo_$N.svg"; $fPng = Join-Path $tmp "rasgo_$N.png"
    [IO.File]::WriteAllText($fSvg, $svgN)
    & $ink $fSvg --export-type=png --export-area-drawing --export-width=$N --export-filename=$fPng 2>$null | Out-Null
    $src = [Drawing.Bitmap]::FromFile($fPng)
    # árvore: regular / suprimido / realçado
    $reg = Compor $src $N $branco $false; Salvar24 $reg (Join-Path $Saida "RasgoArvore_$N.bmp")
    $reg.Save((Join-Path $tmp "prev_$N.png")); Salvar24 $reg (Join-Path $Saida "RasgoArvore_${N}_realce.bmp"); $reg.Dispose()
    $sup = Compor $src $N $branco $true;  Salvar24 $sup (Join-Path $Saida "RasgoArvore_${N}_suprimido.bmp"); $sup.Dispose()
    # botão
    $bot = Compor $src $N $cinza $false; Salvar24 $bot (Join-Path $Saida "RasgoChaveta_$N.bmp"); $bot.Dispose()
    $src.Dispose()
    "ok $N (traço $w)"
}
