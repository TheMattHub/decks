# Builds the marketing demo from fictional content only: generated illustrations, a warm template generated
# from a palette, and a six-slide image-led campaign deck for "Northern Trails", an invented outdoor brand.
# Needs Windows with PowerPoint and uv. Run: pwsh -File examples\build_marketing_demo.ps1
$ErrorActionPreference = "Stop"
$skill = Join-Path $PSScriptRoot "..\skills\making-presentations"
. (Join-Path $skill "scripts\PptxCom.ps1")
$out = Join-Path $PSScriptRoot "..\docs\screenshots"
$work = Join-Path ([System.IO.Path]::GetTempPath()) "decks-demo-marketing"
$art = Join-Path $work "art"
New-Item -ItemType Directory -Force $out, $work | Out-Null
$out = (Resolve-Path $out).Path
uv run (Join-Path $PSScriptRoot "make_art.py") $art
function Art($name) { Join-Path $art "$name.jpg" }

$ink = "#2B2622"; $paper = "#FAF6F0"; $rust = "#C8553D"; $pine = "#2E6E5E"; $sand = "#E9DFD1"
$head = "Bahnschrift"; $body = "Segoe UI"
$tpl = Join-Path $work "Northern Trails.potx"
New-TemplateFromTheme $tpl $ink $paper @($rust, $pine, "#D9A441") $head $body "#4A403A" $sand

# A rectangle with no line; $alpha 0 = solid, 1 = clear.
function Panel($slide, $l, $t, $w, $h, $color, $alpha = 0) {
    $r = $slide.Shapes.AddShape(1, $l, $t, $w, $h)
    $r.Fill.ForeColor.RGB = ConvertTo-Bgr $color; $r.Fill.Transparency = [single]$alpha; $r.Line.Visible = 0
    $r
}
function Text($slide, $l, $t, $w, $h, $text, $font, $size, $color, [switch]$Bold) {
    $b = Add-TextBox $slide $l $t $w $h $text $font $size $color
    if ($Bold) { $b.TextFrame.TextRange.Font.Bold = -1 }
    $b
}

$s = New-DeckFromTemplate $tpl
try {
    $d = $s.Deck; $W = $d.PageSetup.SlideWidth; $H = $d.PageSetup.SlideHeight

    # 1. Cover: a full-bleed picture, a dark band for the title.
    $c = Add-LayoutSlide $d '^Blank$'
    [void](Add-PictureFill $c (Art "dawn") 0 0 $W $H)
    [void](Panel $c 0 300 $W 240 $ink 0.15)
    [void](Text $c 64 322 800 80 "Walk further, carry less" $head 52 $paper -Bold)
    [void](Text $c 66 404 800 40 "Northern Trails  |  spring campaign 2027" $body 20 $sand)
    Set-Notes $c "0:00. One line: the lightest pack we have made, for the walks people actually do."

    # 2. Who it is for: three picture cards.
    $p = Add-LayoutSlide $d '^Blank$'
    [void](Text $p 64 40 800 50 "Who we make it for" $head 34 $ink -Bold)
    $x = 64
    foreach ($card in @(@("forest", "Weekend walkers", "Two days, one bag, no planning spreadsheet."),
                        @("lake", "Long-distance hikers", "Every gram counted over 300 km."),
                        @("night", "Night owls", "Head out after work, back before midnight."))) {
        [void](Add-PictureFill $p (Art $card[0]) $x 112 262 200)
        [void](Panel $p $x 312 262 160 "#FFFFFF")
        [void](Text $p ($x + 18) 326 230 30 $card[1] $head 20 $ink -Bold)
        [void](Text $p ($x + 18) 362 230 80 $card[2] $body 15 "#4A403A")
        $x += 285
    }

    # 3. A customer voice: picture left, quote right.
    $q = Add-LayoutSlide $d '^Blank$'
    [void](Panel $q 0 0 $W $H $sand)
    [void](Add-PictureFill $q (Art "lake") 0 0 420 $H)
    [void](Text $q 470 150 430 50 ([string][char]0x201C) $head 96 $rust -Bold)
    [void](Text $q 470 230 430 160 "I forgot the pack was on my back until the lake." $head 32 $ink)
    [void](Text $q 470 400 430 30 "Ana, trail tester (fictional)" $body 16 "#6B5E55")

    # 4. Three numbers, native text.
    $n = Add-LayoutSlide $d '^Blank$'
    [void](Panel $n 0 0 $W $H $pine)
    [void](Text $n 64 50 800 50 "Lighter, tested, ready" $head 34 $paper -Bold)
    $x = 64
    foreach ($kpi in @(@("38%", "lighter than our last pack"), @("120 km", "of field tests per prototype"), @("4", "seasons, one shell"))) {
        [void](Text $n $x 190 260 110 $kpi[0] $head 72 $paper -Bold)
        [void](Text $n $x 310 240 60 $kpi[1] $body 18 $sand)
        $x += 290
    }

    # 5. Campaign plan: four native steps on a line.
    $t = Add-LayoutSlide $d '^Blank$'
    [void](Text $t 64 40 800 50 "Campaign plan" $head 34 $ink -Bold)
    $line = $t.Shapes.AddLine(110, 230, 850, 230); $line.Line.ForeColor.RGB = ConvertTo-Bgr $sand; $line.Line.Weight = 4
    $x = 90
    foreach ($step in @(@("March", "Teaser films on trails we know"), @("April", "Field days in four regions"),
                        @("May", "Launch with partner shops"), @("June", "Stories from first owners"))) {
        $dot = $t.Shapes.AddShape(9, $x, 210, 40, 40); $dot.Fill.ForeColor.RGB = ConvertTo-Bgr $rust; $dot.Line.Visible = 0
        [void](Text $t ($x - 30) 270 200 30 $step[0] $head 20 $ink -Bold)
        [void](Text $t ($x - 30) 304 190 80 $step[1] $body 15 "#4A403A")
        $x += 240
    }

    # 6. Close: full-bleed night, the ask.
    $e = Add-LayoutSlide $d '^Blank$'
    [void](Add-PictureFill $e (Art "night") 0 0 $W $H)
    [void](Text $e 64 330 800 70 "See you on the trail" $head 48 $paper -Bold)
    [void](Text $e 66 404 800 40 "Partner shop kits ship 1 March  |  hello@northerntrails.example" $body 18 $sand)

    Save-Deck $s (Join-Path $work "Northern Trails.pptx")
    Export-SlidePngs $d (Join-Path $work "png") 1600
} finally { Close-Deck $s }

# A 3 x 2 sheet of the six slides for the README.
Add-Type -AssemblyName System.Drawing
$pngs = @(Get-ChildItem (Join-Path $work "png") -Filter *.png | Sort-Object Name)
$w = 640; $h = 360; $gap = 14
$sheet = New-Object System.Drawing.Bitmap (3 * $w + 4 * $gap), (2 * $h + 3 * $gap)
$g = [System.Drawing.Graphics]::FromImage($sheet)
$g.Clear([System.Drawing.Color]::FromArgb(233, 223, 209))
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
for ($i = 0; $i -lt [math]::Min(6, $pngs.Count); $i++) {
    $img = [System.Drawing.Image]::FromFile($pngs[$i].FullName)
    $g.DrawImage($img, $gap + ($i % 3) * ($w + $gap), $gap + [math]::Floor($i / 3) * ($h + $gap), $w, $h)
    $img.Dispose()
}
$g.Dispose(); $sheet.Save((Join-Path $out "marketing-deck.png")); $sheet.Dispose()
Copy-Item (Join-Path $work "png\slide01.png") (Join-Path $out "marketing-cover.png") -Force
"images in $out"
