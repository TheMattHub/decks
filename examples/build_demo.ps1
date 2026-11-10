# Builds the README images from fictional content only: a template generated from a palette, a short deck
# on it (native text, a big-number row, an equation), the deck rendered, and the template's layout sheet.
# Needs Windows with PowerPoint and uv. Run: pwsh -File examples\build_demo.ps1 [-Out folder]
param([string]$Out)     # where the images go; docs\screenshots by default
$ErrorActionPreference = "Stop"
$skill = Join-Path $PSScriptRoot "..\skills\making-presentations"
. (Join-Path $skill "scripts\PptxCom.ps1")
if (-not $Out) { $Out = Join-Path $PSScriptRoot "..\docs\screenshots" }
$work = Join-Path ([System.IO.Path]::GetTempPath()) "decks-demo"
New-Item -ItemType Directory -Force $out, $work | Out-Null
$out = (Resolve-Path $out).Path

# 1. A template from a look: colours and fonts in, a .potx out.
$tpl = Join-Path $work "Northwind.potx"
New-TemplateFromTheme $tpl "#1B2430" "#F4F1EA" @("#D9622B", "#2F6F8F", "#8A9A5B") "Georgia" "Segoe UI" "#2B3644" "#E6E1D6"

# 2. A four-slide deck on it, every word native.
$s = New-DeckFromTemplate $tpl
try {
    $d = $s.Deck
    $c = Add-LayoutSlide $d '^Title$'
    Set-PlaceholderText $c 'title' "Fewer trucks, same day"
    Set-PlaceholderText $c 'subtitle' "Northwind Freight - fleet study, a fictional example"

    $a = Add-LayoutSlide $d '^Title only$'
    Set-PlaceholderText $a 'title' "Four depots reach their bound with 17 trucks"
    $x = 66
    foreach ($card in @(@("North", "4", "trucks, 312 min"), @("East", "5", "trucks, 287 min"),
                        @("South", "3", "trucks, 198 min"), @("West", "5", "trucks, 341 min"))) {
        $bg = $a.Shapes.AddShape(5, $x, 170, 192, 230)
        $bg.Fill.ForeColor.RGB = ConvertTo-Bgr "#1B2430"; $bg.Line.Visible = 0
        [void](Add-TextBox $a ($x + 18) 186 160 30 $card[0] "Segoe UI" 16 "#E6E1D6")
        [void](Add-TextBox $a ($x + 18) 220 160 90 $card[1] "Georgia" 66 "#D9622B")
        [void](Add-TextBox $a ($x + 18) 330 160 40 $card[2] "Segoe UI" 16 "#F4F1EA")
        $x += 210
    }
    Set-Notes $a "1:00. Each depot reaches its lower bound; one truck fewer costs 2 to 6 minutes."

    $m = Add-LayoutSlide $d '^Title and content$'
    Set-PlaceholderText $m 'title' "Every parcel rides exactly one truck"
    Set-PlaceholderText $m 'body' "Assignment, written as LaTeX and converted to native PowerPoint math:"
    $eq = Add-TextBox $m 66 260 828 80 '$$\sum_{t=1}^{T} y_{pt} = 1 \quad \forall p$$' "Cambria Math" 32 "#1B2430"
    $eq.TextFrame.TextRange.ParagraphFormat.Alignment = 2

    $b = Add-LayoutSlide $d '^Section$'
    Set-PlaceholderText $b 'title' "Backup"
    foreach ($sl in @($d.Slides)) { Remove-EmptyPlaceholders $sl }

    Save-Deck $s (Join-Path $work "Northwind.pptx")
} finally { Close-Deck $s }

# 3. Equations to native math, then render the slides.
uv run (Join-Path $skill "scripts\pptx_equations.py") (Join-Path $work "Northwind.pptx") (Join-Path $work "Northwind-eq.pptx")
$s = Open-Deck (Join-Path $work "Northwind-eq.pptx") -ReadOnly
try { Export-SlidePngs $s.Deck (Join-Path $work "png") 1600 } finally { Close-Deck $s }

# 4. The template's layouts as one labelled sheet.
$s = Open-Deck $tpl -ReadOnly
try { Export-LayoutSheet $s.Deck (Join-Path $out "layout-sheet.png") 3 420 } finally { Close-Deck $s }

# 5. The deck as a 2 x 2 hero image.
$pngs = @(Get-ChildItem (Join-Path $work "png") -Filter *.png | Sort-Object Name | Select-Object -First 4 | ForEach-Object { $_.FullName })
Write-TileSheet $pngs (Join-Path $out "demo-deck.png") 2 800 450 16 -background "#E6E1D6"
Copy-Item (Join-Path $work "png\slide03.png") (Join-Path $out "equation-slide.png") -Force
"images in $out"
