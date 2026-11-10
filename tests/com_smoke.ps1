# Smoke test of PptxCom.ps1 against PowerPoint itself: local only (CI has no PowerPoint). Builds throwaway
# decks in a temp folder, prints one PASS/FAIL line per check and exits 1 on any failure. Run it under both
#   pwsh -NoProfile -File tests\com_smoke.ps1
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests\com_smoke.ps1
$ErrorActionPreference = "Stop"
$repo = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$lib = Join-Path $repo "skills\making-presentations\scripts\PptxCom.ps1"
. $lib
$shell = (Get-Process -Id $PID).Path
$work = Join-Path ([System.IO.Path]::GetTempPath()) ("pptxcom-smoke-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory $work | Out-Null
$repoBefore = git -C $repo status --porcelain
$script:failed = 0

# A check passes when its block's last output is $true; anything else (or an exception) is shown as the reason.
function Check([string]$name, [scriptblock]$test) {
    try {
        $r = @(& $test)
        if ($r.Count -and $r[-1] -eq $true) { "PASS  $name" + $(if ($r.Count -gt 1) { "  ($($r[0..($r.Count - 2)] -join ' '))" }) } else { "FAIL  $name  ($($r -join ' '))"; $script:failed++ }
    } catch { "FAIL  $name  ($($_.Exception.Message))"; $script:failed++ }
}
function Throws([scriptblock]$block, [string]$pattern) {
    try { & $block | Out-Null; "did not throw" } catch { if ($_.Exception.Message -match $pattern) { $true } else { "threw: " + $_.Exception.Message } }
}
function Wait-NoPowerPoint($seconds) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while (@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -and $sw.Elapsed.TotalSeconds -lt $seconds) { Start-Sleep -Milliseconds 100 }
    if (@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count) { -1 } else { [math]::Round($sw.Elapsed.TotalSeconds, 1) }
}
# Runs a script in a new process of this same shell, so process start-up and shut-down are real.
function Invoke-Child([string]$code, [string[]]$arguments) {
    $f = Join-Path $work ("child-" + [guid]::NewGuid().ToString("N").Substring(0, 8) + ".ps1")
    Set-Content -LiteralPath $f -Value $code -Encoding ASCII
    $out = & $shell -NoProfile -ExecutionPolicy Bypass -File $f @arguments 2>&1
    if ($LASTEXITCODE) { throw "child failed: $($out -join ' ')" }
    $out
}

$ownPowerPoint = @(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
if ($ownPowerPoint) { "NOTE  PowerPoint was already running: process exit checks are skipped" }
"temp folder: $work"

# --- B04: sessions, release, retry, in fresh processes -------------------------------------------------
$tpl = Join-Path $work "Smoke.potx"
$base = Join-Path $work "base.pptx"
$base2 = Join-Path $work "base2.pptx"
$childHead = 'param($lib, $a, $b, $c); $ErrorActionPreference = "Stop"; . $lib' + "`n"
Check "B04 a script builds a template and a deck" {
    Invoke-Child ($childHead + @'
New-TemplateFromTheme $a "#1B2430" "#F4F1EA" @("#D9622B", "#2F6F8F") "Georgia" "Segoe UI"
foreach ($f in $b, $c) {
    $s = New-DeckFromTemplate $a
    try { $sl = Add-LayoutSlide $s.Deck '^Title only$'; Set-PlaceholderText $sl 'title' "Base deck"; Save-Deck $s $f } finally { Close-Deck $s }
}
'@) @($lib, $tpl, $base, $base2) | Out-Null
    (Test-Path $base) -and (Test-Path $base2)
}
Check "B04 a second process right after the first works without a wait" {
    $out = Invoke-Child ($childHead + @'
$s = Open-Deck $a -ReadOnly; try { "slides " + $s.Deck.Slides.Count } finally { Close-Deck $s }
'@) @($lib, $base)
    ($out -join '') -eq "slides 1"
}
Check "B04 two sessions: closing one keeps the other; close then open; PowerPoint exits within 5 s" {
    $out = Invoke-Child ($childHead + @'
$s1 = Open-Deck $a -ReadOnly; $s2 = Open-Deck $b -ReadOnly
Close-Deck $s1
"second " + $s2.Deck.Slides.Count
$sw = [System.Diagnostics.Stopwatch]::StartNew(); Close-Deck $s2; "close " + $sw.Elapsed.TotalSeconds
$s3 = Open-Deck $a -ReadOnly; "reopen " + $s3.Deck.Slides.Count; Close-Deck $s3
$sw = [System.Diagnostics.Stopwatch]::StartNew()
while (@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -and $sw.Elapsed.TotalSeconds -lt 15) { Start-Sleep -Milliseconds 100 }
"exit " + $sw.Elapsed.TotalSeconds
'@) @($lib, $base, $base2)
    $exit = [double](($out | Where-Object { $_ -match '^exit ' }) -replace '^exit ', '')
    if ($ownPowerPoint) { $exit = 0 }
    if (($out -contains "second 1") -and ($out -contains "reopen 1") -and $exit -le 5) { "exit after $exit s"; $true } else { $out -join '; ' }
}

# --- checks in this process ------------------------------------------------------------------------------
# In a function, so its COM references go out of scope before the final process check.
function Invoke-ComChecks {
    $s = New-DeckFromTemplate $tpl
    try {
        $d = $s.Deck

        # B03: groups (nested) and table cells
        $g = Add-LayoutSlide $d '^Blank$'
        [void](Add-TextBox $g 20 20 200 30 "Plain box" "Segoe UI" 18 "#000000")
        $a1 = Add-TextBox $g 20 60 200 30 "Group alpha" "Segoe UI" 18 "#000000"
        $a2 = Add-TextBox $g 20 100 200 30 "Group beta" "Segoe UI" 18 "#000000"
        [void]$g.Shapes.Range([object[]]@($a1.Name, $a2.Name)).Group()
        $n1 = Add-TextBox $g 300 60 200 30 "Nested gamma" "Segoe UI" 18 "#000000"
        $n2 = Add-TextBox $g 300 100 200 30 "Nested inner" "Segoe UI" 18 "#000000"
        $inner = $g.Shapes.Range([object[]]@($n1.Name, $n2.Name)).Group()
        $n3 = Add-TextBox $g 300 140 200 30 "Nested outer" "Segoe UI" 18 "#000000"
        [void]$g.Shapes.Range([object[]]@($inner.Name, $n3.Name)).Group()
        $t = $g.Shapes.AddTable(2, 2, 20, 300, 400, 80)
        $t.Table.Cell(2, 2).Shape.TextFrame.TextRange.Text = "Cell delta"
        Check "B03 Find-Slide finds text in a nested group" { (Find-Slide $d '^Nested gamma$').SlideIndex -eq $g.SlideIndex }
        Check "B03 Find-Slide finds text in a table cell" { (Find-Slide $d '^Cell delta$').SlideIndex -eq $g.SlideIndex }
        Check "B03 Find-Shape returns a grouped shape and a cell's shape" {
            ((Find-Shape $g 'beta').TextFrame.TextRange.Text -eq "Group beta") -and ((Find-Shape $g 'delta').TextFrame.TextRange.Text -eq "Cell delta")
        }
        Check "B03 Get-SlideTexts includes groups and cells" {
            $x = @(Get-SlideTexts $g); ($x -contains "Nested outer") -and ($x -contains "Cell delta") -and ($x -contains "Plain box")
        }
        Set-SlideRole $g "walker"
        Check "B03 Get-SlideLibrary lists grouped and cell text" {
            $l = @(Get-SlideLibrary $d | Where-Object { $_.Role -eq "walker" })[0]
            (($l.Shapes -join '|') -match 'Nested gamma') -and (($l.Shapes -join '|') -match 'Cell delta')
        }
        $g.Name = "walker-done"

        # B02: keys resolve before any write
        $p = Add-LayoutSlide $d '^Blank$'
        [void](Add-TextBox $p 20 20 300 40 "Alpha" "Segoe UI" 20 "#000000")
        [void](Add-TextBox $p 20 80 300 40 "Beta" "Segoe UI" 20 "#000000")
        Set-SlideRole $p "pair"
        Check "B02 swapped keys give 'Beta one', 'Alpha two'" {
            $new = Add-FromLibrary $d "pair" @{ '^Alpha' = 'Beta one'; '^Beta' = 'Alpha two' }
            (@(Get-SlideTexts $new) -join '|') -eq "Beta one|Alpha two"
        }
        Check "B02 a key matching nothing throws and leaves no copy" {
            $n = $d.Slides.Count
            $r = Throws { Add-FromLibrary $d "pair" @{ '^Alpha' = 'x'; '^Gamma' = 'y' } } "matches '\^Gamma'"
            ($r -eq $true) -and ($d.Slides.Count -eq $n)
        }
        Check "B02 two keys on one shape throw and leave no copy" {
            $n = $d.Slides.Count
            $r = Throws { Add-FromLibrary $d "pair" @{ '^Al' = 'x'; '^Alp' = 'y' } } "same shape"
            ($r -eq $true) -and ($d.Slides.Count -eq $n)
        }
        Remove-LibrarySlides $d

        # B06: Edit-DeckText
        $e = Add-LayoutSlide $d '^Blank$'
        $box = Add-TextBox $e 20 20 600 40 "Revenue grew 38% in Q3" "Segoe UI" 20 "#000000"
        $box.TextFrame.TextRange.Characters(14, 3).Font.Bold = -1
        $e2 = Add-LayoutSlide $d '^Blank$'
        [void](Add-TextBox $e2 20 20 600 40 "Same line" "Segoe UI" 20 "#000000")
        $e3 = Add-LayoutSlide $d '^Blank$'
        [void](Add-TextBox $e3 20 20 600 40 "Same line" "Segoe UI" 20 "#000000")
        Check "B06 replaces once and keeps the bold run" {
            $sh = Edit-DeckText $d '38%' '41%'
            $r = $sh.TextFrame.TextRange
            ($r.Text -eq "Revenue grew 41% in Q3") -and ($r.Characters(14, 3).Font.Bold -eq -1) -and ($r.Characters(1, 7).Font.Bold -eq 0)
        }
        Check "B06 a second run throws 'not found'" { Throws { Edit-DeckText $d '38%' '41%' } 'not found' }
        Check "B06 case-sensitive" { Throws { Edit-DeckText $d 'revenue' 'x' } 'not found' }
        Check "B06 two occurrences throw naming both slides" {
            Throws { Edit-DeckText $d 'Same line' 'x' } ("slides {0}, {1}" -f $e2.SlideIndex, $e3.SlideIndex)
        }
        Check "B06 a slide as target, and a table cell" {
            $a = Edit-DeckText $e3 'Same line' 'Other line'
            $b = Edit-DeckText $d 'Cell delta' 'Cell epsilon'
            ($a.TextFrame.TextRange.Text -eq "Other line") -and ((Find-Shape $g 'epsilon').TextFrame.TextRange.Text -eq "Cell epsilon")
        }

        # B07: the shared leftover cases, exclude, extra patterns, comments
        $cases = @([System.IO.File]::ReadAllLines((Join-Path $repo "tests\leftover_cases.txt"), [System.Text.Encoding]::UTF8) |
            ForEach-Object { $_.TrimEnd("`r") } | Where-Object { $_ -match "^(hit|ok)`t" } | ForEach-Object {
                $k, $v = $_ -split "`t", 2; [pscustomobject]@{ Hit = $k -eq "hit"; Text = $v } })
        $lo = Add-LayoutSlide $d '^Blank$'
        $y = 0
        foreach ($c in $cases) { [void](Add-TextBox $lo 20 $y 900 20 $c.Text "Segoe UI" 10 "#000000"); $y += 22 }
        [void]$lo.Comments.Add2(10, 10, "Reviewer", "R", "Check this number", "AD", "reviewer")
        $found = @(Test-DeckLeftovers $d | Where-Object { $_.Slide -eq $lo.SlideIndex })
        Check "B07 agrees with tests/leftover_cases.txt ($($cases.Count) cases)" {
            $wrong = @(foreach ($c in $cases) { if ([bool](@($found | Where-Object { $_.Where -ne "comments" -and $_.Text -eq $c.Text }).Count) -ne $c.Hit) { $c.Text } })
            if ($wrong.Count) { "disagree: " + ($wrong -join " | ") } else { $true }
        }
        Check "B07 one finding for a slide with comments" {
            $c = @($found | Where-Object { $_.Where -eq "comments" }); ($c.Count -eq 1) -and ($c[0].Match -eq 1)
        }
        Check "B07 exclude drops a hit, extra patterns add" {
            $f = @(Test-DeckLeftovers $d @('northern') @('^lorem') | Where-Object { $_.Slide -eq $lo.SlideIndex })
            (-not @($f | Where-Object { $_.Text -match '^Lorem' }).Count) -and [bool]@($f | Where-Object { $_.Text -match 'northern depot' }).Count
        }
        Check "B07 finds leftovers in groups and cells" {
            $gs = Add-LayoutSlide $d '^Blank$'
            $q1 = Add-TextBox $gs 20 20 200 30 "TODO in a group" "Segoe UI" 18 "#000000"
            $q2 = Add-TextBox $gs 20 60 200 30 "fine" "Segoe UI" 18 "#000000"
            [void]$gs.Shapes.Range([object[]]@($q1.Name, $q2.Name)).Group()
            $tb = $gs.Shapes.AddTable(1, 1, 20, 200, 300, 40); $tb.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Lorem cell"
            $f = @(Test-DeckLeftovers $d | Where-Object { $_.Slide -eq $gs.SlideIndex })
            $ok = ($f.Count -eq 2) -and [bool]@($f | Where-Object { $_.Where -eq ("shape {0} 1,1" -f $tb.Name) }).Count
            $gs.Delete(); $ok
        }

        # B08a: stale renders go, numbers are padded
        $png = Join-Path $work "png"
        New-Item -ItemType Directory -Force $png | Out-Null
        foreach ($n in "slide01.png", "slide2.png", "slide99.png", "keep.txt") { Set-Content (Join-Path $png $n) "old" }
        Check "B08a Export-SlidePngs replaces old renders, two-digit names" {
            Export-SlidePngs $d $png 320
            $want = @(1..$d.Slides.Count | ForEach-Object { "slide{0:D2}.png" -f $_ })
            $got = @(Get-ChildItem $png -Filter "slide*.png" | Sort-Object Name | ForEach-Object { $_.Name })
            $fresh = @(Get-ChildItem $png -Filter "slide*.png" | Where-Object { $_.Length -le 10 }).Count -eq 0
            (($got -join ',') -eq ($want -join ',')) -and $fresh -and (Test-Path (Join-Path $png "keep.txt"))
        }

        # W-B: Save-Deck checks itself and records the deck
        Check "W-B Save-Deck without a path throws a clear message" { Throws { Save-Deck $s } 'no path' }
        Check "W-B Save-Deck writes <name>.dump.txt: hash, then the dump with fonts, groups and cells" {
            Save-Deck $s $wb
            $rec = @(Get-Content -LiteralPath (Join-Path $work "wb.dump.txt") -Encoding UTF8)
            ($rec[0] -eq "sha256 " + (Get-FileHash $wb -Algorithm SHA256).Hash.ToLower()) -and
                [bool]($rec -match '^  \[.+\] type 17 .+  Segoe UI 18  Plain box$') -and
                [bool]($rec -match '^    \[.+\]  Segoe UI 18  Nested gamma$') -and
                [bool]($rec -match '^    \[.+ 2,2\]  .+ Cell epsilon$')
        }
        Check "W-B after Save-Deck to a path, the session is that file: Get-DeckHandEdits reads its record" {
            ($s.Path -eq $wb) -and (@(Get-DeckHandEdits $s).Count -eq 0)
        }

        # W-C: contact sheet
        Check "W-C Export-ContactSheet: one labelled tile per slide, sheet of the expected size" {
            $d.Slides.Item(2).SlideShowTransition.Hidden = -1
            $f = Join-Path $work "contact.png"
            Export-ContactSheet $d $f 4 240
            $rows = [math]::Ceiling($d.Slides.Count / 4)
            (Get-PngSize $f) -eq ("{0}x{1}" -f (4 * 252 + 12), ($rows * (135 + 28 + 12) + 12))
        }

        # F4: PDF, slides and notes pages
        # Twice each on one open deck: PowerPoint itself writes only the first PDF saved from an open deck.
        $shown = @(@($d.Slides) | Where-Object { $_.SlideShowTransition.Hidden -ne -1 }).Count
        Check "F4 Export-DeckPdf: a page per shown slide, landscape, twice; the deck stays the .pptx" {
            $d.PrintOptions.OutputType = 2                       # a setting of the deck's own, to see it untouched
            $r = foreach ($n in 1, 2) {
                $f = Join-Path $work "deck$n.pdf"; Export-DeckPdf $d $f; $p = Get-PdfPages $f
                ($p.Count -eq $shown) -and ($p.Landscape -eq $true)
            }
            "$($r -join ',')"
            ($r -notcontains $false) -and ($d.FullName -eq $wb) -and ($d.PrintOptions.OutputType -eq 2)
        }
        Check "F4 Export-DeckPdf -Notes: notes pages (portrait), twice" {
            $r = foreach ($n in 1, 2) {
                $f = Join-Path $work "notes$n.pdf"; Export-DeckPdf $d $f -Notes; $p = Get-PdfPages $f
                ($p.Count -eq $shown) -and ($p.Landscape -eq $false)
            }
            "$($r -join ',')"
            ($r -notcontains $false) -and ($d.FullName -eq $wb) -and ($d.PrintOptions.OutputType -eq 2)
        }
        Check "F4 Export-DeckPdf -Notes without the PDF printer: one clear line, nothing written" {
            $PdfPrinter = "No Such Printer 7"           # shadows the helper's name for this block only
            $f = Join-Path $work "noprinter.pdf"; $tmpBefore = @(Get-ChildItem ([System.IO.Path]::GetTempPath()) -Filter "deckpdf-*").Count
            $r = Throws { Export-DeckPdf $d $f -Notes } "'No Such Printer 7' printer, which is not installed"
            ($r -eq $true) -and -not (Test-Path $f) -and ($d.FullName -eq $wb) -and
                (@(Get-ChildItem ([System.IO.Path]::GetTempPath()) -Filter "deckpdf-*").Count -eq $tmpBefore)
        }
    } finally { Close-Deck $s }
}
function Get-PngSize([string]$path) {
    Add-Type -AssemblyName System.Drawing
    $i = [System.Drawing.Image]::FromFile($path); try { "{0}x{1}" -f $i.Width, $i.Height } finally { $i.Dispose() }
}
# Page count and orientation of the first page, read from the PDF's page objects (PowerPoint writes them
# uncompressed).
function Get-PdfPages([string]$path) {
    $text = [System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($path))
    if (-not $text.StartsWith("%PDF")) { throw "$path is not a PDF" }
    $box = [regex]::Match($text, '/MediaBox\s*\[\s*([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*\]')
    [pscustomobject]@{ Count = [regex]::Matches($text, '/Type\s*/Page(?![s\w])').Count
        Landscape = ([double]$box.Groups[3].Value - [double]$box.Groups[1].Value) -gt ([double]$box.Groups[4].Value - [double]$box.Groups[2].Value) }
}
$wb = Join-Path $work "wb.pptx"
Invoke-ComChecks

function Invoke-HandEditChecks {
    Check "W-B Get-DeckHandEdits: nothing for an unchanged deck" {
        $s = Open-Deck $wb; try { @(Get-DeckHandEdits $s).Count -eq 0 } finally { Close-Deck $s }
    }
    Check "W-B Get-DeckHandEdits: the changed lines after a hand edit" {
        $s = Open-Deck $wb
        try { [void](Edit-DeckText $s.Deck 'Plain box' 'Plain hand box'); $s.Deck.Save() } finally { Close-Deck $s }
        $s = Open-Deck $wb -ReadOnly
        try { $e = @(Get-DeckHandEdits $s) } finally { Close-Deck $s }
        $ok = ($e.Count -eq 2) -and ($e[0] -match '^- .+  Plain box$') -and ($e[1] -match '^\+ .+  Plain hand box$')
        if ($ok) { $true } else { $e -join ' | ' }
    }
    Check "W-B Get-DeckHandEdits throws without a record" {
        Remove-Item (Join-Path $work "wb.dump.txt")
        $s = Open-Deck $wb -ReadOnly; try { Throws { Get-DeckHandEdits $s } 'no record' } finally { Close-Deck $s }
    }
    Check "W-B Open-Deck's backup of a .potx is a .potx" {
        $s = Open-Deck $tpl; Close-Deck $s
        @(Get-ChildItem $work -Filter "Smoke.bak-*.potx").Count -eq 1
    }
    Check "W-C Export-LayoutSheet: same sheet size as before the shared tile function" {
        $f = Join-Path $work "layouts.png"
        $s = Open-Deck $tpl -ReadOnly; try { Export-LayoutSheet $s.Deck $f } finally { Close-Deck $s }
        (Get-PngSize $f) -eq "1488x942"        # 9 layouts, 3 across at 480 px
    }
}
Invoke-HandEditChecks

# --- W-A: Test-DeckLayout ------------------------------------------------------------------------------
function Invoke-LayoutChecks {
    $s = New-DeckFromTemplate $tpl
    try {
        $d = $s.Deck
        $ok = Add-LayoutSlide $d '^Blank$'
        [void](Add-TextBox $ok 0 0 960 540 "Full slide, 12 pt" "Segoe UI" 12 "#000000")
        $t = $ok.Shapes.AddTable(1, 1, 20, 100, 300, 20)
        $t.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = ("A cell grows with its text. " * 6)
        $rot = Add-TextBox $ok 800 100 300 40 "Rotated past the edge" "Segoe UI" 14 "#000000"; $rot.Rotation = 30
        [void](Add-PictureFill $ok (Join-Path $work "png\slide01.png") -20 300 400 260)     # a photo bleeding off the edge
        $bad = Add-LayoutSlide $d '^Blank$'
        $small = Add-TextBox $bad 20 20 300 40 "Fine " "Segoe UI" 14 "#000000"
        $small.TextFrame.TextRange.InsertAfter("tiny").Font.Size = 10
        $over = Add-TextBox $bad 20 100 200 60 ("Too much text for this box. " * 12) "Segoe UI" 14 "#000000"
        $off = Add-TextBox $bad -10 300 200 40 "Off the left" "Segoe UI" 14 "#000000"
        $shrink = Add-LayoutSlide $d '^Title and content$'
        Set-PlaceholderText $shrink 'title' "Too many bullets"
        Set-PlaceholderText $shrink 'body' ((1..25 | ForEach-Object { "Bullet line number $_ with some words" }) -join "`n")
        $f = @(Test-DeckLayout $d)
        Check "W-A overflow in a body placeholder that shrinks text on overflow" {
            $x = @($f | Where-Object { $_.Slide -eq $shrink.SlideIndex })
            ($x | ForEach-Object { "$($_.Check): $($_.Detail)" }) -join '; '
            ($x.Count -eq 1) -and ($x[0].Check -eq "overflow") -and ($x[0].Where -eq "shape " + (Get-Placeholder $shrink 'body').Name)
        }
        $shrink.Delete()
        Check "W-A nothing on a clean slide (full-bleed box, growing table cell, rotated shape, bleeding photo)" {
            $x = @($f | Where-Object { $_.Slide -eq $ok.SlideIndex }); if ($x.Count) { ($x | ForEach-Object { "$($_.Check) $($_.Where) $($_.Detail)" }) -join '; ' } else { $true }
        }
        Check "W-A small-text, overflow and off-slide, one each" {
            $x = @($f | Where-Object { $_.Slide -eq $bad.SlideIndex })
            ($x | ForEach-Object { "$($_.Check): $($_.Detail)" }) -join '; '
            ($x.Count -eq 3) -and [bool]@($x | Where-Object { $_.Check -eq "small-text" -and $_.Where -eq "shape " + $small.Name }).Count -and
                [bool]@($x | Where-Object { $_.Check -eq "overflow" -and $_.Where -eq "shape " + $over.Name }).Count -and
                [bool]@($x | Where-Object { $_.Check -eq "off-slide" -and $_.Where -eq "shape " + $off.Name }).Count
        }
        Check "W-A the minimum scales with the slide width" {
            # COM leaves text sizes as they are when the slide grows: at 1920 pt, 6 means 12 and 7 means 14.
            $d.PageSetup.SlideWidth = 1920; $d.PageSetup.SlideHeight = 1080
            $x = @(Test-DeckLayout $d 6 | Where-Object { $_.Check -eq "small-text" })     # the 10 pt run
            $y = @(Test-DeckLayout $d 7 | Where-Object { $_.Check -eq "small-text" })     # and the 12 pt box
            "6: " + (($x | ForEach-Object { $_.Detail }) -join ', ') + "; 7: " + (($y | ForEach-Object { $_.Detail }) -join ', ')
            ($x.Count -eq 1) -and ($y.Count -eq 2)
        }
    } finally { Close-Deck $s }
}
Invoke-LayoutChecks

# --- F1: starter library -------------------------------------------------------------------------------
function Invoke-StarterChecks {
    $libPath = Join-Path $work "Smoke library.pptx"
    New-StarterLibrary $tpl $libPath
    $roles = "cover", "section", "cards", "big-numbers", "quote", "timeline", "closing"
    # A 4:3 copy of the template: x and y scale differently there.
    $tpl43 = Join-Path $work "Smoke43.pptx"; $lib43 = Join-Path $work "Smoke43 library.pptx"
    $s = New-DeckFromTemplate $tpl
    try { $s.Deck.PageSetup.SlideWidth = 720; $s.Deck.PageSetup.SlideHeight = 540; Save-Deck $s $tpl43 } finally { Close-Deck $s }
    New-StarterLibrary $tpl43 $lib43
    $s = Open-Deck $lib43 -ReadOnly
    try {
        Check "F1 on a 4:3 template the library's circles stay round, and it has no layout findings" {
            $tl = @(@($s.Deck.Slides) | Where-Object { (Get-SlideRole $_) -eq "timeline" })[0]
            $ovals = @(@($tl.Shapes) | Where-Object { $_.Type -eq 1 -and $_.AutoShapeType -eq 9 } | ForEach-Object { "{0:N1}x{1:N1}" -f $_.Width, $_.Height })
            $f = @(Test-DeckLayout $s.Deck)
            ($ovals -join ' ') + ($f | ForEach-Object { "; $($_.Slide) $($_.Check) $($_.Detail)" })
            ($ovals.Count -eq 4) -and -not @($ovals | Where-Object { $_ -ne "24.0x24.0" }).Count -and -not $f.Count
        }
    } finally { Close-Deck $s }
    $s = New-DeckFromTemplate $libPath -KeepSlides
    try {
        $d = $s.Deck
        Check "F1 the library itself has no layout findings; the cards have an outline (light 2 may be the background)" {
            $f = @(Test-DeckLayout $d)
            $cards = @(@($d.Slides) | Where-Object { (Get-SlideRole $_) -eq "cards" })[0]
            $outlined = @(@($cards.Shapes) | Where-Object { $_.Type -eq 1 -and $_.Line.Visible -eq -1 }).Count
            "outlined $outlined" + ($f | ForEach-Object { "; $($_.Slide) $($_.Check) $($_.Detail)" })
            (-not $f.Count) -and ($outlined -eq 3)
        }
        Check "F1 New-StarterLibrary tags one pattern per role, every sample text starting [placeholder]" {
            $l = @(Get-SlideLibrary $d)
            $bad = @(foreach ($p in $l) { foreach ($t in @(Get-SlideTexts $d.Slides.Item($p.Slide))) { if ($t -and -not $t.StartsWith("[placeholder]") -and $t -ne [string][char]0x201C) { $t } } })
            (($l | ForEach-Object { $_.Role }) -join ',') -eq ($roles -join ',') -and -not $bad.Count
        }
        Check "F1 a deck built from every role, every placeholder replaced: no leftovers, no layout findings" {
            foreach ($role in $roles) {
                # The pattern's own texts without the prefix: realistic lengths for the boxes.
                $pattern = @(@($d.Slides) | Where-Object { (Get-SlideRole $_) -eq $role })[0]
                $texts = @{}
                foreach ($t in @(Get-SlideTexts $pattern)) { if ($t.StartsWith("[placeholder] ")) { $texts['^' + [regex]::Escape($t) + '$'] = $t.Substring(14) } }
                [void](Add-FromLibrary $d $role $texts)
            }
            Remove-LibrarySlides $d
            Export-ContactSheet $d (Join-Path $work "starter.png")
            $f = @(Test-DeckLeftovers $d) + @(Test-DeckLayout $d)
            if ($f.Count) { ($f | ForEach-Object { "$($_.Slide) $($_.Where) $($_.Check)$($_.Match) $($_.Detail)$($_.Text)" }) -join '; ' }
            else { ($d.Slides.Count -eq $roles.Count) }
        }
    } finally { Close-Deck $s }
}
Invoke-StarterChecks

# --- F2: Invoke-FinishManifest -------------------------------------------------------------------------
function Invoke-ManifestChecks {
    $media = Join-Path $work "media"
    New-Item -ItemType Directory -Force $media | Out-Null
    # A real clip: PowerPoint refuses a file that is not media. ffmpeg comes with imageio-ffmpeg, through uv.
    # Under Windows PowerShell, a native command's stderr line is an error, which "Stop" makes fatal.
    $ff = & { $ErrorActionPreference = "Continue"; uv run --with imageio-ffmpeg python -c "import imageio_ffmpeg; print(imageio_ffmpeg.get_ffmpeg_exe())" 2>$null }
    if ($ff) { & $ff -y -loglevel error -f lavfi -i testsrc=size=320x180:rate=10 -t 1 -pix_fmt yuv420p (Join-Path $media "ab12.mp4") }
    $long = "Costs fell 12% in the third quarter, with every depot on the new schedule from July"
    $s = New-DeckFromTemplate $tpl
    try {
        $d = $s.Deck
        $one = Add-LayoutSlide $d '^Blank$'; [void](Add-TextBox $one 40 40 800 60 "Q3   review" "Segoe UI" 32 "#000000")
        $two = Add-LayoutSlide $d '^Blank$'; [void](Add-TextBox $two 40 40 800 60 "Results" "Segoe UI" 32 "#000000")
        $cost = Add-TextBox $two 40 140 800 60 $long "Segoe UI" 18 "#000000"
        $g1 = Add-TextBox $two 40 260 300 40 "Grouped point" "Segoe UI" 18 "#000000"
        $g2 = Add-TextBox $two 40 300 300 40 "Its label" "Segoe UI" 18 "#000000"
        $grp = $two.Shapes.Range([object[]]@($g1.Name, $g2.Name)).Group()
        $three = Add-LayoutSlide $d '^Blank$'; [void](Add-TextBox $three 40 40 800 60 "Film" "Segoe UI" 32 "#000000")
        $n38 = Add-TextBox $three 40 400 200 60 "38%" "Segoe UI" 32 "#000000"
        [void](Add-TextBox $three 300 400 400 60 "38% of trucks" "Segoe UI" 32 "#000000")
        Set-Content -LiteralPath (Join-Path $media "broken.mp4") -Value "not a video"
        function Write-Manifest($slides) {
            $f = Join-Path $work ("manifest-" + [guid]::NewGuid().ToString("N").Substring(0, 6) + ".json")
            [System.IO.File]::WriteAllText($f, (@{ deck = "Q3 review"; canvas_width = 1920; canvas_height = 1080; slides = $slides } | ConvertTo-Json -Depth 6))
            $f
        }
        $good = Write-Manifest @(
            @{ index = 1; id = "cover"; title = "Q3 review"; transition = "fade"; builds = @(); videos = @() },
            @{ index = 2; id = "results"; title = "Results"; transition = "push"; videos = @()
               builds = @(@{ effect = "rise"; text = $long.Substring(0, 60) }, @{ effect = "pop"; text = "Grouped point" }, @{ effect = "fade"; text = "" }) },
            @{ index = 3; id = "film"; title = "Film"; transition = "magic"; builds = @(@{ effect = "fade"; text = "38%" })
               videos = @(@{ src = "/_blob/ab12"; poster = "/_blob/missing"; box = @{ x = 192; y = 108; w = 960; h = 540 } }) })
        $unreadable = Write-Manifest @(
            @{ index = 1; id = "cover"; title = "Q3 review"; transition = "fade"; builds = @(); videos = @() },
            @{ index = 2; id = "results"; title = "Results"; transition = $null; videos = @(); builds = @(@{ effect = "pop"; text = "Grouped point" }) },
            @{ index = 3; id = "film"; title = "Film"; transition = $null; builds = @()
               videos = @(@{ src = "/_blob/ab12"; poster = $null; box = $null }, @{ src = "/_blob/broken"; poster = $null; box = $null }) })
        $bad = Write-Manifest @(
            @{ index = 1; id = "cover"; title = "Q3 overview"; transition = "spin"; builds = @(); videos = @() },
            @{ index = 2; id = "results"; title = "Results"; transition = $null; videos = @(); builds = @(@{ effect = "rise"; text = "Revenue" }) },
            @{ index = 3; id = "film"; title = ""; transition = $null; builds = @(); videos = @(@{ src = "/_blob/nothere"; poster = $null; box = $null }) },
            @{ index = 4; id = "extra"; title = "Extra"; transition = $null; builds = @(); videos = @() })
        Check "F2 a manifest that cannot be applied throws once, listing every item, and changes nothing" {
            $before = @(Get-DeckDump $d) -join "`n"
            $msg = try { Invoke-FinishManifest $d $bad $media; "did not throw" } catch { $_.Exception.Message }
            $want = "3 slides, the manifest 4|slide 1: title 'Q3 overview'|unknown transition 'spin'|0 shapes start with 'Revenue'|no file for video '/_blob/nothere'|slide 4 \(extra\): not in"
            $missing = @($want -split '\|' | Where-Object { $msg -notmatch $_ })
            if ($missing.Count) { "missing: " + ($missing -join ', ') + " in: " + $msg } else { (@(Get-DeckDump $d) -join "`n") -eq $before }
        }
        Check "F2 a video file PowerPoint cannot read throws and changes nothing (the video before it goes too)" {
            if (-not $ff) { return "no ffmpeg (uv with imageio-ffmpeg) to make a clip" }
            $before = @(Get-DeckDump $d) -join "`n"
            $msg = try { Invoke-FinishManifest $d $unreadable $media; "did not throw" } catch { $_.Exception.Message }
            $after = @(Get-DeckDump $d) -join "`n"
            if ($msg -notmatch "slide 3: PowerPoint cannot insert '.+broken\.mp4'") { $msg } else { $after -eq $before }
        }
        Check "F2 transitions, builds by text (a 60-character prefix, a group, '38%' beside '38% of trucks'), the video in its scaled box" {
            if (-not $ff) { return "no ffmpeg (uv with imageio-ffmpeg) to make a clip" }
            $warn = @(Invoke-FinishManifest $d $good $media 3>&1)          # it outputs nothing but warnings
            $seq = $two.TimeLine.MainSequence
            $v = @(@($three.Shapes) | Where-Object { $_.Type -eq $MsoMedia })
            "warnings $($warn.Count)"
            ($one.SlideShowTransition.EntryEffect -eq $PpFade) -and ($two.SlideShowTransition.EntryEffect -eq $PpPushUp) -and
                ($three.SlideShowTransition.EntryEffect -eq $PpMorph) -and ($seq.Count -eq 2) -and
                ($seq.Item(1).Shape.Name -eq $cost.Name) -and ($seq.Item(1).EffectType -eq $MsoRiseUp) -and
                ($seq.Item(2).Shape.Name -eq $grp.Name) -and ($seq.Item(2).EffectType -eq $MsoZoom) -and
                ($v.Count -eq 1) -and ([math]::Round($v[0].Left) -eq 96) -and ([math]::Round($v[0].Top) -eq 54) -and
                ([math]::Round($v[0].Width) -eq 480) -and ($warn.Count -eq 2) -and
                [bool]@(@($three.TimeLine.MainSequence) | Where-Object { $_.Shape.Name -eq $n38.Name -and $_.EffectType -eq $MsoFadeIn }).Count
        }
    } finally { Close-Deck $s }
}
Invoke-ManifestChecks

Check "W-B a record written under the other shell gives no hand edits under this one" {
    $other = if ($PSVersionTable.PSEdition -eq "Core") { "powershell.exe" } else { "pwsh" }
    $f = Join-Path $work "shells.pptx"
    $code = Join-Path $work "shells.ps1"
    # 112.49999 is where the two shells' "{0:N0}" disagreed (113 against 112).
    Set-Content -LiteralPath $code -Encoding ASCII -Value ($childHead + @'
$s = New-DeckFromTemplate $a
try { $sl = Add-LayoutSlide $s.Deck '^Blank$'; $x = Add-TextBox $sl 20 20 300 40 "Rounding" "Segoe UI" 20 "#000000"
      $x.Left = [single]112.49999; $x.Width = [single]200.5; Save-Deck $s $b } finally { Close-Deck $s }
'@)
    $log = & $other -NoProfile -ExecutionPolicy Bypass -File $code $lib $tpl $f 2>&1
    if ($LASTEXITCODE) { return "child failed: $($log -join ' ')" }
    # The dumps themselves: Get-DeckHandEdits would stop at the unchanged file hash.
    $was = @(Get-Content -LiteralPath (Join-Path $work "shells.dump.txt") -Encoding UTF8 | Select-Object -Skip 1)
    $s = Open-Deck $f -ReadOnly; try { $now = @(Get-DeckDump $s.Deck) } finally { Close-Deck $s }
    $e = @(Compare-Object $was $now -CaseSensitive | ForEach-Object { $_.SideIndicator + " " + $_.InputObject })
    if ($e.Count) { $e -join ' | ' } else { @($now -match 'Rounding$') -join ''; $true }
}

# --- the example scripts, writing into temp ------------------------------------------------------------
foreach ($ex in @(@("build_demo.ps1", "demo-deck.png", "1648x948", "layout-sheet.png", "1308x840", "equation-slide.png"),
                  @("build_marketing_demo.ps1", "marketing-deck.png", "1976x762", "marketing-cover.png"))) {
    Check "examples\$($ex[0]) -Out <temp>" {
        $dir = Join-Path $work ("ex-" + $ex[0])
        $log = & $shell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo "examples\$($ex[0])") -Out $dir 2>&1
        if ($LASTEXITCODE) { return "exit $LASTEXITCODE`: $($log -join ' ')" }
        $missing = @($ex[1..($ex.Count - 1)] | Where-Object { $_ -match '\.png$' -and -not (Test-Path (Join-Path $dir $_)) })
        if ($missing.Count) { return "missing $($missing -join ', ')" }
        $wrong = @(for ($i = 1; $i -lt $ex.Count - 1; $i++) {
            if ($ex[$i + 1] -match '^\d+x\d+$' -and (Get-PngSize (Join-Path $dir $ex[$i])) -ne $ex[$i + 1]) { "$($ex[$i]) is $(Get-PngSize (Join-Path $dir $ex[$i]))" }
        })
        if ($wrong.Count) { $wrong -join ', ' } else { $true }
    }
}

# --- G1, G2: a PowerPoint the helper did not start is never quit or stopped ------------------------------
# Only when no PowerPoint ran before the test: the stand-in for the person's PowerPoint would otherwise be
# the person's real one.
function Invoke-GuardChecks {
    Check "G2 Stop-PowerPoint stops the instance its own process started, unsaved work and all" {
        [void](Wait-NoPowerPoint 15)         # the examples' PowerPoint may still be closing
        $out = Invoke-Child ($childHead + @'
$app = Get-PowerPointApp; $p = $app.Presentations.Add(0); [void]$p.Slides.Add(1, 12)
$id = @(Get-Process POWERPNT)[0].Id
Stop-PowerPoint
if (Get-Process -Id $id -ErrorAction SilentlyContinue) { "running" } else { "stopped" }
'@) @($lib)
        $out -join ' '; ($out -join ' ') -eq "stopped"
    }
    if ((Wait-NoPowerPoint 15) -lt 0) { "FAIL  G1/G2 skipped: a PowerPoint is still running"; $script:failed++; return }
    # PowerPoint is one instance: a person who opens it while a script runs joins the script's process.
    Check "G3 the person joins the script's PowerPoint: Close-Deck does not quit it, Stop-PowerPoint refuses" {
        try {
            $s = Open-Deck $base -ReadOnly
            $own = [string]$script:PptxComStartedPid
            [void](Invoke-Child '$app = New-Object -ComObject PowerPoint.Application; $app.Visible = -1' @())   # on its start screen
            Close-Deck $s
            # A quitting PowerPoint refuses the person's next call, which then starts a second process.
            [void](Invoke-Child '$app = New-Object -ComObject PowerPoint.Application; $p = $app.Presentations.Add(-1); [void]$p.Slides.Add(1, 12)' @())
            $joined = (@(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) -join ',') -eq $own
            $r = Throws { Stop-PowerPoint } 'visible window'
            "joined $joined, Stop-PowerPoint $r"
            $joined -and ($r -eq $true) -and ((@(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) -join ',') -eq $own)
        } finally {
            # The stand-in is the test's own: its draft closes unsaved and it quits through COM, from a child
            # process, whose COM references all go when it ends (under Windows PowerShell one left in this
            # process kept PowerPoint running).
            if (@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count) {
                [void](Invoke-Child '$app = New-Object -ComObject PowerPoint.Application; foreach ($p in @($app.Presentations)) { $p.Saved = -1; $p.Close() }; $app.Quit()' @())
            }
        }
    }
    if ((Wait-NoPowerPoint 15) -lt 0) { "FAIL  G1/G2 skipped: the G3 stand-in is still running"; $script:failed++; return }
    # The stand-in: a PowerPoint with a window, started before the helper attaches, on its start screen.
    $person = New-Object -ComObject PowerPoint.Application
    $person.Visible = -1
    $draft = $null
    try {
        Check "G1 Close-Deck leaves running a PowerPoint the helper did not start, even on its start screen" {
            $s = Open-Deck $base -ReadOnly; $was = $s.WasRunning; Close-Deck $s
            Start-Sleep -Seconds 2
            $was -and (@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -eq 1) -and ($person.Presentations.Count -eq 0)
        }
        $draft = $person.Presentations.Add(-1); [void]$draft.Slides.Add(1, 12)
        Check "G2 Stop-PowerPoint throws for the person's PowerPoint, naming the unsaved presentation, and leaves it" {
            $r = Throws { Stop-PowerPoint } ("unsaved changes in " + [regex]::Escape($draft.FullName))
            $r; ($r -eq $true) -and ($draft.Slides.Count -eq 1)
        }
    } finally {
        if ($draft) { $draft.Close(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($draft) }
        $person.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($person)
        $draft = $null; $person = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    }
}
if (-not $ownPowerPoint) { Invoke-GuardChecks }


# --- the end: no PowerPoint left, nothing written outside temp ------------------------------------------
[GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
if (-not $ownPowerPoint) {
    Check "no PowerPoint process left" { $t = Wait-NoPowerPoint 15; if ($t -ge 0) { "after $t s"; $true } else { "still running after 15 s" } }
    # Never Stop-Process: Stop-PowerPoint refuses when the instance has a window or unsaved work.
    if (@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count) { try { Stop-PowerPoint } catch { "NOTE  " + $_.Exception.Message } }
}
Check "nothing written in the repository" { $now = git -C $repo status --porcelain; (($now -join "`n") -eq ($repoBefore -join "`n")) }
if ($script:failed) { "$script:failed check(s) failed"; exit 1 }
Remove-Item -Recurse -Force $work
"all checks passed"
