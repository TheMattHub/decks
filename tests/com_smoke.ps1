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

Check "B08b `$PptxComVersion equals plugin.json's version" {
    $PptxComVersion -eq (Get-Content -Raw (Join-Path $repo ".claude-plugin\plugin.json") | ConvertFrom-Json).version
}

# --- the end: no PowerPoint left, nothing written outside temp ------------------------------------------
[GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
if (-not $ownPowerPoint) {
    Check "no PowerPoint process left" { $t = Wait-NoPowerPoint 15; if ($t -ge 0) { "after $t s"; $true } else { "still running after 15 s" } }
}
Check "nothing written in the repository" { $now = git -C $repo status --porcelain; (($now -join "`n") -eq ($repoBefore -join "`n")) }
if ($script:failed) { "$script:failed check(s) failed"; exit 1 }
Remove-Item -Recurse -Force $work
"all checks passed"
