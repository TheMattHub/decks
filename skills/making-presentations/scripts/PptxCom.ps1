# Helpers for finishing and editing a .pptx with PowerPoint itself (COM, Windows), so every video,
# transition, build-in and embedded font is native. Copy it into the project's slides/ folder and
# dot-source it from a deck script there:
#
#   . "$PSScriptRoot\PptxCom.ps1"
#   $s = Open-Deck "$PSScriptRoot\deck.pptx"        # read-write: refuses if the file is open, backs it up first
#   try { $slide = Find-Slide $s.Deck '^THE ANSWER$'; Set-Notes $slide "2:30. Say the number." ; Save-Deck $s }
#   finally { Close-Deck $s }
#
# Keep this file ASCII: Windows PowerShell 5.1 reads a script without a BOM as ANSI, so write
# non-ASCII text in deck scripts as [char]0x00B7 (middle dot) and the like.

$PptxComVersion = "0.3.0"     # the plugin's version (tests/check_ps.ps1 keeps them equal)

# PowerPoint enumerations the deck scripts use.
$PpFade = 3849; $PpPushUp = 3848; $PpMorph = 3953          # ppEntryEffect (slide transitions)
$MsoFadeIn = 10; $MsoRiseUp = 34; $MsoZoom = 23            # msoAnimEffect (build-ins)
$OnClick = 1; $WithPrevious = 2; $AfterPrevious = 3         # msoAnimTrigger
$MsoPicture = 13; $MsoMedia = 16                           # msoShapeType
$PdfPrinter = "Microsoft Print to PDF"                     # Export-DeckPdf -Notes prints to it

# "#RRGGBB" -> the BGR integer COM expects in .RGB (a plain 0xRRGGBB gives the wrong colour).
function ConvertTo-Bgr([string]$hex) {
    $h = $hex.TrimStart('#')
    [int]("0x" + $h.Substring(4, 2) + $h.Substring(2, 2) + $h.Substring(0, 2))
}

# Opens a deck without a window. Read-write unless -ReadOnly; read-write refuses when another program
# (usually PowerPoint with the deck open) holds the file, and copies it to <name>.bak-<time>.pptx first.
function Open-Deck([string]$path, [switch]$ReadOnly) {
    $path = (Resolve-Path $path).Path
    if (-not $ReadOnly) {
        try { $fs = [System.IO.File]::Open($path, 'Open', 'ReadWrite', 'None'); $fs.Close() }
        catch { throw "$path is open in another program (close it in PowerPoint first)" }
        $bak = Join-Path (Split-Path $path) ([System.IO.Path]::GetFileNameWithoutExtension($path) +
            ".bak-" + (Get-Date -Format "yyyyMMdd-HHmmss") + [System.IO.Path]::GetExtension($path))
        Copy-Item -LiteralPath $path $bak
    }
    # Attaches to a running PowerPoint if there is one; Close-Deck leaves it running in that case.
    $app = Get-PowerPointApp
    $wasRunning = -not (Test-OwnPowerPoint)
    $deck = $app.Presentations.Open($path, [bool]$ReadOnly, $false, $false)   # ReadOnly, Untitled, WithWindow
    [pscustomobject]@{ App = $app; Deck = $deck; Path = $path; WasRunning = $wasRunning }
}

# The PowerPoint application. A PowerPoint still shutting down (the previous script just ended) rejects
# calls for a few seconds, so retry for up to 20 s, creating the object again each time.
# PowerPoint is one instance per Windows session, so this attaches to the person's PowerPoint when it runs.
# A POWERPNT process that was not there before is the one this process started: its id is kept, and only
# that instance is ever quit (Close-Deck) or stopped without checks (Stop-PowerPoint).
function Get-PowerPointApp {
    $before = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $deadline = (Get-Date).AddSeconds(20)
    while ($true) {
        try {
            $app = New-Object -ComObject PowerPoint.Application
            [void]$app.Presentations.Count       # the first call is the one a closing PowerPoint rejects
            break
        } catch {
            $busy = $_.Exception.Message -match 'rejected|RPC server is unavailable|0x80010001|0x800706BA|0x800706B5|0x80080005'
            if (-not $busy -or (Get-Date) -gt $deadline) { throw }
            Start-Sleep -Milliseconds 500
        }
    }
    $new = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
    if ($new.Count) { $script:PptxComStartedPid = $new[0].Id }
    $app
}

# True when the running PowerPoint is the one this process started (a PowerPoint sitting on its start
# screen is still the person's: whether a presentation is open says nothing about who started it).
function Test-OwnPowerPoint {
    if (-not $script:PptxComStartedPid) { return $false }
    [bool](Get-Process -Id $script:PptxComStartedPid -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -eq 'POWERPNT' })
}

# Ends a hung PowerPoint without risking anyone's work. Stops it only when it has no visible window and,
# unless this process started it, no presentation with unsaved changes: PowerPoint is one instance, so a
# person who opens it while a script runs joins the script's process. A PowerPoint that does not answer
# within 10 s cannot say what is unsaved, so then only the window decides. Throws, changing nothing, in
# every other case. Never Stop-Process POWERPNT directly: the person's unsaved work lives there.
function Stop-PowerPoint {
    $p = @(Get-Process POWERPNT -ErrorAction SilentlyContinue) | Select-Object -First 1
    if (-not $p) { return }
    $unsaved = @()
    if ($p.Id -ne $script:PptxComStartedPid) {
        # Asked on another thread, so a hung PowerPoint cannot hang this one; strings come back, no COM objects.
        $ps = [powershell]::Create().AddScript({
            $app = New-Object -ComObject PowerPoint.Application
            foreach ($d in @($app.Presentations)) { if ($d.Saved -eq 0) { $d.FullName } }
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
        })
        $job = $ps.BeginInvoke()
        $answered = $job.AsyncWaitHandle.WaitOne(10000)
        $unsaved = if ($answered) { @($ps.EndInvoke($job) | ForEach-Object { [string]$_ }) } else { @() }
        if ($answered) { $ps.Dispose() } else { [void]$ps.BeginStop($null, $null) }   # a hung call cannot be waited for
    }
    if ($unsaved.Count) { throw "PowerPoint was not stopped: unsaved changes in " + ($unsaved -join ", ") }
    if ((Get-Process -Id $p.Id -ErrorAction SilentlyContinue).MainWindowHandle -ne [IntPtr]::Zero) {
        throw "PowerPoint was not stopped: it has a visible window (the person's; close it there)"
    }
    # A windowless PowerPoint may exit by itself once the probe lets go of it.
    try { Stop-Process -Id $p.Id -Force -ErrorAction Stop } catch { if (-not $p.HasExited) { throw } }
    [void]$p.WaitForExit(10000)
    if ($p.Id -eq $script:PptxComStartedPid) { $script:PptxComStartedPid = $null }
}

# Saves as .pptx with the fonts embedded (needs fonts whose licence allows embedding, e.g. OFL), warns for
# each font that cannot be embedded, and records the deck beside a .pptx as <name>.dump.txt (its hash, then
# its dump) so Get-DeckHandEdits can tell later what was changed by hand.
function Save-Deck($session, [string]$path) {
    if (-not $path) { $path = $session.Path }
    else { $path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path) }   # COM resolves relative paths against its own folder
    if (-not $path) { throw "Save-Deck: no path; a deck from New-DeckFromTemplate needs one (Save-Deck `$s out.pptx)" }
    $start = Get-Date
    $session.Deck.SaveAs($path, 24, -1)     # ppSaveAsOpenXMLPresentation, EmbedTrueTypeFonts
    # A hidden dialog (font embedding) can leave SaveAs returning without writing: check the file itself.
    # The second of slack covers file systems that store coarse write times.
    $f = Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
    if (-not $f -or $f.LastWriteTime -lt $start.AddSeconds(-1)) { throw "Save-Deck: $path was not written (a hidden PowerPoint dialog?)" }
    $session.Path = $path         # the deck is that file now: Get-DeckHandEdits reads its record
    foreach ($font in @($session.Deck.Fonts)) {
        if ($font.Embeddable -eq 0) { Write-Warning "font '$($font.Name)' cannot be embedded: viewers without it see a substitute" }
    }
    if ([System.IO.Path]::GetExtension($path) -eq ".pptx") {
        $record = @("sha256 " + (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLower()) + @(Get-DeckDump $session.Deck)
        $record | Set-Content -Encoding UTF8 (Get-DeckRecordPath $path)
    }
}

# The record Save-Deck writes beside a deck.
function Get-DeckRecordPath([string]$path) {
    Join-Path (Split-Path $path) ([System.IO.Path]::GetFileNameWithoutExtension($path) + ".dump.txt")
}

# What changed in the session's file since Save-Deck last recorded it: nothing when the file is unchanged,
# else the dump lines that differ, "- " for the recorded ones and "+ " for the current ones. Throws when
# there is no record beside the deck.
function Get-DeckHandEdits($session) {
    $record = Get-DeckRecordPath $session.Path
    if (-not (Test-Path -LiteralPath $record)) { throw "no record ${record}: save the deck once with Save-Deck first" }
    $lines = @(Get-Content -LiteralPath $record -Encoding UTF8)
    if ($lines[0] -eq ("sha256 " + (Get-FileHash -LiteralPath $session.Path -Algorithm SHA256).Hash.ToLower())) { return }
    $was = @($lines | Select-Object -Skip 1)
    $now = @(Get-DeckDump $session.Deck)
    $diff = @(Compare-Object $was $now -CaseSensitive)
    $diff | Where-Object { $_.SideIndicator -eq '<=' } | ForEach-Object { "- " + $_.InputObject }
    $diff | Where-Object { $_.SideIndicator -eq '=>' } | ForEach-Object { "+ " + $_.InputObject }
}

# Closes the deck and drops the session's COM references: PowerPoint, started windowless, stays alive while
# any reference lives, and the runtime frees the intermediate ones only on a garbage collection.
# Quits PowerPoint only when this process started it, no presentation is left and it has no window: a person
# who opened PowerPoint meanwhile joined this instance and made it visible.
function Close-Deck($session) {
    $session.Deck.Close()
    if (-not $session.WasRunning -and $session.App.Presentations.Count -eq 0 -and $session.App.Visible -eq 0) { $session.App.Quit() }
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($session.Deck)
    # One release, not a final one: every session in this process shares one wrapper around the application.
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($session.App)
    $session.Deck = $null; $session.App = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
}

# Every text-bearing shape on a slide as {Shape; Name}, in z-order, inside groups (nested too) and table
# cells. A cell's shape has no name of its own, so Name says "<table> r,c".
function Get-TextItems($slide) {
    foreach ($sh in @($slide.Shapes)) { Get-ShapeTextItems $sh }
}
function Get-ShapeTextItems($shape) {
    if ($shape.Type -eq 6) {            # msoGroup
        foreach ($child in @($shape.GroupItems)) { Get-ShapeTextItems $child }
    } elseif ($shape.HasTable) {
        $t = $shape.Table
        for ($r = 1; $r -le $t.Rows.Count; $r++) {
            for ($c = 1; $c -le $t.Columns.Count; $c++) {
                [pscustomobject]@{ Shape = $t.Cell($r, $c).Shape; Name = "{0} {1},{2}" -f $shape.Name, $r, $c }
            }
        }
    } elseif ($shape.HasTextFrame) {
        [pscustomobject]@{ Shape = $shape; Name = $shape.Name }
    }
}

# Slides and shapes are found by their text, never by index: hand edits move and renumber slides.
function Get-SlideTexts($slide) {
    @(Get-TextItems $slide) | ForEach-Object { $_.Shape.TextFrame.TextRange.Text }
}
function Find-Slide($deck, [string]$pattern) {
    @($deck.Slides) | Where-Object { (Get-SlideTexts $_) -match $pattern } | Select-Object -First 1
}
function Find-Shape($slide, [string]$pattern) {
    @(Get-TextItems $slide) | Where-Object { $_.Shape.TextFrame.TextRange.Text -match $pattern } |
        Select-Object -First 1 | ForEach-Object { $_.Shape }
}

# Replaces the one occurrence of $old (literal, case-sensitive) in a deck or on one slide, keeping the run's
# formatting (assigning .Text would drop a bold number's bold). Throws when $old is missing (the edit is
# already applied, or the text was changed by hand) or occurs more than once. Returns the shape.
function Edit-DeckText($target, [string]$old, [string]$new) {
    $slides = if ($target.PSObject.Properties['Slides']) { @($target.Slides) } else { @($target) }
    $hits = @(foreach ($s in $slides) {
        foreach ($it in @(Get-TextItems $s)) {
            $n = [regex]::Matches([string]$it.Shape.TextFrame.TextRange.Text, [regex]::Escape($old)).Count
            for ($i = 0; $i -lt $n; $i++) { [pscustomobject]@{ Slide = $s.SlideIndex; Shape = $it.Shape } }
        }
    })
    if ($hits.Count -eq 0) { throw "'$old' not found: already applied, or changed by hand" }
    if ($hits.Count -gt 1) {
        throw ("'$old' occurs {0} times, on slides {1}" -f $hits.Count, ((@($hits | ForEach-Object { $_.Slide }) | Select-Object -Unique) -join ", "))
    }
    $shape = $hits[0].Shape
    [void]$shape.TextFrame.TextRange.Replace($old, $new, 0, -1, 0)     # After, MatchCase, WholeWords
    $shape
}
function Get-Picture($slide) { @($slide.Shapes) | Where-Object { $_.Type -eq $MsoPicture } | Select-Object -First 1 }

function Set-Notes($slide, [string]$text) {
    $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $text
}

# A fixed-size, word-wrapped text box; $color is "#RRGGBB".
function Add-TextBox($slide, $left, $top, $width, $height, [string]$text, [string]$font, $size, [string]$color) {
    $b = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
    $b.TextFrame.AutoSize = 0; $b.Height = [single]$height; $b.TextFrame.WordWrap = -1
    # [single] always: PowerShell caches a COM property set by its argument type, so an int size in one call
    # and a fractional one in a later call fail with "Unable to cast ... to System.Int32".
    $r = $b.TextFrame.TextRange; $r.Text = $text; $r.Font.Name = $font; $r.Font.Size = [single]$size
    $r.Font.Color.RGB = ConvertTo-Bgr $color
    $b
}

# Scales a picture or video to fit inside $frame (any shape) with a margin, centred, aspect kept.
function Set-FitInto($shape, $frame, $margin = 15) {
    $shape.LockAspectRatio = -1
    $k = [math]::Min(($frame.Width - 2 * $margin) / $shape.Width, ($frame.Height - 2 * $margin) / $shape.Height)
    $shape.Width = [single]($shape.Width * $k)
    $shape.Left = [single]($frame.Left + ($frame.Width - $shape.Width) / 2)
    $shape.Top = [single]($frame.Top + ($frame.Height - $shape.Height) / 2)
}

# Replaces the slide's first picture with $file in the same box (give it the same aspect ratio).
function Set-SwapPicture($slide, [string]$file) {
    $old = Get-Picture $slide
    $new = $slide.Shapes.AddPicture((Resolve-Path $file).Path, 0, -1, $old.Left, $old.Top, $old.Width, $old.Height)
    $old.Delete(); $new
}

# A picture cropped to fill a box exactly (photos, full-bleed covers), centred; the crop stays editable.
function Add-PictureFill($slide, [string]$file, $left, $top, $width, $height) {
    $m = $slide.Shapes.AddPicture((Resolve-Path $file).Path, 0, -1, 0, 0)
    # Crops are measured on the picture's own size, and PowerPoint may have shrunk it on insert: reset first.
    $m.ScaleHeight(1, -1); $m.ScaleWidth(1, -1)
    $m.LockAspectRatio = 0
    $w0 = $m.Width; $h0 = $m.Height
    if ($w0 / $h0 -gt $width / $height) {
        $c = [single](($w0 - $h0 * $width / $height) / 2); $m.PictureFormat.CropLeft = $c; $m.PictureFormat.CropRight = $c
    } else {
        $c = [single](($h0 - $w0 * $height / $width) / 2); $m.PictureFormat.CropTop = $c; $m.PictureFormat.CropBottom = $c
    }
    $m.Left = [single]$left; $m.Top = [single]$top; $m.Width = [single]$width; $m.Height = [single]$height
    $m
}

# Embeds a video that plays when the slide opens; $poster is the still shown before it plays.
# Compress first (H.264, CRF 24-28): PowerPoint embeds the file as it is.
function Add-Video($slide, [string]$file, [string]$poster, $left, $top, $width, $height) {
    $v = $slide.Shapes.AddMediaObject2((Resolve-Path $file).Path, $false, $true, $left, $top, $width, $height)
    if ($poster) { $v.MediaFormat.SetDisplayPictureFromFile((Resolve-Path $poster).Path) }
    $v.AnimationSettings.PlaySettings.PlayOnEntry = -1
    $v
}

function Set-Transition($slide, $effect = $PpFade, $duration = 0.7) {
    $t = $slide.SlideShowTransition; $t.EntryEffect = $effect; $t.Duration = $duration
}

# One build-in group: the first shape on $trigger, the others with it (a card and its labels).
function Add-BuildIn($slide, $shapes, $effect = $MsoFadeIn, $trigger = $OnClick, $duration = 0.5) {
    $seq = $slide.TimeLine.MainSequence; $first = $true
    foreach ($sh in @($shapes)) {
        $e = $seq.AddEffect($sh, $effect, 0, $(if ($first) { $trigger } else { $WithPrevious }))
        $e.Timing.Duration = $duration; $first = $false
    }
}

# Applies a manifest from slides_manifest.py (its docstring is the schema) to the PowerPoint export of a
# Claude Slides deck: transitions, build-ins and videos, which the export drops. Slide i of the manifest is
# slide i of the deck; its title, whitespace collapsed and compared literally, must be one of the slide's
# texts (an empty title is not checked). A build finds its shape by text: the one top-level shape holding
# exactly the build's text, else the one holding a text that starts with it. A build without text (an image)
# is skipped with a warning. A video's file is the one in $mediaDir named like the last segment of its src
# (with or without extension); a video without a box fills the slide. Everything is checked first: when
# anything cannot be applied, it throws once listing every item and the deck is not changed. Videos go in
# before anything else, since only inserting tells whether PowerPoint can read a file: when one fails, the
# videos this call added are deleted and it throws. Run it once per export (it adds, never replaces).
function Invoke-FinishManifest($deck, [string]$manifestPath, [string]$mediaDir) {
    $m = [System.IO.File]::ReadAllText((Resolve-Path $manifestPath).Path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    function Squash($t) { ([string]$t -replace '\s+', ' ').Trim() }
    # Windows PowerShell turns an empty JSON list into $null, and @($null) has one item: hence the filter.
    function Items($list) { @($list | Where-Object { $null -ne $_ }) }
    $transitions = @{ fade = $PpFade; push = $PpPushUp; magic = $PpMorph }
    $effects = @{ rise = $MsoRiseUp; fade = $MsoFadeIn; pop = $MsoZoom }
    $media = if ($mediaDir) { @(Get-ChildItem -LiteralPath $mediaDir -File) } else { @() }
    function Find-Media([string]$src) {
        $leaf = ($src -split '[\\/]')[-1]
        if ($leaf) { $media | Where-Object { $_.Name -eq $leaf -or $_.BaseName -eq $leaf } | Select-Object -First 1 }
    }
    $W = $deck.PageSetup.SlideWidth; $H = $deck.PageSetup.SlideHeight
    $k = $W / $m.canvas_width
    $problems = New-Object System.Collections.ArrayList
    $todo = New-Object System.Collections.ArrayList
    $slides = Items $m.slides
    if ($slides.Count -ne $deck.Slides.Count) { [void]$problems.Add("the deck has $($deck.Slides.Count) slides, the manifest $($slides.Count)") }
    foreach ($ms in $slides) {
        $n = [int]$ms.index
        if ($n -lt 1 -or $n -gt $deck.Slides.Count) { [void]$problems.Add("slide $n ($($ms.id)): not in the deck"); continue }
        $slide = $deck.Slides.Item($n)
        $title = Squash $ms.title
        if ($title -and @(Get-SlideTexts $slide | ForEach-Object { Squash $_ }) -cnotcontains $title) {
            [void]$problems.Add("slide ${n}: title '$title' is not on it (slides moved, or the export changed)")
        }
        if ($ms.transition) {
            if ($transitions.ContainsKey([string]$ms.transition)) { [void]$todo.Add(@("transition", $slide, $transitions[[string]$ms.transition])) }
            else { [void]$problems.Add("slide ${n}: unknown transition '$($ms.transition)'") }
        }
        foreach ($b in (Items $ms.builds)) {
            $text = Squash $b.text
            if (-not $effects.ContainsKey([string]$b.effect)) { [void]$problems.Add("slide ${n}: unknown build '$($b.effect)' on '$text'"); continue }
            if (-not $text) { Write-Warning "slide ${n}: a '$($b.effect)' build on an element without text cannot be found by text; add it by hand"; continue }
            # Animations go on top-level shapes: a match inside a group or table animates the group or table.
            $hits = @(foreach ($sh in @($slide.Shapes)) {
                if (@(Get-ShapeTextItems $sh | Where-Object { (Squash $_.Shape.TextFrame.TextRange.Text).StartsWith($text, [System.StringComparison]::Ordinal) }).Count) { $sh }
            })
            # "38%" also starts "38% of trucks": the shape holding exactly the text wins.
            $exact = @(foreach ($sh in $hits) { if (@(Get-ShapeTextItems $sh | Where-Object { (Squash $_.Shape.TextFrame.TextRange.Text) -ceq $text }).Count) { $sh } })
            if ($exact.Count -eq 1) { $hits = $exact }
            if ($hits.Count -eq 1) { [void]$todo.Add(@("build", $slide, $effects[[string]$b.effect], $hits[0])) }
            else { [void]$problems.Add("slide ${n}: $($hits.Count) shapes start with '$text' (one needed for the '$($b.effect)' build)") }
        }
        foreach ($v in (Items $ms.videos)) {
            $file = Find-Media $v.src
            if (-not $mediaDir) { [void]$problems.Add("slide ${n}: video '$($v.src)' needs a media folder"); continue }
            if (-not $file) { [void]$problems.Add("slide ${n}: no file for video '$($v.src)' in $mediaDir"); continue }
            $poster = $null
            if ($v.poster) {
                $poster = Find-Media $v.poster
                if (-not $poster) { Write-Warning "slide ${n}: no poster file for '$($v.poster)' in ${mediaDir}: the video shows its first frame" }
            }
            $box = if ($v.box) { @(($v.box.x * $k), ($v.box.y * $k), ($v.box.w * $k), ($v.box.h * $k)) } else { @(0, 0, $W, $H) }
            [void]$todo.Add(@("video", $slide, $file.FullName, $(if ($poster) { $poster.FullName } else { "" }), $box))
        }
    }
    if ($problems.Count) { throw ("Invoke-FinishManifest: nothing applied, because:`n  " + ($problems -join "`n  ")) }
    $added = New-Object System.Collections.ArrayList
    foreach ($t in $todo) {
        if ($t[0] -ne "video") { continue }
        $sl = $t[1]; $b = $t[4]; $n0 = $sl.Shapes.Count
        try { [void]$added.Add((Add-Video $sl $t[2] $t[3] ([single]$b[0]) ([single]$b[1]) ([single]$b[2]) ([single]$b[3]))) }
        catch {
            # A poster that fails leaves the new video behind: it is above every shape that was there.
            while ($sl.Shapes.Count -gt $n0) { $sl.Shapes.Item($sl.Shapes.Count).Delete() }
            foreach ($v in $added) { $v.Delete() }
            throw ("Invoke-FinishManifest: nothing applied, because:`n  slide $($sl.SlideIndex): PowerPoint cannot insert '$($t[2])': " + $_.Exception.Message)
        }
    }
    foreach ($t in $todo) {
        switch ($t[0]) {
            "transition" { Set-Transition $t[1] $t[2] }
            "build" { Add-BuildIn $t[1] $t[3] $t[2] }
        }
    }
}

# --- templates ------------------------------------------------------------------------------------
# Template paths live in ~/.claude/decks.json, outside any repository, so company files are never committed:
#   { "templates": { "acme": { "path": "C:\\...\\Template Acme.pptx", "library": "C:\\...\\Acme library.pptx", "use": "client and marketing decks" } } }
# -Library returns the slide library registered beside the template instead (see Slide library).
function Get-DeckTemplatePath([string]$name, [switch]$Library) {
    $cfg = Join-Path $env:USERPROFILE ".claude\decks.json"
    if (-not (Test-Path $cfg)) { throw "no $cfg; register the template there first" }
    $t = (Get-Content -Raw $cfg | ConvertFrom-Json).templates.$name
    if (-not $t) { throw "template '$name' is not in $cfg" }
    if ($Library) {
        if (-not $t.library) { throw "template '$name' has no library in $cfg" }
        return $t.library
    }
    $t.path
}

function ConvertFrom-Bgr([int]$bgr) { "#{0:X2}{1:X2}{2:X2}" -f ($bgr -band 0xFF), (($bgr -shr 8) -band 0xFF), (($bgr -shr 16) -band 0xFF) }

# Layouts with their placeholders, theme colours and fonts, as markdown: read it to pick layouts.
function Get-TemplateProfile($deck) {
    $roles = "dark 1", "light 1", "dark 2", "light 2", "accent 1", "accent 2", "accent 3", "accent 4", "accent 5", "accent 6"
    $m = $deck.Designs.Item(1).SlideMaster.Theme
    "# Template profile: " + $deck.Name
    ""
    "Slide size: {0} x {1} pt" -f $deck.PageSetup.SlideWidth, $deck.PageSetup.SlideHeight
    "Fonts: headings {0}, body {1}" -f $m.ThemeFontScheme.MajorFont.Item(1).Name, $m.ThemeFontScheme.MinorFont.Item(1).Name
    "Theme colours: " + (@(for ($i = 1; $i -le 10; $i++) { "{0} {1}" -f $roles[$i - 1], (ConvertFrom-Bgr $m.ThemeColorScheme.Colors($i).RGB) }) -join ", ")
    ""
    "| Design | Layout | Placeholders (type: left,top width x height) |"
    "|---|---|---|"
    foreach ($d in @($deck.Designs)) {
        foreach ($l in @($d.SlideMaster.CustomLayouts)) {
            $ph = @($l.Shapes.Placeholders) | ForEach-Object {
                "{0}: {1:N0},{2:N0} {3:N0}x{4:N0}" -f (Get-PlaceholderKind $_.PlaceholderFormat.Type), $_.Left, $_.Top, $_.Width, $_.Height
            }
            "| {0} | {1} | {2} |" -f $d.Name, $l.Name, ($ph -join "; ")
        }
    }
}

# ppPlaceholderType -> a short name used by Get-Placeholder.
function Get-PlaceholderKind([int]$type) {
    switch ($type) { 1 { "title" } 3 { "title" } 2 { "body" } 7 { "body" } 4 { "subtitle" } 18 { "picture" }
        8 { "chart" } 12 { "table" } 13 { "date" } 15 { "footer" } 16 { "number" } default { "type$type" } }
}

# A new, unsaved deck from a .pptx or .potx template (the file is never written), without its sample slides;
# -KeepSlides keeps them, for a slide library.
function New-DeckFromTemplate([string]$template, [switch]$KeepSlides) {
    $app = Get-PowerPointApp
    $wasRunning = -not (Test-OwnPowerPoint)
    $deck = $app.Presentations.Open((Resolve-Path $template).Path, $false, $true, $false)   # Untitled: a copy
    if (-not $KeepSlides) { while ($deck.Slides.Count -gt 0) { $deck.Slides.Item(1).Delete() } }
    [pscustomobject]@{ App = $app; Deck = $deck; Path = $null; WasRunning = $wasRunning }
}

# --- slide library --------------------------------------------------------------------------------
# A library deck holds designed pattern slides, each tagged with a role in its COM Name ("role:<name>"):
# invisible to the audience, and it survives hand edits and moves. A new deck is a copy of the library
# that keeps the patterns (New-DeckFromTemplate -KeepSlides); Add-FromLibrary duplicates one and replaces its text, Remove-LibrarySlides
# drops the unused patterns at the end.

# Tags a slide with a role (case-study, cover, big-number, ...).
function Set-SlideRole($slide, [string]$role) { $slide.Name = "role:" + $role }

# The slide's role, or $null when it is untagged.
function Get-SlideRole($slide) {
    if ($slide.Name -match '^role:(.+)$') { $Matches[1] } else { $null }
}

# One {Role; Slide (index); Title; Shapes} per tagged slide, to see which text shapes a script must replace.
function Get-SlideLibrary($deck) {
    foreach ($s in @($deck.Slides)) {
        $role = Get-SlideRole $s
        if (-not $role) { continue }
        $texts = @(@(Get-TextItems $s) | Where-Object { $_.Shape.TextFrame.HasText })
        $shapes = @($texts | ForEach-Object {
            $t = $_.Shape.TextFrame.TextRange.Text -replace "[\r\n\v]+", " / "
            "{0}: {1}" -f $_.Name, $t.Substring(0, [math]::Min(40, $t.Length))
        })
        $title = if ($texts.Count -gt 0) { $texts[0].Shape.TextFrame.TextRange.Text } else { "" }
        [pscustomobject]@{ Role = $role; Slide = $s.SlideIndex; Title = $title; Shapes = $shapes }
    }
}

# Duplicates the first slide tagged $role, puts the copy just before the first pattern slide, untags it, and
# sets the text of the shape matching each key regex (Find-Shape) in $texts. Every key is matched against the
# pattern's text before any is written, so one new text never decides where another goes. Throws, deleting
# the copy, when a key matches nothing or two keys match the same shape. Returns the new slide.
function Add-FromLibrary($deck, [string]$role, [hashtable]$texts) {
    $src = @(@($deck.Slides) | Where-Object { (Get-SlideRole $_) -eq $role }) | Select-Object -First 1
    if (-not $src) { throw "no slide tagged role '$role'" }
    $copy = $src.Duplicate().Item(1)
    # A fresh unique name: the copy may carry the pattern's role tag, and names must be unique.
    $copy.Name = "built-" + [guid]::NewGuid().ToString("N").Substring(0, 8)
    $first = @(@($deck.Slides) | Where-Object { Get-SlideRole $_ }) | Select-Object -First 1
    $copy.MoveTo($first.SlideIndex)
    # Keys resolve to positions in the walker's list: COM gives no reliable identity to compare shapes by.
    $items = @(Get-TextItems $copy)
    $at = @{}
    foreach ($key in $texts.Keys) {
        $i = -1
        for ($j = 0; $j -lt $items.Count; $j++) { if ($items[$j].Shape.TextFrame.TextRange.Text -match $key) { $i = $j; break } }
        if ($i -lt 0) { $copy.Delete(); throw "no shape on the '$role' slide matches '$key'" }
        if ($at.ContainsKey($i)) { $copy.Delete(); throw "'$key' and '$($at[$i])' match the same shape on the '$role' slide" }
        $at[$i] = $key
    }
    foreach ($i in $at.Keys) {
        $items[$i].Shape.TextFrame.TextRange.Text = ([string]$texts[$at[$i]] -replace "`r?`n", "`r")
    }
    $copy
}

# Deletes every slide still tagged with a role (the unused patterns).
function Remove-LibrarySlides($deck) {
    $tagged = @(@($deck.Slides) | Where-Object { Get-SlideRole $_ })
    foreach ($s in $tagged) { $s.Delete() }
}

# A starter slide library for a template that has none (one from New-TemplateFromTheme): designed pattern
# slides tagged cover, section, cards, big-numbers, quote, timeline and closing, saved to $out. Colours are
# theme colours and fonts are the theme's heading and body fonts, so the library follows the template.
# Geometry is drawn for 960 x 540 and scaled to the slide size. Every sample text starts with
# "[placeholder]", so one left unreplaced is a Test-DeckLeftovers finding.
function New-StarterLibrary([string]$template, [string]$out) {
    $s = New-DeckFromTemplate $template
    try {
        $d = $s.Deck
        $kx = $d.PageSetup.SlideWidth / 960; $ky = $d.PageSetup.SlideHeight / 540
        $layouts = @($d.Designs.Item(1).SlideMaster.CustomLayouts)
        $layout = @($layouts | Where-Object { $_.Name -eq 'Blank' }) + @($layouts | Sort-Object { $_.Shapes.Placeholders.Count }) | Select-Object -First 1
        $dk1 = 1; $lt1 = 2; $dk2 = 3; $lt2 = 4; $acc1 = 5                  # msoThemeColorIndex
        function Pattern($role, $bg) {
            $sl = $d.Slides.AddSlide($d.Slides.Count + 1, $layout)
            foreach ($p in @($sl.Shapes.Placeholders)) { $p.Delete() }
            $sl.FollowMasterBackground = 0; $sl.Background.Fill.Solid(); $sl.Background.Fill.ForeColor.ObjectThemeColor = $bg
            Set-SlideRole $sl $role
            $sl
        }
        # A circle or square scales by the smaller factor, about its centre, so it stays round on a 4:3 slide.
        # -Outline draws a thin line: a card's fill (light 2) may equal the background (light 1).
        function Shape($sl, $type, $x, $y, $w, $h, $color, [switch]$Outline) {
            $sw = $w * $kx; $sh = $h * $ky
            if ($w -eq $h) { $sw = $sh = $w * [math]::Min($kx, $ky) }
            $r = $sl.Shapes.AddShape($type, [single](($x + $w / 2) * $kx - $sw / 2), [single](($y + $h / 2) * $ky - $sh / 2), [single]$sw, [single]$sh)
            $r.Fill.ForeColor.ObjectThemeColor = $color; $r.Line.Visible = 0
            if ($Outline) { $r.Line.Visible = -1; $r.Line.ForeColor.ObjectThemeColor = $dk2; $r.Line.Weight = [single]0.75 }
        }
        # "+mj-lt" and "+mn-lt" are the theme's heading and body fonts. -Bottom anchors the text to the box's
        # foot, so a title of one line or three ends at the same height, just above what follows it.
        function Text($sl, $x, $y, $w, $h, $text, $size, $color, [switch]$Heading, [switch]$Bold, [switch]$Raw, [switch]$Bottom) {
            if (-not $Raw) { $text = "[placeholder] " + $text }
            $b = Add-TextBox $sl ($x * $kx) ($y * $ky) ($w * $kx) ($h * $ky) $text $(if ($Heading) { "+mj-lt" } else { "+mn-lt" }) ($size * $kx) "#000000"
            $b.TextFrame.TextRange.Font.Color.ObjectThemeColor = $color
            if ($Bold) { $b.TextFrame.TextRange.Font.Bold = -1 }
            if ($Bottom) { $b.TextFrame.VerticalAnchor = 4 }             # msoAnchorBottom
        }
        $names = "one", "two", "three", "four"
        $col = (832 - 2 * 24) / 3                                       # three columns between 64 pt margins

        $c = Pattern "cover" $dk1
        Text $c 64 120 832 230 "Deck title in a few words" 52 $lt1 -Heading -Bold -Bottom
        Shape $c 1 64 364 72 6 $acc1
        Text $c 64 384 832 40 "Subtitle - who, where, when" 22 $lt2

        $c = Pattern "section" $dk2
        Text $c 64 150 832 160 "Section title" 44 $lt1 -Heading -Bold -Bottom
        Shape $c 1 64 324 72 6 $acc1
        Text $c 64 344 832 36 "Part 1 - what this part answers" 20 $lt2

        $c = Pattern "cards" $lt1
        Text $c 64 44 832 90 "Slide title that makes the point" 32 $dk1 -Heading -Bold
        for ($i = 0; $i -lt 3; $i++) {
            $x = 64 + $i * ($col + 24)
            Shape $c 1 $x 160 $col 280 $lt2 -Outline
            Shape $c 1 $x 160 $col 6 $acc1
            Text $c ($x + 20) 186 ($col - 40) 66 ("Card " + $names[$i]) 24 $dk1 -Heading -Bold
            Text $c ($x + 20) 256 ($col - 40) 170 ("Two or three lines on what card " + $names[$i] + " says.") 18 $dk2
        }

        $c = Pattern "big-numbers" $dk1
        Text $c 64 44 832 90 "Slide title that makes the point" 32 $lt1 -Heading -Bold
        $nums = "38%", "120", "4"
        for ($i = 0; $i -lt 3; $i++) {
            $x = 64 + $i * ($col + 24)
            # Tall enough for the sample: "[placeholder]" alone takes two lines at this size.
            Text $c $x 134 $col 258 $nums[$i] 72 $lt1 -Heading -Bold -Bottom
            Shape $c 1 $x 404 48 6 $acc1
            Text $c $x 420 ($col - 16) 80 ("what number " + $names[$i] + " counts") 20 $lt2
        }

        $c = Pattern "quote" $lt2
        Text $c 80 60 120 150 ([string][char]0x201C) 120 $acc1 -Heading -Bold -Raw   # an opening quote mark
        Text $c 96 190 768 170 "A sentence a customer or expert said, in their own words." 34 $dk1 -Heading
        Text $c 96 380 768 40 "Name, role, organisation" 18 $dk2

        $c = Pattern "timeline" $lt1
        Text $c 64 44 832 90 "Plan in four steps" 32 $dk1 -Heading -Bold
        $ln = $c.Shapes.AddLine([single]((64 + 16) * $kx), [single](250 * $ky), [single]((64 + 3 * 208 + 16) * $kx), [single](250 * $ky))
        $ln.Line.ForeColor.ObjectThemeColor = $acc1; $ln.Line.Weight = [single](3 * $kx)
        for ($i = 0; $i -lt 4; $i++) {
            $x = 64 + $i * 208
            Shape $c 9 $x 234 32 32 $acc1                                 # msoShapeOval
            Text $c $x 262 192 64 ("Step " + $names[$i]) 22 $dk1 -Heading -Bold -Bottom
            Text $c $x 330 184 130 ("What happens in step " + $names[$i] + ".") 18 $dk2
        }

        $c = Pattern "closing" $dk1
        Text $c 64 120 832 230 "The ask, or the one thing to remember" 48 $lt1 -Heading -Bold -Bottom
        Shape $c 1 64 364 72 6 $acc1
        Text $c 64 384 832 40 "Next step - name, e-mail" 22 $lt2

        Save-Deck $s $out
    } finally { Close-Deck $s }
}

# Appends a slide on the first layout (any design) whose name matches $pattern.
function Add-LayoutSlide($deck, [string]$pattern) {
    foreach ($d in @($deck.Designs)) {
        $l = @($d.SlideMaster.CustomLayouts) | Where-Object { $_.Name -match $pattern } | Select-Object -First 1
        if ($l) { return $deck.Slides.AddSlide($deck.Slides.Count + 1, $l) }
    }
    throw "no layout matches '$pattern'"
}

# The $n-th placeholder of a kind (title, body, subtitle, picture, ...), counted left to right, top to bottom.
function Get-Placeholder($slide, [string]$kind, [int]$n = 1) {
    # @(...) around the whole pipeline: one match would otherwise come back as a bare COM object.
    $all = @(@($slide.Shapes.Placeholders) | Where-Object { (Get-PlaceholderKind $_.PlaceholderFormat.Type) -eq $kind } |
        Sort-Object { [math]::Round($_.Top / 20) }, Left)
    if ($all.Count -lt $n) { throw "slide $($slide.SlideIndex) has no $kind placeholder number $n" }
    $all[$n - 1]
}

# Text into a placeholder keeps the template's styling; lines split on "`n" become paragraphs (bullets in a body).
function Set-PlaceholderText($slide, [string]$kind, [string]$text, [int]$n = 1) {
    (Get-Placeholder $slide $kind $n).TextFrame.TextRange.Text = ($text -replace "`r?`n", "`r")
}

# Removes the placeholders still empty, so they do not show "Click to add text" in the editor.
function Remove-EmptyPlaceholders($slide) {
    foreach ($p in @($slide.Shapes.Placeholders)) {
        if ($p.HasTextFrame -and -not $p.TextFrame.HasText) { $p.Delete() }
    }
}

# A .potx from a look chosen elsewhere (a Claude Slides deck, another AI tool, a brand page): theme colours,
# heading and body fonts, and the standard layouts restyled. Colours are "#RRGGBB"; 2 to 6 accents.
function New-TemplateFromTheme([string]$outPath, [string]$dark, [string]$light, [string[]]$accents,
                               [string]$headingFont, [string]$bodyFont, [string]$dark2, [string]$light2) {
    if (-not $dark2) { $dark2 = $dark }; if (-not $light2) { $light2 = $light }
    $app = Get-PowerPointApp
    $wasRunning = -not (Test-OwnPowerPoint)
    $p = $app.Presentations.Add($false)
    try {
        $p.PageSetup.SlideWidth = 960; $p.PageSetup.SlideHeight = 540           # 16:9
        $master = $p.Designs.Item(1).SlideMaster
        $c = $master.Theme.ThemeColorScheme
        $c.Colors(1).RGB = ConvertTo-Bgr $dark; $c.Colors(2).RGB = ConvertTo-Bgr $light
        $c.Colors(3).RGB = ConvertTo-Bgr $dark2; $c.Colors(4).RGB = ConvertTo-Bgr $light2
        for ($i = 0; $i -lt 6; $i++) { $c.Colors(5 + $i).RGB = ConvertTo-Bgr $accents[[math]::Min($i, $accents.Count - 1)] }
        $c.Colors(11).RGB = ConvertTo-Bgr $accents[0]
        $f = $master.Theme.ThemeFontScheme
        $f.MajorFont.Item(1).Name = $headingFont; $f.MinorFont.Item(1).Name = $bodyFont
        # Office's standard layouts in their fixed order, renamed in English (Office names them in its own
        # language); the two vertical-text layouts go. The section layout gets the dark background.
        $names = "Title", "Title and content", "Section", "Two content", "Comparison", "Title only", "Blank",
                 "Content with caption", "Picture with caption"
        while ($master.CustomLayouts.Count -gt $names.Count) { $master.CustomLayouts.Item($master.CustomLayouts.Count).Delete() }
        for ($i = 1; $i -le $names.Count; $i++) { $master.CustomLayouts.Item($i).Name = $names[$i - 1] }
        $sec = $master.CustomLayouts.Item(3)
        $sec.FollowMasterBackground = $false; $sec.Background.Fill.Solid(); $sec.Background.Fill.ForeColor.RGB = ConvertTo-Bgr $dark
        foreach ($ph in @($sec.Shapes.Placeholders)) { if ($ph.HasTextFrame) { $ph.TextFrame.TextRange.Font.Color.RGB = ConvertTo-Bgr $light } }
        $p.SaveAs($outPath, 26)        # ppSaveAsOpenXMLTemplate (.potx)
    } finally { Close-Deck ([pscustomobject]@{ App = $app; Deck = $p; WasRunning = $wasRunning }) }
}

# One PNG per slide, to look at the deck the way it will be projected. Old slide*.png in $dir go first (a
# shorter deck would leave stale renders); numbers are zero-padded to the slide count's digits, at least 2.
function Export-SlidePngs($deck, [string]$dir, $width = 1600) {
    $dir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($dir)
    New-Item -ItemType Directory -Force $dir | Out-Null
    Get-ChildItem -LiteralPath $dir -Filter "slide*.png" | Remove-Item -Force
    $h = [int]($width * $deck.PageSetup.SlideHeight / $deck.PageSetup.SlideWidth)
    $digits = [math]::Max(2, ([string]$deck.Slides.Count).Length)
    foreach ($s in @($deck.Slides)) {
        $s.Export((Join-Path $dir ("slide" + ([string]$s.SlideIndex).PadLeft($digits, '0') + ".png")), "PNG", $width, $h)
    }
}

# A plain-text dump of every slide (shapes, boxes, first font and size, text, notes, effects), as lines;
# text inside groups and table cells gets an indented line of its own. Save-Deck records it, and
# Get-DeckHandEdits diffs it to see what a person changed by hand.
function Get-DeckDump($deck) {
    function Format-Text($shape) {
        $r = $shape.TextFrame.TextRange
        if (-not $r.Text) { return "" }
        $f = $r.Runs(1).Font
        "{0} {1}  {2}" -f $f.Name, $f.Size, ($r.Text -replace "[\r\n\v]+", " / ")
    }
    foreach ($s in @($deck.Slides)) {
        "=== slide {0}  layout '{1}'  transition {2}  effects {3}" -f $s.SlideIndex, $s.CustomLayout.Name,
            $s.SlideShowTransition.EntryEffect, $s.TimeLine.MainSequence.Count
        foreach ($sh in @($s.Shapes)) {
            $inside = $sh.Type -eq 6 -or $sh.HasTable
            $t = if (-not $inside -and $sh.HasTextFrame) { Format-Text $sh } else { "" }
            # Rounded by hand: "{0:N0}" rounds a COM single differently in Windows PowerShell 5.1 and PowerShell 7
            # (112.49999 gives 113 or 112), and a record from one shell must compare equal in the other.
            $box = @($sh.Left, $sh.Top, $sh.Width, $sh.Height) | ForEach-Object { [long][math]::Round([double]$_, [MidpointRounding]::AwayFromZero) }
            "  [{0}] type {1}  {2},{3} {4}x{5}  {6}" -f $sh.Name, $sh.Type, $box[0], $box[1], $box[2], $box[3], $t
            if ($inside) { foreach ($it in @(Get-ShapeTextItems $sh)) { "    [{0}]  {1}" -f $it.Name, (Format-Text $it.Shape) } }
        }
        $n = $s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
        if ($n) { "  notes: " + ($n -replace "[\r\n\v]+", " / ") }
    }
}
function Write-DeckDump($deck, [string]$outFile) {
    Get-DeckDump $deck | Set-Content -Encoding UTF8 $outFile
}

# --- checks and previews --------------------------------------------------------------------------

# Text left over from drafting (placeholders, build notes, file paths, variant labels), as a list of
# {Slide; Where; Match; Text}. Scans shapes, groups, table cells and notes; prints nothing itself.
# $extraPatterns add to the defaults; a hit whose text matches an $exclude regex is dropped. A slide with
# review comments gives one finding (Where "comments", Match their count). The default patterns are shared
# with pptx_check.py and tests/leftover_cases.txt: change them together.
function Test-DeckLeftovers($deck, [string[]]$extraPatterns, [string[]]$exclude) {
    $euro = [string][char]0x20AC
    $patterns = @('lorem|ipsum', '\bTODO\b|\bTBD\b|\bFIXME\b',
        ('\[insert|\[placeholder|\[' + $euro + '__\]|\[[A-Z ]*__\]'), 'click to add', 'made with claude',
        '\b(option|variant) [A-D]\b', '[A-Za-z]:\x5C|\.(py|ps1|csv|json|md)\b') + @($extraPatterns | Where-Object { $_ })
    $exclude = @($exclude | Where-Object { $_ })
    # One list entry per text: (slide index, location, text).
    $pending = New-Object System.Collections.ArrayList
    foreach ($s in @($deck.Slides)) {
        foreach ($it in @(Get-TextItems $s)) { [void]$pending.Add(@($s.SlideIndex, ("shape " + $it.Name), $it.Shape.TextFrame.TextRange.Text)) }
        foreach ($sh in @($s.NotesPage.Shapes)) {
            if ($sh.Type -eq 14 -and $sh.HasTextFrame -and $sh.PlaceholderFormat.Type -eq 2) { [void]$pending.Add(@($s.SlideIndex, "notes", $sh.TextFrame.TextRange.Text)) }
        }
    }
    foreach ($p in $pending) {
        $text = [string]$p[2]
        if (@($exclude | Where-Object { [regex]::IsMatch($text, $_, 'IgnoreCase') }).Count) { continue }
        foreach ($re in $patterns) {
            $m = [regex]::Match($text, $re, 'IgnoreCase')
            if ($m.Success) {
                $t = $text -replace "[\r\n\v]+", " / "
                [pscustomobject]@{ Slide = $p[0]; Where = $p[1]; Match = $m.Value; Text = $t.Substring(0, [math]::Min(80, $t.Length)) }
            }
        }
    }
    foreach ($s in @($deck.Slides)) {
        $n = $s.Comments.Count
        if ($n -gt 0) {
            $t = [string]$s.Comments.Item(1).Text -replace "[\r\n\v]+", " / "
            [pscustomobject]@{ Slide = $s.SlideIndex; Where = "comments"; Match = $n; Text = $t.Substring(0, [math]::Min(80, $t.Length)) }
        }
    }
}

# Layout faults, as a list of {Slide; Where; Check; Detail}; prints nothing. $minPt is for a 960 pt wide
# slide and scales with the slide's width. Checks: small-text (a run under the minimum; date, footer and
# slide-number placeholders are skipped), overflow (a frame that does not grow to fit its text, shrink on
# overflow included, whose text is over 2 pt taller than the shape; table rows grow, so cells are skipped),
# off-slide (a top-level unrotated shape leaving the slide by over 1 pt; pictures are skipped, since a photo
# may bleed off the edge on purpose), font (a deck font that cannot be embedded: Slide 0, Where "deck").
function Test-DeckLayout($deck, $minPt = 12) {
    $W = $deck.PageSetup.SlideWidth; $H = $deck.PageSetup.SlideHeight
    $min = $minPt * $W / 960
    foreach ($s in @($deck.Slides)) {
        foreach ($sh in @($s.Shapes)) {
            if ($sh.Type -ne $MsoPicture -and $sh.Rotation -eq 0 -and ($sh.Left -lt -1 -or $sh.Top -lt -1 -or $sh.Left + $sh.Width -gt $W + 1 -or $sh.Top + $sh.Height -gt $H + 1)) {
                [pscustomobject]@{ Slide = $s.SlideIndex; Where = "shape " + $sh.Name; Check = "off-slide"
                    Detail = "{0:N0},{1:N0} {2:N0}x{3:N0} on a {4:N0}x{5:N0} slide" -f $sh.Left, $sh.Top, $sh.Width, $sh.Height, $W, $H }
            }
            if ($sh.Type -eq 14 -and @("date", "footer", "number") -contains (Get-PlaceholderKind $sh.PlaceholderFormat.Type)) { continue }
            foreach ($it in @(Get-ShapeTextItems $sh)) {
                $r = $it.Shape.TextFrame.TextRange
                if (-not $r.Text.Trim()) { continue }
                # COM gives each run's effective size, inherited ones included.
                $small = @(for ($i = 1; $i -le $r.Runs().Count; $i++) {
                    $run = $r.Runs($i); if ($run.Text.Trim() -and $run.Font.Size -lt $min) { $run.Font.Size }
                })
                $t = $r.Text -replace "[\r\n\v]+", " / "; $t = $t.Substring(0, [math]::Min(40, $t.Length))
                if ($small.Count) {
                    [pscustomobject]@{ Slide = $s.SlideIndex; Where = "shape " + $it.Name; Check = "small-text"
                        Detail = "{0} pt under {1:0.#} pt: {2}" -f ($small | Measure-Object -Minimum).Minimum, $min, $t }
                }
                # BoundHeight is measured on the slide, so a rotated shape's text looks taller than it is: skipped.
                if (-not $sh.HasTable -and $it.Shape.Rotation -eq 0 -and $it.Shape.TextFrame.AutoSize -ne 1 -and $r.BoundHeight -gt $it.Shape.Height + 2) {
                    [pscustomobject]@{ Slide = $s.SlideIndex; Where = "shape " + $it.Name; Check = "overflow"
                        Detail = "text {0:N0} pt tall in a {1:N0} pt box: {2}" -f $r.BoundHeight, $it.Shape.Height, $t }
                }
            }
        }
    }
    foreach ($font in @($deck.Fonts)) {
        if ($font.Embeddable -eq 0) { [pscustomobject]@{ Slide = 0; Where = "deck"; Check = "font"; Detail = "'$($font.Name)' cannot be embedded" } }
    }
}

# One PNG sheet with a labelled thumbnail ("design / layout") of every custom layout, each empty slot
# showing its kind name. Uses temporary slides on $deck and deletes them, so a -ReadOnly deck works.
function Export-LayoutSheet($deck, [string]$outPng, $columns = 3, $thumbWidth = 480) {
    $outPng = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($outPng)   # absolute, so a bare file name has a parent folder
    $w = [int]$thumbWidth; $h = [int]($w * $deck.PageSetup.SlideHeight / $deck.PageSetup.SlideWidth)
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("layoutsheet-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force $tmp | Out-Null
    $before = $deck.Slides.Count
    $items = @()
    try {
        foreach ($d in @($deck.Designs)) {
            foreach ($l in @($d.SlideMaster.CustomLayouts)) {
                $slide = $deck.Slides.AddSlide($deck.Slides.Count + 1, $l)
                try {
                    foreach ($p in @($slide.Shapes.Placeholders)) {
                        if ($p.HasTextFrame -and -not $p.TextFrame.HasText) {
                            $p.TextFrame.TextRange.Text = Get-PlaceholderKind $p.PlaceholderFormat.Type
                        }
                    }
                    $file = Join-Path $tmp ("t{0:D3}.png" -f $items.Count)
                    $slide.Export($file, "PNG", $w, $h)
                    $items += [pscustomobject]@{ File = $file; Label = $d.Name + " / " + $l.Name }
                } finally { $slide.Delete() }
            }
        }
        Write-TileSheet @($items | ForEach-Object { $_.File }) $outPng $columns $w $h 12 @($items | ForEach-Object { $_.Label })
    } finally {
        while ($deck.Slides.Count -gt $before) { $deck.Slides.Item($deck.Slides.Count).Delete() }
        Remove-Item -Recurse -Force $tmp
    }
}

# Every slide as a thumbnail on one PNG sheet, labelled with its number and "(hidden)" for a hidden slide:
# the whole deck at a glance.
function Export-ContactSheet($deck, [string]$outPng, $columns = 4, $thumbWidth = 480) {
    $outPng = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($outPng)
    $w = [int]$thumbWidth; $h = [int]($w * $deck.PageSetup.SlideHeight / $deck.PageSetup.SlideWidth)
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("contactsheet-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force $tmp | Out-Null
    try {
        $files = @(); $labels = @()
        foreach ($s in @($deck.Slides)) {
            $files += Join-Path $tmp ("s{0:D4}.png" -f $s.SlideIndex)
            $s.Export($files[-1], "PNG", $w, $h)
            $labels += [string]$s.SlideIndex + $(if ($s.SlideShowTransition.Hidden -eq -1) { " (hidden)" })
        }
        Write-TileSheet $files $outPng $columns $w $h 12 $labels
    } finally { Remove-Item -Recurse -Force $tmp }
}

# Internal: draws PNG files as $w x $h tiles, $columns across, $gap apart, on a $background ("#RRGGBB")
# sheet. With $labels, each tile gets a thin frame and its label underneath. Shared by Export-LayoutSheet,
# Export-ContactSheet and the example scripts (which pass no labels).
function Write-TileSheet([string[]]$files, [string]$outPng, $columns, $w, $h, $gap = 12, [string[]]$labels, [string]$background = "#FFFFFF") {
    Add-Type -AssemblyName System.Drawing
    $labelH = if ($labels) { 28 } else { 0 }
    $rows = [math]::Ceiling($files.Count / $columns)
    $bmp = New-Object System.Drawing.Bitmap ([int]($columns * ($w + $gap) + $gap)), ([int]($rows * ($h + $labelH + $gap) + $gap))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $font = New-Object System.Drawing.Font("Segoe UI", 11)
    try {
        $g.Clear([System.Drawing.ColorTranslator]::FromHtml($background))
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        for ($i = 0; $i -lt $files.Count; $i++) {
            $x = $gap + ($i % $columns) * ($w + $gap); $y = $gap + [math]::Floor($i / $columns) * ($h + $labelH + $gap)
            $img = [System.Drawing.Image]::FromFile($files[$i])
            try { $g.DrawImage($img, [int]$x, [int]$y, [int]$w, [int]$h) } finally { $img.Dispose() }
            if ($labels) {
                $g.DrawRectangle([System.Drawing.Pens]::LightGray, [int]$x, [int]$y, [int]$w, [int]$h)
                $g.DrawString($labels[$i], $font, [System.Drawing.Brushes]::Black, [single]$x, [single]($y + $h + 4))
            }
        }
        New-Item -ItemType Directory -Force (Split-Path -Parent $outPng) | Out-Null
        $bmp.Save($outPng, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally { $font.Dispose(); $g.Dispose(); $bmp.Dispose() }
}

# The deck as a PDF: one page per slide, or with -Notes each slide with its speaker notes (the version to
# read alone, from the same source). Hidden slides are left out. The deck itself is not changed.
function Export-DeckPdf($deck, [string]$outPdf, [switch]$Notes) {
    $outPdf = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($outPdf)
    if ($Notes) {
        Add-Type -AssemblyName System.Drawing
        if (@([System.Drawing.Printing.PrinterSettings]::InstalledPrinters) -notcontains $PdfPrinter) {
            throw "Export-DeckPdf -Notes needs the '$PdfPrinter' printer, which is not installed (Windows Settings > Printers & scanners)"
        }
    }
    # Works on a throwaway copy: PowerPoint writes only the first PDF saved from an open deck (later saves
    # return without writing), a PDF save ignores PrintOptions.OutputType, and ExportAsFixedFormat fails
    # under late binding. Notes pages therefore go through the "Microsoft Print to PDF" printer.
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("deckpdf-" + [guid]::NewGuid().ToString("N") + ".pptx")
    $deck.SaveCopyAs($tmp)
    $start = Get-Date
    $app = $deck.Application
    $copy = $null
    try {
        $copy = $app.Presentations.Open($tmp, -1, 0, 0)                 # ReadOnly, Untitled no, no window
        if ($Notes) {
            $po = $copy.PrintOptions
            $po.OutputType = 5; $po.PrintInBackground = 0; $po.ActivePrinter = $PdfPrinter   # ppPrintOutputNotesPages
            $po.PrintHiddenSlides = 0; $po.PrintComments = 0       # as the slides PDF: no hidden slides, no comment pages
            $copy.PrintOut(1, $copy.Slides.Count, $outPdf, 1, 0)       # From, To, PrintToFile, Copies, Collate
        } else {
            $copy.SaveAs($outPdf, 32)                                    # ppSaveAsPDF
        }
    } finally {
        if ($copy) { Close-Deck ([pscustomobject]@{ App = $app; Deck = $copy; WasRunning = $true }) }
        Remove-Item -LiteralPath $tmp -Force
    }
    $f = Get-Item -LiteralPath $outPdf -ErrorAction SilentlyContinue
    if (-not $f -or $f.LastWriteTime -lt $start.AddSeconds(-1)) { throw "Export-DeckPdf: $outPdf was not written" }
}
