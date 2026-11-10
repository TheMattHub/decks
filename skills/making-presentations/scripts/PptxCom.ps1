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
            ".bak-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".pptx")
        Copy-Item -LiteralPath $path $bak
    }
    # Attaches to a running PowerPoint if there is one; Close-Deck leaves it running in that case.
    $app = Get-PowerPointApp
    $wasRunning = $app.Presentations.Count -gt 0
    $deck = $app.Presentations.Open($path, [bool]$ReadOnly, $false, $false)   # ReadOnly, Untitled, WithWindow
    [pscustomobject]@{ App = $app; Deck = $deck; Path = $path; WasRunning = $wasRunning }
}

# The PowerPoint application. A PowerPoint still shutting down (the previous script just ended) rejects
# calls for a few seconds, so retry for up to 20 s, creating the object again each time.
function Get-PowerPointApp {
    $deadline = (Get-Date).AddSeconds(20)
    while ($true) {
        try {
            $app = New-Object -ComObject PowerPoint.Application
            [void]$app.Presentations.Count       # the first call is the one a closing PowerPoint rejects
            return $app
        } catch {
            $busy = $_.Exception.Message -match 'rejected|RPC server is unavailable|0x80010001|0x800706BA|0x80080005'
            if (-not $busy -or (Get-Date) -gt $deadline) { throw }
            Start-Sleep -Milliseconds 500
        }
    }
}

# Saves as .pptx with the fonts embedded (needs fonts whose licence allows embedding, e.g. OFL).
function Save-Deck($session, [string]$path) {
    if (-not $path) { $path = $session.Path }
    else { $path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path) }   # COM resolves relative paths against its own folder
    $session.Deck.SaveAs($path, 24, -1)     # ppSaveAsOpenXMLPresentation, EmbedTrueTypeFonts
}

# Closes the deck and drops the session's COM references: PowerPoint, started windowless, stays alive while
# any reference lives, and the runtime frees the intermediate ones only on a garbage collection.
function Close-Deck($session) {
    $session.Deck.Close()
    if (-not $session.WasRunning -and $session.App.Presentations.Count -eq 0) { $session.App.Quit() }
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($session.Deck)
    # One release, not a final one: every session in this process shares one wrapper around the application.
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($session.App)
    $session.Deck = $null; $session.App = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
}

# Every text-bearing shape on a slide as {Shape; Name}, in z-order, inside groups (nested too) and table
# cells. A cell's shape has no name of its own, so Name says "<table> r,c".
function Get-TextItems($slide) {
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
    foreach ($sh in @($slide.Shapes)) { Get-ShapeTextItems $sh }
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
    $r = $b.TextFrame.TextRange; $r.Text = $text; $r.Font.Name = $font; $r.Font.Size = $size
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
    $wasRunning = $app.Presentations.Count -gt 0
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
    $wasRunning = $app.Presentations.Count -gt 0
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

# A plain-text dump of every slide (shapes, boxes, text, notes, effects). Diff two dumps to see what a
# person changed by hand before editing a deck that has hand edits.
function Write-DeckDump($deck, [string]$outFile) {
    $lines = foreach ($s in @($deck.Slides)) {
        "=== slide {0}  layout '{1}'  transition {2}  effects {3}" -f $s.SlideIndex, $s.CustomLayout.Name,
            $s.SlideShowTransition.EntryEffect, $s.TimeLine.MainSequence.Count
        foreach ($sh in @($s.Shapes)) {
            $t = if ($sh.HasTextFrame) { ($sh.TextFrame.TextRange.Text -replace "[\r\n\v]+", " / ") } else { "" }
            "  [{0}] type {1}  {2:N0},{3:N0} {4:N0}x{5:N0}  {6}" -f $sh.Name, $sh.Type, $sh.Left, $sh.Top, $sh.Width, $sh.Height, $t
        }
        $n = $s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
        if ($n) { "  notes: " + ($n -replace "[\r\n\v]+", " / ") }
    }
    $lines | Set-Content -Encoding UTF8 $outFile
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

# One PNG sheet with a labelled thumbnail ("design / layout") of every custom layout, each empty slot
# showing its kind name. Uses temporary slides on $deck and deletes them, so a -ReadOnly deck works.
function Export-LayoutSheet($deck, [string]$outPng, $columns = 3, $thumbWidth = 480) {
    $outPng = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($outPng)   # absolute, so a bare file name has a parent folder
    Add-Type -AssemblyName System.Drawing
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
    } finally {
        while ($deck.Slides.Count -gt $before) { $deck.Slides.Item($deck.Slides.Count).Delete() }
    }
    $labelH = 28; $gap = 12
    $rows = [math]::Ceiling($items.Count / $columns)
    $bmp = New-Object System.Drawing.Bitmap ([int]($columns * ($w + $gap) + $gap)), ([int]($rows * ($h + $labelH + $gap) + $gap))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.Clear([System.Drawing.Color]::White)
        $font = New-Object System.Drawing.Font("Segoe UI", 11)
        for ($i = 0; $i -lt $items.Count; $i++) {
            $x = $gap + ($i % $columns) * ($w + $gap); $y = $gap + [math]::Floor($i / $columns) * ($h + $labelH + $gap)
            $img = [System.Drawing.Image]::FromFile($items[$i].File)
            try { $g.DrawImage($img, [int]$x, [int]$y, $w, $h) } finally { $img.Dispose() }
            $g.DrawRectangle([System.Drawing.Pens]::LightGray, [int]$x, [int]$y, $w, $h)
            $g.DrawString($items[$i].Label, $font, [System.Drawing.Brushes]::Black, [single]$x, [single]($y + $h + 4))
        }
        New-Item -ItemType Directory -Force (Split-Path -Parent $outPng) | Out-Null
        $bmp.Save($outPng, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally { $g.Dispose(); $bmp.Dispose(); Remove-Item -Recurse -Force $tmp }
}
