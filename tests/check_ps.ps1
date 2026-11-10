# Parses every .ps1 under skills/, examples/ and tests/ and checks it is ASCII-only (Windows PowerShell 5.1
# reads BOM-less scripts as ANSI), and that PptxCom.ps1's $PptxComVersion equals the plugin's version.
$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$failed = 0
foreach ($f in Get-ChildItem (Join-Path $root "skills"), (Join-Path $root "examples"), (Join-Path $root "tests") -Recurse -Filter *.ps1) {
    $tokens = $null; $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errors)
    foreach ($e in $errors) { Write-Error "$($f.Name):$($e.Extent.StartLineNumber): $($e.Message)" -ErrorAction Continue; $failed++ }
    $bad = @([System.IO.File]::ReadAllBytes($f.FullName) | Where-Object { $_ -gt 127 }).Count
    if ($bad) { Write-Error "$($f.Name): $bad non-ASCII bytes" -ErrorAction Continue; $failed++ }
    "$($f.Name): parses, ASCII-only"
}
$lib = Get-Content -Raw (Join-Path $root "skills/making-presentations/scripts/PptxCom.ps1")
$ps = [regex]::Match($lib, '(?m)^\$PptxComVersion\s*=\s*"([^"]+)"').Groups[1].Value
$plugin = (Get-Content -Raw (Join-Path $root ".claude-plugin/plugin.json") | ConvertFrom-Json).version
if ($ps -ne $plugin) { Write-Error "`$PptxComVersion '$ps' differs from plugin.json's version '$plugin'" -ErrorAction Continue; $failed++ }
else { "`$PptxComVersion $ps = plugin.json" }
if ($failed) { exit 1 }
