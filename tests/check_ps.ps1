# Parses PptxCom.ps1 and checks it is ASCII-only (Windows PowerShell 5.1 reads BOM-less scripts as ANSI).
$path = Join-Path $PSScriptRoot "..\skills\making-presentations\scripts\PptxCom.ps1"
$tokens = $null; $errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $path), [ref]$tokens, [ref]$errors)
if ($errors.Count) { $errors | ForEach-Object { Write-Error $_.Message }; exit 1 }
$bad = @([System.IO.File]::ReadAllBytes((Resolve-Path $path)) | Where-Object { $_ -gt 127 }).Count
if ($bad) { Write-Error "$bad non-ASCII bytes in PptxCom.ps1"; exit 1 }
"PptxCom.ps1: parses, ASCII-only"
