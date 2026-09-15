$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path $PSScriptRoot 'create_excel_mvp.ps1'
$scriptText = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
$escapedRoot = $PSScriptRoot.Replace("'", "''")
$scriptText = $scriptText.Replace('$PSScriptRoot', "'$escapedRoot'")
Invoke-Expression $scriptText
