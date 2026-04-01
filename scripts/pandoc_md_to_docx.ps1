[CmdletBinding()]
param(
  [Parameter(Mandatory = $false, Position = 0)]
  [string]$InputPath,

  [Parameter(Mandatory = $false, Position = 1)]
  [string]$OutputPath,

  [Alias("h")]
  [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TargetScript = Join-Path $ScriptDir "..\skills\markdown-to-docx\scripts\pandoc_md_to_docx.ps1"
$ResolvedTarget = [System.IO.Path]::GetFullPath($TargetScript)

if (-not (Test-Path -LiteralPath $ResolvedTarget)) {
  throw "Compatibility wrapper target not found: $ResolvedTarget"
}

$ForwardArgs = @()

if ($Help.IsPresent) {
  $ForwardArgs += "-h"
}

if ($PSBoundParameters.ContainsKey("InputPath")) {
  $ForwardArgs += @("-InputPath", $InputPath)
}

if ($PSBoundParameters.ContainsKey("OutputPath")) {
  $ForwardArgs += @("-OutputPath", $OutputPath)
}

& $ResolvedTarget @ForwardArgs
