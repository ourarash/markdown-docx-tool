[CmdletBinding()]
param(
  [Parameter(Mandatory = $false, Position = 0)]
  [string]$InputPath,

  [Parameter(Mandatory = $false, Position = 1)]
  [string]$OutputPath,

  [string]$ReferenceDoc,

  [string]$MetadataFile,

  [string]$OutputDir,

  [switch]$TableOfContents
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ScriptDir
$DefaultReferenceDoc = Join-Path $ScriptDir "reference.docx"

function Show-Usage {
  @"
Usage:
  .\scripts\pandoc_md_to_docx.ps1 -InputPath <input.md> [-OutputPath <output.docx>]
  .\scripts\pandoc_md_to_docx.ps1 -InputPath <input.md> -OutputDir <dir> [-TableOfContents]
  .\scripts\pandoc_md_to_docx.ps1 -InputPath <input.md> -ReferenceDoc <reference.docx> [-MetadataFile <metadata.yaml>]

Examples:
  .\scripts\pandoc_md_to_docx.ps1 -InputPath .\samples\showcase.md
  .\scripts\pandoc_md_to_docx.ps1 -InputPath .\README.md -OutputPath .\output\repo-readme.docx
  .\scripts\pandoc_md_to_docx.ps1 -InputPath .\samples\showcase.md -OutputDir .\output -TableOfContents
  .\scripts\pandoc_md_to_docx.ps1 -InputPath C:\docs\report.md
"@ | Write-Host
}

function Require-Command {
  param([string]$CommandName)

  if (-not (Get-Command $CommandName -ErrorAction SilentlyContinue)) {
    throw "Required command not found: $CommandName"
  }
}

function Resolve-ExistingPath {
  param([string]$Candidate)

  if ([System.IO.Path]::IsPathRooted($Candidate)) {
    if (-not (Test-Path -LiteralPath $Candidate)) {
      throw "Markdown file not found: $Candidate"
    }

    return (Resolve-Path -LiteralPath $Candidate).Path
  }

  if (Test-Path -LiteralPath $Candidate) {
    return (Resolve-Path -LiteralPath $Candidate).Path
  }

  $RepoCandidate = Join-Path $RepoRoot $Candidate
  if (Test-Path -LiteralPath $RepoCandidate) {
    return (Resolve-Path -LiteralPath $RepoCandidate).Path
  }

  throw "Markdown file not found: $Candidate"
}

function Resolve-OutputPath {
  param([string]$Candidate)

  if ([System.IO.Path]::IsPathRooted($Candidate)) {
    return [System.IO.Path]::GetFullPath($Candidate)
  }

  return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Candidate))
}

function Resolve-OptionalExistingPath {
  param([string]$Candidate)

  if ([string]::IsNullOrWhiteSpace($Candidate)) {
    return $null
  }

  return Resolve-ExistingPath $Candidate
}

function Get-TemporaryDirectory {
  $Path = [System.IO.Path]::Combine(
    [System.IO.Path]::GetTempPath(),
    [System.Guid]::NewGuid().ToString()
  )

  [System.IO.Directory]::CreateDirectory($Path) | Out-Null
  return $Path
}

function AutoFit-DocxTables {
  param([string]$DocxPath)

  $TempDir = Get-TemporaryDirectory
  $SourceZipPath = Join-Path ([System.IO.Path]::GetTempPath()) ("{0}-source.zip" -f [System.Guid]::NewGuid())
  $OutputZipPath = Join-Path ([System.IO.Path]::GetTempPath()) ("{0}-output.zip" -f [System.Guid]::NewGuid())

  try {
    # PowerShell's Expand-Archive only accepts .zip paths, even though .docx
    # files are ZIP containers internally.
    Copy-Item -LiteralPath $DocxPath -Destination $SourceZipPath -Force
    Expand-Archive -LiteralPath $SourceZipPath -DestinationPath $TempDir -Force

    $DocumentXml = Join-Path $TempDir "word/document.xml"
    $Content = [System.IO.File]::ReadAllText($DocumentXml)
    $Content = $Content.Replace('<w:tblLayout w:type="fixed"/>', '')
    $Content = $Content.Replace('<w:tblW w:type="pct" w:w="5000"/>', '<w:tblW w:type="auto" w:w="0"/>')
    [System.IO.File]::WriteAllText($DocumentXml, $Content)

    if (Test-Path -LiteralPath $OutputZipPath) {
      Remove-Item -LiteralPath $OutputZipPath -Force
    }

    Compress-Archive -Path (Join-Path $TempDir '*') -DestinationPath $OutputZipPath -Force
    Copy-Item -LiteralPath $OutputZipPath -Destination $DocxPath -Force
  }
  finally {
    if (Test-Path -LiteralPath $TempDir) {
      Remove-Item -LiteralPath $TempDir -Recurse -Force
    }

    if (Test-Path -LiteralPath $SourceZipPath) {
      Remove-Item -LiteralPath $SourceZipPath -Force
    }

    if (Test-Path -LiteralPath $OutputZipPath) {
      Remove-Item -LiteralPath $OutputZipPath -Force
    }
  }
}

if ([string]::IsNullOrWhiteSpace($InputPath)) {
  Show-Usage
  exit 1
}

if ($InputPath -in @('-h', '--help', '/?')) {
  Show-Usage
  exit 0
}

Require-Command "pandoc"

if ($PSBoundParameters.ContainsKey('OutputPath') -and $PSBoundParameters.ContainsKey('OutputDir')) {
  throw "Use either -OutputPath or -OutputDir, not both."
}

$ResolvedReferenceDoc = if ($PSBoundParameters.ContainsKey('ReferenceDoc')) {
  Resolve-ExistingPath $ReferenceDoc
}
else {
  $DefaultReferenceDoc
}

if (-not (Test-Path -LiteralPath $ResolvedReferenceDoc)) {
  throw "Reference document not found: $ResolvedReferenceDoc"
}

$ResolvedInput = Resolve-ExistingPath $InputPath

if (-not (Test-Path -LiteralPath $ResolvedInput -PathType Leaf)) {
  throw "Input is not a file: $ResolvedInput"
}

if ($PSBoundParameters.ContainsKey('OutputPath')) {
  $ResolvedOutput = Resolve-OutputPath $OutputPath
}
elseif ($PSBoundParameters.ContainsKey('OutputDir')) {
  $ResolvedOutputDir = Resolve-OutputPath $OutputDir
  $ResolvedOutput = Join-Path $ResolvedOutputDir (
    "{0}.docx" -f [System.IO.Path]::GetFileNameWithoutExtension($ResolvedInput)
  )
}
else {
  $ResolvedOutput = [System.IO.Path]::ChangeExtension($ResolvedInput, ".docx")
}

$ResolvedMetadataFile = Resolve-OptionalExistingPath $MetadataFile

$InputDir = Split-Path -Parent $ResolvedInput
$OutputDir = Split-Path -Parent $ResolvedOutput
$OutputBase = [System.IO.Path]::GetFileNameWithoutExtension($ResolvedOutput)
$MediaDir = Join-Path $OutputDir ("{0}_media" -f $OutputBase)

if (-not (Test-Path -LiteralPath $OutputDir)) {
  New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$ResourcePath = "{0};{1}" -f $InputDir, $RepoRoot

$PandocArgs = @(
  $ResolvedInput,
  "--from=markdown+smart",
  "--to=docx",
  "--wrap=none",
  "--resource-path=$ResourcePath",
  "--extract-media=$MediaDir",
  "--dpi=300",
  "--reference-doc=$ResolvedReferenceDoc"
)

if ($TableOfContents.IsPresent) {
  $PandocArgs += "--toc"
}

if ($ResolvedMetadataFile) {
  $PandocArgs += "--metadata-file=$ResolvedMetadataFile"
}

$PandocArgs += @("-o", $ResolvedOutput)

Push-Location $RepoRoot
try {
  & pandoc @PandocArgs
}
finally {
  Pop-Location
}

AutoFit-DocxTables $ResolvedOutput
Write-Host ("Wrote: {0}" -f $ResolvedOutput)
