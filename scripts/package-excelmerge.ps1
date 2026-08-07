[CmdletBinding()]
param(
    [ValidatePattern('^[A-Za-z0-9._-]+$')]
    [string]$RuntimeIdentifier = "win-x64",

    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release",

    [string]$OutputDirectory = "artifacts",

    [switch]$SkipBinaryCheck
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Remove-ExactPath {
    param([Parameter(Mandatory)][string]$Path)

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
}

function Copy-PublishFiles {
    param(
        [Parameter(Mandatory)][string]$PublishPath,
        [Parameter(Mandatory)][string]$DestinationPath,
        [Parameter(Mandatory)][string]$DisplayName
    )

    $publishFiles = Get-ChildItem -LiteralPath $PublishPath -File -Recurse |
        Where-Object { $_.Extension -ine ".pdb" }
    if (-not $publishFiles) {
        throw "No runtime files were found in the published $DisplayName output."
    }

    foreach ($publishFile in $publishFiles) {
        $relativePath = [IO.Path]::GetRelativePath($PublishPath, $publishFile.FullName)
        $destinationFile = Join-Path $DestinationPath $relativePath
        if (Test-Path -LiteralPath $destinationFile -PathType Leaf) {
            throw "CLI and Desktop publish outputs contain the same file name: $relativePath"
        }
        New-Item -ItemType Directory -Path (Split-Path -Parent $destinationFile) -Force | Out-Null
        Copy-Item -LiteralPath $publishFile.FullName -Destination $destinationFile -Force
    }
}

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$cliProject = Join-Path $repoRoot "src/ExcelMerge.Cli/ExcelMerge.Cli.csproj"
$desktopProject = Join-Path $repoRoot "src/ExcelMerge.Desktop/ExcelMerge.Desktop.csproj"
$packageName = "ExcelMerge-$RuntimeIdentifier"
$outputRoot = if ([IO.Path]::IsPathRooted($OutputDirectory)) {
    [IO.Path]::GetFullPath($OutputDirectory)
} else {
    [IO.Path]::GetFullPath((Join-Path $repoRoot $OutputDirectory))
}

if ([string]::IsNullOrWhiteSpace($OutputDirectory) -or
    $outputRoot -eq [IO.Path]::GetPathRoot($outputRoot)) {
    throw "OutputDirectory must identify a directory below a filesystem root."
}

$packageRoot = Join-Path $outputRoot $packageName
$archivePath = Join-Path $outputRoot "$packageName.zip"
$cliPublishRoot = Join-Path $outputRoot ".publish-cli-$RuntimeIdentifier"
$desktopPublishRoot = Join-Path $outputRoot ".publish-desktop-$RuntimeIdentifier"
$cliName = if ($RuntimeIdentifier.StartsWith("win-", [StringComparison]::OrdinalIgnoreCase)) {
    "ExcelMerge.Cli.exe"
} else {
    "ExcelMerge.Cli"
}
$desktopName = if ($RuntimeIdentifier.StartsWith("win-", [StringComparison]::OrdinalIgnoreCase)) {
    "ExcelMerge.Desktop.exe"
} else {
    "ExcelMerge.Desktop"
}
$publishedCli = Join-Path $cliPublishRoot $cliName
$publishedDesktop = Join-Path $desktopPublishRoot $desktopName

$requiredFiles = @(
    $cliProject,
    $desktopProject,
    (Join-Path $repoRoot "README.md"),
    (Join-Path $repoRoot "docs/AI-INSTALL.md"),
    (Join-Path $repoRoot "docs/ai-git-merge-driver.md"),
    (Join-Path $repoRoot "docs/user-guide.zh-CN.md")
)

foreach ($requiredFile in $requiredFiles) {
    if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
        throw "Required package input was not found: $requiredFile"
    }
}

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw "The .NET SDK is required but dotnet was not found on PATH."
}

New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
Remove-ExactPath $packageRoot
Remove-ExactPath $cliPublishRoot
Remove-ExactPath $desktopPublishRoot
Remove-ExactPath $archivePath

try {
    function Publish-NativeAot {
        param(
            [Parameter(Mandatory)][string]$ProjectPath,
            [Parameter(Mandatory)][string]$PublishPath,
            [Parameter(Mandatory)][string]$DisplayName
        )

        $publishArguments = @(
            "publish",
            $ProjectPath,
            "--configuration", $Configuration,
            "--runtime", $RuntimeIdentifier,
            "--self-contained", "true",
            "--output", $PublishPath,
            "-p:PublishAot=true",
            "-p:PublishSingleFile=true",
            "-p:DebugType=None"
        )

        Write-Host "Publishing AOT $DisplayName for $RuntimeIdentifier..."
        & dotnet @publishArguments
        if ($LASTEXITCODE -ne 0) {
            throw "dotnet publish failed for $DisplayName with exit code $LASTEXITCODE."
        }
    }

    Publish-NativeAot -ProjectPath $cliProject -PublishPath $cliPublishRoot -DisplayName "CLI"
    Publish-NativeAot -ProjectPath $desktopProject -PublishPath $desktopPublishRoot -DisplayName "Desktop"

    if (-not (Test-Path -LiteralPath $publishedCli -PathType Leaf)) {
        throw "The published CLI was not found at $publishedCli."
    }
    if (-not (Test-Path -LiteralPath $publishedDesktop -PathType Leaf)) {
        throw "The published Desktop app was not found at $publishedDesktop."
    }

    if (-not $SkipBinaryCheck) {
    Write-Host "Checking the published CLI..."
    & $publishedCli --version
        if ($LASTEXITCODE -ne 0) {
            throw "The published CLI did not pass the --version check."
        }
    }

    $binRoot = Join-Path $packageRoot "bin"
    $docsRoot = Join-Path $packageRoot "docs"
    New-Item -ItemType Directory -Path $binRoot, $docsRoot -Force | Out-Null

    Copy-PublishFiles -PublishPath $cliPublishRoot -DestinationPath $binRoot -DisplayName "CLI"
    Copy-PublishFiles -PublishPath $desktopPublishRoot -DestinationPath $binRoot -DisplayName "Desktop"
    Copy-Item -LiteralPath (Join-Path $repoRoot "README.md") -Destination $packageRoot
    Copy-Item -LiteralPath (Join-Path $repoRoot "docs/AI-INSTALL.md") -Destination $docsRoot
    Copy-Item -LiteralPath (Join-Path $repoRoot "docs/ai-git-merge-driver.md") -Destination $docsRoot
    Copy-Item -LiteralPath (Join-Path $repoRoot "docs/user-guide.zh-CN.md") -Destination $docsRoot

    $packagedCli = Join-Path $binRoot $cliName
    $packagedDesktop = Join-Path $binRoot $desktopName
    $cliHash = (Get-FileHash -LiteralPath $packagedCli -Algorithm SHA256).Hash
    $desktopHash = (Get-FileHash -LiteralPath $packagedDesktop -Algorithm SHA256).Hash
    @(
        "ExcelMerge package"
        "Project source: https://github.com/ksgfk/ExcelMerge"
        "Runtime identifier: $RuntimeIdentifier"
        "Configuration: $Configuration"
        "Generated UTC: $([DateTime]::UtcNow.ToString("O"))"
        "CLI: bin/$cliName"
        "CLI SHA256: $cliHash"
        "Desktop: bin/$desktopName"
        "Desktop SHA256: $desktopHash"
        "AI installation guide: docs/AI-INSTALL.md"
        "AI Git guide: docs/ai-git-merge-driver.md"
        "User guide: docs/user-guide.zh-CN.md"
    ) | Set-Content -LiteralPath (Join-Path $packageRoot "PACKAGE-MANIFEST.txt") -Encoding UTF8

    Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $archivePath -CompressionLevel Optimal

    Write-Host ""
    Write-Host "Package directory: $packageRoot"
    Write-Host "Package archive:   $archivePath"
}
catch {
    Remove-ExactPath $packageRoot
    Remove-ExactPath $archivePath
    throw
}
finally {
    Remove-ExactPath $cliPublishRoot
    Remove-ExactPath $desktopPublishRoot
}
