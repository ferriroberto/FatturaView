<#
Build script for creating a release build and preparing files for the installer.
Usage:
  Open a Developer PowerShell for Visual Studio and run:
    .\build_release.ps1

This script will:
 - find the solution file in the repository root (or .vcxproj files if no .sln)
 - build the solution or projects in Release x64 (if available) using msbuild
 - locate the produced executable(s) and copy them together with required resources
   into the `installer\dist` folder so the Inno Setup script can package them.

Requires: msbuild in PATH (Visual Studio developer tools) or run from Developer PowerShell.
#>

param(
    [string]$Configuration = "Release",
    [string]$Platform = "x64"
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Search upward up to 6 levels for a solution or project file (handles running from installer folder)
$repoRoot = $scriptDir
$foundRoot = $null
$sln = $null
$proj = $null
for ($i = 0; $i -lt 6; $i++) {
    # look for .sln and .vcxproj in this folder (non-recursive to be deterministic)
    $slnCandidate = Get-ChildItem -Path $repoRoot -Filter "*.sln" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($slnCandidate) { $sln = $slnCandidate; $foundRoot = $repoRoot; break }

    $projCandidate = Get-ChildItem -Path $repoRoot -Filter "*.vcxproj" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($projCandidate) { $proj = $projCandidate; $foundRoot = $repoRoot; break }

    $parent = Split-Path $repoRoot -Parent
    if ($parent -eq $repoRoot -or [string]::IsNullOrEmpty($parent)) { break }
    $repoRoot = $parent
}

if (-not $foundRoot) {
    Write-Error "No solution (.sln) or project (.vcxproj) files found in the repository or parent folders. Run this script from the repository or provide a .sln/.vcxproj.";
    exit 1
}

Set-Location $foundRoot
Write-Host "Preparing build in: $foundRoot"

# If we found a project but not a solution, leave $sln null and later build projects
if ($sln) { $sln = Get-ChildItem -Path $foundRoot -Filter "*.sln" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 }

function Find-MSBuild {
    # Prefer msbuild in PATH
    $ms = Get-Command msbuild -ErrorAction SilentlyContinue
    if ($ms) { return $ms.Path }

    # Try vswhere to locate MSBuild
    $vswhere = "C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $msPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find **\MSBuild\**\Bin\MSBuild.exe 2>$null | Select-Object -First 1
        if ($msPath -and (Test-Path $msPath)) { return $msPath }
    }

    # Common installation locations (check a few typical VS years)
    $candidates = @(
        'C:\Program Files (x86)\Microsoft Visual Studio\\**\\MSBuild\\Current\\Bin\\MSBuild.exe',
        'C:\Program Files\\Microsoft Visual Studio\\**\\MSBuild\\Current\\Bin\\MSBuild.exe'
    )
    foreach ($c in $candidates) {
        $found = Get-ChildItem -Path $c -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) { return $found.FullName }
    }

    return $null
}

$msbuildExe = Find-MSBuild
if (-not $msbuildExe) {
    Write-Error "MSBuild not found. Open Developer PowerShell for Visual Studio or ensure MSBuild is installed and in PATH."
    exit 1
}

if ($sln) {
    Write-Host "Building solution: $($sln.FullName) [$Configuration|$Platform] using MSBuild: $msbuildExe"
    $msbuildArgs = "`"$($sln.FullName)`" /t:Build /p:Configuration=$Configuration;Platform=$Platform"
    $proc = Start-Process -FilePath $msbuildExe -ArgumentList $msbuildArgs -NoNewWindow -Wait -PassThru -ErrorAction SilentlyContinue
    if ($proc.ExitCode -ne 0) {
        Write-Error "MSBuild failed with exit code $($proc.ExitCode)."
        exit $proc.ExitCode
    }
}
else {
    # Build each vcxproj found at repository root (recursive)
    $projects = Get-ChildItem -Path $foundRoot -Filter "*.vcxproj" -Recurse -ErrorAction SilentlyContinue
    if (-not $projects -or $projects.Count -eq 0) {
        Write-Error "No solution (.sln) or project (.vcxproj) files found in the repository root."
        exit 1
    }

    foreach ($projItem in $projects) {
        Write-Host "Building project: $($projItem.FullName) [$Configuration|$Platform] using MSBuild: $msbuildExe"
        $msbuildArgs = "`"$($projItem.FullName)`" /t:Build /p:Configuration=$Configuration;Platform=$Platform"
        $proc = Start-Process -FilePath $msbuildExe -ArgumentList $msbuildArgs -NoNewWindow -Wait -PassThru -ErrorAction SilentlyContinue
        if ($proc.ExitCode -ne 0) {
            Write-Error "MSBuild failed for project $($projItem.Name) with exit code $($proc.ExitCode)."
            exit $proc.ExitCode
        }
    }
}

# Prepare dist folder
$distDir = Join-Path (Get-Location) "installer\dist"
if (Test-Path $distDir) { Remove-Item -Recurse -Force $distDir }
New-Item -ItemType Directory -Path $distDir | Out-Null

# Copy LICENSE
if (Test-Path (Join-Path (Get-Location) "LICENSE")) {
    Copy-Item -Path (Join-Path (Get-Location) "LICENSE") -Destination $distDir
}

# Find built executables matching project name or output name
Write-Host "Searching for built exes..."
$exeFiles = Get-ChildItem -Path (Get-Location) -Include *.exe -Recurse -ErrorAction SilentlyContinue | Where-Object {
    $_.FullName -notmatch "\\packages\\" -and $_.FullName -notmatch "\\.git\\" -and $_.Length -gt 1024
}

if (-not $exeFiles -or $exeFiles.Count -eq 0)
{
    Write-Warning "No executable found under repository after build. You may need to build the project manually in Visual Studio first.";
    exit 0
}

# Heuristic: prefer exe whose path contains Configuration and Platform
$preferred = $exeFiles | Where-Object { $_.FullName -match "\\$Configuration" -and $_.FullName -match "\\$Platform" }
if ($preferred -and $preferred.Count -gt 0) { $exeFiles = $preferred }

# Copy top executables (limit to those in project folder) to dist
$copied = 0
foreach ($exe in $exeFiles)
{
    # Avoid toolchain exes
    $dest = Join-Path $distDir $exe.Name
    Copy-Item -Path $exe.FullName -Destination $dest -Force
    Write-Host "Copied $($exe.FullName) -> $dest"
    $copied++
    if ($copied -ge 3) { break }
}

Write-Host "Dist folder prepared: $distDir"
Write-Host "Now you can run Inno Setup Compiler (ISCC) on installer\setup.iss to create the installer."
Write-Host "Example: ISCC.exe installer\setup.iss"

exit 0
