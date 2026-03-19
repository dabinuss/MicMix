param(
    [string]$Version = "",
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot

function Assert-SafeVersion {
    param([Parameter(Mandatory = $true)][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "Version must not be empty."
    }
    if ($Value.Length -gt 64) {
        throw "Version is too long (max 64 characters)."
    }
    if ($Value -notmatch '^[0-9A-Za-z][0-9A-Za-z._-]*$') {
        throw "Invalid version '$Value'. Allowed characters: A-Z a-z 0-9 . _ -"
    }
}

function Resolve-PathInside {
    param(
        [Parameter(Mandatory = $true)][string]$BasePath,
        [Parameter(Mandatory = $true)][string]$CandidatePath,
        [Parameter(Mandatory = $true)][string]$Label
    )

    $baseFull = [System.IO.Path]::GetFullPath($BasePath)
    $candidateFull = [System.IO.Path]::GetFullPath($CandidatePath)
    $sep = [string][System.IO.Path]::DirectorySeparatorChar
    $altSep = [string][System.IO.Path]::AltDirectorySeparatorChar
    $basePrefix = $baseFull
    if (-not $basePrefix.EndsWith($sep) -and -not $basePrefix.EndsWith($altSep)) {
        $basePrefix += $sep
    }

    if (-not $candidateFull.StartsWith($basePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label resolved outside base directory. Base: $baseFull Candidate: $candidateFull"
    }

    return $candidateFull
}

if ([string]::IsNullOrWhiteSpace($Version)) {
    $versionFile = Join-Path $repoRoot "VERSION"
    if (-not (Test-Path $versionFile)) {
        throw "VERSION file not found: $versionFile"
    }
    $Version = (Get-Content -LiteralPath $versionFile -Raw).Trim()
}
Assert-SafeVersion -Value $Version

$buildDll = Join-Path $repoRoot "build\$Configuration\micmix.dll"
if (-not (Test-Path $buildDll)) {
    throw "Build artifact missing: $buildDll. Build first (cmake --build build --config $Configuration)."
}
$buildHostExe = Join-Path $repoRoot "build\$Configuration\micmix_vst_host.exe"
if (-not (Test-Path $buildHostExe)) {
    throw "Build artifact missing: $buildHostExe. Build first (cmake --build build --config $Configuration)."
}

$icon16 = Join-Path $repoRoot "assets\branding\MicMixIcon16.png"
if (-not (Test-Path $icon16)) {
    throw "Missing icon: $icon16"
}

$distDir = Join-Path $repoRoot "dist"
$stageRoot = Join-Path $repoRoot "build\package_stage"
$stageDir = Resolve-PathInside -BasePath $stageRoot -CandidatePath (Join-Path $stageRoot "MicMix-$Version") -Label "Stage directory"
$pluginDir = Resolve-PathInside -BasePath $stageDir -CandidatePath (Join-Path $stageDir "plugins") -Label "Plugin directory"
$pluginIconDir = Resolve-PathInside -BasePath $pluginDir -CandidatePath (Join-Path $pluginDir "micmix_win64") -Label "Plugin icon directory"

Remove-Item -LiteralPath $stageDir -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $pluginIconDir -Force | Out-Null
New-Item -ItemType Directory -Path $distDir -Force | Out-Null

Copy-Item -LiteralPath $buildDll -Destination (Join-Path $pluginDir "micmix_win64.dll") -Force
Copy-Item -LiteralPath $buildHostExe -Destination (Join-Path $pluginDir "micmix_vst_host.exe") -Force
Copy-Item -LiteralPath $icon16 -Destination (Join-Path $pluginIconDir "1.png") -Force
Copy-Item -LiteralPath $icon16 -Destination (Join-Path $pluginIconDir "t.png") -Force
Copy-Item -LiteralPath $icon16 -Destination (Join-Path $pluginIconDir "menu_icon.png") -Force

$packageIni = @"
Name = MicMix
Type = Plugin
Author = Dabinuss
Version = $Version
Platforms = win64
Description = Mixes an additional audio source into captured microphone audio.
Api = 26
"@
Set-Content -LiteralPath (Join-Path $stageDir "package.ini") -Value $packageIni -Encoding ASCII

$zipPath = Resolve-PathInside -BasePath $distDir -CandidatePath (Join-Path $distDir "MicMix-$Version-win64.zip") -Label "Zip artifact"
$ts3PluginPath = Resolve-PathInside -BasePath $distDir -CandidatePath (Join-Path $distDir "MicMix-$Version-win64.ts3_plugin") -Label "Plugin artifact"
Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $ts3PluginPath -Force -ErrorAction SilentlyContinue

Push-Location $stageDir
try {
    Compress-Archive -Path * -DestinationPath $zipPath -CompressionLevel Optimal
} finally {
    Pop-Location
}
Move-Item -LiteralPath $zipPath -Destination $ts3PluginPath -Force

Write-Host "Created:" $ts3PluginPath
Write-Host "SHA256:" ((Get-FileHash -LiteralPath $ts3PluginPath -Algorithm SHA256).Hash)
