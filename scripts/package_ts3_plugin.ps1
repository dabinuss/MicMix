param(
    [string]$Version = "",
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Version)) {
    $versionFile = Join-Path $repoRoot "VERSION"
    if (-not (Test-Path $versionFile)) {
        throw "VERSION file not found: $versionFile"
    }
    $Version = (Get-Content -LiteralPath $versionFile -Raw).Trim()
}

$buildDll = Join-Path $repoRoot "build\$Configuration\micmix.dll"
if (-not (Test-Path $buildDll)) {
    throw "Build artifact missing: $buildDll. Build first (cmake --build build --config $Configuration)."
}

$icon16 = Join-Path $repoRoot "assets\branding\MicMixIcon16.png"
if (-not (Test-Path $icon16)) {
    throw "Missing icon: $icon16"
}

$distDir = Join-Path $repoRoot "dist"
$stageRoot = Join-Path $repoRoot "build\package_stage"
$stageDir = Join-Path $stageRoot "MicMix-$Version"
$pluginDir = Join-Path $stageDir "plugins"
$pluginIconDir = Join-Path $pluginDir "micmix_win64"

Remove-Item -LiteralPath $stageDir -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $pluginIconDir -Force | Out-Null
New-Item -ItemType Directory -Path $distDir -Force | Out-Null

Copy-Item -LiteralPath $buildDll -Destination (Join-Path $pluginDir "micmix_win64.dll") -Force
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

$zipPath = Join-Path $distDir "MicMix-$Version-win64.zip"
$ts3PluginPath = Join-Path $distDir "MicMix-$Version-win64.ts3_plugin"
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
