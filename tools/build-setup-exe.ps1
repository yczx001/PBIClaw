param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$Version = "",
    [switch]$KeepIntermediates
)

$ErrorActionPreference = "Stop"
$env:DOTNET_CLI_UI_LANGUAGE = "en-US"
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = "1"
$env:DOTNET_CLI_TELEMETRY_OPTOUT = "1"
$env:DOTNET_NOLOGO = "1"

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw "dotnet not found. Install .NET SDK 8 or later."
}

$sdkList = & dotnet --list-sdks 2>$null
if (-not $sdkList) {
    throw "No .NET SDK detected. Install .NET SDK 8 (x64): https://aka.ms/dotnet-download"
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$toolProject = Join-Path $repoRoot "src\PbiMetadataTool\PbiMetadataTool.csproj"
$setupProject = Join-Path $repoRoot "src\PbiMetadataInstaller\PbiMetadataInstaller.csproj"
$toolJsonTemplate = Join-Path $repoRoot "external-tools\PBIClaw.pbitool.json"
$versionScript = Join-Path $repoRoot "tools\get-build-version.ps1"

if (-not (Test-Path $toolProject)) { throw "Missing tool project: $toolProject" }
if (-not (Test-Path $setupProject)) { throw "Missing setup project: $setupProject" }
if (-not (Test-Path $toolJsonTemplate)) { throw "Missing pbitool json: $toolJsonTemplate" }
if (-not (Test-Path $versionScript)) { throw "Missing version script: $versionScript" }

if ([string]::IsNullOrWhiteSpace($Version)) {
    $Version = (& $versionScript).Trim()
}
Write-Host "Build version: $Version"

$toolPublishDir = Join-Path $repoRoot "dist\publish"
$setupPayloadDir = Join-Path $repoRoot "src\PbiMetadataInstaller\Resources\payload"
$setupOutDir = Join-Path $repoRoot "dist\setup"

foreach ($d in @($toolPublishDir, $setupPayloadDir, $setupOutDir)) {
    if (Test-Path $d) {
        Remove-Item $d -Recurse -Force
    }
    New-Item -ItemType Directory -Path $d -Force | Out-Null
}

Write-Host "Step 1/3: publish PBIClaw.exe..."
& dotnet publish $toolProject `
    -c $Configuration `
    -r $Runtime `
    --self-contained false `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -o $toolPublishDir

if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed for tool."
}

$toolExe = Join-Path $toolPublishDir "PBIClaw.exe"
if (-not (Test-Path $toolExe)) {
    throw "Tool exe not found: $toolExe"
}

Write-Host "Step 2/3: stage installer payload..."
$payloadExe = Join-Path $setupPayloadDir "PBIClaw.exe"
$payloadJson = Join-Path $setupPayloadDir "PBIClaw.pbitool.json"

Copy-Item $toolExe $payloadExe -Force
Copy-Item $toolJsonTemplate $payloadJson -Force

$json = Get-Content $payloadJson -Raw | ConvertFrom-Json
$json.path = "PBIClaw.exe"
$json.version = $Version
$json.name = "PBI Claw"
if (-not [string]::IsNullOrWhiteSpace($json.iconData) -and $json.iconData -notmatch '^[a-z]+/[a-z0-9+.-]+;base64,') {
    $json.iconData = "image/png;base64,$($json.iconData.Trim())"
}
$json | ConvertTo-Json -Depth 20 | Set-Content $payloadJson -Encoding UTF8

Write-Host "Step 3/3: build setup exe..."
& dotnet publish $setupProject `
    -c $Configuration `
    -r $Runtime `
    --self-contained false `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:Version=$Version `
    -p:FileVersion=$Version `
    -p:AssemblyVersion=$Version `
    -o $setupOutDir

if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed for setup."
}

$setupExe = Join-Path $setupOutDir "PBIClawSetup.exe"
if (-not (Test-Path $setupExe)) {
    throw "Setup exe not found: $setupExe"
}

Write-Host ""
Write-Host "Setup EXE created:"
Write-Host $setupExe
Write-Host "Tool version: $Version"
Write-Host ""
Write-Host "Double-click setup exe, choose EXE install folder, then allow Administrator permission."
Write-Host "Optional silent style argument:"
Write-Host "`"$setupExe`" --install-dir `"C:\Program Files\PBI Claw`""

if (-not $KeepIntermediates) {
    if (Test-Path $toolPublishDir) {
        Remove-Item $toolPublishDir -Recurse -Force
    }

    if (Test-Path $setupPayloadDir) {
        Remove-Item $setupPayloadDir -Recurse -Force
    }
}
