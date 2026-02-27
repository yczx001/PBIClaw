param(
    [switch]$NoBuild,
    [string]$Version,
    [switch]$Machine
)

$ErrorActionPreference = "Stop"
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$projectPath = Join-Path $repoRoot "src\PbiMetadataTool\PbiMetadataTool.csproj"
$debugExe = Join-Path $repoRoot "src\PbiMetadataTool\bin\Debug\net8.0-windows\PBIClaw.exe"
$toolJsonTemplate = Join-Path $repoRoot "external-tools\PBIClaw.pbitool.json"
$versionScript = Join-Path $repoRoot "tools\get-build-version.ps1"

if (-not (Test-Path $projectPath)) {
    throw "Missing project: $projectPath"
}
if (-not (Test-Path $toolJsonTemplate)) {
    throw "Missing pbitool json template: $toolJsonTemplate"
}
if (-not (Test-Path $versionScript)) {
    throw "Missing version script: $versionScript"
}

if ([string]::IsNullOrWhiteSpace($Version)) {
    $Version = (& $versionScript).Trim()
}

if (-not $NoBuild) {
    Write-Host "Building Debug executable..."
    & dotnet build $projectPath -c Debug `
        -p:Version=$Version `
        -p:FileVersion=$Version `
        -p:AssemblyVersion=$Version
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet build failed."
    }
}

if (-not (Test-Path $debugExe)) {
    throw "Debug exe not found: $debugExe"
}

$template = Get-Content $toolJsonTemplate -Raw | ConvertFrom-Json
$iconData = $template.iconData
if (-not [string]::IsNullOrWhiteSpace($iconData) -and $iconData -notmatch '^[a-z]+/[a-z0-9+.-]+;base64,') {
    $iconData = "image/png;base64,$($iconData.Trim())"
}

if ($Machine) {
    $targets = @()
    if (-not [string]::IsNullOrWhiteSpace($env:CommonProgramFiles)) {
        $targets += (Join-Path $env:CommonProgramFiles "Microsoft Shared\Power BI Desktop\External Tools")
    }
    $commonX86 = [Environment]::GetEnvironmentVariable("CommonProgramFiles(x86)")
    if (-not [string]::IsNullOrWhiteSpace($commonX86)) {
        $targets += (Join-Path $commonX86 "Microsoft Shared\Power BI Desktop\External Tools")
    }
    $targets = $targets | Select-Object -Unique
} else {
    $docs = [Environment]::GetFolderPath("MyDocuments")
    $targets = @(
        (Join-Path $docs "Power BI Desktop\External Tools"),
        (Join-Path $docs "Power BI Desktop Store App\External Tools")
    ) | Select-Object -Unique
}

$written = @()
foreach ($externalToolsDir in $targets) {
    try {
        New-Item -ItemType Directory -Path $externalToolsDir -Force | Out-Null
        $jsonPath = Join-Path $externalToolsDir "PBIClaw.Debug.pbitool.json"
        $payload = [ordered]@{
            name = "PBI Claw (Debug) v$Version"
            description = "Local debug entry, no installer required"
            path = $debugExe
            arguments = "--server `"%server%`" --database `"%database%`" --external-tool"
            iconData = $iconData
            version = $Version
        }
        $payload | ConvertTo-Json -Depth 20 | Set-Content -Path $jsonPath -Encoding UTF8
        $written += $jsonPath
    }
    catch [System.UnauthorizedAccessException] {
        throw "No permission to write: $externalToolsDir . Re-run this script in an Administrator terminal."
    }
}

Write-Host ""
Write-Host "Debug external tool registered:"
$written | ForEach-Object { Write-Host $_ }
$scopeText = if ($Machine) { "machine" } else { "user" }
Write-Host "Scope: $scopeText"
Write-Host "Version: $Version"
Write-Host ""
Write-Host "Usage:"
Write-Host "1) Keep this repo on disk"
Write-Host "2) Restart Power BI Desktop once, then click 'PBI Claw (Debug) v$Version'"
Write-Host "3) After each code change, rerun: register-debug-external-tool.ps1 (auto bumps version)"
Write-Host "4) Then click the debug tool again (no reinstall needed)"
