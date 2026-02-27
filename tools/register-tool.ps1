[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'

$root = Split-Path $PSScriptRoot -Parent
$exePath = Join-Path $root 'src\PbiMetadataTool\bin\Debug\net8.0-windows\PBIClaw.exe'
$targetDir = 'C:\Program Files (x86)\Common Files\Microsoft Shared\Power BI Desktop\External Tools'
$jsonPath = Join-Path $targetDir 'PBIClaw.Debug.pbitool.json'

# 编译
if (-not (Test-Path $exePath)) {
    Write-Host '未找到可执行文件，正在编译...'
    $proj = Join-Path $root 'src\PbiMetadataTool\PbiMetadataTool.csproj'
    dotnet build $proj -c Debug
    if ($LASTEXITCODE -ne 0) { Write-Host '编译失败'; Read-Host; exit 1 }
}

# 写 json
New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
$template = Get-Content (Join-Path $root 'external-tools\PBIClaw.pbitool.json') | ConvertFrom-Json
$iconData = "image/png;base64,$($template.iconData)"

$json = [ordered]@{
    name        = 'PBI Claw'
    description = 'PBI Claw Power BI External Tool'
    path        = $exePath
    arguments   = '--server "%server%" --database "%database%" --external-tool'
    iconData    = $iconData
    version     = '1.0.0'
}
$json | ConvertTo-Json | Set-Content -Path $jsonPath -Encoding UTF8

Write-Host "注册成功: $jsonPath"
Write-Host '请重启 Power BI Desktop'
Read-Host '按 Enter 退出'
