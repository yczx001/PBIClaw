param(
    [switch]$Deep
)

$ErrorActionPreference = "SilentlyContinue"

function Remove-PathIfExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [switch]$Recurse
    )

    if (Test-Path -LiteralPath $Path) {
        if ($Recurse) {
            Remove-Item -LiteralPath $Path -Recurse -Force
        } else {
            Remove-Item -LiteralPath $Path -Force
        }
    }
}

Write-Host "== Clear Windows Icon Cache =="
Write-Host ""

Write-Host "[1/5] Stop Explorer..."
Get-Process explorer | Stop-Process -Force
Start-Sleep -Milliseconds 900

Write-Host "[2/5] Remove icon/thumb cache files..."
Remove-PathIfExists "$env:LOCALAPPDATA\IconCache.db"
Remove-PathIfExists "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\iconcache*" -Recurse
Remove-PathIfExists "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\thumbcache*" -Recurse

if ($Deep) {
    Write-Host "[3/5] Deep clean StartMenu/Shell caches..."
    Get-Process StartMenuExperienceHost | Stop-Process -Force
    Get-Process ShellExperienceHost | Stop-Process -Force
    Start-Sleep -Milliseconds 500

    Remove-PathIfExists "$env:LOCALAPPDATA\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\*" -Recurse
    Remove-PathIfExists "$env:LOCALAPPDATA\Packages\Microsoft.Windows.ShellExperienceHost_cw5n1h2txyewy\TempState\*" -Recurse
} else {
    Write-Host "[3/5] Skip deep clean (-Deep not specified)..."
}

Write-Host "[4/5] Rebuild icon cache..."
& "$env:SystemRoot\System32\ie4uinit.exe" -ClearIconCache | Out-Null

Write-Host "[5/5] Start Explorer..."
Start-Process explorer.exe

Write-Host ""
Write-Host "Done. If old icons remain, unpin/pin again and reboot if needed."
