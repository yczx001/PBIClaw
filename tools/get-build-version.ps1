param(
    [switch]$Utc
)

$now = if ($Utc) { (Get-Date).ToUniversalTime() } else { Get-Date }

# 4-part numeric version compatible with AssemblyVersion:
# yyyy.M.(ddHH).(mmss)
$version = "{0}.{1}.{2}.{3}" -f `
    $now.Year, `
    $now.Month, `
    ($now.Day * 100 + $now.Hour), `
    ($now.Minute * 100 + $now.Second)

Write-Output $version
