# ------------------------------------------------------------
# 1. Overview of Script
#   - OS 名（例: Windows 10, Windows Server 2019）
#   - ビルド番号 (例: 19041)
#   - バージョン文字列 (例: 10.0.19041)
#   - サブバージョン/ビルド日付を取得する場合は追加
# ------------------------------------------------------------

# 2. get system info
$os = Get-CimInstance Win32_OperatingSystem

# 3, extract necessary property
$winName      = $os.Caption           # "Microsoft Windows 10 Pro"
$versionFull  = $os.Version           # "10.0.19041"
$buildNumber  = $os.BuildNumber       # "19041"

# 4. Output
Write-Host "OS: $winName"
Write-Host "version: $versionFull"
Write-Host "build Number: $buildNumber"

# 5. If you want to use JSON
#$result = @{
#    Name          = $winName
#    Version       = $versionFull
#    BuildNumber   = $buildNumber
#}
#[Console]::WriteLine((ConvertTo-Json $result))
