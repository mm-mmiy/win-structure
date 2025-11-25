# ------------------------------------------------------------
# 1. Overview of Script
#   - OS Name (e.g.: Windows 10, Windows Server 2019)
#   - BUild number (e.g.: 19041)
#   - Version (e.g.: 10.0.19041)
#   - Sub Version/ Add ths to get build date
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
