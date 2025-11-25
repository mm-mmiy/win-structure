<#
.SYNOPSIS
    LanScope Cat Version Check Script

.DESCRIPTION
    Searches for LanScope Cat version information using multiple methods.
    Priority order:
    1. LSP*.ver file in C:\Windows (Official method)
    2. Direct registry access
    3. Uninstall registry keys
    4. Get-Package cmdlet (Package Manager)

.NOTES
    Author: Auto-generated
    Latest Version: 9472
#>

#region Configuration Variables
# Latest version number
$LATEST_VERSION = "9472"

# Registry paths (in priority order)
$REGISTRY_PATHS = @(
    "HKLM:\SOFTWARE\WOW6432Node\LanScopeCat\Agent",
    "HKLM:\SOFTWARE\LanScopeCat\Agent",
    "HKLM:\SOFTWARE\WOW6432Node\LanScopeCat",
    "HKLM:\SOFTWARE\LanScopeCat"
)

# Uninstall registry paths
$UNINSTALL_REGISTRY_PATHS = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*LanScope*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*LanScope*"
)

# Windows directory path
$WINDOWS_PATH = "C:\Windows"

# Box download URL for update file
$LSC_LATEST_IN_BOX = "https://mmmacromill.box.com/s/ct3ab42y2atjxtw3a8ogbvtgbfm0mpfw"

# Temporary directory for downloads
$TEMP_DOWNLOAD_DIR = "$env:TEMP\LanScopeCat_Update"
#endregion

#region Initialization
$script:found = $false
$script:exitCode = 1
$script:versionNumber = $null
$script:versionSource = $null
#endregion

#region Function Definitions
function Write-SectionHeader {
    param([string]$Title)
    Write-Output "`n[$Title]"
}

function Write-Success {
    param([string]$Message)
    Write-Output "  ✓ $Message"
}

function Write-WarningMessage {
    param([string]$Message)
    Write-Warning "  ⚠ $Message"
}

function Write-Info {
    param([string]$Message)
    Write-Output "  $Message"
}

function Get-VersionFromLspFile {
    <#
    .SYNOPSIS
        Retrieves version number from LSP*.ver file in C:\Windows
    #>
    Write-SectionHeader "Method 1: LSP*.ver File Check (Official Method)"
    
    try {
        if (-not (Test-Path $WINDOWS_PATH)) {
            Write-Info "Windows directory not found: $WINDOWS_PATH"
            return $false
        }
        
        $lspFiles = Get-ChildItem -Path $WINDOWS_PATH -Filter "LSP*.ver" -ErrorAction SilentlyContinue
        
        if (-not $lspFiles) {
            Write-Info "LSP*.ver file not found in: $WINDOWS_PATH"
            return $false
        }
        
        foreach ($file in $lspFiles) {
            Write-Info "File found: $($file.Name)"
            Write-Info "Full path: $($file.FullName)"
            
            # Extract version number from filename (e.g., LSP9472.ver -> 9472)
            if ($file.Name -match "LSP(\d+)\.ver") {
                $script:versionNumber = $matches[1]
                Write-Info "Extracted version number: $script:versionNumber"
                
                # Check if it's the latest version
                if ($script:versionNumber -eq $LATEST_VERSION) {
                    Write-Success "Latest version ($LATEST_VERSION) - No update required"
                }
                else {
                    Write-WarningMessage "Version $script:versionNumber (Latest: $LATEST_VERSION) - Update required"
                }
                
                # Display file content (optional)
                try {
                    $fileContent = Get-Content -Path $file.FullName -ErrorAction SilentlyContinue
                    if ($fileContent) {
                        Write-Info "File content:"
                        $fileContent | ForEach-Object { Write-Output "    $_" }
                    }
                }
                catch {
                    # Ignore file content read errors
                }
                
                $script:versionSource = "LSP*.ver file"
                return $true
            }
            else {
                Write-WarningMessage "Could not extract version number from filename: $($file.Name)"
            }
        }
    }
    catch {
        Write-WarningMessage "Error checking LSP*.ver file: $($_.Exception.Message)"
    }
    
    return $false
}

function Get-VersionFromRegistry {
    <#
    .SYNOPSIS
        Retrieves version information directly from registry
    #>
    Write-SectionHeader "Method 2: Direct Registry Check"
    
    foreach ($path in $REGISTRY_PATHS) {
        if (-not (Test-Path $path)) {
            continue
        }
        
        try {
            $info = Get-ItemProperty -Path $path -ErrorAction Stop
            
            # Version information priority: DisplayVersion > ProductVersion > Version
            $version = $null
            if ($info.DisplayVersion) {
                $version = $info.DisplayVersion
            }
            elseif ($info.ProductVersion) {
                $version = $info.ProductVersion
            }
            elseif ($info.Version) {
                $version = $info.Version
            }
            
            if ($version) {
                Write-Success "Found in: $path"
                Write-Info "Version: $version"
                
                # Display other version fields as well
                if ($info.Version -and $info.Version -ne $version) {
                    Write-Info "  (Version field: $($info.Version))"
                }
                if ($info.ProductVersion -and $info.ProductVersion -ne $version) {
                    Write-Info "  (ProductVersion field: $($info.ProductVersion))"
                }
                if ($info.DisplayVersion -and $info.DisplayVersion -ne $version) {
                    Write-Info "  (DisplayVersion field: $($info.DisplayVersion))"
                }
                
                $script:versionSource = "Registry: $path"
                return $true
            }
        }
        catch {
            Write-WarningMessage "Error reading registry path $path : $($_.Exception.Message)"
        }
    }
    
    return $false
}

function Get-VersionFromUninstallRegistry {
    <#
    .SYNOPSIS
        Retrieves version information from uninstall registry keys
    #>
    Write-SectionHeader "Method 3: Uninstall Registry Keys Check"
    
    foreach ($pattern in $UNINSTALL_REGISTRY_PATHS) {
        try {
            $uninstallKeys = Get-ItemProperty -Path $pattern -ErrorAction SilentlyContinue
            if (-not $uninstallKeys) {
                continue
            }
            
            foreach ($key in $uninstallKeys) {
                $displayName = $key.DisplayName
                if ($displayName -match "LanScope|Lanscope" -and $key.DisplayVersion) {
                    Write-Success "Found in: $($key.PSPath)"
                    Write-Info "Display Name: $displayName"
                    Write-Info "Version: $($key.DisplayVersion)"
                    
                    if ($key.Publisher) {
                        Write-Info "Publisher: $($key.Publisher)"
                    }
                    
                    $script:versionSource = "Uninstall Registry: $($key.PSPath)"
                    return $true
                }
            }
        }
        catch {
            # Ignore errors and continue
        }
    }
    
    return $false
}

function Get-VersionFromPackage {
    <#
    .SYNOPSIS
        Retrieves package information via Get-Package cmdlet
    #>
    Write-SectionHeader "Method 4: Get-Package Check (Package Manager)"
    Write-Info "Note: Get-Package retrieves information from package providers (Programs, msi, msu, etc.)"
    
    try {
        # Wildcard search
        $packages = Get-Package -Name "*LanScope*" -ErrorAction SilentlyContinue
        
        if (-not $packages) {
            # Case-insensitive search
            $packages = Get-Package -ErrorAction SilentlyContinue | Where-Object { 
                $_.Name -match "LanScope|Lanscope" 
            }
        }
        
        if ($packages) {
            Write-Info "Found $($packages.Count) package(s):"
            
            foreach ($package in $packages) {
                Write-Success "Package found"
                Write-Info "Name: $($package.Name)"
                Write-Info "Version: $($package.Version)"
                
                if ($package.ProviderName) {
                    Write-Info "Provider: $($package.ProviderName)"
                }
                if ($package.Source) {
                    Write-Info "Source: $($package.Source)"
                }
                if ($package.Status) {
                    Write-Info "Status: $($package.Status)"
                }
                
                if ($package.Version) {
                    $script:versionSource = "Get-Package: $($package.Name)"
                    return $true
                }
            }
        }
        else {
            Write-Info "No packages found matching 'LanScope'"
        }
    }
    catch {
        Write-WarningMessage "Get-Package failed: $($_.Exception.Message)"
        Write-Info "Note: Get-Package requires PackageManagement module (PowerShell 5.0+)"
    }
    
    return $false
}

function Compare-Version {
    <#
    .SYNOPSIS
        Compares current version with latest version
    .OUTPUTS
        Returns $true if update is required, $false otherwise
    #>
    param(
        [string]$CurrentVersion,
        [string]$LatestVersion
    )
    
    if ([string]::IsNullOrEmpty($CurrentVersion)) {
        return $false
    }
    
    try {
        $current = [int]$CurrentVersion
        $latest = [int]$LatestVersion
        return $current -lt $latest
    }
    catch {
        Write-WarningMessage "Error comparing versions: $($_.Exception.Message)"
        return $false
    }
}

function Download-UpdateFromBox {
    <#
    .SYNOPSIS
        Downloads update file from Box
    #>
    Write-SectionHeader "Downloading Update from Box"
    
    try {
        # Create temporary directory if it doesn't exist
        if (-not (Test-Path $TEMP_DOWNLOAD_DIR)) {
            New-Item -ItemType Directory -Path $TEMP_DOWNLOAD_DIR -Force | Out-Null
            Write-Info "Created temporary directory: $TEMP_DOWNLOAD_DIR"
        }
        
        # Convert Box shared link to direct download URL
        # Box shared links need to be converted to direct download format
        $downloadUrl = $LSC_LATEST_IN_BOX
        
        # Try to get direct download URL from Box shared link
        # Box shared links format: https://app.box.com/s/... or https://mmmacromill.box.com/s/...
        # Direct download format: https://app.box.com/shared/static/... or https://dl.boxcloud.com/...
        if ($downloadUrl -match "box\.com/s/([^/]+)") {
            $shareId = $matches[1]
            # Try direct download URL format
            $directUrl = "https://app.box.com/shared/static/$shareId"
            Write-Info "Attempting to download from: $directUrl"
        }
        else {
            $directUrl = $downloadUrl
        }
        
        # Determine file extension (could be .exe, .msi, .zip, etc.)
        # For now, try common extensions
        $possibleExtensions = @(".exe", ".msi", ".zip", ".msu")
        $downloadedFile = $null
        
        foreach ($ext in $possibleExtensions) {
            $fileName = "LanScopeCat_Update$ext"
            $filePath = Join-Path $TEMP_DOWNLOAD_DIR $fileName
            
            try {
                Write-Info "Attempting to download: $fileName"
                
                # Use Invoke-WebRequest to download
                $response = Invoke-WebRequest -Uri $directUrl -OutFile $filePath -ErrorAction Stop
                
                if (Test-Path $filePath) {
                    $fileInfo = Get-Item $filePath
                    if ($fileInfo.Length -gt 0) {
                        Write-Success "File downloaded successfully: $filePath"
                        Write-Info "File size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB"
                        $downloadedFile = $filePath
                        break
                    }
                    else {
                        Remove-Item $filePath -Force
                    }
                }
            }
            catch {
                Write-Info "Failed to download with extension $ext : $($_.Exception.Message)"
                if (Test-Path $filePath) {
                    Remove-Item $filePath -Force -ErrorAction SilentlyContinue
                }
            }
        }
        
        # If direct download failed, try the original URL
        if (-not $downloadedFile) {
            Write-Info "Trying original Box URL..."
            $fileName = "LanScopeCat_Update.zip"
            $filePath = Join-Path $TEMP_DOWNLOAD_DIR $fileName
            
            try {
                # Box shared links might require authentication or special handling
                # Try using Invoke-WebRequest with headers
                $headers = @{
                    "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
                }
                
                $response = Invoke-WebRequest -Uri $LSC_LATEST_IN_BOX -Headers $headers -OutFile $filePath -ErrorAction Stop
                
                if (Test-Path $filePath -and (Get-Item $filePath).Length -gt 0) {
                    Write-Success "File downloaded successfully: $filePath"
                    $downloadedFile = $filePath
                }
            }
            catch {
                Write-WarningMessage "Failed to download from Box URL: $($_.Exception.Message)"
                Write-Info "Note: Box shared links may require authentication or manual download"
                Write-Info "Please download the file manually from: $LSC_LATEST_IN_BOX"
                return $null
            }
        }
        
        return $downloadedFile
    }
    catch {
        Write-WarningMessage "Error downloading update file: $($_.Exception.Message)"
        return $null
    }
}

function Install-Update {
    <#
    .SYNOPSIS
        Installs the downloaded update file
    #>
    param(
        [string]$UpdateFilePath
    )
    
    Write-SectionHeader "Installing Update"
    
    if (-not $UpdateFilePath -or -not (Test-Path $UpdateFilePath)) {
        Write-WarningMessage "Update file not found: $UpdateFilePath"
        return $false
    }
    
    try {
        $fileExtension = [System.IO.Path]::GetExtension($UpdateFilePath).ToLower()
        Write-Info "Update file: $UpdateFilePath"
        Write-Info "File type: $fileExtension"
        
        # Check if running as administrator
        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if (-not $isAdmin) {
            Write-WarningMessage "Administrator privileges required for installation"
            Write-Info "Please run this script as Administrator to install the update"
            return $false
        }
        
        switch ($fileExtension) {
            ".exe" {
                Write-Info "Executing installer..."
                $process = Start-Process -FilePath $UpdateFilePath -ArgumentList "/S", "/quiet" -Wait -PassThru
                if ($process.ExitCode -eq 0) {
                    Write-Success "Update installed successfully"
                    return $true
                }
                else {
                    Write-WarningMessage "Installer exited with code: $($process.ExitCode)"
                    return $false
                }
            }
            ".msi" {
                Write-Info "Installing MSI package..."
                $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i", "`"$UpdateFilePath`"", "/quiet", "/norestart" -Wait -PassThru
                if ($process.ExitCode -eq 0) {
                    Write-Success "Update installed successfully"
                    return $true
                }
                else {
                    Write-WarningMessage "MSI installer exited with code: $($process.ExitCode)"
                    return $false
                }
            }
            ".msu" {
                Write-Info "Installing Windows Update package..."
                $process = Start-Process -FilePath "wusa.exe" -ArgumentList "`"$UpdateFilePath`"", "/quiet", "/norestart" -Wait -PassThru
                if ($process.ExitCode -eq 0) {
                    Write-Success "Update installed successfully"
                    return $true
                }
                else {
                    Write-WarningMessage "Windows Update installer exited with code: $($process.ExitCode)"
                    return $false
                }
            }
            ".zip" {
                Write-Info "Extracting ZIP file..."
                $extractPath = Join-Path $TEMP_DOWNLOAD_DIR "Extracted"
                if (Test-Path $extractPath) {
                    Remove-Item $extractPath -Recurse -Force
                }
                Expand-Archive -Path $UpdateFilePath -DestinationPath $extractPath -Force
                
                # Look for installer in extracted files
                $installers = Get-ChildItem -Path $extractPath -Filter "*.exe", "*.msi" -Recurse
                if ($installers) {
                    $installer = $installers[0]
                    Write-Info "Found installer: $($installer.FullName)"
                    return Install-Update -UpdateFilePath $installer.FullName
                }
                else {
                    Write-WarningMessage "No installer found in ZIP file"
                    Write-Info "Please extract and run the installer manually from: $extractPath"
                    return $false
                }
            }
            default {
                Write-WarningMessage "Unsupported file type: $fileExtension"
                Write-Info "Please install the update manually: $UpdateFilePath"
                return $false
            }
        }
    }
    catch {
        Write-WarningMessage "Error installing update: $($_.Exception.Message)"
        return $false
    }
}

function Show-Summary {
    <#
    .SYNOPSIS
        Displays final results
    #>
    Write-Output "`n======================================"
    
    if ($script:found) {
        if ($script:versionNumber) {
            Write-Success "LanScope Cat version information found"
            Write-Info "Version: $script:versionNumber"
            Write-Info "Source: $script:versionSource"
            
            if ($script:versionNumber -eq $LATEST_VERSION) {
                Write-Info "Status: Latest version - No update required"
            }
            else {
                Write-Info "Status: Update required (Latest: $LATEST_VERSION)"
            }
        }
        else {
            Write-Success "LanScope Cat version information found (via other methods)"
            Write-Info "Source: $script:versionSource"
        }
    }
    else {
        Write-Output "✗ Cannot find version information of LanScope Cat"
        Write-Output "`nChecked methods:"
        Write-Output "  [Method 1] LSP*.ver file in C:\Windows: Not found"
        Write-Output "  [Method 2] Direct registry paths:"
        
        foreach ($path in $REGISTRY_PATHS) {
            $exists = if (Test-Path $path) { "Exists (but no version info)" } else { "Not found" }
            Write-Output "    $path : $exists"
        }
        
        Write-Output "  [Method 3] Uninstall registry keys: No version information found"
        Write-Output "  [Method 4] Get-Package: No version information found"
        Write-Output "`nSuggestions:"
        Write-Output "  - LanScope Cat may not be installed on this system"
        Write-Output "  - Manually check C:\Windows folder for LSP*.ver files"
        Write-Output "  - Verify installation directory manually"
        Write-Output "  - Check if running with administrator privileges"
    }
}
#endregion

#region Main Processing
Write-Output "======================================"
Write-Output "LanScope Cat Version Check Script"
Write-Output "Version 0.0.1, Auther: mmiy powered by Gen AI"
Write-Output "======================================"

# Try each method in order (exit when found)
if (Get-VersionFromLspFile) {
    $script:found = $true
    $script:exitCode = 0
}
elseif (Get-VersionFromRegistry) {
    $script:found = $true
    $script:exitCode = 0
}
elseif (Get-VersionFromUninstallRegistry) {
    $script:found = $true
    $script:exitCode = 0
}
elseif (Get-VersionFromPackage) {
    $script:found = $true
    $script:exitCode = 0
}

# Display final results
Show-Summary

# Check if update is required and perform update
if ($script:found -and $script:versionNumber) {
    if (Compare-Version -CurrentVersion $script:versionNumber -LatestVersion $LATEST_VERSION) {
        Write-Output "`n======================================"
        Write-Output "Update Required - Starting Update Process"
        Write-Output "======================================"
        
        $updateFile = Download-UpdateFromBox
        
        if ($updateFile) {
            $updateSuccess = Install-Update -UpdateFilePath $updateFile
            
            if ($updateSuccess) {
                Write-Output "`n======================================"
                Write-Success "Update completed successfully"
                Write-Info "Please restart your computer if required"
                Write-Output "======================================"
                $script:exitCode = 0
            }
            else {
                Write-Output "`n======================================"
                Write-WarningMessage "Update installation failed"
                Write-Info "Update file is available at: $updateFile"
                Write-Info "Please install manually if needed"
                Write-Output "======================================"
                $script:exitCode = 1
            }
        }
        else {
            Write-Output "`n======================================"
            Write-WarningMessage "Failed to download update file"
            Write-Info "Please download manually from: $LSC_LATEST_IN_BOX"
            Write-Output "======================================"
            $script:exitCode = 1
        }
    }
}

exit $script:exitCode
#endregion
