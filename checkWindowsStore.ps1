<#
Updated 25th Nov 2025

.SYNOPSIS
    Windows Store Availability Check and Registry Update Script

.DESCRIPTION
    Checks if Windows Store is available and enabled.
    If restricted, updates registry to enable Windows Store.
    
    Checks multiple methods:
    1. Registry policy check (RemoveWindowsStore)
    2. Windows Store app installation status
    3. AppX package status

.NOTES
    Author: Auto-generated
    Requires: Administrator privileges for registry updates
#>

#region Configuration Variables
# Registry paths for Windows Store policy
$REGISTRY_POLICY_PATH = "HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore"
$REGISTRY_POLICY_VALUE = "RemoveWindowsStore"

# Windows Store app name
$WINDOWS_STORE_APP_NAME = "Microsoft.WindowsStore"
#endregion

#region Initialization
$script:isAvailable = $false
$script:isRestricted = $false
$script:restrictionReason = $null
$script:exitCode = 0
$script:registryUpdated = $false
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

function Test-Administrator {
    <#
    .SYNOPSIS
        Checks if script is running with administrator privileges
    #>
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-WindowsStoreRegistryPolicy {
    <#
    .SYNOPSIS
        Checks Windows Store registry policy to determine if it's restricted
    #>
    Write-SectionHeader "Method 1: Registry Policy Check"
    
    try {
        if (-not (Test-Path $REGISTRY_POLICY_PATH)) {
            Write-Info "Registry policy path not found: $REGISTRY_POLICY_PATH"
            Write-Info "Windows Store is not restricted by policy (default: enabled)"
            return $false
        }
        
        $policy = Get-ItemProperty -Path $REGISTRY_POLICY_PATH -Name $REGISTRY_POLICY_VALUE -ErrorAction SilentlyContinue
        
        if ($null -eq $policy) {
            Write-Info "Policy value '$REGISTRY_POLICY_VALUE' not found"
            Write-Info "Windows Store is not restricted by this policy"
            return $false
        }
        
        $removeStoreValue = $policy.$REGISTRY_POLICY_VALUE
        
        Write-Info "Found registry policy: $REGISTRY_POLICY_PATH\$REGISTRY_POLICY_VALUE"
        Write-Info "Current value: $removeStoreValue"
        
        if ($removeStoreValue -eq 1) {
            Write-WarningMessage "Windows Store is RESTRICTED by registry policy"
            $script:restrictionReason = "Registry policy RemoveWindowsStore = 1"
            return $true
        }
        elseif ($removeStoreValue -eq 0) {
            Write-Success "Windows Store is ENABLED by registry policy"
            return $false
        }
        else {
            Write-Info "Unexpected policy value: $removeStoreValue"
            return $false
        }
    }
    catch {
        Write-WarningMessage "Error checking registry policy: $($_.Exception.Message)"
        return $false
    }
}

function Test-WindowsStoreAppInstalled {
    <#
    .SYNOPSIS
        Checks if Windows Store app is installed
    #>
    Write-SectionHeader "Method 2: Windows Store App Installation Check"
    
    try {
        # Check using Get-AppxPackage
        $storeApp = Get-AppxPackage -Name $WINDOWS_STORE_APP_NAME -ErrorAction SilentlyContinue
        
        if ($storeApp) {
            Write-Success "Windows Store app is installed"
            Write-Info "Package Name: $($storeApp.Name)"
            Write-Info "Version: $($storeApp.Version)"
            Write-Info "Publisher: $($storeApp.Publisher)"
            
            # Check if app is provisioned for all users
            $provisionedApp = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq $WINDOWS_STORE_APP_NAME } -ErrorAction SilentlyContinue
            if ($provisionedApp) {
                Write-Info "Status: Provisioned for all users"
            }
            else {
                Write-Info "Status: Installed for current user only"
            }
            
            return $true
        }
        else {
            Write-WarningMessage "Windows Store app is NOT installed"
            Write-Info "This may indicate Windows Store has been removed"
            return $false
        }
    }
    catch {
        Write-WarningMessage "Error checking Windows Store app: $($_.Exception.Message)"
        Write-Info "Note: Get-AppxPackage may not be available on Windows Server"
        return $false
    }
}

function Test-WindowsStoreService {
    <#
    .SYNOPSIS
        Checks Windows Store related services
    #>
    Write-SectionHeader "Method 3: Windows Store Service Check"
    
    try {
        # Check AppX Deployment Service
        $appxService = Get-Service -Name "AppXSvc" -ErrorAction SilentlyContinue
        
        if ($appxService) {
            Write-Info "AppX Deployment Service found"
            Write-Info "Status: $($appxService.Status)"
            Write-Info "StartType: $($appxService.StartType)"
            
            if ($appxService.Status -eq "Running") {
                Write-Success "AppX Deployment Service is running"
                return $true
            }
            else {
                Write-WarningMessage "AppX Deployment Service is not running (Status: $($appxService.Status))"
                return $false
            }
        }
        else {
            Write-Info "AppX Deployment Service not found (may not be available on this Windows edition)"
            return $false
        }
    }
    catch {
        Write-WarningMessage "Error checking AppX service: $($_.Exception.Message)"
        return $false
    }
}

function Update-WindowsStoreRegistry {
    <#
    .SYNOPSIS
        Updates registry to enable Windows Store
    #>
    Write-SectionHeader "Updating Registry to Enable Windows Store"
    
    # Check administrator privileges
    if (-not (Test-Administrator)) {
        Write-WarningMessage "Administrator privileges required to update registry"
        Write-Info "Please run this script as Administrator"
        return $false
    }
    
    try {
        # Ensure registry path exists
        if (-not (Test-Path $REGISTRY_POLICY_PATH)) {
            Write-Info "Creating registry path: $REGISTRY_POLICY_PATH"
            New-Item -Path $REGISTRY_POLICY_PATH -Force | Out-Null
        }
        
        # Set RemoveWindowsStore to 0 (enabled)
        Write-Info "Setting $REGISTRY_POLICY_VALUE to 0 (enabled)..."
        Set-ItemProperty -Path $REGISTRY_POLICY_PATH -Name $REGISTRY_POLICY_VALUE -Value 0 -Type DWord -Force
        
        # Verify the change
        $verify = Get-ItemProperty -Path $REGISTRY_POLICY_PATH -Name $REGISTRY_POLICY_VALUE -ErrorAction Stop
        if ($verify.$REGISTRY_POLICY_VALUE -eq 0) {
            Write-Success "Registry updated successfully"
            Write-Info "Windows Store policy is now set to ENABLED"
            $script:registryUpdated = $true
            return $true
        }
        else {
            Write-WarningMessage "Registry update verification failed"
            return $false
        }
    }
    catch {
        Write-WarningMessage "Error updating registry: $($_.Exception.Message)"
        return $false
    }
}

function Remove-WindowsStoreRegistryRestriction {
    <#
    .SYNOPSIS
        Removes the registry restriction entirely (alternative method)
    #>
    Write-SectionHeader "Removing Registry Restriction"
    
    # Check administrator privileges
    if (-not (Test-Administrator)) {
        Write-WarningMessage "Administrator privileges required"
        return $false
    }
    
    try {
        if (Test-Path $REGISTRY_POLICY_PATH) {
            $policy = Get-ItemProperty -Path $REGISTRY_POLICY_PATH -Name $REGISTRY_POLICY_VALUE -ErrorAction SilentlyContinue
            
            if ($policy) {
                Write-Info "Removing registry value: $REGISTRY_POLICY_VALUE"
                Remove-ItemProperty -Path $REGISTRY_POLICY_PATH -Name $REGISTRY_POLICY_VALUE -Force
                Write-Success "Registry restriction removed"
                $script:registryUpdated = $true
                return $true
            }
            else {
                Write-Info "Registry restriction value does not exist"
                return $true
            }
        }
        else {
            Write-Info "Registry policy path does not exist - no restriction to remove"
            return $true
        }
    }
    catch {
        Write-WarningMessage "Error removing registry restriction: $($_.Exception.Message)"
        return $false
    }
}

function Show-Summary {
    <#
    .SYNOPSIS
        Displays final results
    #>
    Write-Output "`n======================================"
    Write-Output "Summary"
    Write-Output "======================================"
    
    if ($script:isRestricted) {
        Write-WarningMessage "Windows Store is RESTRICTED"
        Write-Info "Reason: $script:restrictionReason"
        
        if ($script:registryUpdated) {
            Write-Output "`n✓ Registry has been updated to enable Windows Store"
            Write-Info "Note: You may need to restart your computer for changes to take effect"
            Write-Info "Note: If Windows Store app is not installed, you may need to reinstall it"
        }
        else {
            Write-Output "`n⚠ Registry update was not performed"
            Write-Info "To enable Windows Store, run this script as Administrator"
        }
    }
    else {
        Write-Success "Windows Store appears to be AVAILABLE"
        Write-Info "No registry restrictions found"
        
        if (-not $script:isAvailable) {
            Write-Info "Note: Windows Store app may not be installed, but it's not restricted by policy"
        }
    }
    
    Write-Output "======================================"
}
#endregion

#region Main Processing
Write-Output "======================================"
Write-Output "Windows Store Availability Check Script"
Write-Output "Date: 25th Nov 2025, Author: mmiy & GenAI"
Write-Output "======================================"

# Check if running as administrator
$isAdmin = Test-Administrator
if ($isAdmin) {
    Write-Success "Running with Administrator privileges"
}
else {
    Write-Info "Running without Administrator privileges (registry updates will be skipped)"
}

# Check Windows Store availability using multiple methods
$registryRestricted = Test-WindowsStoreRegistryPolicy
$appInstalled = Test-WindowsStoreAppInstalled
$serviceRunning = Test-WindowsStoreService

# Determine overall status
if ($registryRestricted) {
    $script:isRestricted = $true
    $script:isAvailable = $false
}
else {
    $script:isRestricted = $false
    if ($appInstalled -or $serviceRunning) {
        $script:isAvailable = $true
    }
    else {
        $script:isAvailable = $false
        Write-Info "Windows Store app may not be installed, but no policy restrictions found"
    }
}

# Display summary
Show-Summary

# If restricted and running as admin, update registry
if ($script:isRestricted -and $isAdmin) {
    Write-Output "`n======================================"
    Write-Output "Attempting to Enable Windows Store"
    Write-Output "======================================"
    
    # Try to update registry (set to 0)
    $updateSuccess = Update-WindowsStoreRegistry
    
    if (-not $updateSuccess) {
        Write-Info "`nTrying alternative method: Removing registry restriction..."
        $removeSuccess = Remove-WindowsStoreRegistryRestriction
        
        if ($removeSuccess) {
            Write-Success "Registry restriction removed successfully"
        }
    }
    
    Write-Output "`n======================================"
    if ($script:registryUpdated) {
        Write-Success "Windows Store has been enabled via registry update"
        Write-Info "Please restart your computer for changes to take full effect"
        Write-Info "If Windows Store app is missing, you may need to reinstall it using:"
        Write-Info "  Get-AppxPackage -allusers Microsoft.WindowsStore | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register `"$($_.InstallLocation)\AppXManifest.xml`"}"
        $script:exitCode = 0
    }
    else {
        Write-WarningMessage "Failed to update registry"
        Write-Info "Please check error messages above"
        $script:exitCode = 1
    }
    Write-Output "======================================"
}
elseif ($script:isRestricted -and -not $isAdmin) {
    Write-Output "`n======================================"
    Write-WarningMessage "Windows Store is restricted, but cannot update registry"
    Write-Info "Please run this script as Administrator to enable Windows Store"
    Write-Output "======================================"
    $script:exitCode = 1
}

exit $script:exitCode
#endregion

