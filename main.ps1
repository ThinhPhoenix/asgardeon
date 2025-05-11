# Asgardeon Loader Script
# Version: 1.0

# Load required .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Base URL for fetching modules
$baseUrl = "https://raw.githubusercontent.com/thinhphoenix/asgardeon/main/methods"

function Import-AsgardeonModule {
    param (
        [string]$ModuleName
    )
    
    try {
        $moduleUrl = "$baseUrl/$ModuleName.ps1"
        $moduleContent = Invoke-RestMethod -Uri $moduleUrl -UseBasicParsing
        Invoke-Expression $moduleContent
        return $true
    } catch {
        Write-Error "Failed to import module $ModuleName from $moduleUrl. Error: $_"
        return $false
    }
}

function Initialize-Asgardeon {
    # Load all required modules
    $modules = @(
        "resetTaskBar", 
        "setWallpaper", 
        "setTheme", 
        "setLockScreen", 
        "hideActivate", 
        "activateWin", 
        "asgardeonUI"
    )
    
    $failed = $false
    foreach ($module in $modules) {
        $success = Import-AsgardeonModule -ModuleName $module
        if (-not $success) {
            $failed = $true
            break
        }
    }
    
    if ($failed) {
        [System.Windows.Forms.MessageBox]::Show("Failed to load one or more Asgardeon modules. Please check your internet connection and try again.", "Module Loading Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    
    return $true
}

function Start-Asgardeon {
    $initSuccess = Initialize-Asgardeon
    if ($initSuccess) {
        # Start the main UI
        Show-AsgardeonUI
    }
}

# Individual function runners for direct function invocation
function Set-AsgardeonTheme {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("Light", "Dark")]
        [string]$Theme
    )
    
    $initSuccess = Initialize-Asgardeon
    if ($initSuccess) {
        Set-WindowsTheme -Theme $Theme
    }
}

function Set-AsgardeonWallpaper {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ImagePath
    )
    
    $initSuccess = Initialize-Asgardeon
    if ($initSuccess) {
        if (Test-Path $ImagePath) {
            Set-WallPaper -ImagePath $ImagePath
        } else {
            Write-Error "Image file not found at path: $ImagePath"
        }
    }
}

function Set-AsgardeonLockScreen {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ImagePath
    )
    
    $initSuccess = Initialize-Asgardeon
    if ($initSuccess) {
        if (Test-Path $ImagePath) {
            Set-LockScreenWallpaper -ImagePath $ImagePath
        } else {
            Write-Error "Image file not found at path: $ImagePath"
        }
    }
}

function Start-WindowsActivation {
    $initSuccess = Initialize-Asgardeon
    if ($initSuccess) {
        Activate-WindowsOffice
    }
}

function Enable-HideActivation {
    $initSuccess = Initialize-Asgardeon
    if ($initSuccess) {
        Hide-WindowsActivation
    }
}

# Export functions (this is important for remote invocation)
Export-ModuleMember -Function Start-Asgardeon, Set-AsgardeonTheme, Set-AsgardeonWallpaper, Set-AsgardeonLockScreen, Start-WindowsActivation, Enable-HideActivation

# Default action when script is invoked without parameters
Start-Asgardeon