# Asgardeon Loader Script
# Version: 1.1

# Load required .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Base URL for fetching modules
$baseUrl = "https://raw.githubusercontent.com/thinhphoenix/asgardeon/main/methods"

# First, define all required functions
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
        Write-Host "Failed to import module $ModuleName from $moduleUrl. Error: $_" -ForegroundColor Red
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
            Write-Host "Failed to load module: $module" -ForegroundColor Red
            break
        } else {
            Write-Host "Successfully loaded module: $module" -ForegroundColor Green
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
        # Verify function exists before calling
        if (Get-Command -Name Show-AsgardeonUI -ErrorAction SilentlyContinue) {
            # Start the main UI
            Show-AsgardeonUI
        } else {
            [System.Windows.Forms.MessageBox]::Show("UI module was not properly loaded. Please check your internet connection and try again.", "Module Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

function Set-AsgardeonTheme {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("Light", "Dark")]
        [string]$Theme
    )
    
    $initSuccess = Initialize-Asgardeon
    if ($initSuccess -and (Get-Command -Name Set-WindowsTheme -ErrorAction SilentlyContinue)) {
        Set-WindowsTheme -Theme $Theme
    } else {
        Write-Host "Failed to initialize theme module" -ForegroundColor Red
    }
}

function Set-AsgardeonWallpaper {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ImagePath
    )
    
    $initSuccess = Initialize-Asgardeon
    if ($initSuccess -and (Get-Command -Name Set-WallPaper -ErrorAction SilentlyContinue)) {
        if (Test-Path $ImagePath) {
            Set-WallPaper -ImagePath $ImagePath
        } else {
            Write-Host "Image file not found at path: $ImagePath" -ForegroundColor Red
        }
    } else {
        Write-Host "Failed to initialize wallpaper module" -ForegroundColor Red
    }
}

function Set-AsgardeonLockScreen {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ImagePath
    )
    
    $initSuccess = Initialize-Asgardeon
    if ($initSuccess -and (Get-Command -Name Set-LockScreenWallpaper -ErrorAction SilentlyContinue)) {
        if (Test-Path $ImagePath) {
            Set-LockScreenWallpaper -ImagePath $ImagePath
        } else {
            Write-Host "Image file not found at path: $ImagePath" -ForegroundColor Red
        }
    } else {
        Write-Host "Failed to initialize lock screen module" -ForegroundColor Red
    }
}

function Start-WindowsActivation {
    $initSuccess = Initialize-Asgardeon
    if ($initSuccess -and (Get-Command -Name Activate-WindowsOffice -ErrorAction SilentlyContinue)) {
        Activate-WindowsOffice
    } else {
        Write-Host "Failed to initialize activation module" -ForegroundColor Red
    }
}

function Enable-HideActivation {
    $initSuccess = Initialize-Asgardeon
    if ($initSuccess -and (Get-Command -Name Hide-WindowsActivation -ErrorAction SilentlyContinue)) {
        Hide-WindowsActivation
    } else {
        Write-Host "Failed to initialize hide activation module" -ForegroundColor Red
    }
}

# Default action when script is invoked without parameters
# First check if any parameters were specified
$params = $args
if ($params.Count -eq 0) {
    # No parameters, start the UI
    Start-Asgardeon
}