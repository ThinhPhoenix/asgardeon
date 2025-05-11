function Set-LockScreenWallpaper {
    param (
        [string]$ImagePath
    )
    
    try {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
        
        if (!(Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        
        New-ItemProperty -Path $regPath -Name "LockScreenImagePath" -Value $ImagePath -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $regPath -Name "LockScreenImageStatus" -Value 1 -PropertyType DWORD -Force | Out-Null
        
        $destinationPath = "$env:WINDIR\System32\oobe\info\backgrounds"
        if (!(Test-Path $destinationPath)) {
            New-Item -Path $destinationPath -ItemType Directory -Force | Out-Null
        }
        
        Copy-Item -Path $ImagePath -Destination "$destinationPath\backgroundDefault.jpg" -Force
        
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error setting lock screen wallpaper: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}