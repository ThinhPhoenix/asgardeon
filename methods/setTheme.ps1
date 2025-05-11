function Set-WindowsTheme {
    param (
        [string]$Theme
    )
    
    try {
        if ($Theme -eq "Dark") {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 0
            
            Reset-Taskbar
            return $true
        }
        elseif ($Theme -eq "Light") {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 1
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 1
            
            Reset-Taskbar
            return $true
        }
        else {
            return $false
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error setting theme: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}