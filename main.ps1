Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Reset-Taskbar {
    try {
        $explorer = Get-Process -Name explorer -ErrorAction SilentlyContinue
        
        if ($explorer) {
            $explorer | Stop-Process -Force
            Start-Sleep -Seconds 1
            Start-Process explorer
            return $true
        }
        else {
            Start-Process explorer
            return $true
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error resetting taskbar: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

function Set-WallPaper {
    param (
        [string]$ImagePath
    )
    
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    
    public class Wallpaper {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
    }
"@
    
    $SPI_SETDESKWALLPAPER = 0x0014
    $SPIF_UPDATEINIFILE = 0x01
    $SPIF_SENDCHANGE = 0x02
    
    $ret = [Wallpaper]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $ImagePath, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)
    return $ret
}

# Function to get current wallpaper path
function Get-CurrentWallpaper {
    try {
        $wallpaperPath = (Get-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name Wallpaper).Wallpaper
        return $wallpaperPath
    } catch {
        return $null
    }
}

# Function to get current lock screen wallpaper
function Get-CurrentLockScreenWallpaper {
    try {
        # Try to get from registry (may not work depending on Windows version and permissions)
        $lockScreenPath = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP" -Name "LockScreenImagePath" -ErrorAction SilentlyContinue).LockScreenImagePath
        
        # If that doesn't work, try the default location
        if (!$lockScreenPath -or !(Test-Path $lockScreenPath)) {
            $defaultPath = "$env:WINDIR\System32\oobe\info\backgrounds\backgroundDefault.jpg"
            if (Test-Path $defaultPath) {
                return $defaultPath
            }
        }
        
        return $lockScreenPath
    } catch {
        return $null
    }
}

# Load image safely with error handling
function Load-ImageSafely {
    param (
        [System.Windows.Forms.PictureBox]$pictureBox,
        [string]$imagePath
    )
    
    try {
        if ([string]::IsNullOrEmpty($imagePath) -or -not (Test-Path $imagePath)) {
            $pictureBox.Image = $null
            $pictureBox.BackColor = [System.Drawing.Color]::LightGray
            return $false
        }
        
        $image = [System.Drawing.Image]::FromFile($imagePath)
        $pictureBox.Image = $image
        $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        return $true
    } catch {
        $pictureBox.Image = $null
        $pictureBox.BackColor = [System.Drawing.Color]::LightGray
        return $false
    }
}

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

function Set-ActivationHiding {
    param (
        [bool]$Enable
    )
    
    try {
        $startupFolder = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\Start Menu\Programs\Startup")
        $scriptPath = [System.IO.Path]::Combine($startupFolder, "wis_hideactivate.bat")
        
        if ($Enable) {
            # Create the script
            $batchContent = @"
@echo off
echo Hide Windows Activation - $(Get-Date)
timeout /t 2 /nobreak
taskkill /F /IM explorer.exe
timeout /t 1 /nobreak
start explorer.exe
exit
"@
            Set-Content -Path $scriptPath -Value $batchContent
            [System.Windows.Forms.MessageBox]::Show("Activation hiding enabled! Restart your PC to apply the change.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $statusLabel.Text = "Hiding Windows activation enabled!"
            $statusLabel.ForeColor = [System.Drawing.Color]::Green
        }
        else {
            # Remove the script if it exists
            if (Test-Path -Path $scriptPath) {
                Remove-Item -Path $scriptPath
                [System.Windows.Forms.MessageBox]::Show("Activation hiding disabled. Changes will take effect after restarting your PC.", "Script Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $statusLabel.Text = "Hiding Windows activation disabled"
                $statusLabel.ForeColor = [System.Drawing.Color]::Blue
            }
        }
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error with Windows activation hiding: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

function Activate-WindowsOffice {
    try {
        $activateForm = New-Object System.Windows.Forms.Form
        $activateForm.Text = "Activate Windows/Office"
        $activateForm.Size = New-Object System.Drawing.Size(400, 300)
        $activateForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $activateForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $activateForm.MaximizeBox = $false
        $activateForm.MinimizeBox = $false
        
        $warningLabel = New-Object System.Windows.Forms.Label
        $warningLabel.Text = "WARNING:`nThis will attempt to activate Windows and/or Office`nusing an online activation script.`n`nOnly proceed if you understand the implications."
        $warningLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $warningLabel.Size = New-Object System.Drawing.Size(380, 100)
        $warningLabel.Location = New-Object System.Drawing.Point(10, 20)
        $warningLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $warningLabel.ForeColor = [System.Drawing.Color]::Red
        $activateForm.Controls.Add($warningLabel)
        
        $activateButton = New-Object System.Windows.Forms.Button
        $activateButton.Text = "Run Activation"
        $activateButton.Location = New-Object System.Drawing.Point(120, 130)
        $activateButton.Size = New-Object System.Drawing.Size(150, 40)
        
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(120, 190)
        $cancelButton.Size = New-Object System.Drawing.Size(150, 40)
        
        $activateButton.Add_Click({
            $statusLabel.Text = "Running activation script..."
            $statusLabel.ForeColor = [System.Drawing.Color]::Blue
            $form.Refresh()
            
            try {
                $tempScriptPath = [System.IO.Path]::Combine($env:TEMP, "activate_script.ps1")
                Set-Content -Path $tempScriptPath -Value 'irm https://get.activated.win | iex'
                
                $startInfo = New-Object System.Diagnostics.ProcessStartInfo
                $startInfo.FileName = "powershell.exe"
                $startInfo.Arguments = "-ExecutionPolicy Bypass -File `"$tempScriptPath`""
                $startInfo.Verb = "runas"                
                [System.Diagnostics.Process]::Start($startInfo)
                
                $statusLabel.Text = "Activation script launched!"
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
            }
            catch {
                $statusLabel.Text = "Activation failed: $_"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
                [System.Windows.Forms.MessageBox]::Show("Failed to run activation script: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            
            $activateForm.Close()
        })
        
        $cancelButton.Add_Click({
            $activateForm.Close()
        })
        
        $activateForm.Controls.Add($activateButton)
        $activateForm.Controls.Add($cancelButton)
        
        $activateForm.ShowDialog() | Out-Null
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error with activation process: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

# Create improved UI with better layout and organization
$form = New-Object System.Windows.Forms.Form
$form.Text = "Asgardeon"
$form.Size = New-Object System.Drawing.Size(470, 550)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.BackColor = [System.Drawing.Color]::WhiteSmoke

# Create a scrollable panel with only vertical scrolling
$scrollPanel = New-Object System.Windows.Forms.Panel
$scrollPanel.AutoScroll = $true
$scrollPanel.AutoScrollMinSize = New-Object System.Drawing.Size(430, 650) # Set minimum content width to prevent horizontal scrolling
$scrollPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.Controls.Add($scrollPanel)

# Create a container panel within the scroll panel
$containerPanel = New-Object System.Windows.Forms.Panel
$containerPanel.Size = New-Object System.Drawing.Size(430, 650) # Taller to allow scrolling
$containerPanel.Location = New-Object System.Drawing.Point(10, 10)
$scrollPanel.Controls.Add($containerPanel)

# Create a stylish text logo instead of the image
$logoLabel = New-Object System.Windows.Forms.Label
$logoLabel.Text = "Asgardeon"
$logoLabel.Size = New-Object System.Drawing.Size(410, 80)
$logoLabel.Location = New-Object System.Drawing.Point(10, 20)
$logoLabel.Font = New-Object System.Drawing.Font("Arial", 28, [System.Drawing.FontStyle]::Bold)
$logoLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 255) # Blue color
$logoLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$containerPanel.Controls.Add($logoLabel)

# Theme section in a GroupBox
$themeGroup = New-Object System.Windows.Forms.GroupBox
$themeGroup.Text = "Theme"
$themeGroup.Size = New-Object System.Drawing.Size(410, 70)
$themeGroup.Location = New-Object System.Drawing.Point(10, 120)
$containerPanel.Controls.Add($themeGroup)

# Theme controls inside group
$themeLabel = New-Object System.Windows.Forms.Label
$themeLabel.Text = "Select Theme:"
$themeLabel.Size = New-Object System.Drawing.Size(90, 25)
$themeLabel.Location = New-Object System.Drawing.Point(15, 30)
$themeLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$themeGroup.Controls.Add($themeLabel)

$themeComboBox = New-Object System.Windows.Forms.ComboBox
$themeComboBox.Location = New-Object System.Drawing.Point(110, 30)
$themeComboBox.Size = New-Object System.Drawing.Size(150, 25)
$themeComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$themeComboBox.Items.Add("Light")
$themeComboBox.Items.Add("Dark")
$themeGroup.Controls.Add($themeComboBox)

# Set current theme value
try {
    $currentTheme = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue
    if ($currentTheme -and $currentTheme.AppsUseLightTheme -eq 0) {
        $themeComboBox.SelectedIndex = 1 # Dark
    } else {
        $themeComboBox.SelectedIndex = 0 # Light
    }
} catch {
    $themeComboBox.SelectedIndex = 0
}

$applyThemeButton = New-Object System.Windows.Forms.Button
$applyThemeButton.Text = "Apply"
$applyThemeButton.Location = New-Object System.Drawing.Point(280, 30)
$applyThemeButton.Size = New-Object System.Drawing.Size(110, 25)
$themeGroup.Controls.Add($applyThemeButton)

$applyThemeButton.Add_Click({
    $selectedTheme = $themeComboBox.SelectedItem
    $statusLabel.Text = "Setting $selectedTheme theme and resetting taskbar..."
    $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    $form.Refresh()
    
    $success = Set-WindowsTheme -Theme $selectedTheme
    if ($success) {
        $statusLabel.Text = "Theme changed to $selectedTheme successfully!"
        $statusLabel.ForeColor = [System.Drawing.Color]::Green
    } else {
        $statusLabel.Text = "Error changing theme"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
    }
})

# Wallpaper section in a GroupBox
$wallpaperGroup = New-Object System.Windows.Forms.GroupBox
$wallpaperGroup.Text = "Personalization"
$wallpaperGroup.Size = New-Object System.Drawing.Size(410, 150)
$wallpaperGroup.Location = New-Object System.Drawing.Point(10, 200)
$containerPanel.Controls.Add($wallpaperGroup)

# Desktop wallpaper controls
$wallpaperLabel = New-Object System.Windows.Forms.Label
$wallpaperLabel.Text = "Desktop Wallpaper:"
$wallpaperLabel.Size = New-Object System.Drawing.Size(120, 30)
$wallpaperLabel.Location = New-Object System.Drawing.Point(15, 30)
$wallpaperLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$wallpaperGroup.Controls.Add($wallpaperLabel)

$wallpaperButton = New-Object System.Windows.Forms.Button
$wallpaperButton.Text = "Browse..."
$wallpaperButton.Location = New-Object System.Drawing.Point(135, 30)
$wallpaperButton.Size = New-Object System.Drawing.Size(170, 30)
$wallpaperGroup.Controls.Add($wallpaperButton)

# Wallpaper preview
$wallpaperPreview = New-Object System.Windows.Forms.PictureBox
$wallpaperPreview.Location = New-Object System.Drawing.Point(320, 25)
$wallpaperPreview.Size = New-Object System.Drawing.Size(70, 40)
$wallpaperPreview.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$wallpaperPreview.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$wallpaperPreview.BackColor = [System.Drawing.Color]::LightGray
$wallpaperGroup.Controls.Add($wallpaperPreview)

# Load current wallpaper
$currentWallpaperPath = Get-CurrentWallpaper
Load-ImageSafely -pictureBox $wallpaperPreview -imagePath $currentWallpaperPath

$wallpaperButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Image Files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp"
    $openFileDialog.Title = "Select Desktop Wallpaper Image"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $wallpaperPath = $openFileDialog.FileName
        $success = Set-WallPaper -ImagePath $wallpaperPath
        
        if ($success) {
            # Update the preview
            Load-ImageSafely -pictureBox $wallpaperPreview -imagePath $wallpaperPath
            [System.Windows.Forms.MessageBox]::Show("Desktop wallpaper set successfully!", "Wallpaper Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Failed to set desktop wallpaper.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Lock screen controls
$lockScreenLabel = New-Object System.Windows.Forms.Label
$lockScreenLabel.Text = "Lock Screen:"
$lockScreenLabel.Size = New-Object System.Drawing.Size(120, 30)
$lockScreenLabel.Location = New-Object System.Drawing.Point(15, 85)
$lockScreenLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$wallpaperGroup.Controls.Add($lockScreenLabel)

$lockScreenButton = New-Object System.Windows.Forms.Button
$lockScreenButton.Text = "Browse..."
$lockScreenButton.Location = New-Object System.Drawing.Point(135, 85)
$lockScreenButton.Size = New-Object System.Drawing.Size(170, 30)
$wallpaperGroup.Controls.Add($lockScreenButton)

# Lock screen preview
$lockScreenPreview = New-Object System.Windows.Forms.PictureBox
$lockScreenPreview.Location = New-Object System.Drawing.Point(320, 80)
$lockScreenPreview.Size = New-Object System.Drawing.Size(70, 40)
$lockScreenPreview.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$lockScreenPreview.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$lockScreenPreview.BackColor = [System.Drawing.Color]::LightGray
$wallpaperGroup.Controls.Add($lockScreenPreview)

# Load current lock screen
$currentLockScreenPath = Get-CurrentLockScreenWallpaper
Load-ImageSafely -pictureBox $lockScreenPreview -imagePath $currentLockScreenPath

$lockScreenButton.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Note: Setting lock screen wallpaper requires Administrator privileges. The script will attempt to set it, but may fail without proper permissions.", "Administrator Rights Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Image Files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp"
    $openFileDialog.Title = "Select Lock Screen Wallpaper Image"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $lockWallpaperPath = $openFileDialog.FileName
        $success = Set-LockScreenWallpaper -ImagePath $lockWallpaperPath
        
        if ($success) {
            # Update the preview
            Load-ImageSafely -pictureBox $lockScreenPreview -imagePath $lockWallpaperPath
            [System.Windows.Forms.MessageBox]::Show("Lock screen wallpaper set successfully!", "Lock Screen Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Failed to set lock screen wallpaper. Try running the script as Administrator.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Activation section in a GroupBox
$activationGroup = New-Object System.Windows.Forms.GroupBox
$activationGroup.Text = "Windows Activation"
$activationGroup.Size = New-Object System.Drawing.Size(410, 120)
$activationGroup.Location = New-Object System.Drawing.Point(10, 360)
$containerPanel.Controls.Add($activationGroup)

# Hide activation switch & label
$hideActivateLabel = New-Object System.Windows.Forms.Label
$hideActivateLabel.Text = "Hide Activation Watermark:"
$hideActivateLabel.Size = New-Object System.Drawing.Size(160, 25)
$hideActivateLabel.Location = New-Object System.Drawing.Point(15, 30)
$hideActivateLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$activationGroup.Controls.Add($hideActivateLabel)

$hideActivateSwitch = New-Object System.Windows.Forms.CheckBox
$hideActivateSwitch.Text = "Enabled"
$hideActivateSwitch.Location = New-Object System.Drawing.Point(185, 30)
$hideActivateSwitch.Size = New-Object System.Drawing.Size(70, 25)
$hideActivateSwitch.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$activationGroup.Controls.Add($hideActivateSwitch)

# Small help label
$hideActivateHelp = New-Object System.Windows.Forms.Label
$hideActivateHelp.Text = "(Requires restart)"
$hideActivateHelp.Size = New-Object System.Drawing.Size(130, 25)
$hideActivateHelp.Location = New-Object System.Drawing.Point(260, 30)
$hideActivateHelp.Font = New-Object System.Drawing.Font("Arial", 8)
$hideActivateHelp.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$hideActivateHelp.ForeColor = [System.Drawing.Color]::Gray
$activationGroup.Controls.Add($hideActivateHelp)

# Activate Windows/Office button
$activateButton = New-Object System.Windows.Forms.Button
$activateButton.Text = "Activate Windows/Office"
$activateButton.Location = New-Object System.Drawing.Point(105, 70)
$activateButton.Size = New-Object System.Drawing.Size(200, 35)
$activationGroup.Controls.Add($activateButton)

# Check if hiding script exists and set checkbox state accordingly
$startupFolder = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\Start Menu\Programs\Startup")
$scriptPath = [System.IO.Path]::Combine($startupFolder, "wis_hideactivate.bat")
$hideActivateSwitch.Checked = (Test-Path -Path $scriptPath)

# Add event handler for the checkbox
$hideActivateSwitch.Add_CheckedChanged({
    $statusLabel.Text = "Updating activation hiding settings..."
    $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    $form.Refresh()
    
    $success = Set-ActivationHiding -Enable $hideActivateSwitch.Checked
    if (-not $success) {
        $statusLabel.Text = "Error changing activation hiding setting"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
    }
})

$activateButton.Add_Click({
    $statusLabel.Text = "Preparing activation options..."
    $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    $form.Refresh()
    
    $success = Activate-WindowsOffice
    if (-not $success) {
        $statusLabel.Text = "Error with activation function"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
    }
})

# Exit button
$quitButton = New-Object System.Windows.Forms.Button
$quitButton.Text = "Exit"
$quitButton.Location = New-Object System.Drawing.Point(155, 490)
$quitButton.Size = New-Object System.Drawing.Size(120, 35)
$containerPanel.Controls.Add($quitButton)

$quitButton.Add_Click({
    $form.Close()
})

# Status label at the bottom of the form (outside the scroll panel)
$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$statusPanel.Height = 30
$statusPanel.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($statusPanel)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready"
$statusLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$statusPanel.Controls.Add($statusLabel)

# Ensure scroll panel is above status panel
$form.Controls.SetChildIndex($statusPanel, 1)
$form.Controls.SetChildIndex($scrollPanel, 0)

$form.ShowDialog() | Out-Null