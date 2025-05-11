Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Custom Button Style Function with requested color #6087CF
function Create-CustomButton {
    param (
        [string]$Text,
        [System.Drawing.Point]$Location,
        [System.Drawing.Size]$Size,
        [string]$ToolTip = ""
    )
    
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = $Location
    $button.Size = $Size
    
    # Enhanced styling with the requested color
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.BackColor = [System.Drawing.Color]::FromArgb(96, 135, 207) # #6087CF
    $button.ForeColor = [System.Drawing.Color]::White
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(66, 95, 153) # Darker border
    $button.FlatAppearance.BorderSize = 1
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    
    # Add hover effects
    $button.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(116, 155, 227) # Lighter blue hover
    })
    
    $button.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(96, 135, 207) # Back to #6087CF
    })
    
    # Add pressed effect
    $button.Add_MouseDown({
        $this.BackColor = [System.Drawing.Color]::FromArgb(56, 95, 167) # Darker blue when pressed
    })
    
    $button.Add_MouseUp({
        $this.BackColor = [System.Drawing.Color]::FromArgb(116, 155, 227) # Back to hover color
    })
    
    # Add tooltip if provided - FIXED HERE
    if ($ToolTip -ne "") {
        $tooltipControl = New-Object System.Windows.Forms.ToolTip
        $tooltipControl.SetToolTip($button, $ToolTip)
    }
    
    return $button
}

# Apply OneDark theme to form controls
function Apply-OneDarkTheme {
    param (
        [System.Windows.Forms.Form]$form
    )
    
    # OneDark color palette
    $darkBackground = [System.Drawing.Color]::FromArgb(40, 44, 52)      # #282c34
    $darkForeground = [System.Drawing.Color]::FromArgb(171, 178, 191)   # #abb2bf
    $darkSelection = [System.Drawing.Color]::FromArgb(62, 68, 81)       # #3e4451
    $darkComment = [System.Drawing.Color]::FromArgb(92, 99, 112)        # #5c6370
    $darkBlue = [System.Drawing.Color]::FromArgb(97, 175, 239)          # #61afef
    
    # Update form colors
    $form.BackColor = $darkBackground
    $form.ForeColor = $darkForeground
    
    # Update all panels
    foreach ($control in $form.Controls) {
        if ($control -is [System.Windows.Forms.Panel]) {
            $control.BackColor = $darkBackground
            $control.ForeColor = $darkForeground
            
            # Update all contained controls
            foreach ($childControl in $control.Controls) {
                if ($childControl -is [System.Windows.Forms.Panel]) {
                    $childControl.BackColor = $darkBackground
                    $childControl.ForeColor = $darkForeground
                    
                    # Update GroupBoxes
                    foreach ($groupBox in $childControl.Controls) {
                        if ($groupBox -is [System.Windows.Forms.GroupBox]) {
                            $groupBox.BackColor = $darkSelection
                            $groupBox.ForeColor = $darkForeground
                            
                            # Update controls within GroupBox
                            foreach ($gbControl in $groupBox.Controls) {
                                if ($gbControl -is [System.Windows.Forms.Label]) {
                                    $gbControl.ForeColor = $darkForeground
                                    $gbControl.BackColor = [System.Drawing.Color]::Transparent
                                }
                                elseif ($gbControl -is [System.Windows.Forms.ComboBox]) {
                                    $gbControl.BackColor = $darkBackground
                                    $gbControl.ForeColor = $darkForeground
                                }
                                elseif ($gbControl -is [System.Windows.Forms.CheckBox]) {
                                    $gbControl.ForeColor = $darkForeground
                                    $gbControl.BackColor = [System.Drawing.Color]::Transparent
                                }
                                # PictureBox gets a dark background
                                elseif ($gbControl -is [System.Windows.Forms.PictureBox]) {
                                    $gbControl.BackColor = $darkBackground
                                }
                            }
                        }
                    }
                }
            }
        }
        elseif ($control -is [System.Windows.Forms.StatusBar]) {
            $control.BackColor = $darkSelection
            $control.ForeColor = $darkForeground
        }
    }
    
    # Update status panel and label
    $statusPanel.BackColor = $darkSelection
    $statusLabel.ForeColor = $darkForeground
    $statusLabel.BackColor = [System.Drawing.Color]::Transparent
    
    # Help text has different color
    $hideActivateHelp.ForeColor = $darkComment
}

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
            $pictureBox.BackColor = [System.Drawing.Color]::FromArgb(40, 44, 52) # OneDark background
            return $false
        }
        
        $image = [System.Drawing.Image]::FromFile($imagePath)
        $pictureBox.Image = $image
        $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        return $true
    } catch {
        $pictureBox.Image = $null
        $pictureBox.BackColor = [System.Drawing.Color]::FromArgb(40, 44, 52) # OneDark background
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
            $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(152, 195, 121) # OneDark green
        }
        else {
            # Remove the script if it exists
            if (Test-Path -Path $scriptPath) {
                Remove-Item -Path $scriptPath
                [System.Windows.Forms.MessageBox]::Show("Activation hiding disabled. Changes will take effect after restarting your PC.", "Script Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $statusLabel.Text = "Hiding Windows activation disabled"
                $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(97, 175, 239) # OneDark blue
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
        $activateForm.BackColor = [System.Drawing.Color]::FromArgb(40, 44, 52) # OneDark background
        $activateForm.ForeColor = [System.Drawing.Color]::FromArgb(171, 178, 191) # OneDark foreground
        
        $warningLabel = New-Object System.Windows.Forms.Label
        $warningLabel.Text = "WARNING:`nThis will attempt to activate Windows and/or Office`nusing an online activation script.`n`nOnly proceed if you understand the implications."
        $warningLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $warningLabel.Size = New-Object System.Drawing.Size(380, 100)
        $warningLabel.Location = New-Object System.Drawing.Point(10, 20)
        $warningLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $warningLabel.ForeColor = [System.Drawing.Color]::FromArgb(224, 108, 117) # OneDark red
        $warningLabel.BackColor = [System.Drawing.Color]::Transparent
        $activateForm.Controls.Add($warningLabel)
        
        # Use custom styled buttons for activate and cancel
        $activateButton = Create-CustomButton -Text "Run Activation" `
                                             -Location (New-Object System.Drawing.Point(120, 130)) `
                                             -Size (New-Object System.Drawing.Size(150, 40)) `
                                             -ToolTip "Run Windows/Office activation script"
        
        $cancelButton = Create-CustomButton -Text "Cancel" `
                                          -Location (New-Object System.Drawing.Point(120, 190)) `
                                          -Size (New-Object System.Drawing.Size(150, 40))
        # Set cancel button to a different color scheme
        $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(92, 99, 112) # OneDark comment gray
        $cancelButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(62, 68, 81) # OneDark selection
        
        $cancelButton.Add_MouseEnter({
            $this.BackColor = [System.Drawing.Color]::FromArgb(112, 119, 132) # Lighter gray
        })
        
        $cancelButton.Add_MouseLeave({
            $this.BackColor = [System.Drawing.Color]::FromArgb(92, 99, 112) # OneDark comment gray
        })
        
        $activateButton.Add_Click({
            $statusLabel.Text = "Running activation script..."
            $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(97, 175, 239) # OneDark blue
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
                $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(152, 195, 121) # OneDark green
            }
            catch {
                $statusLabel.Text = "Activation failed: $_"
                $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(224, 108, 117) # OneDark red
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
$form.Size = New-Object System.Drawing.Size(465, 570)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.BackColor = [System.Drawing.Color]::FromArgb(40, 44, 52) # OneDark background

# Create a scrollable panel with only vertical scrolling
$scrollPanel = New-Object System.Windows.Forms.Panel
$scrollPanel.AutoScroll = $true
$scrollPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$scrollPanel.AutoScrollMinSize = New-Object System.Drawing.Size(0, 600) # Set only height for vertical scrolling
$scrollPanel.BackColor = [System.Drawing.Color]::FromArgb(40, 44, 52) # OneDark background
$form.Controls.Add($scrollPanel)

# Create a container panel within the scroll panel with fixed width
$containerPanel = New-Object System.Windows.Forms.Panel
$containerPanel.Width = 430
$containerPanel.Height = 600 # Taller to allow scrolling
$containerPanel.Location = New-Object System.Drawing.Point(10, 10)
$containerPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$containerPanel.BackColor = [System.Drawing.Color]::FromArgb(40, 44, 52) # OneDark background
$scrollPanel.Controls.Add($containerPanel)

# Theme section in a GroupBox
$themeGroup = New-Object System.Windows.Forms.GroupBox
$themeGroup.Text = "Theme"
$themeGroup.Size = New-Object System.Drawing.Size(410, 80)
$themeGroup.Location = New-Object System.Drawing.Point(10, 20)
$themeGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$themeGroup.BackColor = [System.Drawing.Color]::FromArgb(62, 68, 81) # OneDark selection
$themeGroup.ForeColor = [System.Drawing.Color]::FromArgb(171, 178, 191) # OneDark foreground
$containerPanel.Controls.Add($themeGroup)

# Theme controls inside group
$themeLabel = New-Object System.Windows.Forms.Label
$themeLabel.Text = "Select Theme:"
$themeLabel.Size = New-Object System.Drawing.Size(90, 25)
$themeLabel.Location = New-Object System.Drawing.Point(15, 35)
$themeLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$themeLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 178, 191) # OneDark foreground
$themeLabel.BackColor = [System.Drawing.Color]::Transparent
$themeGroup.Controls.Add($themeLabel)

$themeComboBox = New-Object System.Windows.Forms.ComboBox
$themeComboBox.Location = New-Object System.Drawing.Point(110, 35)
$themeComboBox.Size = New-Object System.Drawing.Size(180, 25)
$themeComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$themeComboBox.Items.Add("Light")
$themeComboBox.Items.Add("Dark")
$themeComboBox.BackColor = [System.Drawing.Color]::FromArgb(40, 44, 52) # OneDark background
$themeComboBox.ForeColor = [System.Drawing.Color]::FromArgb(171, 178, 191) # OneDark foreground
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

# Replace with custom styled button
$applyThemeButton = Create-CustomButton -Text "Apply" `
                                       -Location (New-Object System.Drawing.Point(310, 35)) `
                                       -Size (New-Object System.Drawing.Size(90, 25)) `
                                       -ToolTip "Apply the selected theme to Windows"
$themeGroup.Controls.Add($applyThemeButton)

$applyThemeButton.Add_Click({
    $selectedTheme = $themeComboBox.SelectedItem
    $statusLabel.Text = "Setting $selectedTheme theme and resetting taskbar..."
    $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(97, 175, 239) # OneDark blue
    $form.Refresh()
    
    $success = Set-WindowsTheme -Theme $selectedTheme
    if ($success) {
        $statusLabel.Text = "Theme changed to $selectedTheme successfully!"
        $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(152, 195, 121) # OneDark green
    } else {
        $statusLabel.Text = "Error changing theme"
        $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(224, 108, 117) # OneDark red
    }
})
$applyThemeButton.Location = New-Object System.Drawing.Point(300, 35)

# Wallpaper section in a GroupBox
$wallpaperGroup = New-Object System.Windows.Forms.GroupBox
$wallpaperGroup.Text = "Personalization"
$wallpaperGroup.Size = New-Object System.Drawing.Size(410, 160)
$wallpaperGroup.Location = New-Object System.Drawing.Point(10, 110)
$wallpaperGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$wallpaperGroup.BackColor = [System.Drawing.Color]::FromArgb(62, 68, 81) # OneDark selection
$wallpaperGroup.ForeColor = [System.Drawing.Color]::FromArgb(171, 178, 191) # OneDark foreground
$containerPanel.Controls.Add($wallpaperGroup)

# Desktop wallpaper controls
$wallpaperLabel = New-Object System.Windows.Forms.Label
$wallpaperLabel.Text = "Desktop Wallpaper:"
$wallpaperLabel.Size = New-Object System.Drawing.Size(120, 30)
$wallpaperLabel.Location = New-Object System.Drawing.Point(15, 30)
$wallpaperLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$wallpaperLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 178, 191) # OneDark foreground
$wallpaperLabel.BackColor = [System.Drawing.Color]::Transparent
$wallpaperGroup.Controls.Add($wallpaperLabel)

# Replace with custom styled button
$wallpaperButton = Create-CustomButton -Text "Browse..." `
                                       -Location (New-Object System.Drawing.Point(135, 30)) `
                                       -Size (New-Object System.Drawing.Size(170, 30)) `
                                       -ToolTip "Select a new desktop wallpaper image"
$wallpaperGroup.Controls.Add($wallpaperButton)

# Wallpaper preview
$wallpaperPreview = New-Object System.Windows.Forms.PictureBox
$wallpaperPreview.Location = New-Object System.Drawing.Point(315, 25)
$wallpaperPreview.Size = New-Object System.Drawing.Size(80, 45)
$wallpaperPreview.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$wallpaperPreview.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$wallpaperPreview.BackColor = [System.Drawing.Color]::FromArgb(40, 44, 52) # OneDark background
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
$lockScreenLabel.Location = New-Object System.Drawing.Point(15, 90)
$lockScreenLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lockScreenLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 178, 191) # OneDark foreground
$lockScreenLabel.BackColor = [System.Drawing.Color]::Transparent
$wallpaperGroup.Controls.Add($lockScreenLabel)

# Replace with custom styled button
$lockScreenButton = Create-CustomButton -Text "Browse..." `
                                       -Location (New-Object System.Drawing.Point(135, 90)) `
                                       -Size (New-Object System.Drawing.Size(170, 30)) `
                                       -ToolTip "Select a new lock screen wallpaper image"
$wallpaperGroup.Controls.Add($lockScreenButton)

# Lock screen preview
$lockScreenPreview = New-Object System.Windows.Forms.PictureBox
$lockScreenPreview.Location = New-Object System.Drawing.Point(315, 85)
$lockScreenPreview.Size = New-Object System.Drawing.Size(80, 45)
$lockScreenPreview.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$lockScreenPreview.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$lockScreenPreview.BackColor = [System.Drawing.Color]::FromArgb(40, 44, 52) # OneDark background
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
$activationGroup.Size = New-Object System.Drawing.Size(410, 130)
$activationGroup.Location = New-Object System.Drawing.Point(10, 280)
$activationGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$activationGroup.BackColor = [System.Drawing.Color]::FromArgb(62, 68, 81) # OneDark selection
$activationGroup.ForeColor = [System.Drawing.Color]::FromArgb(171, 178, 191) # OneDark foreground
$containerPanel.Controls.Add($activationGroup)

# Hide activation switch & label
$hideActivateLabel = New-Object System.Windows.Forms.Label
$hideActivateLabel.Text = "Hide Activation Watermark:"
$hideActivateLabel.Size = New-Object System.Drawing.Size(160, 25)
$hideActivateLabel.Location = New-Object System.Drawing.Point(15, 30)
$hideActivateLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$hideActivateLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 178, 191) # OneDark foreground
$hideActivateLabel.BackColor = [System.Drawing.Color]::Transparent
$activationGroup.Controls.Add($hideActivateLabel)

$hideActivateSwitch = New-Object System.Windows.Forms.CheckBox
$hideActivateSwitch.Text = "Enabled"
$hideActivateSwitch.Location = New-Object System.Drawing.Point(185, 30)
$hideActivateSwitch.Size = New-Object System.Drawing.Size(70, 25)
$hideActivateSwitch.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$hideActivateSwitch.ForeColor = [System.Drawing.Color]::FromArgb(171, 178, 191) # OneDark foreground
$hideActivateSwitch.BackColor = [System.Drawing.Color]::Transparent
$activationGroup.Controls.Add($hideActivateSwitch)

# Small help label
$hideActivateHelp = New-Object System.Windows.Forms.Label
$hideActivateHelp.Text = "(Requires restart)"
$hideActivateHelp.Size = New-Object System.Drawing.Size(130, 25)
$hideActivateHelp.Location = New-Object System.Drawing.Point(260, 30)
$hideActivateHelp.Font = New-Object System.Drawing.Font("Arial", 8)
$hideActivateHelp.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$hideActivateHelp.ForeColor = [System.Drawing.Color]::FromArgb(92, 99, 112) # OneDark comment
$hideActivateHelp.BackColor = [System.Drawing.Color]::Transparent
$activationGroup.Controls.Add($hideActivateHelp)

# Custom style for activation button
$activateButton = Create-CustomButton -Text "Activate Windows/Office" `
                                     -Location (New-Object System.Drawing.Point(110, 75)) `
                                     -Size (New-Object System.Drawing.Size(200, 35)) `
                                     -ToolTip "Run Windows and Office activation script"

# Set a different color for this special button
$activateButton.BackColor = [System.Drawing.Color]::FromArgb(224, 108, 117) # OneDark red for emphasis
$activateButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(190, 80, 90) # Darker border

$activateButton.Add_MouseEnter({
    $this.BackColor = [System.Drawing.Color]::FromArgb(235, 130, 140) # Lighter red
})

$activateButton.Add_MouseLeave({
    $this.BackColor = [System.Drawing.Color]::FromArgb(224, 108, 117) # OneDark red
})

$activationGroup.Controls.Add($activateButton)

# Check if hiding script exists and set checkbox state accordingly
$startupFolder = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\Start Menu\Programs\Startup")
$scriptPath = [System.IO.Path]::Combine($startupFolder, "wis_hideactivate.bat")
$hideActivateSwitch.Checked = (Test-Path -Path $scriptPath)

# Add event handler for the checkbox
$hideActivateSwitch.Add_CheckedChanged({
    $statusLabel.Text = "Updating activation hiding settings..."
    $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(97, 175, 239) # OneDark blue
    $form.Refresh()
    
    $success = Set-ActivationHiding -Enable $hideActivateSwitch.Checked
    if (-not $success) {
        $statusLabel.Text = "Error changing activation hiding setting"
        $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(224, 108, 117) # OneDark red
    }
})

$activateButton.Add_Click({
    $statusLabel.Text = "Preparing activation options..."
    $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(97, 175, 239) # OneDark blue
    $form.Refresh()
    
    $success = Activate-WindowsOffice
    if (-not $success) {
        $statusLabel.Text = "Error with activation function"
        $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(224, 108, 117) # OneDark red
    }
})

# Status label at the bottom of the form (outside the scroll panel)
$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$statusPanel.Height = 30
$statusPanel.BackColor = [System.Drawing.Color]::FromArgb(62, 68, 81) # OneDark selection
$form.Controls.Add($statusPanel)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready"
$statusLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 178, 191) # OneDark foreground
$statusLabel.BackColor = [System.Drawing.Color]::Transparent
$statusPanel.Controls.Add($statusLabel)

# Ensure scroll panel is above status panel
$form.Controls.SetChildIndex($statusPanel, 1)
$form.Controls.SetChildIndex($scrollPanel, 0)

# Set a keyboard shortcut to close the form (Escape key)
$form.KeyPreview = $true
$form.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
        $form.Close()
    }
})

# Apply OneDark theme
Apply-OneDarkTheme -form $form

$form.ShowDialog() | Out-Null