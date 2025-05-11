function Show-AsgardeonUI {
    param(
        [System.Windows.Forms.Label]$StatusLabel
    )
    
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Asgardeon"
    $form.Size = New-Object System.Drawing.Size(400, 450)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # Create title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Tweaks"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.Size = New-Object System.Drawing.Size(380, 30)
    $titleLabel.Location = New-Object System.Drawing.Point(10, 15)
    $form.Controls.Add($titleLabel)

    # Create status label
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Ready"
    $statusLabel.Size = New-Object System.Drawing.Size(380, 20)
    $statusLabel.Location = New-Object System.Drawing.Point(10, 385)
    $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($statusLabel)
    
    # Create all buttons
    $themeButton = New-Object System.Windows.Forms.Button
    $themeButton.Text = "Change Theme (Light/Dark)"
    $themeButton.Location = New-Object System.Drawing.Point(100, 70)
    $themeButton.Size = New-Object System.Drawing.Size(200, 40)
    $form.Controls.Add($themeButton)
    
    $wallpaperButton = New-Object System.Windows.Forms.Button
    $wallpaperButton.Text = "Set Desktop Wallpaper"
    $wallpaperButton.Location = New-Object System.Drawing.Point(100, 120)
    $wallpaperButton.Size = New-Object System.Drawing.Size(200, 40)
    $form.Controls.Add($wallpaperButton)
    
    $lockScreenButton = New-Object System.Windows.Forms.Button
    $lockScreenButton.Text = "Set Lock Screen Wallpaper"
    $lockScreenButton.Location = New-Object System.Drawing.Point(100, 170)
    $lockScreenButton.Size = New-Object System.Drawing.Size(200, 40)
    $form.Controls.Add($lockScreenButton)
    
    $hideActivateButton = New-Object System.Windows.Forms.Button
    $hideActivateButton.Text = "Hide Windows Activation"
    $hideActivateButton.Location = New-Object System.Drawing.Point(100, 220)
    $hideActivateButton.Size = New-Object System.Drawing.Size(200, 40)
    $form.Controls.Add($hideActivateButton)
    
    $activateButton = New-Object System.Windows.Forms.Button
    $activateButton.Text = "Activate Windows/Office"
    $activateButton.Location = New-Object System.Drawing.Point(100, 270)
    $activateButton.Size = New-Object System.Drawing.Size(200, 40)
    $form.Controls.Add($activateButton)
    
    $quitButton = New-Object System.Windows.Forms.Button
    $quitButton.Text = "Quit"
    $quitButton.Location = New-Object System.Drawing.Point(100, 330)
    $quitButton.Size = New-Object System.Drawing.Size(200, 40)
    $form.Controls.Add($quitButton)
    
    # Add button click events
    $themeButton.Add_Click({
        $themeForm = New-Object System.Windows.Forms.Form
        $themeForm.Text = "Theme Selection"
        $themeForm.Size = New-Object System.Drawing.Size(300, 200)
        $themeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $themeForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $themeForm.MaximizeBox = $false
        $themeForm.MinimizeBox = $false
        
        $themeLabel = New-Object System.Windows.Forms.Label
        $themeLabel.Text = "Select a system theme:"
        $themeLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $themeLabel.Size = New-Object System.Drawing.Size(280, 25)
        $themeLabel.Location = New-Object System.Drawing.Point(10, 20)
        $themeForm.Controls.Add($themeLabel)
        
        $lightButton = New-Object System.Windows.Forms.Button
        $lightButton.Text = "Light"
        $lightButton.Location = New-Object System.Drawing.Point(50, 60)
        $lightButton.Size = New-Object System.Drawing.Size(80, 35)
        
        $darkButton = New-Object System.Windows.Forms.Button
        $darkButton.Text = "Dark"
        $darkButton.Location = New-Object System.Drawing.Point(150, 60)
        $darkButton.Size = New-Object System.Drawing.Size(80, 35)
        
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(100, 120)
        $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
        
        $lightButton.Add_Click({
            $statusLabel.Text = "Setting Light theme and resetting taskbar..."
            $statusLabel.ForeColor = [System.Drawing.Color]::Blue
            $form.Refresh()
            
            $success = Set-WindowsTheme -Theme "Light"
            if ($success) {
                $statusLabel.Text = "Theme changed to Light successfully!"
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
                [System.Windows.Forms.MessageBox]::Show("Successfully set theme to Light and reset taskbar", "Theme Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                $statusLabel.Text = "Error changing theme"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
            }
            $themeForm.Close()
        })
        
        $darkButton.Add_Click({
            $statusLabel.Text = "Setting Dark theme and resetting taskbar..."
            $statusLabel.ForeColor = [System.Drawing.Color]::Blue
            $form.Refresh()
            
            $success = Set-WindowsTheme -Theme "Dark"
            if ($success) {
                $statusLabel.Text = "Theme changed to Dark successfully!"
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
                [System.Windows.Forms.MessageBox]::Show("Successfully set theme to Dark and reset taskbar", "Theme Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                $statusLabel.Text = "Error changing theme"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
            }
            $themeForm.Close()
        })
        
        $cancelButton.Add_Click({
            $themeForm.Close()
        })
        
        $themeForm.Controls.Add($lightButton)
        $themeForm.Controls.Add($darkButton)
        $themeForm.Controls.Add($cancelButton)
        
        $themeForm.ShowDialog() | Out-Null
    })
    
    $wallpaperButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Image Files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp"
        $openFileDialog.Title = "Select Desktop Wallpaper Image"
        
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $wallpaperPath = $openFileDialog.FileName
            $success = Set-WallPaper -ImagePath $wallpaperPath
            
            if ($success) {
                [System.Windows.Forms.MessageBox]::Show("Desktop wallpaper set successfully!", "Wallpaper Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Failed to set desktop wallpaper.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    
    $lockScreenButton.Add_Click({
        [System.Windows.Forms.MessageBox]::Show("Note: Setting lock screen wallpaper requires Administrator privileges. The script will attempt to set it, but may fail without proper permissions.", "Administrator Rights Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Image Files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp"
        $openFileDialog.Title = "Select Lock Screen Wallpaper Image"
        
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $lockWallpaperPath = $openFileDialog.FileName
            $success = Set-LockScreenWallpaper -ImagePath $lockWallpaperPath
            
            if ($success) {
                [System.Windows.Forms.MessageBox]::Show("Lock screen wallpaper set successfully!", "Lock Screen Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Failed to set lock screen wallpaper. Try running the script as Administrator.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    
    $hideActivateButton.Add_Click({
        $statusLabel.Text = "Managing Windows activation notification..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()
        
        $success = Hide-WindowsActivation
        if (-not $success) {
            $statusLabel.Text = "Error with hide activation function"
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
    
    $quitButton.Add_Click({
        $form.Close()
    })
    
    $form.ShowDialog() | Out-Null
}