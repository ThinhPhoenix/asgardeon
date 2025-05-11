function Hide-WindowsActivation {
    try {
        $startupFolder = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\Start Menu\Programs\Startup")
        $scriptPath = [System.IO.Path]::Combine($startupFolder, "wis_hideactivate.bat")
        
        $hideActivateForm = New-Object System.Windows.Forms.Form
        $hideActivateForm.Text = "Hide Windows Activation"
        $hideActivateForm.Size = New-Object System.Drawing.Size(400, 250)
        $hideActivateForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $hideActivateForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $hideActivateForm.MaximizeBox = $false
        $hideActivateForm.MinimizeBox = $false
        
        $descriptionLabel = New-Object System.Windows.Forms.Label
        $descriptionLabel.Text = "This will hide the Windows activation watermark`nby restarting the explorer process at startup."
        $descriptionLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $descriptionLabel.Size = New-Object System.Drawing.Size(380, 50)
        $descriptionLabel.Location = New-Object System.Drawing.Point(10, 20)
        $hideActivateForm.Controls.Add($descriptionLabel)
        
        $startButton = New-Object System.Windows.Forms.Button
        $startButton.Text = "Enable Hiding"
        $startButton.Location = New-Object System.Drawing.Point(50, 90)
        $startButton.Size = New-Object System.Drawing.Size(120, 40)
        
        $deleteButton = New-Object System.Windows.Forms.Button
        $deleteButton.Text = "Disable Hiding"
        $deleteButton.Location = New-Object System.Drawing.Point(220, 90)
        $deleteButton.Size = New-Object System.Drawing.Size(120, 40)
        
        $closeButton = New-Object System.Windows.Forms.Button
        $closeButton.Text = "Cancel"
        $closeButton.Location = New-Object System.Drawing.Point(150, 160)
        $closeButton.Size = New-Object System.Drawing.Size(100, 35)
        
        $startButton.Add_Click({
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
            [System.Windows.Forms.MessageBox]::Show("Done! Restart your PC to apply the change.`nScript created at: $scriptPath", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $statusLabel.Text = "Hiding Windows activation enabled!"
            $statusLabel.ForeColor = [System.Drawing.Color]::Green
            $hideActivateForm.Close()
        })
        
        $deleteButton.Add_Click({
            if (Test-Path -Path $scriptPath) {
                Remove-Item -Path $scriptPath
                [System.Windows.Forms.MessageBox]::Show("Hiding script has been removed from startup.`nChanges will take effect after restarting your PC.", "Script Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $statusLabel.Text = "Hiding Windows activation disabled"
                $statusLabel.ForeColor = [System.Drawing.Color]::Blue
            } else {
                [System.Windows.Forms.MessageBox]::Show("No activation hiding script found in Startup.", "No Script Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            $hideActivateForm.Close()
        })
        
        $closeButton.Add_Click({
            $hideActivateForm.Close()
        })
        
        $hideActivateForm.Controls.Add($startButton)
        $hideActivateForm.Controls.Add($deleteButton)
        $hideActivateForm.Controls.Add($closeButton)
        
        $hideActivateForm.ShowDialog() | Out-Null
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error with Windows activation hiding: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}