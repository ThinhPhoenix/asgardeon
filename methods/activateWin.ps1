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