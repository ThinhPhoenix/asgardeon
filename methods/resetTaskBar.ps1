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