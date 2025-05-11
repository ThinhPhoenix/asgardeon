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