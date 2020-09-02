$AppLocation = "$PSScriptRoot\gitswitch.bat" 
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\gitswitch.lnk")
$Shortcut.TargetPath = $AppLocation
$Shortcut.Description ="Git Switcher"
$Shortcut.Save()