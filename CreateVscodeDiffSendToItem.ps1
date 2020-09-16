$ExePath = [System.IO.Path]::GetFullPath("$env:AppData\..\Local\Programs\Microsoft VS Code\Code.exe")
$ShortcutLocation = [System.IO.Path]::GetFullPath("$env:AppData\Microsoft\Windows\SendTo\VS Code Diff Tool.lnk")
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutLocation)
$Shortcut.TargetPath = $ExePath
$Shortcut.Arguments = "--new-window --disable-extensions --diff"
$Shortcut.Save()
Write-Output "Created compare action $ShortcutLocation"
