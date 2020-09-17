$ExeInstallationRelativePath = "Microsoft VS Code\Code.exe"

if ((Test-Path -Path "${env:AppData}\..\Local\Programs\$ExeInstallationRelativePath")){
$ExePath = Resolve-Path -Path "${env:AppData}\..\Local\Programs\$ExeInstallationRelativePath"
}

if ((Test-Path -Path "${env:ProgramFiles(x86)}\$ExeInstallationRelativePath")){
$ExePath = Resolve-Path -Path "${env:ProgramFiles(x86)}\$ExeInstallationRelativePath"
}

if ((Test-Path -Path "${env:ProgramFiles}\$ExeInstallationRelativePath")){
$ExePath = Resolve-Path -Path "${env:ProgramFiles}\$ExeInstallationRelativePath"
}

$ExePathString = $ExePath.ToString()

$ShortcutLocation = [System.IO.Path]::GetFullPath("${env:AppData}\Microsoft\Windows\SendTo\VS Code Diff.lnk")
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutLocation)
$Shortcut.TargetPath = $ExePathString
$Shortcut.Arguments = "--new-window --disable-extensions --diff"
$Shortcut.Save()
Write-Output "Created compare action $ShortcutLocation to $ExePathString"
