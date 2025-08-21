Option Explicit
Dim objShell, objShortcut, fso, desktopPath, shortcutPath, targetPath, iconPath, appDir

Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Get Desktop path
desktopPath = objShell.SpecialFolders("Desktop")

' Get app folder (the folder where this script is located)
appDir = fso.GetParentFolderName(WScript.ScriptFullName)

' Path to your DailyAutomation.vbs
targetPath = fso.BuildPath(appDir, "DailyAutomation.vbs")

' Path to your app.ico (next to main/DailyAutomation.vbs)
iconPath = fso.BuildPath(appDir, "app.ico")

' Create shortcut on Desktop
shortcutPath = desktopPath & "\DailyAutomation.lnk"
Set objShortcut = objShell.CreateShortcut(shortcutPath)

objShortcut.TargetPath = targetPath
objShortcut.WorkingDirectory = appDir

' Use app.ico if it exists
If fso.FileExists(iconPath) Then
    objShortcut.IconLocation = iconPath
Else
    objShortcut.IconLocation = "C:\Windows\System32\wscript.exe,0"
End If

objShortcut.WindowStyle = 1
objShortcut.Save

MsgBox "Shortcut created on Desktop for DailyAutomation.vbs", vbInformation, "Shortcut Creator"
