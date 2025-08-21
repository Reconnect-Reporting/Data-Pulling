' CreateShortcut.vbs — creates a Desktop shortcut to DailyAutomation.bat with app.ico
Option Explicit
Dim shell, fso, desktop, scriptDir, batPath, icoPath, link
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

desktop   = shell.SpecialFolders("Desktop")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
batPath   = fso.BuildPath(scriptDir, "DailyAutomation.bat")
icoPath   = fso.BuildPath(scriptDir, "app.ico")

Set link = shell.CreateShortcut(fso.BuildPath(desktop, "Daily Automation.lnk"))
link.TargetPath = batPath
link.WorkingDirectory = scriptDir
link.IconLocation = icoPath
link.WindowStyle = 1
link.Save

MsgBox "Desktop shortcut created: Daily Automation.lnk", vbInformation, "Setup complete"
