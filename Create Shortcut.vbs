' CreateShortcut.vbs
Option Explicit

' ==== EDIT THESE IF NEEDED ====
Dim APP_NAME:   APP_NAME   = "Daily Automation"            ' shortcut display name
Dim TARGET_REL: TARGET_REL = "DailyAutomation.vbs"         ' target file relative to this .vbs
Dim ICON_REL:   ICON_REL   = "app.ico"                     ' icon file relative to this .vbs (optional)
' ==============================

Dim fso, shell, scriptDir, targetPath, iconPath, desktopLnk, startMenuLnk

Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' Resolve paths
targetPath = fso.BuildPath(scriptDir, TARGET_REL)
iconPath   = fso.BuildPath(scriptDir, ICON_REL)

If Not fso.FileExists(targetPath) Then
  WScript.Echo "Target not found: " & targetPath
  WScript.Quit 1
End If

' Where to place shortcuts
desktopLnk   = fso.BuildPath(shell.SpecialFolders("Desktop"),  APP_NAME & ".lnk")
startMenuLnk = fso.BuildPath(shell.SpecialFolders("Programs"), APP_NAME & ".lnk")

' Create both shortcuts
CreateShortcut desktopLnk, targetPath, scriptDir, iconPath
CreateShortcut startMenuLnk, targetPath, scriptDir, iconPath

WScript.Echo "Shortcuts created:" & vbCrLf & " - " & desktopLnk & vbCrLf & " - " & startMenuLnk

' -------- Helpers --------
Sub CreateShortcut(linkPath, tgtPath, workDir, icoPath)
  Dim lnk
  Set lnk = shell.CreateShortcut(linkPath)

  ' If targeting a .vbs file, launch via wscript.exe explicitly (safer than relying on default association)
  If LCase(fso.GetExtensionName(tgtPath)) = "vbs" Then
    lnk.TargetPath = "wscript.exe"
    lnk.Arguments  = """" & tgtPath & """"
  Else
    lnk.TargetPath = tgtPath
    lnk.Arguments  = ""
  End If

  lnk.WorkingDirectory = workDir
  lnk.Description      = APP_NAME
  lnk.WindowStyle      = 1  ' 1=normal, 7=minimized

  If fso.FileExists(icoPath) Then
    lnk.IconLocation = icoPath & ",0"
  Else
    ' Fallback icon
    lnk.IconLocation = "%SystemRoot%\System32\shell32.dll,2"
  End If

  lnk.Save
End Sub
