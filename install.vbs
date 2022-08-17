Option Explicit

' Create notification
Dim box
box = MsgBox("This script will install the encoder in your AppData." & vbCrLf & "Click 'OK' to install " & vbCrLf & "Click 'Cancel' to cancel.", 1 + 64 + vbSystemModal, "Installation Script")

' Check result
If box = 2 Then
    box = MsgBox("The installation has been cancelled.", 0 + 16, "Installation Script")
    WScript.Quit
End If

' Get shell and FSO
Dim shell, fso
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Get current directory & app data folders
Dim current_dir : current_dir = fso.GetParentFolderName(WScript.ScriptFullName)
Dim appdata : appdata = shell.ExpandEnvironmentStrings("%APPDATA%")

' Sleep for 10 seconds
box = MsgBox("Downloading files now...", 0 + 64, "Installation Script")

' Download files
shell.Run "powershell -Command ""Invoke-WebRequest https://github.com/Yochran/windows-script-encoding/releases/download/2.0.0/encoder.exe -OutFile encoder.exe"""
shell.Run "powershell -Command ""Invoke-WebRequest https://github.com/Yochran/windows-script-encoding/releases/download/2.0.0/encoder.vbs -OutFile encoder.vbs"""

' Sleep for 2 seconds
WScript.Sleep(5000)

Dim target_folder : target_folder = appdata & "\windows-script-encoding"

' Create folder in app data
fso.CreateFolder target_folder

' Move files to folder
fso.CopyFile current_dir & "\encoder.exe", target_folder & "\encoder.exe"
fso.CopyFile current_dir & "\encoder.vbs", target_folder & "\encoder.vbs"

' Prompt for desktop shortcut
box = MsgBox("Installed files successfully." & vbCrLf & "Would you like to create a desktop shortcut?", 4 + 32, "Installation Script")

' Check result
If box = 6 Then
    ' Create shortcut
    Dim link, desktop : desktop = shell.SpecialFolders("Desktop")
    Set link = shell.CreateShortcut(desktop & "\encoder.lnk")

    link.Description = "Shortcut to encoder.exe"
    link.IconLocation = target_folder & "\encoder.exe,1"
    link.TargetPath = target_folder & "\encoder.exe"
    link.WorkingDirectory = target_folder
    link.Save

    box = MsgBox("Desktop shortcut created successfully.", 0 + 64, "Installation Script")
End If

' Final dialog
box = MsgBox("windows-script-encoder was successfully installed.", 0 + 64, "Installation Script")

' Delete installed files
fso.DeleteFile current_dir & "\encoder.exe"
fso.DeleteFile current_dir & "\encoder.vbs"