' Double-click this to launch Crucible with NO console window popping up
' (feels like a normal installed app). Keep it in the same folder as
' report_builder.py and Run_Crucible.bat.
' To put this on your Desktop: right-click this file > Send to > Desktop
' (create shortcut), then you can rename/move the shortcut freely.

Set WshShell = CreateObject("WScript.Shell")
folder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = folder
WshShell.Run "python report_builder.py", 0, False
