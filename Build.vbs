Option Explicit

' This is a test build file

Public Sub Clean()
	WScript.Echo "Cleaning application files..."
End Sub

Public Sub Build()
	Depends("Clean")
	WScript.Echo "Building all application files..."
End Sub

Public Sub ListAllFiles()
	Depends("Clean")
	
	Dim Path, Folder, File
	
	Path = FileSystem.GetParentFolderName(WScript.ScriptFullName)
	WScript.Echo "Enumerating files in '" & Path & "'..."
	
	Set Folder = FileSystem.GetFolder(Path)
	For Each File In Folder.Files
		WScript.Echo File.Name
	Next
	Set Folder = Nothing
End Sub

Public Sub Test()
	Depends("Build")
	Depends("Clean")
	Depends("ListAllFiles")
	
	WScript.Echo "Running applications test suite..."
End Sub
