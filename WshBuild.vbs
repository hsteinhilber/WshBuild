Option Explicit
' WshScript Build System
' Copyright (c) 2010 Harry Steinhilber, Jr.

' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.

Dim WshShell 
Dim WshNetwork
Dim FileSystem
Dim ExecutedTasks

Const UserBuildFile = "Build.vbs"

Sub Main()
	On Error Resume Next
	Initialize
	EnsureRunningInConsole
	DisplayLogo 
	ImportFiles
	ExecuteTasks
	ExitBuild Err.Number, Err.Description 
End Sub

Sub Initialize()
	Set WshShell      = CreateObject("WScript.Shell")
	Set WshNetwork    = CreateObject("WScript.Network")
	Set FileSystem    = CreateObject("Scripting.FileSystemObject")
	Set ExecutedTasks = CreateObject("Scripting.Dictionary") 
End Sub

Sub Cleanup()
	Set ExecutedTasks = Nothing
	Set FileSystem    = Nothing
	Set WshNetwork    = Nothing
	Set WshShell      = Nothing
End Sub

Sub DisplayLogo()
	WScript.Echo "Windows Scripting Host Build System Version 1.0"
	WScript.Echo "Copyright (C) 2010 Harry Steinhilber, Jr."
	WScript.Echo 
End Sub

Sub EnsureRunningInConsole()
	If Right(WScript.FullName, 11) = "wscript.exe" Then
		Dim Command, Argument, ReturnValue
		Command = "cscript """ & WScript.ScriptFullName & """ "
		For Each Argument In WScript.Arguments
			Command = Command & Argument & " "
		Next 
		ExitBuild WshShell.Run(Command, 1, True), ""
	End If
End Sub 

Sub ExitBuild(ByVal ExitCode, ByVal Description) 
	Cleanup
	If ExitCode <> 0 Then 
		WScript.Echo "Build Failed - Error: " & ExitCode & " - " & Description
	Else
		WScript.Echo "Build Successful"
	End If
	WScript.Quit ExitCode
End Sub

Sub ImportFiles()
	ImportAddinFiles
	Import UserBuildFile
End Sub

Sub ImportAddinFiles()
	Dim Path, Folder, File

	Path = FileSystem.GetParentFolderName(WScript.ScriptFullName)
	Path = FileSystem.BuildPath(Path, "Addins")
	If FileSystem.FolderExists(Path) Then 
		WScript.Echo "Loading add-ins from " & Path
		Set Folder = FileSystem.GetFolder(Path)
		For Each File In Folder.Files
			If FileSystem.GetExtensionName(File.Name) = "vbs" Then
				On Error Resume Next
				WScript.Echo "Importing " & FileSystem.GetBaseName(File.Name) & " extensions"
				Import File.Path
				If Err <> 0 Then WScript.Echo "Failed to load " & File.Name & " extensions - " & Err.Description
				On Error Goto 0 
			End If
		Next
		Set Folder = Nothing
		WScript.Echo 
		Err.Clear
	End If
End Sub

Sub Import(ByVal FileName)
	Dim TextStream, CodeText

	FileName = WshShell.ExpandEnvironmentStrings(FileName)
	FileName = FileSystem.GetAbsolutePathName(FileName)

	Set TextStream = FileSystem.OpenTextFile(FileName)
	CodeText = TextStream.ReadAll
	TextStream.Close
	ExecuteGlobal CodeText
End Sub

Sub DisplayTaskStatus(ByVal TaskName, ByVal Status)
	WScript.Echo "[" & TaskName & "]: " & Status & vbNewLine
End Sub

Sub ExecuteTasks()
	Dim Argument 

	For Each Argument In WScript.Arguments
		ExecuteTask Argument
	Next
End Sub

Public Sub ExecuteTask(ByVal TaskName) 
	On Error Resume Next

	DisplayTaskStatus TaskName, ""
	Execute "Call " & TaskName
	WScript.Echo 
	If Err <> 0 Then
		Dim Num, Src, Desc: Num = Err.Number: Src = Err.Source : Desc = Err.Description
		DisplayTaskStatus TaskName, Err.Description
		On Error Goto 0: Err.Raise Num, Src, Desc
	End If
	DisplayTaskStatus TaskName, "Success"
End Sub

Public Sub Depends(ByVal TaskName)
	If ExecutedTasks.Exists(TaskName) Then Exit Sub

	ExecuteTask TaskName
	ExecutedTasks.Add TaskName, True
End Sub 

' Call entry point
Call Main()
