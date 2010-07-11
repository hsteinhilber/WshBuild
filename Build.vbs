Option Explicit

' This is a test build file

Public Sub Clean()
	WScript.Echo "Cleaning application files..."
End Sub

Public Sub Build()
	Depends("Clean")
	WScript.Echo "Building all application files..."
End Sub

Public Sub AnotherTask()
	Depends("Clean")
	
	WScript.Echo "Another task that depends on the system being clean."
End Sub

Public Sub Test()
	Depends("Build")
	Depends("Clean")
	Depends("AnotherTask")
	
	WScript.Echo "Running applications test suite..."
End Sub
