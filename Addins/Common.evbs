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

Public Function BuildVersion(ByVal BaseVersion)
	' Takes a base version in the form of [Major.Minor] and returns a full version.
	' Result: A version number in the form of [Major.Minor.Build.Revision].
	BuildVersion = BaseVersion & "." & BuildFromDate() & "." & RevisionFromTime()
End Function

Public Function BuildFromDate()
	' Converts the current date into a build number.
	' Result: A build number in the range of 0 to 9999.
	BuildFromDate = DateDiff("D", "01/01/2001", Now()) Mod 10000
End Function

Public Function RevisionFromTime()
	' Converts the time since midnight into a revision number.
	RevisionFromTime = DateDiff("S", DateValue(Now()), Now()) \ 2
End Function 

Public Function MapPath(ByVal FileName)
	FileName = WshShell.ExpandEnvironmentStrings(FileName)
	MapPath = FileSystem.GetAbsolutePathName(FileName)
End Function

Public Function CopyFileToFolder(ByVal FileName, ByVal FolderPath)
	Dim BaseName
	BaseName = FileSystem.GetFile(MapPath(FileName)).Name
	FileSystem.CopyFile MapPath(FileName), FileSystem.BuildPath(MapPath(FolderPath), BaseName), True
End Function 

Public Function StrStartsWith(ByVal Text, ByVal SearchText)
  StrStartsWith = (LCase(Left(Text, Len(SearchText))) = LCase(SearchText))
End Function

Public Function StrEndsWith(ByVal Text, ByVal SearchText)
  StrEndsWith = (LCase(Right(Text, Len(SearchText))) = LCase(SearchText))
End Function 
