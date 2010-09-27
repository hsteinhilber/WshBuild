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

Public Const acTable = 0
Public Const acQuery = 1
Public Const acForm = 2
Public Const acReport = 3
Public Const acMacro = 4
Public Const acModule = 5

Private Const acExportDelim = 2
Private Const acStructureOnly = 0
Private Const acAppendData = 2
Private Const adSaveCreateOverWrite = 2

Public Access ' InitializeAccessAddin dependency task will initialize this variable for use in tasks

Class AccessAddin
  Private InnerApplication

  Public Sub Class_Initialize()
    Set InnerApplication = CreateObject("Access.Application")
  End Sub

  Public Sub Class_Terminate()   
    If Not InnerApplication Is Nothing Then InnerApplication.Quit
    Set InnerApplication = Nothing
  End Sub

  Public Property Get Application()
    Set Application = InnerApplication
  End Property

  Public Sub OpenDatabase(ByVal DatabaseFileName)
    DatabaseFileName = MapPath(DatabaseFileName)
    If CheckFileExtension(DatabaseFileName, ".adp") Or CheckFileExtension(DatabaseFileName, ".ade") Then
      InnerApplication.OpenAccessProject DatabaseFileName
    Else
      InnerApplication.OpenCurrentDatabase DatabaseFileName
    End If
  End Sub

  Public Sub CloseDatabase()
    InnerApplication.CloseCurrentDatabase
  End Sub

  Public Sub CompileDatabase()
    Dim SourcePath, DestinationPath

    SourcePath = InnerApplication.CurrentProject.FullName
    DestinationPath = Left(SourcePath, Len(SourcePath) - 1) & "e"

    ' HACK: Using an undocumented SysCmd to build MDE/ADE files
    Me.CloseDatabase
    InnerApplication.SysCmd 603, CStr(SourcePath), CStr(DestinationPath)
    Me.OpenDatabase SourcePath
  End Sub

  Public Property Get Properties(ByVal Name)
    On Error Resume Next
    Properties = InnerApplication.CurrentProject.Properties(Name)
    If Err = 2455 Then
      InnerApplication.CurrentProject.Properties.Add Name, vbNullString
      Properties = vbNullString
    End If
  End Property
  Public Property Let Properties(ByVal Name, ByVal Value)
    InnerApplication.CurrentProject.Properties.Add Name, Value
  End Property

  Public Sub ExportObjects(ByVal ExportPath)
    On Error Resume Next 
    
    Dim FilePath: FilePath = InnerApplication.CurrentProject.FullName
    ExportPath = MapPath(ExportPath)
    If Not FileSystem.FolderExists(ExportPath) Then FileSystem.CreateFolder ExportPath

    ExportMetaData ExportPath

    ExportUserInterfaceElements ExportPath

    If CheckFileExtension(FilePath, ".adp") Or CheckFileExtension(FilePath, ".ade") Then    
        ExportAdpDataElements ExportPath
    Else    
        ExportMdbDataElements ExportPath
    End If
  End Sub 

  Public Sub GenerateDatabase(ByVal SourcePath, ByVal DataPath, ByVal OutputPath)
    On Error Resume Next

    Dim XmlDocument: Set XmlDocument = CreateObject("Msxml2.DomDocument.6.0")
    XmlDocument.load FileSystem.BuildPath(SourcePath, "Project.xml")

    Dim FileName, FileType
    FileType = XmlDocument.selectSingleNode("database/@type").nodeValue
    FileName = XmlDocument.selectSingleNode("database/@name").nodeValue
    FileName = FileSystem.BuildPath(OutputPath, FileName)

    If FileSystem.FileExists(FileName) Then _
        FileSystem.MoveFile FileName, FileName & "." & BackupTimestamp & ".bak"

    If FileType = "1" Then
      ' InnerApplication.NewAccessProject FileName
      ' ImportAdpDataElements XmlDocument, SourcePath, DataPath
    ElseIf FileType = "2" Then
      InnerApplication.NewCurrentDatabase FileName
      ImportMdbDataElements XmlDocument, SourcePath, DataPath, OutputPath
    End If

    ImportProperties XmlDocument

    ImportUserInterfaceElements SourcePath

    InnerApplication.CloseCurrentDatabase
    Set XmlDocument = Nothing
  End Sub

  Private Sub ExportMetaData(ByVal ExportPath)
    Dim XmlDocument: Set XmlDocument = CreateObject("Msxml2.DomDocument.6.0")
    Dim FilePath: FilePath = FileSystem.BuildPath(ExportPath, "Project.xml")
    Dim Project: Set Project = InnerApplication.CurrentProject

    XmlDocument.documentElement = XmlDocument.createElement("database")
    XmlDocument.documentElement.setAttribute "name", Project.Name
    XmlDocument.documentElement.setAttribute "type", Project.ProjectType

    Dim Parent, Child, Property
    Set Parent = XmlDocument.createElement("properties")
    XmlDocument.documentElement.appendChild Parent

    For Each Property In Project.Properties
      Set Child = XmlDocument.createElement("property")
      Child.setAttribute "name", Property.Name
      Child.setAttribute "value", Property.Value
      Parent.appendChild Child
    Next

    WriteXmlDocument XmlDocument, FilePath
    Set XmlDocument = Nothing
  End Sub

  Private Sub WriteXmlDocument(ByVal XmlDocument, ByVal FilePath)
    Dim Reader, Writer, Stream

    Set Stream = CreateObject("ADODB.Stream")
    Set Writer = CreateObject("Msxml2.MXXMLWriter.6.0")
    Set Reader = CreateObject("Msxml2.SAXXMLReader.6.0")

    Stream.Open
    Stream.Charset = "UTF-8"

    Writer.encoding = "UTF-8"
    Writer.indent = True
    Writer.omitXMLDeclaration = False
    Writer.output = Stream

    Set Reader.contentHandler = Writer
    Set Reader.dtdHandler = Writer
    Set Reader.errorHandler = Writer

    Reader.putProperty "http://xml.org/sax/properties/declaration-handler", Writer
    Reader.putProperty "http://xml.org/sax/properties/lexical-handler", Writer

    Reader.parse XmlDocument.XML
    Writer.flush

    Stream.SaveToFile FilePath, adSaveCreateOverWrite
    Stream.Close

    Set Reader = Nothing
    Set Writer = Nothing
    Set Stream = Nothing
  End Sub

  Private Sub ExportUserInterfaceElements(ByVal ExportPath)
    Dim Project

    Set Project = InnerApplication.CurrentProject
    ExportProjectObjects Project.AllForms, FileSystem.BuildPath(ExportPath, "Forms"), _
      acForm, ".frm"
    ExportProjectObjects Project.AllReports, FileSystem.BuildPath(ExportPath, "Reports"), _
      acReport, ".rpt"
    ExportProjectObjects Project.AllMacros, FileSystem.BuildPath(ExportPath, "Macros"), _
      acMacro, ".mcr"
    ExportProjectObjects Project.AllModules, FileSystem.BuildPath(ExportPath, "Modules"), _
      acModule, ".bas"
  End Sub

  Private Sub ExportMdbDataElements(ByVal ExportPath)
    Dim Database, DbObject, OutputPath
    On Error Resume Next 
    
    Set Database = InnerApplication.CurrentDb 
    
    WScript.Echo "Exporting Tables..."
    OutputPath = FileSystem.BuildPath(ExportPath, "Tables")
    If FileSystem.FolderExists(OutputPath) Then FileSystem.DeleteFolder OutputPath
    FileSystem.CreateFolder OutputPath
    For Each DbObject In Database.TableDefs
        If Not StrStartsWith(DbObject.Name, "msys") And Not StrStartsWith(DbObject.Name, "~") Then
            If Not IsLinkedTable(DbObject) Then
                InnerApplication.ExportXML acTable, DbObject.Name, , FileSystem.BuildPath(OutputPath, DbObject.Name & ".xsd")
            End If
        End If
    Next

    WScript.Echo "Exporting Linked Tables..."
    OutputPath = FileSystem.BuildPath(ExportPath, "Project.xml")
    ExportLinkedTables Database.TableDefs, _
      FileSystem.GetParentFolderName(Database.Name), OutputPath
    
    WScript.Echo "Exporting Relationships..."
    OutputPath = FileSystem.BuildPath(ExportPath, "Project.xml")
    ExportRelationships Database.Relations, OutputPath
    
    WScript.Echo "Exporting Queries..."
    OutputPath = FileSystem.BuildPath(ExportPath, "Queries")
    If FileSystem.FolderExists(OutputPath) Then FileSystem.DeleteFolder OutputPath
    FileSystem.CreateFolder OutputPath
    For Each DbObject In Database.QueryDefs
        If Not StrStartsWith(DbObject.Name, "~") Then
            InnerApplication.SaveAsText acQuery, DbObject.Name, FileSystem.BuildPath(OutputPath, DbObject.Name)
        End If
    Next 
  End Sub

  Private Sub ExportLinkedTables(ByVal Tables, ByVal DatabasePath, ByVal OutputFilePath)
    Dim XmlDocument, XmlLinkedTables, XmlLink
    Dim Table, LinkPath, LocalName, SourceName

    Set XmlDocument = CreateObject("Msxml2.DomDocument.6.0")
    XmlDocument.load OutputFilePath

    Set XmlLinkedTables = XmlDocument.selectSingleNode("database/linked-tables")
    If XmlLinkedTables Is Nothing Then 
      Set XmlLinkedTables = XmlDocument.createElement("linked-tables")
      XmlDocument.documentElement.appendChild XmlLinkedTables
    End If

    For Each Table In Tables
      If IsLinkedTable(Table) Then
        LocalName = Table.Name
        SourceName = Table.SourceTableName
        LinkPath = Mid(Table.Connect, 11)
        If InStr(LinkPath, DatabasePath) = 1 Then 
          LinkPath = Replace(LinkPath, DatabasePath, "")
          LinkPath = FileSystem.BuildPath(".", LinkPath)
        End If
      
        Set XmlLink = XmlDocument.createElement("link")
        XmlLink.setAttribute "name", LocalName
        XmlLink.setAttribute "source", SourceName
        XmlLink.setAttribute "database", LinkPath
        XmlLinkedTables.appendChild XmlLink
      End If
    Next

    WriteXmlDocument XmlDocument, OutputFilePath
    Set XmlLink = Nothing
    Set XmlLinkedTables = Nothing
    Set XmlDocument = Nothing
  End Sub

  Private Sub ExportRelationships(ByVal Relations, ByVal OutputFilePath)
    Dim XmlDocument, RelationshipsElement
    
    Set XmlDocument = CreateObject("Msxml2.DomDocument.6.0")
    XmlDocument.Load OutputFilePath
    
    Set RelationshipsElement = XmlDocument.selectSingleNode("database/relationships")
    If RelationshipsElement Is Nothing Then
      Set RelationshipsElement = XmlDocument.createElement("relationships")
      XmlDocument.documentElement.appendChild RelationshipsElement
    End If
    
    Dim Relation
    For Each Relation In Relations
      Dim RelationElement: Set RelationElement = XmlDocument.createElement("relation")
      RelationshipsElement.appendChild RelationElement
      
      RelationElement.setAttribute "name", Relation.Name
      RelationElement.setAttribute "table", Relation.Table
      RelationElement.setAttribute "foreign-table", Relation.ForeignTable
      RelationElement.setAttribute "attributes", Relation.Attributes
      
      If Relation.Fields.Count > 0 Then
        Dim FieldsElement: Set FieldsElement = XmlDocument.createElement("fields")
        RelationElement.appendChild FieldsElement
        
        Dim Field, FieldElement
        For Each Field In Relation.Fields
          Set FieldElement = XmlDocument.createElement("field")
          FieldsElement.appendChild FieldElement
          
          FieldElement.setAttribute "name", Field.Name
          FieldElement.setAttribute "foreign-name", Field.ForeignName
        Next
      End If
    Next
    
    WriteXmlDocument XmlDocument, OutputFilePath
    Set XmlDocument = Nothing
  End Sub

  Private Sub ExportAdpDataElements(ByVal ExportPath)
    On Error Resume Next
    'TODO: Find a useful way of exporting SQL server objects

    WScript.Echo "Exporting Tables..."
    
    WScript.Echo "Exporting Foreign Keys..."

    WScript.Echo "Exporting Indices..."

    WScript.Echo "Exporting Views..."

    WScript.Echo "Exporting Stored Procedures..."
  End Sub 

  Private Sub ExportProjectObjects(ByVal Objects, ByVal ExportPath, ByVal ObjectType, ByVal Extension)
    Dim AccessObject 
    
    WScript.Echo "Exporting " & FileSystem.GetBaseName(ExportPath) & "..."
    If FileSystem.FolderExists(ExportPath) Then FileSystem.DeleteFolder ExportPath
    FileSystem.CreateFolder ExportPath
    
    For Each AccessObject In Objects
        InnerApplication.SaveAsText ObjectType, AccessObject.Name, FileSystem.BuildPath(ExportPath, AccessObject.Name & Extension)
    Next
  End Sub 

  Private Sub ImportMdbDataElements(ByVal XmlDocument, ByVal SourcePath, ByVal DataPath, ByVal OutputPath)
    Dim Database, DbObject
    Dim File, Folder
    Set Database = InnerApplication.CurrentDb

    WScript.Echo "Generating Tables..."
    Set Folder = FileSystem.GetFolder(FileSystem.BuildPath(SourcePath, "Tables"))
    For Each File In Folder.Files
      InnerApplication.ImportXML File.Path, acStructureOnly
    Next

    WScript.Echo "Generating Linked Tables..."
    Dim LinkNode, TableName, SourceName, LinkPath
    For Each LinkNode In XmlDocument.selectNodes("database/linked-tables/link")
      TableName = LinkNode.selectSingleNode("@name").nodeValue
      SourceName = LinkNode.selectSingleNode("@source").nodeValue
      LinkPath = LinkNode.selectSingleNode("@database").nodeValue
      If Not LinkPath = MapPath(LinkPath) Then ' This was not already an absolute path
        LinkPath = MapPath(FileSystem.BuildPath(OutputPath, LinkPath))
      End If
      
      Set DbObject = Database.CreateTableDef(TableName)
      DbObject.SourceTableName = SourceName
      DbObject.Connect = ";DATABASE=" & LinkPath
      Database.TableDefs.Append DbObject
    Next
    Database.TableDefs.Refresh
    
    If FileSystem.FolderExists(DataPath) Then
      WScript.Echo "Generating Data..."
      Set Folder = FileSystem.GetFolder(DataPath)
      For Each File In Folder.Files
        InnerApplication.ImportXML File.Path, acAppendData
      Next
    End If

    WScript.Echo "Generating Queries..."
    ' Import Queries

  End Sub
  
  Private Sub ImportProperties(ByVal XmlDocument)
    Dim PropertiesElement, PropertyElement
    Dim Name, Value

    Set PropertiesElement = XmlDocument.selectSingleNode("database/properties")
    For Each PropertyElement In PropertiesElement
      Name = PropertyElement.selectSingleNode("@name").nodeValue
      Value = PropertyElement.selectSingleNode("@value").nodeValue
      InnerApplication.CurrentProject.Properties.Add Name, Value
    Next
  End Sub
  
  Private Sub ImportUserInterfaceElements(ByVal SourcePath)
    ImportModules FileSystem.BuildPath(SourcePath, Modules)

    ' 10) ** Unknown: Import Queries (Possibly modify to only export SQL and then create QueryDef and set SQL)
    ' 11) ** Unknown: Import Forms
    ' 12) ** Unknown: Import Reports
    ' 13) ** Unknown: Import Macros
  End Sub

  Private Sub ImportModules(ByVal SourcePath)
    Dim Folder, File

    Set Folder = FileSystem.GetFolder(SourcePath)
    For Each File in Folder.Files
      InnerApplication.VBE.ActiveVBProject.VBComponents.Import File.Path
    Next
  End Sub

  Private Function BackupTimestamp()
    BackupTimestamp = Year(Now) & Right(100 + Month(Now),2) & Right(100 + Day(Now),2) & _
      Right(100 + Hour(Now),2) & Right(100 + Minute(Now),2) & Right(100 + Second(Now),2)
  End Function

  Private Function CheckFileExtension(ByVal Path, ByVal Extension)
    CheckFileExtension = (LCase(FileSystem.GetExtensionName(Path)) = LCase(Extension))
  End Function

  Private Function IsLinkedTable(ByVal Table)
    IsLinkedTable = (Len(Table.Connect) > 0)
  End Function 
End Class

Public Sub InitializeAccessAddin()
  Set Access = New AccessAddin
End Sub
