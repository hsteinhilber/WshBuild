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
Private Const dbText = 10

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
    If FileSystem.FolderExists(ExportPath) Then FileSystem.DeleteFolder ExportPath
    FileSystem.CreateFolder ExportPath

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

    ImportUserInterfaceElements SourcePath

    ImportMetaData XmlDocument

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

    ExportProjectProperties XmlDocument, Project
    ExportDatabaseProperties XmlDocument, InnerApplication.CurrentDb
    ExportObjectDescriptions XmlDocument, Project, InnerApplication.CurrentDb
    
    WriteXmlDocument XmlDocument, FilePath
    Set XmlDocument = Nothing
  End Sub

  Private Sub ExportProjectProperties(ByVal XmlDocument, ByVal Project)
    Dim Parent, Child, Property

    Set Parent = XmlDocument.createElement("properties")
    XmlDocument.documentElement.appendChild Parent

    For Each Property In Project.Properties
      Set Child = XmlDocument.createElement("property")
      Child.setAttribute "name", Property.Name
      Child.setAttribute "value", Property.Value
      Parent.appendChild Child
    Next
  End Sub

  Private Sub ExportDatabaseProperties(ByVal XmlDocument, ByVal Database)
    Dim Parent, Child, Property, ExcludeProperties
    If Database Is Nothing Then Exit Sub

    ExcludeProperties = Array("Name", "Transactions", "Updatable", "CollatingOrder", _
      "Version", "RecordsAffected", "ReplicaID", "DesignMasterID", "Connection", _
      "ProjVer", "Build", "AccessVersion", "Connect")
    Set Parent = XmlDocument.createElement("db-properties")
    XmlDocument.documentElement.appendChild Parent

    For Each Property In Database.Properties
      If Not Contains(ExcludeProperties, Property.Name) Then
        Set Child = XmlDocument.createElement("property")
        Child.setAttribute "name", Property.Name
        Child.setAttribute "value", Property.Value
        Child.setAttribute "type", Property.Type
        Parent.appendChild Child
      End If
    Next
  End Sub

  Private Sub ExportObjectDescriptions(ByVal XmlDocument, ByVal Project, ByVal Database)
    If Database Is Nothing Then
      ExportDescriptions XmlDocument, "forms", Project.AllForms
      ExportDescriptions XmlDocument, "reports", Project.AllReports
      ExportDescriptions XmlDocument, "macros", Project.AllMacros
      ExportDescriptions XmlDocument, "modules", Project.AllModules
    Else
      ExportDescriptions XmlDocument, "tables", Database.Containers("Tables").Documents
      ExportDescriptions XmlDocument, "forms", Database.Containers("Forms").Documents
      ExportDescriptions XmlDocument, "reports", Database.Containers("Reports").Documents
      ExportDescriptions XmlDocument, "macros", Database.Containers("Scripts").Documents
      ExportDescriptions XmlDocument, "modules", Database.Containers("Modules").Documents
    End If
  End Sub

  Private Sub ExportDescriptions(ByVal XmlDocument, ByVal Parent, ByVal Collection)
    Dim DescriptionsElement, ParentElement, ChildElement
    Dim AccessObject, Description

    Set DescriptionsElement = XmlDocument.selectSingleNode("database/descriptions")
    If DescriptionsElement Is Nothing Then 
      Set DescriptionsElement = XmlDocument.createElement("descriptions")
      XmlDocument.documentElement.appendChild DescriptionsElement
    End If

    Set ParentElement = DescriptionsElement.selectSingleNode(Parent)
    If ParentElement Is Nothing Then
      Set ParentElement = XmlDocument.createElement(Parent)
      DescriptionsElement.appendChild ParentElement
    End If

    For Each AccessObject In Collection
      Set ChildElement = XmlDocument.createElement("object")
      ChildElement.setAttribute "name", AccessObject.Name
      On Error Resume Next
      Description = vbNullString
      Description = AccessObject.Properties("Description").Value
      On Error Goto 0
      ChildElement.appendChild XmlDocument.createTextNode(Description)
      ParentElement.appendChild ChildElement
    Next
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
            InnerApplication.SaveAsText acQuery, DbObject.Name, FileSystem.BuildPath(OutputPath, DbObject.Name & ".qry")
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

	ImportRelationships XmlDocument
	
    ImportProjectObjects FileSystem.BuildPath(SourcePath, "Queries"), acQuery
  End Sub
  
  Private Sub ImportProjectObjects(ByVal ImportPath, ByVal ObjectType)
    Dim Folder, File, ObjectName
    
    WScript.Echo "Generating " & FileSystem.GetBaseName(ImportPath) & "..."
    Set Folder = FileSystem.GetFolder(ImportPath)
    For Each File In Folder.Files
      ObjectName = FileSystem.GetBaseName(File.Name)
        InnerApplication.LoadFromText ObjectType, ObjectName, File.Path
    Next
  End Sub 

  Private Sub ImportMetaData(ByVal XmlDocument)
    ImportProperties XmlDocument
    ImportDatabaseProperties XmlDocument
    ImportObjectDescriptions XmlDocument
  End Sub

  Private Sub ImportProperties(ByVal XmlDocument)
    Dim PropertiesElement, PropertyElement
    Dim Name, Value

    Set PropertiesElement = XmlDocument.selectSingleNode("database/properties")
    For Each PropertyElement In PropertiesElement.selectNodes("property")
      Name = PropertyElement.selectSingleNode("@name").nodeValue
      Value = PropertyElement.selectSingleNode("@value").nodeValue
      InnerApplication.CurrentProject.Properties.Add Name, Value
    Next
  End Sub
  
  Private Sub ImportDatabaseProperties(ByVal XmlDocument)
    Dim PropertiesElement, PropertyElement
    Dim Name, Value, ValueType, Database

    Set Database = Application.CurrentDb
    Set PropertiesElement = XmlDocument.selectSingleNode("database/db-properties")
    For Each PropertyElement In PropertiesElement.selectNodes("property")
      Name = PropertyElement.selectSingleNode("@name").nodeValue
      Value = PropertyElement.selectSingleNode("@value").nodeValue
      ValueType = CInt(PropertyElement.selectSingleNode("@type").nodeValue)
      On Error Resume Next 
      Database.Properties(Name).Value = Value
      If Err = 3270 Then
        On Error Goto 0
        Dim Property: Set Property = Database.CreateProperty(Name, ValueType, Value)
        Database.Properties.Append Property
      ElseIf Err <> 0 Then
        On Error Goto 0
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
      End If
    Next
  End Sub

  Private Sub ImportObjectDescriptions(ByVal XmlDocument)
    Dim Database, Project

    Set Database = InnerApplication.CurrentDb
    Set Project = InnerApplication.CurrentProject

    ' HACK: This is very hackish and needs major refactoring
    If Database Is Nothing Then
      ImportDescriptions Project.AllForms, XmlDocument.selectNodes("database/descriptions/forms/object"), "AddAccessObjectProperty"
      ImportDescriptions Project.AllReports, XmlDocument.selectNodes("database/descriptions/reports/object"), "AddAccessObjectProperty"
      ImportDescriptions Project.AllMacros, XmlDocument.selectNodes("database/descriptions/macros/object"), "AddAccessObjectProperty"
      ImportDescriptions Project.AllModules, XmlDocument.selectNodes("database/descriptions/modules/object"), "AddAccessObjectProperty"
    Else
      ImportDescriptions Database.Containers("Tables").Documents, XmlDocument.selectNodes("database/descriptions/tables/object"), "AddContainerDocumentProperty"
      ImportDescriptions Database.Containers("Forms").Documents, XmlDocument.selectNodes("database/descriptions/forms/object"), "AddContainerDocumentProperty"
      ImportDescriptions Database.Containers("Reports").Documents, XmlDocument.selectNodes("database/descriptions/reports/object"), "AddContainerDocumentProperty"
      ImportDescriptions Database.Containers("Scripts").Documents, XmlDocument.selectNodes("database/descriptions/macros/object"), "AddContainerDocumentProperty"
      ImportDescriptions Database.Containers("Modules").Documents, XmlDocument.selectNodes("database/descriptions/modules/object"), "AddContainerDocumentProperty"
    End If
  End Sub

  Private Sub ImportDescriptions(ByVal Collection, ByVal Objects, ByVal AddMethod)
    Dim ObjectElement, Name, Description, AccessObject

    For Each ObjectElement In Objects
      Name = ObjectElement.selectSingleNode("@name").nodeValue
      Description = ObjectElement.text
      If Len(Description) > 0 Then
        Set AccessObject = Collection(Name)
        Execute AddMethod & " AccessObject, ""Description"", Description"
      End If
    Next
  End Sub
  
  Private Sub ImportRelationships(ByVal XmlDocument)
    Dim RelationshipsElement, Database
    
	WScript.Echo "Generating Relationships..."
    Set RelationshipsElement = XmlDocument.selectSingleNode("database/relationships")
	Set Database = InnerApplication.CurrentDb
   
    Dim RelationElement
    For Each RelationElement In RelationshipsElement.childNodes
	  Dim Name, Table, ForeignTable, Attributes, Relation
	  Name = RelationElement.selectSingleNode("@name").nodeValue
	  Table = RelationElement.selectSingleNode("@table").nodeValue
	  ForeignTable = RelationElement.selectSingleNode("@foreign-table").nodeValue
	  Attributes = RelationElement.selectSingleNode("@attributes").nodeValue	  
      Set Relation = Database.CreateRelation(Name, Table, ForeignTable, Attributes)
	  
	  Dim FieldElement
	  For Each FieldElement In RelationElement.selectNodes("fields/field")
        Dim FieldName, ForeignName, Field
		FieldName = FieldElement.selectSingleNode("@name").nodeValue
		ForeignName = FieldElement.selectSingleNode("@foreign-name").nodeValue
		Set Field = Relation.CreateField(FieldName)
		Field.ForeignName = ForeignName
		Relation.Fields.Append Field
	  Next
	  
	  Database.Relations.Append Relation
    Next
  End Sub

  Private Sub AddContainerDocumentProperty(ByVal Document, ByVal Name, ByVal Value)
    On Error Resume Next
    Dim Property
    Set Property = Document.CreateProperty(Name, dbText, Value)
    Document.Properties.Append Property
    If Err = 3367 Then Document.Properties(Name).Value = Value: Err.Clear
  End Sub

  Private Sub AddAccessObjectProperty(ByVal AccessObject, ByVal Name, ByVal Value)
    AccessObject.Properties.Add Name, Value
  End Sub

  Private Sub ImportUserInterfaceElements(ByVal SourcePath)
    ImportProjectObjects FileSystem.BuildPath(SourcePath, "Forms"), acForm
    ImportProjectObjects FileSystem.BuildPath(SourcePath, "Reports"), acReport
    ImportProjectObjects FileSystem.BuildPath(SourcePath, "Macros"), acMacro
    ImportProjectObjects FileSystem.BuildPath(SourcePath, "Modules"), acModule
  End Sub

  Public Function Contains(ByRef Values, ByVal Search)
    Contains = InStr(vbNullChar & Join(Values, vbNullChar) & vbNullChar, _
        vbNullChar & Search & vbNullChar) > 0
  End Function

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
