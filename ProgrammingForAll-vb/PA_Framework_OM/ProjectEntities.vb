Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class ProjectEntities
  Inherits PAEnts

  Overrides Function GetSelectDBItems() As String
    GetSelectDBItems = _
              strDBO & " ID " & _
              strDBO & ",EntityName " & _
              strDBO & ",EntityCollectionName " & _
              strDBO & ",EntityDBTableName " & _
              strDBO & ",EntityShortName " & _
              strDBO & ",EntityItems " & _
              strDBO & ",GenerateFlag " & _
              strDBO & ",InputFile " & _
              strDBO & ",FilesToGenerate " & _
              strDBO & ",ExcelSheetName " & _
              strDBO & ",ExcelFormName " & _
              strDBO & ",DateTimeCodeGenerated " & _
              strDBO & ",LastUpdate "
  End Function

  Overrides Sub LoadEntityItemsForThisEntity(ByRef ent As Object)

    ent.ID = "" & AppSettings.rst.Fields("ID").Value
    ent.EntityName = "" & AppSettings.rst.Fields("EntityName").Value
    ent.EntityCollectionName = "" & AppSettings.rst.Fields("EntityCollectionName").Value
    ent.EntityDBTableName = "" & AppSettings.rst.Fields("EntityDBTableName").Value
    ent.EntityShortName = "" & AppSettings.rst.Fields("EntityShortName").Value
    ent.EntityItems = "" & AppSettings.rst.Fields("EntityItems").Value
    ent.GenerateFlag = "" & AppSettings.rst.Fields("GenerateFlag").Value
    ent.InputFile = "" & AppSettings.rst.Fields("InputFile").Value
    ent.FilesToGenerate = "" & AppSettings.rst.Fields("FilesToGenerate").Value
    ent.ExcelSheetName = "" & AppSettings.rst.Fields("ExcelSheetName").Value
    ent.ExcelFormName = "" & AppSettings.rst.Fields("ExcelFormName").Value
    ent.DateTimeCodeGenerated = AppSettings.rst.Fields("DateTimeCodeGenerated").Value

    Call BLFunctions.BuildChildEntityObjects(cPAProjects, ent, "ProjectEntityItems", "" & AppSettings.rst.Fields("EntityItems").Value)
  End Sub

  Overrides Sub RecordChanges(ByVal sOperation As String, ByRef ent As Object)
    Call gRecordChanges(sOperation, "EntityName", ent.EntityName, "en:")
    Call gRecordChanges(sOperation, "EntityCollectionName", ent.EntityCollectionName, "en:")
    Call gRecordChanges(sOperation, "EntityShortName", ent.EntityShortName, "en:")
    Call gRecordChanges(sOperation, "EntityDBTableName", ent.EntityDBTableName, "en:")
    Call gRecordChanges(sOperation, "GenerateFlag", ent.GenerateFlag, "en:")
    Call gRecordChanges(sOperation, "InputFile", ent.InputFile, "en:")
    Call gRecordChanges(sOperation, "FilesToGenerate", ent.FilesToGenerate, "en:")
    Call gRecordChanges(sOperation, "DateTimeCodeGenerated", ent.DateTimeCodeGenerated, "en:")
    Call gRecordChanges(sOperation, "ExcelSheetName", ent.ExcelSheetName, "esn:")
    Call gRecordChanges(sOperation, "ExcelFormName", ent.ExcelFormName, "efn:")
    Call gRecordChanges(sOperation, "EntityItems", ent.ChildEntityString("ProjectEntityItems"), "EIs:")

  End Sub

  Overrides Function GetSelectFromTable() As String
    GetSelectFromTable = " FROM " & strDBO & "ProjectEntities "
  End Function

  Overrides Function CreateNewEntity() As PAEnt
    CreateNewEntity = New ProjectEntity
  End Function

  Overrides Sub SetCurrentEntity(ByRef ent As PAEnt)
    currePrjEnt_ = ent
  End Sub

  Overrides Sub BuildDomainModel()
    '  Dim eb_ As Course
    '  Dim e1_ As Entity1
    '  'now add eCrse_s to e1_s
    '  For Each e1_ In cEntity1s.Items  'assumed cEntity1s have been loaded
    '    Set e1_.ChildCourses = e1_.BuildChildEntityObjects(e1_.ChildCourses, e1_.ChildEntityString(CurreCrse_))
    '    'For Each eCrse_ In e1_.ChildCourses.Items
    '    ' If Not eCrse_.ContainsParentEntity(e1_) Then  'BuildObjectModel is called from AddtoDB() too. hence to check for already existing.
    '    '  Call eCrse_.ParentEntity1s.Add(e1_)   'add parents to Course
    '    'End If
    '    'Next eCrse_
    '  Next e1_
  End Sub
End Class
