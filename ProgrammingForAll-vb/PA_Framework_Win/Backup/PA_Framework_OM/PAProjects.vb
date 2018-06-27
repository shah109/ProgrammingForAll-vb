Option Explicit On

Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class PAProjects
  Inherits PAEnts

  Private sItemByProjectNameHashTable As Hashtable

  Overrides Sub LoadEntityItemsForThisEntity(ByRef ent As Object)

    ent.ID = "" & AppSettings.rst.Fields("ID").Value
    ent.ProjectName = "" & AppSettings.rst.Fields("ProjectName").Value
    ent.ProjectDescription = "" & AppSettings.rst.Fields("ProjectDescription").Value

    Call BLFunctions.BuildChildEntityObjects(cPAProjects, ent, "ProjectEntities", "" & AppSettings.rst.Fields("ProjectEntities").Value)
  End Sub

  Overrides Sub Recordchanges(ByVal sOperation As String, ByRef ent As Object)
    Call gRecordChanges(sOperation, "ProjectName", ent.ProjectName, "pn:")
    Call gRecordChanges(sOperation, "ProjectDescription", ent.ProjectDescription, "pd:")

    Call gRecordChanges(sOperation, "ProjectEntities", ent.ChildEntityString("ProjectEntities"), "PrjEnts:")
  End Sub

  Overrides Function GetSelectDBItems() As String
    GetSelectDBItems = _
              strDBO & " ID " & _
              strDBO & ",ProjectName " & _
              strDBO & ",ProjectDescription " & _
              strDBO & ",ProjectEntities " & _
              strDBO & ",LastUpdate "
  End Function

  Overrides Function GetSelectFromTable() As String
    GetSelectFromTable = " FROM " & strDBO & "PAProjects "
  End Function

  Overrides Function CreateNewEntity() As PAEnt
    CreateNewEntity = New PAProject
  End Function

  Overrides Sub SetCurrentEntity(ByRef ent As PAEnt)
    currePrj_ = ent
  End Sub

  Overrides Sub BuildDomainModel()

  End Sub


End Class
