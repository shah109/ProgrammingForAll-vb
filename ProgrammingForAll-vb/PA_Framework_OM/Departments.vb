Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Departments
  Inherits PAEnts

  Overrides Sub LoadEntityItemsForThisEntity(ByRef ent As Object)
    ent.ID = "" & AppSettings.rst.Fields("ID").Value
    ent.EntityItem_1 = "" & AppSettings.rst.Fields("DeptName").Value
    ent.EntityItem_2 = "" & AppSettings.rst.Fields("Description").Value
  End Sub

  Overrides Sub RecordChanges(ByVal sOperation As String, ByRef ent As Object)
    Call gRecordChanges(sOperation, "DeptName", ent.EntityItem_1, "ei1")
    Call gRecordChanges(sOperation, "Description", ent.EntityItem_2, "ei2")
    'Call gRecordChanges(sOperation, "EntityBs", ent.ChildEntityString("Courses"), "eB_str:")
  End Sub

  Overrides Function GetSelectDBItems() As String
    GetSelectDBItems = _
          strDBO & " ID " & _
          strDBO & ",DeptName " & _
          strDBO & ",Description " & _
          strDBO & ",EntityBs " & _
          strDBO & ",LastUpdate "
  End Function

  Overrides Function GetSelectFromTable() As String
    GetSelectFromTable = " FROM " & strDBO & "Departments "
  End Function

  Overrides Function CreateNewEntity() As PAEnt
    CreateNewEntity = New Department
  End Function

  Overrides Sub SetCurrentEntity(ByRef ent As PAEnt)
    curreDpt_ = ent
  End Sub

  Overrides Sub BuildDomainModel()
  End Sub

  Public Function GetCollection() As Departments
    Return Me
  End Function
End Class

