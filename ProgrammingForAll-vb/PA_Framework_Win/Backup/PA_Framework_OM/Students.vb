Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices
<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Students
  Inherits PAEnts

  Overrides Sub LoadEntityItemsForThisEntity(ByRef ent As Object)
    ent.ID = "" & AppSettings.rst.Fields("ID").Value
    ent.EntityItem_1 = "" & AppSettings.rst.Fields("EntityItem_1").Value
    ent.EntityItem_2 = "" & AppSettings.rst.Fields("EntityItem_2").Value

    Call BLFunctions.BuildChildEntityObjects(cPAProjects, ent, "Persons", "" & AppSettings.rst.Fields("Persons").Value)
  End Sub

  Overrides Sub RecordChanges(ByVal sOperation As String, ByRef ent As Object)
    Call gRecordChanges(sOperation, "EntityItem_1", ent.EntityItem_1, "ei1:")
    Call gRecordChanges(sOperation, "EntityItem_2", ent.EntityItem_2, "ei2:")
    Call gRecordChanges(sOperation, "Persons", ent.ChildEntityString("Persons"), "Prsns:")
  End Sub

  Overrides Function GetSelectDBItems() As String
    GetSelectDBItems = _
              strDBO & " ID " & _
              strDBO & ", EntityItem_1 " & _
              strDBO & ", EntityItem_2 " & _
              strDBO & ", Persons " & _
              strDBO & ", LastUpdate "
  End Function

  Overrides Function GetSelectFromTable() As String
    GetSelectFromTable = " FROM " & strDBO & "Students "
  End Function

  Overrides Function CreateNewEntity() As PAEnt
    CreateNewEntity = New Student
  End Function

  Overrides Sub SetCurrentEntity(ByRef ent As PAEnt)
    curreStdt_ = ent
  End Sub

  Overrides Sub BuildDomainModel()

  End Sub

End Class
