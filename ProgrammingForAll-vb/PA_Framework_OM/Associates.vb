Option Explicit On
'Imports System
'Imports System.Collections.Generic
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Associates
  Inherits PAEnts

  Overrides Sub LoadEntityItemsForThisEntity(ByRef ent As Object)
    ent.ID = "" & AppSettings.rst.Fields("ID").Value
    ent.Comments = "" & AppSettings.rst.Fields("Comments").Value

    Call BLFunctions.BuildChildEntityObjects(cPAProjects, ent, "Persons", "" & AppSettings.rst.Fields("Persons").Value)
  End Sub

  Overrides Sub RecordChanges(ByVal sOperation As String, ByRef ent As Object)
    Call gRecordChanges(sOperation, "Comments", ent.Comments, "cmnts:")
    Call gRecordChanges(sOperation, "Persons", ent.ChildEntityString("Persons"), "Prsns:")
  End Sub

  Overrides Function GetSelectDBItems() As String
    GetSelectDBItems = _
            strDBO & " ID " & _
            strDBO & ",Persons " & _
            strDBO & ",Comments " & _
            strDBO & ",LastUpdate "
  End Function

  Overrides Function GetSelectFromTable() As String
    GetSelectFromTable = " FROM " & strDBO & "Associates "
  End Function

  Overrides Function CreateNewEntity() As PAEnt
    CreateNewEntity = New Associate
  End Function

  Overrides Sub SetCurrentEntity(ByRef ent As PAEnt)
    curreAssoc_ = ent
  End Sub

  Overrides Sub BuildDomainModel()

  End Sub

End Class
