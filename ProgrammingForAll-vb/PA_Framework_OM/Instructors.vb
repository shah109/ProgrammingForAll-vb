'Option Explicit On
Imports System
Imports System.Collections.Generic
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Instructors
  Inherits PAEnts 'System.Collections.CollectionBase

  Overrides Sub LoadEntityItemsForThisEntity(ByRef ent As Object)
    ent.ID = "" & AppSettings.rst.Fields("ID").Value
    ent.InstructorName = "" & AppSettings.rst.Fields("InstructorName").Value
    ent.Comments = "" & AppSettings.rst.Fields("Comments").Value
    ent.DateStarted = AppSettings.rst.Fields("DateStarted").Value
    Call BLFunctions.BuildChildEntityObjects(cPAProjects, ent, "Persons", "" & AppSettings.rst.Fields("Persons").Value)

  End Sub

  Overrides Sub RecordChanges(ByVal sOperation As String, ByRef ent As Object)
    Call gRecordChanges(sOperation, "Comments", ent.Comments, "cmnts")
    Call gRecordChanges(sOperation, "InstructorName", ent.InstructorName, "InstNme")
    Call gRecordChanges(sOperation, "Persons", ent.ChildEntityString("Persons"), "Prsns")
    Call gRecordChanges(sOperation, "DateStarted", ent.DateStarted, "DteStrt")
  End Sub

  Overrides Function GetSelectDBItems() As String
    GetSelectDBItems = _
      strDBO & " ID " & _
      strDBO & ",InstructorName " & _
      strDBO & ",Persons " & _
      strDBO & ",DateStarted " & _
      strDBO & ",Comments " & _
      strDBO & ",LastUpdate "
  End Function

  Overrides Function GetSelectFromTable() As String
    GetSelectFromTable = " FROM " & strDBO & "Instructors "
  End Function

  Overrides Function CreateNewEntity() As PAEnt
    CreateNewEntity = New Instructor
  End Function

  Overrides Sub SetCurrentEntity(ByRef ent As PAEnt)
    curreInst_ = ent
  End Sub

  Overrides Sub BuildDomainModel()

  End Sub
End Class
