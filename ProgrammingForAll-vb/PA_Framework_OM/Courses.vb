
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals


<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Courses
  Inherits PAEnts

  Overrides Sub LoadEntityItemsForThisEntity(ByRef ent As Object)
    ent.ID = "" & AppSettings.rst.Fields("ID").Value
    ent.Name = "" & AppSettings.rst.Fields("CourseName").Value
    ent.Description = "" & AppSettings.rst.Fields("Description").Value

    Call BLFunctions.BuildChildEntityObjects(cPAProjects, ent, "Students", "" & AppSettings.rst.Fields("Students").Value)
    Call BLFunctions.BuildChildEntityObjects(cPAProjects, ent, "Instructors", "" & AppSettings.rst.Fields("Instructors").Value)
    Call BLFunctions.BuildChildEntityObjects(cPAProjects, ent, "Departments", "" & AppSettings.rst.Fields("Departments").Value)
  End Sub

  Overrides Sub RecordChanges(ByVal sOperation As String, ByRef ent As Object)
    Call gRecordChanges(sOperation, "CourseName", ent.Name, "crs:")
    Call gRecordChanges(sOperation, "Description", ent.Description, "dsc:")
    Call gRecordChanges(sOperation, "Students", ent.ChildEntityString("Students"), "Stds:")
    Call gRecordChanges(sOperation, "Departments", ent.ChildEntityString("Departments"), "Dpts:")
    Call gRecordChanges(sOperation, "Instructors", ent.ChildEntityString("Instructors"), "Instr:")
  End Sub

  Overrides Function GetSelectDBItems() As String
    GetSelectDBItems = _
    strDBO & " ID " & _
    strDBO & ",CourseName " & _
    strDBO & ",Description " & _
    strDBO & ",Departments " & _
    strDBO & ",Students " & _
    strDBO & ",EntityUs " & _
    strDBO & ",Instructors " & _
    strDBO & ",LastUpdate "
  End Function

  Overrides Function GetSelectFromTable() As String
    GetSelectFromTable = " FROM " & strDBO & " Courses "
  End Function

  Overrides Function CreateNewEntity() As PAEnt
    CreateNewEntity = New Course
  End Function

  Overrides Sub SetCurrentEntity(ByRef ent As PAEnt)
    curreCrse_ = ent
  End Sub

  Overrides Sub BuildDomainModel()
  End Sub

End Class
