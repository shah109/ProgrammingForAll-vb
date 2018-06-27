Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class CourseStudents
  Inherits PAEnts

  Overrides Sub LoadEntityItemsForThisEntity(ByRef ent As Object)
    ent.ID = "" & AppSettings.rst.Fields("ID").Value
    ent.Comments = "" & AppSettings.rst.Fields("Comments").Value

    Call BLFunctions.BuildChildEntityObjects(cPAProjects, ent, "Students", "" & AppSettings.rst.Fields("StudentID").Value)
    Call BLFunctions.BuildChildEntityObjects(cPAProjects, ent, "Courses", "" & AppSettings.rst.Fields("CourseID").Value)
  End Sub

  Overrides Sub RecordChanges(ByVal sOperation As String, ByRef ent As Object)
    Call gRecordChanges(sOperation, "Comments", ent.Comments, "cmnts:")
    Call gRecordChanges(sOperation, "StudentID", ent.ChildEntityString("Students"), "stds:")
    Call gRecordChanges(sOperation, "CourseID", ent.ChildEntityString("Courses"), "crs:")
  End Sub

  Overrides Function GetSelectDBItems() As String
    GetSelectDBItems = _
              strDBO & " ID " & _
              strDBO & ",StudentID " & _
              strDBO & ",CourseID " & _
              strDBO & ",Comments " & _
              strDBO & ",LastUpdate "
  End Function

  Overrides Function GetSelectFromTable() As String
    GetSelectFromTable = " FROM " & strDBO & "CourseStudents "
  End Function

  Overrides Function CreateNewEntity() As PAEnt
    CreateNewEntity = New CourseStudent
  End Function

  Overrides Sub SetCurrentEntity(ByRef ent As PAEnt)
    curreCrsStd_ = ent
  End Sub

  Overrides Sub BuildDomainModel()
    'If this entity represents a join table, use the following loop to fill the associated child entities.
    'CourseStudents represents student as the child of course
    'If this is not done here, all child entities are not filled when the ui is navigated. Hence the parent entities are not shown filled unless the child entities are first navigated and fillchildentities is called by GetEntityDependencies() 
    Dim crs As Course
    Dim std As New Student()
    For Each crs In cCourses
      Call UIFunctions.FillChildEntities(cPAProjects, crs, std, "Students")
    Next crs
  End Sub
End Class
