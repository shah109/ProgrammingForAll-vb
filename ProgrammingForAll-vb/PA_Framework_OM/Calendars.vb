Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Calendars
  Inherits PAEnts

  Overrides Sub LoadEntityItemsForThisEntity(ByRef ent As Object)
    ent.ID = "" & AppSettings.rst.Fields("ID").Value
    ent.LectureDate = "" & AppSettings.rst.Fields("LectureDate").Value
    ent.Location = "" & AppSettings.rst.Fields("LectureLocation").Value
    ent.Comments = "" & AppSettings.rst.Fields("Comments").Value

    Call BLFunctions.BuildChildEntityObjects(cPAProjects, ent, "Courses", "" & AppSettings.rst.Fields("CourseID").Value)
    Call BLFunctions.BuildChildEntityObjects(cPAProjects, ent, "Instructors", "" & AppSettings.rst.Fields("InstructorID").Value)
  End Sub

  Overrides Sub RecordChanges(ByVal sOperation As String, ByRef ent As Object)
    Call gRecordChanges(sOperation, "LectureDate", ent.LectureDate, "lectDte:")
    Call gRecordChanges(sOperation, "LectureLocation", ent.Location, "loc:")
    Call gRecordChanges(sOperation, "Comments", ent.Comments, "cmnts:")
    Call gRecordChanges(sOperation, "CourseID", ent.ChildEntityString("Courses"), "crss")
    Call gRecordChanges(sOperation, "InstructorID", ent.ChildEntityString("Instructors"), "Inst")
  End Sub

  Overrides Function GetSelectDBItems() As String
    GetSelectDBItems = _
            strDBO & " ID " & _
            strDBO & ",CourseID " & _
            strDBO & ",InstructorID " & _
            strDBO & ",LectureDate " & _
            strDBO & ",LectureLocation " & _
            strDBO & ",Comments " & _
            strDBO & ",LastUpdate "
  End Function

  Overrides Function GetSelectFromTable() As String
    Return " FROM " & strDBO & "Calendars "
  End Function

  Overrides Function CreateNewEntity() As PAEnt
    CreateNewEntity = New Calendar
  End Function

  Overrides Sub SetCurrentEntity(ByRef ent As PAEnt)
    curreCal_ = ent
  End Sub

  Overrides Sub BuildDomainModel()

  End Sub

  Public Shared Function CompareDate(ByRef Calendar1 As Object, ByRef Calendar2 As Object) As Integer
    'Return CType(Calendar1, Calendar).LectureDate.com > CType(Calendar2, Calendar).LectureDate
    ' MsgBox(DateTime.Compare(Calendar1.LectureDate, Calendar2.LectureDate))

    Return DateTime.Compare(Calendar1.LectureDate, Calendar2.LectureDate)
  End Function

End Class
