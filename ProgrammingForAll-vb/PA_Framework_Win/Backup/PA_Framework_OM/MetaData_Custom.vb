Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Partial Public Class PAProjects

  Public Function CreateObjectFromString(ByVal sPropertyName As String) As Object
    'Creates object from child property name 
    CreateObjectFromString = Nothing
    Select Case sPropertyName
      'For each entity, add a case statement
      Case "Course", "Courses"
        CreateObjectFromString = New Course
      Case "Student", "Students"
        CreateObjectFromString = New Student
      Case "Person", "Persons"
        CreateObjectFromString = New Person
      Case "CalendarItems", "Calendar"
        CreateObjectFromString = New Calendar
      Case "Attendances", "Attendance"
        CreateObjectFromString = New Attendance
      Case "Instructor", "Instructors"
        CreateObjectFromString = New Instructor
      Case "Department", "Departments"
        CreateObjectFromString = New Department
      Case "CourseStudent", "CourseStudents"
        CreateObjectFromString = New CourseStudent
      Case "Associate", "Associates"
        CreateObjectFromString = New Associate
      Case "PAProject", "PAProjects"
        CreateObjectFromString = New PAProject
      Case "ProjectEntity", "ProjectEntities"
        CreateObjectFromString = New ProjectEntity
      Case "ProjectEntityItem", "ProjectEntityItems"
        CreateObjectFromString = New ProjectEntityItem
    End Select
    If CreateObjectFromString Is Nothing Then
      '  MsgBox("Error: CreateObjectFromString does not seem to contain the case for " & sPropertyName)
      Call AppSettings.WriteToErrorLog("Error: CreateObjectFromString does not seem to contain the case for " & sPropertyName)
    End If
  End Function

  Public Sub DBLoadWithUpdateLogItems(ByVal UpLIs As UpdateLogItems)
    'Input UpLIs contains the collection of all updatelog items since last update.
    Dim ul As UpdateLogItem
    For Each ul In UpLIs
      Select Case ul.sTableID
        Case "Attendance"
          Call cAttendances.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
        Case "Course"
          Call cCourses.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
        Case "Department"
          Call cDepartments.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
        Case "Student"
          Call cStudents.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
        Case "Calendar"
          Call cCalendars.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
        Case "Instructor"
          Call cInstructors.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
        Case "Person"
          Call cPersons.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
        Case "CourseStudent"
          Call cCourseStudents.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
        Case "cPAProject"
          Call cPAProjects.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
        Case "ProjectEntity"
          Call cProjectEntities.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)
        Case "ProjectEntityItem"
          Call cProjectEntityItems.LoadSingleEntity(ul.sKeyFieldNumber, ul.sOperation)





        Case Else
          MsgBox("Error in DBLoadWithUpdateLogItems(); UpdatelogItem '" & ul.sTableID & "' not present")
          Call AppSettings.WriteToErrorLog("Error from DBLoadWithUpdateLogItems(); UpdatelogItem '" & ul.sTableID & "' not present")
      End Select
    Next ul
  End Sub

End Class

