
Option Explicit On
Imports PA_Framework_Lib

Partial Public Class OMGlobals

  Public ePrj_ As PAProject
  Public ePrjEnt_ As ProjectEntity
  Public ePrjEntItm_ As ProjectEntityItem

  Public eStd_ As Student
  Public eCrse_ As Course
  Public ePrsn_ As Person
  Public eDept_ As Department
  Public eAssoc_ As Associate
  Public eInst_ As Instructor
  Public eCal_ As Calendar
  Public eCrsStd_ As CourseStudent
  Public eAtt_ As Attendance

  Public Shared cPAProjects As New PAProjects
  Public Shared cProjectEntities As New ProjectEntities
  Public Shared cProjectEntityItems As New ProjectEntityItems

  Public Shared cStudents As New Students
  Public Shared cCourses As New Courses
  Public Shared cPersons As New Persons
  Public Shared cAttendances As New Attendances
  Public Shared cCalendars As New Calendars
  Public Shared cInstructors As New Instructors
  Public Shared cAssociates As New Associates
  Public Shared cCourseStudents As New CourseStudents
  Public Shared cDepartments As New Departments
  'Public cAppArrays As AppArrays

  'SECTION 2
  'Define the current entity for each entity in the solution. Used to track the current selected entity on the user interface
  Public Shared currePrj_ As New PAProject
  Public Shared currePrjEnt_ As New ProjectEntity
  Public Shared currePrjEntItm_ As New ProjectEntityItem

  Public Shared curreDpt_ As New Department
  Public Shared curreCrse_ As New Course
  Public Shared curreStdt_ As New Student
  Public Shared currePrsn_ As New Person
  Public Shared curreCal_ As New Calendar
  Public Shared curreAtt_ As New Attendance
  Public Shared curreInst_ As New Instructor
  Public Shared curreAssoc_ As New Associate
  Public Shared curreCrsStd_ As New CourseStudent

  Public Sub LoadEntities()
    'loads all entity collections, starting with independent ones first and then the dependent entities.
    'metadata entities
    cProjectEntityItems.Load(ePrjEntItm_)
    cProjectEntities.Load(ePrjEnt_)
    cPAProjects.Load(ePrj_)

    'pa institute entities
    cPersons.Load(ePrsn_)
    cInstructors.Load(eInst_)
    cAssociates.Load(eAssoc_)
    cStudents.Load(eStd_)
    cDepartments.Load(eDept_)
    cCourses.Load(eCrse_)
    cCalendars.Load(eCal_)
    cAttendances.Load(eAtt_)
    cCourseStudents.Load(eCrsStd_)
  End Sub

  Public Function getProjectEntityItems() As ProjectEntityItems
    Return cProjectEntityItems
  End Function

  Public Function getProjectEntities() As ProjectEntities
    Return cProjectEntities
  End Function

  Public Function getPAProjects() As PAProjects
    Return cPAProjects
  End Function

  Public Function getPersons() As Persons
    Return cPersons
  End Function

  Public Function getInstructors() As Instructors
    Return cInstructors
  End Function

  Public Function getAssociates() As Associates
    Return OMGlobals.cAssociates
  End Function

  Public Function getStudents() As Students
    Return OMGlobals.cStudents
  End Function

  Public Function getDepartments() As Departments
    Return OMGlobals.cDepartments
  End Function

  Public Function getCourses() As Courses
    Return OMGlobals.cCourses
  End Function

  Public Function getCalendars() As Calendars
    Return OMGlobals.cCalendars
  End Function

  Public Function getAttendances() As Attendances
    Return cAttendances
  End Function

  Public Function getCourseStudents() As CourseStudents
    Return OMGlobals.cCourseStudents
  End Function

  Public Function getChangeHistory()
    Return OMGlobals.cChangeHistorys
  End Function

  Public Function GetCollection(ByRef obj1 As Object) As Object
    Select Case TypeName(obj1)
      Case "Courses"
        GetCollection = cCourses
      Case Else
        GetCollection = Nothing
    End Select
  End Function
End Class