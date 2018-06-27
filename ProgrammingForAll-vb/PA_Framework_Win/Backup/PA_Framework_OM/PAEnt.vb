Public Class PAEnt
  Dim sLoadorder As Integer  'Load order
  'Public Loadorder1 As Integer
  Dim sID As String

  Public IsDirty As Boolean
  Public mContainer As PAEnts
  Protected sLastUpdate As Date

  'SECTION 1: Add three declarations for each child entities of your entity as follows:
  Public mChildDepartments As New Departments
  Protected sChildDepartmentsString As String
  Public mAvChildDepartments As New Departments
  '2
  Public mChildCourses As New Courses
  Protected sChildCoursesString As String = ""
  Public mAvChildCourses As New Courses
  '3
  Public mChildStudents As New Students
  Protected sChildStudentsString As String = ""
  Public mAvChildStudents As New Students

  '4
  Public mChildCourseStudents As New CourseStudents
  Protected sChildCourseStudentsString As String = ""
  Public mAvChildCourseStudents As New CourseStudents

  '5
  Public mChildCalendarItems As New Calendars
  Protected sChildCalendarItemsString As String = ""
  Public mAvChildCalendarItems As New Calendars

  '6
  Public mChildAttendances As New Attendances
  Protected sChildAttendancesString As String = ""
  Public mAvChildAttendances As New Attendances

  '7
  Public mChildPersons As New Persons
  Protected sChildPersonsString As String = ""
  Public mAvChildPersons As New Persons

  '8
  Public mChildInstructors As New Instructors
  Protected sChildInstructorsString As String
  Public mAvChildInstructors As New Instructors

  '9
  Public mChildAssociates As New Associates
  Protected sChildAssociatesString As String = ""
  Public mAvChildAssociates As New Associates

  Public mChildProjectEnttities As New ProjectEntities
  Protected sChildProjectEnttitiesString As String
  Public mAvChildProjectEnttities As New ProjectEntities

  Public mChildProjectEntityItems As New ProjectEntityItems
  Protected sChildProjectEntityItemsString As String
  Public mAvChildProjectEntityItems As New ProjectEntityItems

  'SECTION 2: Add two declarations for each Parent entities of your entity as follows:
  ' 1
  Protected mParentDepartments As New Departments
  Protected mAvParentDepartments As New Departments
  '2
  Protected mParentCourses As New Courses
  Protected mAvParentCourses As New Courses
  '3
  Protected mParentStudents As New Students
  Protected mAvParentStudents As New Students
  '4
  Protected mParentCourseStudents As New CourseStudents
  Protected mAvParentCourseStudents As New CourseStudents
  '5
  Protected mParentCalendarItems As New Calendars
  Protected mAvParentCalendarItems As New Calendars
  '6
  Protected mParentAttendances As New Attendances
  Protected mAvParentAttendances As New Attendances
  '7
  Protected mParentPersons As New Persons
  Protected mAvParentPersons As New Persons
  '8
  Protected mParentInstructors As New Instructors
  Protected mAvParentInstructors As New Instructors
  '9
  Protected mParentAssociates As New Associates
  Protected mAvParentAssociates As New Associates

  Protected mParentPAProjects As New PAProjects
  Protected mAvParentPAProjects As New PAProjects

  Protected mParentProjectEntities As New ProjectEntities
  Protected mAvParentProjectEntities As New ProjectEntities

  Public Property LoadOrder() As Integer
    Get
      LoadOrder = sLoadorder
    End Get
    Set(ByVal value As Integer)
      sLoadorder = value
    End Set
  End Property

  Public Property ID() As String
    Get
      ID = sID
    End Get
    Set(ByVal value As String)
      sID = value
    End Set
  End Property

  Public Property Lastupdate() As Date
    Get
      Lastupdate = sLastUpdate
    End Get
    Set(ByVal value As Date)
      sLastUpdate = value
    End Set
  End Property

  Public Overridable Function ChildEntityString(ByVal ent As String) As String
    ChildEntityString = ""
    Select Case ent
      Case "Department", "Departments"
        ChildEntityString = sChildDepartmentsString
      Case "Course", "Courses"
        ChildEntityString = sChildCoursesString
      Case "Student", "Students"
        ChildEntityString = sChildStudentsString
      Case "CourseStudent", "CourseStudents"
        ChildEntityString = sChildCourseStudentsString
      Case "Calendar", "CalendarItems"
        ChildEntityString = sChildCalendarItemsString
      Case "Attendance", "Attendances"
        ChildEntityString = sChildAttendancesString
      Case "Person", "Persons"
        ChildEntityString = sChildPersonsString
      Case "Instructor", "Instructors"
        ChildEntityString = sChildInstructorsString
      Case "Associate", "Associates"
        ChildEntityString = sChildAssociatesString

      Case "ProjectEntity", "ProjectEntities"
        ChildEntityString = sChildProjectEnttitiesString
      Case "ProjectEntityItem", "ProjectEntityItems"
        ChildEntityString = sChildProjectEntityItemsString

    End Select
  End Function

  Public Overridable Sub ChildEntityString(ByVal ent As String, ByVal strEnt As String)
    Select Case ent
      Case "Department", "Departments"
        sChildDepartmentsString = strEnt
      Case "Course", "Courses"
        sChildCoursesString = strEnt
      Case "Student", "Students"
        sChildStudentsString = strEnt
      Case "CourseStudent", "CourseStudents"
        sChildCourseStudentsString = strEnt
      Case "Calendar", "CalendarItems"
        sChildCalendarItemsString = strEnt
      Case "Attendance", "Attendances"
        sChildAttendancesString = strEnt
      Case "Person", "Persons"
        sChildPersonsString = strEnt
      Case "Instructor", "Instructors"
        sChildInstructorsString = strEnt
      Case "Associate", "Associates"
        sChildAssociatesString = strEnt

      Case "ProjectEntity", "ProjectEntities"
        sChildProjectEnttitiesString = strEnt
      Case "ProjectEntityItem", "ProjectEntityItems"
        sChildProjectEntityItemsString = strEnt

    End Select
  End Sub

  Public Overridable Function ChildEntities(ByVal Ent As String) As PAEnts
    ChildEntities = Nothing
    Select Case Ent
      Case "Department", "Departments"
        ChildEntities = mChildDepartments
      Case "Course", "Courses"
        ChildEntities = mChildCourses
      Case "Student", "Students"
        ChildEntities = mChildStudents
      Case "CourseStudent", "CourseStudents"
        ChildEntities = mChildCourseStudents
      Case "Calendar", "CalendarItems"
        ChildEntities = mChildCalendarItems
      Case "Attendance", "Attendances"
        ChildEntities = mChildAttendances
      Case "Person", "Persons"
        ChildEntities = mChildPersons
      Case "Instructor", "Instructors"
        ChildEntities = mChildInstructors
      Case "Associate", "Associates"
        ChildEntities = mChildAssociates

      Case "ProjectEntity", "ProjectEntities"
        ChildEntities = mChildProjectEnttities
      Case "ProjectEntityItem", "ProjectEntityItems"
        ChildEntities = mChildProjectEntityItems

    End Select
  End Function

  Public Overridable Function AvailableChildEntities(ByRef objChld As PAEnt) As PAEnts
    AvailableChildEntities = Nothing
    Select Case TypeName(objChld)
      Case "Department", "Departments"
        AvailableChildEntities = mAvChildDepartments
      Case "Course", "Courses"
        AvailableChildEntities = mAvChildCourses
      Case "Student", "Students"
        AvailableChildEntities = mAvChildStudents
      Case "CourseStudent", "CourseStudents"
        AvailableChildEntities = mAvChildCourseStudents
      Case "Calendar", "Calendars"
        AvailableChildEntities = mAvChildCalendarItems
      Case "Attendance", "Attendances"
        AvailableChildEntities = mAvChildAttendances
      Case "Person", "Persons"
        AvailableChildEntities = mAvChildPersons
      Case "Instructor", "Instructors"
        AvailableChildEntities = mAvChildInstructors
      Case "Associate", "Associates"
        AvailableChildEntities = mAvChildAssociates

      Case "ProjectEntity", "ProjectEntities"
        AvailableChildEntities = mAvChildProjectEnttities
      Case "ProjectEntityItem", "ProjectEntityItems"
        AvailableChildEntities = mAvChildProjectEntityItems

    End Select
  End Function

  Public Overridable Function ParentEntities(ByRef objPar As PAEnt) As PAEnts
    ParentEntities = Nothing
    Select Case TypeName(objPar)
      Case "Department", "Departments"
        ParentEntities = mParentDepartments
      Case "Course", "Courses"
        ParentEntities = mParentCourses
      Case "Student", "Students"
        ParentEntities = mParentStudents
      Case "CourseStudent", "CourseStudents"
        ParentEntities = mParentCourseStudents
      Case "Calendar", "Calendars"
        ParentEntities = mParentCalendarItems
      Case "Attendance", "Attendances"
        ParentEntities = mParentAttendances
      Case "Person", "Persons"
        ParentEntities = mParentPersons
      Case "Instructor", "Instructors"
        ParentEntities = mParentInstructors
      Case "Associate", "Associates"
        ParentEntities = mParentAssociates

      Case "ProjectEntity", "ProjectEntities"
        ParentEntities = mParentProjectEntities
      Case "PAProject", "PAProjects"
        ParentEntities = mParentPAProjects

    End Select
  End Function

  Public Overridable Function AvailableParentEntities(ByRef objPar As PAEnt) As PAEnts
    AvailableParentEntities = Nothing
    Select Case TypeName(objPar)
      Case "Department", "Departments"
        AvailableParentEntities = mAvParentDepartments
      Case "Course", "Courses"
        AvailableParentEntities = mAvParentCourses
      Case "Student", "Students"
        AvailableParentEntities = mAvParentStudents
      Case "CourseStudent", "CourseStudents"
        AvailableParentEntities = mParentCourseStudents
      Case "Calendar", "Calendars"
        AvailableParentEntities = mAvParentCalendarItems
      Case "Attendance", "Attendances"
        AvailableParentEntities = mAvParentAttendances
      Case "Person", "Persons"
        AvailableParentEntities = mAvParentPersons
      Case "Instructor", "Instructors"
        AvailableParentEntities = mAvParentInstructors
      Case "Associate", "Associates"
        AvailableParentEntities = mAvParentAssociates

      Case "ProjectEntity", "ProjectEntities"
        AvailableParentEntities = mAvParentProjectEntities
      Case "PAProject", "PAProjects"
        AvailableParentEntities = mAvParentPAProjects

    End Select
  End Function

  'Public Function Clone() As Object Implements System.ICloneable.Clone
  '  Return New PAEnt With {.ID = ID}
  'End Function

End Class
