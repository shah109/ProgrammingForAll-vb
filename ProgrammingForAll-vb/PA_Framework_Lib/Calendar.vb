Option Explicit On

Public Class Calendar
  'sPH_cls_DateTime:
  'gfhiii
  'sPH_cls_DateTime:End

  Public Loadorder As Integer  'Load order
  'PH:For Each Entity Item. Create fields with clear naming conventions and with the required data types
  'sPH_cls_Decl:
  Dim sid As String
  Dim sLectureDate As Date
  Dim sLectureLocation As String
  Dim sComments As String
  'sPH_cls_Decl:End
  'PH: End

  ''PH:For Each Child. Add two fields for each child entity this entity supports.
  'Replace all variable of Entity3 with your own <EntityName> variables
  'Dim mChildEntity3s As New Entity3s
  'Dim sChildEntity3sString As String
  ''PH: End
  'sPH_cls_ChildDecl:
  Dim mChildCourses As New Courses
  Dim sChildCoursesString As String
  Dim mChildAvailableCourses As New Courses


  Dim mChildInstructors As New Instructors
  Dim sChildInstructorsString As String
  Dim mChildAvailableInstructors As Instructors

  Dim mChildStudents As New Students
  Dim sChildStudentsString As String
  Dim mChildAvailableStudents As New Students
  'sPH_cls_ChildDecl:End


  ''PH:For Each Parent, add a field for each of the parent entity this entity supports. Replace all variable of Entity1 with your own <EntityName> variables
  'Dim mParentEntity1s As New Entity1s
  Dim mParentAttendances As New Attendances
  Dim mParentAvailableAttendances As New Attendances
  ''PH: Ends
  'sPH_cls_ParentDecl:

  'sPH_cls_ParentDecl:End

  Public mContainer As Object
  Dim sLastUpdate As Date
  'PH:For Each Entity Item. Create a Property. Use clear naming conventions and take care to match the data types

  'sPH_cls_Properties:



  Public Property ID() As String
    Get
      ID = sid
    End Get
    Set(ByVal value As String)
      sid = value
    End Set
  End Property

  Public Property LectureDate() As Date
    Get
      LectureDate = sLectureDate
    End Get
    Set(ByVal value As Date)
      sLectureDate = value
    End Set
  End Property

  Public Property Location() As String
    Get
      Location = sLectureLocation
    End Get
    Set(ByVal value As String)
      sLectureLocation = value
    End Set
  End Property
  Public Property Comments() As String
    Get
      Comments = sComments
    End Get
    Set(ByVal value As String)
      sComments = value
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

  'sPH_cls_Properties:End


  'Child Entity Methods
  Public Function ChildEntities(ByVal sEnt As String) As Object
    Select Case sEnt
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3", "Entity3s"
      '    Set ChildEntities = mChildEntity3s
      '  'PH: End
      'sPH_cls_ChildEntities:
      Case "Courses"
        ChildEntities = mChildCourses
      Case "Students"
        ChildEntities = mChildStudents
      Case "Instructors"
        ChildEntities = mChildInstructors
        'sPH_cls_ChildEntities:End
    End Select
  End Function

  Public Function ChildAvailableEntities(ByVal objPar As Object) As Object
    Select Case TypeName(objPar)
      ''PH: For Each Parent entity, add a case statement
      Case "Student", "Students"
        ChildAvailableEntities = Me.mChildAvailableStudents
      Case "Course", "Courses"
        ChildAvailableEntities = Me.mChildAvailableCourses
      Case "Instructor", "Instructors"
        ChildAvailableEntities = Me.mChildAvailableInstructors
    End Select
  End Function

  Public Function ChildEntityString(ByVal sEnt As String) As String
    Select Case sEnt
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3", "Entity3s":
      '    ChildEntityString = sChildEntity3sString
      '  'PH: End
      'sPH_cls_GetChildString:
      Case "Courses"
        ChildEntityString = sChildCoursesString
      Case "Students"
        ChildEntityString = sChildStudentsString
      Case "Instructors"
        ChildEntityString = sChildInstructorsString
        'sPH_cls_GetChildString:End
    End Select
  End Function

  Public Function ChildEntityString(ByVal sEnt As String, ByVal strEnt As String)
    Select Case sEnt
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3s", "Entity3":
      '    sChildEntity3sString = strEnt
      '  'PH: End
      'sPH_cls_LetChildString:
      Case "Courses"
        sChildCoursesString = strEnt
      Case "Students"
        sChildStudentsString = strEnt
      Case "Instructors"
        sChildInstructorsString = strEnt
        'sPH_cls_LetChildString:End
    End Select
  End Function

  Public Function BuildChildEntityObjects(ByVal strPar As String, ByVal strEnt As String) As Object
    Call BLFunctions.gBuildChildEntityObjects(Me, strPar, strEnt)
  End Function

  Public Function BuildChildEntityString(ByRef enChlds As String) As String
    Call gBuildChildEntityString(Me, enChlds)
  End Function



  ' Parent Entity Methods.
  Public Function ParentEntities(ByVal objPar As Object) As Object
    Select Case TypeName(objPar)
      ''PH: For Each Parent entity, add a case statement
      'Case "Department", "Departments"
      'ParentEntities = mParentDepartments
      ' 'PH: End
      'sPH_cls_ParentEntities:

      'sPH_cls_ParentEntities:End
    End Select
  End Function

  Public Function ParentAvailableEntities(ByVal objPar As Object) As Object
    Select Case TypeName(objPar)
      ''PH: For Each Parent entity, add a case statement
      'Case "Attendance", "Attendances"
      'ParentAvailableEntities = mParentAvailableAttendances
      ' 'PH: End
      'sPH_cls_ParentEntities:

      'sPH_cls_ParentEntities:End
    End Select
  End Function

  Sub New()
    mContainer = cCalendars
  End Sub



End Class
