Option Explicit On

Public Class Course
  Public Loadorder As Integer  'Load order
  Dim sid As String
  Dim sEntityItem_1 As String
  Dim sEntityItem_2 As String
  Dim sStudents As String

  Dim mChildStudents As New Students
  Dim sChildStudentsString As String
  Dim mChildAvailableStudents As New Students

  Dim mChildPersons As New Persons
  Dim sChildPersonsString As String
  Dim mChildAvailablePersons As New Persons

  Dim mChildInstructors As New Instructors
  Dim sChildInstructorsString As String
  Dim mChildAvailableInstructors As New Instructors

  Dim mParentDepartments As New Departments
  Dim mParentAvailableDepartments As New Departments
  Dim mParentCalendars As New Calendars
  Dim mParentAvailableCalendars As New Calendars

  Public mContainer As Object
  Dim sLastUpdate As Date

  Public Property ID() As String
    Get
      ID = sid
    End Get
    Set(ByVal value As String)
      sid = value
    End Set
  End Property

  Public Property EntityItem_1() As String
    Get
      EntityItem_1 = sEntityItem_1
    End Get
    Set(ByVal value As String)
      sEntityItem_1 = value
    End Set
  End Property

  Public Property EntityItem_2() As String
    Get
      EntityItem_2 = sEntityItem_2
    End Get
    Set(ByVal value As String)
      sEntityItem_2 = value
    End Set
  End Property

  Public Property Students() As String
    Get
      Students = sStudents
    End Get
    Set(ByVal value As String)
      sStudents = value
    End Set
  End Property

  'sPH_cls_Properties:End
  Public Property Lastupdate() As Date
    Get
      Lastupdate = sLastUpdate
    End Get
    Set(ByVal value As Date)
      sLastUpdate = value
    End Set
  End Property

  'Child Entity Methods
  Public Function ChildEntities(ByVal ent As String) As Object
    Select Case ent
      Case "Student", "Students"
        ChildEntities = mChildStudents
      Case "Person", "Persons"
        ChildEntities = mChildPersons
      Case "Instructor", "Instructors"
        ChildEntities = mChildInstructors
    End Select
  End Function

  Public Function ChildEntityString(ByVal ent As String) As String
    Select Case ent
      Case "Student", "Students"
        ChildEntityString = sChildStudentsString
      Case "Person", "Persons"
        ChildEntityString = sChildPersonsString
        ' Instructors()
      Case "Instructor", "Instructors"
        ChildEntityString = sChildInstructorsString

    End Select
  End Function

  Public Function ChildEntityString(ByVal ent As String, ByVal strEnt As String)
    Select Case ent
      Case "Student", "Students"
        sChildStudentsString = strEnt
      Case "Person", "Persons"
        sChildPersonsString = strEnt
      Case "Instructor", "Instructors"
        sChildInstructorsString = strEnt
    End Select
  End Function

  Public Function BuildChildEntityObjects(ByVal objPar As String, ByVal strEnt As String) As Object
    Call BLFunctions.gBuildChildEntityObjects(Me, objPar, strEnt)
  End Function

  Public Function BuildChildEntityString(ByRef enChlds As String) As String
    Call gBuildChildEntityString(Me, enChlds)
  End Function

  Public Function ChildAvailableEntities(ByVal objPar As Object) As Object
    Select TypeName(objPar)
      Case "Student", "Students"
        ChildAvailableEntities = Me.mChildAvailableStudents
      Case "Person", "Persons"
        ChildAvailableEntities = Me.mChildAvailablePersons
      Case "Instructor", "Instructors"
        ChildAvailableEntities = Me.mChildAvailableInstructors
    End Select
  End Function

  ' Parent Entity Methods.
  Public Function ParentEntities(ByVal objPar As Object) As Object
    Select Case TypeName(objPar)
      Case "Department", "Departments"
        ParentEntities = mParentDepartments
      Case "Calendar", "Calendars"
        ParentEntities = mParentCalendars
 
    End Select
  End Function

  Public Function ParentAvailableEntities(ByVal objPar As Object) As Object
    Select Case TypeName(objPar)
      Case "Department", "Departments"
        ParentAvailableEntities = mParentAvailableDepartments
      Case "Calendar", "Calendars"
        ParentAvailableEntities = mParentAvailableCalendars
    End Select
  End Function

  Public Sub New()
    mContainer = cCourses
  End Sub

  Public ReadOnly Property PersonName() As String
    Get
      If mChildPersons.Count <> 0 Then
        PersonName = mChildPersons.Item(0).FirstName
      Else
        PersonName = ""
      End If
    End Get
  End Property

End Class
