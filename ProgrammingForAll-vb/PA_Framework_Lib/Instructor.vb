Option Explicit On

Public Class Instructor
  Public Loadorder As Integer  'Load order
  Dim sid As String
  Dim sComments As String
  'Dim sStudents As String

  
  Dim mChildPersons As New Persons
  Dim sChildPersonsString As String
  Dim mChildAvailablePersons As New Persons

  Dim mParentCourses As New Courses
  Dim mParentAvailableCourses As New Courses
 

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

  Public Property Comments() As String
    Get
      Comments = sComments
    End Get
    Set(ByVal value As String)
      sComments = value
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
       Case "Person", "Persons"
        ChildEntities = mChildPersons
    End Select
  End Function

  Public Function ChildEntityString(ByVal ent As String) As String
    Select Case ent
 
      Case "Person", "Persons"
        ChildEntityString = sChildPersonsString
    End Select
  End Function

  Public Function ChildEntityString(ByVal ent As String, ByVal strEnt As String)
    Select Case ent
 
      Case "Person", "Persons"
        sChildPersonsString = strEnt
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
 
      Case "Person", "Persons"
        ChildAvailableEntities = Me.mChildAvailablePersons
    End Select
  End Function

  ' Parent Entity Methods.
  Public Function ParentEntities(ByVal objPar As Object) As Object
    Select Case TypeName(objPar)
      Case "Course", "Courses"
        ParentEntities = mParentCourses

    End Select
  End Function

  Public Function ParentAvailableEntities(ByVal objPar As Object) As Object
    Select Case TypeName(objPar)
      Case "Course", "Courses"
        ParentAvailableEntities = mParentAvailableCourses
 
    End Select
  End Function

  Public Sub New()
    mContainer = cInstructors
  End Sub

 

End Class
