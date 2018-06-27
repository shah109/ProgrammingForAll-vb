Option Explicit On

Public Class Department
  'sPH_cls_DateTime:
  'gfhiii
  'sPH_cls_DateTime:End

  Public Loadorder As Integer  'Load order
  'PH:For Each Entity Item. Create fields with clear naming conventions and with the required data types
  'sPH_cls_Decl:
  Dim sid As String
  Dim sEntityItem_1 As String
  Dim sEntityItem_2 As String
  Dim sCourses As String
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
  'sPH_cls_ChildDecl:End

  'Dim mChildEntityCs As New EntityCs
  'Dim sChildEntityCsString As String

  ''PH:For Each Parent, add a field for each of the parent entity this entity supports. Replace all variable of Entity1 with your own <EntityName> variables
  'Dim mParentEntity1s As New Entity1s
Dim mChildAvailableCourses as new Courses
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

  Public Property Courses() As String
    Get
      Courses = sCourses
    End Get
    Set(ByVal value As String)
      sCourses = value
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
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3", "Entity3s"
      '    Set ChildEntities = mChildEntity3s
      '  'PH: End
      'sPH_cls_ChildEntities:
      Case "Course", "Courses"
        ChildEntities = mChildCourses
        'sPH_cls_ChildEntities:End
    End Select
  End Function

  Public Function ChildEntityString(ByVal ent As String) As String
    Select Case ent
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3", "Entity3s":
      '    ChildEntityString = sChildEntity3sString
      '  'PH: End
      'sPH_cls_GetChildString:
      Case "Course", "Courses"
        ChildEntityString = sChildCoursesString
        'sPH_cls_GetChildString:End
    End Select
  End Function

  Public Function ChildEntityString(ByVal ent As String, ByVal strEnt As String)
    Select Case ent
      Case "Course", "Courses"
        sChildCoursesString = strEnt
    End Select
  End Function

  Public Sub BuildChildEntityObjects(ByVal objPar As String, ByVal strEnt As String)
    Call BLFunctions.gBuildChildEntityObjects(Me, objPar, strEnt)
  End Sub


  Public Function BuildChildEntityString(ByRef enChlds As String) As String
    Call gBuildChildEntityString(Me, enChlds)
  End Function

  Public Function ChildAvailableEntities(ByVal objPar As Object) As Object
    Select TypeName(objPar)
      Case "Course", "Courses"
        ChildAvailableEntities = Me.mChildAvailableCourses
    End Select
  End Function

  ' Parent Entity Methods.
  Public Function ParentEntities(ByVal objPar As Object)
    Select Case TypeName(objPar)
      ''PH: For Each Parent entity, add a case statement
      '  Case "Entity1", "Entity1s":
      '    Set ParentEntities = mParentEntity1s
      ' 'PH: End
      'sPH_cls_ParentEntities:

      'sPH_cls_ParentEntities:End
    End Select
  End Function

  Public Sub New()
    mContainer = cDepartments
  End Sub

  Public Function GetItemProperty(ByVal strP As String) As String
    Select Case strP
      Case "EntityItem_1"
        GetItemProperty = Me.EntityItem_1
      Case "EntityItem_2"
        GetItemProperty = Me.EntityItem_2
    End Select
  End Function

End Class
