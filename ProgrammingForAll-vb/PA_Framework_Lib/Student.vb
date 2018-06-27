Option Explicit On

Public Class Student
  'sPH_cls_DateTime:
  'gfhiii
  'sPH_cls_DateTime:End

  Public Loadorder As Integer  'Load order
  'PH:For Each Entity Item. Create fields with clear naming conventions and with the required data types
  'sPH_cls_Decl:
  Dim sid As String
  Dim sEntityItem_1 As String
  Dim sEntityItem_2 As String
  Dim sEntityBs As String
  'sPH_cls_Decl:End
  'PH: End

  ''PH:For Each Child. Add two fields for each child entity this entity supports.
  'Replace all variable of Entity3 with your own <EntityName> variables
  'Dim mChildEntity3s As New Entity3s
  'Dim sChildEntity3sString As String
  ''PH: End
  'sPH_cls_ChildDecl:
  
  'sPH_cls_ChildDecl:End

  'Dim mChildCourses As New Courses
  'Dim sChildCoursesString As String

  ''PH:For Each Parent, add a field for each of the parent entity this entity supports. Replace all variable of Entity1 with your own <EntityName> variables
  Dim mParentCourses As New Courses
  Dim mParentAvailableCourses As New Courses
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
 
    End Select
  End Function

  Public Function ChildEntityString(ByVal ent As String) As String
    Select Case ent
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3", "Entity3s":
      '    ChildEntityString = sChildEntity3sString
 
    End Select
  End Function

  Public Function ChildEntityString(ByVal ent As String, ByVal strEnt As String)
    Select Case ent
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3s", "Entity3":
      '    sChildEntity3sString = strEnt

    End Select
  End Function

  Public Function BuildChildEntityObjects(ByVal objPar As String, ByVal strEnt As String) As Object
    Call BLFunctions.gBuildChildEntityObjects(Me, objPar, strEnt)
  End Function


  Public Function BuildChildEntityString(ByRef enChlds As String) As String
    Call gBuildChildEntityString(Me, enChlds)
  End Function

  ' Parent Entity Methods.
  Public Function ParentEntities(ByVal objPar As Object)
    Select Case TypeName(objPar)
      ''PH: For Each Parent entity, add a case statement
      Case "Course", "Courses"
        ParentEntities = mParentCourses
 
    End Select
  End Function

  Public Function ParentAvailableEntities(ByVal objPar As Object) As Object
    Select Case TypeName(objPar)
      ''PH: For Each Parent entity, add a case statement
      Case "Course", "Courses"
        ParentAvailableEntities = mParentAvailableCourses
  
    End Select
  End Function

  Public Sub New()
    mContainer = cStudents
  End Sub

  'Public Function GetItemProperty(ByVal strP As String) As String
  '  Select Case strP
  '    Case "EntityItem_1"
  '      GetItemProperty = Me.EntityItem_1
  '    Case "EntityItem_2"
  '      GetItemProperty = Me.EntityItem_2
  '  End Select
  'End Function

End Class
