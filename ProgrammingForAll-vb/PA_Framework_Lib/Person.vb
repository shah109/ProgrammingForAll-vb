Option Explicit On
Public Class Person
  'sPH_cls_DateTime:
  'Framework File Member Copied : 9/22/2011 6:18:13 PM
  'sPH_cls_DateTime:End
  'sh-Member


  Public Loadorder As Integer  'Load order
  'PH:For Each Entity Item. Create fields with clear naming conventions and with the required data types
  'sPH_cls_Decl:
  Dim sid As String
  Dim sFirstName As String
  Dim sMiddleName As String
  Dim sLastName As String
  Dim sLoginID As String
  Dim sEmail As String
  Dim sPhone As String
  Dim sAccessRight As Integer
  Dim sDateJoined As Date
  Dim sRemarks As String
  Dim sEntityItem_1 As String
  Dim sEntityItem_2 As String
  'sPH_cls_Decl:End
  'PH: End

  ''PH:For Each Child. Add two fields for each child entity this entity supports.
  'Replace all variable of Entity3 with your own <EntityName> variables
  'Dim mChildEntity3s As New Entity3s
  'Dim sChildEntity3sString As String
  ''PH: End
  'sPH_cls_ChildDecl:
  '''No Entries
  'sPH_cls_ChildDecl:End

  'Dim mChildEntityCs As New EntityCs
  'Dim sChildEntityCsString As String

  ''PH:For Each Parent, add a field for each of the parent entity this entity supports. Replace all variable of Entity1 with your own <EntityName> variables
  'Dim mParentEntity1s As New Entity1s
  ''PH: Ends
  'sPH_cls_ParentDecl:
  Dim mParentCourses As New Courses
  Dim mParentAvailableCourses As New Courses
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
  Public Property FirstName() As String
    Get
      FirstName = sFirstName
    End Get
    Set(ByVal value As String)
      sFirstName = value
    End Set
  End Property
  Public Property MiddleName() As String
    Get
      MiddleName = sMiddleName
    End Get
    Set(ByVal value As String)
      sMiddleName = value
    End Set
  End Property

  Public Property LastName() As String
    Get
      LastName = sLastName
    End Get
    Set(ByVal value As String)
      sLastName = value
    End Set
  End Property
  Public Property LoginID() As String
    Get
      LoginID = sLoginID
    End Get
    Set(ByVal value As String)
      sLoginID = value
    End Set
  End Property

  Public Property Email() As String
    Get
      Email = sEmail
    End Get
    Set(ByVal value As String)
      sEmail = value
    End Set
  End Property

  Public Property Phone() As String
    Get
      Phone = sPhone
    End Get
    Set(ByVal value As String)
      sPhone = value
    End Set
  End Property

  Public Property AccessRight() As String
    Get
      AccessRight = sAccessRight
    End Get
    Set(ByVal value As String)
      sAccessRight = value
    End Set
  End Property

  Public Property DateJoined() As Date
    Get
      DateJoined = sDateJoined
    End Get
    Set(ByVal value As Date)
      sDateJoined = value
    End Set
  End Property

  Public Property Remarks() As String
    Get
      Remarks = sRemarks
    End Get
    Set(ByVal value As String)
      sRemarks = value
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
      '''No Entries
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
      '''No Entries
      'sPH_cls_GetChildString:End
    End Select
  End Function

  Public Function ChildEntityString(ByVal ent As String, ByVal strEnt As String)
    Select Case ent
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3s", "Entity3":
      '    sChildEntity3sString = strEnt
      '  'PH: End
      'sPH_cls_LetChildString:
      '''No Entries
      'sPH_cls_LetChildString:End
    End Select
  End Function

  Public Function BuildChildEntityObjects(ByVal strPar As String, ByVal strEnt As String) As Object
    Call BLFunctions.gBuildChildEntityObjects(Me, strPar, strEnt)
  End Function

  Public Function BuildChildEntityString(ByRef enChlds As String) As String
    Call gBuildChildEntityString(Me, enChlds)
  End Function

  Public Function ChildAvailableEntities(ByVal objPar As Object) As Object
    Select TypeName(objPar)
      ''PH: For Each Parent entity, add a case statement
      'Case "Student", "Students"
       ' ChildAvailableEntities = Me.mChildAvailableStudents
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
      Case "Course", "Courses"
        ParentEntities = mParentCourses
        'sPH_cls_ParentEntities:End
    End Select
  End Function

Public Function ParentAvailableEntities(ByVal objPar As Object) As Object
    Select Case TypeName(objPar)
      ''PH: For Each Parent entity, add a case statement
      Case "Course", "Courses"
        ParentAvailableEntities = mParentAvailableCourses
        ' 'PH: End
        'sPH_cls_ParentEntities:

        'sPH_cls_ParentEntities:End
    End Select
  End Function

  Public Sub New()
    mContainer = cPersons
  End Sub

  Public Function GetItemProperty(ByVal strP As String) As String
    Select Case strP
      'sPH_cls_GetItemProperty:
      Case "FirstName"
        GetItemProperty = Me.FirstName
      Case "LastName"
        GetItemProperty = Me.LastName
        'sPH_cls_GetItemProperty:End
    End Select
  End Function



End Class
