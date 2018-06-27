Public Class Project
  'sPH_cls_Decl:
  Dim sid As String
  Dim _ProjectName As String
  Dim _ProjectDescription As String
  Dim _ProjectEntities As ProjectEntities
  Dim _ProjectEntitiesString As String
  'sPH_cls_Decl:End

  Dim mChildProjectEntities As New ProjectEntities
  Dim sChildProjectEntitiesString As String
  Public mContainer As Object
  Dim dLastUpdate As Date

  Public Loadorder As Integer  'Load order

  'sPH_cls_Properties:
  Public Property ID() As String
    Get
      ID = sid
    End Get
    Set(ByVal value As String)
      sid = value
    End Set
  End Property

  Property ProjectName() As String
    Get
      ProjectName = _ProjectName
    End Get
    Set(ByVal value As String)
      _ProjectName = value
    End Set
  End Property

  Property ProjectDescription() As String
    Get
      ProjectDescription = _ProjectDescription
    End Get
    Set(ByVal value As String)
      _ProjectDescription = value
    End Set
  End Property

  Property ProjectEntitiesString() As String
    Get
      ProjectEntitiesString = _ProjectEntitiesString
    End Get
    Set(ByVal value As String)
      _ProjectEntitiesString = value
    End Set
  End Property

  Property ChildprojectEntities() As ProjectEntities
    Get
      ChildprojectEntities = mChildProjectEntities
    End Get
    Set(ByVal value As ProjectEntities)
      mChildProjectEntities = value
    End Set
  End Property

  'sPH_cls_Properties:End
  Property Lastupdate() As Date
    Get
      Lastupdate = dLastUpdate
    End Get
    Set(ByVal value As Date)
      dLastUpdate = value
    End Set
  End Property
  'Child Entity Methods
  Public Function ChildEntities(ByVal sEnt As String) As Object
    Select Case sEnt
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3", "Entity3s"
      '    Set ChildEntities = mChildEntity3s
      '  'PH: End
      'sPH_cls_ChildEntities:
      'sPH_cls_ChildEntities
      Case "ProjectEntity", "ProjectEntities"
        ChildEntities = mChildProjectEntities
        'sPH_cls_ChildEntities:End
    End Select
  End Function

  Public Function GetChildEntityString(ByVal sEnt As String) As String
    Select Case sEnt
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3", "Entity3s":
      '    ChildEntityString = sChildEntity3sString
      '  'PH: End
      'sPH_cls_GetChildString:
      Case "ProjectEntity", "ProjectEntities"
        GetChildEntityString = sChildProjectEntitiesString
        'sPH_cls_GetChildStringgdfgsx
        'sPH_cls_GetChildString:End
    End Select
  End Function

  Public Sub SetChildEntityString(ByVal sEnt As String, ByVal strEnt As String)
    Select Case sEnt
      '  'PH: For Each child entity, add a case statement
      '  Case "Entity3s", "Entity3":
      '    sChildEntity3sString = strEnt
      '  'PH: End
      'sPH_cls_LetChildString:
      Case "ProjectEntity", "ProjectEntities"
        sChildProjectEntitiesString = strEnt
        'sPH_cls_LetChildStringytrd
        'sPH_cls_LetChildString:End
    End Select
  End Sub

  Public Sub BuildChildEntityObjects(ByVal strPar As String, ByVal strEnt As String)
    Call BLFunctions.gBuildChildEntityObjects(Me, strPar, strEnt)
  End Sub

  Public Sub BuildChildEntityString(ByRef enChlds As String)
    Call gBuildChildEntityString(Me, enChlds)
  End Sub

  ' Parent Entity Methods.
  Public Function ParentEntities(ByVal objPar As Object)
    Select Case TypeName(objPar)
      ''PH: For Each Parent entity, add a case statement
      '  Case "Entity1", "Entity1s":
      '    Set ParentEntities = mParentEntity1s
      ' 'PH: End
      'sPH_cls_ParentEntities:
      'pholder
      'sPH_cls_ParentEntities:End
    End Select
  End Function

  Public Sub New()
    mContainer = cProjects
  End Sub

  'Public Function GetItemProperty(ByVal strP As String) As String
  '  Select Case strP
  '    'sPH_cls_GetItemProperty:
  '    Case "ProjectName"
  '      GetItemProperty = Me.ProjectName
  '    Case "ProjectDescription"
  '      GetItemProperty = Me.ProjectDescription
  '      'sPH_cls_GetItemProperty:End
  '  End Select
  'End Function

End Class
