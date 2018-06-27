Option Explicit On
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class PAProject
  Inherits PAEnt

  Dim _ProjectName As String
  Dim _ProjectDescription As String
  Dim _ProjectEntities As ProjectEntities
  Dim _ProjectEntitiesString As String

  Dim mChildProjectEntities As New ProjectEntities
  Dim sChildProjectEntitiesString As String

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

  'Child Entity Methods
  Overrides Function ChildEntities(ByVal sEnt As String) As PAEnts
    ChildEntities = Nothing
    Select Case sEnt
      Case "ProjectEntity", "ProjectEntities"
        ChildEntities = mChildProjectEntities
    End Select
  End Function

  Public Function GetChildEntityString(ByVal sEnt As String) As String
    GetChildEntityString = ""
    Select Case sEnt
      Case "ProjectEntity", "ProjectEntities"
        GetChildEntityString = sChildProjectEntitiesString
    End Select
  End Function

  Public Sub SetChildEntityString(ByVal sEnt As String, ByVal strEnt As String)
    Select Case sEnt
      Case "ProjectEntity", "ProjectEntities"
        sChildProjectEntitiesString = strEnt
    End Select
  End Sub

  ' Parent Entity Methods.
  Overrides Function ParentEntities(ByRef objPar As PAEnt) As PAEnts
    ParentEntities = Nothing
    Select Case TypeName(objPar)
      '  Case "Entity1", "Entity1s":
      '    Set ParentEntities = mParentEntity1s

    End Select
  End Function

  Public Sub New()
    mContainer = cPAProjects
  End Sub

  Public Function GetItemProperty(ByVal sName As String) As String
    Select Case sName
      Case "ProjectName"
        Return Me.ProjectName
      Case "ProjectDescription"
        Return Me.ProjectDescription
    End Select
    Return String.Empty
  End Function


End Class
