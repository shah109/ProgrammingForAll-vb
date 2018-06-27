
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class ProjectEntityItem
  Inherits PAEnt

  Private sDBColName As String
  Private sDBColNameType As String
  Private sChildName As String
  Private sChildRelationship As String
  Private sRelationshipType As String
  Private sSQLName As String
  Private sInternalName As String
  Private sInternalNameType As String
  Private sPropertyName As String
  Private sPropertyNameType As String
  Private sULName As String
  Private sSheetDisplayOrder As String
  Private sFormControlName As String
  Private sChildListSetNeeded As String
  Private sParentListSetNeeded As String
  Private sControlReference As String
  Private sGenerateFlag As String

  Public Property ChildListSetNeeded() As String
    Get
      Return sChildListSetNeeded
    End Get
    Set(ByVal value As String)
      sChildListSetNeeded = value
    End Set
  End Property

  Public Property ChildName() As String
    Get
      Return sChildName
    End Get
    Set(ByVal value As String)
      sChildName = value
    End Set
  End Property

  Public Property ChildRelationship() As String
    Get
      Return sChildRelationship
    End Get
    Set(ByVal value As String)
      sChildRelationship = value
    End Set
  End Property

  Public Property ControlReference() As String
    Get
      Return sControlReference
    End Get
    Set(ByVal value As String)
      sControlReference = value
    End Set
  End Property

  Public Property DBColName() As String
    Get
      Return sDBColName
    End Get
    Set(ByVal value As String)
      sDBColName = value
    End Set
  End Property

  Public Property DBColNameType() As String
    Get
      Return sDBColNameType
    End Get
    Set(ByVal value As String)
      sDBColNameType = value
    End Set
  End Property

  Public Property FormControlName() As String
    Get
      Return sFormControlName
    End Get
    Set(ByVal value As String)
      sFormControlName = value
    End Set
  End Property

  Public Property GenerateFlag() As String
    Get
      Return sGenerateFlag
    End Get
    Set(ByVal value As String)
      sGenerateFlag = value
    End Set
  End Property

  Public Property InternalName() As String
    Get
      Return sInternalName
    End Get
    Set(ByVal value As String)
      sInternalName = value
    End Set
  End Property

  Public Property InternalNameType() As String
    Get
      Return sInternalNameType
    End Get
    Set(ByVal value As String)
      sInternalNameType = value
    End Set
  End Property

  Public Property ParentListSetNeeded() As String
    Get
      Return sParentListSetNeeded
    End Get
    Set(ByVal value As String)
      sParentListSetNeeded = value
    End Set
  End Property

  Public Property PropertyName() As String
    Get
      Return sPropertyName
    End Get
    Set(ByVal value As String)
      sPropertyName = value
    End Set
  End Property

  Public Property PropertyNameType() As String
    Get
      Return sPropertyNameType
    End Get
    Set(ByVal value As String)
      sPropertyNameType = value
    End Set
  End Property

  Public Property RelationshipType() As String
    Get
      Return sRelationshipType
    End Get
    Set(ByVal value As String)
      sRelationshipType = value
    End Set
  End Property

  Public Property SheetDisplayOrder() As String
    Get
      Return sSheetDisplayOrder
    End Get
    Set(ByVal value As String)
      sSheetDisplayOrder = value
    End Set
  End Property

  Public Property ULName() As String
    Get
      Return sULName
    End Get
    Set(ByVal value As String)
      sULName = value
    End Set
  End Property

  Sub New()
    mContainer = cProjectEntityItems
  End Sub

  Public Function GetItemProperty(ByVal sName As String) As String
    GetItemProperty = String.Empty
    Select Case sName
      Case "PropertyName"
        Return Me.PropertyName
      Case "PropertyNameType"
        Return Me.PropertyNameType
    End Select
  End Function
End Class
