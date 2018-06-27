Public Class ProjectEntityItem

  Private sID As String
  Private nLoadorder As Integer
  Private sEntityItem_1 As String
  Private sEntityItem_2 As String

  Private sItemDBName As String
  Private sItemDBNameType As String
  Private sItemChildName As String
  Private sItemChildRelationship As String
  Private sItemRelationshipType As String
  Private sItemSQLName As String
  Private sItemInternalName As String
  Private sItemInternalNameType As String
  Private sItemPropertyName As String
  Private sItemPropertyNameType As String
  Private sItemULName As String
  Private sItemSheetDisplayOrder As String
  Private sItemFormControlName As String
  Private sItemChildListSetNeeded As String
  Private sItemParentListSetNeeded As String
  Private sItemControlReference As String
  Private sItemGenerateFlag As String
  Private dLastUpdate As DateTime
  Public mContainer As ProjectEntityItems

  Public Property ID() As String
    Get
      Return sID
    End Get
    Set(ByVal value As String)
      sID = value
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

  Public Property Loadorder() As Integer
    Get
      Return nLoadorder
    End Get
    Set(ByVal value As Integer)
      nLoadorder = value
    End Set
  End Property

  Public Property ItemChildListSetNeeded() As String
    Get
      Return sItemChildListSetNeeded
    End Get
    Set(ByVal value As String)
      sItemChildListSetNeeded = value
    End Set
  End Property

  Public Property ItemChildName() As String
    Get
      Return sItemChildName
    End Get
    Set(ByVal value As String)
      sItemChildName = value
    End Set
  End Property

  Public Property ItemChildRelationship() As String
    Get
      Return sItemChildRelationship
    End Get
    Set(ByVal value As String)
      sItemChildRelationship = value
    End Set
  End Property

  Public Property ItemControlReference() As String
    Get
      Return sItemControlReference
    End Get
    Set(ByVal value As String)
      sItemControlReference = value
    End Set
  End Property

  Public Property Lastupdate() As DateTime
    Get
      Return dLastUpdate
    End Get
    Set(ByVal value As DateTime)
      dLastUpdate = value
    End Set
  End Property

  Public Property ItemDBName() As String
    Get
      Return sItemDBName
    End Get
    Set(ByVal value As String)
      sItemDBName = value
    End Set
  End Property

  Public Property ItemDBNameType() As String
    Get
      Return sItemDBNameType
    End Get
    Set(ByVal value As String)
      sItemDBNameType = value
    End Set
  End Property

  Public Property ItemFormControlName() As String
    Get
      Return sItemFormControlName
    End Get
    Set(ByVal value As String)
      sItemFormControlName = value
    End Set
  End Property

  Public Property ItemGenerateFlag() As String
    Get
      Return sItemGenerateFlag
    End Get
    Set(ByVal value As String)
      sItemGenerateFlag = value
    End Set
  End Property

  Public Property ItemInternalName() As String
    Get
      Return sItemInternalName
    End Get
    Set(ByVal value As String)
      sItemInternalName = value
    End Set
  End Property

  Public Property ItemInternalNameType() As String
    Get
      Return sItemInternalNameType
    End Get
    Set(ByVal value As String)
      sItemInternalNameType = value
    End Set
  End Property

  Public Property ItemParentListSetNeeded() As String
    Get
      Return sItemParentListSetNeeded
    End Get
    Set(ByVal value As String)
      sItemParentListSetNeeded = value
    End Set
  End Property

  Public Property ItemPropertyName() As String
    Get
      Return sItemPropertyName
    End Get
    Set(ByVal value As String)
      sItemPropertyName = value
    End Set
  End Property

  Public Property ItemPropertyNameType() As String
    Get
      Return sItemPropertyNameType
    End Get
    Set(ByVal value As String)
      sItemPropertyNameType = value
    End Set
  End Property

  Public Property ItemRelationshipType() As String
    Get
      Return sItemRelationshipType
    End Get
    Set(ByVal value As String)
      sItemRelationshipType = value
    End Set
  End Property

  Public Property ItemSheetDisplayOrder() As String
    Get
      Return sItemSheetDisplayOrder
    End Get
    Set(ByVal value As String)
      sItemSheetDisplayOrder = value
    End Set
  End Property

  Public Property ItemSQLName() As String
    Get
      Return sItemSQLName
    End Get
    Set(ByVal value As String)
      sItemSQLName = value
    End Set
  End Property

  Public Property ItemULName() As String
    Get
      Return sItemULName
    End Get
    Set(ByVal value As String)
      sItemULName = value
    End Set
  End Property
  Sub New()
    mContainer = cProjectEntityItems
  End Sub
End Class
