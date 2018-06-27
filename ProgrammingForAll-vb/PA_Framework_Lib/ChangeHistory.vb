Public Class ChangeHistory
  Public Loadorder As Integer  'Load order
  Dim sNo As Integer
  Dim sDateTime As DateTime
  Dim sUser As String
  Dim sTable As String
  Dim sKeyField As Integer
  Dim sChanges As String
  Dim sOperation As String
  Public Property ID() As String
    Get
      ID = sNo
    End Get
    Set(ByVal value As String)
      sNo = value
    End Set
  End Property

  Public Property DateTime() As DateTime
    Get
      DateTime = sDateTime
    End Get
    Set(ByVal value As DateTime)
      sDateTime = value
    End Set
  End Property

  Public Property User() As String
    Get
      User = sUser
    End Get
    Set(ByVal value As String)
      sUser = value
    End Set
  End Property

  Public Property Table() As String
    Get
      Table = sTable
    End Get
    Set(ByVal value As String)
      sTable = value
    End Set
  End Property

  Public Property KeyField() As String
    Get
      KeyField = sKeyField
    End Get
    Set(ByVal value As String)
      sKeyField = value
    End Set
  End Property
  Public Property Changes() As String
    Get
      Changes = sChanges
    End Get
    Set(ByVal value As String)
      sChanges = value
    End Set
  End Property

  Public Property Operation() As String
    Get
      Operation = sOperation
    End Get
    Set(ByVal value As String)
      sOperation = value
    End Set
  End Property
End Class
