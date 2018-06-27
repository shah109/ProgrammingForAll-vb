
Public Class ProjectEntity
  Dim nID As Integer
  Public LoadOrder As Integer
  Private sEntityName As String
  Private sEntityCollectionName As String
  Private sEntityShortName As String
  Private sEntityDBTableName As String
  Private sEntityItems As String
  Private bGenerateFlag As Boolean
  Private sInputFile As String
  Private sFilesToGenerate As String
  Private sDateTimeCodeGenerated As DateTime
  Dim dLastUpdate As DateTime
  Public mContainer As ProjectEntities
  Public Property ID() As Integer
    Get
      ID = nID
    End Get
    Set(ByVal value As Integer)
      nID = value
    End Set
  End Property

  Property EntityName() As String
    Get
      EntityName = sEntityName
    End Get
    Set(ByVal value As String)
      sEntityName = value
    End Set
  End Property

  Public Property EntityCollectionName() As String
    Get
      EntityCollectionName = sEntityCollectionName
    End Get
    Set(ByVal value As String)
      sEntityCollectionName = value
    End Set
  End Property

  Public Property EntityShortName() As String
    Get
      EntityShortName = sEntityShortName
    End Get
    Set(ByVal value As String)
      sEntityShortName = value
    End Set
  End Property

  Public Property EntityDBTableName() As String
    Get
      EntityDBTableName = sEntityDBTableName
    End Get
    Set(ByVal value As String)
      sEntityDBTableName = value
    End Set
  End Property

  Public Property EntityItems() As String
    Get
      EntityItems = sEntityItems
    End Get
    Set(ByVal value As String)
      sEntityItems = value
    End Set
  End Property

  Public Property GenerateFlag() As Boolean
    Get
      GenerateFlag = bGenerateFlag
    End Get
    Set(ByVal value As Boolean)
      bGenerateFlag = value
    End Set
  End Property

  Public Property InputFile() As String
    Get
      InputFile = sInputFile
    End Get
    Set(ByVal value As String)
      sInputFile = value
    End Set
  End Property

  Public Property FilesToGenerate() As String
    Get
      FilesToGenerate = sFilesToGenerate
    End Get
    Set(ByVal value As String)
      sFilesToGenerate = value
    End Set
  End Property

  Public Property DateTimeCodeGenerated() As DateTime
    Get
      DateTimeCodeGenerated = sDateTimeCodeGenerated
    End Get
    Set(ByVal value As DateTime)
      sDateTimeCodeGenerated = value
    End Set
  End Property

  Public Property LastUpdate() As DateTime
    Get
      LastUpdate = dLastUpdate
    End Get
    Set(ByVal value As DateTime)
      dLastUpdate = value
    End Set
  End Property

  Sub New()
    Me.mContainer = cProjectEntities
  End Sub

End Class
