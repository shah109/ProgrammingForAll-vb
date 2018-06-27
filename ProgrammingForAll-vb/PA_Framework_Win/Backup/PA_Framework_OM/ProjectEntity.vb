Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class ProjectEntity
  Inherits PAEnt
  'im nID As Integer
  'Public LoadOrder As Integer
  Private sEntityName As String
  Private sEntityCollectionName As String
  Private sEntityShortName As String
  Private sEntityDBTableName As String
  Private sEntityItems As String
  Private bGenerateFlag As Boolean
  Private sInputFile As String
  Private sFilesToGenerate As String
  Private sExcelFormName As String
  Private sExcelSheetName As String
  Private sDateTimeCodeGenerated As DateTime

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

  Public Property ExcelSheetName() As String
    Get
      ExcelSheetName = sExcelSheetName
    End Get
    Set(ByVal value As String)
      sExcelSheetName = value
    End Set
  End Property

  Public Property ExcelFormName() As String
    Get
      ExcelFormName = sExcelFormName
    End Get
    Set(ByVal value As String)
      sExcelFormName = value
    End Set
  End Property

  Public Property DateTimeCodeGenerated() As Date
    Get
      DateTimeCodeGenerated = sDateTimeCodeGenerated
    End Get
    Set(ByVal value As Date)
      sDateTimeCodeGenerated = value
    End Set
  End Property

  Sub New()
    Me.mContainer = cProjectEntities
  End Sub

  Public Function GetItemProperty(ByVal sName As String) As String
    GetItemProperty = String.Empty

    Select Case sName
      Case "EntityName"
        Return Me.EntityName
      Case "EntityShortName"
        Return Me.EntityShortName
    End Select
  End Function
End Class
