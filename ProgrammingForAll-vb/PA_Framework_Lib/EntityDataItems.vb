Option Explicit On
Public Class EntityDataItem
  'sPH_cls_DateTime:
  'gfh
  'sPH_cls_DateTime:End

  Public Loadorder As Integer  'Load order

  Dim sNo As String
  Public EntityName As String
  Public Associations As String
  Public ShortName As String
  Public FormName As String
  Public Sheetname As String

  Public Property EdI_ID() As String
    Get
      EdI_ID = sNo
    End Get
    Set(ByVal value As String)
      sNo = value
    End Set
  End Property
End Class

Public Class EntityDataItems
  Inherits System.Collections.CollectionBase
  'sPH_cls_DateTime:
  'gfh
  'sPH_cls_DateTime:End

  '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  'EntityDataItems Class Collection Source -
  '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  'Private collInt As Collection
  Dim strChanges As String
  Dim EntData As EntityDataItem
  'Dim bChangesMade As Boolean

  'Public Sub New()
  '    collInt = New Collection
  'End Sub

  'Public Function Count() As Integer
  '    Count = Me.List.Count
  'End Function

  Public ReadOnly Property Items() As Collection
    Get
      Return Me.List
    End Get
  End Property

  Public ReadOnly Property item(ByVal val As Object) As EntityDataItem
    Get
      item = Me.List.Item(val)
    End Get
  End Property

  Public Sub Remove(ByRef val As EntityDataItem)
    Me.List.Remove(val)
  End Sub

  Public Sub Add(ByRef val As EntityDataItem)
    Me.List.Add(val)

  End Sub
    Public Function Load()
        Dim nLoadorder As Integer
        Dim eCrse_ As EntityDataItem
        gStrSqlCall = " Select " & _
                         strDBO & "ID as sNo, " & _
                         strDBO & "EntityName, " & _
                         strDBO & "Associations, " & _
                         strDBO & "ShortName, " & _
                         strDBO & "FormName, " & _
                         strDBO & "SheetName " & _
                         "FROM " & strDBO & "EntityDataItems"

        ''''''''''''''''''''''''''''''''''''''''''''''''''

    cnn.Open(MGlobals.GetAppConnString)
    rst.CursorLocation = adUseClient
    rst.Open(gStrSqlCall, cnn, ADODB.CursorTypeEnum.adOpenStatic, 1)
    Me.Clear()
    rst.MoveFirst()
    nloadorder = 1
    Do While Not rst.EOF
      eCrse_ = New EntityDataItem

      eCrse_.EdI_ID = "" & rst.Fields("sNo").Value
      eCrse_.EntityName = "" & rst.Fields("EntityName").Value
      eCrse_.Associations = "" & rst.Fields("Associations").Value
      eCrse_.ShortName = "" & rst.Fields("ShortName").Value
      eCrse_.FormName = "" & rst.Fields("FormName").Value
      eCrse_.Sheetname = "" & rst.Fields("SheetName").Value

      Call cEntityDataItems.Add(eCrse_)
      nLoadorder = nLoadorder + 1
      rst.MoveNext()
    Loop
    rst.Close()
    cnn.Close()

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Dim sqlEntMetaData As String
    'cnn = CreateObject("ADODB.Connection")
    'rst = CreateObject("ADODB.Recordset")

    'DBStrings.gConnectionString = MGlobals.GetAppConnString
    'DBStrings.cnn1.ConnectionString = DBStrings.gConnectionString
    '        DBStrings.cmd1.CommandText = sqlEntMetaData
    '        DBStrings.cmd1.CommandType = CommandType.Text
    '        DBStrings.cmd1.Connection = DBStrings.cnn1
    '        DBStrings.cnn1.Open()
    '        DBStrings.dr1 = DBStrings.cmd1.ExecuteReader
    '        'cnn = New ADODB.Connection
    '        'MGlobals.GetAppConnString()
    '        'cnn.Open(DBStrings.gConnectionString)
    '        'rst.CursorLocation = adUseClient
    '        'rst.Open(sqlEntMetaData, cnn, _
    '        '         ADODB.CursorTypeEnum.adOpenStatic, 1)
    '        If Not DBStrings.dr1.HasRows Then Exit Function
    '        'rst.MoveFirst()
    '        cEntityDataItems.Clear()

    '        Do While DBStrings.dr1.Read
    '            EntData = New EntityDataItem

    '            EntData.EdI_ID = DBStrings.dr1.Item("sNo")
    '            EntData.EntityName = "" & DBStrings.dr1.Item("EntityName")
    '            EntData.Associations = "" & DBStrings.dr1.Item("Associations")
    '            EntData.ShortName = "" & DBStrings.dr1.Item("ShortName")
    '            EntData.FormName = "" & DBStrings.dr1.Item("FormName")
    '            EntData.Sheetname = "" & DBStrings.dr1.Item("SheetName")
    '            Call cEntityDataItems.Add(EntData)
    '            ' rst.MoveNext()
    '        Loop
    'exitloop:

    '        DBStrings.cmd1.Dispose()
    '        DBStrings.cnn1.Close()

  End Function

  Public Function GetAssociation(ByVal ParentEntity As String, ByVal ChldEntity As String, ByRef sChildPropertyName As String, ByRef sJoinTable As String) As String
    'Gets the Parent-child relationship given a parent and a child returns 1-M.M-M or nil
    'sParentEntity: parent entity Input
    'sChildEntity : Child Entity Input
    'sChildPropertyName: Property name of the child entity. Input
    'sJoinTable:Join Table is output
    'Returns association (1-M, M-M etc)

    Dim eD As EntityDataItem
    Dim i As Integer
    Dim Assoc() As String
    Dim det() As String
    sJoinTable = "CSR"
    For Each eD In cEntityDataItems
      If eD.EntityName = ParentEntity Then
        Assoc = Split(eD.Associations, ";")
        For i = 1 To UBound(Assoc) - 1  '0 is out of range
          det = Split(Assoc(i), ",")
          If Trim(det(0)) = ChldEntity Then
            sChildPropertyName = Trim(det(3))
            If Trim(det(1)) = "M-M" Then sJoinTable = Trim(det(2))
            GetAssociation = Trim(det(1))
            Exit Function
          End If
        Next i
      End If
    Next eD
    GetAssociation = "CSR"
  End Function

  Public Function GetJoinTable(ByVal ParentEntity As String, ByVal sChildPropertyName As String, ByRef sChildEntity As String) As String
    'returns the JoinTable given a parent entity and a child property
    'sParentEntity: parent entity Input
    'sChildPropertyName: Property name of the child entity. Input
    'If relationship is not M-M with join table, returns 'Nil'
    '0:Childentity, 1:Relationship with parent, 2:Jointable name if jointable relationship, else 'CSR', 3:ChildPropertyname
    Dim eD As EntityDataItem
    Dim i As Integer
    Dim Assoc() As String
    Dim det() As String
    For Each eD In cEntityDataItems
      If eD.EntityName = ParentEntity Then
        Assoc = Split(eD.Associations, ";")
        For i = 1 To UBound(Assoc) - 1  '0 is out of range
          det = Split(Assoc(i), ",")
          If Trim(det(3)) = sChildPropertyName Then
            gsChildEntityNameFromJT = Trim(det(0))
            If Trim(det(1)) = "M-M" Then
              GetJoinTable = Trim(det(2))
              Exit Function
            End If
          End If
        Next i
      End If
    Next eD
    GetJoinTable = "CSR"
  End Function

  'Public Function IsJoinTable(sJTName As String) As Boolean
  ''Determines if the given string is a jointable name
  '    Dim eD As EntityDataItem
  '    Dim i As Integer
  '    Dim Assoc() As String
  '    Dim det() As String
  '    For Each eD In EntityDataColl
  '        Assoc = Split(eD.Associations, ";")
  '        If UBound(Assoc) = 0 Then GoTo nextloop
  '        For i = 1 To UBound(Assoc) - 1  '0 is out of range
  '            det = Split(Assoc(i), ",")
  '            If Trim(det(2)) = "M-M" Then
  '                If sJTName = Trim(det(3)) Then
  '                    IsJoinTable = True
  '                    Exit Function
  '                End If
  '            End If
  'nextloop:
  '        Next i
  '    Next eD
  '    IsJoinTable = False
  'End Function


  Function ParentHasM1Relationship(ByVal sChild As String) As Boolean
    'Finds out if the child has a M-1 relationship with any parent.
    'Needed because M-1 childs need a dummy entity for combo boxes.
    Dim eD As EntityDataItem
    For Each eD In cEntityDataItems
      If GetAssociation(eD.EntityName, sChild, MGlobals.gsChildPropertyName, gsJoinTable) = "M-1" Then
        ParentHasM1Relationship = True
        Exit Function
      End If
    Next eD
    ParentHasM1Relationship = False
  End Function

End Class