Imports ADODB
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Public Class PAEnts
  Inherits System.Collections.CollectionBase
  Private _sNoHashTable As New Hashtable
  Public Event UpdateEvent()
  Private _LOHashtable As New Hashtable
  Private LoadOrder1 As Integer
  Sub New()
    LoadOrder1 = 1
  End Sub

  Public Sub Add(ByRef val As PAEnt)
    val.LoadOrder = LoadOrder1

    Me.List.Add(val)
    'val.LoadOrder = Me.List.IndexOf(val)
    _sNoHashTable.Add(val.ID, val)
    _LOHashtable.Add(val.LoadOrder, val)
    LoadOrder1 = LoadOrder1 + 1
  End Sub

  Public Function IndexOf(ByVal value As PAEnt) As Integer
    Return Me.List.IndexOf(value)
    'Me.List.
  End Function 'IndexOf



  Public Function GetEntityByLoadOrder(ByVal lo As Integer) As PAEnt
    GetEntityByLoadOrder = Nothing
    Dim ent As PAEnt
    For Each ent In Me
      If ent.LoadOrder = lo Then
        GetEntityByLoadOrder = ent
        Exit Function
      End If
    Next
  End Function
  'Public Property LoadOrder(ByVal sid As Integer)
  '  Get
  '    Return Me.LoadOrder
  '  End Get
  '  Set(ByVal value)
  '    sid.LoadOrder= 
  '  End Set
  'End Property
  'Public ReadOnly Property ItemByLO(ByVal sid As Integer) As Object
  '  Get
  '    Return Me.List.Item(sid)
  '  End Get
  'End Property

  Public Property Item(ByVal sid As String) As Object
    Get
      Return _sNoHashTable.Item(sid)
    End Get
    Set(ByVal value As Object)
      _sNoHashTable.Item(sid) = value
    End Set
  End Property

  Public Property LoadOrder(ByVal sid As Integer) As Object
    Get
      Return _LOHashtable.Item(sid)
    End Get
    Set(ByVal value As Object)
      _LOHashtable.Item(sid) = value
    End Set
  End Property

  Public Sub Remove(ByRef val As Object)
    Me.List.Remove(val)
    _sNoHashTable.Remove(val.ID)
    '_LOHashtable.Remove(Me.LoadOrder1)
  End Sub

  Public Sub RemoveAll()
    Me.List.Clear()
    _sNoHashTable.Clear()
    '_LOHashtable.Clear()
  End Sub

  Public Sub MoveUp(ByRef ent As PAEnt)
    If Me.List.IndexOf(ent) = 0 Then Exit Sub
    Dim temp As New PAEnt
    Dim n As Integer
    n = Me.List.IndexOf(ent)
    Me.List.RemoveAt(n)
    Me.List.Insert(n - 1, ent)

  End Sub

  Public Function Contains(ByVal id As String) As Boolean
    Dim ent As PAEnt
    For Each ent In Me
      If id = ent.ID Then
        Contains = True
        Exit Function
      End If
    Next
    Contains = False
  End Function

  Public Function LoadSingleEntity(ByVal sEnt As String, ByVal sOperation As String) As Integer
    Dim ent As PAEnt
    ent = Me.Item(sEnt)
    If sOperation = "D" Then
      Me.Remove(ent)
      Exit Function
    End If
    gStrSqlCall = GetSQLQuery("LoadSingleEntity", sEnt)

    If AppSettings.cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      AppSettings.cnn.Open(AppSettings.GetAppConnString)
    End If
    If AppSettings.rst.State = ADODB.ObjectStateEnum.adStateOpen Then
      AppSettings.rst.Close()
    End If
    AppSettings.rst.CursorLocation = adUseClient
    AppSettings.rst.Open(gStrSqlCall, AppSettings.cnn, adOpenStatic, 1)
    If AppSettings.rst.RecordCount <> 1 Then
      LoadSingleEntity = 0  '
      AppSettings.rst.Close()
      AppSettings.cnn.Close()
      Exit Function
    End If
    If sOperation = "N" Then
      ent = CreateNewEntity()
    End If
    Call Me.LoadEntityItemsForThisEntity(ent)

    ent.Lastupdate = AppSettings.rst.Fields("LastUpdate").Value
    ent.mContainer = Me
    ent.Loadorder = Me.Count
    If sOperation = "N" Then
      Call Me.Add(ent)
    End If

    AppSettings.rst.Close()
    AppSettings.cnn.Close()
    LoadSingleEntity = 1
  End Function

  Public Sub Load(ByRef ent As PAEnt)
    'Loads entities from db table into entity collections of the object model.
    gStrSqlCall = GetSQLQuery("Load", "ID")

    If AppSettings.cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      AppSettings.cnn.Open(AppSettings.GetAppConnString)
    End If

    'AppSettings.cnn.Open(AppSettings.GetAppConnString)
    AppSettings.rst.CursorLocation = adUseClient
    AppSettings.rst.Open(gStrSqlCall, AppSettings.cnn, ADODB.CursorTypeEnum.adOpenStatic, 1)
    Call Me.RemoveAll()
    AppSettings.rst.MoveFirst()
    'LoadOrder1 = 0
    Do While Not AppSettings.rst.EOF
      ent = CreateNewEntity()
      Call Me.LoadEntityItemsForThisEntity(ent)
      'LoadOrder1 = +1
      'ent.LoadOrder = LoadOrder1
      ent.Lastupdate = AppSettings.rst.Fields("LastUpdate").Value
      ent.mContainer = Me
      Call Me.Add(ent)
      AppSettings.rst.MoveNext()
    Loop
    AppSettings.rst.Close()
    AppSettings.cnn.Close()
    'Build the domain model now
    Call Me.BuildDomainModel()
  End Sub

  Public Function UpdateDB(ByRef ent As PAEnt) As Integer
    'Updates the database table for the entity'.
    'Return 0 if the db record has been updated by other user, 1 if successful, -1 if nothing to update

    Dim nCount As Integer
    Dim dLastUpdate As Date
    gStrSqlCall = GetSQLQuery("UpdateDB", ent.ID)
    If AppSettings.cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      AppSettings.cnn.Open(AppSettings.GetAppConnString)
    End If

    AppSettings.rst.Open(Source:=gStrSqlCall, ActiveConnection:=AppSettings.cnn, _
             CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
    nCount = AppSettings.rst.RecordCount
    If nCount <> 1 Then
      AppSettings.rst.Close()
      AppSettings.cnn.Close()
      UpdateDB = 0
      Exit Function
    End If
    dLastUpdate = AppSettings.rst.Fields("LastUpdate").Value
    If dLastUpdate <> ent.Lastupdate Then
      AppSettings.rst.Close()
      AppSettings.cnn.Close()
      UpdateDB = 0
      Exit Function
    Else
      AppSettings.rst.Fields("LastUpdate").Value = Now
      ent.Lastupdate = AppSettings.rst.Fields("LastUpdate").Value  'update the local last update so you can mame changes without refreshing.
    End If
    gbChangesMade = False
    gstrChanges = ""
    Call RecordChanges("Update", ent)
    AppSettings.rst.Update()
    AppSettings.rst.Close()
    'Me.Item(ent.ID) = ent

    Call SetCurrentEntity(ent)
    'No need to enter a record in update log if no change has been made
    If gbChangesMade = True Then
      Call UpdateLogTable(TypeName(ent), "U", ent.ID, gstrChanges)
      UpdateDB = 1
      RaiseEvent UpdateEvent()
    Else
      UpdateDB = -1
    End If
    AppSettings.cnn.Close()
    'Call LoadSingleEntity(ent.ID, "U")
  End Function

  Public Function AddtoDB(ByRef ent As PAEnt) As Integer
    'returns 0 if unsuccessful, total count if successful, 
    Dim strSQL As String
    Dim nCount As Integer
    strSQL = GetSQLQuery("AddtoDB", "ID")
    Try
      If AppSettings.cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
        AppSettings.cnn.Open(AppSettings.GetAppConnString)
      End If
    Catch
      MsgBox(Err.Description)
      AppSettings.cnn.Close()
    End Try

    AppSettings.rst.Open(Source:=strSQL, ActiveConnection:=AppSettings.cnn, CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
    nCount = AppSettings.rst.RecordCount
    'If Not cEntityDataItems.IsJoinTable(TypeName(ent)) Then 'if this is a join table, the item has not yet been added to the collection
    'nCount = nCount + 1
    'End If
    If nCount <> Me.Count Then  'other user has added a record
      AppSettings.rst.Close()
      AppSettings.cnn.Close()
      AddtoDB = 0
      Exit Function
    End If
    AppSettings.rst.Close()
    strSQL = GetSQLQuery("AddtoDB", "MaxID")
    Dim nMax As Integer
    AppSettings.rst.CursorLocation = adUseClient

    AppSettings.rst.Open(strSQL, AppSettings.cnn, ADODB.CursorTypeEnum.adOpenStatic, adLockOptimistic)
    nMax = AppSettings.rst.Fields("MaxID").Value
    AppSettings.rst.Close()

    ent.ID = nMax + 1

    'now add the record
    AppSettings.rst.CursorType = adOpenKeySet
    AppSettings.rst.LockType = adLockOptimistic
    AppSettings.rst.Open(TypeName(ent.mContainer), AppSettings.cnn, , , adCmdTable)
    AppSettings.rst.AddNew()
    gstrChanges = ""
    Call gRecordChanges("Add", "ID", ent.ID, "id:") 'ID needs to be directly done because recordchanges does not cover ID
    Call RecordChanges("Add", ent)

    AppSettings.rst("LastUpdate").Value = Now
    ent.Lastupdate = AppSettings.rst("LastUpdate").Value
    ent.mContainer = Me
    AppSettings.rst.Update()
    AppSettings.rst.Close()
    '''''''''''''''''''''''''''''''''''
    Me.Add(ent)
    '''''''''''''''''''''''''''''''''''
    ent.Loadorder = Me.Count

    Call SetCurrentEntity(ent)
    'now update the log
    Call UpdateLogTable(TypeName(ent), "N", ent.ID, gstrChanges)
    AppSettings.cnn.Close()
    AddtoDB = nCount + 1
  End Function

  Public Function DeleteFromDB(ByRef ent As PAEnt) As Integer
    Dim lAffected As Long
    gStrSqlCall = GetSQLQuery("DeleteFromDB", ent.ID)
    If AppSettings.cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      AppSettings.cnn.Open(AppSettings.GetAppConnString)
    End If
    AppSettings.cmd.ActiveConnection = AppSettings.cnn
    AppSettings.cmd.CommandText = gStrSqlCall
    AppSettings.cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
    AppSettings.rst = AppSettings.cmd.Execute(lAffected, ADODB.ExecuteOptionEnum.adExecuteNoRecords)
    AppSettings.cnn.Close()
    If lAffected = 1 Then

      OMGlobals.gstrChanges = "Deleted : " & ent.ID
      AppSettings.cnn.Open(gStrConnGenEBA)
      Call UpdateLogTable(TypeName(ent), "D", ent.ID, gstrChanges)
      AppSettings.cnn.Close()
      Call Me.Remove(ent)

    End If
    If AppSettings.cnn.State = ADODB.ObjectStateEnum.adStateOpen Then
      AppSettings.cnn.Close()
    End If
    DeleteFromDB = lAffected
  End Function

  Function GetSQLQuery(ByVal strFunc As String, ByVal strID As String) As String
    Dim strSelect As String = String.Empty
    Dim strSelectMax As String = String.Empty
    Dim strDelete As String = String.Empty
    Dim strFrom As String = String.Empty
    Dim strItems As String = String.Empty
    Dim strWhereIDlike As String = String.Empty
    Dim strOrderyBy As String = String.Empty
    GetSQLQuery = String.Empty
    strSelect = " Select " '& _
    'strDBO & "ID "
    strSelectMax = " Select " & strDBO & "Max(ID) as MaxID "
    strDelete = "Delete "
    strItems = GetSelectDBItems()
    strFrom = GetSelectFromTable()
    strWhereIDlike = " WHERE " & strDBO & "ID like " & strID

    If strFunc = "LoadSingleEntity" Or strFunc = "UpdateDB" Then
      GetSQLQuery = strSelect & strItems & strFrom & strWhereIDlike
    ElseIf strFunc = "Load" Then
      GetSQLQuery = strSelect & strItems & strFrom & " ORDER BY " & strDBO & strID & ";"
    ElseIf strFunc = "DeleteFromDB" Then
      GetSQLQuery = strDelete & strFrom & strWhereIDlike & ";"
    ElseIf strFunc = "AddtoDB" Then
      Select Case strID
        Case "ID"
          GetSQLQuery = strSelect & strID & strFrom & ";"
        Case "MaxID"
          GetSQLQuery = strSelectMax & strFrom & ";"
      End Select
    End If
  End Function

  Overridable Sub BuildDomainModel()
  End Sub

  Overridable Function GetSelectFromTable() As String
    GetSelectFromTable = ""
  End Function

  Overridable Function GetSelectDBItems() As String
    GetSelectDBItems = ""
  End Function

    Overridable Sub LoadEntityItemsForThisEntity(ByRef ent As Object)
    End Sub

    Overridable Sub RecordChanges(ByVal sOperation As String, ByRef ent As Object)
    'must be overridden by the derived entity class. Each item to be updated into the db must have a call to gRecordChanges()
    'Call gRecordChanges(sOperation, "<dbcolumnName>", ent.ChildEntityString("<PropertyName>"), "Stds:")
  End Sub

  Overridable Function CreateNewEntity() As PAEnt
    CreateNewEntity = New Object
  End Function

  Overridable Sub SetCurrentEntity(ByRef ent As PAEnt)
  End Sub
End Class



