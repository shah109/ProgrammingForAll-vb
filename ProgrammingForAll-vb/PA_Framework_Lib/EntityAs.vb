Option Explicit On

Public Class EntityAs
  Inherits System.Collections.CollectionBase
  'sPH_cls_DateTime:
  'gfh
  'sPH_cls_DateTime:End
  Private _sNoHashTable As New Hashtable

  'Private collInt As Collection
    Dim strChanges As String
  'Public Shared currEnt As EntityA
  Public Event Changed()

  ' Public Sub New()
  '  collInt = New Collection
  'End Sub

  Public Sub Add(ByRef val As EntityA)
    Me.List.Add(val)
    _sNoHashTable.Add(val.ID, val)
  End Sub


  Public ReadOnly Property Item(ByVal index As Integer) As EntityA
    Get
      Return Me.List.Item(index)
    End Get
  End Property

  Public Property Item(ByVal sid As String) As EntityA
    Get
      Return _sNoHashTable.Item(sid)
    End Get
    Set(ByVal value As EntityA)
      _sNoHashTable.Item(sid) = value
    End Set
  End Property

  Public Sub Remove(ByRef val As EntityA)
    Me.List.Clear()
    Me._sNoHashTable.Clear()
  End Sub


  Public Function LoadSingleEntity(ByVal sId As String, ByVal sOperation As String) As Integer
    Dim eA_ As EntityA
    'Dim sUser As String

    gStrSqlCall = GetSQLQuery("LoadSingleEntity", sid)
    cnn.Open(gStrConnGenEBA)
    rst.CursorLocation = adUseClient
    rst.Open(gStrSqlCall, cnn, _
             adOpenStatic, 1)
    If rst.RecordCount <> 1 Then
      LoadSingleEntity = 0  '
      rst.Close()
      cnn.Close()
      Exit Function
    End If
    If sOperation = "N" Then
      eA_ = New EntityA
    ElseIf sOperation = "U" Then
      eA_ = Me.item(sid)
    End If
    Call Me.LoadEntityItemsForThisEntity(eA_)

    eA_.Lastupdate = rst.Fields("sLastUpdate").Value
    eA_.mContainer = Me

    eA_.Loadorder = Me.Count
    If sOperation = "N" Then
      Call cEntityAs.Add(eA_)
    End If
    rst.Close()
    cnn.Close()
    LoadSingleEntity = 1
  End Function

  Public Function Load()
    Dim eA_ As EntityA
    Dim nLoadorder As Integer
    gStrSqlCall = GetSQLQuery("Load", "EntityItem_1")

    cnn.Open(MGlobals.GetAppConnString)
    rst.CursorLocation = adUseClient
    rst.Open(gStrSqlCall, cnn, ADODB.CursorTypeEnum.adOpenStatic, 1)
    Me.Clear()
    rst.MoveFirst()
    'if this entity has a M-1 rel with any other entity, then a blank entity must be present for cbos.
    If cEntityDataItems.ParentHasM1Relationship("EntityA") Then
      eA_ = New EntityA   'this member is loaded for 0 index entries cbo's for M-1 relationships
      eA_.Loadorder = 0
      eA_.ID = 0
      Call cEntityAs.Add(eA_)
    End If
    nLoadorder = 1
    Do While Not rst.EOF
      eA_ = New EntityA

      Call Me.LoadEntityItemsForThisEntity(eA_)

      eA_.Lastupdate = rst.Fields("sLastUpdate").Value
      eA_.Loadorder = nLoadorder
      eA_.mContainer = Me
      Call cEntityAs.Add(eA_)
      nLoadorder = nLoadorder + 1
      rst.MoveNext()
    Loop
    rst.Close()
    cnn.Close()
    'Build the domain model now
    'Call Me.BuildDomainModel
  End Function

  Sub BuildDomainModel()
    '  Dim eb_ As EntityA
    '  Dim e1_ As Entity1
    '  'now add eA_s to e1_s
    '  For Each e1_ In cEntity1s.Items  'assumed cEntity1s have been loaded
    '    Set e1_.ChildEntityAs = e1_.BuildChildEntityObjects(e1_.ChildEntityAs, e1_.ChildEntityString(CurreA_))
    '    'For Each eA_ In e1_.ChildEntityAs.Items
    '    ' If Not eA_.ContainsParentEntity(e1_) Then  'BuildObjectModel is called from AddtoDB() too. hence to check for already existing.
    '    '  Call eA_.ParentEntity1s.Add(e1_)   'add parents to EntityA
    '    'End If
    '    'Next eA_
    '  Next e1_
  End Sub

  Public Function UpdateDB(ByRef eA_ As EntityA) As Integer
    'Updates the database table for the entity'.
    'Return 0 if the db record has been updated by other user.
    'Return the loadorder if successfull in update.
    Dim nCount As Integer
    gbChangesMade = False
    gStrSqlCall = GetSQLQuery("UpdateDB", eA_.ID)
    If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      cnn.Open(gStrConnGenEBA)
    End If

    rst.Open(Source:=gStrSqlCall, ActiveConnection:=cnn, _
             CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
    nCount = rst.RecordCount
    If nCount <> 1 Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    End If
    If rst.Fields("sLastUpdate").Value <> eA_.Lastupdate Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    Else
      rst.Fields("sLastUpdate").Value = Now
      eA_.Lastupdate = rst.Fields("sLastUpdate").Value  'update the local last update so you can mame changes without refreshing.
    End If
    strChanges = ""
    ''PH:For Each Entity Item. Add to record changes int the DB (function UpdateDB()
    'Call BLFunctions.gRecordChanges(eA_, "sEntityItem_1", eA_.EntityItem_1, "fn:")
    ''PH: End
    'sPH_col_Update:
    Call RecordChanges("Update", eA_)
    ' Call BLFunctions.gRecordChanges("sEntityItem_1", eA_.EntityItem_1, "ei1:")
    'Call BLFunctions.gRecordChanges("sEntityItem_2", eA_.EntityItem_2, "ei2:")
    'sPH_col_Update:End

    ''PH: For Each Child Entity, add this line for recording the changes for child entities for this object.
    ' Call eA_.BuildChildEntityString(eA_.ChildEntities(e3_Par))
    ' Call BLFunctions.gRecordChanges(eA_, "sEntity3s", eA_.ChildEntityString(e3_Par), "e3_str:")
    ''PH: End
    'sPH_col_UpdateChild:
    'followng two commented
    'Call eA_.BuildChildEntityString("EntityBs")
    'Call BLFunctions.gRecordChanges(eA_, "sEntityBs", eA_.ChildEntityString("EntityBs"), "eB_str:")

    'sPH_col_UpdateChild:End

    rst.Update()
    rst.Close()
    MGlobals.curreA_ = eA_

    'No need to enter a record in update log if no change has been made
    If gbChangesMade = True Then
      AppSettings.LastUpdateID = gUpdateLogTable("eA_", "U", eA_.ID)
    End If
    cnn.Close()
    Call DoChanged()
    UpdateDB = nCount
  End Function

  Public Function AddtoDB(ByRef eA_ As EntityA) As Integer
    Dim strSQL As String
    Dim nCount As Integer
    strSQL = "select Id from EntityAs"
    cnn.Open(gStrConnGenEBA)
    rst.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
    nCount = rst.RecordCount
    rst.Close()
    If nCount <> cEntityAs.Count Then  'other user has added a record
      rst.Close()
      cnn.Close()
      AddtoDB = 0
      Exit Function
    End If
    strSQL = "select max(ID) as MaxNo from EntityAs"
    Dim nMax As Integer
    rst.CursorLocation = adUseClient
    rst.Open(strSQL, cnn, adOpenStatic, adLockOptimistic)
    nMax = rst.Fields("MaxNo").Value
    rst.Close()
    eA_.ID = nMax + 1
    'now add the record
    rst.CursorType = adOpenKeySet
    rst.LockType = adLockOptimistic
    rst.Open("EntityAs", cnn, , , adCmdTable)
    rst.AddNew()

    ''PH:For Each Entity Item. Add the items to the db (function AddtoDB)
    'Call gAddChanges("EntityItem_2", eA_.EntityItem_2, "ln:")
    ''PH: End
    'sPH_col_Add:

    Call BLFunctions.gRecordChanges1("Add", "ID", eA_.ID, "id:")
    Call RecordChanges("Add", eA_)
    'Call gAddChanges("EntityItem_1", eA_.EntityItem_1, "ei1:")
    'Call gAddChanges("EntityItem_2", eA_.EntityItem_2, "ei2:")
    'sPH_col_Add:End

    ''PH: For Each Child Entity, add this line for saving the child string with e3_ is the child.
    ' Call gAddChanges("Entity_Cs", eA_.ChildEntityString(e3_Par), "eA_str:")
    ''PH: End
    'sPH_col_AddChild:
    '    Call eA_.BuildChildEntityString(eA_.ChildEntities(CurreB_))
    '    Call gAddChanges("EntityBs", eA_.ChildEntityString(CurreB_), "eB_str:")
    'sPH_col_AddChild:End

    rst("LastUpdate").Value = Now
    eA_.Lastupdate = rst("LastUpdate").Value
    eA_.mContainer = Me
    rst.Update()
    rst.Close()

    Call Me.Add(eA_)
    eA_.Loadorder = Me.Count

    Call Me.BuildDomainModel()
    MGlobals.curreA_ = eA_
    'now update the log
    AppSettings.LastUpdateID = gUpdateLogTable("eA_", "N", eA_.ID)
    cnn.Close()
    Call DoChanged()
    AddtoDB = nCount + 1
  End Function

  Public Function DeleteFromDB(ByVal eA_ As EntityA) As Integer
    Dim lAffected As Long
    Dim nCount As Integer
    Dim par As Object
    gStrSqlCall = " Delete " & _
                  "FROM " & strDBO & "EntityAs " & _
                  " where " & strDBO & "ID LIKE '" & eA_.ID & "'"

    MGlobals.cmd.ActiveConnection = cnn
    MGlobals.cmd.CommandText = gStrSqlCall
    MGlobals.cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
    rst = MGlobals.cmd.Execute(lAffected, ADODB.ExecuteOptionEnum.adExecuteNoRecords)
    If lAffected = 1 Then
      Dim sEbID As String
      sEbID = eA_.ID
      MGlobals.gstrChanges = "Deleted: " & "ID" & _
                            "fn:" & eA_.EntityItem_1 & _
                            ", ln:" & eA_.EntityItem_2
      Call cEntityAs.Remove(eA_)
      cnn.Open(gStrConnGenEBA)
      AppSettings.LastUpdateID = gUpdateLogTable("eA_", "D", sEbID)
      cnn.Close()
      Call DoChanged()
    End If
    MGlobals.cmd = Nothing
    DeleteFromDB = lAffected
  End Function

  Function GetSQLQuery(ByVal strFunc As String, ByVal strID As String) As String
    'This SQL query is used in 3 functions: LoadSingleEntity(). Load() and UpdateDB()
    Dim strTemp As String
    'PH: For Each Entity Item, add to the Select
    ' strDBO & "EntityItem_2 as sEntityItem_2, " & _
    'PH: End
    'sPH_col_Sql:
    strTemp = " Select " & _
              strDBO & "ID As sID, " & _
              strDBO & "EntityItem_1 As sEntityItem_1, " & _
              strDBO & "EntityItem_2 As sEntityItem_2, " & _
              strDBO & "EntityBs As sEntityBs, " & _
              strDBO & "LastUpdate As sLastUpdate " & _
              "FROM " & strDBO & "Departments "
    'sPH_col_Sql:End

    If strFunc = "LoadSingleEntity" Or strFunc = "UpdateDB" Then
      strTemp = strTemp & " WHERE " & strDBO & "ID like " & strID
    ElseIf strFunc = "Load" Then
      strTemp = strTemp & " ORDER BY " & strDBO & strID
    End If
    GetSQLQuery = strTemp
  End Function

  Function LoadEntityItemsForThisEntity(ByVal ent As EntityA)
    ''PH:For Each Entity Item. Add it for loading
    'eA_.EntityItem_1 = "" & rst.Fields("sEntityItem_1")
    ''PH End
    'sPH_col_Load:
    ent.ID = "" & rst.Fields("sID").Value
    ent.EntityItem_1 = "" & rst.Fields("sEntityItem_1").Value
    ent.EntityItem_2 = "" & rst.Fields("sEntityItem_2").Value
    ent.EntityBs = "" & rst.Fields("sEntityBs").Value
    'sPH_col_Load:End

    ''PH: For Each Child Entity, add the two lines for building child entities for this object.
    ' Call ent.BuildChildEntityObjects(e3_Par, "" & rst.Fields("sEntity3s"))
    ''PH: End
    'sPH_col_LoadChildEntities:
    'Call ent.BuildChildEntityObjects("EntityBs", "" & rst.Fields("sEntityBs").Value)
    'sPH_col_LoadChildEntities:End
  End Function

  'Private Sub DoChanged()
  ''Call Globals.CallFormsEvent
  '    On Error Resume Next
  '    Call frmEntityAs.EntToRaiseChanged.RaiseChanged
  'End Sub

  Public Sub RaiseChanged()
    RaiseEvent Changed()
  End Sub

  Private Sub RecordChanges(ByVal sOperation As String, ByVal ent As EntityA)
    Call BLFunctions.gRecordChanges1(sOperation, "sEntityItem_1", ent.EntityItem_1, "ei1:")
    Call BLFunctions.gRecordChanges1(sOperation, "sEntityItem_2", ent.EntityItem_2, "ei2:")
  End Sub

End Class
