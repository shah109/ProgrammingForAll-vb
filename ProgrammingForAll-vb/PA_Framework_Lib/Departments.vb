Option Explicit On

Public Class Departments
  Inherits System.Collections.CollectionBase
  'sPH_cls_DateTime:
  'gfh
  'sPH_cls_DateTime:End
  Private _sNoHashTable As New Hashtable

  Dim nDecLoadOrder As Integer  'comment
  'Private collInt As Collection
  Dim strChanges As String
  'Public Shared currEnt As Department
  Public Event Changed()

  ' Public Sub New()
  '  collInt = New Collection
  'End Sub

  Public Sub Add(ByRef val As Department)
    Me.List.Add(val)
    _sNoHashTable.Add(val.ID, val)
  End Sub


  'Public ReadOnly Property Item(ByVal index As Integer) As Department
  '  Get
  '    Return Me.List.Item(index)
  '  End Get
  'End Property

  Public Property Item(ByVal sid As String) As Department
    Get
      Return _sNoHashTable.Item(sid)
    End Get
    Set(ByVal value As Department)
      _sNoHashTable.Item(sid) = value
    End Set
  End Property

  Public Sub Remove(ByRef val As Department)
    Me.List.Remove(val)
    _sNoHashTable.Remove(val.ID)
  End Sub

  Public Sub RemoveAll()
    Me.List.Clear()
    _sNoHashTable.Clear()
  End Sub

  Public Function LoadSingleEntity(ByVal sId As String, ByVal sOperation As String) As Integer
    Dim eDpt_ As Department
    'Dim sUser As String

    gStrSqlCall = GetSQLQuery("LoadSingleEntity", sId)
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
      eDpt_ = New Department
    ElseIf sOperation = "U" Then
      eDpt_ = Me.Item(sId)
    End If
    Call Me.LoadEntityItemsForThisEntity(eDpt_)

    eDpt_.Lastupdate = rst.Fields("LastUpdate").Value
    eDpt_.mContainer = Me

    eDpt_.Loadorder = Me.Count
    If sOperation = "N" Then
      eDpt_.Loadorder = Me.Count - nDecLoadOrder
      Call cDepartments.Add(eDpt_)
    End If
    rst.Close()
    cnn.Close()
    LoadSingleEntity = 1
  End Function

  Public Function Load()
    Dim eDpt_ As Department
    Dim nLoadorder As Integer
    gStrSqlCall = GetSQLQuery("Load", "ID")

    cnn.Open(MGlobals.GetAppConnString)
    rst.CursorLocation = adUseClient
    rst.Open(gStrSqlCall, cnn, ADODB.CursorTypeEnum.adOpenStatic, 1)
    Me.RemoveAll()
    rst.MoveFirst()
    nDecLoadOrder = 0
    'if this entity has a M-1 rel with any other entity, then a blank entity must be present for cbos.
    If cEntityDataItems.ParentHasM1Relationship("Department") Then
      eDpt_ = New Department   'this member is loaded for 0 index entries cbo's for M-1 relationships
      eDpt_.Loadorder = 0
      eDpt_.ID = ""
      nDecLoadOrder = 1
      Call cDepartments.Add(eDpt_)
    End If
    nLoadorder = 1
    Do While Not rst.EOF
      eDpt_ = New Department

      Call Me.LoadEntityItemsForThisEntity(eDpt_)

      eDpt_.Lastupdate = rst.Fields("LastUpdate").Value
      eDpt_.Loadorder = nLoadorder
      eDpt_.mContainer = Me
      Call cDepartments.Add(eDpt_)
      nLoadorder = nLoadorder + 1
      rst.MoveNext()
    Loop
    rst.Close()
    cnn.Close()
    'Build the domain model now
    'Call Me.BuildDomainModel
  End Function

  Sub BuildDomainModel()
    '  Dim eb_ As Department
    '  Dim e1_ As Entity1
    '  'now add eDpt_s to e1_s
    '  For Each e1_ In cEntity1s.Items  'assumed cEntity1s have been loaded
    '    Set e1_.ChildDepartments = e1_.BuildChildEntityObjects(e1_.ChildDepartments, e1_.ChildEntityString(CurreDpt_))
    '    'For Each eDpt_ In e1_.ChildDepartments.Items
    '    ' If Not eDpt_.ContainsParentEntity(e1_) Then  'BuildObjectModel is called from AddtoDB() too. hence to check for already existing.
    '    '  Call eDpt_.ParentEntity1s.Add(e1_)   'add parents to Department
    '    'End If
    '    'Next eDpt_
    '  Next e1_
  End Sub

  Public Function UpdateDB(ByRef eDpt_ As Department) As Integer
    'Updates the database table for the entity'.
    'Return 0 if the db record has been updated by other user.
    'Return the loadorder if successfull in update.
    Dim nCount As Integer
    gbChangesMade = False
    gStrSqlCall = GetSQLQuery("UpdateDB", eDpt_.ID)
    If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
     cnn.Open(MGlobals.GetAppConnString)
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
    If rst.Fields("LastUpdate").Value <> eDpt_.Lastupdate Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    Else
      rst.Fields("LastUpdate").Value = Now
      eDpt_.Lastupdate = rst.Fields("LastUpdate").Value  'update the local last update so you can mame changes without refreshing.
    End If
    strChanges = ""
    ''PH:For Each Entity Item. Add to record changes int the DB (function UpdateDB()
    'Call BLFunctions.gRecordChanges(eDpt_, "sEntityItem_1", eDpt_.EntityItem_1, "fn:")
    ''PH: End
    'sPH_col_Update:
    Call RecordChanges("Update", eDpt_)
    ' Call BLFunctions.gRecordChanges("sEntityItem_1", eDpt_.EntityItem_1, "ei1:")
    'Call BLFunctions.gRecordChanges("sEntityItem_2", eDpt_.EntityItem_2, "ei2:")
    'sPH_col_Update:End

    ''PH: For Each Child Entity, add this line for recording the changes for child entities for this object.
    ' Call eDpt_.BuildChildEntityString(eDpt_.ChildEntities(e3_Par))
    ' Call BLFunctions.gRecordChanges(eDpt_, "sEntity3s", eDpt_.ChildEntityString(e3_Par), "e3_str:")
    ''PH: End
    'sPH_col_UpdateChild:
    'followng two commented
    'Call eDpt_.BuildChildEntityString("EntityBs")
    'Call BLFunctions.gRecordChanges1("Update", "EntityBs", eDpt_.ChildEntityString("EntityBs"), "eB_str:")

    'sPH_col_UpdateChild:End

    rst.Update()
    rst.Close()
    MGlobals.curreDpt_ = eDpt_

    'No need to enter a record in update log if no change has been made
    If gbChangesMade = True Then
      AppSettings.LastUpdateID = gUpdateLogTable("eDpt_", "U", eDpt_.ID)
    End If
    cnn.Close()
    Call DoChanged()
    UpdateDB = nCount
  End Function

  Public Function AddtoDB(ByRef eDpt_ As Department) As Integer
    Dim strSQL As String
    Dim nCount As Integer
    strSQL = "select ID from Departments"
    If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      cnn.Open(MGlobals.GetAppConnString)
    End If
    'If rst.State = ADODB.ObjectStateEnum.adStateOpen Then rst.Close()
    rst.Open(Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
    nCount = rst.RecordCount
    If cEntityDataItems.ParentHasM1Relationship("Department") Then nCount = nCount + 1
    If nCount <> cDepartments.Count - 1 Then  'other user has added a record
      rst.Close()
      cnn.Close()
      AddtoDB = 0
      Exit Function
    End If
    rst.Close()
    strSQL = "select max(ID) as MaxNo from Departments"
    Dim nMax As Integer
    rst.CursorLocation = adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, adLockOptimistic)
    nMax = rst.Fields("MaxNo").Value
    rst.Close()
    eDpt_.ID = nMax + 1
    'now add the record
    rst.CursorType = adOpenKeySet
    rst.LockType = adLockOptimistic
    rst.Open("Departments", cnn, , , adCmdTable)
    rst.AddNew()

    ''PH:For Each Entity Item. Add the items to the db (function AddtoDB)
    'Call gAddChanges("EntityItem_2", eDpt_.EntityItem_2, "ln:")
    ''PH: End
    'sPH_col_Add:

    Call BLFunctions.gRecordChanges1("Add", "ID", eDpt_.ID, "id:")
    Call RecordChanges("Add", eDpt_)
    'sPH_col_Add:End

    ''PH: For Each Child Entity, add this line for saving the child string with e3_ is the child.
    ' Call gAddChanges("Entity_Cs", eDpt_.ChildEntityString(e3_Par), "eDpt_str:")
    ''PH: End
    'sPH_col_AddChild:
    '    Call eDpt_.BuildChildEntityString(eDpt_.ChildEntities(CurreB_))
    '    Call gAddChanges("EntityBs", eDpt_.ChildEntityString(CurreB_), "eB_str:")
    'sPH_col_AddChild:End

    rst("LastUpdate").Value = Now
    eDpt_.Lastupdate = rst("LastUpdate").Value
    'eDpt_.mContainer = Me
    rst.Update()
    rst.Close()

    Call Me.Add(eDpt_)
    eDpt_.Loadorder = Me.Count- nDecLoadOrder

    Call Me.BuildDomainModel()
    MGlobals.curreDpt_ = eDpt_
    'now update the log
    AppSettings.LastUpdateID = gUpdateLogTable("eDpt_", "N", eDpt_.ID)
    cnn.Close()
    Call DoChanged()
    AddtoDB = nCount + 1
  End Function

  Public Function DeleteFromDB(ByVal eDpt_ As Department) As Integer
    Dim lAffected As Long
    Dim nCount As Integer
    Dim par As Object
    gStrSqlCall = " Delete " & _
                  "FROM " & strDBO & "Departments " & _
                  " where " & strDBO & "ID LIKE '" & eDpt_.ID & "'"

    MGlobals.cmd.ActiveConnection = cnn
    MGlobals.cmd.CommandText = gStrSqlCall
    MGlobals.cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
    rst = MGlobals.cmd.Execute(lAffected, ADODB.ExecuteOptionEnum.adExecuteNoRecords)
    If lAffected = 1 Then
      Dim sEbID As String
      sEbID = eDpt_.ID
      MGlobals.gstrChanges = "Deleted: " & "ID" & _
                            "fn:" & eDpt_.EntityItem_1 & _
                            ", ln:" & eDpt_.EntityItem_2
      Call cDepartments.Remove(eDpt_)
      cnn.Open(gStrConnGenEBA)
      AppSettings.LastUpdateID = gUpdateLogTable("eDpt_", "D", sEbID)
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
              strDBO & "ID , " & _
              strDBO & "EntityItem_1 , " & _
              strDBO & "EntityItem_2 , " & _
              strDBO & "EntityBs , " & _
              strDBO & "LastUpdate " & _
              "FROM " & strDBO & "Departments "
    'sPH_col_Sql:End

    If strFunc = "LoadSingleEntity" Or strFunc = "UpdateDB" Then
      strTemp = strTemp & " WHERE " & strDBO & "ID like " & strID
    ElseIf strFunc = "Load" Then
      strTemp = strTemp & " ORDER BY " & strDBO & strID
    End If
    GetSQLQuery = strTemp
  End Function

  Function LoadEntityItemsForThisEntity(ByVal ent As Department)
    ''PH:For Each Entity Item. Add it for loading
    'eDpt_.EntityItem_1 = "" & rst.Fields("sEntityItem_1")
    ''PH End
    'sPH_col_Load:
    ent.ID = "" & rst.Fields("ID").Value
    ent.EntityItem_1 = "" & rst.Fields("EntityItem_1").Value
    ent.EntityItem_2 = "" & rst.Fields("EntityItem_2").Value
    'ent.Courses = "" & rst.Fields("sEntityBs").Value
    'sPH_col_Load:End

    ''PH: For Each Child Entity, add the two lines for building child entities for this object.
    ' Call ent.BuildChildEntityObjects(e3_Par, "" & rst.Fields("sEntity3s"))
    ''PH: End
    'sPH_col_LoadChildEntities:
    Call ent.BuildChildEntityObjects("Courses", "" & rst.Fields("EntityBs").Value)
    'sPH_col_LoadChildEntities:End
  End Function

  'Private Sub DoChanged()
  ''Call Globals.CallFormsEvent
  '    On Error Resume Next
  '    Call frmDepartments.EntToRaiseChanged.RaiseChanged
  'End Sub

  Public Sub RaiseChanged()
    RaiseEvent Changed()
  End Sub

  Private Sub RecordChanges(ByVal sOperation As String, ByVal ent As Department)
    Call BLFunctions.gRecordChanges1(sOperation, "EntityItem_1", ent.EntityItem_1, "ei1:")
    Call BLFunctions.gRecordChanges1(sOperation, "EntityItem_2", ent.EntityItem_2, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "EntityBs", ent.ChildEntityString("Courses"), "eB_str:")
  End Sub

End Class

