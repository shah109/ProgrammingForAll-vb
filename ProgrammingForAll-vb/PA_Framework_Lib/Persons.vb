Option Explicit On

Public Class Persons
  Inherits System.Collections.CollectionBase

  Private _sNoHashTable As New Hashtable
  Private _LOHashTable As New Hashtable

  Dim nDecLoadOrder As Integer  'comment
  Dim strChanges As String

  Public Event Changed()

  Public Sub Add(ByRef val As Person)
    Me.List.Add(val)
    _sNoHashTable.Add(val.ID, val)
    _LOHashTable.Add(val.Loadorder, val)
  End Sub

  Public ReadOnly Property ItemLO(ByVal index As Integer) As Person
    Get
      Return Me._LOHashTable.Item(index)
    End Get
  End Property

  Public ReadOnly Property Item(ByVal sid As String) As Person
    Get
      Return _sNoHashTable.Item(sid)
    End Get
    
  End Property

  Public Sub Remove(ByRef val As Person)
    Me.List.Remove(val)
    _sNoHashTable.Remove(val.ID)
    _LOHashTable.Remove(val.Loadorder)
  End Sub

  Public Sub RemoveAll()
    Me.List.Clear()    'alway remove the first item
    _sNoHashTable.Clear()
    _LOHashTable.Clear()
  End Sub

  Public Function LoadSingleEntity(ByVal sId As String, ByVal sOperation As String) As Integer
    Dim ePrsn_ As Person

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
      ePrsn_ = New Person
    ElseIf sOperation = "U" Then
      ePrsn_ = Me.Item(sId)
    End If
    Call Me.LoadEntityItemsForThisEntity(ePrsn_)

    ePrsn_.Lastupdate = rst.Fields("LastUpdate").Value
    ePrsn_.mContainer = Me

    ePrsn_.Loadorder = Me.Count
    If sOperation = "N" Then
      Call cPersons.Add(ePrsn_)
    End If
    rst.Close()
    cnn.Close()
    LoadSingleEntity = 1
  End Function

  Public Function Load()
    Dim ePrsn_ As Person
    Dim nLoadorder As Integer
    gStrSqlCall = GetSQLQuery("Load", "ID")

    cnn.Open(MGlobals.GetAppConnString)
    rst.CursorLocation = adUseClient
    rst.Open(gStrSqlCall, cnn, ADODB.CursorTypeEnum.adOpenStatic, 1)
    Me.RemoveAll()
    rst.MoveFirst()
    'if this entity has a M-1 rel with any other entity, then a blank entity must be present for cbos.
    If cEntityDataItems.ParentHasM1Relationship("Person") Then
      ePrsn_ = New Person   'this member is loaded for 0 index entries cbo's for M-1 relationships
      ePrsn_.Loadorder = 0
      ePrsn_.ID = ""
      Call cPersons.Add(ePrsn_)
    End If
    nLoadorder = 1
    Do While Not rst.EOF
      ePrsn_ = New Person

      Call Me.LoadEntityItemsForThisEntity(ePrsn_)

      ePrsn_.Lastupdate = rst.Fields("LastUpdate").Value
      ePrsn_.Loadorder = nLoadorder
      ePrsn_.mContainer = Me
      Call cPersons.Add(ePrsn_)
      nLoadorder = nLoadorder + 1
      rst.MoveNext()
    Loop
    rst.Close()
    cnn.Close()
    'Build the domain model now
    'Call Me.BuildDomainModel
  End Function

  Sub BuildDomainModel()
    '  Dim eb_ As Person
    '  Dim e1_ As Entity1
    '  'now add ePrsn_s to e1_s
    '  For Each e1_ In cEntity1s.Items  'assumed cEntity1s have been loaded
    '    Set e1_.ChildPersons = e1_.BuildChildEntityObjects(e1_.ChildPersons, e1_.ChildEntityString(currePrsn_))
    '    'For Each ePrsn_ In e1_.ChildPersons.Items
    '    ' If Not ePrsn_.ContainsParentEntity(e1_) Then  'BuildObjectModel is called from AddtoDB() too. hence to check for already existing.
    '    '  Call ePrsn_.ParentEntity1s.Add(e1_)   'add parents to Person
    '    'End If
    '    'Next ePrsn_
    '  Next e1_
  End Sub

  Public Function UpdateDB(ByRef ePrsn_ As Person) As Integer
    'Updates the database table for the entity'.
    'Return 0 if the db record has been updated by other user.
    'Return the loadorder if successfull in update.
    Dim nCount As Integer
    gbChangesMade = False
    gStrSqlCall = GetSQLQuery("UpdateDB", ePrsn_.ID)
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
    If rst.Fields("LastUpdate").Value <> ePrsn_.Lastupdate Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    Else
      rst.Fields("LastUpdate").Value = Now
      ePrsn_.Lastupdate = rst.Fields("LastUpdate").Value  'update the local last update so you can mame changes without refreshing.
    End If
    strChanges = ""
    ''PH:For Each Entity Item. Add to record changes int the DB (function UpdateDB()
    'Call BLFunctions.gRecordChanges(ePrsn_, "sEntityItem_1", ePrsn_.EntityItem_1, "fn:")
    ''PH: End
    'sPH_col_Update:
    Call RecordChanges("Update", ePrsn_)
    ' Call BLFunctions.gRecordChanges("sEntityItem_1", ePrsn_.EntityItem_1, "ei1:")
    'Call BLFunctions.gRecordChanges("sEntityItem_2", ePrsn_.EntityItem_2, "ei2:")
    'sPH_col_Update:End

    ''PH: For Each Child Entity, add this line for recording the changes for child entities for this object.
    ' Call ePrsn_.BuildChildEntityString(ePrsn_.ChildEntities(e3_Par))
    ' Call BLFunctions.gRecordChanges(ePrsn_, "sEntity3s", ePrsn_.ChildEntityString(e3_Par), "e3_str:")
    ''PH: End
    'sPH_col_UpdateChild:
    'followng two commented
    'Call ePrsn_.BuildChildEntityString("EntityBs")
    'Call BLFunctions.gRecordChanges(ePrsn_, "sEntityBs", ePrsn_.ChildEntityString("EntityBs"), "eB_str:")

    'sPH_col_UpdateChild:End

    rst.Update()
    rst.Close()
    MGlobals.currePrsn_ = ePrsn_

    'No need to enter a record in update log if no change has been made
    If gbChangesMade = True Then
      AppSettings.LastUpdateID = gUpdateLogTable("ePrsn_", "U", ePrsn_.ID)
    End If
    cnn.Close()
    Call DoChanged()
    UpdateDB = nCount
  End Function

  Public Function AddtoDB(ByRef ePrsn_ As Person) As Integer
    Dim strSQL As String
    Dim nCount As Integer
    strSQL = "select Id from Persons"
    Try
      If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
        cnn.Open(MGlobals.GetAppConnString)
      End If
    Catch
      MsgBox(Err.Description)
    End Try
    'rst.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    rst.Open(Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
    'rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
    nCount = rst.RecordCount
    If cEntityDataItems.ParentHasM1Relationship("Person") Then nCount = nCount + 1
    If nCount <> cPersons.Count - 1 Then  'other user has added a record
      rst.Close()
      cnn.Close()
      AddtoDB = 0
      Exit Function
    End If
    rst.Close()
    strSQL = "select max(ID) as MaxNo from Persons"
    Dim nMax As Integer
    rst.CursorLocation = adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, adLockOptimistic)
    nMax = rst.Fields("MaxNo").Value
    rst.Close()
    ePrsn_.ID = nMax + 1
    'now add the record
    rst.CursorType = adOpenKeySet
    rst.LockType = adLockOptimistic
    rst.Open("Persons", cnn, , , adCmdTable)
    rst.AddNew()

    ''PH:For Each Entity Item. Add the items to the db (function AddtoDB)
    'Call gAddChanges("EntityItem_2", ePrsn_.EntityItem_2, "ln:")
    ''PH: End
    'sPH_col_Add:

    Call BLFunctions.gRecordChanges1("Add", "ID", ePrsn_.ID, "id:")
    Call RecordChanges("Add", ePrsn_)
    'Call gAddChanges("EntityItem_1", ePrsn_.EntityItem_1, "ei1:")
    'Call gAddChanges("EntityItem_2", ePrsn_.EntityItem_2, "ei2:")
    'sPH_col_Add:End

    ''PH: For Each Child Entity, add this line for saving the child string with e3_ is the child.
    ' Call gAddChanges("Entity_Cs", ePrsn_.ChildEntityString(e3_Par), "ePrsn_str:")
    ''PH: End
    'sPH_col_AddChild:
    '    Call ePrsn_.BuildChildEntityString(ePrsn_.ChildEntities(CurreB_))
    '    Call gAddChanges("EntityBs", ePrsn_.ChildEntityString(CurreB_), "eB_str:")
    'sPH_col_AddChild:End

    rst("LastUpdate").Value = Now
    ePrsn_.Lastupdate = rst("LastUpdate").Value
    ePrsn_.mContainer = Me
    rst.Update()
    rst.Close()

    Call Me.Add(ePrsn_)
    ePrsn_.Loadorder = Me.Count - nDecLoadOrder

    Call Me.BuildDomainModel()
    MGlobals.currePrsn_ = ePrsn_
    'now update the log
    AppSettings.LastUpdateID = gUpdateLogTable("ePrsn_", "N", ePrsn_.ID)
    cnn.Close()
    Call DoChanged()
    AddtoDB = nCount + 1
  End Function

  Public Function DeleteFromDB(ByVal ePrsn_ As Person) As Integer
    Dim lAffected As Long
    gStrSqlCall = " Delete " & _
                  "FROM " & strDBO & "Persons " & _
                  " where " & strDBO & "ID LIKE '" & ePrsn_.ID & "'"

    MGlobals.cmd.ActiveConnection = cnn
    MGlobals.cmd.CommandText = gStrSqlCall
    MGlobals.cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
    rst = MGlobals.cmd.Execute(lAffected, ADODB.ExecuteOptionEnum.adExecuteNoRecords)
    If lAffected = 1 Then
      Dim sEbID As String
      sEbID = ePrsn_.ID
      MGlobals.gstrChanges = "Deleted: " & "ID" & _
                            "fn:" & ePrsn_.FirstName & _
                            ", ln:" & ePrsn_.LastName
      Call cPersons.Remove(ePrsn_)
      cnn.Open(gStrConnGenEBA)
      AppSettings.LastUpdateID = gUpdateLogTable("ePrsn_", "D", sEbID)
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
      strDBO & "ID, " & _
      strDBO & "FirstName, " & _
      strDBO & "MiddleName, " & _
      strDBO & "LastName , " & _
      strDBO & "LoginID , " & _
      strDBO & "Email, " & _
      strDBO & "Phone , " & _
      strDBO & "AccessRight, " & _
      strDBO & "DateJoined , " & _
      strDBO & "Remarks, " & _
      strDBO & "LastUpdate " & _
      "FROM " & strDBO & "Persons "
    'sPH_col_Sql:End
    If strFunc = "LoadSingleEntity" Or strFunc = "UpdateDB" Then
      strTemp = strTemp & " WHERE " & strDBO & "ID like " & strID
    ElseIf strFunc = "Load" Then
      strTemp = strTemp & " ORDER BY " & strDBO & strID
    End If
    GetSQLQuery = strTemp
  End Function

  Function LoadEntityItemsForThisEntity(ByVal ent As Person)
    ''PH:For Each Entity Item. Add it for loading
    'ePrsn_.EntityItem_1 = "" & rst.Fields("sEntityItem_1")
    ''PH End
    'sPH_col_Load:
    ent.ID = "" & rst.Fields("ID").Value
    ent.FirstName = "" & rst.Fields("FirstName").Value
    ent.MiddleName = "" & rst.Fields("MiddleName").Value
    ent.LastName = "" & rst.Fields("LastName").Value
    ent.LoginID = "" & rst.Fields("LoginID").Value
    ent.Email = "" & rst.Fields("Email").Value
    ent.Phone = "" & rst.Fields("Phone").Value
    ent.AccessRight = "" & rst.Fields("AccessRight").Value
    ent.DateJoined = "" & rst.Fields("DateJoined").Value
    ent.Remarks = "" & rst.Fields("Remarks").Value

    'sPH_col_Load:End

    ''PH: For Each Child Entity, add the two lines for building child entities for this object.
    ' Call ent.BuildChildEntityObjects(e3_Par, "" & rst.Fields("sEntity3s"))
    ''PH: End
    'sPH_col_LoadChildEntities:
    'Call ent.BuildChildEntityObjects("Students", "" & rst.Fields("EntityCs").Value)
    'sPH_col_LoadChildEntities:End
  End Function

  Public Sub RaiseChanged()
    RaiseEvent Changed()
  End Sub

  Private Sub RecordChanges(ByVal sOperation As String, ByVal ent As Person)
    Call BLFunctions.gRecordChanges1(sOperation, "FirstName", ent.FirstName, "fn:")
    Call BLFunctions.gRecordChanges1(sOperation, "MiddleName", ent.MiddleName, "mn:")
    Call BLFunctions.gRecordChanges1(sOperation, "LastName", ent.LastName, "ln:")
    Call BLFunctions.gRecordChanges1(sOperation, "LoginID", ent.LoginID, "logid:")
    Call BLFunctions.gRecordChanges1(sOperation, "Email", ent.Email, "em:")
    Call BLFunctions.gRecordChanges1(sOperation, "Phone", ent.Phone, "phn:")
    Call BLFunctions.gRecordChanges1(sOperation, "AccessRight", ent.AccessRight, "ar:")
    Call BLFunctions.gRecordChanges1(sOperation, "DateJoined", ent.DateJoined, "dj:")
    Call BLFunctions.gRecordChanges1(sOperation, "Remarks", ent.Remarks, "rmks:")

  End Sub

End Class
