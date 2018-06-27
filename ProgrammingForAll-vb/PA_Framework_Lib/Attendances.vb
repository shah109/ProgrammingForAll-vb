Option Explicit On
Public Class Attendances
  Inherits System.Collections.CollectionBase

  Private _sNoHashTable As New Hashtable

  Dim nDecLoadOrder As Integer  'comment
  Dim strChanges As String

  Public Event Changed()
  Public Sub Add(ByRef val As Attendance)
    Me.List.Add(val)
    _sNoHashTable.Add(val.ID, val)
  End Sub


  'Public ReadOnly Property Item(ByVal index As Integer) As Attendance
  '  Get
  '    Return Me.List.Item(index)
  '  End Get
  'End Property

  Public Property Item(ByVal sid As String) As Attendance
    Get
      Return _sNoHashTable.Item(sid)
    End Get
    Set(ByVal value As Attendance)
      _sNoHashTable.Item(sid) = value
    End Set
  End Property

  Public Sub Remove(ByRef val As Attendance)
    Me.List.Remove(val)
    _sNoHashTable.Remove(val.ID)
  End Sub

  Public Sub RemoveAll()
    Me.List.Clear()    'alway remove the first item
    _sNoHashTable.Clear()
  End Sub

  Public Function LoadSingleEntity(ByVal sId As String, ByVal sOperation As String) As Integer
    Dim eAtt_ As Attendance

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
      eAtt_ = New Attendance
    ElseIf sOperation = "U" Then
      eAtt_ = Me.Item(sId)
    End If
    Call Me.LoadEntityItemsForThisEntity(eAtt_)

    eAtt_.Lastupdate = rst.Fields("LastUpdate").Value
    eAtt_.mContainer = Me

    eAtt_.Loadorder = Me.Count
    If sOperation = "N" Then
      Call cAttendances.Add(eAtt_)
    End If
    rst.Close()
    cnn.Close()
    LoadSingleEntity = 1
  End Function

  Public Function Load()
    Dim eAtt_ As Attendance
    Dim nLoadorder As Integer
    gStrSqlCall = GetSQLQuery("Load", "ID")

    cnn.Open(MGlobals.GetAppConnString)
    rst.CursorLocation = adUseClient
    rst.Open(gStrSqlCall, cnn, ADODB.CursorTypeEnum.adOpenStatic, 1)
    Me.RemoveAll()
    rst.MoveFirst()
    'if this entity has a M-1 rel with any other entity, then a blank entity must be present for cbos.
    If cEntityDataItems.ParentHasM1Relationship("Attendance") Then
      eAtt_ = New Attendance   'this member is loaded for 0 index entries cbo's for M-1 relationships
      eAtt_.Loadorder = 0
      eAtt_.ID = 0
      Call cAttendances.Add(eAtt_)
    End If
    nLoadorder = 1
    Do While Not rst.EOF
      eAtt_ = New Attendance

      Call Me.LoadEntityItemsForThisEntity(eAtt_)

      eAtt_.Lastupdate = rst.Fields("LastUpdate").Value
      eAtt_.Loadorder = nLoadorder
      eAtt_.mContainer = Me
      Call cAttendances.Add(eAtt_)
      nLoadorder = nLoadorder + 1
      rst.MoveNext()
    Loop
    rst.Close()
    cnn.Close()
    'Build the domain model now
    'Call Me.BuildDomainModel
  End Function

  Sub BuildDomainModel()
    '  Dim eb_ As Attendance
    '  Dim e1_ As Entity1
    '  'now add eCrse_s to e1_s
    '  For Each e1_ In cEntity1s.Items  'assumed cEntity1s have been loaded
    '    Set e1_.ChildAttendances = e1_.BuildChildEntityObjects(e1_.ChildAttendances, e1_.ChildEntityString(CurreCrse_))
    '    'For Each eAtt_ In e1_.ChildAttendances.Items
    '    ' If Not eAtt_.ContainsParentEntity(e1_) Then  'BuildObjectModel is called from AddtoDB() too. hence to check for already existing.
    '    '  Call eAtt_.ParentEntity1s.Add(e1_)   'add parents to Attendance
    '    'End If
    '    'Next eAtt_
    '  Next e1_
  End Sub

  Public Function UpdateDB(ByRef eAtt_ As Attendance) As Integer
    'Updates the database table for the entity'.
    'Return 0 if the db record has been updated by other user.
    'Return the loadorder if successfull in update.
    Dim nCount As Integer
    gbChangesMade = False
    gStrSqlCall = GetSQLQuery("UpdateDB", eAtt_.ID)
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
    If rst.Fields("LastUpdate").Value <> eAtt_.Lastupdate Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    Else
      rst.Fields("LastUpdate").Value = Now
      eAtt_.Lastupdate = rst.Fields("LastUpdate").Value  'update the local last update so you can mame changes without refreshing.
    End If
    strChanges = ""
    ''PH:For Each Entity Item. Add to record changes int the DB (function UpdateDB()
    'Call BLFunctions.gRecordChanges(eAtt_, "sEntityItem_1", eAtt_.EntityItem_1, "fn:")
    ''PH: End
    'sPH_col_Update:
    Call RecordChanges("Update", eAtt_)
    ' Call BLFunctions.gRecordChanges("sEntityItem_1", eAtt_.EntityItem_1, "ei1:")
    'Call BLFunctions.gRecordChanges("sEntityItem_2", eAtt_.EntityItem_2, "ei2:")
    'sPH_col_Update:End

    ''PH: For Each Child Entity, add this line for recording the changes for child entities for this object.
    ' Call eAtt_.BuildChildEntityString(eAtt_.ChildEntities(e3_Par))
    ' Call BLFunctions.gRecordChanges(eAtt_, "sEntity3s", eAtt_.ChildEntityString(e3_Par), "e3_str:")
    ''PH: End
    'sPH_col_UpdateChild:
    'followng two commented
    'Call eAtt_.BuildChildEntityString("EntityBs")
    'Call BLFunctions.gRecordChanges(eAtt_, "sEntityBs", eAtt_.ChildEntityString("EntityBs"), "eB_str:")

    'sPH_col_UpdateChild:End

    rst.Update()
    rst.Close()
    MGlobals.curreAtt_ = eAtt_

    'No need to enter a record in update log if no change has been made
    If gbChangesMade = True Then
      AppSettings.LastUpdateID = gUpdateLogTable("eAtt_", "U", eAtt_.ID)
    End If
    cnn.Close()
    Call DoChanged()
    UpdateDB = nCount
  End Function

  Public Function AddtoDB(ByRef eAtt_ As Attendance) As Integer
    Dim strSQL As String
    Dim nCount As Integer
    strSQL = "select Id from Attendances"
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
    If cEntityDataItems.ParentHasM1Relationship("Attendance") Then nCount = nCount + 1
    If nCount <> cAttendances.Count - 1 Then  'other user has added a record
      rst.Close()
      cnn.Close()
      AddtoDB = 0
      Exit Function
    End If
    rst.Close()
    strSQL = "select max(ID) as MaxNo from Attendances"
    Dim nMax As Integer
    rst.CursorLocation = adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, adLockOptimistic)
    nMax = rst.Fields("MaxNo").Value
    rst.Close()
    eAtt_.ID = nMax + 1
    'now add the record
    rst.CursorType = adOpenKeySet
    rst.LockType = adLockOptimistic
    rst.Open("Attendances", cnn, , , adCmdTable)
    rst.AddNew()

    ''PH:For Each Entity Item. Add the items to the db (function AddtoDB)
    'Call gAddChanges("EntityItem_2", eAtt_.EntityItem_2, "ln:")
    ''PH: End
    'sPH_col_Add:

    Call BLFunctions.gRecordChanges1("Add", "ID", eAtt_.ID, "id:")
    Call RecordChanges("Add", eAtt_)
    'Call gAddChanges("EntityItem_1", eAtt_.EntityItem_1, "ei1:")
    'Call gAddChanges("EntityItem_2", eAtt_.EntityItem_2, "ei2:")
    'sPH_col_Add:End

    ''PH: For Each Child Entity, add this line for saving the child string with e3_ is the child.
    ' Call gAddChanges("Entity_Cs", eAtt_.ChildEntityString(e3_Par), "eCrse_str:")
    ''PH: End
    'sPH_col_AddChild:
    '    Call eAtt_.BuildChildEntityString(eAtt_.ChildEntities(CurreB_))
    '    Call gAddChanges("EntityBs", eAtt_.ChildEntityString(CurreB_), "eB_str:")
    'sPH_col_AddChild:End

    rst("LastUpdate").Value = Now
    eAtt_.Lastupdate = rst("LastUpdate").Value
    eAtt_.mContainer = Me
    rst.Update()
    rst.Close()

    Call Me.Add(eAtt_)
    eAtt_.Loadorder = Me.Count - nDecLoadOrder

    Call Me.BuildDomainModel()
    MGlobals.curreAtt_ = eAtt_
    'now update the log
    AppSettings.LastUpdateID = gUpdateLogTable("eAtt_", "N", eAtt_.ID)
    cnn.Close()
    Call DoChanged()
    AddtoDB = nCount + 1
  End Function

  Public Function DeleteFromDB(ByVal eAtt_ As Attendance) As Integer
    Dim lAffected As Long
    Dim strID As String = eAtt_.ID

    gStrSqlCall = " Delete " & _
                  "FROM " & strDBO & "Attendances " & _
                  " where " & strDBO & "ID LIKE '" & eAtt_.ID & "'"

    If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      cnn.Open(MGlobals.GetAppConnString)
    End If

    MGlobals.cmd.ActiveConnection = cnn
    MGlobals.cmd.CommandText = gStrSqlCall
    MGlobals.cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
    rst = MGlobals.cmd.Execute(lAffected, ADODB.ExecuteOptionEnum.adExecuteNoRecords)
    cnn.Close()
    If lAffected = 1 Then
      MGlobals.gstrChanges = "Deleted: " & "ID"
      Call cAttendances.Remove(eAtt_)
      cnn.Open(gStrConnGenEBA)
      AppSettings.LastUpdateID = gUpdateLogTable("eAtt_", "D", strID)
      cnn.Close()
      Call DoChanged()
    End If
    If cnn.State = ADODB.ObjectStateEnum.adStateOpen Then
      cnn.Close()
    End If
    MGlobals.cmd = Nothing
    DeleteFromDB = lAffected

    ''''''''''''''''''''''''''''''''
    'Dim lAffected As Long
    'Dim nCount As Integer

    'gStrSqlCall = " Delete " & _
    '              "FROM " & strDBO & "Attendances " & _
    '              " where " & strDBO & "ID LIKE '" & eAtt_.ID & "'"

    'Globals.gADOcmd.ActiveConnection = gStrConnGenEBA
    'Globals.gADOcmd.CommandText = gStrSqlCall
    'Globals.gADOcmd.CommandType = adCmdText
    'Globals.gADOcmd.Execute(lAffected, adExecuteNoRecords)
    'If lAffected = 1 Then
    '  Dim sEbID As String
    '  sEbID = eAtt_.ID
    '  Globals.gstrChanges = "Deleted: " & "ID"
    '  Call cAttendances.Remove(eAtt_)
    '  cnn.Open(gStrConnGenEBA)
    '  frmSettings.LastUpdateID = gUpdateLogTable("eAtt_", "D", sEbID)
    '  cnn.Close()
    '  Call DoChanged()
    'End If
    'Globals.gADOcmd = Nothing
    'DeleteFromDB = lAffected
    '''''''''''''''''''''''''''''''''

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
                strDBO & "StudentID, " & _
                strDBO & "CalendarID , " & _
                strDBO & "Comments , " & _
                strDBO & "LastUpdate " & _
                "FROM " & strDBO & "Attendances "    'sPH_col_Sql:End

    If strFunc = "LoadSingleEntity" Or strFunc = "UpdateDB" Then
      strTemp = strTemp & " WHERE " & strDBO & "ID like " & strID
    ElseIf strFunc = "Load" Then
      strTemp = strTemp & " ORDER BY " & strDBO & strID
    End If
    GetSQLQuery = strTemp
  End Function

  Function LoadEntityItemsForThisEntity(ByVal ent As Attendance)
    ''PH:For Each Entity Item. Add it for loading
    'eAtt_.EntityItem_1 = "" & rst.Fields("sEntityItem_1")
    ''PH End
    'sPH_col_Load:
    ent.ID = "" & rst.Fields("ID").Value
    ent.Comments = "" & rst.Fields("Comments").Value

    'ent.Students = "" & rst.Fields("sEntityBs").Value
    'sPH_col_Load:End

    ''PH: For Each Child Entity, add the two lines for building child entities for this object.
    ' Call ent.BuildChildEntityObjects(e3_Par, "" & rst.Fields("sEntity3s"))
    ''PH: End
    'sPH_col_LoadChildEntities:
    Call ent.BuildChildEntityObjects("Students", "" & rst.Fields("StudentID").Value)
    Call ent.BuildChildEntityObjects("CalendarItems", "" & rst.Fields("CalendarID").Value)
    'sPH_col_LoadChildEntities:End
  End Function

  Public Sub RaiseChanged()
    RaiseEvent Changed()
  End Sub

  Private Sub RecordChanges(ByVal sOperation As String, ByVal ent As Attendance)
    Call BLFunctions.gRecordChanges1(sOperation, "Comments", ent.Comments, "cmnts:")
    Call BLFunctions.gRecordChanges1(sOperation, "StudentID", ent.ChildEntityString("Students"), "Stds:")
    Call BLFunctions.gRecordChanges1(sOperation, "CalendarID", ent.ChildEntityString("Calendars"), "Cals:")
  End Sub
End Class
