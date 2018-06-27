Option Explicit On

Public Class Projects
  Inherits System.Collections.CollectionBase
  Private _sNoHashTable As New Hashtable
  'sPH_col_DateTime:
  'gfh
  'sPH_col_DateTime:End
  'Projects
  Dim nDecLoadOrder As Integer  'comment
  'Private collInt As Collection
  Public mContained As Object
  Dim strChanges As String
  Public Event Changed()
  Dim strSqlCall As String
  Dim strConnFramework As String
  
  Public Sub Add(ByRef val As Project)
    Me.List.Add(val)
    _sNoHashTable.Add(val.ID, val)
  End Sub

  Public ReadOnly Property item(ByVal index As Integer) As Project
    Get
      Return Me.List.Item(index)
    End Get
  End Property

  Public Property Item(ByVal sid As String) As Project
    Get
      Return _sNoHashTable.Item(sid)
    End Get
    Set(ByVal value As Project)
      _sNoHashTable.Item(sid) = value
    End Set
  End Property

  Public Sub Remove(ByRef val As Project)
    Me.List.Remove(val)
  End Sub

  Public Function LoadSingleEntity(ByVal sid As String, ByVal sOperation As String) As Integer
    Dim ePrjt_ As Project
    strConnFramework = MGlobals.GetAppConnString
    strSqlCall = GetSQLQuery("LoadSingleEntity", sid)
    cnn.Open(strConnFramework)
    rst.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    rst.Open(strSqlCall, cnn, _
              ADODB.CursorTypeEnum.adOpenStatic, 1)
    If rst.RecordCount <> 1 Then
      LoadSingleEntity = 0  '
      rst.Close()
      cnn.Close()
      Exit Function
    End If
    If sOperation = "N" Then
      ePrjt_ = New Project
    ElseIf sOperation = "U" Then
      ePrjt_ = Me.Item(sid)
    End If
    Call Me.LoadEntityItemsForThisEntity(ePrjt_)

    ePrjt_.Lastupdate = rst.Fields("LastUpdate").Value
    ePrjt_.mContainer = Me

    If sOperation = "N" Then
      ePrjt_.Loadorder = Me.Count - nDecLoadOrder
      Call cProjects.Add(ePrjt_)
    End If
    rst.Close()
    cnn.Close()
    LoadSingleEntity = 1
  End Function

  Public Function Load()
    Dim ePrjt_ As Project
    Dim nLoadorder As Integer
    'strConnFramework = MGlobals.GetAppConnString
    gStrSqlCall = GetSQLQuery("Load", "ID")
    cnn.Open(MGlobals.GetAppConnString)
    rst.CursorLocation = adUseClient
   rst.Open(gStrSqlCall, cnn, ADODB.CursorTypeEnum.adOpenStatic, 1)
    Me.Clear()
    rst.MoveFirst()
    nDecLoadOrder = 0
    'if this entity has a M-1 rel with any other entity, then a blank entity must be present for cbos.
    If cEntityDataItems.ParentHasM1Relationship("Project") Then
      ePrjt_ = New Project   'this member is loaded for 0 index entries cbo's for M-1 relationships
      ePrjt_.Loadorder = 0
      ePrjt_.ID = 0
      nDecLoadOrder = 1
      Call cProjects.Add(ePrjt_)
    End If
    nLoadorder = 1
    Do While Not rst.EOF
      ePrjt_ = New Project

      Call Me.LoadEntityItemsForThisEntity(ePrjt_)

      ePrjt_.Lastupdate = rst.Fields("LastUpdate").Value
      ePrjt_.Loadorder = nLoadorder
      ePrjt_.mContainer = Me
      Call cProjects.Add(ePrjt_)
      nLoadorder = nLoadorder + 1
      rst.MoveNext()
    Loop
    rst.Close()
    cnn.Close()
    'Build the domain model now
    'Call Me.BuildDomainModel
  End Function

  Sub BuildDomainModel()
    '  Dim eb_ As Project
    '  Dim e1_ As Entity1
    '  'now add ePrjt_s to e1_s
    '  For Each e1_ In cEntity1s.Items  'assumed cEntity1s have been loaded
    '    Set e1_.ChildProjects = e1_.BuildChildEntityObjects(e1_.ChildProjects, e1_.ChildEntityString(CurrePrjt_))
    '    'For Each ePrjt_ In e1_.ChildProjects.Items
    '    ' If Not ePrjt_.ContainsParentEntity(e1_) Then  'BuildObjectModel is called from AddtoDB() too. hence to check for already existing.
    '    '  Call ePrjt_.ParentEntity1s.Add(e1_)   'add parents to Project
    '    'End If
    '    'Next ePrjt_
    '  Next e1_
  End Sub

  Public Function UpdateDB(ByRef ent As Project) As Integer
    'Updates the database table for the entity'.
    'Return 0 if the db record has been updated by other user.
    'Return the loadorder if successfull in update.
    Dim nCount As Integer
    gbChangesMade = False
    gStrSqlCall = GetSQLQuery("UpdateDB", ent.ID)
    If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
     cnn.Open(MGlobals.GetAppConnString)
    End If
    'If rst.State = ADODB.ObjectStateEnum.adStateOpen Then
    '  rst.Close()
    'End If

    'cnn.Open(gStrConnGenEBA)
    rst.Open(Source:=gStrSqlCall, ActiveConnection:=cnn, CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
    nCount = rst.RecordCount
    If nCount <> 1 Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    End If
    If rst.Fields("LastUpdate").Value <> ent.Lastupdate Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    Else
      rst.Fields("LastUpdate").Value = Now
      ent.Lastupdate = rst.Fields("LastUpdate").Value  'update the local last update so you can mame changes without refreshing.
    End If
    strChanges = ""
    ''PH:For Each Entity Item. Add to record changes int the DB (function UpdateDB()
    'Call BLFunctions.gRecordChanges(ePrjt_, "sEntityItem_1", ePrjt_.EntityItem_1, "fn:")
    ''PH: End
    'sPH_col_Update:
    Call Recordchanges("Update", ent)
    'Call BLFunctions.gRecordChanges("sProjectName", ent.ProjectName, "pn:")
    'Call BLFunctions.gRecordChanges("sProjectDescription", ent.ProjectDescription, "pd:")
    'sPH_col_Update:End

    ''PH: For Each Child Entity, add this line for recording the changes for child entities for this object.
    ' Call ePrjt_.BuildChildEntityString(ePrjt_.ChildEntities(e3_Par))
    ' Call BLFunctions.gRecordChanges(ePrjt_, "sEntity3s", ePrjt_.ChildEntityString(e3_Par), "e3_str:")
    ''PH: End
    'sPH_col_UpdateChild:
    'sPH_col_UpdateChild
    'sPH_col_UpdateChild:End

    rst.Update()
    rst.Close()
    MGlobals.currePrjt_ = ent

    'No need to enter a record in update log if no change has been made
    If gbChangesMade = True Then
      AppSettings.LastUpdateID = gUpdateLogTable("ePrjt_", "U", ent.ID)
    End If
    cnn.Close()
    Call DoChanged()
    UpdateDB = nCount
  End Function

  Public Function AddtoDB(ByRef ent As Project) As Integer
    Dim strSQL As String
    Dim nCount As Integer
    strSQL = "select ID from Projects"

    If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      cnn.Open(MGlobals.GetAppConnString)
    End If

    rst.Open(Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
    nCount = rst.RecordCount
    If cEntityDataItems.ParentHasM1Relationship("Project") Then nCount = nCount + 1
    If nCount <> cProjects.Count - 1 Then  'other user has added a record
      rst.Close()
      cnn.Close()
      AddtoDB = 0
      Exit Function
    End If
    rst.Close()
    strSQL = "select max(ID) as MaxNo from Projects"
    Dim nMax As Integer
    rst.CursorLocation = adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, adLockOptimistic)
    nMax = rst.Fields("MaxNo").Value
    rst.Close()
    ent.ID = nMax + 1
    'now add the record
    rst.CursorType = adOpenKeySet
    rst.LockType = adLockOptimistic
    rst.Open("Projects", cnn, , , adCmdTable)
    rst.AddNew()

    ''PH:For Each Entity Item. Add the items to the db (function AddtoDB)
    'Call gAddChanges("EntityItem_2", ePrjt_.EntityItem_2, "ln:")
    ''PH: End
    'sPH_col_Add:
    Call BLFunctions.gRecordChanges1("Add", "ID", ent.ID, "id:")
    Call Recordchanges("Add", ent)
    'sPH_col_Add:End

    ''PH: For Each Child Entity, add this line for saving the child string with e3_ is the child.
    ' Call gAddChanges("Entity_Cs", ePrjt_.ChildEntityString(e3_Par), "ePrjt_str:")
    ''PH: End
    'sPH_col_AddChild:
    ''''sPH_col_AddChild

    'sPH_col_AddChild:End

    rst("LastUpdate").Value = Now
    ent.Lastupdate = rst("LastUpdate").Value
    'ent.mContainer = Me
    rst.Update()
    rst.Close()

    ent.Loadorder = Me.Count - nDecLoadOrder
    'Call Me.Add(ent)

    Call Me.BuildDomainModel()
    MGlobals.currePrjt_ = ent

    'now update the log
    AppSettings.LastUpdateID = gUpdateLogTable("ePrjt_", "N", ent.ID)
    cnn.Close()
    Call DoChanged()
    AddtoDB = nCount + 1
  End Function

  Public Function DeleteFromDB(ByVal ent As Project) As Integer
    Dim lAffected As Long

    gStrSqlCall = " Delete " & _
                  "FROM " & strDBO & "Projects " & _
                  " where " & strDBO & "ID LIKE '" & ent.ID & "'"
    cmd.ActiveConnection = cnn
    'Globals.gADOcmd.ActiveConnection = gStrConnGenEBA
    cmd.CommandText = gStrSqlCall
    cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
    cmd.Execute(lAffected, ADODB.ExecuteOptionEnum.adExecuteNoRecords)
    If lAffected = 1 Then
      Dim sEbID As String
      sEbID = ent.ID
      MGlobals.gstrChanges = "Deleted: " & "ID"
      Call cProjects.Remove(ent)
      cnn.Open(gStrConnGenEBA)
      AppSettings.LastUpdateID = gUpdateLogTable("ePrjt_", "D", sEbID)
      cnn.Close()
      Call DoChanged()
    End If
    cmd = Nothing
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
              strDBO & "ID as ID, " & _
              strDBO & "ProjectName , " & _
              strDBO & "ProjectDescription , " & _
              strDBO & "ProjectEntities , " & _
              strDBO & "LastUpdate " & _
              "FROM " & strDBO & "Projects "
    'sPH_col_Sql:End

    If strFunc = "LoadSingleEntity" Or strFunc = "UpdateDB" Then
      strTemp = strTemp & " WHERE " & strDBO & "ID like " & strID
    ElseIf strFunc = "Load" Then
      strTemp = strTemp & " ORDER BY " & strDBO & strID
    End If
    GetSQLQuery = strTemp
  End Function

  Function LoadEntityItemsForThisEntity(ByRef ent As Project)
    ''PH:For Each Entity Item. Add it for loading
    'eA_.EntityItem_1 = "" & rst.Fields("sEntityItem_1")
    ''PH End
    'sPH_col_Load:
    ent.ID = "" & rst.Fields("ID").Value
    ent.ProjectName = "" & rst.Fields("ProjectName").Value
    ent.ProjectDescription = "" & rst.Fields("ProjectDescription").Value
    ent.ProjectEntitiesString = "" & rst.Fields("ProjectEntities").Value
    'sPH_col_Load:End

    ''PH: For Each Child Entity, add the two lines for building child entities for this object.
    ' Call ent.BuildChildEntityObjects(e3_Par, "" & rst.Fields("sEntity3s"))
    ''PH: End
    ' If cEntityDataItems.IsJoinTable(TypeName(ent)) = True Then
    'sPH_col_LoadChildEntitiesJT:
    'Call ent.BuildChildEntityObjects("ProjectEntities", "" & rst.Fields("ProjectEntities").Value)
    'sPH_col_LoadChildEntitiesJT:End
    'Else
    'sPH_col_LoadChildEntitiesCS:
    ' Call ent.BuildChildEntityObjects("JTOrders", "" & rst.Fields("sOrders"))
    'sPH_col_LoadChildEntitiesCS:End
    'End If
  End Function

  Public Sub RaiseChanged()
    RaiseEvent Changed()
  End Sub

  Sub Recordchanges(ByVal sOperation As String, ByVal ent As Project)
    Call BLFunctions.gRecordChanges1(sOperation, "ProjectName", ent.ProjectName, "pn:")
    Call BLFunctions.gRecordChanges1(sOperation, "ProjectDescription", ent.ProjectDescription, "pd:")

  End Sub


End Class
